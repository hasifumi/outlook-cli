from __future__ import annotations
import os
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal, Vertical
from textual.widgets import (
    Header, Footer, Label, ListView, ListItem,
    TextArea, Input, Static, Button
)
from textual.screen import ModalScreen
from textual.reactive import reactive
from textual import on


def get_client():
    if os.getenv("OUTLOOK_MOCK"):
        from .mock import OutlookMock
        return OutlookMock()
    from .com import OutlookCOM
    return OutlookCOM()


# ────────────────────────────────────────────
# 送信・返信画面
# ────────────────────────────────────────────
class ComposeScreen(ModalScreen):
    BINDINGS = [
        Binding("ctrl+enter", "send", "送信"),
        Binding("escape", "cancel", "キャンセル"),
        Binding("tab", "focus_next", "次のフィールド", show=False),
    ]

    DEFAULT_CSS = """
    ComposeScreen {
        align: center middle;
    }
    #compose-dialog {
        width: 90%;
        height: 90%;
        background: $surface;
        border: solid $primary;
        padding: 1 2;
    }
    #compose-dialog Label {
        color: $text-muted;
        width: 6;
    }
    .field-row {
        height: 3;
        margin-bottom: 0;
    }
    .field-row Input {
        width: 1fr;
    }
    #body-area {
        height: 1fr;
        margin-top: 1;
        border: solid $primary-darken-2;
    }
    #compose-buttons {
        height: 3;
        align: right middle;
    }
    #candidate-list {
        background: $surface-darken-1;
        border: solid $accent;
        max-height: 6;
        display: none;
        layer: above;
    }
    #candidate-list.visible {
        display: block;
    }
    """

    def __init__(self, client, contacts, mode="new", mail=None):
        super().__init__()
        self.client = client
        self.contacts = contacts
        self.mode = mode
        self.original_mail = mail
        self._candidates: list = []

    def compose(self) -> ComposeResult:
        to_val = ""
        subject_val = ""
        body_val = ""

        if self.mode == "reply" and self.original_mail:
            to_val = self.original_mail.get("from", "")
            subject_val = f"Re: {self.original_mail.get('subject', '')}"
            body_val = f"\n\n--- 元のメール ---\n{self.original_mail.get('body', '')}"
        elif self.mode == "forward" and self.original_mail:
            subject_val = f"Fw: {self.original_mail.get('subject', '')}"
            body_val = f"\n\n--- 転送メッセージ ---\n{self.original_mail.get('body', '')}"

        with Container(id="compose-dialog"):
            with Horizontal(classes="field-row"):
                yield Label("TO:")
                yield Input(value=to_val, placeholder="宛先を入力...", id="input-to")
            yield ListView(id="candidate-list")
            with Horizontal(classes="field-row"):
                yield Label("CC:")
                yield Input(placeholder="CC", id="input-cc")
            with Horizontal(classes="field-row"):
                yield Label("BCC:")
                yield Input(placeholder="BCC", id="input-bcc")
            with Horizontal(classes="field-row"):
                yield Label("件名:")
                yield Input(value=subject_val, placeholder="件名", id="input-subject")
            yield TextArea(body_val, id="body-area")
            with Horizontal(id="compose-buttons"):
                yield Button("送信 [Ctrl+Enter]", id="btn-send", variant="primary")
                yield Button("キャンセル [Esc]", id="btn-cancel")

    @on(Input.Changed, "#input-to")
    def on_to_changed(self, event: Input.Changed) -> None:
        query = event.value.split(";")[-1].strip()
        candidate_list = self.query_one("#candidate-list", ListView)
        if len(query) >= 2:
            matches = [
                c for c in self.contacts
                if query.lower() in c["name"].lower()
                or query.lower() in c["email"].lower()
            ]
            self._candidates = matches[:6]
            candidate_list.clear()
            for i, c in enumerate(self._candidates):
                candidate_list.append(
                    ListItem(Label(f"{c['name']}  {c['email']}"), id=f"c-{i}")
                )
            candidate_list.add_class("visible") if matches else candidate_list.remove_class("visible")
        else:
            self._candidates = []
            candidate_list.clear()
            candidate_list.remove_class("visible")

    @on(ListView.Selected, "#candidate-list")
    def on_candidate_selected(self, event: ListView.Selected) -> None:
        i = int(event.item.id.replace("c-", ""))
        email = self._candidates[i]["email"]
        input_to = self.query_one("#input-to", Input)
        parts = input_to.value.split(";")
        parts[-1] = email
        input_to.value = "; ".join(parts) + "; "
        candidate_list = self.query_one("#candidate-list", ListView)
        candidate_list.clear()
        candidate_list.remove_class("visible")
        input_to.focus()

    def action_send(self) -> None:
        to = self.query_one("#input-to", Input).value.strip().rstrip(";")
        subject = self.query_one("#input-subject", Input).value.strip()
        body = self.query_one("#body-area", TextArea).text
        if not to or not subject:
            self.notify("TO・件名は必須です", severity="error")
            return
        try:
            self.client.send(to=to, subject=subject, body=body)
            self.notify("送信しました", severity="information")
            self.dismiss(True)
        except Exception as e:
            self.notify(f"送信エラー: {e}", severity="error")

    @on(Button.Pressed, "#btn-send")
    def on_send(self) -> None:
        self.action_send()

    @on(Button.Pressed, "#btn-cancel")
    def action_cancel(self) -> None:
        self.dismiss(False)


# ────────────────────────────────────────────
# メイン画面
# ────────────────────────────────────────────
class OutlookTUI(App):
    TITLE = "Outlook TUI"
    BINDINGS = [
        Binding("n", "new_mail", "新規"),
        Binding("r", "reply", "返信"),
        Binding("f", "forward", "転送"),
        Binding("u", "toggle_unread", "未読切替"),
        Binding("d", "delete_mail", "削除"),
        Binding("ctrl+r", "refresh", "更新"),
        Binding("slash", "search", "検索"),
        Binding("q", "quit", "終了"),
        Binding("tab", "switch_pane", "ペイン切替", show=False),
        Binding("escape", "clear_search", "検索解除", show=False),
        # Vim キーバインド
        Binding("j", "vim_down", show=False),
        Binding("k", "vim_up", show=False),
        Binding("h", "vim_left", show=False),
        Binding("l", "vim_right", show=False),
        Binding("g", "vim_top", show=False),
        Binding("G", "vim_bottom", show=False),
    ]

    DEFAULT_CSS = """
    #main-layout {
        layout: horizontal;
        height: 1fr;
    }
    #folder-pane {
        width: 22;
        border-right: solid $primary-darken-2;
        background: $surface-darken-1;
    }
    #folder-pane ListView {
        background: transparent;
    }
    #right-pane {
        width: 1fr;
        layout: vertical;
    }
    #mail-list {
        height: 50%;
        border-bottom: solid $primary-darken-2;
    }
    #mail-list ListView {
        height: 1fr;
    }
    #preview-pane {
        height: 1fr;
        padding: 1 2;
    }
    #preview-subject {
        text-style: bold;
        color: $text;
    }
    #preview-meta {
        color: $text-muted;
        margin-bottom: 1;
    }
    #preview-body {
        height: 1fr;
        border: none;
    }
    #search-bar {
        height: 3;
        display: none;
        background: $surface-darken-1;
        border-bottom: solid $accent;
        padding: 0 1;
        align: left middle;
    }
    #search-bar.visible {
        display: block;
    }
    #search-bar Label {
        width: 6;
        color: $accent;
    }
    .folder-item {
        padding: 0 1;
    }
    .unread-badge {
        color: $accent;
        text-style: bold;
    }
    .mail-unread {
        text-style: bold;
        color: $text;
    }
    .mail-read {
        color: $text-muted;
    }
    #folder-title {
        background: $primary-darken-2;
        color: $text;
        padding: 0 1;
        text-style: bold;
        height: 3;
        content-align: left middle;
    }
    """

    current_folder = reactive("inbox")

    FOLDER_LABELS = {
        "inbox": "受信トレイ",
        "sent": "送信済み",
        "drafts": "下書き",
        "trash": "ゴミ箱",
    }

    async def on_mount(self) -> None:
        self.client = get_client()
        self.contacts: list = self.client.get_contacts()
        self.current_mails: list = []
        self.selected_mail: dict | None = None
        self._subfolders: list[str] = []
        await self.load_folders()
        await self.load_mails("inbox")

    def compose(self) -> ComposeResult:
        yield Header()
        with Horizontal(id="main-layout"):
            with Vertical(id="folder-pane"):
                yield Static("フォルダ", id="folder-title")
                yield ListView(id="folder-list")
            with Vertical(id="right-pane"):
                with Horizontal(id="search-bar"):
                    yield Label("検索:")
                    yield Input(placeholder="キーワードを入力...", id="search-input")
                with Container(id="mail-list"):
                    yield ListView(id="mail-list-view")
                with Container(id="preview-pane"):
                    yield Static("", id="preview-subject")
                    yield Static("", id="preview-meta")
                    yield TextArea("メールを選択してください", id="preview-body", read_only=True)
        yield Footer()

    async def load_folders(self) -> None:
        folder_list = self.query_one("#folder-list", ListView)
        await folder_list.clear()
        for key, label in self.FOLDER_LABELS.items():
            unread = self.client.get_unread_count(key)
            badge = f" ({unread})" if unread > 0 else ""
            folder_list.append(
                ListItem(Label(f"{label}{badge}"), id=f"folder-{key}")
            )
        self._subfolders = self.client.list_subfolders()
        for i, name in enumerate(self._subfolders):
            unread = self.client.get_unread_count(name)
            badge = f" ({unread})" if unread > 0 else ""
            folder_list.append(
                ListItem(Label(f"  {name}{badge}"), id=f"folder-sub-{i}")
            )

    async def load_mails(self, folder: str, keyword: str = "") -> None:
        self.current_folder = folder
        if keyword:
            mails = self.client.search(keyword=keyword, days=30)
        else:
            mails = self.client.list_mails(folder=folder, limit=50)
        self.current_mails = mails
        mail_list = self.query_one("#mail-list-view", ListView)
        await mail_list.clear()
        for i, m in enumerate(mails):
            date_str = m["date"][:10]
            sender = m.get("from_name") or m.get("from", "")
            subject = m.get("subject", "")
            unread_mark = "★ " if m.get("unread") else "　 "
            attach_mark = "📎" if m.get("has_attachments") else "　"
            label_text = f"{unread_mark}{attach_mark} {date_str}  {sender:<14}  {subject}"
            style = "mail-unread" if m.get("unread") else "mail-read"
            mail_list.append(
                ListItem(Label(label_text, classes=style), id=f"mail-{i}")
            )
        self.query_one("#preview-subject", Static).update("")
        self.query_one("#preview-meta", Static).update("")
        self.query_one("#preview-body", TextArea).load_text("メールを選択してください")

    @on(ListView.Selected, "#mail-list-view")
    def on_mail_selected(self, event: ListView.Selected) -> None:
        i = int(event.item.id.replace("mail-", ""))
        mail = self.current_mails[i] if i < len(self.current_mails) else None
        if mail:
            self.selected_mail = mail
            self.query_one("#preview-subject", Static).update(mail.get("subject", ""))
            meta = f"From: {mail.get('from_name', '')} <{mail.get('from', '')}> | {mail.get('date', '')[:16]}"
            self.query_one("#preview-meta", Static).update(meta)
            try:
                full = self.client.read(mail["id"])
                body = full.get("body", "")
            except Exception as e:
                self.notify(f"本文取得エラー: {e}", severity="warning")
                body = mail.get("body", "")
            self.query_one("#preview-body", TextArea).load_text(body)

    @on(ListView.Selected, "#folder-list")
    async def on_folder_selected(self, event: ListView.Selected) -> None:
        item_id = event.item.id
        if item_id.startswith("folder-sub-"):
            i = int(item_id.replace("folder-sub-", ""))
            folder = self._subfolders[i]
        else:
            folder = item_id.replace("folder-", "")
        await self.load_mails(folder)

    def action_switch_pane(self) -> None:
        focused = self.focused
        if focused and "folder" in str(type(focused).__name__).lower():
            self.query_one("#mail-list-view").focus()
        else:
            self.query_one("#folder-list").focus()

    def action_search(self) -> None:
        search_bar = self.query_one("#search-bar")
        search_bar.add_class("visible")
        self.query_one("#search-input", Input).focus()

    async def action_clear_search(self) -> None:
        search_bar = self.query_one("#search-bar")
        search_bar.remove_class("visible")
        self.query_one("#search-input", Input).value = ""
        await self.load_mails(self.current_folder)

    @on(Input.Submitted, "#search-input")
    async def on_search_submitted(self, event: Input.Submitted) -> None:
        keyword = event.value.strip()
        if keyword:
            await self.load_mails(self.current_folder, keyword=keyword)
        self.query_one("#mail-list-view").focus()

    def action_new_mail(self) -> None:
        self.push_screen(
            ComposeScreen(self.client, self.contacts, mode="new"),
            self._after_compose
        )

    def _get_full_mail(self) -> dict | None:
        if not self.selected_mail:
            return None
        try:
            return self.client.read(self.selected_mail["id"])
        except Exception:
            return self.selected_mail

    def action_reply(self) -> None:
        mail = self._get_full_mail()
        if not mail:
            self.notify("返信するメールを選択してください", severity="warning")
            return
        self.push_screen(
            ComposeScreen(self.client, self.contacts, mode="reply", mail=mail),
            self._after_compose
        )

    def action_forward(self) -> None:
        mail = self._get_full_mail()
        if not mail:
            self.notify("転送するメールを選択してください", severity="warning")
            return
        self.push_screen(
            ComposeScreen(self.client, self.contacts, mode="forward", mail=mail),
            self._after_compose
        )

    def _after_compose(self, result) -> None:
        if result:
            self.call_after_refresh(self._reload)

    async def _reload(self) -> None:
        await self.load_folders()
        await self.load_mails(self.current_folder)

    async def action_toggle_unread(self) -> None:
        if not self.selected_mail:
            return
        self.selected_mail["unread"] = not self.selected_mail.get("unread", False)
        await self.load_mails(self.current_folder)
        self.notify("未読/既読を切り替えました")

    async def action_delete_mail(self) -> None:
        if not self.selected_mail:
            self.notify("削除するメールを選択してください", severity="warning")
            return
        try:
            self.client.delete(self.selected_mail["id"])
            self.notify(f"削除しました: {self.selected_mail.get('subject', '')}")
        except Exception as e:
            self.notify(f"削除エラー: {e}", severity="error")
            return
        self.selected_mail = None
        await self.load_folders()
        await self.load_mails(self.current_folder)

    async def action_refresh(self) -> None:
        await self.load_folders()
        await self.load_mails(self.current_folder)
        self.notify("更新しました")

    def _main_list_focused(self) -> ListView | None:
        """フォルダ/メール一覧のListViewにフォーカスがある場合のみ返す"""
        focused = self.focused
        if isinstance(focused, ListView) and focused.id in ("folder-list", "mail-list-view"):
            return focused
        return None

    def action_vim_down(self) -> None:
        lv = self._main_list_focused()
        if lv:
            lv.action_cursor_down()

    def action_vim_up(self) -> None:
        lv = self._main_list_focused()
        if lv:
            lv.action_cursor_up()

    def action_vim_left(self) -> None:
        if self._main_list_focused() is not None:
            self.query_one("#folder-list").focus()

    def action_vim_right(self) -> None:
        if self._main_list_focused() is not None:
            self.query_one("#mail-list-view").focus()

    def action_vim_top(self) -> None:
        lv = self._main_list_focused()
        if lv and len(lv) > 0:
            lv.index = 0

    def action_vim_bottom(self) -> None:
        lv = self._main_list_focused()
        if lv and len(lv) > 0:
            lv.index = len(lv) - 1


def main():
    app = OutlookTUI()
    app.run()


if __name__ == "__main__":
    main()
