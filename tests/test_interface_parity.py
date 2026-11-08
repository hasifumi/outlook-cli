"""OutlookBase の全抽象メソッドを、モックと COM が同じ引数名で実装していることを確かめる。

COM は Windows + Outlook（pywin32）が無いと import できないので、ソースを AST で読む。
実装漏れ・引数名の食い違いは、会社 PC で OutlookCOM() を作った瞬間（TypeError）や
CLI を実行した瞬間にしか分からないため、ここで先に検出する。
"""
import ast
import inspect
from pathlib import Path

import pytest

from outlook_cli.base import OutlookBase
from outlook_cli.mock import OutlookMock

PKG = Path(__file__).resolve().parent.parent / "outlook_cli"


def _abstract_signatures() -> dict:
    sigs = {}
    for name, member in inspect.getmembers(OutlookBase, predicate=inspect.isfunction):
        if getattr(member, "__isabstractmethod__", False):
            sigs[name] = list(inspect.signature(member).parameters)
    return sigs


def _class_methods(source_path: Path, class_name: str) -> dict:
    tree = ast.parse(source_path.read_text(encoding="utf-8-sig"))
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == class_name:
            methods = {}
            for item in node.body:
                if isinstance(item, ast.FunctionDef):
                    args = item.args
                    names = [a.arg for a in args.posonlyargs + args.args + args.kwonlyargs]
                    methods[item.name] = names
            return methods
    raise AssertionError(f"{class_name} が {source_path} に見つかりません")


ABSTRACT = _abstract_signatures()


def test_abstract_methods_are_discovered():
    assert {"send", "move_mail", "save_mail_attachment", "create_event", "forward"} <= set(ABSTRACT)


def test_mock_can_be_instantiated():
    OutlookMock()          # 抽象メソッドの実装漏れがあれば TypeError


@pytest.mark.parametrize("cls_file,cls_name", [("mock.py", "OutlookMock"), ("com.py", "OutlookCOM")])
def test_every_abstract_method_is_implemented_with_same_parameters(cls_file, cls_name):
    methods = _class_methods(PKG / cls_file, cls_name)
    missing = sorted(set(ABSTRACT) - set(methods))
    assert not missing, f"{cls_name} に未実装のメソッド: {missing}"
    mismatched = {
        name: (ABSTRACT[name], methods[name])
        for name in ABSTRACT if methods[name] != ABSTRACT[name]
    }
    assert not mismatched, f"{cls_name} の引数が base と違う: {mismatched}"
