import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from openpyxl import load_workbook
import model as M

HERE = os.path.dirname(os.path.dirname(__file__))


def test_excel_matches_python_base_case():
    """The saved workbook (recalculated in LibreOffice) must agree with the Python mirror."""
    wb = load_workbook(os.path.join(HERE, "ARR_Growth_Model.xlsx"), data_only=True)
    s = wb["Summary"]
    assert s["B2"].value == "Base"
    _, py = M.run(M.inputs("Base"))
    assert abs(s["B3"].value - py["arr_m12"]) < 1
    assert abs(s["B4"].value - py["arr_m24"]) < 1
    assert s["B5"].value == py["month_target_hit"]
    assert abs(s["B6"].value - py["min_cash"]) < 1


def test_scenarios_are_ordered():
    a = [M.run(M.inputs(x))[1]["arr_m24"] for x in ("Bear", "Base", "Bull")]
    assert a[0] < a[1] < a[2]


def test_policy_book_never_negative():
    df, _ = M.run(M.inputs("Bear"))
    assert (df.active > 0).all()
