"""Regression tests for str_term_validator.py.

The core guarantee under test: every revised term is a fixpoint (re-validating
it changes nothing) and, unless flagged NEEDS REVIEW, a second validation run
reports no ERROR - the exact failure mode where a "revised" term was itself
still invalid.
"""
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import str_term_validator as v  # noqa: E402

v.SPELLCHECK_ENABLED = False     # keep tests independent of pyspellchecker


def codes(result):
    return {i.code for i in result.issues}


def load(name):
    return v.read_terms(os.path.join(ROOT, "examples", name))


# --------------------------------------------------------------------------
# Specific repairs
# --------------------------------------------------------------------------

@pytest.mark.parametrize("raw, expected", [
    ("apple, pear, banana", "apple OR pear OR banana"),
    ("*fraud*", "fraud*"),
    ("contract AND AND breach", "contract AND breach"),
    ("AND settlement", "settlement"),
    ("settlement OR", "settlement"),
    ("(missing close", "(missing close)"),
    ('"unclosed phrase', '"unclosed phrase"'),
    ("invoice w/ 5 payment", "invoice W/5 payment"),
    ("invoice w/ payment", "invoice W/5 payment"),
    ("a w/7/ b", "a W/7 b"),
    ("apple w/0 pear", "apple W/1 pear"),
    ('"data breach" AND w/10 "notice"', '"data breach" W/10 "notice"'),
    ('"a b" NOT OR "c d"', '"a b" OR "c d"'),
    ("~apply", "apply~"),
    ("proj#ect", "project"),
    ("term ()", "term"),
    ("app%le", "apple"),
    ("deal OR deal", "deal"),
    ("market analysis", '"market analysis"'),
    ("e-mail", '("e-mail" OR email OR "e mail")'),
    ("merger AND acquisition OR takeover",
     "merger AND (acquisition OR takeover)"),
    ("“smart quoted phrase”", '"smart quoted phrase"'),
    ('"confidential"~~ AND report', '"confidential" AND report'),
    ("report.doc**", "report.doc*"),
    ('"x y" AND NOT stemming()', '"x y"'),
])
def test_repairs(raw, expected):
    assert v.validate_full(1, raw).revised == expected


def test_assumptions_are_labelled():
    r = v.validate_full(1, "invoice w/ payment")
    assumed = [i for i in r.issues if "ASSUMPTION APPLIED" in i.reasoning]
    assert assumed and assumed[0].code == "PROX_DISTANCE_ASSUMED"


# --------------------------------------------------------------------------
# Metadata field:value extraction
# --------------------------------------------------------------------------

def test_field_conditions_are_extracted_and_validated():
    r = v.validate_full(
        1, 'custodian:Smith AND "wire transfer" AND date:2024/31/02 '
           'AND ext:xlxs AND auther:"Jones"')
    assert r.revised == '"wire transfer"'
    assert [c["field"] for c in r.conditions] == \
        ["custodian", "date", "ext", "auther"]
    notes = " | ".join(c["note"] for c in r.conditions)
    assert "not a possible calendar date" in notes
    assert "'xlsx'" in notes            # transposition beats the shorter 'xls'
    assert "'author'" in notes
    assert "FIELD_SYNTAX_UNSUPPORTED" in codes(r)


def test_date_range_end_is_checked():
    r = v.validate_full(1, '"audit" AND date:2024-01-01 TO 2024-02-40')
    assert len(r.conditions) == 1
    assert "2024-02-40" in r.conditions[0]["note"]


@pytest.mark.parametrize("date, ok", [
    ("2024/02/29", True), ("2023/02/29", False), ("2024-31-12", True),
    ("2024-13-13", False),
])
def test_calendar_dates(date, ok):
    assert (v._check_date_value(date) is None) == ok


def test_quoted_colon_is_not_a_field():
    r = v.validate_full(1, '"re: budget meeting"')
    assert r.conditions == []


def test_url_is_not_a_field():
    r = v.validate_full(1, "http://example.com")
    assert r.conditions == []


# --------------------------------------------------------------------------
# Things that must stay with a human
# --------------------------------------------------------------------------

@pytest.mark.parametrize("raw, code", [
    ("the", "NOISE_WORD_TERM"),
    ("12~~abc", "NUMERIC_RANGE_MALFORMED"),
    ("a" * 460, "OVER_450_CHARS"),
])
def test_needs_review(raw, code):
    r = v.validate_full(1, raw)
    assert r.residual.startswith("NEEDS REVIEW")
    assert code in r.residual


def test_needs_review_terms_are_excluded_from_clean_file(tmp_path):
    results = v.validate_all([(1, "fraud"), (2, "the"), (3, "fraud")])
    out = tmp_path / "clean.txt"
    assert v.write_clean_txt(results, str(out)) == 1
    assert out.read_text(encoding="utf-8").splitlines() == ["fraud"]


# --------------------------------------------------------------------------
# The core guarantee: fixpoint + no errors on a second pass
# --------------------------------------------------------------------------

ALL_TERMS = [t for _, t in load("sample_terms.txt") + load("complex_terms.txt")
             if t.strip()]


@pytest.mark.parametrize("raw", ALL_TERMS, ids=lambda t: t[:30])
def test_revised_term_is_stable_and_error_free(raw):
    first = v.validate_full(1, raw)
    if first.residual.startswith("NEEDS REVIEW"):
        pytest.skip("deliberately left for a human")
    second = v.validate_full(1, first.revised)
    assert second.revised == first.revised, "revision is not a fixpoint"
    assert not [i for i in second.issues if i.severity == "ERROR"]


def test_complex_terms_all_become_runnable():
    for _, raw in load("complex_terms.txt"):
        r = v.validate_full(1, raw)
        assert not r.residual.startswith("NEEDS REVIEW"), r.residual


# --------------------------------------------------------------------------
# I/O
# --------------------------------------------------------------------------

def test_report_can_be_fed_back_in(tmp_path):
    results = v.validate_all(load("complex_terms.txt"))
    report = tmp_path / "report.xlsx"
    v.write_excel(results, str(report), "complex_terms.txt")
    reread = [t for _, t in v.read_terms(str(report))]
    assert reread == [r.revised for r in results]   # the Revised Term column


def test_control_characters_do_not_break_the_report(tmp_path):
    results = v.validate_all([(1, "alpha\x14beta"), (2, "gamma\x07")])
    v.write_excel(results, str(tmp_path / "r.xlsx"), "x")
    assert results[0].revised == '"alpha beta"'


def test_cli_end_to_end(tmp_path):
    report, clean = tmp_path / "r.xlsx", tmp_path / "c.txt"
    v.main([os.path.join(ROOT, "examples", "complex_terms.txt"),
            "-o", str(report), "--clean", str(clean), "--no-spellcheck"])
    assert report.exists()
    assert len(clean.read_text(encoding="utf-8").splitlines()) == 4
