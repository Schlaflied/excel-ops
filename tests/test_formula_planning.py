import pytest

from excel_ops.formula_planning import (
    ConditionalSumFormulaSpec,
    DateAddFormulaSpec,
    FormulaPlanError,
    LookupFormulaSpec,
    plan_formula,
)


def test_modern_lookup_uses_xlookup():
    plan = plan_formula(
        LookupFormulaSpec("Find each employee rate.", "A2", "Rate Card", "$A$2:$A$20", "$B$2:$B$20"),
        target_excel_version="365",
    )

    assert plan.function == "XLOOKUP"
    assert plan.value == '=XLOOKUP(A2,\'Rate Card\'!$A$2:$A$20,\'Rate Card\'!$B$2:$B$20,"",0)'
    assert plan.requires_independent_recalculation is True


@pytest.mark.parametrize("version", ["2016", "2019"])
def test_legacy_lookup_uses_index_match(version):
    plan = plan_formula(
        LookupFormulaSpec("Find each employee rate.", "A2", "Rates", "A2:A20", "B2:B20"),
        target_excel_version=version,
    )

    assert plan.function == "INDEX/MATCH"
    assert "XLOOKUP" not in plan.value
    assert plan.value.startswith("=IFNA(INDEX(")


def test_conditional_sum_is_cross_sheet_and_compatible():
    plan = plan_formula(
        ConditionalSumFormulaSpec("Total approved hours.", "A2", "Time Data", "A2:A100", "D2:D100"),
        target_excel_version="2016",
    )

    assert plan.value == "=SUMIFS('Time Data'!D2:D100,'Time Data'!A2:A100,A2)"
    assert plan.compatibility_strategy == "SUMIFS is supported by all target versions"


def test_conditional_sum_accepts_equal_two_dimensional_ranges():
    plan = plan_formula(
        ConditionalSumFormulaSpec(
            "Total approved values.", "A2", "Data", "a2:$b$10", "D2:E10"
        ),
        target_excel_version="365",
    )

    assert plan.value == "=SUMIFS('Data'!D2:E10,'Data'!A2:$B$10,A2)"


def test_conditional_sum_rejects_mismatched_shapes():
    with pytest.raises(FormulaPlanError, match="same dimensions"):
        plan_formula(
            ConditionalSumFormulaSpec(
                "Total approved values.", "A2", "Data", "A2:B10", "D2:D10"
            ),
            target_excel_version="365",
        )


def test_horizontal_lookup_allows_multiple_return_rows():
    plan = plan_formula(
        LookupFormulaSpec("Find a month.", "A2", "Data", "B1:M1", "B2:M4"),
        target_excel_version="365",
    )

    assert plan.function == "XLOOKUP"


def test_legacy_horizontal_lookup_uses_match_as_index_column():
    plan = plan_formula(
        LookupFormulaSpec("Find a month.", "A2", "Data", "B1:M1", "B2:M2"),
        target_excel_version="2019",
    )

    assert plan.value == (
        "=IFNA(INDEX('Data'!B2:M2,0,MATCH(A2,'Data'!B1:M1,0)),\"\")"
    )


@pytest.mark.parametrize(
    ("lookup_range", "return_range"),
    [("A2:A10", "C2:D10"), ("B1:M1", "B2:M4")],
)
def test_legacy_lookup_rejects_multi_cell_returns(lookup_range, return_range):
    with pytest.raises(FormulaPlanError, match="legacy lookup return_range"):
        plan_formula(
            LookupFormulaSpec(
                "Find a value.", "A2", "Data", lookup_range, return_range
            ),
            target_excel_version="2019",
        )


@pytest.mark.parametrize(
    ("lookup_range", "return_range", "message"),
    [
        ("A2:B10", "C2:C10", "single row or column"),
        ("A2:A10", "C2:C9", "same number of rows"),
        ("B1:M1", "B2:L4", "same number of columns"),
    ],
)
def test_lookup_rejects_ambiguous_or_mismatched_shapes(
    lookup_range, return_range, message
):
    with pytest.raises(FormulaPlanError, match=message):
        plan_formula(
            LookupFormulaSpec(
                "Find a value.", "A2", "Data", lookup_range, return_range
            ),
            target_excel_version="365",
        )


def test_date_plan_uses_explicit_month_offset():
    plan = plan_formula(
        DateAddFormulaSpec("Move the renewal date forward one month.", "C2", 1),
        target_excel_version="2021",
    )

    assert plan.value == "=EDATE(C2,1)"


def test_static_mode_requires_an_independent_value():
    spec = DateAddFormulaSpec("Use the approved renewal date.", "C2", 1)

    with pytest.raises(FormulaPlanError, match="independently computed"):
        plan_formula(spec, target_excel_version="365", output_mode="static")

    plan = plan_formula(
        spec,
        target_excel_version="365",
        output_mode="static",
        computed_value="2026-10-23",
    )
    assert plan.value == "2026-10-23"
    assert plan.function is None
    assert plan.requires_independent_recalculation is False


@pytest.mark.parametrize(
    ("spec", "message"),
    [
        (LookupFormulaSpec("Lookup.", "A2+1", "Rates", "A2:A3", "B2:B3"), "lookup_cell"),
        (ConditionalSumFormulaSpec("Sum.", "A2", "Bad/Sheet", "A2:A3", "B2:B3"), "source_sheet"),
        (DateAddFormulaSpec("Date.", "A2", True), "months"),
    ],
)
def test_unsafe_or_ambiguous_inputs_are_rejected(spec, message):
    with pytest.raises(FormulaPlanError, match=message):
        plan_formula(spec, target_excel_version="365")


def test_sheet_quotes_are_escaped():
    plan = plan_formula(
        LookupFormulaSpec("Lookup.", "A2", "Director's Rates", "A2:A3", "B2:B3"),
        target_excel_version="365",
    )

    assert "'Director''s Rates'" in plan.value
