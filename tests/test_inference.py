from excel_ops.inference import infer_fields, infer_value


def test_identifier_fields_preserve_leading_zeroes():
    result = infer_value("employee_id", "00123")
    assert result.inferred_type == "text_identifier"
    assert result.normalized_value == "00123"
    assert result.reason == "leading_zero_identifier"


def test_phone_and_postal_codes_are_identifiers():
    assert infer_value("phone", "+1 (519) 555-0100").inferred_type == "text_identifier"
    assert infer_value("postal_code", "02139").normalized_value == "02139"


def test_ambiguous_date_is_not_silently_resolved():
    result = infer_value("event_date", "03/04/2026")
    assert result.ambiguous is True
    assert result.normalized_value is None
    assert result.reason == "ambiguous_day_month"


def test_locale_resolves_common_date_formats():
    assert infer_value("event_date", "03/04/2026", "en-US").normalized_value == "2026-03-04"
    assert infer_value("event_date", "03/04/2026", "en-GB").normalized_value == "2026-04-03"
    assert infer_value("日期", "2026年4月3日").normalized_value == "2026-04-03"


def test_excel_serial_date_requires_date_context():
    result = infer_value("invoice_date", 46023)
    assert result.inferred_type == "excel_serial_date"
    assert result.normalized_value == "2026-01-01"
    assert infer_value("quantity", 46023).inferred_type == "integer"


def test_percentage_and_locale_numbers():
    percentage = infer_value("completion_rate", "12.5%", "en-US")
    assert percentage.inferred_type == "percentage"
    assert percentage.normalized_value == "0.125"
    german = infer_value("amount", "1.234,56 €", "de-DE")
    assert german.unit == "currency"
    assert german.normalized_value == "1234.56"


def test_field_report_exposes_low_confidence_and_ambiguity():
    report = infer_fields(
        [
            {"start_date": "03/04/2026", "employee_id": "001"},
            {"start_date": "04/05/2026", "employee_id": "002"},
        ]
    )
    by_field = {item.field: item for item in report}
    assert by_field["start_date"].ambiguous_count == 2
    assert by_field["start_date"].low_confidence_count == 2
    assert by_field["employee_id"].leading_zero_preserved is True
