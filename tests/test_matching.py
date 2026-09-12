from excel_ops.matching import (
    Destination,
    preallocate_records,
    review_results,
    stable_record_id,
)
from excel_ops.models import ExtractedRecord


def record(location: str, *, identifier: str = "DEMO-1") -> ExtractedRecord:
    return ExtractedRecord(
        location=location,
        event_date="2026-09-12",
        identifier=identifier,
        category="inspection",
        confidence=0.99,
        source="synthetic.xlsx:Sheet1!A2",
    )


DESTINATIONS = (
    Destination("Toronto", ("Toronto Office", "Toronto, ON")),
    Destination("London", ("London Office", "London, ON")),
)


def test_matching_order_is_explainable():
    results = preallocate_records(
        [record("Toronto Office"), record("LONDON OFFICE", identifier="DEMO-2"), record("Toronto—ON", identifier="DEMO-3")],
        DESTINATIONS,
    )

    assert [(item.destination_key, item.rule) for item in results] == [
        ("Toronto", "configured_exact"),
        ("London", "casefold_exact"),
        ("Toronto", "normalized_exact"),
    ]


def test_strict_mode_never_accepts_fuzzy_match():
    strict = preallocate_records([record("Tornto Office")], DESTINATIONS)
    suggested = preallocate_records([record("Tornto Office")], DESTINATIONS, strict=False)

    assert strict[0].status == "unmatched"
    assert strict[0].candidates == ()
    assert suggested[0].status == "review"
    assert suggested[0].destination_key is None
    assert suggested[0].candidates[0].destination_key == "Toronto"
    assert suggested[0].rule == "fuzzy_candidate"


def test_multiple_aliases_for_one_destination_allocate_once():
    destinations = (Destination("Toronto", ("Toronto", "Toronto")),)
    result = preallocate_records([record("Toronto")], destinations)[0]

    assert result.status == "matched"
    assert result.destination_key == "Toronto"
    assert len(result.candidates) == 1


def test_alias_collision_between_destinations_is_a_conflict():
    destinations = (
        Destination("North", ("Shared Site",)),
        Destination("South", ("Shared Site",)),
    )
    result = preallocate_records([record("Shared Site")], destinations)[0]

    assert result.status == "conflict"
    assert result.destination_key is None
    assert {item.destination_key for item in result.candidates} == {"North", "South"}
    assert result in review_results([result])


def test_duplicate_records_are_not_allocated_twice():
    duplicate = record("Toronto Office")
    first, second = preallocate_records([duplicate, duplicate], DESTINATIONS)

    assert first.status == "matched"
    assert second.status == "duplicate"
    assert second.reasons == ("duplicate_record_id",)


def test_human_confirmation_uses_stable_record_id_and_is_reusable():
    uncertain = record("T.O. HQ")
    record_id = stable_record_id(uncertain)

    result = preallocate_records(
        [uncertain],
        DESTINATIONS,
        confirmations={record_id: "Toronto"},
    )[0]

    assert result.record_id == record_id
    assert result.status == "matched"
    assert result.destination_key == "Toronto"
    assert result.rule == "human_confirmation"


def test_missing_confirmation_target_returns_to_review():
    uncertain = record("Old office")
    result = preallocate_records(
        [uncertain],
        DESTINATIONS,
        confirmations={stable_record_id(uncertain): "Removed destination"},
    )[0]

    assert result.status == "review"
    assert result.reasons == ("confirmation_target_missing",)


def test_duplicate_destination_keys_are_rejected():
    destinations = (Destination("Toronto", ()), Destination("Toronto", ("T.O.",)))

    try:
        preallocate_records([record("Toronto")], destinations)
    except ValueError as exc:
        assert "unique" in str(exc)
    else:
        raise AssertionError("duplicate destination keys should be rejected")
