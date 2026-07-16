from outlook_mcp import friendly


def test_words_map_known_and_unknown_values():
    assert friendly.importance_word(2) == "high"
    assert friendly.importance_word(99) == "normal"  # unknown -> default
    assert friendly.response_word(3) == "accepted"
    assert friendly.response_word(5) == "not_responded"
    assert friendly.busy_status_word(0) == "free"
    assert friendly.busy_status_word(3) == "out_of_office"
    assert friendly.busy_status_word(99) == "busy"  # unknown -> default
    assert friendly.task_status_word(1) == "in_progress"
    assert friendly.task_status_word(99) == "not_started"


def test_reverse_lookups_are_case_insensitive_and_reject_garbage():
    assert friendly.busy_status_to_id("Out_Of_Office") == 3
    assert friendly.busy_status_to_id("nope") is None
    assert friendly.task_status_to_id("COMPLETE") == 2
    assert friendly.task_status_to_id("nope") is None
