from services.tracking_service import extract_revision_value


def test_extract_revision_value_number():
    assert extract_revision_value("00") == 0
    assert extract_revision_value("01") == 1
    assert extract_revision_value("Rev02") == 2


def test_extract_revision_value_letter():
    assert extract_revision_value("A") == 1001
    assert extract_revision_value("B") == 1002
