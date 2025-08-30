from helpers import _num_to_2dec


def test_num_to_2dec_formats_comma_and_dot():
    assert _num_to_2dec("1,23") == "1.23"
    assert _num_to_2dec("1.23") == "1.23"


def test_num_to_2dec_exact_rounding():
    # Without Decimal this would yield '2.67'
    assert _num_to_2dec("2.675") == "2.68"


def test_num_to_2dec_invalid_returns_original():
    assert _num_to_2dec("abc") == "abc"
