import re

from xl2y import patterns


def _full(p, s):
    return re.fullmatch(p, s) is not None


def test_email():
    assert _full(patterns.EMAIL, "a.b+c@example.com.au")
    assert not _full(patterns.EMAIL, "not an email")


def test_uuid():
    assert _full(patterns.UUID, "550e8400-e29b-41d4-a716-446655440000")
    assert not _full(patterns.UUID, "550e8400")


def test_ipv4():
    assert _full(patterns.IPV4, "192.168.0.1")
    assert not _full(patterns.IPV4, "999.1.1.1")


def test_phone_au():
    for ok in [
        "0412 345 678",
        "0412345678",
        "+61 412 345 678",
        "(02) 9123 4567",
    ]:
        assert _full(patterns.PHONE_AU, ok), ok
    assert not _full(patterns.PHONE_AU, "12345")


def test_abn_checksum():
    assert patterns.abn_valid("51 824 753 556")  # ATO example ABN
    assert not patterns.abn_valid("51 824 753 557")


def test_acn_checksum():
    assert patterns.acn_valid("004 085 616")  # ASIC example ACN
    assert not patterns.acn_valid("004 085 617")


def test_helpers():
    assert _full(patterns.digits(4), "1234") and not _full(
        patterns.digits(4), "12345"
    )
    assert _full(patterns.one_of("Y", "N"), "Y")
    assert _full(patterns.any_of(patterns.EMAIL, patterns.UUID), "a@b.co")
    assert _full(patterns.exact("a.b"), "a.b") and not _full(
        patterns.exact("a.b"), "axb"
    )
