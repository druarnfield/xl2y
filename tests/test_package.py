def test_package_imports():
    import xl2y

    assert xl2y.__all__  # public API declared


def test_error_hierarchy():
    from xl2y.errors import (
        EmptySheetError,
        SchemaError,
        SheetNotFoundError,
        UnsupportedFormatError,
        Xl2yError,
    )

    for exc in (
        UnsupportedFormatError,
        EmptySheetError,
        SheetNotFoundError,
        SchemaError,
    ):
        assert issubclass(exc, Xl2yError)
