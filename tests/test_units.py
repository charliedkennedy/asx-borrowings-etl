from src.borrowings_extractor import normalise_to_millions


def test_normalise_thousands():
    assert normalise_to_millions("294100", "$000") == 294.1


def test_normalise_millions():
    assert normalise_to_millions("294.1", "$m") == 294.1
