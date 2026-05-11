import csv


def test_ticker_never_dropped(tmp_path):
    f = tmp_path / "tickers.csv"
    f.write_text("ticker,company_name\nAAA,AAA Ltd\nBBB,BBB Ltd\n")
    with open(f, newline='', encoding='utf-8') as h:
        rows = list(csv.DictReader(h))
    assert [r['ticker'] for r in rows] == ["AAA", "BBB"]
