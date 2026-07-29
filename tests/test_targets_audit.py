import csv
import re
from pathlib import Path


def test_targets_are_normalized_unique_and_preserve_known_aliases():
    rows = list(csv.DictReader(Path("targets.csv").open(encoding="utf-8-sig")))
    identifiers = [row["ticker"] for row in rows]
    assert len(rows) == len(set(identifiers)) == 289
    assert all(identifier == identifier.upper() for identifier in identifiers)
    assert all(re.fullmatch(r"[A-Z0-9]{2,5}", identifier) for identifier in identifiers)
    by_ticker = {row["ticker"]: row for row in rows}
    assert "Star Entertainment" in by_ticker["SGR"]["aliases"]
    assert "Shopping Centres Australasia" in by_ticker["RGN"]["aliases"]
    assert "Acrow" in by_ticker["ACF"]["aliases"]
    for ticker in ("PEN", "RRC", "TSS", "VRN", "WON", "SCO"):
        assert ticker in by_ticker
