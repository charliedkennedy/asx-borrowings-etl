from datetime import datetime
from src.maturity_mapper import map_date_to_bucket, allocate_annual_to_first_half


def test_date_bucket_mapping():
    assert map_date_to_bucket(datetime(2028, 5, 30)) == "Jun-28"
    assert map_date_to_bucket(datetime(2028, 12, 1)) == "Dec-28"


def test_annual_to_half_bucket():
    assert allocate_annual_to_first_half(2028) == "Jun-28"
