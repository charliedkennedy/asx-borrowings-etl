from datetime import datetime
from .schemas import MATURITY_BUCKETS


def map_date_to_bucket(date_obj: datetime) -> str:
    if date_obj.year > 2035:
        return ">2035"
    yy = str(date_obj.year)[-2:]
    return ("Jun-" if date_obj.month <= 6 else "Dec-") + yy


def allocate_annual_to_first_half(year: int) -> str:
    if year > 2035:
        return ">2035"
    return f"Jun-{str(year)[-2:]}"


def blank_maturity(value: str = "not disclosed"):
    return {b: value for b in MATURITY_BUCKETS}
