from src.asx_maturity_screen import allocate_exact, allocate_bucket, HALVES

def test_exact_and_fy29():
 assert allocate_exact(10,'2027-12-01')['2H27']==10
 assert allocate_exact(10,'2030-01-01')['FY29+']==10

def test_bucket_is_inferred_and_does_not_make_zeroes():
 x=allocate_bucket(12,0,12); assert sum(v or 0 for v in x.values())==12
 assert set(x)==set(HALVES)
