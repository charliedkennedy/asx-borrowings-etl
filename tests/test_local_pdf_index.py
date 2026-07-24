from pathlib import Path
from src.local_pdf_index import normalise_name, parse_filename, match_targets

def row(name, text='Annual Report\nConsolidated Financial Statements\nIntegral Diagnostics'):
    return {'full_path':name,'filename':Path(name).name,'filename_company':parse_filename(Path(name).name)[0], 'financial_year':'2026','acn':'123456789','page_count':100,'first_pages_text':text,'normalised_filename_company':normalise_name(parse_filename(Path(name).name)[0]),'index_error':''}

def test_filename_and_normalisation():
    assert parse_filename('A.P. Eagers Ltd_FY2026_ACN123456789.Pdf') == ('A.P. Eagers Ltd','2026','123456789')
    assert normalise_name('A.P. Eagers & Co. Limited') == 'A P EAGERS AND CO'

def test_local_matched_and_missing_and_ambiguous():
    target={'ticker':'IDX','target_name':'Integral Diagnostics','aliases':['Integral Diagnostics'],'warning':''}
    _, register=match_targets([target],[row('INTEGRAL DIAGNOSTICS LIMITED_FY2026_ACN123456789.pdf')])
    assert register[0]['match_status']=='MATCHED_HIGH'
    _, missing=match_targets([target],[]); assert missing[0]['match_status']=='NO_LOCAL_DOCUMENT'
    _, ambiguous=match_targets([target],[row('INTEGRAL DIAGNOSTICS_FY2026_ACN123456789.pdf'),row('INTEGRAL DIAGNOSTICS_FY2026_ACN123456788.pdf')])
    assert ambiguous[0]['match_status']=='AMBIGUOUS_MATCH'
