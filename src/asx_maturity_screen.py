"""Local-only ASX debt maturity screen. It never performs HTTP requests."""
from __future__ import annotations
import argparse, csv, json, os, sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from .local_pdf_index import build_index, match_targets, read_targets

PILOT=("IDX","MTS","EVT","XRO","BSL","CHC","DMP","GNC","S32","TAH")
ACCEPTED={"MATCHED_HIGH","MATCHED_MEDIUM"}

def parser():
 p=argparse.ArgumentParser(description=__doc__); p.add_argument('--pdf-root',required=True); p.add_argument('--targets',required=True); p.add_argument('--output',default='outputs/asx_maturity_screen.xlsx'); p.add_argument('--pilot',action='store_true'); p.add_argument('--match-only',action='store_true'); return p

def write_csv(path, rows):
 path.parent.mkdir(parents=True,exist_ok=True)
 with path.open('w',newline='',encoding='utf-8') as f:
  w=csv.DictWriter(f,fieldnames=list(rows[0]) if rows else ['ticker']); w.writeheader(); w.writerows(rows)

def workbook(path, rows, register):
 wb=Workbook(); ws=wb.active; ws.title='Summary'; headers=['ticker','target name','match status','match confidence','status','selected PDF','match warning','extraction notes']
 ws.append(headers)
 for r in rows: ws.append([r.get(x,'') for x in headers])
 for c in range(1,len(headers)+1): ws.column_dimensions[get_column_letter(c)].width=24
 ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
 for name,data in [('Document Register',register),('Facilities',[]),('Buckets',[]),('Run Log',rows),('Exceptions',[r for r in rows if r['match status'] not in ACCEPTED])]:
  s=wb.create_sheet(name); h=list(data[0]) if data else ['ticker']; s.append(h)
  for r in data:s.append([r.get(x,'') for x in h])
  s.freeze_panes='A2'; s.auto_filter.ref=s.dimensions
 path.parent.mkdir(parents=True,exist_ok=True); wb.save(path)

def run_local(args):
 out=Path(args.output).parent; indexed=build_index(Path(args.pdf_root),out/'pdf_index.csv'); targets=read_targets(Path(args.targets))
 if args.pilot: targets=[t for t in targets if t['ticker'] in PILOT]
 candidates,register=match_targets(targets,indexed); write_csv(out/'document_match_candidates.csv',candidates); write_csv(out/'document_register.csv',register)
 if args.match_only: print(f'Local matching complete: {out / "document_register.csv"}'); return 0
 if not os.getenv('OPENAI_API_KEY'): print('OPENAI_API_KEY is required only after matching, before extraction.',file=sys.stderr); return 2
 # Documents are deliberately local-only. Extraction integration must consume these paths;
 # unconfirmed documents are represented, never sent to a model.
 rows=[]
 for r in register:
  status='PENDING_EXTRACTION' if r['match_status'] in ACCEPTED else r['match_status']
  rows.append({'ticker':r['ticker'],'target name':r['target_name'],'match status':r['match_status'],'match confidence':r['match_confidence'],'status':status,'selected PDF':r['selected_pdf'],'match warning':r['validation_warning'],'extraction notes':'Local PDF selected; extraction pending.' if status=='PENDING_EXTRACTION' else r['match_reason']})
 (out/'results.jsonl').write_text(''.join(json.dumps(r)+'\n' for r in rows),encoding='utf-8'); workbook(Path(args.output),rows,register); print(f'Workbook: {args.output}'); return 0

def main(): return run_local(parser().parse_args())
if __name__=='__main__': raise SystemExit(main())
