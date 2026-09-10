from __future__ import annotations
import argparse, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

HERE=Path(__file__).resolve().parent
PROCESSOR=HERE/'process_datasets_v4_1_3.py'
BUILDER=HERE/'build_auto_pricing_v4_1_3.py'

def unzip(src:Path,dst:Path):
    dst.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(src,'r') as z:z.extractall(dst)

def main():
    ap=argparse.ArgumentParser(description='End-to-end Dataset Finalizer v4.1.3 + formatted Auto Pricing workbook')
    ap.add_argument('--raw-zip',required=True)
    ap.add_argument('--input-zip',required=True,help='Previous finalized base files ZIP')
    ap.add_argument('--template',required=True,help='Auto Pricing workbook used for formatting/layout')
    ap.add_argument('--remove-party',help='Remove party.xlsx; when supplied it is authoritative')
    ap.add_argument('--sales-120-history',help='Optional replacement sales_120 history xlsx')
    ap.add_argument('--sales-180-history',help='Optional replacement sales_180 history xlsx')
    ap.add_argument('--output-dir',default='output_v4_1_3')
    ap.add_argument('--test-date',help='Regression only, YYYY-MM-DD. Omit in production.')
    a=ap.parse_args()
    out=Path(a.output_dir).resolve(); out.mkdir(parents=True,exist_ok=True)
    work=out/'_work'; raw=work/'raw'; hist=work/'history'
    if work.exists(): shutil.rmtree(work)
    unzip(Path(a.raw_zip),raw); unzip(Path(a.input_zip),hist)
    if a.sales_120_history: shutil.copy2(a.sales_120_history,hist/'sales_120_file.xlsx')
    if a.sales_180_history: shutil.copy2(a.sales_180_history,hist/'sales_180_file.xlsx')
    env=os.environ.copy(); env.update({'RAW_DATA_DIR':str(raw),'FINAL_HISTORY_DIR':str(hist),'OUTPUT_DIR':str(out),'COLUMN_MAPPING_FILE':str(HERE/'column_mapping_aliases.json')})
    if a.remove_party: env['REMOVE_PARTY_FILE']=str(Path(a.remove_party).resolve())
    if a.test_date: env['TEST_PROCESSING_DATE']=a.test_date
    subprocess.run([sys.executable,str(PROCESSOR)],env=env,check=True)
    date_tag=a.test_date or 'current'
    final_book=out/f'Auto_pricing_FINAL_{date_tag}_v4_1_3.xlsx'
    base_zip=out/f'Final_Base_Output_{date_tag}_v4_1_3.zip'
    subprocess.run([sys.executable,str(BUILDER),'--template',str(Path(a.template).resolve()),'--base-dir',str(out),'--output',str(final_book),'--base-zip',str(base_zip)],check=True)
    print('FINAL_WORKBOOK',final_book)
    print('BASE_ZIP',base_zip)
    print('AUDIT',out/'Processing_Audit.xlsx')

if __name__=='__main__': main()
