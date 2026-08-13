from __future__ import annotations
import argparse, json, math, os, shutil, subprocess, sys, zipfile
from pathlib import Path

def run_command(command: list[str], cwd: Path, log_path: Path) -> None:
    command_text = " ".join(command)
    print("\n$ " + command_text, flush=True)
    with log_path.open("a", encoding="utf-8") as log:
        log.write("\n$ " + command_text + "\n")
        process = subprocess.Popen(
            command,
            cwd=str(cwd),
            env={**os.environ, "PYTHONHASHSEED": "0", "PYTHONUNBUFFERED": "1"},
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        assert process.stdout is not None
        tail: list[str] = []
        for line in process.stdout:
            print(line, end="", flush=True)
            log.write(line)
            log.flush()
            tail.append(line.rstrip("\n"))
            if len(tail) > 120:
                tail.pop(0)
        return_code = process.wait()
    if return_code != 0:
        diagnostic = log_path.parent / "RUN_FAILURE_DIAGNOSTIC_v8_2_2.txt"
        diagnostic.write_text(
            "DEALMAKER v8.2.2 RUN FAILURE\n"
            "=============================\n"
            f"Command: {command_text}\n"
            f"Exit code: {return_code}\n\n"
            "Last output lines:\n" + "\n".join(tail) + "\n",
            encoding="utf-8",
        )
        raise RuntimeError(
            f"Stage failed with exit code {return_code}. "
            f"See {diagnostic} and {log_path}."
        )

def main() -> None:
    ap=argparse.ArgumentParser(description="Memory-safe Dealmaker v8 production runner with sequential unique stock allocation.")
    ap.add_argument("--sales",required=True);ap.add_argument("--actions",required=True);ap.add_argument("--stock",required=True);ap.add_argument("--control",required=True)
    ap.add_argument("--output-dir",default="dealmaker_v8_2_output");ap.add_argument("--chunk-size",type=int,default=10)
    ap.add_argument("--sales-sheet",default="");ap.add_argument("--actions-sheet",default="");ap.add_argument("--stock-sheet",default="")
    ap.add_argument("--sales-header-row",type=int,default=None);ap.add_argument("--actions-header-row",type=int,default=None);ap.add_argument("--stock-header-row",type=int,default=None)
    args=ap.parse_args();here=Path(__file__).resolve().parent;output=Path(args.output_dir).resolve();work=output/".work";alloc=output/".allocation"
    if output.exists(): shutil.rmtree(output)
    work.mkdir(parents=True);alloc.mkdir(parents=True)
    run_log = output / "Production_Run_Log_v8_2_2.txt"
    run_log.write_text("DEALMAKER v8.2.2 PRODUCTION RUN LOG\n", encoding="utf-8")
    pipeline=here/"Dealmaker_Full_Pipeline_v6.py";interval=here/"Dealmaker_Step1_Interval_Engine_v10.py"
    preflight=[sys.executable,str(here/"Dealmaker_Production_Preflight_v6.py"),"--sales",str(Path(args.sales).resolve()),"--actions",str(Path(args.actions).resolve()),"--stock",str(Path(args.stock).resolve()),"--control",str(Path(args.control).resolve()),"--output",str(output/"Dealmaker_Data_Preflight_v6.xlsx")]
    for flag,value in [("--sales-sheet",args.sales_sheet),("--actions-sheet",args.actions_sheet),("--stock-sheet",args.stock_sheet)]:
        if value: preflight.extend([flag,value])
    for flag,value in [("--sales-header-row",args.sales_header_row),("--actions-header-row",args.actions_header_row),("--stock-header-row",args.stock_header_row)]:
        if value is not None: preflight.extend([flag,str(value)])
    run_command(preflight,here,run_log)
    stage1=[sys.executable,str(here/"Dealmaker_Stage1_Prepare_v6.py"),"--sales",str(Path(args.sales).resolve()),"--actions",str(Path(args.actions).resolve()),"--control",str(Path(args.control).resolve()),"--pipeline",str(pipeline),"--interval-engine",str(interval),"--work-dir",str(work)]
    for flag,value in [("--sales-sheet",args.sales_sheet),("--actions-sheet",args.actions_sheet)]:
        if value: stage1.extend([flag,value])
    for flag,value in [("--sales-header-row",args.sales_header_row),("--actions-header-row",args.actions_header_row)]:
        if value is not None: stage1.extend([flag,str(value)])
    run_command(stage1,here,run_log)
    stage1b=[sys.executable,str(here/"Dealmaker_Stage1B_Profiles_Stock_v6.py"),"--stock",str(Path(args.stock).resolve()),"--control",str(Path(args.control).resolve()),"--pipeline",str(pipeline),"--interval-engine",str(interval),"--work-dir",str(work)]
    if args.stock_sheet: stage1b.extend(["--stock-sheet",args.stock_sheet])
    if args.stock_header_row is not None: stage1b.extend(["--stock-header-row",str(args.stock_header_row)])
    run_command(stage1b,here,run_log)
    metadata=json.loads((work/"stage1_metadata.json").read_text());chunks=math.ceil(int(metadata["eligible_parties"])/args.chunk_size)
    for chunk in range(chunks):
        run_command([sys.executable,str(here/"Dealmaker_Stage2_Allocate_v6.py"),"--pipeline",str(pipeline),"--work-dir",str(work),"--allocation-dir",str(alloc),"--chunk",str(chunk),"--chunk-size",str(args.chunk_size)],here,run_log)
    run_command([sys.executable,str(here/"Dealmaker_Stage3_Finalize_v6.py"),"--pipeline",str(pipeline),"--work-dir",str(work),"--allocation-dir",str(alloc),"--output-dir",str(output)],here,run_log)
    run_command([sys.executable,str(here/"Dealmaker_Stage4_DealBooks_v6.py"),"--pipeline",str(pipeline),"--output-dir",str(output)],here,run_log)
    run_command([sys.executable,str(here/"Dealmaker_Release_Gate_Audit_v6.py"),"--output-dir",str(output),"--control",str(Path(args.control).resolve())],here,run_log)
    shutil.copy2(Path(args.control).resolve(), output/"Control_Workbook_Used.xlsx")
    review_zip=output/"Dealmaker_Production_Review_v8.zip"
    with zipfile.ZipFile(review_zip,"w",zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(output.iterdir()):
            if path.is_file() and path!=review_zip and path.suffix!=".pkl": zf.write(path,arcname=path.name)
    shutil.rmtree(work,ignore_errors=True);shutil.rmtree(alloc,ignore_errors=True)
    print(f"\nProduction review bundle: {review_zip}",flush=True)

if __name__=="__main__":main()
