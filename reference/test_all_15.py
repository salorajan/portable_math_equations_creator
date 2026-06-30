import subprocess
import os
import sys

# Define inputs and targets
inputs = {
    "md": "latex_testing.md",
    "html": "latex_testing.html",
    "docx": "latex_testing.docx"
}

formats = ["word", "html", "pdf", "epub", "md"]

python_bin = r"C:\salo\acb\quill\env_quill\Scripts\python.exe"
convert_script = "convert_md.py"

results = []

print("==================================================")
print("Testing all 15 document conversions...")
print("==================================================\n")

for in_fmt, in_file in inputs.items():
    if not os.path.exists(in_file):
        print(f"Error: Input file {in_file} does not exist!")
        sys.exit(1)
        
    for out_fmt in formats:
        # Determine output file name
        out_ext = {
            "word": ".docx",
            "html": ".html",
            "pdf": ".pdf",
            "epub": ".epub",
            "md": ".md"
        }[out_fmt]
        
        out_file = f"test_out_{in_fmt}_to_{out_fmt}{out_ext}"
        if os.path.exists(out_file):
            try:
                os.remove(out_file)
            except:
                pass
                
        cmd = [python_bin, convert_script, in_file, out_fmt, out_file]
        print(f"Running: {' '.join(cmd)}")
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
        
        ok = res.returncode == 0 and os.path.exists(out_file) and os.path.getsize(out_file) > 0
        status = "SUCCESS" if ok else "FAILED"
        
        err_msg = ""
        if not ok:
            err_msg = res.stderr or res.stdout or "Output file not found or empty."
            err_msg = err_msg.strip().replace("\n", " ")[:100]
            
        results.append({
            "input": in_file,
            "target": out_fmt,
            "status": status,
            "error": err_msg
        })
        
        # Clean up output file to keep directory clean
        if os.path.exists(out_file):
            try:
                os.remove(out_file)
            except:
                pass

print("\n" + "=" * 60)
print(f"{'INPUT FILE':<20} | {'TARGET':<8} | {'STATUS':<8} | {'ERROR / INFO':<50}")
print("-" * 100)
for r in results:
    print(f"{r['input']:<20} | {r['target']:<8} | {r['status']:<8} | {r['error']:<50}")
print("=" * 100)

# Check if all passed
all_ok = all(r["status"] == "SUCCESS" for r in results)
if all_ok:
    print("\nAll 15 conversions completed successfully!")
    sys.exit(0)
else:
    print("\nSome conversions failed!")
    sys.exit(1)
