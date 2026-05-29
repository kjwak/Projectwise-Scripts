import json
import re
import sys

path = sys.argv[1]
needle = sys.argv[2] if len(sys.argv) > 2 else "movedTo\":\"failed"

for line in open(path, encoding="utf-8", errors="replace"):
    if "QC_PREPEND_LEGACY_FAILED" not in line:
        continue
    if needle not in line:
        continue
    o = json.loads(line)
    d = o["data"]
    print("jobId:", d.get("jobId"), "attempts:", d.get("attempts"))
    pd_raw = d.get("processorData") or ""
    # processorData is JSON string embedded in JSON - try parse with fixes
    try:
        pd = json.loads(pd_raw)
    except json.JSONDecodeError:
        # truncate at last complete key maybe
        m = re.search(r'"stdout":"', pd_raw)
        if m:
            print("stdout snippet (raw):", pd_raw[m.start() : m.start() + 3000])
        pd = {}
    print("exitCode:", pd.get("exitCode"))
    stdout = pd.get("stdout") or ""
    if stdout:
        print("--- stdout ---")
        print(stdout[-5000:])
    stderr = pd.get("stderr") or ""
    if stderr and "ScriptHalted" not in stderr:
        print("--- stderr ---")
        print(stderr[-2000:])
    break
