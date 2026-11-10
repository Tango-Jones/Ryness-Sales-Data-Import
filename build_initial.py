"""One-time: build the starting database from the first PDF + v4 history."""
import sys, time
from ryness_parser import parse_pdf
import sales_db as sdb
pdf, legacy, out = sys.argv[1:4]
t=time.time()
db = sdb.empty_db()
res = parse_pdf(pdf)
import os
s = sdb.import_report(db, res, source_file=sys.argv[4] if len(sys.argv) > 4 else os.path.basename(pdf))
print({k: v for k, v in s.items() if k != "changes"})
rep, cmp_ = sdb.import_legacy(db, legacy, source_file=sys.argv[5] if len(sys.argv) > 5 else os.path.basename(legacy))
db["legacy_report"], db["legacy_compare"] = rep, cmp_
print(rep.method.value_counts())
sdb.save_db(db, out)
print("saved", out, round(time.time()-t,1), "s")
