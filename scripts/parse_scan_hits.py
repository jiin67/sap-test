# parse_scan_hits.py : normalize S18 SCAN logs (old "yN:" and new "aN:" formats) into "OBJ LINE TEXT"
import re, sys
def hits(path):
    t = open(path, "rb").read().decode("utf-16").replace("\r", "")
    obj, out = "", []
    for l in t.split("\n"):
        m = re.match(r"\s*[a-z]?\d*:?\s*\[1\]\d+ \[4\](\S+)", l)
        if m:
            obj = m.group(1); continue
        m = re.match(r"\s*[a-z]?\d*:?\s*\[4\](\d{6})(?: \[11\](.*))?$", l)
        if m and m.group(2) is not None:
            k = (obj, m.group(1), m.group(2).rstrip())
            if k not in out: out.append(k)
    return out
if __name__ == "__main__":
    for obj, ln, tx in hits(sys.argv[1]):
        if len(sys.argv) < 3 or re.search(sys.argv[2], tx, re.I):
            print(obj, ln, tx)
