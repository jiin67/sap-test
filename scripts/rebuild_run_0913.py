# rebuild_run_0913.py : reconstruct the exact command sequence of the 2026-09-13 통합테스트
# (v0.04) from the UTF-16 logs, so that it can be replayed today.
import io, os, re, glob, datetime

LOGS = r"C:\sap-test\logs"
OUT = r"C:\sap-test\results\_work\run_0913_rebuilt.txt"
T0 = datetime.datetime(2026, 9, 13, 17, 0)
T1 = datetime.datetime(2026, 9, 13, 19, 10)

files = []
for p in glob.glob(os.path.join(LOGS, "S1[3-9]_*.txt")) + glob.glob(os.path.join(LOGS, "S20_*.txt")):
    if p.endswith("_utf8.txt"):
        continue
    mt = datetime.datetime.fromtimestamp(os.path.getmtime(p))
    if T0 <= mt <= T1:
        files.append((mt, p))
files.sort()


def read(p):
    raw = open(p, "rb").read()
    for enc in ("utf-16", "utf-8", "cp949"):
        try:
            return raw.decode(enc)
        except Exception:
            pass
    return raw.decode("latin-1")


out = []
for mt, p in files:
    name = os.path.basename(p)
    txt = read(p).replace("\r", "")
    lines = txt.split("\n")
    head = lines[0].lstrip("\ufeff")
    if name.startswith("S17_"):
        m = re.match(r"tag=(\S+) tcode=(\S+) kunnr=(\S+) spmon=(\S+) vkbur=(\S+) mode=(\S+)", head)
        steps = []
        for l in lines:
            s = re.match(r"==== STEP \d+ (\S+) ?(.*)$", l)
            if s:
                typ, arg = s.group(1), s.group(2).strip()
                steps.append(typ if arg == "" else "{}:{}".format(typ, arg))
        # exit code
        ex = [l for l in lines if l.startswith("exit code=")]
        if m:
            out.append("{} S17 {} {} {} {} {} {} {}   # {}".format(
                mt.strftime("%H:%M"), m.group(1), m.group(2), m.group(3), m.group(4), m.group(5),
                m.group(6), " ".join(steps), ex[-1] if ex else ""))
        else:
            out.append("{} S17 ?? {}".format(mt.strftime("%H:%M"), head))
    elif name.startswith("S14_"):
        m = re.match(r"tag=(\S+) IV_ACTION=\[(.*?)\] items=(\d+)", head)
        items = re.findall(r"IT_DATA\[\d+\] CSLTNO=\[(.*?)\] CFDATE1=\[(.*?)\]", txt)
        res = [l.strip() for l in lines if "EV_RESULT" in l][:1]
        act = m.group(2) if m and m.group(2) != "" else "_"
        out.append("{} S14 {} {} {}   # {}".format(
            mt.strftime("%H:%M"), m.group(1) if m else "?", act,
            " ".join("{}:{}".format(a, b) for a, b in items), res[0] if res else ""))
    elif name.startswith("S16_"):
        ops = [l.split("---------- ", 1)[1].strip() for l in lines if l.startswith("---------- ")]
        tag = re.match(r"tag=(\S+)", head)
        out.append("{} S16 {} {}".format(mt.strftime("%H:%M"), tag.group(1) if tag else "?", " ".join(ops)))
    elif name.startswith("S13_"):
        m = re.match(r"mode=(\S+) newval=\[(.*?)\]", head)
        cs = [l.split("CSLTNO ", 1)[1].strip() for l in lines if l.startswith("---------- CSLTNO ")]
        nv = m.group(2) if m and m.group(2) != "" else "_"
        out.append("{} S13 {} {} {}".format(mt.strftime("%H:%M"), m.group(1) if m else "?", nv, " ".join(cs)))
    elif name.startswith("S18_"):
        out.append("{} S18 {} | {}".format(mt.strftime("%H:%M"), name, " ".join(lines[:4])[:220]))
    else:
        out.append("{} {} | {}".format(mt.strftime("%H:%M"), name, head[:200]))

io.open(OUT, "w", encoding="utf-8").write("\n".join(out))
print("\n".join(out))
print("\nlines:", len(out), "->", OUT)
