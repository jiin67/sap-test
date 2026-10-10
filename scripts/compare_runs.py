# compare_runs.py : pair today's logs (tag prefix D14_) with yesterday's (no prefix) and
# show the business-relevant facts side by side: status bars, titles, button texts,
# grid rows (S17) and EV_RESULT / ET_RETURN (S14).
#   usage: python compare_runs.py [tag-substring ...]
import io, os, re, sys, glob

LOGS = r"C:\sap-test\logs"
PFX = "D14_"


def read(p):
    if not os.path.exists(p):
        return None
    raw = open(p, "rb").read()
    for enc in ("utf-16", "utf-8"):
        try:
            return raw.decode(enc).replace("\r", "")
        except Exception:
            pass
    return raw.decode("latin-1")


def facts_s17(txt):
    out = []
    for l in txt.split("\n"):
        s = l.strip()
        if s.startswith("after F8:"):
            out.append("F8  " + re.sub(r"prog=\S+ ", "", s.split("after F8: ", 1)[1]))
        elif s.startswith("==== STEP"):
            out.append(s)
        elif s.startswith("press "):
            out.append("  " + s)
        elif re.match(r"after step \d+:", s):
            m = re.search(r"title=\[(.*?)\] sbar=(.*)$", s)
            if m:
                out.append("  -> title=[{}] sbar={}".format(m.group(1), m.group(2).strip()))
        elif s.startswith("(popup open:"):
            out.append("  " + s)
        elif s.startswith("#ROW") or s.startswith("#COLS"):
            out.append("  " + s)
        elif "GuiTextField wnd[1]/usr/txtSHKONTEXT-PFKEY" in s or "txtSHKONTEXT-PFKEY" in s:
            out.append("  GUI상태 " + s)
        elif s.startswith("GuiTextField wnd[1]") or s.startswith("GuiLabel wnd[1]/usr/lbl[") :
            if "STATUS_INFO" not in s and "SHKONTEXT" not in s:
                out.append("  POP " + s)
        elif s.startswith("exit code="):
            out.append(s)
    return out


def facts_s14(txt):
    out = []
    grab = False
    for l in txt.split("\n"):
        s = l.strip()
        if "EV_RESULT" in s or "ET_RETURN [" in s:
            out.append(s)
        if s.startswith("---- ET_RETURN"):
            grab = True
            continue
        if grab:
            if s.startswith("y") and ":" in s:
                if "엔트리" in s or "CSLTNO [" in s:
                    continue
                out.append("  RET " + s)
            elif s.startswith("----") or s.startswith("exit") or s == "":
                if s.startswith("exit"):
                    out.append(s)
                grab = False
        if s.startswith("exit code") or s.startswith("exit="):
            out.append(s)
    return out


def norm(lines):
    # drop volatile bits (session ids, execution time, dates of today vs yesterday)
    res = []
    for l in lines:
        l = re.sub(r"실행시간: \S+", "", l)
        res.append(l)
    return res


def main():
    filt = sys.argv[1:]
    today = sorted(glob.glob(os.path.join(LOGS, "S1[47]_" + PFX + "*.txt")), key=os.path.getmtime)
    for tp in today:
        if tp.endswith("_utf8.txt"):
            continue
        base = os.path.basename(tp)
        tag = base.split(PFX, 1)[1][:-4]
        if filt and not any(f in tag for f in filt):
            continue
        kind = base[:3]
        yp = os.path.join(LOGS, "{}_{}.txt".format(kind, tag))
        t_txt, y_txt = read(tp), read(yp)
        fx = facts_s17 if kind == "S17" else facts_s14
        t_f = norm(fx(t_txt)) if t_txt else ["(no log)"]
        y_f = norm(fx(y_txt)) if y_txt else ["(no yesterday log)"]
        same = t_f == y_f
        print("=" * 100)
        print("{} {}   {}".format(kind, tag, "SAME AS 09-13" if same else "DIFFERENT"))
        if same:
            for l in t_f:
                if l.strip().startswith(("#ROW", "EV_RESULT", "RET", "-> title", "F8", "GUI")):
                    print("   " + l)
        else:
            n = max(len(t_f), len(y_f))
            for i in range(n):
                a = y_f[i] if i < len(y_f) else ""
                b = t_f[i] if i < len(t_f) else ""
                mark = "  " if a == b else "**"
                if a == b:
                    if b.strip().startswith(("#ROW", "EV_RESULT", "RET", "-> title", "F8")):
                        print(" {} {}".format(mark, b))
                else:
                    print(" {} 09-13: {}".format(mark, a))
                    print(" {} 09-14: {}".format(mark, b))


if __name__ == "__main__":
    main()
