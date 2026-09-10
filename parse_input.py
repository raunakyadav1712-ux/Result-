"""Parse pasted student results into the generator's JSON student list.
Handles one-line-per-subject AND multi-line (each attribute on its own line)."""
import re

SUBJECTS = ("physics", "chemistry", "maths", "math", "biology", "botany",
            "zoology", "science", "sst")


def _canon(word):
    w = word.strip().lower()
    return "Maths" if w in ("math", "maths") else word.strip().title()


def _subject_head(line):
    m = re.match(r'\s*([A-Za-z]+)\b(.*)$', line)
    if not m:
        return None
    head = m.group(1).lower()
    for s in SUBJECTS:
        if head == s or head.startswith(s) or s.startswith(head):
            return _canon(m.group(1)), m.group(2)
    return None


def _nums_labelled(text):
    c = re.search(r'\b(?:c|correct)\s*[:=\-]?\s*(-?\d+)', text, re.I)
    w = re.search(r'\b(?:w|wrong|incorrect|inc)\s*[:=\-]?\s*(-?\d+)', text, re.I)
    u = re.search(r'\b(?:u|l|left|unattempt(?:ed)?|notattempt(?:ed)?)\s*[:=\-]?\s*(-?\d+)',
                  text, re.I)
    if c or w:
        return (int(c.group(1)) if c else 0,
                int(w.group(1)) if w else 0,
                int(u.group(1)) if u else None)
    return None


def _attr(line):
    l = line.lower()
    n = re.search(r'-?\d+', line)
    if not n:
        return None
    v = int(n.group(0))
    if "incorrect" in l or "wrong" in l:
        return "incorrect", v
    if "unattempt" in l or "notattempt" in l or "left" in l:
        return "left", v
    if "correct" in l:
        return "correct", v
    if "mark" in l:
        return "marks", v
    return None


def _apply_inline(sub, rest):
    lab = _nums_labelled(rest)
    if lab:
        sub["correct"], sub["incorrect"] = lab[0], lab[1]
        if lab[2] is not None:
            sub["left"] = lab[2]
        return
    nums = [int(x) for x in re.findall(r'-?\d+', rest)]
    if re.search("mark", rest, re.I) and nums:
        nums = nums[1:]
    if len(nums) >= 1:
        sub["correct"] = nums[0]
    if len(nums) >= 2:
        sub["incorrect"] = nums[1]
    if len(nums) >= 3:
        sub["left"] = nums[2]


def parse_students(text):
    students, cur, sub = [], None, None

    def ensure_student():
        nonlocal cur
        if cur is None:
            cur = {"name": f"Student {len(students)+1}", "subjects": []}
            students.append(cur)

    for raw in text.splitlines():
        line = raw.strip().strip("*").strip()
        if not line:
            continue
        head = _subject_head(line)
        if head:
            ensure_student()
            sub = {"name": head[0], "correct": 0, "incorrect": 0, "left": None}
            cur["subjects"].append(sub)
            _apply_inline(sub, head[1])
            continue
        attr = _attr(line)
        if attr and sub is not None:
            key, val = attr
            if key in ("correct", "incorrect", "left"):
                sub[key] = val
            continue
        absent = bool(re.search(r'\babsent\b', line, re.I))
        prev = re.search(r'\bprev(?:ious)?\s*[:=\-]?\s*(\d+)', line, re.I)
        name = re.sub(r'\babsent\b', '', line, flags=re.I)
        name = re.sub(r'\bprev(?:ious)?\s*[:=\-]?\s*\d+', '', name, flags=re.I)
        name = re.sub(r'^(?:student\s*name|name|roll\s*no\.?)\s*[:\-]?\s*', '', name, flags=re.I)
        name = name.strip(" -:*").strip()
        cur = {"name": name or f"Student {len(students)+1}", "subjects": []}
        if absent:
            cur["absent"] = True
        if prev:
            cur["previous_marks"] = int(prev.group(1))
        students.append(cur)
        sub = None

    out = []
    for s in students:
        for sb in s.get("subjects", []):
            if sb.get("left") is None:
                sb.pop("left", None)
        if s.get("subjects") or s.get("absent"):
            out.append(s)
    return out


def is_neet(students):
    keys = ("bio", "botan", "zoolog")
    return any(any(k in sb["name"].lower() for k in keys)
               for st in students for sb in st.get("subjects", []))
