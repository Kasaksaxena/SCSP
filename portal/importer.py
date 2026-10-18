"""
One importer used by BOTH the Upload page and `manage.py import_students`.
Cleans Excel values so the database and the dashboard filters agree.
"""
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from .models import get_course_category, studentdata

MODES = {"offline": "offline", "off campus": "offline", "online": "online", "on campus": "online"}
CASTES = {"general": "GENERAL", "gen": "GENERAL", "obc": "OBC", "sc": "SC", "st": "ST", "pwd": "PWD"}
GENDERS = {"m": "Male", "male": "Male", "f": "Female", "female": "Female"}
CENTERS = {"inderlok": "inderlok", "janakpuri": "janakpuri", "karkardooma": "karkardooma"}


def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "null"):
        return None
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def to_date(v):
    if not v:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            pass
    return None


def to_bool(v):
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("true", "yes", "1", "y")


@dataclass
class ImportResult:
    created: int = 0
    duplicates: int = 0
    bad_rows: list = field(default_factory=list)
    warnings: Counter = field(default_factory=Counter)
    dry_run: bool = False


def import_workbook(ws, session="", dry_run=False):
    result = ImportResult(dry_run=dry_run)
    headers = [re.sub(r"\s+", "_", str(c.value).strip().lower()) if c.value else "" for c in ws[1]]

    existing = set()
    for a, c, b, n, d in studentdata.objects.values_list("aadhaar", "course_name", "batch_code", "name", "dob"):
        existing.add((a, c, b) if a else ((n or "").lower(), d, c, b))

    to_create = []
    for rownum, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        r = dict(zip(headers, row))
        if not any(v not in (None, "") for v in r.values()):
            continue
        name, course = clean(r.get("name")), clean(r.get("course_name"))
        center = CENTERS.get((clean(r.get("center_name")) or "").lower())
        if not name or not course or not center:
            result.bad_rows.append(f"row {rownum}: needs name, course_name and a valid center_name")
            continue

        aadhaar = re.sub(r"\D", "", clean(r.get("aadhaar")) or "")
        if aadhaar and len(aadhaar) != 12:
            result.warnings["aadhaar not 12 digits -> saved empty"] += 1
            aadhaar = ""
        aadhaar = aadhaar or None
        batch, dob = clean(r.get("batch_code")), to_date(r.get("dob"))
        key = (aadhaar, course, batch) if aadhaar else (name.lower(), dob, course, batch)
        if key in existing:
            result.duplicates += 1
            continue
        existing.add(key)

        try:
            fee = Decimal(str(clean(r.get("fee")) or 0))
        except InvalidOperation:
            fee = Decimal(0)
            result.warnings["fee is not a number -> saved as 0"] += 1
        try:
            hours = int(float(clean(r.get("course_hour")) or 0))
        except ValueError:
            hours = 0
        nsqf = clean(r.get("nsqf"))
        if nsqf and nsqf.isdigit():
            nsqf = f"Level {nsqf}"

        for col, mapping, label in (("mode", MODES, "mode"), ("caste_category", CASTES, "caste_category"),
                                    ("gender", GENDERS, "gender")):
            val = clean(r.get(col))
            if val and val.lower() not in mapping:
                result.warnings[f"unknown {label} value -> default used"] += 1

        t_date, c_date = to_date(r.get("trained_date")), to_date(r.get("certified_date"))
        trained = to_bool(r.get("trained")) or bool(t_date)
        certified = to_bool(r.get("certified")) or bool(c_date)
        if trained and not t_date:
            result.warnings["trained but no trained_date (quarter filters skip them)"] += 1
        if certified and not c_date:
            result.warnings["certified but no certified_date (quarter filters skip them)"] += 1

        s = studentdata(
            session=session or clean(r.get("session")),
            roll_number=clean(r.get("roll_number")), batch_code=batch, name=name,
            father_name=clean(r.get("father_name")), mother_name=clean(r.get("mother_name")),
            dob=dob, gender=GENDERS.get((clean(r.get("gender")) or "").lower()),
            address=clean(r.get("address")), qualifications=clean(r.get("qualifications")),
            aadhaar=aadhaar, course_name=course, course_hour=hours, scheme=clean(r.get("scheme")),
            nsqf=nsqf, mode=MODES.get((clean(r.get("mode")) or "").lower(), "offline"),
            caste_category=CASTES.get((clean(r.get("caste_category")) or "").lower(), "GENERAL"),
            center_name=center, fee=fee, fee_date=to_date(r.get("fee_date")),
            trained=trained, trained_date=t_date, certified=certified, certified_date=c_date,
            placed=to_bool(r.get("placed")), claimed=to_bool(r.get("claimed")),
        )
        s.course_category = get_course_category(s.course_name, s.course_hour)
        s.claimable_amount = s.calculate_claimable_amount()
        to_create.append(s)

    if not dry_run:
        studentdata.objects.bulk_create(to_create, batch_size=500)
    result.created = len(to_create)
    return result
