"""
Dashboard filters (replaces apply_filters in views.py).

Rules:
- Quarters are financial-year quarters: Q1 Apr-Jun, Q2 Jul-Sep, Q3 Oct-Dec, Q4 Jan-Mar.
- Year + Quarter are checked on the SAME date (old code mixed month of one date
  with year of another date).
- Trained = Yes + period  -> trained_date must fall in the period.
- Certified = Yes + period -> certified_date must fall in the period.
- Only a period, no Trained/Certified chosen -> trained_date OR certified_date in period.
- Search: ?q= (or ?search=) matches name, roll number, batch code.
- Sort: ?sort_field=<field>&sort_order=asc|desc (what dashboard.html sends), default by id.
"""
from django.db.models import Q

from .models import studentdata

QUARTER_MONTHS = {"Q1": [4, 5, 6], "Q2": [7, 8, 9], "Q3": [10, 11, 12], "Q4": [1, 2, 3]}

SORTABLE_FIELDS = {
    "roll_number", "batch_code", "name", "course_name", "father_name", "mother_name",
    "dob", "gender", "address", "qualifications", "aadhaar", "scheme", "nsqf",
    "course_hour", "course_category", "center_name", "mode", "caste_category", "fee",
    "claimable_amount", "fee_date", "trained", "certified", "placed", "session", "claimed",
    "trained_date", "certified_date",
}


def _bool(value):
    return str(value).strip().lower() in ("true", "yes", "1")


def _in_period(field, year, months):
    q = Q()
    if year:
        q &= Q(**{f"{field}__year": year})
    if months:
        q &= Q(**{f"{field}__month__in": months})
    return q


def apply_filters(params):
    qs = studentdata.objects.all()

    for param, field in (("center", "center_name"), ("mode", "mode"),
                         ("caste", "caste_category"), ("scheme", "scheme")):
        if params.get(param):
            qs = qs.filter(**{field: params[param]})

    for param in ("placed", "claimed"):
        if params.get(param):
            qs = qs.filter(**{param: _bool(params[param])})

    nsqf = params.get("nsqf")
    if nsqf == "yes":
        qs = qs.exclude(nsqf__isnull=True).exclude(nsqf="")
    elif nsqf == "no":
        qs = qs.filter(Q(nsqf__isnull=True) | Q(nsqf=""))

    text = (params.get("q") or params.get("search") or "").strip()
    if text:
        qs = qs.filter(
            Q(name__icontains=text) | Q(roll_number__icontains=text) | Q(batch_code__icontains=text)
        )

    year = params.get("year")
    year = int(year) if year and str(year).isdigit() else None
    months = QUARTER_MONTHS.get(params.get("quarterly"))
    has_period = bool(year or months)
    trained_in = _in_period("trained_date", year, months)
    certified_in = _in_period("certified_date", year, months)

    trained, certified = params.get("trained"), params.get("certified")
    period_used = False

    if trained == "or":
        qs = qs.filter((Q(trained=True) & trained_in) | (Q(certified=True) & certified_in)
                       if has_period else Q(trained=True) | Q(certified=True))
        period_used = True
    elif trained:
        if _bool(trained):
            qs = qs.filter(Q(trained=True) & trained_in)
            period_used = True
        else:
            qs = qs.filter(trained=False)

    if certified:
        if _bool(certified):
            qs = qs.filter(Q(certified=True) & certified_in)
            period_used = True
        else:
            qs = qs.filter(certified=False)

    if has_period and not period_used and not (trained or certified):
        qs = qs.filter(trained_in | certified_in)

    sort = params.get("sort_field") or params.get("sort")
    direction = params.get("sort_order") or params.get("dir")
    if sort in SORTABLE_FIELDS:
        qs = qs.order_by(("-" if direction == "desc" else "") + sort, "id")
    else:
        qs = qs.order_by("id")
    return qs
