from django.db import migrations, models


MONTH_MAP = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}


def convert_mon_yyyy_to_date(apps, schema_editor):
    studentdata = apps.get_model("portal", "studentdata")
    for student in studentdata.objects.all():
        for field in ["trained_date", "certified_date"]:
            val = getattr(student, field)
            if val and isinstance(val, str):
                parts = str(val).strip().upper().split("-")
                if len(parts) == 2:
                    month_str, year_str = parts
                    month_num = MONTH_MAP.get(month_str)
                    if month_num and year_str.isdigit():
                        setattr(student, field, f"{year_str}-{month_num:02d}-01")
                    else:
                        setattr(student, field, None)
                else:
                    setattr(student, field, None)
        student.save()


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0012_studentdata_claimed"),
    ]

    operations = [
        migrations.RunPython(convert_mon_yyyy_to_date, reverse_code=migrations.RunPython.noop),
        migrations.AlterField(
            model_name="studentdata",
            name="certified_date",
            field=models.DateField(blank=True, null=True),
        ),
        migrations.AlterField(
            model_name="studentdata",
            name="trained_date",
            field=models.DateField(blank=True, null=True),
        ),
    ]
