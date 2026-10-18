"""Fake students for testing filters. No real data. python manage.py seed_fake_students --count 300"""
import random
from datetime import date, timedelta

from django.core.management.base import BaseCommand

from portal.models import studentdata, get_course_category

FIRST = ["Aarav", "Riya", "Kabir", "Sana", "Vivaan", "Isha", "Arjun", "Meera", "Rohan", "Priya"]
LAST = ["Sharma", "Verma", "Khan", "Singh", "Gupta", "Yadav", "Kumar", "Jain"]
COURSES = [("Course on Computer Concepts (CCC)", 80), ("Basic Computer Course (BCC)", 60),
           ("A-Level Course", 600), ("O-Level Course", 450), ("Other Short Term Course", 120)]


class Command(BaseCommand):
    def add_arguments(self, p):
        p.add_argument("--count", type=int, default=200)

    def handle(self, *a, **o):
        rnd, objs = random.Random(7), []
        for i in range(o["count"]):
            course, hours = rnd.choice(COURSES)
            start = date(2023, 4, 1) + timedelta(days=rnd.randint(0, 900))
            trained = rnd.random() < 0.8
            certified = trained and rnd.random() < 0.6
            s = studentdata(
                session=start.strftime("%b-%Y").upper(), roll_number=str(i + 1),
                batch_code=f"FAKE-{rnd.randint(1, 20):03d}",
                name=f"{rnd.choice(FIRST)} {rnd.choice(LAST)}".upper(),
                gender=rnd.choice(["Male", "Female"]), dob=date(2000, 1, 1) + timedelta(days=rnd.randint(0, 3000)),
                aadhaar="9" + "".join(rnd.choice("0123456789") for _ in range(11)),
                course_name=course, course_hour=hours, scheme=rnd.choice(["PMKVY", "NSQF", None]),
                nsqf=rnd.choice(["Level 3", "Level 4", None]), mode=rnd.choice(["offline", "online"]),
                caste_category=rnd.choice(["GENERAL", "OBC", "SC", "ST", "PWD"]),
                center_name=rnd.choice(["inderlok", "janakpuri", "karkardooma"]),
                fee=rnd.choice([3000, 5000, 8000]), fee_date=start,
                trained=trained, trained_date=start + timedelta(days=60) if trained else None,
                certified=certified, certified_date=start + timedelta(days=150) if certified else None,
                placed=rnd.random() < 0.2,
            )
            s.course_category = get_course_category(s.course_name, s.course_hour)
            s.claimable_amount = s.calculate_claimable_amount()
            objs.append(s)
        studentdata.objects.bulk_create(objs, batch_size=500)
        self.stdout.write(self.style.SUCCESS(f"Created {len(objs)} FAKE students"))
