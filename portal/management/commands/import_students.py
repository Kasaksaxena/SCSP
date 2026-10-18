"""
Import students from an Excel file straight into the database in DATABASE_URL.
The Excel file never goes through git. Usage (PowerShell):

    $env:DATABASE_URL="postgresql://..."      # omit to use local SQLite
    python manage.py import_students "D:\\data\\students.xlsx" --session JAN-2025 --dry-run
    python manage.py import_students "D:\\data\\students.xlsx" --session JAN-2025
"""
import openpyxl
from django.core.management.base import BaseCommand, CommandError

from portal.importer import import_workbook


class Command(BaseCommand):
    help = "Import students from an Excel (.xlsx) file."

    def add_arguments(self, parser):
        parser.add_argument("file")
        parser.add_argument("--session", default="", help="e.g. JAN-2025")
        parser.add_argument("--dry-run", action="store_true", help="check only, save nothing")

    def handle(self, *args, **o):
        try:
            ws = openpyxl.load_workbook(o["file"], data_only=True).active
        except Exception as e:
            raise CommandError(f"Cannot open file: {e}")
        r = import_workbook(ws, session=o["session"], dry_run=o["dry_run"])
        self.stdout.write(self.style.SUCCESS(
            f"{'WOULD IMPORT' if r.dry_run else 'IMPORTED'}: {r.created} | "
            f"duplicates skipped: {r.duplicates} | bad rows: {len(r.bad_rows)}"))
        for msg in r.bad_rows[:10]:
            self.stdout.write(f"  bad: {msg}")
        for text, n in r.warnings.items():
            self.stdout.write(self.style.WARNING(f"  warning: {n} rows - {text}"))
