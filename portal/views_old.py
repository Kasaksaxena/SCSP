# from django.shortcuts import render, redirect
# from django.http import JsonResponse
# from django.contrib import messages
# from django.contrib.auth import authenticate, login, logout
# from django.contrib.auth.decorators import login_required
# from django.views.decorators.csrf import ensure_csrf_cookie
# import openpyxl
# from .forms import StudentDataForm
# from .forms import ExcelUploadForm
# from .models import studentdata, NsqfElectronics ,NsqfIT ,Dlc


# # ─── Auth ────────────────────────────────────────────────────────────────────

# def login_view(request):
#     if request.method == 'POST':
#         username = request.POST['username']
#         password = request.POST['password']
#         user = authenticate(request, username=username, password=password)
#         if user:
#             login(request, user)
#             return redirect('dashboard')
#         return render(request, 'login.html', {'error': 'Invalid credentials'})
#     return render(request, 'login.html')


# def logout_view(request):
#     logout(request)
#     return redirect('login')


# # ─── Dashboard ───────────────────────────────────────────────────────────────

# @login_required(login_url='/login')
# @ensure_csrf_cookie
# def dashboard(request):
#     return render(request, 'dashboard.html')


# # ─── Upload ──────────────────────────────────────────────────────────────────

# def parse_bool_field(value):
#     """Excel cells can have True/False, 'yes'/'no', 1/0 — normalize all to bool."""
#     if isinstance(value, bool):
#         return value
#     if str(value).strip().lower() in ['true', 'yes', '1']:
#         return True
#     return False


# def parse_date_field(value):
#     """
#     Expects a string like 'JAN-2024' or 'January-2024' or 'Jan 2024'.
#     Returns normalized 'JAN-2024' or empty string if nothing useful.
#     """
#     if not value:
#         return ''
#     val = str(value).strip().upper().replace(' ', '-')
#     return val  # store as-is, e.g. "JAN-2024"


# @login_required(login_url='/login')
# def upload(request):
#     if request.method == 'POST':
#         form = ExcelUploadForm(request.POST, request.FILES)

#         if form.is_valid():
#             excel_file = request.FILES['file']
#             year = form.cleaned_data['year']
#             session = form.cleaned_data['session']
#             session_label = f"{session.upper()[:3]}-{year}"

#             try:
#                 wb = openpyxl.load_workbook(excel_file)
#                 sheet = wb.active

#                 headers = [
#                     str(cell.value).lower().strip() if cell.value else ''
#                     for cell in sheet[1]
#                 ]

#                 success_count = 0
#                 error_count = 0
#                 duplicate_count = 0

#                 # 🔥 Existing roll numbers (fast lookup)
#                 existing_rolls = set(
#                     studentdata.objects.values_list('roll_number', flat=True)
#                 )

#                 for row in sheet.iter_rows(min_row=2, values_only=True):

#                     if all(cell is None for cell in row):
#                         continue

#                     try:
#                         row_data = {
#                             headers[i]: row[i]
#                             for i in range(len(headers))
#                             if i < len(row)
#                         }

#                         name = str(row_data.get('name') or '').strip()
#                         roll_number = str(row_data.get('roll_number') or '').strip()
#                         course_name = str(row_data.get('course_name') or '').strip()
#                         scheme = str(row_data.get('scheme') or '').strip()

#                         # ❌ Skip duplicate roll number
#                         if roll_number and roll_number in existing_rolls:
#                             duplicate_count += 1
#                             continue

#                         # course_hour
#                         try:
#                             course_hour = int(float(str(row_data.get('course_hour') or 0)))
#                         except:
#                             course_hour = 0

#                         # fee
#                         try:
#                             fee = float(str(row_data.get('fee') or 0))
#                         except:
#                             fee = 0

#                         # mode
#                         mode = str(row_data.get('mode') or 'offline').lower().strip()
#                         if mode not in ['offline', 'online']:
#                             mode = 'offline'

#                         # caste
#                         caste = str(row_data.get('caste_category') or 'GENERAL').upper().strip()
#                         if caste not in ['OBC', 'SC', 'ST', 'PWD', 'GENERAL']:
#                             caste = 'GENERAL'

#                         # center
#                         center = str(row_data.get('center_name') or 'inderlok').lower().strip()
#                         if center not in ['inderlok', 'janakpuri', 'karkardooma']:
#                             center = 'inderlok'

#                         placed = parse_bool_field(row_data.get('placed', False))
#                         nsqf = str(row_data.get('nsqf') or '').strip()

#                         if not name or not course_name or course_hour <= 0:
#                             error_count += 1
#                             continue

#                         trained = parse_bool_field(row_data.get('trained', False))
#                         certified = parse_bool_field(row_data.get('certified', False))

#                         # auto-set dates
#                         trained_date = session_label if trained else ''
#                         certified_date = session_label if certified else ''

#                         # Extract all other personal details
#                         father_name = str(row_data.get('father_name') or '').strip()
#                         mother_name = str(row_data.get('mother_name') or '').strip()
#                         batch_code = str(row_data.get('batch_code') or '').strip().upper()
#                         gender = str(row_data.get('gender') or 'Male').strip()
#                         address = str(row_data.get('address') or '').strip()
#                         qualifications = str(row_data.get('qualifications') or '').strip()

#                         # Handle aadhaar - ensure it's stored as string, not scientific notation
#                         aadhaar_val = row_data.get('aadhaar')
#                         if aadhaar_val:
#                             aadhaar = str(int(float(aadhaar_val))).strip() if isinstance(aadhaar_val, (int, float)) else str(aadhaar_val).strip()
#                         else:
#                             aadhaar = ''

#                         # Parse DOB if exists - must be datetime.date object
#                         dob = None
#                         dob_val = row_data.get('dob')
#                         if dob_val:
#                             try:
#                                 # openpyxl returns datetime.datetime for date cells
#                                 if hasattr(dob_val, 'date'):
#                                     dob = dob_val.date()
#                                 elif isinstance(dob_val, str):
#                                     from datetime import datetime
#                                     dob = datetime.strptime(dob_val, '%Y-%m-%d').date()
#                                 else:
#                                     dob = None
#                             except Exception as dob_err:
#                                 print(f"DOB parse error: {dob_err}")
#                                 dob = None

#                         # Parse fee_date if exists - must be datetime.date object
#                         fee_date = None
#                         fee_date_val = row_data.get('fee_date')
#                         if fee_date_val:
#                             try:
#                                 # openpyxl returns datetime.datetime for date cells
#                                 if hasattr(fee_date_val, 'date'):
#                                     fee_date = fee_date_val.date()
#                                 elif isinstance(fee_date_val, str):
#                                     from datetime import datetime
#                                     fee_date = datetime.strptime(fee_date_val, '%Y-%m-%d').date()
#                                 else:
#                                     fee_date = None
#                             except Exception as fd_err:
#                                 print(f"Fee date parse error: {fd_err}")
#                                 fee_date = None

#                         student = studentdata(
#                             session=session_label,
#                             batch_code=batch_code,
#                             roll_number=roll_number,
#                             name=name,
#                             father_name=father_name,
#                             mother_name=mother_name,
#                             dob=dob,
#                             gender=gender,
#                             address=address,
#                             qualifications=qualifications,
#                             aadhaar=aadhaar,
#                             course_name=course_name,
#                             course_hour=course_hour,
#                             scheme=scheme,
#                             nsqf=nsqf,
#                             mode=mode,
#                             caste_category=caste,
#                             center_name=center,
#                             fee=fee,
#                             fee_date=fee_date,
#                             trained=trained,
#                             trained_date=trained_date,
#                             certified=certified,
#                             certified_date=certified_date,
#                             placed=placed,
#                         )

#                         student.save()
#                         success_count += 1

#                         # 🔥 Add to set to prevent duplicate inside same file
#                         if roll_number:
#                             existing_rolls.add(roll_number)

#                     except Exception as e:
#                         error_count += 1
#                         print(f"Row error: {str(e)}")
#                         import traceback
#                         traceback.print_exc()

#                 messages.success(
#                     request,
#                     f'Uploaded: {success_count} | Duplicate skipped: {duplicate_count} | Errors: {error_count}'
#                 )

#             except Exception as e:
#                 messages.error(request, f'Error reading file: {str(e)}')

#             return redirect('upload')

#     else:
#         form = ExcelUploadForm()

#     return render(request, 'upload.html', {'form': form})


# # ─── Filter (AJAX) ───────────────────────────────────────────────────────────

# @login_required(login_url='/login')
# def filter_students(request):
#     students = studentdata.objects.all()

#     center    = request.GET.get('center')
#     mode      = request.GET.get('mode')
#     caste     = request.GET.get('caste')
#     trained   = request.GET.get('trained')    # 'true' / 'false' / ''
#     certified = request.GET.get('certified')
#     placed    = request.GET.get('placed')
#     session   = request.GET.get('session')
#     scheme    = request.GET.get('scheme')
#     nsqf      = request.GET.get('nsqf')       # 'yes' / 'no'
#     quarterly = request.GET.get('quarterly')  # 'Q1-trained', 'Q1-certified', etc.
#     year      = request.GET.get('year')       # '2024', '2025', etc.

#     if center:
#         students = students.filter(center_name=center)
#     if mode:
#         students = students.filter(mode=mode)
#     if caste:
#         students = students.filter(caste_category=caste)
#     if session:
#         students = students.filter(session__icontains=session)
#     if trained:
#         students = students.filter(trained_date__gt='') if trained == 'true' else students.exclude(trained_date__gt='')
#     if certified:
#         students = students.filter(certified_date__gt='') if certified == 'true' else students.exclude(certified_date__gt='')
#     if placed:
#         students = students.filter(placed=(placed.lower() == 'true'))


#     if scheme:
#         students = students.filter(scheme__icontains=scheme)
#     if nsqf:
#         if nsqf == 'yes':
#             students = students.exclude(nsqf='').exclude(nsqf__isnull=True)
#         elif nsqf == 'no':
#             students = students.filter(nsqf='') | students.filter(nsqf__isnull=True)

#     # Helper function to get quarter from date string like "JAN-2024"
#     def get_quarter_from_date(date_str):
#         """Extract quarter and year from date string like 'JAN-2024'"""
#         if not date_str:
#             return None, None
#         try:
#             parts = date_str.split('-')
#             if len(parts) < 2:
#                 return None, None
#             month_str = parts[0].upper()
#             year_str = parts[1]

#             month_map = {
#                 'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4,
#                 'MAY': 5, 'JUN': 6, 'JUL': 7, 'AUG': 8,
#                 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
#             }

#             month = month_map.get(month_str)
#             if month is None:
#                 return None, year_str

#             # Q1: Jan-Mar (1-3), Q2: Apr-Jun (4-6), Q3: Jul-Sep (7-9), Q4: Oct-Dec (10-12)
#             quarter = (month - 1) // 3 + 1
#             return f'Q{quarter}', year_str
#         except:
#             return None, None

#     # Filter by quarterly status
#     if quarterly:
#         # quarterly format: "Q1-trained" or "Q1-certified"
#         quarter_part, status_type = quarterly.split('-') if '-' in quarterly else (quarterly, '')

#         filtered_students = []
#         for s in students:
#             if status_type == 'trained' and s.trained_date:
#                 q, y = get_quarter_from_date(s.trained_date)
#                 if q == quarter_part and (not year or y == year):
#                     filtered_students.append(s)
#             elif status_type == 'certified' and s.certified_date:
#                 q, y = get_quarter_from_date(s.certified_date)
#                 if q == quarter_part and (not year or y == year):
#                     filtered_students.append(s)

#         students = filtered_students
#     elif year:
#         # Year filter only (without quarterly)
#         filtered_students = []
#         for s in students:
#             if s.trained_date:
#                 _, y = get_quarter_from_date(s.trained_date)
#                 if y == year:
#                     filtered_students.append(s)
#                     continue
#             if s.certified_date:
#                 _, y = get_quarter_from_date(s.certified_date)
#                 if y == year:
#                     filtered_students.append(s)
#                     continue
#             # Also include if session contains year
#             if s.session and year in s.session:
#                 filtered_students.append(s)

#         students = filtered_students

#     # Sorting
#     sort_field = request.GET.get('sort_field', 'name')
#     sort_order = request.GET.get('sort_order', 'asc')

#     # Define sortable fields
#     sortable_fields = {
#         'roll_number': 'roll_number',
#         'batch_code': 'batch_code',
#         'name': 'name',
#         'course_name': 'course_name',
#         'father_name': 'father_name',
#         'mother_name': 'mother_name',
#         'dob': 'dob',
#         'gender': 'gender',
#         'address': 'address',
#         'qualifications': 'qualifications',
#         'aadhaar': 'aadhaar',
#         'scheme': 'scheme',
#         'nsqf': 'nsqf',
#         'course_hour': 'course_hour',
#         'course_category': 'course_category',
#         'center_name': 'center_name',
#         'mode': 'mode',
#         'caste_category': 'caste_category',
#         'fee': 'fee',
#         'claimable_amount': 'claimable_amount',
#         'fee_date': 'fee_date',
#         'trained': 'trained',
#         'certified': 'certified',
#         'placed': 'placed',
#         'session': 'session',
#     }

#     if sort_field in sortable_fields and not isinstance(students, list):
#         db_field = sortable_fields[sort_field]
#         if sort_order == 'desc':
#             students = students.order_by(f'-{db_field}')
#         else:
#             students = students.order_by(db_field)

#     # Pagination
#     page = int(request.GET.get('page', 1))
#     limit = int(request.GET.get('limit', 10))
#     offset = (page - 1) * limit

#     total_count = len(students) if isinstance(students, list) else students.count()
#     total_pages = (total_count + limit - 1) // limit

#     if isinstance(students, list):
#         paginated_students = students[offset:offset + limit]
#     else:
#         paginated_students = students[offset:offset + limit]

#     data = [
#     {
#         'id':              s.id,
#         'roll_number':     s.roll_number,
#         'batch_code':      s.batch_code,
#         'name':            s.name,
#         'father_name':     s.father_name,
#         'mother_name':     s.mother_name,
#         'dob':             s.dob.strftime('%Y-%m-%d') if s.dob else '',
#         'gender':          s.gender,
#         'address':         s.address,
#         'qualifications':  s.qualifications,
#         'aadhaar':         s.aadhaar,
#         'course_name':     s.course_name,
#         'scheme':          s.scheme,
#         'nsqf':            s.nsqf,
#         'course_hour':     s.course_hour,
#         'course_category': s.course_category,
#         'center_name':     s.center_name,
#         'mode':            s.mode,
#         'caste_category':  s.caste_category,
#         'fee':             float(s.fee),
#         'claimable_amount':float(s.claimable_amount),
#         'fee_date':        s.fee_date or '',
#         'trained':         s.trained,
#         'trained_date':    s.trained_date,
#         'certified':       s.certified,
#         'certified_date':  s.certified_date,
#         'placed':          s.placed,
#         'session':         s.session,
#     }
#     for s in paginated_students
# ]
#     return JsonResponse({
#         'results': data,
#         'pagination': {
#             'page': page,
#             'limit': limit,
#             'total_count': total_count,
#             'total_pages': total_pages,
#             'has_next': page < total_pages,
#             'has_prev': page > 1,
#         }
#     })


# @login_required(login_url='/login')
# def download_filtered_data(request):
#     """Download currently filtered data as Excel"""
#     import openpyxl
#     from openpyxl.styles import Font, PatternFill
#     from django.http import HttpResponse

#     students = studentdata.objects.all()

#     center    = request.GET.get('center')
#     mode      = request.GET.get('mode')
#     caste     = request.GET.get('caste')
#     trained   = request.GET.get('trained')
#     certified = request.GET.get('certified')
#     placed    = request.GET.get('placed')
#     session   = request.GET.get('session')
#     scheme    = request.GET.get('scheme')
#     nsqf      = request.GET.get('nsqf')
#     quarterly = request.GET.get('quarterly')
#     year      = request.GET.get('year')

#     if center:
#         students = students.filter(center_name=center)
#     if mode:
#         students = students.filter(mode=mode)
#     if caste:
#         students = students.filter(caste_category=caste)
#     if session:
#         students = students.filter(session__icontains=session)
#     if trained:
#         students = students.filter(trained_date__gt='') if trained == 'true' else students.exclude(trained_date__gt='')
#     if certified:
#         students = students.filter(certified_date__gt='') if certified == 'true' else students.exclude(certified_date__gt='')
#     if placed:
#         students = students.filter(placed=(placed.lower() == 'true'))

#     if scheme:
#         students = students.filter(scheme__icontains=scheme)
#     if nsqf:
#         if nsqf == 'yes':
#             students = students.exclude(nsqf='').exclude(nsqf__isnull=True)
#         elif nsqf == 'no':
#             students = students.filter(nsqf='') | students.filter(nsqf__isnull=True)

#     def get_quarter_from_date(date_str):
#         if not date_str:
#             return None, None
#         try:
#             parts = date_str.split('-')
#             if len(parts) < 2:
#                 return None, None
#             month_str = parts[0].upper()
#             year_str = parts[1]

#             month_map = {
#                 'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4,
#                 'MAY': 5, 'JUN': 6, 'JUL': 7, 'AUG': 8,
#                 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
#             }
#             month = month_map.get(month_str)
#             if month is None:
#                 return None, year_str

#             quarter = (month - 1) // 3 + 1
#             return f'Q{quarter}', year_str
#         except:
#             return None, None

#     if quarterly:
#         quarter_part, status_type = quarterly.split('-') if '-' in quarterly else (quarterly, '')
#         filtered_students = []
#         for s in students:
#             if status_type == 'trained' and s.trained_date:
#                 q, y = get_quarter_from_date(s.trained_date)
#                 if q == quarter_part and (not year or y == year):
#                     filtered_students.append(s)
#             elif status_type == 'certified' and s.certified_date:
#                 q, y = get_quarter_from_date(s.certified_date)
#                 if q == quarter_part and (not year or y == year):
#                     filtered_students.append(s)
#         students = filtered_students
#     elif year:
#         filtered_students = []
#         for s in students:
#             if s.trained_date:
#                 _, y = get_quarter_from_date(s.trained_date)
#                 if y == year:
#                     filtered_students.append(s)
#                     continue
#             if s.certified_date:
#                 _, y = get_quarter_from_date(s.certified_date)
#                 if y == year:
#                     filtered_students.append(s)
#                     continue
#             if s.session and year in s.session:
#                 filtered_students.append(s)
#         students = filtered_students

#     sort_field = request.GET.get('sort_field', 'name')
#     sort_order = request.GET.get('sort_order', 'asc')
#     sortable_fields = {
#         'roll_number': 'roll_number',
#         'batch_code': 'batch_code',
#         'name': 'name',
#         'course_name': 'course_name',
#         'father_name': 'father_name',
#         'mother_name': 'mother_name',
#         'dob': 'dob',
#         'gender': 'gender',
#         'address': 'address',
#         'qualifications': 'qualifications',
#         'aadhaar': 'aadhaar',
#         'scheme': 'scheme',
#         'nsqf': 'nsqf',
#         'course_hour': 'course_hour',
#         'course_category': 'course_category',
#         'center_name': 'center_name',
#         'mode': 'mode',
#         'caste_category': 'caste_category',
#         'fee': 'fee',
#         'claimable_amount': 'claimable_amount',
#         'fee_date': 'fee_date',
#         'trained': 'trained',
#         'certified': 'certified',
#         'placed': 'placed',
#         'session': 'session',
#     }
#     if sort_field in sortable_fields and not isinstance(students, list):
#         db_field = sortable_fields[sort_field]
#         if sort_order == 'desc':
#             students = students.order_by(f'-{db_field}')
#         else:
#             students = students.order_by(db_field)

#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Filtered Students"
#     headers = [
#         'Roll Number', 'Batch Code', 'Name', 'Father Name', 'Mother Name', 'DOB', 'Gender', 'Address',
#         'Qualifications', 'Aadhaar', 'Course Name', 'Scheme', 'NSQF', 'Course Hours', 'Course Category',
#         'Center', 'Mode', 'Caste Category', 'Fee', 'Claimable Amount', 'Fee Date', 'Trained',
#         'Trained Date', 'Certified', 'Certified Date', 'Placed', 'Session'
#     ]
#     for col_num, header in enumerate(headers, 1):
#         cell = ws.cell(row=1, column=col_num, value=header)
#         cell.font = Font(bold=True)
#         cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

#     for row_num, student in enumerate(students, 2):
#         ws.cell(row=row_num, column=1, value=student.roll_number)
#         ws.cell(row=row_num, column=2, value=student.batch_code)
#         ws.cell(row=row_num, column=3, value=student.name)
#         ws.cell(row=row_num, column=4, value=student.father_name)
#         ws.cell(row=row_num, column=5, value=student.mother_name)
#         ws.cell(row=row_num, column=6, value=student.dob.strftime('%Y-%m-%d') if student.dob else '')
#         ws.cell(row=row_num, column=7, value=student.gender)
#         ws.cell(row=row_num, column=8, value=student.address)
#         ws.cell(row=row_num, column=9, value=student.qualifications)
#         ws.cell(row=row_num, column=10, value=student.aadhaar)
#         ws.cell(row=row_num, column=11, value=student.course_name)
#         ws.cell(row=row_num, column=12, value=student.scheme)
#         ws.cell(row=row_num, column=13, value=student.nsqf)
#         ws.cell(row=row_num, column=14, value=student.course_hour)
#         ws.cell(row=row_num, column=15, value=student.course_category)
#         ws.cell(row=row_num, column=16, value=student.center_name)
#         ws.cell(row=row_num, column=17, value=student.mode)
#         ws.cell(row=row_num, column=18, value=student.caste_category)
#         ws.cell(row=row_num, column=19, value=float(student.fee))
#         ws.cell(row=row_num, column=20, value=float(student.claimable_amount))
#         ws.cell(row=row_num, column=21, value=str(student.fee_date) if student.fee_date else '')
#         ws.cell(row=row_num, column=22, value='Yes' if student.trained else 'No')
#         ws.cell(row=row_num, column=23, value=student.trained_date)
#         ws.cell(row=row_num, column=24, value='Yes' if student.certified else 'No')
#         ws.cell(row=row_num, column=25, value=student.certified_date)
#         ws.cell(row=row_num, column=26, value='Yes' if student.placed else 'No')
#         ws.cell(row=row_num, column=27, value=student.session)

#     for column in ws.columns:
#         max_length = 0
#         column_letter = column[0].column_letter
#         for cell in column:
#             try:
#                 if len(str(cell.value)) > max_length:
#                     max_length = len(str(cell.value))
#             except:
#                 pass
#         adjusted_width = min(max_length + 2, 50)
#         ws.column_dimensions[column_letter].width = adjusted_width

#     response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#     response['Content-Disposition'] = 'attachment; filename=filtered_students.xlsx'
#     wb.save(response)
#     return response


# # ─── Download Report ─────────────────────────────────────────────────────────

# @login_required(login_url='/login')
# def download(request):
#     year    = request.GET.get('year', '')
#     session = request.GET.get('session', '')
#     center  = request.GET.get('center', '')

#     students = studentdata.objects.all()
#     if year:
#         students = students.filter(session__icontains=year)
#     if session:
#         students = students.filter(session__icontains=session)
#     if center:
#         students = students.filter(center_name=center)

#     # Group by: category | course_name | center_name | session
#     grouped = {}
#     for s in students:
#         key = f"{s.course_category}|{s.course_name}|{s.center_name}|{s.session}"
#         if key not in grouped:
#             grouped[key] = {
#                 'category':    s.course_category,
#                 'course_name': s.course_name,
#                 'course_hour': s.course_hour,
#                 'center_name': s.center_name,
#                 'session':     s.session,
#                 'scheme':      s.scheme,
#                 'nsqf':        s.nsqf,
#             }
#             for c in ['GENERAL', 'OBC', 'SC', 'ST', 'PWD']:
#                 grouped[key][c] = {'trained': 0, 'certified': 0, 'placed': 0, 'total': 0}

#         caste = s.caste_category
#         grouped[key][caste]['total'] += 1
#         if s.trained_date:
#             grouped[key][caste]['trained'] += 1
#         if s.certified_date:
#             grouped[key][caste]['certified'] += 1
#         if s.placed:
#             grouped[key][caste]['placed'] += 1

#     report_data = list(grouped.values())

#     # Grand totals
#     totals = {c: {'trained': 0, 'certified': 0, 'placed': 0, 'total': 0} for c in ['GENERAL', 'OBC', 'SC', 'ST', 'PWD']}
#     totals['grand_total'] = 0
#     for item in report_data:
#         for c in ['GENERAL', 'OBC', 'SC', 'ST', 'PWD']:
#             for key in ['trained', 'certified', 'placed', 'total']:
#                 totals[c][key] += item[c][key]
#             totals['grand_total'] += item[c]['total']
#     # grand_total was being double-counted above, fix:
#     totals['grand_total'] = sum(totals[c]['total'] for c in ['GENERAL', 'OBC', 'SC', 'ST', 'PWD'])

#     context = {
#         'data': report_data,
#         'totals': totals,
#         'selected_year': year,
#         'selected_session': session,
#         'selected_center': center,
#         'years': [str(y) for y in range(2020, 2026)],
#         'sessions': ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
#                      'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'],
#     }
#     return render(request, 'download.html', context)


# # ─── API endpoint for JS-driven Excel export ─────────────────────────────────

# @login_required(login_url='/login')
# def api_download_data(request):
#     year    = request.GET.get('year', '')
#     session = request.GET.get('session', '')
#     center  = request.GET.get('center', '')

#     students = studentdata.objects.all()
#     if year:
#         students = students.filter(session__icontains=year)
#     if session:
#         students = students.filter(session__icontains=session)
#     if center:
#         students = students.filter(center_name=center)

#     data = [
#         {
#             'course_category':  s.course_category,
#             'course_name':      s.course_name,
#             'course_hour':      s.course_hour,
#             'center_name':      s.center_name,
#             'scheme':           s.scheme ,
#             'nsqf':             s.nsqf ,
#             'session':          s.session,
#             'caste_category':   s.caste_category,
#             'trained_date':     s.trained_date,
#             'certified_date':   s.certified_date,
#             'placed':           s.placed,
#             'fee':              float(s.fee),
#             'claimable_amount': float(s.claimable_amount),
#         }
#         for s in students
#     ]

#     return JsonResponse({'results': data})


# import json
# @login_required(login_url='/login')
# def update_student(request, student_id):
#     if request.method != 'POST':
#         return JsonResponse({'error': 'POST required'}, status=405)

#     try:
#         student = studentdata.objects.get(id=student_id)
#     except studentdata.DoesNotExist:
#         return JsonResponse({'error': 'Student not found'}, status=404)

#     from datetime import datetime
#     current_month = datetime.now().strftime('%b').upper() + '-' + datetime.now().strftime('%Y')

#     try:
#         body = json.loads(request.body)

#         # Safe field updates with type handling
#         student.name = (body.get('name') or student.name or '').strip()
#         student.batch_code = (body.get('batch_code') or student.batch_code or '').strip().upper()
#         student.father_name = (body.get('father_name') or student.father_name or '').strip()
#         student.mother_name = (body.get('mother_name') or student.mother_name or '').strip()
#         student.dob = body.get('dob') or student.dob
#         student.gender = body.get('gender') or student.gender or 'Male'
#         student.address = (body.get('address') or student.address or '').strip()
#         student.qualifications = (body.get('qualifications') or student.qualifications or '').strip()
#         student.aadhaar = (body.get('aadhaar') or student.aadhaar or '').strip()
#         student.course_name = (body.get('course_name') or student.course_name or '').strip()
#         student.scheme = (body.get('scheme') or student.scheme or '').strip()
#         student.nsqf = (body.get('nsqf') or student.nsqf or '').strip()

#         try:
#             student.course_hour = int(body.get('course_hour') or student.course_hour or 0)
#         except (ValueError, TypeError):
#             student.course_hour = 0

#         student.mode = body.get('mode') or student.mode or 'offline'
#         student.caste_category = body.get('caste_category') or student.caste_category or 'GENERAL'
#         student.center_name = body.get('center_name') or student.center_name or 'inderlok'

#         try:
#             student.fee = float(body.get('fee') or student.fee or 0)
#         except (ValueError, TypeError):
#             student.fee = 0.0

#         student.fee_date = body.get('fee_date') or student.fee_date
#         student.placed = body.get('placed', student.placed)
#         student.session = (body.get('session') or student.session or '').strip()

#         # trained logic
#         new_trained = body.get('trained', student.trained)
#         if new_trained and not student.trained:
#             student.trained_date = current_month
#         elif not new_trained:
#             student.trained_date = ''
#         student.trained = new_trained

#         # certified logic
#         new_certified = body.get('certified', student.certified)
#         if new_certified and not student.certified:
#             student.certified_date = current_month
#         elif not new_certified:
#             student.certified_date = ''
#         student.certified = new_certified

#         student.save()

#         return JsonResponse({
#             'success': True,
#             'course_category': student.course_category,
#             'claimable_amount': float(student.claimable_amount),
#             'trained_date': student.trained_date,
#             'certified_date': student.certified_date,
#         })

#     except json.JSONDecodeError as e:
#         return JsonResponse({'success': False, 'error': f'Invalid JSON: {str(e)}'}, status=400)
#     except Exception as e:
#         import traceback
#         error_msg = str(e)
#         tb = traceback.format_exc()
#         print(f"Update student error: {error_msg}\n{tb}")
#         return JsonResponse({'success': False, 'error': error_msg}, status=400)

# def inputView(request):
#     if request.method=='POST':
#         form=StudentDataForm(request.POST)
#         if form.is_valid():
#             student = form.save(commit=False)

#             # Auto-set dates if trained/certified but no date provided
#             if student.trained and not student.trained_date:
#                 student.trained_date = student.session
#             if student.certified and not student.certified_date:
#                 student.certified_date = student.session

#             student.save()
#             return redirect("dashboard")

#     else:
#         form=StudentDataForm()

#     context = {
#         'form': form,
#         'months': ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'],
#         'years': list(range(2020, 2031)),
#     }
#     return render(request,"input.html",context)

# from django import template

# register = template.Library()

# @register.filter(name='add_class')
# def add_class(field, css_class):
#     return field.as_widget(attrs={"class": css_class})


# def get_summary_for_queryset(queryset):
#     summary = {
#         'Total': queryset.count(),
#         'SC': 0,
#         'ST': 0,
#         'OBC': 0,
#         'PWD': 0,
#         'GENERAL': 0,
#         'B': 0,
#         'C': 0,
#         'D': 0,
#         'E': 0,
#     }
#     for s in queryset:
#         caste = s.caste_category
#         if caste in summary:
#             summary[caste] += 1

#         course_cat = (s.course_category or '').strip().upper()
#         if course_cat.startswith('B'):
#             summary['B'] += 1
#         elif course_cat.startswith('C'):
#             summary['C'] += 1
#         elif course_cat.startswith('D'):
#             summary['D'] += 1
#         elif course_cat.startswith('E'):
#             summary['E'] += 1

#     return summary


# @login_required(login_url='/login')
# def overview(request):
#     selected_session = request.GET.get('session', '')

#     students = studentdata.objects.all()
#     if selected_session:
#         students = students.filter(session=selected_session)

#     centers = [
#         ('Inderlok', get_summary_for_queryset(students.filter(center_name='inderlok'))),
#         ('Janakpuri', get_summary_for_queryset(students.filter(center_name='janakpuri'))),
#         ('Karkardooma', get_summary_for_queryset(students.filter(center_name='karkardooma'))),
#     ]

#     sessions = list(
#         studentdata.objects
#             .values_list('session', flat=True)
#             .distinct()
#             .order_by('-session')
#     )

#     context = {
#         'all_record': students.count(),
#         'centers': centers,
#         'sessions': sessions,
#         'selected_session': selected_session,
#     }

#     return render(request, 'overview.html', context)


# # old summary code removed; overview is now implemented above with dynamic sessions and optional filters


# @login_required(login_url='/login')
# def overview_data(request):
#     selected_session = request.GET.get('session', '')

#     students = studentdata.objects.all()
#     if selected_session:
#         students = students.filter(session=selected_session)

#     centers = [
#         {
#             'name': 'Inderlok',
#             'stats': get_summary_for_queryset(students.filter(center_name='inderlok'))
#         },
#         {
#             'name': 'Janakpuri',
#             'stats': get_summary_for_queryset(students.filter(center_name='janakpuri'))
#         },
#         {
#             'name': 'Karkardooma',
#             'stats': get_summary_for_queryset(students.filter(center_name='karkardooma'))
#         },
#     ]

#     sessions = list(
#         studentdata.objects
#             .values_list('session', flat=True)
#             .distinct()
#             .order_by('-session')
#     )

#     return JsonResponse({
#         'all_record': students.count(),
#         'centers': centers,
#         'sessions': sessions,
#         'selected_session': selected_session,
#     })

# def courses(request):
#     It=NsqfIT.objects.all()
#     elctro=NsqfElectronics.objects.all()
#     dlc=Dlc.objects.all()
#     return render(request,"view_courses.html",locals())
