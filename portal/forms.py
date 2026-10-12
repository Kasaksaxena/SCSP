from django import forms
from .models import studentdata 
from datetime import datetime

# ModelForm for studentdata
from django import forms
from .models import studentdata

class StudentDataForm(forms.ModelForm):
    trained_date = forms.DateField(required=False, widget=forms.DateInput(attrs={"class": "form-control", "type": "date"}))
    certified_date = forms.DateField(required=False, widget=forms.DateInput(attrs={"class": "form-control", "type": "date"}))
    
    class Meta:
        model = studentdata
        fields = [
            "session", "batch_code", "roll_number", "name", "father_name", "mother_name", "dob", "gender", "address", "qualifications", "aadhaar", "course_name", 
            "course_hour", "scheme", "nsqf", "mode", 
            "caste_category", "center_name", "fee", "fee_date", 
            "trained", "trained_date", "certified", "certified_date", "placed", "claimed"
        ]
        
        widgets = {
            # Date and Text inputs
            "session": forms.HiddenInput(attrs={"id": "id_session"}),
            "batch_code": forms.TextInput(attrs={"class": "form-control", "placeholder": "Enter Batch Code", "style": "text-transform: uppercase;", "required": "required"}),
            "roll_number": forms.TextInput(attrs={"class": "form-control", "placeholder": "Enter Roll Number"}),
            "name": forms.TextInput(attrs={"class": "form-control", "placeholder": "Full Name", "required": "required"}),
            "father_name": forms.TextInput(attrs={"class": "form-control", "placeholder": "Father's Name", "required": "required"}),
            "mother_name": forms.TextInput(attrs={"class": "form-control", "placeholder": "Mother's Name", "required": "required"}),
            "dob": forms.DateInput(attrs={"class": "form-control", "type": "date", "required": "required"}),
            "gender": forms.Select(attrs={"class": "form-select", "required": "required"}),
            "address": forms.TextInput(attrs={"class": "form-control", "placeholder": "Enter Address", "required": "required"}),
            "qualifications": forms.Select(attrs={"class": "form-select", "required": "required"}),
            "aadhaar": forms.TextInput(attrs={"class": "form-control", "placeholder": "Enter Aadhaar", "maxlength": "12", "required": "required"}),
            "course_name": forms.Select(attrs={"class": "form-select", "required": "required"}),
            "course_hour": forms.NumberInput(attrs={"class": "form-control","placeholder": "Enter course hours", "required": "required"}),
            "scheme": forms.TextInput(attrs={"class": "form-control","placeholder": "Enter Scheme", "required": "required"}),
            "nsqf": forms.Select(attrs={"class": "form-select","placeholder": "Enter NSQF"}),
            "fee": forms.NumberInput(attrs={"class": "form-control","placeholder": "Enter FEE", "required": "required"}),

            # Choice Fields (Dropdowns)
            "mode": forms.Select(attrs={"class": "form-select", "required": "required"}),
            "caste_category": forms.Select(attrs={"class": "form-select", "required": "required"}),
            "center_name": forms.Select(attrs={"class": "form-select", "required": "required"}),
            "fee_date": forms.DateInput(attrs={"class": "form-control", "type": "date", "required": "required"}),

            # Boolean Fields (Checkboxes)
            "trained": forms.CheckboxInput(attrs={"class": "form-check-input"}),
            "trained_date": forms.DateInput(attrs={"class": "form-control", "type": "date"}),
            "certified": forms.CheckboxInput(attrs={"class": "form-check-input"}),
            "certified_date": forms.DateInput(attrs={"class": "form-control", "type": "date"}),
            "placed": forms.CheckboxInput(attrs={"class": "form-check-input"}),
            "claimed": forms.CheckboxInput(attrs={"class": "form-check-input"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Make roll_number optional
        self.fields['roll_number'].required = False
class ExcelUploadForm(forms.Form):
    # Generate years from 2020 to current year + 1
    current_year = datetime.now().year
    YEAR_CHOICES = [(str(year), str(year)) for year in range(2020, current_year + 2)]
    
    SESSION_CHOICES = [
        ('JAN', 'JAN'),
        ('FEB', 'FEB'),
        ('MAR', 'MAR'),
        ('APR', 'APR'),
        ('MAY', 'MAY'),
        ('JUN', 'JUN'),
        ('JUL', 'JUL'),
        ('AUG', 'AUG'),
        ('SEP', 'SEP'),
        ('OCT', 'OCT'),
        ('NOV', 'NOV'),
        ('DEC', 'DEC'),
    ]
    
    year = forms.ChoiceField(choices=YEAR_CHOICES, required=True, widget=forms.Select(attrs={'class': 'filter-select'}))
    session = forms.ChoiceField(choices=SESSION_CHOICES, required=True, widget=forms.Select(attrs={'class': 'filter-select'}))
    file = forms.FileField(widget=forms.FileInput(attrs={'class': 'file-input'}))