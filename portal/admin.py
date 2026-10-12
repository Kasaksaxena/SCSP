from django.contrib import admin
from .models import studentdata , NsqfElectronics  ,NsqfIT ,Dlc
from import_export.admin import ImportExportModelAdmin
# Register your models here.


class StudentDataAdmin(ImportExportModelAdmin):
    list_display=["session",
    "roll_number" ,
    "name",
    "father_name",
    "mother_name",
    "dob",
    "gender",
    "qualifications",
    "address",
    "aadhaar",
    "course_name",    
    "course_hour",  
    "course_category",
    "scheme",       
    "nsqf",     
    "mode",         
    "caste_category", 
    "center_name",    
    "fee",         
    "claimable_amount",
    "fee_date",
    "trained",       
    "trained_date",    
    "certified",     
    "certified_date",  
    "placed"
      ]   

admin.site.register(studentdata,StudentDataAdmin)
admin.site.register(NsqfIT)
admin.site.register(NsqfElectronics)
admin.site.register(Dlc)