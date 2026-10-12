from django.contrib import admin
from django.urls import path

from portal import views

urlpatterns = [
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("", views.dashboard, name="dashboard"),
    path("upload/", views.upload, name="upload"),
    path("filter-students/", views.filter_students, name="filter_students"),
    path("download/", views.download, name="download"),
    path("api/download-data/", views.api_download_data, name="api_download_data"),
    path("overview-data/", views.overview_data, name="overview_data"),
    path(
        "update-student/<int:student_id>/", views.update_student, name="update_student"
    ),
    path("input_student/", views.inputView, name="input"),
    path("overview/", views.overview, name="overview"),
    path("view_courses/", views.courses, name="courses"),
    path(
        "download-filtered-data/",
        views.download_filtered_data,
        name="download_filtered_data",
    ),
    path("sample/", views.sample_upload, name="sample"),
]
