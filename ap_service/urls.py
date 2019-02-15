from django.conf.urls import url

from report.controller import ReportController
from report.export.export_controller import ExportController
urlpatterns = [
    url(r'^reports$', ReportController.as_view()),
    url(r'^export_reports', ExportController.as_view()),
]

