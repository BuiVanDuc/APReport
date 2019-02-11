from django.conf.urls import url

from controller import APReportController
from report.controller import ReportController
urlpatterns = [
    url(r'^report$', APReportController.as_view()),
    url(r'^down_load_report', ReportController.as_view())
]
