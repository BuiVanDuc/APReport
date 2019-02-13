from django.conf.urls import url

from report.controller import ReportController
urlpatterns = [
    url(r'^reports$', ReportController.as_view()),
    url(r'^exports', ReportController.as_view()),
]
 # url(r'^exports_existing')
# exports?date=2019-12-11_..&type_report=2
