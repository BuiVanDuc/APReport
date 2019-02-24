# -*- coding: utf-8 -*-

from rest_framework import status
from rest_framework.generics import CreateAPIView
from rest_framework.parsers import JSONParser
from rest_framework.response import Response

from ap_service.settings import REPORT_DIR
from database.models import StatisticalDataReport
from serializer import ExportMultiReportSerializer
from utils.file_xlsx_utils import generate_report_name, create_xlsx_file_using_template, update_xlsx_file, \
    is_file_existed


class ExportController(CreateAPIView):

    def post(self, request, *args, **kwargs):
        data = JSONParser().parse(request)
        serializer = ExportMultiReportSerializer(data=data)

        if serializer.is_valid():
            validate_data = serializer.validated_data
            export_type = validate_data.get('export_type')
            is_force = validate_data.get('is_force')
            report_ids = validate_data.get('report_ids')

            '''
            {
                "status": 0, # 0 --> Report is not existed, export report failed, 1 --> Report is existed, export report failed, 2 --> Export new report
                "report_id": 1, # Id of report
                "report_url": None # Uri of report if exported successfully else it's None
            }
            '''
            ret_data = list()
            list_reports = list()
            info_report = dict()
            start_date = None
            end_date = None

            for report_id in report_ids:
                try:
                    report = StatisticalDataReport.objects.get(id=report_id)
                    list_reports.append(report)

                    if start_date is None or start_date < report.created_at:
                        start_date = report.created_at
                    if end_date is None or end_date > report.created_at:
                        end_date = report.created_at
                except StatisticalDataReport.DoesNotExist:
                    ret_data.append({"status": 0, "report_id": report_id, "report_url": None})

            if len(list_reports) <= 0:
                return Response(data=None, status=status.HTTP_404_NOT_FOUND)

            if export_type==0:
                # create name report:
                report_name = generate_report_name(export_type, start_date, end_date)

                if is_force or not is_file_existed(report_name):
                    # create report:
                    if report_name and len(report_name) > 0:

                        number_sheet = len(list_reports)
                        report_url = create_xlsx_file_using_template(report_name, number_sheet)
                        number_sheet = 0
                        for data in list_reports:
                            # create report and load data to report and xet if id
                            report_url = update_xlsx_file(data, report_url, number_sheet)
                            number_sheet += 1

                        info_report['report_url'] = report_url
                        info_report['status_report'] = 1
                        return Response(data=ret_data, status=status.HTTP_200_OK)

                elif is_file_existed(report_name):
                    info_report['status_report'] = 2
                    info_report['report_url'] = REPORT_DIR + report_name
                    return Response()
            else:
                for report in list_reports:
                    info_report = dict()
                    date = report.created_at
                    report_name = generate_report_name(export_type, created_at=date)

                    if is_force or not is_file_existed(report_name):
                        report_url = create_xlsx_file_using_template(report_name, number_sheet=1)
                        report_url = update_xlsx_file(data, report_url, index_sheet=0)
                        info_report['report_url'] = report_url
                        info_report['status_report'] = 1
                        report.append(info_report)

                    elif is_file_existed(report_name):
                        info_report['status_report'] = 2
                        info_report['report_url'] = REPORT_DIR + report_name
                        report.append(info_report)

                return Response(data=ret_data, status=status.HTTP_200_OK)

        return Response(data=serializer.errors, status=status.HTTP_406_NOT_ACCEPTABLE)
