# -*- coding: utf-8 -*-
from datetime import datetime

from rest_framework import status
from rest_framework.generics import CreateAPIView
from rest_framework.parsers import JSONParser
from rest_framework.response import Response

from ap_service.settings import STATIC_URL
from database.models import StatisticalDataReport
from report.serializer import ListReportSerializer
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
            is_forced = validate_data.get('is_forced')
            report_ids = validate_data.get('report_ids')

            '''
            {
                "status": 0, # 0 --> Report is not existed, export report failed, 1 --> Report is existed, export report failed, 2 --> Export new report
                "report_id": 1, # Id of report
                "report_url": None # Uri of report if exported successfully else it's None
            '''
            ret_data = list()
            list_reports = list()
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
                return Response(data=ret_data, status=status.HTTP_404_NOT_FOUND)

            # # 0 --> 1 Sheet in 1 File
            if export_type==0:
                # create name report:
                report_name = generate_report_name(export_type, start_date, end_date)

                if report_name and len(report_name) > 0:

                    if is_forced or not is_file_existed(report_name):
                        # create file report:
                        report_path = create_xlsx_file_using_template(report_name, number_sheet=1)

                        if report_path and len(report_path) > 0:

                            if len(list_reports) > 1:
                                # Convert from object to Dict
                                list_data = list()
                                for obj_report in list_reports:
                                    list_data.append(ListReportSerializer(obj_report).data)

                                sum_report = dict()
                                for i in range(1, len(list_data)):
                                    for key, val in list_data[0].items():
                                        if isinstance(list_data[0][key], unicode):
                                            sum_report[key] = ""
                                        elif isinstance(list_data[0][key], int):
                                            list_data[0][key] += list_data[i][key]
                                            sum_report[key] = list_data[0][key]

                                sum_report['created_at'] = datetime.today().strftime('%Y-%m-%d')
                                update_xlsx_file(sum_report, report_path, index_sheet=0)
                            else:
                                # have just one report
                                update_xlsx_file(vars(list_reports[0]), report_path, index_sheet=0)

                            ret_data.append(
                                {"status": 2, "report_id": report_ids, "report_url": STATIC_URL + report_name})
                        else:
                            ret_data.append({"status": 0, "report_id": report_ids, "report_url": None})

                    elif is_file_existed(report_name):
                        ret_data.append({'status': 1, "report_id": report_ids, 'report_url': STATIC_URL + report_name})
                else:
                    ret_data.append({"status": 0, "report_id": report_ids, "report_url": None})
            # 1 --> Multiple sheets in 1 File
            elif export_type == 1:
                report_name = generate_report_name(export_type, start_date, end_date)

                if report_name and len(report_name) > 0:
                    if is_forced or not is_file_existed(report_name):
                        number_sheet = len(list_reports)
                        report_path = create_xlsx_file_using_template(report_name, number_sheet)

                        if report_path and len(report_path) > 0:
                            index = 0
                            for obj_report in list_reports:
                                report_path = update_xlsx_file(ListReportSerializer(obj_report).data, report_path, index_sheet=index)
                                index += 1

                            ret_data.append({"status": 2, "report_id": report_ids, "report_url": report_path})
                        else:
                            ret_data.append({"status": 2, "report_id": report_ids, "report_url": None})
                    elif is_file_existed(report_name):
                        ret_data.append({"status": 1, "report_id": report_ids, "report_url": STATIC_URL + report_name})
                else:
                    ret_data.append({"status": 0, "report_id": report_ids, "report_url": None})
            # 2 --> Multiple Files
            elif export_type == 2:
                for obj_report in list_reports:
                    date = obj_report.created_at
                    report_id = obj_report.id
                    report_name = generate_report_name(export_type, created_at=date)

                    if report_name and len(report_name) > 0:
                        if is_forced or not is_file_existed(report_name):
                            report_path = create_xlsx_file_using_template(report_name, number_sheet=1)

                            if report_path and len(report_path) > 0:
                                update_xlsx_file(ListReportSerializer(obj_report).data, report_path, index_sheet=0)
                                ret_data.append(
                                    {"status": 2, "report_id": report_id, "report_url": STATIC_URL + report_name})
                            else:
                                ret_data.append({"status": 0, "report_id": report_id, "report_url": None})
                        elif is_file_existed(report_name):
                            ret_data.append({"status": 1, "report_id": report_id, "report_url": STATIC_URL + report_name})
                    else:
                        ret_data.append({"status": 0, "report_id": report_id, "report_url": None})
            else:
                return Response(data={"details": "Invalid export type"}, status=status.HTTP_406_NOT_ACCEPTABLE)

            return Response(data=ret_data, status=status.HTTP_200_OK)

        return Response(data=serializer.errors, status=status.HTTP_406_NOT_ACCEPTABLE)
