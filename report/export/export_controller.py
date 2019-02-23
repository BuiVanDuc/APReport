# -*- coding: utf-8 -*-
from datetime import datetime

from rest_framework import status
from rest_framework.generics import CreateAPIView
from rest_framework.parsers import JSONParser
from rest_framework.response import Response

from ap_service.settings import REPORT_DIR
from database.models import StatisticalDataReport
from report.serializer import ListReportSerializer
from serializer import ExportMultiReportSerializer
from utils.date_utils import parse_date_from_string, convert_datetime_to_string
from utils.file_xlsx_utils import generate_name_report, create_xlsx_file_using_template, update_xlsx_file, \
    check_file_existing


class ExportController(CreateAPIView):

    def post(self, request, *args, **kwargs):
        data = JSONParser().parse(request)
        serializer = ExportMultiReportSerializer(data=data)

        if serializer.is_valid():
            # status_report: 1 --> bao cao moi duoc sinh
            # status_report: 2 --> bao cao da ton tai truoc do va duoc tao moi
            export_type = int(data['export_type'])
            is_force = int(data['is_force'])
            report_data = list()
            ret_data = dict()
            status_report = 1

            for report_id in serializer.validated_data.get("list_ids"):
                try:
                    report = StatisticalDataReport.objects.get(id=report_id)
                    report = ListReportSerializer(report).data
                    report_data.append(report)
                except StatisticalDataReport.DoesNotExist:
                    return Response(data=None, status=status.HTTP_404_NOT_FOUND)

            if len(report_data) <= 0:
                return Response(data=None, status=status.HTTP_404_NOT_FOUND)

            # export_type: 0 --> tao bao cao duy nhat
            # export_type: 1 --> sinh nhieu bao cao
            if export_type == 0:
                # create name report:
                start_date_str = data['start_date']
                end_date_str = data['end_date']
                start_date = parse_date_from_string(start_date_str)
                end_date = parse_date_from_string(end_date_str)
                name_report = generate_name_report(export_type, start_date, end_date)
                is_existing = check_file_existing(name_report)

                if is_force or is_existing is False:
                    # create report:
                    if name_report and len(name_report) > 0:
                        number_sheet = len(report_data)
                        path_report = create_xlsx_file_using_template(name_report, number_sheet)
                        number_sheet = 0
                        for data in report_data:
                            # create report and load data to report and xet if id
                            path_report = update_xlsx_file(data, path_report, number_sheet)
                            number_sheet += 1

                        ret_data['path_report'] = path_report
                        ret_data['status_report'] = status_report
                        return Response(data=ret_data, status=status.HTTP_200_OK)

                if is_existing:
                    ret_data['status_report'] = 2
                    ret_data['path_report'] = REPORT_DIR + name_report
                    return Response()
            else:
                list_info = list()
                ret_data = dict()
                for data in report_data:
                    date_str = data['created_at']
                    date = parse_date_from_string(date_str)
                    name_report = generate_name_report(export_type, created_at=date)

                    if is_force:
                        path_report = create_xlsx_file_using_template(name_report, number_sheet=1)
                        path_report = update_xlsx_file(data, path_report, index_sheet=0)
                        ret_data['path'] = path_report
                        list_info.append(ret_data.copy())

                    if check_file_existing(name_report):
                        ret_data['status_report'] = 2
                        ret_data['path'] = REPORT_DIR + name_report
                        list_info.append(ret_data.copy())
                    else:
                        path_report = create_xlsx_file_using_template(name_report, number_sheet=1)
                        path_report = update_xlsx_file(data, path_report, index_sheet=0)
                        ret_data['status_report'] = 1
                        ret_data['path'] = path_report
                        list_info.append(ret_data.copy())

                return Response(data=list_info, status=status.HTTP_200_OK)

        return Response(data=serializer.errors, status=status.HTTP_406_NOT_ACCEPTABLE)
