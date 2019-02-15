# -*- coding: utf-8 -*-
from datetime import datetime

from rest_framework import status
from rest_framework.generics import ListAPIView
from rest_framework.pagination import LimitOffsetPagination
from rest_framework.response import Response

from database.models import StatisticalDataReport
from report.serializer import ListReportSerializer
from utils.date_utils import parse_date_from_string, convert_datetime_to_string
from utils.file_xlsx_utils import create_xlsx_file_using_template, update_xlsx_file, generate_name_report, \
    rename_file_report_existing, find_report_existing, create_new_name_for_xlsx_file, get_path_file_using_name


class ExportController(ListAPIView, LimitOffsetPagination):

    def get(self, request, *args, **kwargs):
        start_date_str = self.request.query_params.get('start_date', None)
        end_date_str = self.request.query_params.get('end_date', None)
        date_str = self.request.query_params.get('date', None)
        type_report =self.request.query_params.get('type_report', None)
        is_force = self.request.query_params.get('is_force', None)
        status_report = 0
        ret_data = dict()
        list_date = list()
        info_report = dict()
        ret_reports = list()
        time_download = datetime.now()
        time_download_str = convert_datetime_to_string(time_download, 1)

        if start_date_str and end_date_str:
            start_date = parse_date_from_string(start_date_str)
            end_date = parse_date_from_string(end_date_str)
            # check range time correctly
            if end_date > start_date:
                queryset = StatisticalDataReport.objects.filter(created_at__range=(end_date, start_date))
                if not queryset:
                    return Response(data={'detail': 'have no report'}, status=status.HTTP_200_OK)
                serializer = ListReportSerializer(queryset, many=True)
                number_objects = queryset.count()
                list_data_reports = serializer.data

                if type_report == 1:
                    list_date = [start_date, end_date]
                    # Generate name report
                    list_name_reports = generate_name_report(list_date, type_report)

                    if len(list_name_reports) > 0:
                        # Check report existing
                        list_reports_existing = find_report_existing(list_name_reports,is_existing=True)

                        if len(list_reports_existing) > 0:
                            status_report = 2
                            number_sheet = number_objects
                            list_new_reports = create_xlsx_file_using_template(number_sheet, type_report,
                                                                               list_date)
                            # Report is existing and check still get report or new report
                            if is_force == 1:
                                # rename file old report to create new report
                                list_old_name_reports = list_reports_existing
                                list_new_name = create_new_name_for_xlsx_file(list_old_name_reports, time_download_str)
                                list_files_rename = rename_file_report_existing(list_old_name_reports, list_new_name)

                                if len(list_files_rename) > 0:
                                    # create new reports using template
                                    if len(list_new_reports) > 0:
                                        # Load data report into file report
                                        list_path_reports = update_xlsx_file(list_data_reports, list_new_reports)
                                        if len(list_path_reports) > 0:
                                            status_report = 1
                                            for path_report in list_path_reports:
                                                info_report['path'] = path_report
                                                info_report['status_report'] = status_report
                                                ret_reports.append(info_report)

                                            ret_data['reports'] = ret_reports
                                            ret_data['is_force'] = is_force
                            elif is_force == 0:
                                list_path_reports = get_path_file_using_name(list_reports_existing)

                                for path_report in list_path_reports:
                                    info_report['path'] = path_report
                                    info_report['status_report'] = status_report
                                    ret_reports.append(info_report)

                                ret_data['reports'] = ret_reports
                                ret_data['is_force'] = is_force
                                return Response(data=ret_data, status=status.HTTP_200_OK)
                            else:
                                list_path_existing = get_path_file_using_name(list_reports_existing)

                                for path_existing in list_path_existing:
                                    info_report['path'] = path_existing
                                    info_report['status_report'] = status_report
                                    ret_reports.append(info_report)

                                ret_data['reports'] = ret_reports
                                return Response(data=ret_data, status=status.HTTP_200_OK)
                        else:
                            number_sheet = number_objects
                            # Create file report
                            list_reports = create_xlsx_file_using_template(number_sheet, type_report, list_date)
                            if len(list_reports) >0:

                                # Load data into file reports
                                list_path_reports = update_xlsx_file(list_data_reports,list_reports)
                                if len(list_path_reports):
                                    status_report =1

                                    for path_report in list_path_reports:
                                        info_report['path'] = path_report
                                        info_report['status_report'] = status_report
                                        ret_reports.append(info_report)

                                    ret_data['reports'] = ret_reports
                                    return Response(data=ret_data, status=status.HTTP_200_OK)
                    return Response(data={'detail':'Internal Server Error '}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

                elif type_report == 2:
                    list_date = list()
                    #
                    for data_report in list_data_reports:
                        date_time_str = data_report['created_at']
                        date_time = parse_date_from_string(date_time_str)
                        list_date.append(date_time)

                    # Create name report:
                    list_name_reports = generate_name_report(list_date, type_report)

                    if len(list_name_reports) > 0:
                        # Check report if report existing
                        list_reports_existing = find_report_existing(list_name_reports,is_existing=True)

                        if len(list_reports_existing) > 0:
                            list_reports_not_existing = find_report_existing(list_name_reports,is_existing=False)

                            if is_force == 1:
                                list_old_name_reports = list_reports_existing
                                list_new_name = create_new_name_for_xlsx_file(list_old_name_reports, time_download_str)
                                list_files_rename = rename_file_report_existing(list_old_name_reports, list_new_name)

                                if len(list_files_rename) > 0:
                                    # Create new report
                                    list_new_reports = create_xlsx_file_using_template(1, type_report, list_date)

                                    if len(list_new_reports) > 0:
                                        # Load data reports into report
                                        list_path_reports = update_xlsx_file(list_data_reports, list_new_reports)

                                        if len(list_path_reports) > 0:
                                            status_report = 1
                                            info_report = dict()
                                            ret_reports = list()

                                            for path_report in list_path_reports:
                                                info_report['path'] = path_report
                                                info_report['status_report'] = status_report
                                                ret_reports.append(info_report)

                                            ret_data['reports'] = ret_reports
                                            ret_data['is_force'] = is_force
                                            return Response(data=ret_data, status=status.HTTP_200_OK)
                            elif is_force == 0:
                                # Find report not existing to create new report

                                list_new_date = list()
                                list_new_data_objs = list()
                                status_report = 2

                                # find date is new crated
                                if len(list_reports_not_existing) > 0:
                                    # Create file reports
                                    for reports_not_existing in list_reports_not_existing:
                                        index = list_name_reports.index(reports_not_existing)
                                        if index:
                                            new_date = list_date[index]
                                            new_data_obj = list_data_reports[index]
                                            list_new_date.append(new_date)
                                            list_new_data_objs.append(new_data_obj)

                                    list_new_reports = create_xlsx_file_using_template(1, type_report, list_new_date)

                                    if len(list_new_reports) > 0:
                                        # Load data into reports:
                                        list_path_reports = update_xlsx_file(list_new_data_objs,
                                                                             list_reports_not_existing)
                                        if len(list_path_reports) > 0:
                                            status_report = 1

                                            for path_report in list_path_reports:
                                                info_report['path'] = path_report
                                                info_report['status_report'] = status_report
                                                ret_reports.append(info_report)

                                list_path_reports = get_path_file_using_name(list_reports_existing)
                                for path_report in list_path_reports:
                                    info_report['path'] = path_report
                                    info_report['status_report'] = status_report
                                    ret_reports.append(info_report)

                                return Response(data=ret_data, status=status.HTTP_200_OK)
                            else:
                                list_paths_existing = get_path_file_using_name(list_reports_existing)
                                for path_existing in list_paths_existing:
                                    info_report['path'] = path_existing
                                    info_report['status_report'] = status_report
                                    ret_reports.append(info_report)

                                return Response(data=ret_data, status=status.HTTP_200_OK)
                        else:
                            # Create new report
                            number_sheet = 1
                            list_new_reports = create_xlsx_file_using_template(number_sheet, type_report,
                                                                               list_data_reports)

                            if len(list_new_reports) > 0:
                                list_name_reports = list_new_reports
                                list_path_reports = update_xlsx_file(list_data_reports, list_name_reports)

                                if len(list_path_reports) > 0:
                                    status_report = 1

                                    for path_report in list_path_reports:
                                        info_report['path'] = path_report
                                        info_report['status_report'] = status_report
                                        ret_reports.append(info_report)
                                    return Response(data=ret_data, status=status.HTTP_200_OK)

                    return Response(data={'detail':'Internal Server Error '}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
            else:
                return Response(
                    data={'detail': 'please input again format date. Ex:?from_date=2010-12-03&end_date=2010-12-04'},
                    status=status.HTTP_404_NOT_FOUND)
                            
        elif date_str:
            # parse date to string
            date = parse_date_from_string(date_str)

            if date:
                list_date.append(date_str)
                queryset = StatisticalDataReport.objects.filter(created_at__date=date)

                if not queryset:
                    return Response(data={'detail': 'have no report at date'}, status=status.HTTP_200_OK)
                # Check report if existing:
                list_name_reports = generate_name_report(list_date)
                serializer = ListReportSerializer(queryset, many=True)
                list_data_reports = serializer.data

                if len(list_name_reports) > 0:
                    list_reports_existing = find_report_existing(list_name_reports,is_existing=True)

                    if len(list_reports_existing) > 0:
                        status_report = 2
                        if is_force == 1:
                            # rename old file
                            list_old_name_reports = list_reports_existing
                            list_new_reports = create_new_name_for_xlsx_file(list_old_name_reports, time_download_str)
                            # rename old file
                            if list_old_name_reports and list_new_reports:
                                list_files_rename = rename_file_report_existing(list_old_name_reports, list_new_reports)
                                # create new xlsx file report

                                list_name_reports = create_xlsx_file_using_template(1, None, list_date)

                                if len(list_name_reports) > 0:
                                    # Load data reports into file report
                                    list_path_reports = update_xlsx_file(list_data_reports, list_name_reports)

                                    if len(list_path_reports) > 0:
                                        for path_report in list_path_reports:
                                            info_report['path'] = path_report
                                            info_report['status_report'] = status_report
                                            ret_reports.append(info_report)

                                        return Response(data=ret_data, status=status.HTTP_200_OK)
                        elif is_force == 0:
                            list_path_reports = get_path_file_using_name(list_reports_existing)

                            if len(list_path_reports) > 0:
                                for path_report in list_path_reports:
                                    info_report['path'] = path_report
                                    info_report['status_report'] = status_report
                                    ret_reports.append(info_report)

                                return Response(data=ret_data, status=status.HTTP_200_OK)
                        else:
                            # Create new report
                            number_sheet =1
                            list_name_reports = create_xlsx_file_using_template(number_sheet,None,list_date)
                            if len(list_name_reports) >0:
                                list_path_reports = update_xlsx_file(list_data_reports, list_name_reports)
                                status_report =1
                                if len(list_data_reports) >0:
                                    for path_report in list_path_reports:
                                        info_report['path'] = path_report
                                        info_report['status_report'] = status_report
                                        ret_reports.append(info_report)
                                    return Response(data=ret_data, status=status.HTTP_200_OK)
                    else:
                        # Create new report
                        list_name_reports = create_xlsx_file_using_template(1, None, list_date)

                        if len(list_name_reports) > 0:
                            # Load data reports into file report
                            list_path_reports = update_xlsx_file(list_data_reports, list_name_reports)

                            if len(list_path_reports) > 0:
                                for path_report in list_path_reports:
                                    info_report['path'] = path_report
                                    info_report['status_report'] = status_report
                                    ret_reports.append(info_report)

                                return Response(data=ret_data, status=status.HTTP_200_OK)
                else:
                    return Response(data={'detail':'Internal Server Error '}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

                return Response(data={'detail': 'please input again correct format date. Ex: 2019-12-11'})
