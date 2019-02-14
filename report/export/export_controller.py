from rest_framework import status
from rest_framework.generics import ListAPIView
from rest_framework.pagination import LimitOffsetPagination
from rest_framework.response import Response
from serializer import ListReportSerializer
from utils import file_xlsx_utils
from database.models import StatisticalDataReport
from utils.date_utils import parse_date_from_string
from  utils.file_xlsx_utils import create_xlsx_file_using_template, update_xlsx_file, generate_path_report, check_report_existing, rename_file_report_existing


class ExportController(ListAPIView, LimitOffsetPagination):

    def get(self, request, *args, **kwargs):
        start_date_str = self.request.query_params.get('start_date', None)
        end_date_str = self.request.query_params.get('end_date', None)
        date_str = self.request.query_params.get('date', None)
        type_report =self.request.query_params.get('type_report', None)
        is_force = self.request.query_params.get('is_force', None)
        is_created = False
        ret_data = dict()
        list_date = list()

        if type_report == 3:
            number_sheet = 1

        if start_date_str and end_date_str:
            start_date = parse_date_from_string(start_date_str)
            end_date = parse_date_from_string(end_date_str)
            # check range time correctly
            if end_date > start_date:
                queryset = StatisticalDataReport.objects.filter(created_at__range=(end_date, start_date))
                if not queryset:
                    pass
                serializer = ListReportSerializer(queryset, many=True)
                list_object_data = serializer.data
                count = len(list_object_data)
                if type_report ==2:
                    number_sheet =count

                list_date.append(start_date_str)
                list_date.append(end_date_str)

                # Generate path report
                list_path_report = generate_path_report(list_date, type_report)
                if len(list_path_report) > 0:
                    # Check report existing
                    list_report_existing = check_report_existing(list_path_report)
                    if len(list_report_existing) >0:
                        is_created = True
                        if is_force:
                            # rename file old report to create new report
                            number_file_rename = rename_file_report_existing(list_report_existing)
                            if number_file_rename >0:
                                # create new reports using tempale
                                list_reports = create_xlsx_file_using_template(number_sheet,type_report, list_date)
                                if list_reports and len(list_reports) >0:
                                    # Load data into report
                                    result_update_file = update_xlsx_file(list_object_data, list_path_report)
                                    if result_update_file >0:
                                        ret_data['list_path_report'] =list_path_report
                                        ret_data['is_force']= is_force
                                        ret_data['type_report'] = type_report
                                        return Response(data=ret_data, status=status.HTTP_200_OK)
                        else:
                            ret_data['list_path_report'] = list_path_report
                            ret_data['is_created'] = is_created
                            ret_data['is_force'] = is_force
                            ret_data['type_report'] = type_report
                            return Response(data=ret_data, status=status.HTTP_200_OK)
                    else:
                        # create report
                        list_reports = create_xlsx_file_using_template(number_sheet, type_report, list_date)
                        if len(list_reports) > 0:
                            # Load data into report
                            result_update_file = update_xlsx_file(list_object_data, list_path_report)
                            if result_update_file > 0:
                                ret_data['list_path_report'] = list_path_report
                                ret_data['type_report'] = type_report
                                return Response(data=ret_data, status=status.HTTP_200_OK)
                else:
                    return Response(data={'detail':'Internal Server Error '}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        elif date_str:
            date = parse_date_from_string(date_str)
            if date:
                queryset =StatisticalDataReport.objects.filter(created_at_date=date_str)
                if not queryset:
                    return Response(data={'detail':'have no report'}, status=status.HTTP_200_OK)
                serializer = ListReportSerializer(queryset, many=True)
                list_object_data = serializer.data
                count = len(list_object_data)
                if type_report == 2:
                    number_sheet = count

                list_date.append(date_str)
                list_path_report = generate_path_report(list_date, type_report)
                if len(list_path_report) > 0:
                    # Check report existing
                    list_report_existing = check_report_existing(list_path_report)
                    if len(list_report_existing) >0:
                        is_created = True
                        if is_force:
                            # rename file old report to create new report
                            number_file_rename = rename_file_report_existing(list_report_existing)
                            if number_file_rename >0:
                                # create new reports using tempale
                                list_reports = create_xlsx_file_using_template(number_sheet,type_report, list_date)
                                if list_reports and len(list_reports) >0:
                                    # Load data into report
                                    result_update_file = update_xlsx_file(list_object_data, list_path_report)
                                    if result_update_file >0:
                                        ret_data['list_path_report'] =list_path_report
                                        ret_data['is_force']= is_force
                                        ret_data['type_report'] = type_report
                                        return Response(data=ret_data, status=status.HTTP_200_OK)
                        else:
                            ret_data['list_path_report'] = list_path_report
                            ret_data['is_created'] = is_created
                            ret_data['is_force'] = is_force
                            ret_data['type_report'] = type_report
                            return Response(data=ret_data, status=status.HTTP_200_OK)
                    else:
                        # create report
                        list_reports = create_xlsx_file_using_template(number_sheet, type_report, list_date)
                        if len(list_reports) > 0:
                            # Load data into report
                            result_update_file = update_xlsx_file(list_object_data, list_path_report)
                            if result_update_file > 0:
                                ret_data['list_path_report'] = list_path_report
                                ret_data['type_report'] = type_report
                                return Response(data=ret_data, status=status.HTTP_200_OK)
                else:
                    return Response(data={'detail':'Internal Server Error '}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        return Response(data={'detail':'Not acceptable'}, status=status.HTTP_406_NOT_ACCEPTABLE)



