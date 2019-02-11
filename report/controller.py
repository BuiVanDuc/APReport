from rest_framework import status
from rest_framework.generics import ListAPIView
from rest_framework.pagination import LimitOffsetPagination
from rest_framework.response import Response

from ap_service.serializer import ListReportSerializer
from core import writing_file
from database.models import StatisticalDataReport
from util.format_date import check_format_date


class ReportController(ListAPIView, LimitOffsetPagination):

    def get(self, request, *args, **kwargs):
        queryset = StatisticalDataReport.objects.all()
        date = self.request.query_params.get('date', None)

        type_report = 1
        try:
            if date == 'all_dates':
                type_report = 2
            elif date == 'searched_by_time':
                type_report = 3
                queryset = queryset.filter(is_looked_for=True)
            else:
                # handle
                check_format_date(date)
                queryset = queryset.filter(crated_at__date=date)
        except Exception as e:
            return Response(data={'Detail': 'NOT ACCEPTABLE'}, status=status.HTTP_406_NOT_ACCEPTABLE)

        if not queryset:
            return Response(data={'detail': 'have no report is searched by time'}, status=status.HTTP_200_OK)

        serializer = ListReportSerializer(instance=queryset, many=True)
        data_report = serializer.data
        number_sheet = len(data_report)

        path_report = writing_file.create_file(number_sheet, type_report)

        writing_file.modify_file(data_report, path_report)

        return Response(data=serializer.data, status=status.HTTP_200_OK)
