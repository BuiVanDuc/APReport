from rest_framework import status
from rest_framework.generics import ListAPIView
from rest_framework.pagination import LimitOffsetPagination
from rest_framework.response import Response

from util import file_xlsx_util
from database.models import StatisticalDataReport
from util.date_util import parse_date_from_string


class ExportController(ListAPIView, LimitOffsetPagination):

    def get(self, request, *args, **kwargs):
        date_str = self.request.query_params.get('date', None)
        type_report = 1
        if date_str:
            date = parse_date_from_string(date_str)
            if date:
                return Response(data={'detail': 'Invalid date string, please use date in format: %Y-%m-%d, For ex: 2019-02-11'},
                status=status.HTTP_406_NOT_ACCEPTABLE)

        else:
            request = StatisticalDataReport.objects.all()

        number_sheet = request.count()
        if number_sheet >0:
            path_report =


        path_report = file_xlsx_util.create_xlsx_file(number_sheet, type_report)

        file_xlsx_util.update_xlsx_file(data_report, path_report)


