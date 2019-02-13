from rest_framework import status
from rest_framework.generics import CreateAPIView, ListAPIView
from rest_framework.pagination import LimitOffsetPagination
from rest_framework.parsers import JSONParser
from rest_framework.response import Response

from database.models import StatisticalDataReport
from serializer import ListReportSerializer
from util.date_util import parse_date_from_string


class ReportController(CreateAPIView, ListAPIView, LimitOffsetPagination):

    def post(self, request, *args, **kwargs):
        data = JSONParser().parse(request)
        serializer = ListReportSerializer(data=data)
        if serializer.is_valid():
            serializer.save()
            return Response(data={'detail': 'Add a new report successfully'}, status=status.HTTP_200_OK)
        return Response(data={'detail': 'NOT ACCEPTABLE'}, status=status.HTTP_406_NOT_ACCEPTABLE)

    def get(self, request, *args, **kwargs):
        try:
            date_str = self.request.query_params.get('date', None)

            if date_str:
                date = parse_date_from_string(date_str)
                if date:
                    queryset = StatisticalDataReport.objects.filter(created_at__date=date_str)
                    if not queryset:
                        return Response(data={'detail': 'No report on the date'}, status=status.HTTP_200_OK)
                    serializer = ListReportSerializer(queryset, many=True)
                    return Response(data={serializer.data}, status=status.HTTP_200_OK)
                return Response(data={'detail': 'No acceptable'}, status=status.HTTP_406_NOT_ACCEPTABLE)

            queryset = StatisticalDataReport.objects.all().order_by('-created_at')
            serializer = ListReportSerializer(queryset, many=True)
            return Response(data=serializer.data, status=status.HTTP_200_OK)
        except ValueError:
            return Response(data={'detail': 'Not Found'}, status=status.HTTP_404_NOT_FOUND)
