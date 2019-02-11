import datetime

from rest_framework import status
from rest_framework.generics import CreateAPIView, ListAPIView
from rest_framework.pagination import LimitOffsetPagination
from rest_framework.parsers import JSONParser
from rest_framework.response import Response

from database.models import StatisticalDataReport
from serializer import ListReportSerializer
from util.format_date import check_format_date

class APReportController(CreateAPIView, ListAPIView, LimitOffsetPagination):

    def post(self, request, *args, **kwargs):
        data = JSONParser().parse(request)
        serializer = ListReportSerializer(data=data)
        if serializer.is_valid():
            serializer.save()
            return Response(data={'detail': 'Add a new report successfully'}, status=status.HTTP_200_OK)
        return Response(data={'detail': 'NOT ACCEPTABLE'}, status=status.HTTP_406_NOT_ACCEPTABLE)

    def get(self, request, *args, **kwargs):
        try:
            queryset = StatisticalDataReport.objects.all()
            date = self.request.query_params.get('date', None)

            if date == 'all_dates':
                queryset = StatisticalDataReport.objects.all().order_by('-crated_at')
                serializer = ListReportSerializer(data=queryset, many=True)
                return Response(data=serializer.data, status=status.HTTP_200_OK)
            # Check format of date
            check_format_date(date)

            queryset = queryset.filter(crated_at__date=date)
            if not queryset:
                return Response(data={'detail': 'have no report in the date'}, status=status.HTTP_200_OK)

            # Cornfirm report is looked for
            object = queryset.get(crated_at__date=date)
            object.is_looked_for = True
            object.save()

            serializer = ListReportSerializer(queryset, many=True)

            return Response(data=serializer.data, status=status.HTTP_200_OK)
        except ValueError:
            return Response(data={'detail': 'NOT ACCEPTABLE'}, status=status.HTTP_406_NOT_ACCEPTABLE)
        except StatisticalDataReport.DoesNotExist:
            pass
        return Response(data={'detail': 'Not Found'}, status=status.HTTP_404_NOT_FOUND)
