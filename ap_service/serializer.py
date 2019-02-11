from rest_framework import serializers

from database.models import StatisticalDataReport


class ListReportSerializer(serializers.ModelSerializer):
    class Meta:
        model = StatisticalDataReport
        fields = '__all__'
