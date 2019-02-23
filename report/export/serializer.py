from rest_framework import serializers


class ExportMultiReportSerializer(serializers.Serializer):
    list_ids = serializers.ListSerializer(child=serializers.IntegerField())
