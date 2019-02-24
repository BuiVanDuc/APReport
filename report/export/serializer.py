from rest_framework import serializers


class ExportMultiReportSerializer(serializers.Serializer):
    report_ids = serializers.ListSerializer(child=serializers.IntegerField())
    export_type = serializers.IntegerField(default=0)  # 0 --> 1 Sheet in 1 File, 1 --> Multiple sheets in 1 File, 2 --> Multiple Files
    is_forced = serializers.IntegerField(default=0)  # 1 --> Generated new report