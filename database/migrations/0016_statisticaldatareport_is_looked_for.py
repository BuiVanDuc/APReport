# -*- coding: utf-8 -*-
# Generated by Django 1.11.17 on 2019-02-08 08:09
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('database', '0015_auto_20190203_0753'),
    ]

    operations = [
        migrations.AddField(
            model_name='statisticaldatareport',
            name='is_looked_for',
            field=models.BooleanField(default=False),
        ),
    ]
