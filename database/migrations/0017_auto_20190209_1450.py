# -*- coding: utf-8 -*-
# Generated by Django 1.11.17 on 2019-02-09 14:50
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('database', '0016_statisticaldatareport_is_looked_for'),
    ]

    operations = [
        migrations.AlterField(
            model_name='statisticaldatareport',
            name='c_ii_2_note',
            field=models.CharField(blank=True, max_length=225),
        ),
    ]
