# -*- coding: utf-8 -*-
# Generated by Django 1.11.17 on 2019-02-11 14:47
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('database', '0018_remove_statisticaldatareport_c_ii_2_note'),
    ]

    operations = [
        migrations.AddField(
            model_name='statisticaldatareport',
            name='c_ii_2_note',
            field=models.CharField(blank=True, max_length=225),
        ),
    ]
