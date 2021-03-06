# -*- coding: utf-8 -*-
# Generated by Django 1.11.17 on 2019-02-01 10:36
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('database', '0013_auto_20190131_2301'),
    ]

    operations = [
        migrations.AddField(
            model_name='statisticaldatareport',
            name='c_i_1_2_amount',
            field=models.IntegerField(default=0),
        ),
        migrations.AddField(
            model_name='statisticaldatareport',
            name='c_i_1_2_note',
            field=models.CharField(blank=True, max_length=225),
        ),
        migrations.AddField(
            model_name='statisticaldatareport',
            name='c_i_2_note',
            field=models.CharField(blank=True, max_length=225),
        ),
        migrations.AddField(
            model_name='statisticaldatareport',
            name='c_ii_2_note',
            field=models.CharField(default=0, max_length=225),
        ),
        migrations.AddField(
            model_name='statisticaldatareport',
            name='d_ii_1_note',
            field=models.CharField(blank=True, max_length=225),
        ),
        migrations.AddField(
            model_name='statisticaldatareport',
            name='d_iii_1_a',
            field=models.CharField(blank=True, max_length=225),
        ),
        migrations.AddField(
            model_name='statisticaldatareport',
            name='d_iii_1_b_note',
            field=models.CharField(blank=True, max_length=225),
        ),
    ]
