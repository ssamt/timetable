# Generated by Django 3.1.4 on 2021-02-18 08:00

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('time_table', '0003_auto_20210108_1747'),
    ]

    operations = [
        migrations.AddField(
            model_name='exceldata',
            name='include_aa',
            field=models.BooleanField(default=False),
        ),
    ]
