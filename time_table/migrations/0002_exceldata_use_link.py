# Generated by Django 3.1.4 on 2021-01-01 05:55

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('time_table', '0001_initial'),
    ]

    operations = [
        migrations.AddField(
            model_name='exceldata',
            name='use_link',
            field=models.BooleanField(default=True),
        ),
    ]
