# Generated by Django 3.1.5 on 2021-05-06 08:34

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('masterdata', '0007_auto_20210423_0310'),
    ]

    operations = [
        migrations.CreateModel(
            name='PersonInCharge',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=200)),
            ],
        ),
    ]
