# Generated by Django 3.2.3 on 2021-07-30 16:58

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('contracts', '0033_auto_20210730_2032'),
    ]

    operations = [
        migrations.AddField(
            model_name='contractproduct',
            name='available',
            field=models.CharField(blank=True, default='T', max_length=1, null=True),
        ),
    ]