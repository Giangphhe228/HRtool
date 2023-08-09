# Generated by Django 4.2.2 on 2023-08-08 08:28

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('OKR', '0001_initial'),
    ]

    operations = [
        migrations.AlterField(
            model_name='formula',
            name='formula_name',
            field=models.CharField(max_length=200),
        ),
        migrations.AlterField(
            model_name='formula',
            name='formula_value',
            field=models.SlugField(max_length=200),
        ),
        migrations.AlterField(
            model_name='objective',
            name='objective_name',
            field=models.CharField(max_length=200, null=True),
        ),
        migrations.AlterField(
            model_name='source',
            name='source_name',
            field=models.CharField(max_length=200),
        ),
    ]
