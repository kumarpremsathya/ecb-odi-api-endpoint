# Generated by Django 5.0 on 2023-12-22 12:25

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='rbi_main2',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('Type', models.CharField(blank=True, max_length=500, null=True)),
                ('Borrower', models.CharField(blank=True, max_length=500, null=True)),
                ('Equivalent_Amount_in_USD', models.CharField(blank=True, max_length=500, null=True)),
                ('Purpose', models.CharField(blank=True, max_length=500, null=True)),
                ('Maturity_Period', models.CharField(blank=True, max_length=500, null=True)),
                ('Lender_Category', models.CharField(blank=True, max_length=5000, null=True)),
                ('Route', models.CharField(blank=True, max_length=500, null=True)),
                ('date_scraped', models.CharField(blank=True, max_length=500, null=True)),
            ],
            options={
                'db_table': 'rbi_main2',
            },
        ),
    ]
