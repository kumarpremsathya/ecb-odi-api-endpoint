from django.db import models
from django.contrib.auth.models import User


class rbi_main2(models.Model):

    Type = models.CharField(max_length=500, blank=True, null=True)
    Borrower  = models.CharField(max_length=500, blank=True, null=True)
    Economic_sector_of_borrower= models.CharField(max_length=500, blank=True, null=True)
    Equivalent_Amount_in_USD= models.CharField(max_length=500, blank=True, null=True)
    Purpose  = models.CharField(max_length=500, blank=True, null=True)
    Maturity_Period= models.CharField(max_length=500, blank=True, null=True)
    Lender_Category  = models.CharField(max_length=5000, blank=True, null=True)
    Month=models.CharField(max_length=500, blank=True, null=True)
    Year= models.CharField(max_length=500, blank=True, null=True)
    Route= models.CharField(max_length=500, blank=True, null=True)
    date_scraped=models.CharField(max_length=500, blank=True, null=True)
    
    class Meta:
        db_table='new_2'


    


    