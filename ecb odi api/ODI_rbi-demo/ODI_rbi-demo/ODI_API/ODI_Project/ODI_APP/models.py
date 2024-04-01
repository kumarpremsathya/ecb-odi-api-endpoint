from django.db import models
from django.contrib.auth.models import User


class rbi_odi(models.Model):
    SI_NO = models.IntegerField(primary_key=True)
    Name_of_the_Indian_Party = models.CharField(max_length=255, default=None, null=True)
    Name_of_the_JV_WOS = models.CharField(max_length=255, default=None, null=True)
    Whether_JV_WOS = models.CharField(max_length=255, default=None, null=True)
    Overseas_Country = models.CharField(max_length=255, default=None, null=True)
    Major_Activity = models.CharField(max_length=255, default=None, null=True)
    FC_Equity = models.FloatField(default=None, null=True)
    FC_Loan = models.FloatField(default=None, null=True)
    FC_Guarantee_Issued = models.FloatField(default=None, null=True)
    FC_Total = models.FloatField(default=None, null=True)
    Month = models.CharField(max_length=255, default=None, null=True)
    Year = models.CharField(max_length=255, default=None, null=True)
    date_scraped = models.DateTimeField(default=models.DateTimeField(auto_now_add=True))
    
    class Meta:
        db_table='odi'