from django.urls import path
from  .views import *

urlpatterns = [

    path('login/', Login.as_view(), name='login'),
    path('logout/', Logout.as_view(), name='logout'),
   
    path('<str:whether>/getSebiOrderByDate/',GetOrderdateView.as_view(), name='getSebiOrderByDate'),
    path('download_pdf/<str:filename>/', DownloadPDFView.as_view(), name='download_pdf'),  
    path('download_all_pdfs/<str:date>/<str:whether>/', DownloadAllPDFsView.as_view(), name='download_all_pdfs'),
    
    
    path('<str:whether>/download_zip/', DownloadZipView.as_view(), name='download_zip'),
    # path('getOdiOrderByDate/',GetOrderdateView.as_view(), name='GetOrderdateView'),
    path('zipdownload/',zipdownload.as_view(), name='zipdownload'),
    path('downloadzip1/',downloadzip1.as_view(), name='downloadzip1'),
    path('<path:invalid_path>', Custom404View.as_view(), name='not_found'),
    # path('download_zip/', DownloadZipView.as_view(), name='download_zip'),
    # path('getOrdermonthView/',GetOrdermonthView.as_view(), name='GetOrdermonthView'),
    # path('getOrdertotalview/',GetOrdertotalview.as_view(), name='GetOrdertotalview')
]




