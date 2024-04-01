from django.urls import path
from  ECB_APP.views import *

urlpatterns = [


    path('getEcbOrderByDate/',GetOrderdateView.as_view(), name='GetOrderdateView'),
    path('<path:invalid_path>', Custom404View.as_view(), name='not_found'),
    
    # path('<path:invalid_path>', NotFoundView.as_view(), name='not_found'),
    # path('getOrdermonthView/',GetOrdermonthView.as_view(), name='GetOrdermonthView'),
    # path('getOrdertotalview/',GetOrdertotalview.as_view(), name='GetOrdertotalview'),
    path('Ecblogin/', Login.as_view(), name='login'),
    path('Ecblogout/', Logout.as_view(), name='logout'),

]