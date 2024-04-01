from django.urls import path
from  .views import *

urlpatterns = [

    path('login/', Login.as_view(), name='login'),
    path('logout/', Logout.as_view(), name='logout'),
   
    path('getOdiOrderByDate/',GetOrderdateView.as_view(), name='GetOrderdateView'),
    path('<path:invalid_path>', Custom404View.as_view(), name='not_found'),
    # path('getOrdermonthView/',GetOrdermonthView.as_view(), name='GetOrdermonthView'),
    # path('getOrdertotalview/',GetOrdertotalview.as_view(), name='GetOrdertotalview')
]




