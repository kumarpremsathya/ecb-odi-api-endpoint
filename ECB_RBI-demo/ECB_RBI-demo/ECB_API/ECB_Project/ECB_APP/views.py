from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework import status
from rest_framework.permissions import IsAuthenticated, AllowAny
from rest_framework.authtoken.models import Token
from django.contrib.auth import authenticate
from ECB_APP.models import *
from datetime import datetime
from django.http import Http404



class Login(APIView):
    # permission_classes = (AllowAny,)
    permission_classes = (IsAuthenticated,)
    def post(self, request):
        username = request.data.get("username")
        password = request.data.get("password")

        if username is None or password is None:
            return Response({'message': 'Please provide both username and password'},
                            status=status.HTTP_400_BAD_REQUEST)
        user = authenticate(username=username, password=password)  
        if not user:
            return Response({'status': 'Failed', 'message': 'Invalid Credentials'},
                            status=status.HTTP_404_NOT_FOUND)
        token, _ = Token.objects.get_or_create(user=user)
        return Response({'status': 'Success', 'token': token.key},
                        status=status.HTTP_200_OK)


class Logout(APIView):
    
    def get(self, request):
        request.user.auth_token.delete()
        return Response("User Logged out successfully")
class NotFoundView(APIView):
    def get(self, request, *args, **kwargs):
        return Response({"result": "Page not found", 'status': status.HTTP_404_NOT_FOUND})
    
def validate(date_text):
    try:
        datetime.strptime(date_text, '%Y-%m-%d')
        return True
    except ValueError:
        return False

#filter date view

class GetOrderdateView(APIView):
    permission_classes = (IsAuthenticated,)
    def get(self, request):
        try:
            limit = int(request.GET.get('limit', 50))
            offset = int(request.GET.get('offset', 0))
        except ValueError:
            return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

        date = str(request.GET.get('date', None))

        if date:
            if not validate(date):
                return Response({"result": "Invalid date format , should be  YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

            # Check for misspelled parameters
            valid_parameters = {'limit', 'offset', 'date'}
            provided_parameters = set(request.GET.keys())

            if not valid_parameters.issuperset(provided_parameters):
                return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

            try:
                order_details = rbi_main2.objects.filter(date_scraped__startswith=date).values('SI_NO','Type','Borrower','Economic_sector_of_borrower','Equivalent_Amount_in_USD','Purpose','Maturity_Period','Lender_Category','Month','Year','Route','date_scraped')[offset:limit]
                total_count = rbi_main2.objects.filter(date_scraped__startswith=date).count()
                if len(order_details) > 0:
                    return Response({"result": order_details,'total_count': total_count,  'status': status.HTTP_200_OK})
                else:
                    return Response({"result": "No Data Provided in your specific date !!!.", 'status': status.HTTP_401_UNAUTHORIZED})
            except TimeoutError:
                return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
            except Exception as err:
                return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
        else:
            raise Http404("Page not found")
class Custom404View(APIView):
    def get(self, request, *args, **kwargs):
        return Response({"result": "Page not found", 'status': status.HTTP_404_NOT_FOUND})


        

# To filter month and year

# class GetOrdermonthView(APIView):
#     permission_classes = (IsAuthenticated,)

#     def get(self, request):
#         limit = int(request.GET.get('limit', 2))
#         offset = int(request.GET.get('offset', 0))
#         month = str(request.GET.get('Month', None))
#         year = str(request.GET.get('Year', None))

#         if month and year:
#             try:
#                 order_details = rbi_main2.objects.filter(Month=month, Year=year).values('S_No', 'Type', 'Borrower', 'Economic_sector_of_borrower', 'Equivalent_Amount_in_USD', 'Purpose', 'Maturity_Period', 'Lender_Category', 'Month', 'Year', 'Route', 'date_scraped')[offset:limit]

#                 total_count = rbi_main2.objects.filter(Month=month, Year=year).count()

#                 if len(order_details) > 0:
#                     return Response({"result": order_details, 'total_count': total_count, 'status': HTTP_200_OK})
#                 else:
#                     return Response({"result": "No Data Provided !!!.", 'status': HTTP_400_BAD_REQUEST})
#             except Exception as err:
#                 return Response({"result": "Exception occurred: {}".format(err), 'status': HTTP_400_BAD_REQUEST})
#         else:
#             return Response({"result": "Month and Year are required!!!", 'status': HTTP_400_BAD_REQUEST})

  
 #To filter total view 
        
# class GetOrdertotalview(APIView):
#         permission_classes = (IsAuthenticated,)
#         def get(self, request):
#             limit = int(request.GET.get('limit', 2))
#             offset = int(request.GET.get('offset', 0))
#             try:
#             # Define the base queryset
#             #    base_queryset = rbi_main2.objects.all()
#                order_details = rbi_main2.objects.all().values('S_No', 'Type', 'Borrower', 'Economic_sector_of_borrower', 'Equivalent_Amount_in_USD', 'Purpose', 'Maturity_Period', 'Lender_Category', 'Month', 'Year', 'Route', 'date_scraped')[offset:limit]
#                total_count = rbi_main2.objects.all().count()
#                if len(order_details) > 0:
#                     return Response({"result": order_details, 'total_count': total_count, 'status': HTTP_200_OK})
#                else:
#                     return Response({"result": "No Data Provided !!!.", 'status': HTTP_400_BAD_REQUEST})
#             except Exception as err:
#                 return Response({"result": "Exception occurred: {}".format(err), 'status': HTTP_400_BAD_REQUEST})





             
        
        