from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework import status
from rest_framework.permissions import IsAuthenticated, AllowAny
from rest_framework.authtoken.models import Token
from django.contrib.auth import authenticate
from .models import *
from datetime import datetime
from django.http import Http404
from rest_framework import status
import zipfile
from io import BytesIO
from django.shortcuts import render, HttpResponse
import json
from django.http import JsonResponse,request

from django.http import HttpResponse
import json
import zipfile
from io import BytesIO
from django.urls import reverse
from django.views import View
import io
import csv

from django.http import HttpResponse
import zipfile
from io import BytesIO
from django.views import View
from .models import rbi_odi,sebi_orders
from django.core.serializers.json import DjangoJSONEncoder


from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.http import Http404
from django.core.exceptions import ValidationError
# from django.core.validators import validate_date


import os
from django.conf import settings
from datetime import datetime
import json

from urllib.parse import urlencode


from rest_framework import renderers
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework import status

from urllib.parse import unquote
from django.http import QueryDict



# Define the custom HTML renderer
class HTMLRenderer(renderers.BaseRenderer):
    media_type = 'text/html'
    format = 'html'
    charset = 'utf-8'

    def render(self, data, accepted_media_type=None, renderer_context=None):
        html_content = '<html><body>'
        for entry in data['result']:
            html_content += f'<p>Date: {entry["date_of_order"]}</p>'
            html_content += f'<p>Title: {entry["title_of_order"]}</p>'
            html_content += f'<p>Type Of Order: {entry["type_of_order"]}</p>'
            html_content += f'<p>Link to Order: {entry["link_to_order"]}</p>'
            html_content += f'<p>PDF File Path: {entry["pdf_file_path"]}</p>'
            html_content += f'<p>PDF File Name: {entry["pdf_file_name"]}</p>'
            html_content += f'<p>Date Scraped: {entry["date_scraped"]}</p>'
            html_content += '<hr />'
            
         # Add the total PDF download link to the HTML content
        html_content += f'<p>{data["total_pdf_download_link"]}</p>'    
        html_content += '</body></html>'
        return html_content.encode(self.charset)


# Define your API view
from django.http import QueryDict

# class GetOrderdateView(APIView):
#     renderer_classes = [HTMLRenderer]  # Set the custom renderer

#     def get(self, request, *args, **kwargs):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#         date = str(request.GET.get('date', None))
#         whether = kwargs.get('whether')  # Extract 'whether' parameter from URL

#         if date:
#             try:
#                 validate_date(date)
#             except ValidationError:
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters"}, status=status.HTTP_400_BAD_REQUEST)

#             # Filter queryset based on 'whether' parameter
#             if whether == 'ed_cgm':
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='ed_cgm').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
#             elif whether == 'ao_cgm':
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='ao_cgm').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
#             else:
#                 return Response({"result": "Invalid 'whether' parameter"}, status=status.HTTP_400_BAD_REQUEST)

#             total_count = order_details.count()

#             if len(order_details) > 0:
#                 # Modify 'pdf_file_name' field to include a hyperlink for downloading the PDF file
#                 for entry in order_details:
#                     # Construct the URL to the Download PDF view with the PDF filename as a parameter
#                     download_pdf_url = reverse('download_pdf', kwargs={'filename': entry['pdf_file_name']})
#                     entry['pdf_file_name'] = f'<a href="{download_pdf_url}">Download PDF</a>'

#                 # Construct the URL for the "Download All PDFs" view
#                 download_all_pdfs_url = reverse('download_all_pdfs', kwargs={'date': date, 'whether': whether})
                
                
#                # Extract PDF filenames from order_details
#                 pdf_filenames = [entry['pdf_file_name'] for entry in order_details]

#                 # Remove HTML tags and extract filenames
#                 cleaned_pdf_filenames = []
#                 for filename in pdf_filenames:
#                     # Split filename by '/' to extract the last part (filename)
#                     parts = filename.split('/')
#                     if parts:
#                         filename = parts[-1].strip()  # Extract filename and remove leading/trailing whitespace
                        
#                         # Remove any remaining HTML tags
#                         filename = filename.replace('>', '').replace('<', '')
                        
#                         cleaned_pdf_filenames.append(filename)

#                 # Construct the "Download All PDFs" link with cleaned PDF filenames
#                 download_all_pdfs_url = reverse('download_all_pdfs', kwargs={'date': date, 'whether': whether})
#                 download_all_pdfs_url += f'?pdf_filenames={",".join(cleaned_pdf_filenames)}'

            
#                 # # Construct the query parameters for the "Download All PDFs" link
#                 # pdf_filenames = [entry['pdf_file_name'] for entry in order_details]
#                 # query_parameters = QueryDict(mutable=True)
#                 # query_parameters.update({'pdf_filenames': pdf_filenames})
#                 # download_all_pdfs_url += '?' + query_parameters.urlencode()

#                 # Create the "Download All PDFs" link
#                 total_pdf_download_link = f'<a href="{download_all_pdfs_url}">Download All PDFs</a>'

#                 # Return JSON response with results including the total PDF download link
#                 return Response({"result": order_details, 'total_count': total_count, 'total_pdf_download_link': total_pdf_download_link}, status=status.HTTP_200_OK)
#             else:
#                 return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
#         else:
#             raise Http404("Page not found")

# class DownloadAllPDFsView(View):
#     def get(self, request, date, whether):
#         try:
#             validate_date(date)

#             # Retrieve PDF filenames from request
#             pdf_filenames = request.GET.getlist('pdf_filenames', [])
#             if not pdf_filenames:
#                 return HttpResponse("No PDFs specified for download.", status=400)

#             # Extract cleaned PDF filenames from HTML links
#             cleaned_pdf_filenames = []
#             for filename in pdf_filenames:
#                 cleaned_filename = extract_clean_filename(filename)
#                 if cleaned_filename:
#                     cleaned_pdf_filenames.append(cleaned_filename)

#             print("Cleaned PDF Filenames:", cleaned_pdf_filenames)  # Debugging statement

#             # Filter queryset based on 'whether' parameter and cleaned filenames
#             if whether in ['ed_cgm', 'ao_cgm']:
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order=whether, pdf_file_name__in=cleaned_pdf_filenames).values('pdf_file_name')
                
#                 print("Order Details:", order_details)  # Debugging statement

#                 # Retrieve PDF file paths from the queryset
#                 pdf_paths = [os.path.join(settings.MEDIA_ROOT, entry['pdf_file_name']) for entry in order_details if entry['pdf_file_name']]
                
#                 print("PDF Paths:", pdf_paths)  # Debugging statement
                
#                 if pdf_paths:
#                     # Create an HttpResponse object with content type as zip
#                     response = HttpResponse(content_type='application/zip')

#                     # Set the zip file name
#                     response['Content-Disposition'] = 'attachment; filename="pdf_files.zip"'

#                     # Create a zip file
#                     with zipfile.ZipFile(response, 'w') as zip_file:
#                         # Add each PDF file to the zip
#                         for pdf_path in pdf_paths:
#                             if os.path.exists(pdf_path):  # Check if the file exists before adding to zip
#                                 zip_file.write(pdf_path, arcname=os.path.basename(pdf_path))

#                     return response
#                 else:
#                     # If no PDF files were found, return an empty response
#                     return HttpResponse("No PDF files found.", status=404)
#             else:
#                 return HttpResponse("Invalid 'whether' parameter.", status=400)
#         except ValidationError:
#             return HttpResponse("Incorrect date format, should be YYYY-MM-DD", status=status.HTTP_422_UNPROCESSABLE_ENTITY)
#         except Exception as e:
#             print("Error:", e)  # Debugging statement
#             return HttpResponse("An error occurred.", status=500)

# def extract_clean_filename(filename):
#     # Split by '/' to get the last part containing the filename
#     filename_parts = filename.split('/')[-1]
#     # Split by '>' to get the filename part after the last '>'
#     filename_parts = filename_parts.split('>')[-1]
#     # Split by '<' to remove any trailing characters and get the clean filename
#     cleaned_filename = filename_parts.split('<')[0].strip()
#     return cleaned_filename






def validate_date(date_string):
    try:
        datetime.strptime(date_string, '%Y-%m-%d')
        return True
    except:
       return False





# Define your API view
class GetOrderdateView(APIView):
    renderer_classes = [HTMLRenderer]  # Set the custom renderer
    # renderer_classes = [renderers.JSONRenderer, renderers.BrowsableAPIRenderer]

    def get(self, request, *args, **kwargs):
        try:
            limit = int(request.GET.get('limit', 50))
            offset = int(request.GET.get('offset', 0))
        except ValueError:
            return Response({"result": "Invalid limit or offset value, must be an integer"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

        date = str(request.GET.get('date', None))
        whether = kwargs.get('whether')  # Extract 'whether' parameter from URL

        if date:
            try:
                validate_date(date)
            except ValidationError:
                return Response({"result": "Incorrect date format, should be YYYY-MM-DD"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

            valid_parameters = {'limit', 'offset', 'date'}
            provided_parameters = set(request.GET.keys())

            if not valid_parameters.issuperset(provided_parameters):
                return Response({"result": "Invalid query parameters, check spelling for given parameters"}, status=status.HTTP_400_BAD_REQUEST)

            # Filter queryset based on 'whether' parameter
            if whether == 'ed_cgm':
                order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='ed_cgm').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
            elif whether == 'ao_cgm':
                order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='ao_cgm').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
            # elif whether == 'settlementorder':
            #     order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='settlementorder').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
            
            else:
                return Response({"result": "Invalid 'whether' parameter"}, status=status.HTTP_400_BAD_REQUEST)

            total_count = order_details.count()
            
            
            if len(order_details) > 0:
                # Modify 'pdf_file_name' field to include a hyperlink for downloading the PDF file
                for entry in order_details:
                    # Construct the URL to the Download PDF view with the PDF filename as a parameter
                    download_pdf_url = reverse('download_pdf', kwargs={'filename': entry['pdf_file_name']})
                    entry['pdf_file_name'] = f'<a href="{download_pdf_url}">Download PDF</a>'
                
                
                #Create a total download link for all PDFs displayed
                # total_pdf_download_url = reverse('download_all_pdfs', kwargs={'date': date, 'whether': whether})
                total_pdf_download_url = reverse('download_all_pdfs', kwargs={'date': date, 'whether': whether}) + f'?limit={limit}&offset={offset}'

                total_pdf_download_link = f'<a href="{total_pdf_download_url}">Download All PDFs</a>'
                
                # Return JSON response with results including the total PDF download link
                return Response({"result": order_details, 'total_count': total_count, 'total_pdf_download_link': total_pdf_download_link}, status=status.HTTP_200_OK)
            else:
                return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
        else:
            raise Http404("Page not found")


class DownloadPDFView(View):
    def get(self, request, filename):
        pdf_path = os.path.join(settings.MEDIA_ROOT, filename)
        # 

        if os.path.exists(pdf_path):
            response = HttpResponse(content_type='application/zip')
            zip_file = zipfile.ZipFile(response, 'w')

            # Add the PDF file to the ZIP archive
            zip_file.write(pdf_path, os.path.basename(pdf_path))

            # Close the ZIP file
            zip_file.close()

            # Set the appropriate content-disposition header for the response
            response['Content-Disposition'] = f'attachment; filename="{filename}.zip"'
            return response
        else:
            # Handle case when PDF file doesn't exist
            return HttpResponse("PDF file not found.", status=404)
        
        
class DownloadAllPDFsView(View):
    def get(self, request, date, whether):
        try:
            # Extract the 'limit' parameter from the request.GET dictionary
            limit = int(request.GET.get('limit', 50))
            offset = int(request.GET.get('offset', 0))
            validate_date(date)

            valid_parameters = {'limit', 'offset', 'date'}  # Include limit and offset in valid parameters
            provided_parameters = set(request.GET.keys())

            if not valid_parameters.issuperset(provided_parameters):
                return HttpResponse("Invalid query parameters, check spelling for given parameters", status=status.HTTP_400_BAD_REQUEST)

            # Filter queryset based on 'whether' parameter
            if whether in ['ed_cgm', 'ao_cgm']:
                # Query the necessary number of PDFs based on both limit and offset parameters
                order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order=whether).values('pdf_file_name')[offset:offset+limit]

                # Filter the PDFs based on sr_no values
                pdf_paths = [os.path.join(settings.MEDIA_ROOT, entry['pdf_file_name']) for entry in order_details if entry['pdf_file_name']]

                if pdf_paths:
                    # Create an HttpResponse object with content type as zip
                    response = HttpResponse(content_type='application/zip')

                    # Set the zip file name
                    response['Content-Disposition'] = 'attachment; filename="pdf_files.zip"'

                    # Create a zip file
                    with zipfile.ZipFile(response, 'w') as zip_file:
                        # Add each PDF file to the zip
                        for pdf_path in pdf_paths:
                            zip_file.write(pdf_path, arcname=os.path.basename(pdf_path))

                    return response
                else:
                    # If no PDF files were found, return an empty response
                    return HttpResponse("No PDF files found.", status=404)
            else:
                return HttpResponse("Invalid 'whether' parameter.", status=400)
        except ValueError:
            return HttpResponse("Invalid limit or offset value, must be an integer", status=status.HTTP_422_UNPROCESSABLE_ENTITY)
        except ValidationError:
            return HttpResponse("Incorrect date format, should be YYYY-MM-DD", status=status.HTTP_422_UNPROCESSABLE_ENTITY)
        except Exception as e:
            print("Error:", e)  # Debugging statement
            return HttpResponse("An error occurred.", status=500)
        


# class DownloadAllPDFsView(View):
#     def get(self, request, date, whether):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#             validate_date(date)

#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return HttpResponse("Invalid query parameters, check spelling for given parameters", status=status.HTTP_400_BAD_REQUEST)

#             # Filter queryset based on 'whether' parameter
#             if whether in ['ed_cgm', 'ao_cgm']:
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order=whether).values('pdf_file_name')[offset:limit]

#                 # Retrieve PDF file paths from the queryset
#                 pdf_paths = [os.path.join(settings.MEDIA_ROOT, entry['pdf_file_name']) for entry in order_details if entry['pdf_file_name']]

#                 if pdf_paths:
#                     # Create an HttpResponse object with content type as zip
#                     response = HttpResponse(content_type='application/zip')

#                     # Set the zip file name
#                     response['Content-Disposition'] = 'attachment; filename="pdf_files.zip"'

#                     # Create a zip file
#                     with zipfile.ZipFile(response, 'w') as zip_file:
#                         # Add each PDF file to the zip
#                         for pdf_path in pdf_paths:
#                             zip_file.write(pdf_path, arcname=os.path.basename(pdf_path))

#                     return response
#                 else:
#                     # If no PDF files were found, return an empty response
#                     return HttpResponse("No PDF files found.", status=404)
#             else:
#                 return HttpResponse("Invalid 'whether' parameter.", status=400)
#         except ValueError:
#             return HttpResponse("Invalid limit or offset value, must be an integer", status=status.HTTP_422_UNPROCESSABLE_ENTITY)
#         except ValidationError:
#             return HttpResponse("Incorrect date format, should be YYYY-MM-DD", status=status.HTTP_422_UNPROCESSABLE_ENTITY)
#         except Exception as e:
#             print("Error:", e)  # Debugging statement
#             return HttpResponse("An error occurred.", status=500)









# class DownloadAllPDFsView(View):
#     def get(self, request, date, whether):
#         try:
#             # Retrieve the PDF file paths based on the provided date and whether parameters
#             if whether in ['ed_cgm', 'ao_cgm']:
#                 # Filter sebi_orders queryset based on the provided date and whether parameters
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order=whether).values('pdf_file_name')
                
#                 # Retrieve PDF file paths from the queryset
#                 pdf_paths = [os.path.join(settings.MEDIA_ROOT, entry['pdf_file_name']) for entry in order_details if entry['pdf_file_name']]

#                 if pdf_paths:
#                     # Create an HttpResponse object with content type as zip
#                     response = HttpResponse(content_type='application/zip')

#                     # Set the zip file name
#                     response['Content-Disposition'] = 'attachment; filename="pdf_files.zip"'

#                     # Create a zip file
#                     with zipfile.ZipFile(response, 'w') as zip_file:
#                         # Add each PDF file to the zip
#                         for pdf_path in pdf_paths:
#                             zip_file.write(pdf_path, arcname=os.path.basename(pdf_path))

#                     return response
#                 else:
#                     # If no PDF files were found, return an empty response
#                     return HttpResponse("No PDF files found.", status=404)
#             else:
#                 return HttpResponse("Invalid 'whether' parameter.", status=400)
#         except Exception as e:
#             print("Error:", e)  # Debugging statement
#             return HttpResponse("An error occurred.", status=500)









class DownloadZipView(APIView):
    def get(self, request, *args, **kwargs):
        date = request.GET.get('date')
        whether = kwargs.get('whether')

        # Fetch PDF file names from the API response
        pdf_file_names = request.GET.getlist('pdf_file_name')

        # Create a zip file containing the PDF files
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
            for pdf_file_name in pdf_file_names:
                pdf_file_path = os.path.join(settings.MEDIA_ROOT, pdf_file_name.strip('/'))
                zip_file.write(pdf_file_path, os.path.basename(pdf_file_path))

        # Prepare response with zip file
        response = HttpResponse(zip_buffer.getvalue(), content_type='application/zip')
        response['Content-Disposition'] = f'attachment; filename={date}_orders.zip'
        return response












# class GetOrderdateView(APIView):
#     def get(self, request, *args, **kwargs):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#         date = str(request.GET.get('date', None))
#         whether = kwargs.get('whether')  # Extract 'whether' parameter from URL

#         if date:
#             try:
#                 validate_date(date)
#             except ValidationError:
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters"}, status=status.HTTP_400_BAD_REQUEST)

#             # Filter queryset based on 'whether' parameter
#             if whether == 'ed_cgm':
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='ed_cgm').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
#             elif whether == 'ao_cgm':
#                 order_details = sebi_orders.objects.filter(date_scraped__startswith=date, type_of_order='ao_cgm').values('sr_no', 'date_of_order',  'title_of_order', 'type_of_order', 'link_to_order',  'pdf_file_path', 'pdf_file_name', 'updated_date',  'date_scraped')[offset:limit]
#             else:
#                 return Response({"result": "Invalid 'whether' parameter"}, status=status.HTTP_400_BAD_REQUEST)

#             total_count = order_details.count()

#             if len(order_details) > 0:
#                 # Convert order_details to JSON
#                 order_details_json = json.dumps(list(order_details), default=str)

#                 # Create a BytesIO object to store zip file contents
#                 zip_buffer = BytesIO()
#                 with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
#                     # Add order_details as a JSON file in the zip archive
#                     zip_file.writestr("order_details.json", order_details_json)
                
#                 # Set the buffer's file pointer to the beginning
#                 zip_buffer.seek(0)
                
#                 # Create download link for the zip file
#                 download_zipfile = request.build_absolute_uri('/api/v1/{}/download_zip/?date={}'.format(whether, date))   # Assuming download URL is 'download/'

#                 # Return JSON response with results and download link
#                 return Response({"result": order_details, 'total_count': total_count, 'download_zipfile': download_zipfile}, status=status.HTTP_200_OK)
#             else:
#                 return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
#         else:
#             raise Http404("Page not found")
          

# class GetOrderdateView(APIView):
#     def get(self, request, *args, **kwargs):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#         date = str(request.GET.get('date', None))
#         whether = kwargs.get('whether')  # Extract 'whether' parameter from URL

#         if date:
#             try:
#                 validate_date(date)
#             except ValidationError:
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters"}, status=status.HTTP_400_BAD_REQUEST)

#             # Filter queryset based on 'whether' parameter
#             if whether == 'WOS':
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='WOS').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#             elif whether == 'JV':
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='JV').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#             else:
#                 return Response({"result": "Invalid 'whether' parameter"}, status=status.HTTP_400_BAD_REQUEST)

#             total_count = order_details.count()

#             if len(order_details) > 0:
#                 # Convert order_details to JSON
#                 order_details_json = json.dumps(list(order_details), default=str)

#                 # Create a BytesIO object to store zip file contents
#                 zip_buffer = BytesIO()
#                 with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
#                     # Add order_details as a JSON file in the zip archive
#                     zip_file.writestr("order_details.json", order_details_json)
                
#                 # Set the buffer's file pointer to the beginning
#                 zip_buffer.seek(0)
                
#                 # Create download link for the zip file
#                 download_zipfile = request.build_absolute_uri('/api/v1/{}/download_zip/?date={}'.format(whether, date))   # Assuming download URL is 'download/'

#                 # Return JSON response with results and download link
#                 return Response({"result": order_details, 'total_count': total_count, 'download_zipfile': download_zipfile}, status=status.HTTP_200_OK)
#             else:
#                 return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
#         else:
#             raise Http404("Page not found")
        
 
# # class DownloadZipView(APIView):
#     def get(self, request, *args, **kwargs):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#         date = str(request.GET.get('date', None))
#         whether = kwargs.get('whether')  # Extract 'whether' parameter from URL

#         if date:
#             try:
#                 validate_date(date)
#             except ValidationError:
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD"}, status=status.HTTP_422_UNPROCESSABLE_ENTITY)

#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters"}, status=status.HTTP_400_BAD_REQUEST)

#             # Filter queryset based on 'whether' parameter
#             if whether == 'WOS':
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='WOS').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#             elif whether == 'JV':
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='JV').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#             else:
#                 return Response({"result": "Invalid 'whether' parameter"}, status=status.HTTP_400_BAD_REQUEST)

#             total_count = order_details.count()

#             if len(order_details) > 0:
#                 # Create a BytesIO object to store zip file contents
#                 # zip_buffer = BytesIO()
#                 # with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
#                 #     for entry in order_details:
#                 #         # Serialize entry to JSON string
#                 #         entry_json = json.dumps(entry, default=str)
#                 #         # Add entry as file in the zip archive
#                 #         zip_file.writestr(f"{entry['SI_NO']}.json", entry_json)
                
#                 # # Set the buffer's file pointer to the beginning
#                 # zip_buffer.seek(0)

                
#                 # Serialize queryset to JSON
#                 json_data = json.dumps(list(order_details), cls=DjangoJSONEncoder, indent=4)
                
#                 # Create a BytesIO object to hold the zip file in memory
#                 buffer = io.BytesIO()
#                 with zipfile.ZipFile(buffer, 'w') as zip_file:
#                     # Write JSON data
#                     zip_file.writestr('data.json', json_data)
                
#                 buffer.seek(0)
#                 # Prepare respons
                
                
#                 # Return the zip file as response
#                 response = HttpResponse(buffer, content_type='application/zip')
#                 response['Content-Disposition'] = 'attachment; filename="order_details.zip"'
#                 return response
#             else:
#                 return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
#         else:
#             raise Http404("Page not found")   
        
            

# class GetOrderdateView(APIView):
#     def get(self, request, *args, **kwargs):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

#         date = str(request.GET.get('date', None))
#         whether = self.kwargs.get('whether')  # Extract 'whether' parameter from URL

#         if date:
#             if not validate_date(date):
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

#             try:
#                 # Filter queryset based on 'whether' parameter
#                 if whether == 'WOS':
#                     order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='WOS').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#                     total_count = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='WOS').count()
#                 elif whether == 'JV':
#                     order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='JV').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#                     total_count = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='JV').count()
#                 else:
#                     return Response({"result": "Invalid 'whether' parameter", 'status': status.HTTP_400_BAD_REQUEST})

#                 if len(order_details) > 0:
#                     return Response({"result": order_details, 'total_count': total_count, 'status': status.HTTP_200_OK})
#                 else:
#                     return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
#             except TimeoutError:
#                 return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
#             except Exception as err:
#                 return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
#         else:
#             raise Http404("Page not found")

 
    
# class GetOrderdateView(APIView):
#     # permission_classes = (IsAuthenticated,)
    
#     def get(self, request):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#             date = request.GET.get('date')
#             whether = request.GET.get('whether')  # New parameter

#             if not validate_date(date):
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})
            
#             if whether not in ['WOS', 'JV']:
#                 return Response({"result": "Invalid 'whether' parameter", 'status': status.HTTP_400_BAD_REQUEST})
            
#             valid_parameters = {'limit', 'offset', 'date', 'whether'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

#             # Filter queryset based on 'whether' parameter
#             if whether == 'WOS':
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='WOS').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#             elif whether == 'JV':
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date, Whether_JV_WOS='JV').values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
            
#             total_count = order_details.count()
#             for entry in order_details:
#                 for key, value in entry.items():
#                     if value == "":
#                         entry[key] = "Null"
#             if len(order_details) > 0:
#                 return Response({"result": order_details, 'total_count': total_count, 'status': status.HTTP_200_OK})
#             else:
#                 return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
        
#         except TimeoutError:
#             return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
#         except Exception as err:
#             return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})

        
class zipdownload(APIView):
    # permission_classes = (IsAuthenticated,)
    
    def get(self, request):
        try:
            limit = int(request.GET.get('limit', 50))
            offset = int(request.GET.get('offset', 0))
        except ValueError:
            return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

        date = str(request.GET.get('date', None))

        if date:
            if not validate(date):
                return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})
            
            valid_parameters = {'limit', 'offset', 'date'}
            provided_parameters = set(request.GET.keys())

            if not valid_parameters.issuperset(provided_parameters):
                return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

            try:
                order_details = rbi_odi.objects.filter(date_scraped__startswith=date).values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
                
                # Convert datetime objects to strings
                for entry in order_details:
                    entry['date_scraped'] = entry['date_scraped'].strftime('%Y-%m-%d %H:%M:%S')
                
                total_count = rbi_odi.objects.filter(date_scraped__startswith=date).count()
                
                # Serialize queryset to JSON
                json_data = json.dumps(list(order_details), cls=DjangoJSONEncoder, indent=4)
                
                # Create a BytesIO object to hold the zip file in memory
                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, 'w') as zip_file:
                    # Write JSON data
                    zip_file.writestr('data.json', json_data)
                
                buffer.seek(0)
                # Prepare response for downloading the zip file
                response = HttpResponse(buffer, content_type='application/zip')
                response['Content-Disposition'] = 'attachment; filename="data.zip"'
                return response

            except TimeoutError:
                return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
            except Exception as err:
                return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
        else:
            raise Http404("Page not found")




class downloadzip1(APIView):
    # permission_classes = (IsAuthenticated,)
    
    def get(self, request):
        try:
            limit = int(request.GET.get('limit', 50))
            offset = int(request.GET.get('offset', 0))
        except ValueError:
            return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

        date = str(request.GET.get('date', None))

        if date:
            if not validate(date):
                return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})
            
            valid_parameters = {'limit', 'offset', 'date'}
            provided_parameters = set(request.GET.keys())

            if not valid_parameters.issuperset(provided_parameters):
                return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

            try:
                order_details = rbi_odi.objects.filter(date_scraped__startswith=date).values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
                total_count = rbi_odi.objects.filter(date_scraped__startswith=date).count()
                
                # Create a BytesIO object to hold the zip file in memory
                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, 'w') as zip_file:
                    # Create a CSV file in the zip for the data
                    csv_data = io.StringIO()
                    csv_writer = csv.writer(csv_data)
                    csv_writer.writerow(order_details[0].keys())  # Write header
                    for entry in order_details:
                        csv_writer.writerow(entry.values())
                    zip_file.writestr('data.csv', csv_data.getvalue())
                
                buffer.seek(0)
                # Prepare response for downloading the zip file
                response = HttpResponse(buffer, content_type='application/zip')
                response['Content-Disposition'] = 'attachment; filename="data.zip"'
                return response

            except TimeoutError:
                return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
            except Exception as err:
                return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
        else:
            raise Http404("Page not found")
        


# class zipdownload(APIView):
#     # permission_classes = (IsAuthenticated,)
    
#     def get(self, request):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

#         date = str(request.GET.get('date', None))

#         if date:
#             if not validate(date):
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})
#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

#             try:
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date).values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#                 total_count = rbi_odi.objects.filter(date_scraped__startswith=date).count()
                
#                 # Create a BytesIO object to store zip file contents
#                 zip_buffer = BytesIO()
#                 with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
#                     for entry in order_details:
#                         # Serialize entry to JSON string
#                         entry_json = json.dumps(entry, default=str)
#                         # Add entry as file in the zip archive
#                         zip_file.writestr(f"{entry['SI_NO']}.json", entry_json)
                
#                 # Set the buffer's file pointer to the beginning
#                 zip_buffer.seek(0)
                
#                 # Create an HTTP response with the zip file as the content
#                 response = HttpResponse(zip_buffer, content_type='application/zip')
#                 response['Content-Disposition'] = 'attachment; filename="order_details.zip"'
                
#                 # Return the HTTP response
#                 return response
            
#             except TimeoutError:
#                 return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
#             except Exception as err:
#                 return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
#         else:
#             raise Http404("Page not found")


        
class Login(APIView):   
    # permission_classes = (AllowAny,)

    def post(self, request):
        username = request.data.get("username")
        password = request.data.get("password")

        if username is None or password is None:
            return Response({'message': 'Please provide both username and password'},
                            status.HTTP_400_BAD_REQUEST)
        user = authenticate(username=username, password=password)  
        if not user:
            return Response({'status': 'Failed', 'message': 'Invalid Credentials'},
                            status.HTTP_404_NOT_FOUND)
        token, _ = Token.objects.get_or_create(user=user)
        return Response({'status': 'Success', 'token': token.key},
                        status.HTTP_200_OK)


class Logout(APIView):
    

    def get(self, request):
        request.user.auth_token.delete()
        return Response("User Logged out successfully")

class Custom404View(APIView):
    def get(self, request, *args, **kwargs):
        return Response({"result": "Page not found", 'status': status.HTTP_404_NOT_FOUND})

def validate(date_text):
    try:
        datetime.strptime(date_text, '%Y-%m-%d')
        return True
    except:
        return False


# #filter date view

# class GetOrderdateView(APIView):
#     #  permission_classes = (IsAuthenticated,)
#      def get(self, request):
        
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

#         date = str(request.GET.get('date', None))

#         if date:
#             if not validate(date):
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})
#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

#             try:
                
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date).values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#                 total_count = rbi_odi.objects.filter(date_scraped__startswith=date).count()
#                 for entry in order_details:
#                     for key, value in entry.items():
#                         if value == "":
#                             entry[key] = "Null"
#                 if len(order_details) > 0:
#                     return Response({"result": order_details,'total_count': total_count,  'status': status.HTTP_200_OK})
#                 else:
#                     return Response({"result": "No Data Provided in your specific date!!!.", 'status': status.HTTP_401_UNAUTHORIZED})
#             except TimeoutError:
#                 return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
#             except Exception as err:
#                 return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
#         else:
#             raise Http404("Page not found")





# class GetOrderdateView(APIView):
#     # permission_classes = (IsAuthenticated,)
    
#     def get(self, request):
#         try:
#             limit = int(request.GET.get('limit', 50))
#             offset = int(request.GET.get('offset', 0))
#         except ValueError:
#             return Response({"result": "Invalid limit or offset value, must be an integer", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})

#         date = str(request.GET.get('date', None))

#         if date:
#             if not validate(date):
#                 return Response({"result": "Incorrect date format, should be YYYY-MM-DD", 'status': status.HTTP_422_UNPROCESSABLE_ENTITY})
#             valid_parameters = {'limit', 'offset', 'date'}
#             provided_parameters = set(request.GET.keys())

#             if not valid_parameters.issuperset(provided_parameters):
#                 return Response({"result": "Invalid query parameters, check spelling for given parameters", 'status': status.HTTP_400_BAD_REQUEST})

#             try:
#                 order_details = rbi_odi.objects.filter(date_scraped__startswith=date).values('SI_NO','Name_of_the_Indian_Party','Name_of_the_JV_WOS','Whether_JV_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#                 total_count = rbi_odi.objects.filter(date_scraped__startswith=date).count()
                
#                 # Create a BytesIO object to store zip file contents
#                 zip_buffer = BytesIO()
#                 with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
#                     for entry in order_details:
#                         # Serialize entry to JSON string
#                         entry_json = json.dumps(entry, default=str)
#                         # Add entry as file in the zip archive
#                         zip_file.writestr(f"{entry['SI_NO']}.json", entry_json)
                
#                 # Set the buffer's file pointer to the beginning
#                 zip_buffer.seek(0)
                
#                 # Return the zip file as response
#                 response = HttpResponse(zip_buffer, content_type='application/zip')
#                 response['Content-Disposition'] = 'attachment; filename="order_details.zip"'
#                 return response
            
#             except TimeoutError:
#                 return Response({"result": "timeout error", 'status': status.HTTP_502_BAD_GATEWAY})
#             except Exception as err:
#                 return Response({"result": f"An internal server error occurred: {err}", 'status': status.HTTP_500_INTERNAL_SERVER_ERROR})
#         else:
#             raise Http404("Page not found")





# To filter month and year

# class GetOrdermonthView(APIView):
#     permission_classes = (IsAuthenticated,)

#     def get(self, request):
#         limit = int(request.GET.get('limit', 5))
#         offset = int(request.GET.get('offset', 0))
#         month = str(request.GET.get('Month', None))
#         year = str(request.GET.get('Year', None))

#         if month and year:
#             try:
#                 order_details = rbi_odi.objects.filter(Month=month, Year=year).values('Sr_No','Name_of_the_Indian_Party','Name_of_the_JV_or_WOS','Whether_JV_or_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]

#                 total_count = rbi_odi.objects.filter(Month=month, Year=year).count()

#                 if len(order_details) > 0:
#                     return Response({"result": order_details, 'total_count': total_count, 'status': HTTP_200_OK})
#                 else:
#                     return Response({"result": "No Data Provided !!!.", 'status': HTTP_400_BAD_REQUEST})
#             except Exception as err:
#                 return Response({"result": "Exception occurred: {}".format(err), 'status': HTTP_400_BAD_REQUEST})
#         else:
#             return Response({"result": "Month and Year are required!!!", 'status': HTTP_400_BAD_REQUEST})
    
 
 
# To filter Total data  
        
# class GetOrdertotalview(APIView):
#         permission_classes = (IsAuthenticated,)
#         def get(self, request):
#             limit = int(request.GET.get('limit', 5))
#             offset = int(request.GET.get('offset', 0))
#             try:
#             # Define the base queryset
#             #    base_queryset = rbi_main2.objects.all()
#                order_details = rbi_odi.objects.all().values('Sr_No','Name_of_the_Indian_Party','Name_of_the_JV_or_WOS','Whether_JV_or_WOS','Overseas_Country','Major_Activity','FC_Equity','FC_Loan','FC_Guarantee_Issued','FC_Total','Month','Year','date_scraped')[offset:limit]
#                total_count = rbi_odi.objects.all().count()
#                if len(order_details) > 0:
#                     return Response({"result": order_details, 'total_count': total_count, 'status': HTTP_200_OK})
#                else:
#                     return Response({"result": "No Data Provided !!!.", 'status': HTTP_400_BAD_REQUEST})
#             except Exception as err:
#                 return Response({"result": "Exception occurred: {}".format(err), 'status': HTTP_400_BAD_REQUEST})