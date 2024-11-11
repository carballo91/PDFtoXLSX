import os
import pdfplumber
import pandas as pd
import re
from django.shortcuts import render
from django.http import HttpResponse
from .forms import PDFUploadForm
from django.urls import reverse
from io import BytesIO
from .helpers import PDFEditor

def upload_pdf(request):
    if request.method == "POST":
        form = PDFUploadForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_files = request.FILES.getlist('pdf_file')  # Get the list of uploaded files
            download_urls = []  # To store download URLs
            filenames = []

            for pdf_file in pdf_files:
                # Create an instance of the PDFEditor class
                pdf_editor = PDFEditor(pdf_file)

                # Check if the uploaded file is a valid PDF
                if not pdf_editor.is_valid_pdf():
                    return render(request, 'upload.html', {'form': form, 'message': True})

                # Extract text from the first page (or relevant page) to check for "Run Date"
                first_page_text = pdf_editor.extract_text()
                decoded = pdf_editor.processText(first_page_text)

                df = None
                # Determine which PDF processing method to call based on the presence of "Run Date"
                if "Earned Commission Statement" in first_page_text:
                    # Process using method for the PDF with "Run Date" and "Agents"
                    df, output_name = pdf_editor.process_pdf_type1()
                elif "Foresters Financial" in first_page_text:
                    # Process using another method for different PDF structures
                    df, output_name = pdf_editor.forester_financial()  # Adjust for another type of PDF
                
                elif "SENIOR ADVISOR SERVICES AND\nINSURANCE SERVICES, LLC" in first_page_text:
                    df, output_name = pdf_editor.assurity_commission()
                
                elif "CURR JNT POLICY #  NAME PLAN ANNIV YR PREM RATE EARNINGS EXPLANATION PAY DISTR AMT DISTR TO DISTR FR" in decoded:
                    df, output_name = pdf_editor.kansas_city_life()
                elif "Agent # Writing Agent Name Policy # Name St Plan Code Mo/Yr Paid Date Dur. Premium Rate Comm Advance" in first_page_text:
                    df, output_name = pdf_editor.sentinel()
                elif "Member ID Name Company Product HICN Override Date Date Period Year Retro Amount" in first_page_text:
                    df, output_name = pdf_editor.bcbs_la_commisions()
                elif "Current ContractSubscriber Name Company MOP OED Due Date Product Name Premium Elapsed Comm. % Commission" in first_page_text:
                    df, output_name = pdf_editor.bcbs_la_compensation()
                else:
                    df, output_name = pdf_editor.essence_file()

                if df is None:
                    return render(request, 'upload.html', {'form': form, 'message': True})

                # Save the DataFrame to Excel and get the file path
                file_path, filename = pdf_editor.save_to_excel(df, output_name)

                # Create the download link
                download_url = reverse("download_file", args=[output_name])
                download_urls.append(download_url)
                filenames.append(filename)

            download_data = zip(download_urls, filenames)
            return render(request, 'upload.html', {'download_data': download_data, 'form': form})

    else:
        form = PDFUploadForm()

    return render(request, 'upload.html', {'form': form})

from django.http import FileResponse

def download_file(request,filename):
    file_path = os.path.join("/tmp",f"{filename}.xlsx")
    if os.path.exists(file_path):
        return FileResponse(open(file_path, 'rb'), as_attachment=True, filename=f'{filename}.xlsx')
    else:
        return HttpResponse("File not found", status=404)
