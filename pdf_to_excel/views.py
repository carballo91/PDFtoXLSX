import os
from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from .forms import PDFUploadForm
from django.urls import reverse
from .helpers import PDFEditor
from io import BytesIO
import zipfile
import urllib
import time

def upload_pdf(request):
    if request.method == "POST":
        form = PDFUploadForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_files = request.FILES.getlist('pdf_file')  # Get the list of uploaded files
            filenames = []
            
            
            for pdf_file in pdf_files:
                # Create an instance of the PDFEditor class
                pdf_editor = PDFEditor(pdf_file)
                
                # Check if the uploaded file is a valid PDF
                if not pdf_editor.is_valid_pdf():
                    return render(request, 'upload.html', {'form': form, 'message': True})

                # start_time = time.time()
                # Extract text from the PDF and determine processing method
                try:
                    first_page_text = pdf_editor.extract_text(pages=2)
                except IndexError:
                    first_page_text = pdf_editor.extract_text(pages=1)
                decoded = pdf_editor.processText(first_page_text)
                # print(first_page_text)
                #print(first_page_text)
                df = None
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
                elif "Member ID Writing ID Name Product State Date Term Date Term Code Period Type Retro Amount" in first_page_text:
                    df, output_name = pdf_editor.essence_file()
                elif "Blue Shield of California" in first_page_text: 
                    df,output_name = pdf_editor.blueshield_of_california()
                elif "Member ID Name Line of BusinessProduct MBI Effective Term Date Signed Date Period Cycle Retro ?Commissio" in first_page_text:
                    df,output_name = pdf_editor.providence()
                elif "Policy Number Insured Name Issue Policy Type Issue Date Commission Reason Date Chargeback Producer Commission" in first_page_text or "Policy Insured Name Issue Policy Type Issue Base Rate Reason Date Chargeback Producer Commission" in first_page_text or "Policy Insured Name Issue Policy Issue Date Base Rate Reason Date Chargeback Producer Commission" in first_page_text or "Policy Number Insured Name Issue Policy Type Issue Date Base Rate Reason Date Chargeback Producer Commission" in first_page_text:
                    df, output_name = pdf_editor.cincinatti()
                else:
                    df,output_name = pdf_editor.polish_falcons()
                # Add other conditions as needed...
                if df is None:
                    return render(request, 'upload.html', {'form': form, 'message': True})

                # Save the DataFrame to Excel and store the file path
                filename = pdf_editor.save_to_excel(df, output_name)
                filenames.append(filename)
                #print(f"Filename is {filename}")
                # end_time = time.time()
                # print(f"Time taken to run everything : {end_time - start_time:.2f} seconds")
            if len(filenames) > 1:
                # Multiple files - create a zip file
                in_memory_zip = BytesIO()
                with zipfile.ZipFile(in_memory_zip, 'w') as zipf:
                    for file_path in filenames:
                        zipf.write(file_path, os.path.basename(file_path))
                in_memory_zip.seek(0)

                # Save the zip file to temporary storage and generate download URL
                zip_filename = "processed_pdfs.zip"
                zip_path = os.path.join("/tmp", zip_filename)
                with open(zip_path, 'wb') as f:
                    f.write(in_memory_zip.getvalue())

                download_url = reverse("download_file", args=[urllib.parse.quote(zip_filename)])
                name = download_url.split("/")[2]
            else:
                # Single file, no zip needed
                single_file_name = os.path.basename(filenames[0])
                download_url = reverse("download_file", args=[urllib.parse.quote(single_file_name)])
                name = download_url.split("/")[2]

            return render(request, 'upload.html', {'download_url': download_url, 'form': form,"name":name})
    else:
        form = PDFUploadForm()

    return render(request, 'upload.html', {'form': form})


def download_file(request, filename):
    # Decode the filename and construct the file path
    file_path = os.path.join("/tmp", urllib.parse.unquote(filename))
    if os.path.exists(file_path):
        return FileResponse(open(file_path, 'rb'), as_attachment=True, filename=filename)
    else:
        return HttpResponse("File not found", status=404)
