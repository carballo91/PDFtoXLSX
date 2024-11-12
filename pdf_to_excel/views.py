import os
from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from .forms import PDFUploadForm
from django.urls import reverse
from .helpers import PDFEditor
import urllib.parse
from django.urls import reverse


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
                elif "Member ID Writing ID Name Product State Date Term Date Term Code Period Type Retro Amount" in first_page_text:
                    df, output_name = pdf_editor.essence_file()

                if df is None:
                    return render(request, 'upload.html', {'form': form, 'message': True})

                # Save the DataFrame to Excel and get the file path
                filename = pdf_editor.save_to_excel(df, output_name)
                filenames.append(filename)
     
            if len(filenames) > 1: # Create the zip file containing all the Excel files 
   
                download_url = reverse("download_file", args=["processed_pdfs.zip"])
                name = download_url.split("/")[2]
                print(name)
            else: # Single file, do not create a zip 
                single_file_name = os.path.basename(filenames[0]) 
                download_url = reverse("download_file", args=[urllib.parse.quote(single_file_name)])
                name = download_url.split("/")[2]
                print(name)
            return render(request, 'upload.html', {'download_url': download_url, 'form': form,"name":name,})

    else:
        form = PDFUploadForm()

    return render(request, 'upload.html', {'form': form})

def download_file(request, filename):
    file_path = os.path.join("/tmp", urllib.parse.unquote(filename))
    if os.path.exists(file_path):
        return FileResponse(open(file_path, 'rb'), as_attachment=True, filename=filename)
    else:
        return HttpResponse("File not found", status=404)
