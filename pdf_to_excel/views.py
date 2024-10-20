import os
import pdfplumber
import pandas as pd
import re
from django.conf import settings
from django.shortcuts import render
from django.http import HttpResponse
from .forms import PDFUploadForm
from django.utils.html import format_html
from django.urls import reverse

def extract_text(pdf, start_page):
    page = pdf.pages[start_page]
    return page.extract_text()

def extract_large_table(pdf, start_page, agent_table_number):
    full_table = []
    table_found = False
    for page_num in range(start_page, len(pdf.pages)):
        page = pdf.pages[page_num]
        tables = page.extract_tables()
        if tables:
            if not table_found:
                full_table.extend(tables[agent_table_number])
                table_found = True
            elif len(tables) > agent_table_number:
                full_table.extend(tables[agent_table_number])
        elif table_found:
            break
    return full_table

from django.core.files.uploadedfile import InMemoryUploadedFile
from io import BytesIO

def upload_pdf(request):
    if request.method == "POST":
        form = PDFUploadForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_file = request.FILES['pdf_file']
            if not pdf_file.name.endswith(".pdf"):
                return render(request, 'upload.html', {'form': form, 'message': True})
            #Gets the name of the file and removes the pdf extension
            output_name = pdf_file.name.rstrip(".pdf")
            # Use BytesIO to handle the file in memory
            pdf_file_memory = BytesIO(pdf_file.read())
            
            with pdfplumber.open(pdf_file_memory) as pdf:
                text = extract_text(pdf, 4)
                
                agents = re.findall(r"Subtotals for Agent (\w+)\s+([A-Z,.\s]+)", text)
                run_date = re.findall(r"Run Date:\s(\d{2}/\d{2}/\d{4})",text)
                if not agents:
                    return render(request, 'upload.html', {'form': form, 'message': True})
                new_data, extra_cols = [], []

                for i, agent in enumerate(agents):
                    large_table = extract_large_table(pdf, 4, i)
                  
                    for row in large_table:
                        if re.search(r"^[A-Z]+[,-]", row[0]):
                            row_data = {
                                "Run Date": run_date[0],
                                "Carrier": "Royal Neighbors",
                                "Agent Name": agent[1],
                                "Agent ID": agent[0],
                                "Insured's Name": row[0] if len(row) > 0 else None,
                                "Certificate": row[1] if len(row) > 1 else None,
                                "Prod ID": row[2] if len(row) > 2 else None,
                                "Issue Date": row[3] if len(row) > 3 else None,
                                "Mode": row[4] if len(row) > 4 else None,
                                "Paid To Date": row[5] if len(row) > 5 else None,
                                "1st Yr Rnwl": row[6] if len(row) > 6 else None,
                                "Split%": row[7] if len(row) > 7 else None,
                                "Prem": row[8] if len(row) > 8 else None,
                                "Comm%": row[9] if len(row) > 9 else None,
                                "Earned": row[10] if len(row) > 10 else None,
                                "Applied To Advance": row[11] if len(row) > 11 else None,
                                "Amt To Pay": row[12] if len(row) > 12 else None,
                            }
                            new_data.append(row_data)
                        if re.search(r"^\$", row[0]):
                            extra_cols.append(row)

                for d, e in zip(new_data, extra_cols):
                    d["Cert Adv Balance"] = e[0]
                    d["Comment"] = e[1]
                #Column for hyperlink
                for i in range(1):
                    new_data[i]["Converted from .pdf by"] = ""

                # Convert to pandas DataFrame
                df = pd.DataFrame(new_data)
           
                # Use a BytesIO stream to store the Excel output in memory
                output = BytesIO()

                # Write the DataFrame to the BytesIO stream
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False,sheet_name="Sheet1")
                    worksheet = writer.sheets["Sheet1"]
                    # Add hyperlinks in the "Converted from .pdf by" column
                    website_url = "https://comtrack.io"  # The URL you want the hyperlink to point to
                    
                    for row_num in range(1, 2):  # Start from 1 to skip the header
                        # Insert hyperlink in the 'Converted from .pdf by' column (assuming it's the second-last column)
                        worksheet.write_url(row_num, df.columns.get_loc('Converted from .pdf by'), 
                            website_url, string="ComTrack.io")

                # Ensure the pointer is at the start of the stream
                output.seek(0)
                # Save the file to a temporary location
                filename = f'{output_name}.xlsx'
                file_path = os.path.join('/tmp', filename)

                # Write the BytesIO content to the file
                with open(file_path, 'wb') as file:
                    file.write(output.read())

                # Create the download link
                download_url = reverse("download_file",args=[output_name])

            return render(request, 'upload.html', {'download_url': download_url})

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
