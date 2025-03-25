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
from pdfminer.pdfdocument import PDFPasswordIncorrect

def upload_pdf(request):
    if request.method == "POST":
        form = PDFUploadForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_files = request.FILES.getlist('pdf_file')  # Get the list of uploaded files
            filenames = []
            
            
            for pdf_file in pdf_files:
                # print(f"PDF File is {pdf_file}")
                # Create an instance of the PDFEditor class
                pdf_editor = PDFEditor(pdf_file)
                
                # Check if the uploaded file is a valid PDF
                if not pdf_editor.is_valid_pdf():
                    print("It is not valid")
                    return render(request, 'upload.html', {'form': form, 'message': True})

                try:
                    pdf_name = pdf_editor.pdf_output_name.split()[2]
                except IndexError:
                    pdf_name = pdf_editor.pdf_output_name
         
                
                # start_time = time.time()
                # Extract text from the PDF and determine processing method
                passwords = [None,"2646","WG500","LBL22728"]
                pw = pdf_name.rstrip("Z")
                passwords.append(pw)
                for password in passwords:
                    try:
                        first_page_text = pdf_editor.extract_text(pages=3,password=password)
                        break
                    except IndexError:
                        first_page_text = pdf_editor.extract_text(pages=1,password=password)
                        break
                    except PDFPasswordIncorrect:
                        continue
                decoded = pdf_editor.processText(first_page_text)
                # print(first_page_text)
                # print(first_page_text)
                
                df = None
                output_name = ""
  
                if "Royal Neighbors of America" in first_page_text:
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
                elif "Blue Shield of California" in first_page_text or "Group Number Subscriber ID Customer Name Product Eff Date Period Gross Premium Base Premium* Rate Commission Paid" in first_page_text: 
                    df,output_name = pdf_editor.blueshield_of_california()
                elif "Member ID Name Line of BusinessProduct MBI Effective Term Date Signed Date Period Cycle Retro ?Commissio" in first_page_text:
                    df,output_name = pdf_editor.providence()
                elif "Policy Number Insured Name Issue Policy Type Issue Date Commission Reason Date Chargeback Producer Commission" in first_page_text or "Policy Insured Name Issue Policy Type Issue Base Rate Reason Date Chargeback Producer Commission" in first_page_text or "Policy Insured Name Issue Policy Issue Date Base Rate Reason Date Chargeback Producer Commission" in first_page_text or "Policy Number Insured Name Issue Policy Type Issue Date Base Rate Reason Date Chargeback Producer Commission" in first_page_text:
                    df, output_name = pdf_editor.cincinatti()
                elif "Policy Insured/Anuitant Plan Date Mode Value Premium Age Year Agent Share Date Payment Percent Earned Advanced Repaid to Agent" in first_page_text:
                    df,output_name = pdf_editor.polish_falcons()
                elif "Writing Agent Policy No. Description Code Dur. Date Due Mths. Premium Rate Commission" in first_page_text:
                    df,output_name = pdf_editor.kskj_Life()
                elif "PFA Financial Life" in first_page_text:
                    df,output_name = pdf_editor.polish_falcons2()
                elif "DATE PAYEE AGENT PAYEE MEMBER MEMBER AGENT PRODUCT TRANSACTION AMOUNT PAYOUT CREDIT DEBIT" in first_page_text:
                    df,output_name = pdf_editor.USAHealth()
                elif "Divisions of Health Care Service Corporation, a Mutual Legal Reserve Company, an Independent Licensee of the Blue Cross and Blue Shield Association" in first_page_text:
                    df,output_name = pdf_editor.bcbs()
                elif "NUMBER INSURED MD AGE PRD LV PAID DATE YR RATE PREMIUM COMMISSION PAID REMAINING NET" in first_page_text:
                    df,output_name = pdf_editor.family_benefit_life()
                elif "Broker # Referral(s) Commission Adjustment(s)" in first_page_text:
                    df,output_name = pdf_editor.river_health()
                elif "Month Subscribers Schedule Premium Paid Commission" in first_page_text:
                    df,output_name = pdf_editor.kaiser_permanente()
                elif "Delta Dental of Colorado" in first_page_text:
                    df,output_name = pdf_editor.delta_dental_colorado()
                elif "Group No. Group Name Billing Period Adj. Period Invoice Total Stoploss Total Agent Rate Calculation" in first_page_text:
                    df,output_name = pdf_editor.allied()
                elif "Delta Dental of Virginia" in first_page_text:
                    df,output_name = pdf_editor.delta_dental_virginia()
                elif "Peek Performance Insurance" in first_page_text:
                    df,output_name = pdf_editor.peek_performance()
                elif "LifeShield National Insurance" in first_page_text:
                    df,output_name = pdf_editor.life_shield()
                elif "Liberty Bankers Life Insurance Company" in first_page_text:
                    df,output_name = pdf_editor.libery_bankers()
                elif "ID Name Effective Date Coverage Period LOB Plan / Adjustment Description Fst / Ren Premium Month Rate Commission Due" in first_page_text:
                    df,output_name = pdf_editor.baylor_scott()
                elif "Liberty Bankers Insurance Group" in first_page_text:
                    df,output_name = pdf_editor.liberty_bankers_life()
                elif "CHECK DATE LAST PAY DATE AGENCY # AGENCY (group) CLIENT # CLIENT INVOICE # COMM RATE PREMIUM NET AGENT C MO KM T M SEG M RAS TE C CO ON UNTR TACT" in first_page_text:
                    df,output_name = pdf_editor.sentara_aca()
                elif "STEPHENS-MATTHEWS MARKETING" in first_page_text:
                    df,output_name = pdf_editor.stevens_matthews()
                elif "United American Insurance Company" in first_page_text or "Globe Life Insurance Company of New York" in first_page_text:
                    df,output_name = pdf_editor.united_american()
                elif "Agent NPN Agent Name Member ID Member HICN Member First Member Last Member State Effective Date" in first_page_text:
                    df,output_name = pdf_editor.jefferson_health()
                elif "Broker Contract ID Name Prem Premium Premium Count Rate Commission Type Commission" in first_page_text:
                    df,output_name = pdf_editor.health_first_fl()
                elif f"Number Month Received Rate % Earned" in first_page_text:
                    df,output_name = pdf_editor.inshore()
                elif "NIPPON LIFE BENEFITS" in first_page_text:
                    df,output_name = pdf_editor.nippon_life()
                # Add other conditions as needed...
                # if df is None:
                #     print(f"Df is none {output_name}")
                #     return render(request, 'upload.html', {'form': form, 'message': True})

                # Save the DataFrame to Excel and store the file path
                filename = pdf_editor.save_to_excel(df, output_name)
                # print(f"Type of filename is {type(filename)}")
                if filename is not None:
                    filenames.append(filename)

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
                try:
                    single_file_name = os.path.basename(filenames[0])
                except IndexError:
                    return render(request, 'upload.html', {'form': form, 'message': True})
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
