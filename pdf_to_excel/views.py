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
from .helpers2 import PDFS

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
                extended_pdf_editor = PDFS(pdf_file)
                
                
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
                passwords = [None,"2646","WG500","LBL22728","7964","mphc4apps"]
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
                    df, output_name = extended_pdf_editor.royal_neighbors()
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
                    df, output_name = pdf_editor.bcbs_la_commisions("BCBS LA")
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
                elif "Month Subscribers Schedule Premium Paid Commission" in first_page_text or "Received Date Received % Dues Paid Rate Total Commission" in first_page_text:
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
                elif "POLICY INSURED'S NAME PLAN CODE PAID TO PREMIUM PERCENT EARNED AMT TO PAY FICA APPL TO ADV BALANCE" in first_page_text:
                    column_ranges = [
                        (24,65),
                        (67,150),
                        (153,192),
                        (193,228),
                        (246,273),
                        (275,315),
                        (320,355),
                        (360,400),
                        (401,430),
                        (443,473),
                    ]
                    df,output_name = pdf_editor.cigna_ms_lisa(column_ranges)
                elif "POLICY INSURED'S NAME PLAN CODE PAID TO PREMIUM PERCENT EARNED AMT TO PAY FICA APPL" in first_page_text:
                    column_ranges = [
                        (24,76),
                        (76,190),
                        (196,246),
                        (246,296),
                        (324,354),
                        (359,402),
                        (425,455),
                        (482,515),
                        (520,556),
                        (556,580),
                    ]
                    df,output_name = pdf_editor.cigna_ms_lisa(column_ranges)
                elif "Producer Producer Name E&O End Date Active Producer BC Commission HMO Commission Life Commission Total Commission" in first_page_text:
                    df,output_name = pdf_editor.bcbs_lousiana("BCBS LA")  
                elif "Producer Producer Name E&O End Date Active Producer PWMS Commission PWAR Commission Total Commission" in first_page_text:
                    df,output_name = pdf_editor.bcbs_lousiana("Primewell MAPD")
                elif "Member Name Member ID Status Selling Agent Effective Date Term Date Amount" in first_page_text:
                    df,output_name = pdf_editor.martins_point()
                elif "www.careington.com" in first_page_text:
                    df,output_name = pdf_editor.carington()
                elif "Northeast Delta Dental" in first_page_text:
                    df,output_name = pdf_editor.delta_dental_northeast()
                elif "WRITING SOURCE TRAN EFFECT PAID PAID NAME POLICY PLAN CURR PREMIUM FEE % COMMISSION" in first_page_text:
                    column_ranges = [
                        (0,54),
                        (54,74),
                        (74,105),
                        (105,137),
                        (137,159),
                        (160,178),
                        (178,221),
                        (221,265),
                        (265,300),
                        (300,315),
                        (320,355),
                        (373,390),
                        (408,443),
                        (450,496),
                        (496,532),
                        (546,581),
                    ]
                    df,output_name = pdf_editor.general_agent_center(column_ranges)
                elif "SUBSCRIBER NAME ALTERNATE ID CONTRACT TYPE EFF DATE C DUE DATE PREMIUM PERCENT AMOUNT REA AGENT" in first_page_text:
                    column_ranges = [
                        (15,41),
                        (47,64),
                        (96,185),
                    ]
        
                    column_ranges_two = [
                        (15,97),
                        (97,151),
                        (155,190),
                        (205,237),
                        (238,246),
                        (246,277),
                        (284,317),
                        (320,355),
                        (370,394),
                        (395,406),
                        (407,433),
                    ]
                    df,output_name = pdf_editor.bcbs_sc(column_ranges,column_ranges_two)
                elif "Policy Policyholder Transaction Premium Premium PremiumPaid Commission Commission" in first_page_text:
                    df,output_name = pdf_editor.cigna_global()
                elif "Customer ID Customer Name Prod Coverage First Eff Date Age at Disability Member Bill Eff Paid Thru Prem Member Comm CommComm" in first_page_text:
                    df, output_name = extended_pdf_editor.BCBS_Nebraska()
                elif "Kaiser Foundation Health Plan of the Northwest" in first_page_text:
                    df,output_name = extended_pdf_editor.kaiser_permanente_northwest()
                elif "Case Number External ID Description Coverage Lives Paid Compensation Writing" in first_page_text:
                    column_ranges = [
                        (32,86),
                        (105,154),
                        (172,246),
                        (258,298),
                        (341,354),
                        (389,425),
                        (471,502),
                        (504,539),
                    ]
                    df,output_name = extended_pdf_editor.pivot_health(column_ranges)
                elif "CERT# HOLDER EFFECT PAID PAID PREMIUM COMM SRC 1ST 1ST REN REF" in first_page_text:
                    df,output_name = extended_pdf_editor.sons_of_norway()
                elif "Group ID Group Name Total Premium Premium Paid Total Paid Prem Med/Den Comm Retro ?Commission" in first_page_text:
                    df,output_name = extended_pdf_editor.providence2()
                elif "GROUP/SUBSCR NAME SUBSCR ID ALTERNATE ID TYPE EFF DATE C DUE DATE TYPE RACTS PAID PERCENT AMOUNT REA AGENT" in first_page_text:
                    column_ranges = [
                        (15,41),
                        (47,64),
                        (96,185),
                    ]
        
                    column_ranges_two = [
                        (15,97),
                        (97,172),
                        (177,223),
                        (223,255),
                        (256,263),
                        (263,295),
                        (295,315),
                        (324,335),
                        (338,371),
                        (374,408),
                        (425,448),
                        (448,461),
                        (461,487),
                    ]
                    df,output_name = pdf_editor.bcbs_sc(column_ranges,column_ranges_two)
                    
                else:
                    df,output_name = extended_pdf_editor.bcbs_kc()
                # elif "Kaiser Foundation Health Plan of Georgia" in first_page_text:
                #     df,output_name = pdf_editor.kaiser_georgia()
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
