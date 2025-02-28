import os
import re
import pandas as pd
from io import BytesIO
import pdfplumber
import time
import fitz

class PDFEditor:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.pdf_output_name = pdf_file.name.rstrip(".pdf")
        
    def rotate_pdf(self, rotation=0):
        self.pdf_file.seek(0)
        pdf_bytes = self.pdf_file.read()  # Read the uploaded file as bytes
        doc = fitz.open("pdf", pdf_bytes)  # Open from bytes
        
        for page in doc:
            if page.rotation != 0:
                page.set_rotation(rotation)  # Rotate each page

        # Save the rotated PDF into a byte stream and return it
        output_stream = BytesIO()
        doc.save(output_stream)
        doc.close()
        output_stream.seek(0)  # Move cursor to the beginning
        return output_stream
    
    def is_valid_pdf(self):
        """Check if the uploaded file is a valid PDF."""
        # print(self.pdf_file.name.lower().endswith(".pdf"))
        return self.pdf_file.name.lower().endswith(".pdf")

    
    def extract_page_text(self, pdf_path, page_num):
        with pdfplumber.open(pdf_path) as pdf:
            page_text = pdf.pages[page_num].extract_text()
            # gc.collect()  # Call garbage collection to free up memory
            return page_text or ""

    def extract_text(self,start=0, pages=None):
        """Extract text from specified pages or all pages in the PDF sequentially."""
        
        with pdfplumber.open(self.pdf_file) as pdf:
            if pages is None:
                pages = range(start,len(pdf.pages))
            else:
                pages = range(pages)
            text_list = []
            for page_num in pages:
                page_text = self.extract_page_text(self.pdf_file, page_num)
                text_list.append(page_text)
                # gc.collect()  # Optional: Call garbage collection after each page extraction

        # Join extracted text from all pages.
        return "\n".join(text_list)
    
    def extract_text_delta(self,text,start=0, pages=None):
        """Extract text from specified pages or all pages in the PDF sequentially."""
        
        with pdfplumber.open(text) as pdf:
            if pages is None:
                pages = range(start,len(pdf.pages))
            else:
                pages = range(pages)
            text_list = []
            for page_num in pages:
                page_text = self.extract_page_text(text, page_num)
                text_list.append(page_text)
                # gc.collect()  # Optional: Call garbage collection after each page extraction

        # Join extracted text from all pages.
        return "\n".join(text_list)
    
    #Used for proccess pdf type1
    def extract_large_table(self,start_page, agent_table_number):
        full_table = []
        table_found = False
        with pdfplumber.open(self.pdf_file) as pdf:

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
    
    def extract_tables_from_pdf(self):
        all_tables = []
        with pdfplumber.open(self.pdf_file) as pdf:
            for page in pdf.pages:
                # Extract tables from the page
                tables = page.extract_tables()
                all_tables.extend(tables)
        return all_tables
    
    def extract_tables(self):
        tables = []  # List to store extracted tables
        with pdfplumber.open(self.pdf_file) as pdf:
            for page in pdf.pages:  # Iterate through each page
                # Extract tables from the current page
                page_tables = page.extract_tables()
                for table in page_tables:
                    # Convert the table into a DataFrame and append to the list
                    df = pd.DataFrame(table[1:], columns=table[0])  # Skip header
                    tables.append(df)
        return tables
    
    # Uses PyMuPDF
    def extract_text_from_range(self,start_page,end_page=None):
        # start_time = time.time()
        extracted_text = ""

        self.pdf_file.seek(0) 
        file_content = self.pdf_file.read()
        
        with fitz.open(stream=file_content,filetype="pdf") as pdf:
            total_pages = len(pdf)
            # Validate the range
            if start_page < 0:
                print(f"Invalid page range. The PDF has {total_pages} pages.")
                return ""
            if end_page is None:
                end_page = total_pages
            
            for page_number in range(start_page, end_page):
                page = pdf[page_number]  # Access each page in the range
                extracted_text += page.get_text() + "\n"  # Append text from each page, add newline for separation

        # end_time = time.time()


        return extracted_text
    
    # ------------------------------------------- KANSAS COMMISSIONS LIFE PDF EXTRA FUNCTIONS --------------------------------------------------
    def cidToChar(self, cidx):
        """Convert (cid:XX) to a character by adding 29 to the numeric part."""
        return chr(int(re.findall(r'\(cid\:(\d+)\)', cidx)[0]) + 29)

    def processText(self, text):
        """Process the input text, replacing (cid:XX) patterns and returning clean lines."""
        data = []
        for line in text.split('\n'):
            # Skip empty lines or lines with only '(cid:3)'
            if line and line != '(cid:3)':
                # Replace all (cid:XX) patterns in the line with corresponding characters
                for cid in re.findall(r'\(cid\:\d+\)', line):
                    line = line.replace(cid, self.cidToChar(cid))
                data.append(line.strip("'"))
        return data
    # ------------------------------------------- KANSAS COMMISSIONS LIFE PDF EXTRA FUNCTIONS --------------------------------------------------
    def process_pdf_type1(self):
        """Process the first type of PDF (with 'Run Date' and 'Agents')."""
        output_name = self.pdf_output_name

        # Extract text from page 4
        text = self.extract_text()

        # Extract 'Run Date' and 'Agents' information
        agents = re.findall(r"Subtotals for Agent (\w+)\s+([A-Z,.\s]+)", text)
        run_date = re.findall(r"Run Date:\s(\d{2}/\d{2}/\d{4})", text)

        if not agents:
            return None, None  # Handle invalid case

        new_data, extra_cols = [], []

        for i, agent in enumerate(agents):
            large_table = self.extract_large_table(4, i)

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

        # Add custom column
        for i in range(1):
            new_data[i]["Converted from .pdf by"] = ""

        # Convert to pandas DataFrame
        df = pd.DataFrame(new_data)
        
        return df, output_name


    def forester_financial(self):
        """Process another type of PDF (define structure here)."""
        output_name = self.pdf_output_name
        
        # Extract text from the necessary pages and apply regex as needed
        text = self.extract_text()  # Assume relevant data is on page 1

        data_table = []
        pattern = r"(\d+)\s([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z-]+){0,2})\s([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z-]+){0,2})\s(\d+)\s(\w+)\s(\d+\/\d+\/\d+)\s(Premium Payment)\s(\$(?:\d+,)?\d+.\d+)(\s\$?\w+)?(\s\$?\w+)?\s(\d+.\d+)\s(\(?\$(?:\d+,)?\d+.\d+\)?)"
        text_parts = re.search(r"Renewal LIFE\nCommission(.*?)Renewal Override LIFE", text, re.DOTALL)
        if text_parts:
            match = re.findall(pattern,text_parts.group(1),re.DOTALL)
            for renewal_commission_results in match:
                data_table.append({
                    "Carrier": "Foresters",
                    "type" : "Renewal Commissions LIFE",
                    "Writing Producer #":renewal_commission_results[0],
                    "Writing Producer Name":renewal_commission_results[1],
                    "Insured Name":renewal_commission_results[2],
                    "Policy #": renewal_commission_results[3],
                    "Plan Code": renewal_commission_results[4],
                    "Transaction Date": renewal_commission_results[5],
                    "Transaction Type": f"{renewal_commission_results[6]} US",
                    "Compensation Basis Amount": renewal_commission_results[7],
                    "% Repl. Factor":renewal_commission_results[8],
                    "% Share": renewal_commission_results[9],
                    "Comp Rate %": renewal_commission_results[10],
                    "Amount Due": renewal_commission_results[11]})
                   
        renewal_override = re.search(r"Renewal Override LIFE(.*?)Current Balance:",text,re.DOTALL)

        match = re.findall(pattern,renewal_override.group(1),re.DOTALL)
        for renewal_override_results in match:
            data_table.append({
                "Carrier": "Foresters",
                "type" : "Renewal Override LIFE",
                "Writing Producer #":renewal_override_results[0],
                "Writing Producer Name":renewal_override_results[1],
                "Insured Name":renewal_override_results[2],
                "Policy #": renewal_override_results[3],
                "Plan Code": renewal_override_results[4],
                "Transaction Date": renewal_override_results[5],
                "Transaction Type": f"{renewal_override_results[6]} US",
                "Compensation Basis Amount": renewal_override_results[7],
                "% Repl. Factor":renewal_override_results[8],
                "% Share": renewal_override_results[9],
                "Comp Rate %": renewal_override_results[10],
                "Amount Due": renewal_override_results[11]})
        

        # Add custom column
        for i in range(1):
            data_table[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data_table)

        return df, output_name
    # ------------------------------------------- ASSURITY PDF EXTRA FUNCTIONS --------------------------------------------------
    # Function to extract name-amount pairs from sections
    """
    This function is used to extract data from Assurity PDF's
    """
    def first_columns(self,text,title,endsearch,lst,date):
        pattern = r"^([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z]+){0,2})\s(\$(?:\d+,)?\d+.\d+)\s[a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z]+){0,2}\s\$(?:\d+,)?\d+.\d+$"
        
        columns = re.search(rf"{title} (.*?){endsearch}",text,re.DOTALL)
        match = re.findall(pattern,columns.group(1),re.MULTILINE)
        for pfyc in match:
            lst.append({
                "Carrier": "Assurity",
                "Statement Date": date,     
                "Type": title,
                "Agent Name": pfyc[0],
                "Amount": pfyc[1]
            })
            
    def second_columns(self,text,title,endsearch,lst,date):
        pattern = r"([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z]+){0,2})\s(\$(?:\d+,)?\d+.\d+)$"
        
        columns = re.search(rf"{title}(.*?){endsearch}",text,re.DOTALL)
        match = re.findall(pattern,columns.group(1),re.MULTILINE)
        for pfyc in match:
            lst.append({
                "Carrier": "Assurity",
                "Statement Date": date,     
                "Type": title,
                "Agent Name": pfyc[0],
                "Amount": pfyc[1]
            })
  
    # ------------------------------------------- ASSURITY PDF EXTRA FUNCTIONS --------------------------------------------------
    
    def assurity_commission(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        data = []
        date = re.search(r"as of (\d{2}/\d{2}/\d{4})", text, re.DOTALL)
        date = date.group(1) if date else "N/A"

        self.first_columns(text,"PAID FIRST YEAR COMMISSIONS","Total PFYC",data,date)
        self.first_columns(text,"RENEWAL COMMISSIONS","Total Renewal",data,date)
        self.second_columns(text,"YTD PAID FIRST YEAR COMMISSIONS","Total YTD PFYC",data,date)
        self.second_columns(text," YTD RENEWAL COMMISSIONS","Total YTD Renewal",data,date)

        for i in range(1):
            data[i]["Converted from .pdf by"] = "" 
        
        df = pd.DataFrame(data)
        return df, output_name

    def kansas_city_life(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        decoded_text = self.processText(text)
        
        info = []    
        #Matching the an specific info
        for line in decoded_text:
            match = re.search(r"\s*AGENCY:(.+)?PAY PERIOD:",line)
            if match:
                agency = str(match.group(1)).strip()
                break
        for line in decoded_text:
            match = re.search(r"(\d+\/\d+)\s+(\d+)\s+(\d+)\s+([a-zA-Z]+[\s-]?[a-zA-Z]+)\s+((?:\w+\s+)+)(\d+\/\d+)\s+(\d+)\s+(-?\$(?:\d+,)?\d+.\d{0,2})\s+(\d+.\d{0,2})\s+(-?\$(?:\d+,)?\d+.\d{0,2})\s+(\w+\s+)+(-?\$(?:\d+,)?\d+.\d{0,2})\s(\s|.+)\s(\s|.+)\s(\w+)",line)
            if match:
                info.append(list(match.groups()))
        #Removing leading whitespace from matched info
        info = [list(map(str.strip,i)) for i in info]
        #Removing spaces from Plan
        for i in info:
            i[4] = re.sub(r"\s","",i[4])
        data = []
        for row in info:
            data.append({
                "Carrier": "Kansas City Life",
                "Agency": agency,
                "CURR": row[0],
                "JNT": row[1],
                "POLICY #": row[2],
                "NAME": row[3],
                "PLAN": row[4],
                "ANNIV": row[5],
                "YR": row[6],
                "PREM": row[7],
                "RATE": row[8],
                "EARNINGS": row[9],
                "EXPLANATIONS": row[10],
                "PAY": row[11],
                "DISTR AMT": row[12],
                "DISTR TO": row[13],
                "DISTR FR": row[14]
            })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df, output_name
    
    def sentinel(self):
        output_name = self.pdf_output_name
        text = self.extract_text()

        agency = re.search(r"Sentinel Security Life Insurance Company Commission Statement([^,].*?)LLC",text,re.DOTALL)
        agency = agency.group(1).replace(",","").strip("\n")
        
        date = re.search(r"Commission Period: \d+ - (\d+)",text,re.DOTALL)
        date = date.group(1)

        pattern = r"(\w+)\s+((?:[a-zA-Z\s]+)?[a-zA-Z]+,\s[a-zA-Z]+(?:\s[a-zA-Z]+)?)\s+(\d+)\s+([a-zA-Z]+\s[a-zA-Z.]+(?:\s[a-zA-Z]{3,})?)\s([a-zA-Z]{2})\s(\d+)\s(\d+\/\d+)\s(\d+)\s(\d+\/\d+)\s(\d+)\s(\(?\$(?:\d+,)?\d+.\d+\)?)\s(\d+.\d+\s%)\s(\(?\$(?:\d+,)?\d+.\d+\)?)\s(\(?\$(?:\d+,)?\d+.\d+\)?)"

        types = re.search(rf'Agent # Writing Agent Name Policy # Name St Plan Code Mo/Yr Paid Date Dur. Premium Rate Comm Advance\n(.*?){pattern}',text,re.DOTALL)
        if types:
            line,tpe = types.group(1).split("\n",maxsplit=1)
            
        match = re.findall(pattern,text,re.MULTILINE)
        data = []
        for row in match:
            data.append({
                "Carrier": "Sentinel Security",
                "Agency": agency,
                "Line": line,
                "Type": tpe,
                "Data": date,
                "Writing Agent #": row[0],
                "Writing Agent Name": row[1],
                "Policy #": row[2],
                "Name": row[3],
                "App St": row[4],
                "Plan Code": row[5],
                "Eff Mo/Yr": row[6],
                "Mths Paid": row[7],
                "Paid to Date": row[8],
                "Policy Dur.": row[9],
                "Comm Premium": row[10],
                "Comm Rate": row[11],
                "Payable Comm": row[12],
                "Applied to Advance": row[13]
            })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name

    def bcbs_la_commisions(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        date = re.search(r"Activity Ending Date: (\d+\/\d+\/\d+)",text,re.DOTALL)
        date = date.group(1)
        #Info patterns
        pattern = r"(\d+)\s([a-zA-Z]+(?:\s[a-zA-Z]+){0,3})\s(\w+)\s(\d+(?:\s[a-zA-Z]+){0,4})\s(\w+)\s(.*?)(\d+\/\d+\/\d+)\s(.*?)(\d+\/\d+\/\d+)\s(\d+\/\d+)\s(\d+)\s(\w+)\s(\$ (?:\d+,)?\d+.\d+)"
        tpe = re.findall(r"Writing Producer (\d+)\s(\w+\s\w+)\n(\w+(?:\s\w+)?)\n(.*?)Total for",text,re.DOTALL)

        data = []
        for t in tpe:
     
            x = re.findall(rf"{pattern}",t[3])

            for y in x:
                data.append({
                    "Carrier": "BCBS LA",
                    "Statement Type": "Statements of Commissions",
                    "Date": date,
                    "Type": t[2],
                    "Writing Producer": t[1],
                    "Writing ID": t[0],
                    "Member ID": y[0],
                    "Name": y[1],
                    "Company": y[2],
                    "Product": y[3],
                    "HICN": y[4],
                    "Override": y[5],
                    "Effective Date": y[6],
                    "Term Date": y[7],
                    "Signed Date": y[8],
                    "Period": y[9],
                    "Cycle Year": y[10],
                    "Retro": y[11],
                    "Amount": y[12], 
                })

        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name
        
    def bcbs_la_compensation(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        data = []
        date = re.search(r"Activity Ending Date: (\d+\/\d+\/\d+)",text,re.DOTALL)
        date = date.group(1)
        
        pay_entity = re.search(r"Pay Entity: ([a-zA-Z-]+(?:\s[a-zA-Z-]+){0,3}) Activity Ending Date",text,re.DOTALL)
        pay_entity = pay_entity.group(1)
        
        #Info patterns
        pattern = r"(\d+)\s([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z])?)\s(\w+)\s(\w+)\s(\d+\/\d+\/\d+)\s(\d+\/\d+\/\d+)\s(.*?)\s(\(?\$\s(?:\d+,)?\d+.\d+\)?)\s(\d+)\s(\d+.\d+%)\s(\(?\$\s(?:\d+,)?\d+.\d+\)?)"
        tpe = re.findall(r"Producer (\d+) NPN (\d+) Producer Name ([a-zA-Z-]+(?:\s[a-zA-Z-]+){0,3}) Total(.*?)Total Individual Payment",text,re.DOTALL)
        
        for t in tpe:
            id,npn,name = t[0],t[1],t[2]
            
            x = re.findall(rf"{pattern}",t[3])
 
            for y in x:
                data.append({
                    "Carrier": "BCBS LA",
                    "Statement Type": "Compensation Statement",
                    "Pay Entity": pay_entity,
                    "Statement Date": date,
                    "Producer Name": name,
                    "Producer ID": id,
                    "Producer NPN": npn,
                    "Current Contract ID":y[0],
                    "Subscriber Name": y[1],
                    "Company": y[2],
                    "MOP": y[3],
                    "OED": y[4],
                    "Due Date": y[5],
                    "Product Name": y[6],
                    "Premium Collected": y[7],
                    "Elapsed Months": y[8],
                    "Comm. %": y[9],
                    "Commission Due": y[10],   
                })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name
    
    def essence_file(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        
        data = []
        
        agency = re.search(r"(.*?)Statement Of Commissions",text)
        
        writing_agent_pattern =r'Writing Agent \d+ ((?:[a-zA-Z]*\s)?[a-zA-Z]+, (?:[a-zA-Z]*\s)?[a-zA-Z]*)\n'
        writing_agent = re.search(writing_agent_pattern, text)
        
        statement_date = re.search(r"Commission Period: (\d{2}\/\d{2} - \d{2}\/\d{2})",text)
        
        pattern = r'(\d+) (\d+) ((?:[a-zA-Z]*\s)?[a-zA-Z]+, (?:[a-zA-Z]*\s)?[a-zA-Z]*) (\w+) (\w{2}) (\d{2}\/\d{2}\/\d{4}) ((?:\d+\/\d+\s)?)((?:\w+\s)?)(\d{2}\/\d{4}) (\w+) (\w+) (\$\d+.\d+)'

        filtered = re.findall(r"Writing Agent(.*?)Total for (?:Annual Renewals|New Enrollments)",text,re.DOTALL)
        info = re.findall(pattern, str(filtered), re.DOTALL)
        statement_type = re.search(r"(Annual Renewals|New Enrollments)",str(filtered))

        for row in info:
            data.append({
                "Carrier": "Essence Healthcare",
                "Statement Type": "Statement of Commissions",
                "Agency": agency.group(1),
                "Statement Date": statement_date.group(1),
                "Type": statement_type.group(1),
                "Writing Agent": writing_agent.group(1),
                "Writing ID": row[1],
                "Member ID": row[0],
                "Name": row[2],
                "Product": row[3],
                "Policy State": row[4],
                "Effective Date": row[5],
                "Term Date": row[6],
                "Term Code": row[7],
                "Period": row[8],
                "CMS Payment Type": row[9],
                "Retro": row[10],
                "Commission Ammount": row[11],
            })
        
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name
        
    def blueshield_of_california(self):
        carrier = "Blue Shield CA"
        output_name = self.pdf_output_name
        text = self.extract_text(0,1)
        tablas = self.extract_tables_from_pdf()
        data = []
        
        date_pattern = r'Statement Date: ([0-9\/]+)'
        date = re.search(date_pattern,text)

        for tabla in tablas:
            for row in tabla:
                if row[0] != None:
                    if "Commission" in row[0]:
                        commission_type = row[0]
                        print(f"La commission es: {row}")
                    if "Writing Producer" in row[0]:
                        producer_info = row[1:]
                        producer = [info for info in producer_info if info != None]
                        if len(producer) == 3:
                            producer[2] = producer[2].replace("NPN","")
                            producer_id,producer_name,producer_npn = producer
                        else:
                            producer_id,producer_name = producer
                            producer_npn = ""
                        
                            
                    if re.match(r'X\d+',row[0]):
                        row = [col for col in row if col != None]
                        data.append({
                            "Carrier" : carrier,
                            "Statement Date" : date.group(1),
                            "Writing Producer ID" : producer_id,
                            "Producer NPN": producer_npn,
                            "Producer Name": producer_name,
                            "Group Number": row[0],
                            "Subscriber ID": row[1],
                            "Customer Name": row[2],
                            "Product": row[3],
                            "Effective Date": row[4],
                            "Term Date": "",
                            "Period": row[5],
                            "Gross Premium": row[6],
                            "Base Premium": row[7],
                            "Commission Rate": row[8],
                            "Cycle Year": "",
                            "Commission Paid": row[9],
                            "Commission Type": commission_type,
                        })
                    elif re.match(r'\d+',row[0]):
                        row = [col for col in row if col != None]
                        data.append({
                            "Carrier" : carrier,
                            "Statement Date" : date.group(1),
                            "Writing Producer ID" : producer_id,
                            "Producer NPN": producer_npn,
                            "Producer Name": producer_name,
                            "Group Number": "",
                            "Subscriber ID": row[0],
                            "Customer Name": row[1],
                            "Product": row[2],
                            "Effective Date": row[3],
                            "Term Date": row[4],
                            "Period": row[5],
                            "Gross Premium": "",
                            "Base Premium": "",
                            "Commission Rate": "",
                            "Cycle Year": row[6],
                            "Commission Paid": row[7],
                            "Commission Type": commission_type,
                        })
                        
            
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name
    
    def providence(self):
        # start_time = time.time()
        carrier = "Providence Med Adv"
        output_name =  self.pdf_output_name
        text = self.extract_text_from_range(0)

        # agency = re.match(r'^([a-zA-Z ]+)\n\d+',text,re.DOTALL)
        agency = re.search(r'\w+\n([a-zA-Z ]+?)\nNPN',text,re.DOTALL|re.MULTILINE)
        commission_period = re.search(r'Commission Period: ([0-9- ]+)',text)

        data = []
        
        producer_pattern = r'Producer(.*?)Writing Producer\s+(\w+)\s+([a-zA-Z-, ]+)'
        enrollments_renewals_pattern = r'^(New Enrollment|Renewals)(.*?)(?:Total for New Enrollments|Total for Renewals)'
        clients_info = r'^([a-zA-Z-,0-9]+)\n(\w+)\s?([a-zA-Z \-,0-9]+)\n(-?\$[0-9 .]+)\n([a-zA-Z]{1})\n(\d+\/\d+\/\d+)\n(\d+\/\d+\/\d+\n)?(|\d+\/\d+\/\d+)\n(\d+)?\n?([0-9]{2}\/\d+)\n([a-zA-Z]+\s[a-zA-Z]+|[a-zA-Z]+)\s([a-zA-Z]+)$'
        client_info_pattern = r'^(x[a-zA-Z0-9]+)?\n?(\w+)\s?([a-zA-Z \-,0-9]+)\n(-?\$[0-9 .]+)\n([a-zA-Z]{1})\n([a-zA-Z -]+)\n(\d+\/\d+\/\d+)\n(\d+\/\d+\/\d+)?\n?(\d+\/\d+\/\d+)?\n?(\d+\n)?(\d+\/\d+)\n([a-zA-Z -]+)\n'
        
        # Creates a list of all the producers
        filtered = re.findall(producer_pattern,text,re.DOTALL|re.MULTILINE)
        # Loops through the list of producers
        for f in filtered:
            # Creates list of  New Enrollment to Total New Enrollments or Renewals for each producer
            new = re.findall(enrollments_renewals_pattern,f[0],re.DOTALL|re.MULTILINE)
            # Loops through the list of New Enrollment to Total New Enrollments or Renewals
            for n in new:
                # Gets the transaction type
                # Creates a list of members information for each producer, loops through it and stores information in a list
    
                info = re.findall(clients_info,n[1],re.DOTALL|re.MULTILINE)
        
                client_info = re.findall(client_info_pattern,n[1],re.DOTALL|re.MULTILINE)
                
                if client_info:
                    
                    for i in client_info:
                        transaction_type = n[0]
                        data.append({
                            "Carrier": carrier,
                            "Agency": agency.group(1),
                            "Document Type": "Statement Of Commissions",
                            "Commission Period": commission_period.group(1),
                            "Producer Name": f[2],
                            "Producer ID": f[1],
                            "Transaction Type": transaction_type,
                            "Member ID": i[1],
                            "Name": i[2],
                            "Line of Business": i[11],
                            "Product": i[5],
                            "MBI": i[0],
                            "Effective Date": i[6],
                            "Term Date": i[7],
                            "Signed Date": i[8],
                            "Period": i[10],
                            "Cycle Year": i[9],
                            "Retro": i[4],
                            "Commission Ammount": i[3],        
                        })
                    
                                  
                for i in info:
                    transaction_type = n[0]
                    data.append({
                        "Carrier": carrier,
                        "Agency": agency.group(1),
                        "Document Type": "Statement Of Commissions",
                        "Commission Period": commission_period.group(1),
                        "Producer Name": f[2],
                        "Producer ID": f[1],
                        "Transaction Type": transaction_type,
                        "Member ID": i[1],
                        "Name": i[2],
                        "Line of Business": i[10],
                        "Product": i[11],
                        "MBI": i[0],
                        "Effective Date": i[5],
                        "Term Date": i[6],
                        "Signed Date": i[7],
                        "Period": i[9],
                        "Cycle Year": i[8],
                        "Retro": i[4],
                        "Commission Ammount": i[3],        
                    })
        
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""

        df = pd.DataFrame(data)
        # end_time = time.time()
        # print(f"Time taken to run Providence method : {end_time - start_time:.2f} seconds")
        return df, output_name
    
    def cincinatti(self):
        output_name = self.pdf_output_name
        agency_text = self.extract_text(pages=1)
        agency_pattern = r'([a-zA-Z0-9 ,]+)\nFrom (\d+\/\d+\/\d+ to \d+\/\d+\/\d+)'
        match = re.search(agency_pattern,agency_text)
        agency, payment_period = match.groups()

        info_pattern = r'(\d+)\s+([a-zA-Z -.]+)\s+(\d+)\s+(|\w+)\s?(\d+\/\d+\/\d+)\s([0-9.,]+)\s([0-9.]+%)?\s?(\w+)?\s(\d+\/\d+\/\d+)?\s*(-?[0-9]{,5}\.?\d{,2}?)?\s*([a-zA-Z]+ [a-zA-Z]+\s|[a-zA-Z\n]+\s)?(\d{8}\s)?([-0-9.,]+)$'
        
        text = self.extract_tables_from_pdf()
        carrier = "Cincinnati Equitable"
        data = []
        x = ""
        for table in text:
            reason = table[1][7]
            if reason == None:
                reason = table[1][6] if table[1][6] is not None else table[1][10]
            info = table[6:-2]
            for row in info:
                for i in row:
                    if i is None:
                        i = ""
                    x = x + " " + i
                x = x + "\n"
                results = re.findall(info_pattern,x)
                for i in results:
                    data.append({
                        "Carrier": carrier,
                        "Agency": agency,
                        "Payment Period": payment_period,
                        "Policy Number": i[0],
                        "Insured Name": i[1],
                        "Issue Age": i[2],
                        "Policy Type": i[3],
                        "Issue Date": i[4],
                        "Base": i[5],
                        "Rate": i[6],
                        reason + " " + "Reason": i[7],
                        reason + " " + "Date": i[8],
                        "Chargeback Amount": i[9],
                        "Producer": i[10], 
                        "Producer Code": i[11],
                        "Commission Paid": i[12],
                    })

        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name
        
    def polish_falcons(self):
        output_name = self.pdf_output_name
        data = []
        carrier = "Polish Falcons of America"
        text = self.extract_text(pages=1)

        agency_pattern = r'^\d+\/\d+\/\d+.*\n\w+\s(.*)'
        match = re.search(agency_pattern,text,re.MULTILINE)
        agency = match.group(1)
        
        text = self.extract_text_from_range(0)
        tables_pattern = r'(.*?)\s*(\w+ - [a-zA-Z ,.]+)'
        tables = re.findall(tables_pattern,text,re.DOTALL)
        
        info_pattern = r'(\w+)\s*(\d+ \(\w+\))\n(\w)\s*([0-9.,]+)\s*([0-9.]+)\s*(\d+)\s*([0-9.]+)\s*([0-9.]+)\s*(-?[0-9.]+)\s*(\d+)\s*(\w+)\s*([a-zA-Z-. ]+)\s*(-?[0-9.]+)\s*([0-9.]+)\s*(-?[0-9.]+)\s*(\d+\/\d+\/\d+)\s*(\d+\/\d+\/\d+)'

        for table in tables:
            info = re.findall(info_pattern,table[0])
            for i in info:
                data.append({
                    "Carrier": carrier,
                    "Agency": agency,
                    "Policy": i[0],
                    "Insured/Anuitant": i[11],
                    "Plan": i[1],
                    "Issue Date": i[16],
                    "Mode": i[2],
                    "Value": i[3],
                    "Base Premium": i[4],
                    "Age": i[5],
                    "Year": i[9],
                    "Selling Agent": i[10],
                    "Agent Share": "",
                    "Payment Date": i[15],
                    "Payment": i[8],
                    "Percent": i[7],
                    "Earned": i[12],
                    "Advanced": i[13],
                    "Repaid": i[14],
                    "Paid to Agent": i[6]
                })
            data.append({
                "Extra": table[1]
                })

        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name
    
    def polish_falcons2(self):
        output_name = self.pdf_output_name
        carrier = "Polish Falcons of America"
        data = []

        # tables = self.extract_text_from_range(start_page=0)
        tables = self.extract_tables_from_pdf()
        
        selling_agent_pattern = r'Selling Agent: (\w+) ([a-zA-Z 0-9]+)'
        # print(f"Table is {tables}")
        agent_no = ""
        agency = ""
        agent_count = 0
        for table in tables:
            if len(table) > 2:
                if "Selling Agent" in table[1][0] or "Selling Agent" in table[2][0]:
                    if agent_count != 0:
                        data.append({
                        "Extra": agent_no + " - " + agency
                        })
                    
                    selling_agent = re.match(selling_agent_pattern,table[1][0])
                    if selling_agent is None:
                        selling_agent = re.match(selling_agent_pattern,table[2][0])
                    agent_no,agency = selling_agent.groups() 
                    info = [x for x in table[2:] if x[0] and x[1]]
                    for row in info:
                        # print(row)
                        data.append({
                            "Carrier": carrier,
                            "Agency": agency,
                            "Policy": row[0],
                            "Insured/Anuitant": row[1],
                            "Plan": row[3],
                            "Issue Date": row[2],
                            "Mode": "",
                            "Value": "",
                            "Base Premium": row[5],
                            "Age": "",
                            "Year": row[6],
                            "Selling Agent": agent_no,
                            "Agent Share": "",
                            "Payment Date": row[4],
                            "Payment": "",
                            "Percent": row[8],
                            "Earned": row[9],
                            "Advanced": row[10],
                            "Repaid": row[11],
                            "Paid to Agent": row[12]
                        })
                    agent_count += 1
                else:
                    info = [x for x in table if x[0] and x[1]]
                    if not info:
                        continue
                    for row in info:
                        # print(row)
                        data.append({
                            "Carrier": carrier,
                            "Agency": agency,
                            "Policy": row[0],
                            "Insured/Anuitant": row[1],
                            "Plan": row[3],
                            "Issue Date": row[2],
                            "Mode": "",
                            "Value": "",
                            "Base Premium": row[5],
                            "Age": "",
                            "Year": row[6],
                            "Selling Agent": agent_no,
                            "Agent Share": "",
                            "Payment Date": row[4],
                            "Payment": "",
                            "Percent": row[8],
                            "Earned": row[9],
                            "Advanced": row[10],
                            "Repaid": row[11],
                            "Paid to Agent": row[12]
                        })
            
        data.append({
                        "Extra": agent_no + " - " + agency
                        })
              
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name
                

    def kskj_Life(self):
        output_name = self.pdf_output_name
        data = []
        carrier = "KSKJ Life"
        text = self.extract_text()

        date_pattern = r'period ending (\d+\/\d+\/\d+)'
        date = re.search(date_pattern,text)
        
        # agency_pattern = r'.+?\n.+?\n(.+)'
        # agency = re.search(agency_pattern,text)
        
        agency_pattern = r'^[a-zA-Z0-9 -]+\n[a-zA-Z0-9 ,-]+\n([a-zA-Z0-9 ,]+)(?:TRANSACTION CODES)?'
        agency_name = re.search(agency_pattern,text)
        if "TRANSACTION CODES" in agency_name.group(1):
            agency = re.sub(r" TRANSACTION CODES","",agency_name.group(1))
        else:
            agency = agency_name.group(1)
        
        agent_pattern = r'Agt Code \d+ (.*?) Pay Method'
        agent = re.search(agent_pattern,text)
        
        info_pattern = r'Writing Agent(.+?)Total Commission Amount'
        info = re.findall(info_pattern,text,re.DOTALL)
        
        client_pattern = r'^(\d+) (\w+) ([a-zA-Z ,-]+) *(\d) (\d+ )?(\d+\/\d+ )?(-?\d{,2}) ([0-9.,-]+) ?([0-9.,-]+) ?([0-9.,-]+)$'
        
        for i in info:
            clients = re.findall(client_pattern,i,re.MULTILINE)
            for client in clients:
                data.append({
                    "Carrier": carrier,
                    "Date": date.group(1),
                    "Agency": agency,
                    "Writing Agent": client[0],
                    "Agt Name": agent.group(1),
                    "Policy No.": client[1],
                    "Description": client[2],
                    "Code": client[3],
                    "Dur.": client[4],
                    "Date Due": client[5],
                    "Mths": client[6],
                    "Premium": client[7],
                    "Rate": client[8],
                    "Commission": client[9], 
                })

        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name
    
    def USAHealth(self):
        output_name = self.pdf_output_name
        data = []
        tables = self.extract_tables_from_pdf()     
        for table in tables:
            for row in table:
                if row[0] != "DATE":
                    name,policy_no = row[2].split(" |")
                    agent_name,agent_no = row[3].split("\n(")
                    agent_no = agent_no.strip(")")
                    data.append({
                        "Date":row[0],
                        "TypeID": "",
                        "Carrier": "USA Health Plans",
                        "FMO": "",
                        "Age": "",
                        "Agency": row[1].split("\n(")[0],
                        "Applied to Advance": "",
                        "Chargeback Amount": "",
                        "Client First": "",
                        "Client Full Name": name,
                        "Client Last": "",
                        "CMS Payment Type": "",
                        "Code": "",
                        "Commission Amount": row[7],
                        "Cycle Year": "",
                        "Date Due": "",
                        "Description": row[5],
                        "Document Type": "",
                        "Duration": "",
                        "Effective Date": "",
                        "Line of Business": "",
                        "Mode": "",
                        "Months": "",
                        "Override": "",
                        "Plan": "",
                        "Policy No": policy_no,
                        "Premium": "",
                        "Product": row[4],
                        "Rate": "",
                        "Retro": "",
                        "Signed Date": "",
                        "Split": "",
                        "State": "",
                        "Termination Date": "",
                        "Termination Reason": "",
                        "Value": "",
                        "Writing Agent": agent_name,
                        "Writing Agent NPN": "",
                        "Writing Agent Number": agent_no,
                    })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name        

    
    def bcbs(self):
        states = {
            "Illinois": "IL",
            "Montana": "MT",
            "New Mexico": "NM",
            "Oklahoma": "OK",
            "Texas": "TX",
        }
        output_name = self.pdf_output_name
        data = []

        text = self.extract_text(start=2)
        
        period_ending_pattern = r'For Period Ending: ([a-zA-Z]+ [0-9]{,2},? [0-9]{4})'
        state_pattern = r'State: ([a-zA-Z]+)'
        producer_pattern = r'([0-9]+)-\w+-([a-zA-Z0-9 ]+)'
        producers_pattern = r'Producer\/Sub-Producer: (.*)Totals for GA\/Producer'
           
        try:
            producers = re.findall(producers_pattern,text,re.DOTALL)
            period_ending = re.search(period_ending_pattern,text).group(1)
            state = re.search(state_pattern,text).group(1)
        except AttributeError:
            return
               
        try:
            state_initials = states[state]
        except KeyError:
            state_initials = ""
        
        carrier = f"BCBS {state_initials}"
        
        for producer in producers:
            producer_info = re.search(producer_pattern,producer)
            producer_no = producer_info.group(1)
            producer_name = producer_info.group(2)
            producer_tables = re.findall(r'^Individual (.*?)Totals for Individual',producer,re.MULTILINE | re.DOTALL)
            
            for table in producer_tables:
                commission_type = re.search(r'([a-zA-Z0-9 ]+)\nSource ID',table).group(1)
                rows = re.findall(r'([0-9]+) ?([a-zA-Z, ]+) (\d+\/\d+\/\d+) ?(\d+\/\d+\/\d+ )?([a-zA-Z\/\$\%_ ]+) (\d+ )?([0-9.]+) (\d+) ([0-9]{2}\/[0-9]{2}\/[0-9]{2}) ?([0-9.]+ )?([0-9.\$\%]+)?\n(\w+) (\w+¬?) ([0-9]{2}\/[0-9]{2}\/[0-9]{2}) (\w+) ([0-9]{2}\/[0-9]{2}\/[0-9]{2}) ([0-9.]+) ([0-9.]+ )?([0-9.]+)',table,re.DOTALL|re.MULTILINE)
                for column in rows:
                    data.append({
                        "Period Ending": period_ending,
                        "Carrier": carrier,
                        "Producer/Sub-Producer": producer_name,
                        "Producer No": producer_no,
                        "Type": commission_type,
                        "Acct/Policy": column[0],
                        "Group No": column[11],
                        "Acct/Pol Name": column[1],
                        "Product Name": column[12],
                        "Orig Eff Dt": column[2],
                        "PR Eff Dt": column[13],
                        "Cancel Dt": column[3],
                        "Calc Method": column[4],
                        "Funding Type": column[14],
                        "Contracts": column[5],
                        "Split %": column[6],
                        "Pol Mos": column[7],
                        "Pd From Dt": column[8],
                        "Pd To Dt": column[15],
                        "Premium Applied": column[9],
                        "Comm Rate": column[10],
                        "Comm Amt": column[16],
                        "YTD Premium": column[17],
                        "YTD Commission": column[18]
                    })

        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name   
        
    def family_benefit_life(self):
        output_name = self.pdf_output_name
        data = []
        text = self.extract_text()
       
        date_pattern = r'THRU ([a-zA-Z0-9 ,]+)'
        producers_pattern = 'PRODUCED BY (.*?)ENDING BALANCE'
        writing_agent_pattern = r'^([a-zA-Z -]+)$'
        agency_pattern = r'NAME AGENT NUMBER STATUS\n(.*) \w+ \d+'
        
        date = re.search(date_pattern,text).group(1)
        agency = re.search(agency_pattern,text).group(1)
        producers = re.findall(producers_pattern,text,re.DOTALL|re.MULTILINE)
        
        for producer in producers:
            writing_agent = re.match(writing_agent_pattern,producer,re.MULTILINE).group(1)
            rows = re.findall(r'(\w+) ([a-zA-z- .]+) ([A-Z]{,2}) ([0-9]{,3}) (?:[0-9]+ )?(?:\d+\/\d+) (\d+\/\d+) ([0-9]+ )?(1?[0-9]{,2}.[0-9]{2} )?([0-9.,-]+) (1?[0-9]{,2}.[0-9-]+ )?(?:[0-9.,]+ )?(?:[0-9.,]+ )?([0-9.-]+)',producer,re.DOTALL|re.MULTILINE)
            for column in rows:
                data.append({
                        "Date":date,
                        "TypeID": "",
                        "Carrier": "Family Benefit Life",
                        "FMO": "",
                        "Age": column[3],
                        "Agency": agency,
                        "Applied to Advance": column[8],
                        "Chargeback Amount": "",
                        "Client First": "",
                        "Client Full Name": column[1],
                        "Client Last": "",
                        "CMS Payment Type": "",
                        "Code": "",
                        "Commission Amount": column[9],
                        "Cycle Year": "",
                        "Date Due": "",
                        "Description": "",
                        "Document Type": "",
                        "Duration": column[5],
                        "Effective Date": column[4],
                        "Line of Business": "",
                        "Mode": column[2],
                        "Months": "",
                        "Override": "",
                        "Plan": "",
                        "Policy No": column[0],
                        "Premium": column[7],
                        "Product": "",
                        "Rate": column[6],
                        "Retro": "",
                        "Signed Date": "",
                        "Split": "",
                        "State": "",
                        "Termination Date": "",
                        "Termination Reason": "",
                        "Value": "",
                        "Writing Agent": writing_agent,
                        "Writing Agent NPN": "",
                        "Writing Agent Number": "",
                })
        
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name  
    
    def river_health(self):
        output_name = self.pdf_output_name
        data = []
        carrier = "River Health"
        text = self.extract_text()
        
        statement_date_pattern = r'Statement Date (\d+\/\d+\/\d+)' 
        statement_date = re.search(statement_date_pattern,text).group(1)
        
        tables_pattern = r'Adjustment\(s\)(.+)Total Commission'
        tables = re.findall(tables_pattern,text,re.DOTALL)
        
        producers_pattern = r'([a-zA-Z ]+) \d+ ([a-zA-Z0-9 +]+) (\$ \d+.\d+)\s?(\$ \d+.\d+)?'
        
        for table in tables:
            rows = re.findall(producers_pattern,table,re.DOTALL|re.MULTILINE)
            for column in rows:
                data.append({
                        "Date":statement_date,
                        "TypeID": "",
                        "Carrier": carrier,
                        "FMO": "",
                        "Age": "",
                        "Agency": "",
                        "Applied to Advance": "",
                        "Chargeback Amount": "",
                        "Client First": "",
                        "Client Full Name": column[1],
                        "Client Last": "",
                        "CMS Payment Type": "",
                        "Code": "",
                        "Commission Amount": column[2],
                        "Cycle Year": "",
                        "Date Due": "",
                        "Description": "",
                        "Document Type": "",
                        "Duration": "",
                        "Effective Date": "",
                        "Line of Business": "",
                        "Mode": "",
                        "Months": "",
                        "Override": "",
                        "Plan": "",
                        "Policy No": "",
                        "Premium": "",
                        "Product": "",
                        "Rate": "",
                        "Retro": "",
                        "Signed Date": "",
                        "Split": "",
                        "State": "",
                        "Termination Date": "",
                        "Termination Reason": "",
                        "Value": "",
                        "Writing Agent": column[0],
                        "Writing Agent NPN": "",
                        "Writing Agent Number": "",
                })
                if column[3]:
                    data.append({
                        "Date":statement_date,
                        "TypeID": "",
                        "Carrier": carrier,
                        "FMO": "",
                        "Age": "",
                        "Agency": "",
                        "Applied to Advance": "",
                        "Chargeback Amount": "",
                        "Client First": "",
                        "Client Full Name": column[1],
                        "Client Last": "",
                        "CMS Payment Type": "",
                        "Code": "",
                        "Commission Amount": column[3],
                        "Cycle Year": "",
                        "Date Due": "",
                        "Description": "",
                        "Document Type": "",
                        "Duration": "",
                        "Effective Date": "",
                        "Line of Business": "",
                        "Mode": "",
                        "Months": "",
                        "Override": "",
                        "Plan": "",
                        "Policy No": "",
                        "Premium": "",
                        "Product": "",
                        "Rate": "",
                        "Retro": "",
                        "Signed Date": "",
                        "Split": "",
                        "State": "",
                        "Termination Date": "",
                        "Termination Reason": "",
                        "Value": "",
                        "Writing Agent": column[0],
                        "Writing Agent NPN": "",
                        "Writing Agent Number": "",
                })
                    
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name   
    
    def kaiser_permanente(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        data = []
        
        agency_pattern = r'Commission Month: [a-zA-Z0-9\/ -]+\n(.*?)Total'
        agency = re.search(agency_pattern,text,re.DOTALL).group(1)
        
        vendor_pattern = r'Vendor # (\w+)'
        vendor = re.search(vendor_pattern,text,re.DOTALL|re.MULTILINE).group(1)
        
        vendorID_pattern = r'Vendor ID (\w+)'
        vendorID = re.search(vendorID_pattern,text,re.DOTALL|re.MULTILINE).group(1)
        
        date_pattern = r'Commission Month: ([0-9\/ -]+)'
        date = re.search(date_pattern,text,re.DOTALL|re.MULTILINE).group(1)
        
        subscribers_pattern = r'Paid Commission(.*?)Totals'
        subscribers_table = re.findall(subscribers_pattern,text,re.DOTALL|re.MULTILINE)
        
        kpif_pattern = r'KPIF(.*?)KPIF Members'
        kpif_table = re.findall(kpif_pattern,text,re.DOTALL|re.MULTILINE)
        
        agent_pattern = r'([a-zA-Z0-9, ]+) \((\w+)\)'
        subscribers_list_pattern = r'(\d+) ([a-zA-Z0-9 &-]+) ([a-zA-Z]{2,4}) ([0-9]{,4}) (\$\d+.\d+) (\$\d?,?\d+.\d+) (\d+%) (\$\d?,?\d+.\d+)\n([a-zA-Z ]+)?'
        for subscribers in subscribers_table:
            # print(subscriber)
            agent, agent_id = re.search(agent_pattern,subscribers,re.DOTALL|re.MULTILINE).groups()
            print(agent)
            print(agent_id)
            subscribers_list = re.findall(subscribers_list_pattern,subscribers,re.DOTALL|re.MULTILINE)
            for subscriber in subscribers_list:
                data.append({
                    "Carrier": "Kaiser Permanente Colorado",
                    "Agency": agency,
                    "Vendor #": vendor,
                    "AP Vendor ID": vendorID,
                    "Commission Month": date,
                    "Agent Name": agent,
                    "Agent ID": agent_id,
                    "Policy Number":subscriber[0],
                    "Client Name": subscriber[1] + " " + subscriber[8] if "For questions" not in subscriber[8] and subscriber[8] != "" else subscriber[1],
                    "Renewal Month": subscriber[2],
                    "Subscribers": subscriber[3],
                    "Amount from PSPM Schedule": subscriber[4],
                    "Payment Received Date": "",
                    "Premium": subscriber[5],
                    "%Dues Paid": subscriber[6],
                    "Commission": subscriber[7],
                    "Rate": "",
                    "Commission Last Paid": "",
                    "Total Commission": "",
                })
                
        for subscribers in kpif_table: 
            subscribers_list_pattern = r'([a-zA-Z0-9, ]+ \(\w+\).+)'

            subscribers_list = re.search(subscribers_list_pattern, subscribers,re.DOTALL|re.MULTILINE).group(1)

            header_pattern = re.compile(r"^[A-Za-z]+,[A-Za-z ]+ \(\w+\)$")

            blocks = []
            current_block = []

            for line in subscribers_list.split("\n"):
                if header_pattern.match(line):
                    # If we already have a block, save it before starting a new one
                    if current_block:
                        blocks.append("\n".join(current_block))
                        current_block = []
                # Always add the line to the current block
                current_block.append(line)

            # Add the last collected block
            if current_block:
                blocks.append("\n".join(current_block))
        kpif_member_pattern = r'([a-zA-Z, ]+) (\d+\/\d+\/\d+ )?(\$\d?,?\d+.\d+ )?(\d?,?\d+.\d+ % )?(\$\d?,?\d+.\d+ )?(\d+\/\d+\/\d+) (\$\d?,?\d+.\d+)'
        for table in blocks:
            # print(table)
            agent, agent_id = re.search(agent_pattern,table,re.DOTALL|re.MULTILINE).groups()
            subscribers = re.findall(kpif_member_pattern,table,re.DOTALL|re.MULTILINE)
            for i in range(len(subscribers)):
                data.append({
                    "Carrier": "Kaiser Permanente Colorado",
                    "Agency": agency,
                    "Vendor #": vendor,
                    "AP Vendor ID": vendorID,
                    "Commission Month": date,
                    "Agent Name": agent,
                    "Agent ID": agent_id,
                    "Policy Number":"",
                    "Client Name": subscribers[i][0],
                    "Renewal Month": "",
                    "Subscribers": "",
                    "Amount from PSPM Schedule": "",
                    "Payment Received Date": subscribers[i][1] if subscribers[i][1] != "" else next(
                        (subscribers[j][1] for j in range(i - 1, -1, -1) if subscribers[j][1] != ""), ""
                    ),
                    "Premium": subscribers[i][2] if subscribers[i][2] != "" else next(
                        (subscribers[j][2] for j in range(i - 1, -1, -1) if subscribers[j][2] != ""), ""
                    ),
                    "%Dues Paid": subscribers[i][3] if subscribers[i][3] != "" else next(
                        (subscribers[j][3] for j in range(i - 1, -1, -1) if subscribers[j][3] != ""), ""
                    ),
                    "Commission": "",
                    "Rate": subscribers[i][4] if subscribers[i][4] != "" else next(
                        (subscribers[j][4] for j in range(i - 1, -1, -1) if subscribers[j][4] != ""), ""
                    ),
                    "Commission Last Paid": subscribers[i][5],
                    "Total Commission": subscribers[i][6],
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name  
            
    def delta_dental_colorado(self):
        output_name = self.pdf_output_name
        carrier = "Delta Dental of Colorado"
        data = []
        text = self.rotate_pdf()
       
        # text = self.extract_text(start=1)
        full_text = self.extract_text_delta(text=text)
        # print(f"After rotation {full_text}")
        date_pattern = r'Billing Period: ([0-9\/ -]+)'
        date = re.search(date_pattern,full_text).group(1)

        broker_agency_pattern = r'Broker ID: (\w+)\n([a-zA-Z0-9 ]+)'
        broker_id, agency = re.search(broker_agency_pattern,full_text,re.DOTALL|re.MULTILINE).groups()
        
        commissions_pattern = r'([a-zA-Z0-9 \/\'&\,.-]+) (?:\(\w+\) )?(\([a-zA-Z0-9 -]+\))(.*?)Total Entries'
        commissions_tables = re.findall(commissions_pattern,full_text,re.DOTALL|re.MULTILINE)
        
        rows_pattern = r'(?:[0-9-]+) ([a-zA-Z -]+ )?([A-Z]{,4} [0-9]{4}) (-?\d?,?\d+.\d+) (-?\d?,?\d+.\d+) (-?\d?,?\d+.\d+) (-?\d?,?\d+.?\d?) (-?\$\d?,?\d+.\d+)'
        
        for tables in commissions_tables:
            group_name,group_number = tables[0],tables[1].strip("()")
            if "-" in group_number:
                group_number = group_number.split("-")[0]
            # groups = re.search(group_name_number_pattern,tables[0]).groups()
            rows = re.findall(rows_pattern,tables[2],re.DOTALL|re.MULTILINE)
            for row in rows:
                # print(row)
                data.append({
                    "Carrier": carrier,
                    "Billing Period": date,
                    "Broker ID": broker_id,
                    "Agency": agency,
                    "Group Name": group_name,
                    "Group Number": group_number,
                    "Subscriber Name": row[0],
                    "Billing Month": row[1],
                    "Invoice Amount": row[2],
                    "Premium Received": row[3],
                    "Commission Basis": row[4],
                    "Rate": row[5],
                    "Commission Amount": row[6]
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name  
    
    def allied(self):
        output_name = self.pdf_output_name
        data = []
        carrier = "Allied Benefit"
        text = self.extract_text()
        
        payment_agency_date_pattern = r'Payment type: (\w+)\n([a-zA-Z0-9 -]+)\nReport Date (\d+\/\d+\/\d+)'
        payment_type, agency, date = re.search(payment_agency_date_pattern,text,re.DOTALL|re.MULTILINE).groups()
        
        tables_pattern = r'Paid Amount(.*?)Total for Group'
        tables = re.findall(tables_pattern,text,re.MULTILINE|re.DOTALL)
        
        client_pattern = r'(\w+) ([a-zA-Z0-9 -&]+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) (\$\d*,?\d+.\d+) (\$\d*,?\d+.\d+) (\d+%) ([a-zA-Z0-9 ]+) (\d+) (\$\d*,?\d+.\d+)'
        agent_pattern = r'Writing Agent Number: (\d+) Writing Agent Name: ([a-zA-Z0-9 -]+)\nWriting Agent 2 No: (\d+ )?Writing Agent 2 Name: ?([a-zA-Z0-9 -]+)?'
        for table in tables:
            clients = re.findall(client_pattern,table,re.MULTILINE|re.DOTALL)
            agents = re.findall(agent_pattern,table,re.DOTALL|re.MULTILINE)
            for i in range(len(agents)):
                for j in range(len(clients)):
                    data.append({
                        "Carrier": carrier,
                        "Agency": agency,
                        "Payment Type": payment_type,
                        "Report Date": date,
                        "Writing Agent Number": agents[i][0],
                        "Writing Agent Name": agents[i][1],
                        "Writing Agent 2 No": agents[i][2],
                        "Writing Agent 2 Name": agents[i][3],
                        "Group No.": clients[j][0],
                        "Group Name": clients[j][1],
                        "Billing Period": clients[j][2],
                        "Adj. Period": clients[j][3],
                        "Invoice Total": clients[j][4],
                        "Stoploss Total": clients[j][5],
                        "Agent Rate": clients[j][6],
                        "Calculation Method": clients[j][7],
                        "Census Ct.": clients[j][8],
                        "Paid Amount": clients[j][9],
                    })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
    
    def delta_dental_virginia(self):
        carrier = "Delta Dental of Virginia"
        output_name = self.pdf_output_name
        data = []
        text = self.extract_text()

        agency_pattern = r'Agency: (.*?) Tax'
        agency = re.search(agency_pattern,text).group(1)
        
        agents_pattern = r'(\d+) (\w+) ([a-zA-Z0-9 ]+) (\$ \d*,?\d+.\d+) (\$ \d*,?\d+ ?\d*.\d+) ([a-zA-Z0-9-]+) ([a-zA-Z0-9 ]+)'
        
        tables_pattern = r'PolicyHolder (?:.*?)\n(.*)'
        tables = re.findall(tables_pattern,text,re.MULTILINE|re.DOTALL)
        for table in tables:
            rows = re.findall(agents_pattern,table,re.MULTILINE|re.DOTALL)
            for columns in rows:
                data.append({
                    "Carrier": carrier,
                    "Agency": agency,
                    "PolicyHolder #": columns[0],
                    "PH Last Name": columns[1],
                    "Product Name": columns[2],
                    "Premium": columns[3],
                    "Comm": columns[4],
                    "Month": columns[5],
                    "Broker Name": columns[6],
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
    
    def peek_performance(self):
        output_name = self.pdf_output_name
        data = []
        text = self.extract_text_from_range(0)
        
        agency_pattern = '\d+\/\d+\/\d+$\n[a-zA-Z0-9 ]+\n([a-zA-Z0-9 ]+)'
        agency = re.search(agency_pattern,text,re.MULTILINE|re.DOTALL).group(1)
        
        statements_pattern = r'(.*?)End of Statement'
        statements = re.findall(statements_pattern,text,re.MULTILINE|re.DOTALL)
    
        
        types_pattern = r'([a-zA-Z0-9 ]+)\nPolicy(.*?)Total'
        clients_pattern = r'(?:^(\w+)\n(?:([A-Z]+, [A-Z]+) ([A-Z0-9 ]+)\n|([A-Z]* ?[A-Z]+,? ?[A-Z]+ ?[a-zA-Z]{0,1}) ((?:AMBR)* [A-Z0-9 ]+)\n|([A-Z]+, [A-Z]+ ?[a-zA-Z]*)\n(.*?)\n)\n?([A-Z0-9-, ]+) (\d+\/\d+\/\d+)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n|^(?:(\w+) ([a-zA-Z]+, [a-zA-Z]+ ?[A-Z]?)|(\w+) ([a-zA-Z,]+ ?-? [a-zA-Z]+ ?[a-zA-Z]{0,1}) ([a-zA-Z0-9 ]+))\n([a-zA-Z0-9 ]+)?\n?^([a-zA-Z0-9, -]+) (\d+\/\d+\/\d+)\n(\d+\/\d+\/\d+)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n)'
        for statement in statements:
            types = re.findall(types_pattern,statement,re.MULTILINE|re.DOTALL)
            for tpe in types:
                
                print(tpe[1])
                clients = re.findall(clients_pattern,tpe[1],re.DOTALL|re.MULTILINE)
                print(len(clients))
                adjusted = []
                for client in clients:
                    formatted_client = [c for c in client if c != ""]
                    adjusted.append(formatted_client)
                    data.append({
                        "FMO": "Peek Performance",
                        "Agency": agency,
                        "Type": tpe[0],
                        "Policy No.": formatted_client[0],
                        "Policyholder": formatted_client[1],
                        "Product": formatted_client[2],
                        "Writing Agent": formatted_client[3],
                        "Effective": formatted_client[4],
                        "Paid-To": formatted_client[5],
                        "Comm. Premium": formatted_client[6],
                        "Rate": formatted_client[7],
                        "Earned Credit": formatted_client[8],
                        "Earned Commission": formatted_client[9],
                        "Escrow Adjustment": formatted_client[10],
                    })
                print(len(adjusted))
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
        
                
            
    def save_to_excel(self, df, output_name):
        """Save DataFrame to an Excel file and return the file path."""
        if df is None or df.empty:
            return
        output = BytesIO()

        # Write DataFrame to Excel file in memory
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
            worksheet = writer.sheets["Sheet1"]
            website_url = "https://comtrack.io"
            for row_num in range(1, 2):
                worksheet.write_url(row_num, df.columns.get_loc('Converted from .pdf by'), website_url, string="ComTrack.io")

        # Save the file to /tmp for consistent access
        output.seek(0)
        filename = f"{output_name}.xlsx"
        file_path = os.path.join("/tmp", filename)

        with open(file_path, 'wb') as file:
            file.write(output.read())

        return file_path
