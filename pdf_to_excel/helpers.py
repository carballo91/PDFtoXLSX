import os
import re
import pandas as pd
from io import BytesIO
import pdfplumber
from concurrent.futures import ProcessPoolExecutor
import gc

class PDFEditor:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.pdf_output_name = pdf_file.name.rstrip(".pdf")

    def is_valid_pdf(self):
        """Check if the uploaded file is a valid PDF."""
        return self.pdf_file.name.endswith(".pdf")

    
    def extract_page_text(self, pdf_path, page_num):
        with pdfplumber.open(pdf_path) as pdf:
            page_text = pdf.pages[page_num].extract_text()
            gc.collect()  # Call garbage collection to free up memory
            return page_text or ""

    def extract_text(self, pages=None):
        """Extract text from specified pages or all pages in the PDF sequentially."""
        
        with pdfplumber.open(self.pdf_file) as pdf:
            if pages is None:
                pages = range(len(pdf.pages))
        
            text_list = []
            for page_num in pages:
                page_text = self.extract_page_text(self.pdf_file, page_num)
                text_list.append(page_text)
                gc.collect()  # Optional: Call garbage collection after each page extraction

        # Join extracted text from all pages.
        return "\n".join(text_list)
    #Used for proccess pdf type1
    def extract_large_table(self,start_page, agent_table_number):
        full_table = []
        table_found = False
        with pdfplumber.open(self.pdf_file) as pdf:
            print(f"Lenght is {len(pdf.pages)}")
            print(agent_table_number)
            for page_num in range(start_page, len(pdf.pages)):
                page = pdf.pages[page_num]
                tables = page.extract_tables()
                print(f"Tables length {len(tables)}")
                if tables:
                    if not table_found:
                        full_table.extend(tables[agent_table_number])
                        table_found = True
                    elif len(tables) > agent_table_number:
                        full_table.extend(tables[agent_table_number])
                elif table_found:
                    break
        return full_table
    
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
            print(len(large_table))

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

        #print(f"Todo el texto es\n{text}\nAqui termina")
        data_table = []
        pattern = r"(\d+)\s([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z-]+){0,2})\s([a-zA-Z-]+\s[a-zA-Z-]+(?:\s[a-zA-Z-]+){0,2})\s(\d+)\s(\w+)\s(\d+\/\d+\/\d+)\s(Premium Payment)\s(\$(?:\d+,)?\d+.\d+)(\s\$?\w+)?(\s\$?\w+)?\s(\d+.\d+)\s(\(?\$(?:\d+,)?\d+.\d+\)?)"
        text_parts = re.search(r"Renewal LIFE\nCommission(.*?)Renewal Override LIFE", text, re.DOTALL)
        if text_parts:
            print(text_parts.group(1))
            match = re.findall(pattern,text_parts.group(1),re.DOTALL)
            for renewal_commission_results in match:
                print(renewal_commission_results)
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
        #print(df)

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
            #print(pfyc)
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
            #print(pfyc)
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
        print(text)
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
                #print(y[0])

        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name
        
    def bcbs_la_compensation(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        data = []
        print(text)
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

    def save_to_excel(self, df, output_name):
        """Save DataFrame to an Excel file and return the file path."""
        output = BytesIO()

        # Write DataFrame to Excel file in memory
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
            worksheet = writer.sheets["Sheet1"]

            # Add hyperlinks in the "Converted from .pdf by" column
            website_url = "https://comtrack.io"
            for row_num in range(1, 2):  # Start from 1 to skip the header
                worksheet.write_url(row_num, df.columns.get_loc('Converted from .pdf by'),
                                    website_url, string="ComTrack.io")

        # Save the file to a temporary location
        output.seek(0)
        filename = f'{output_name}.xlsx'
        file_path = os.path.join('/tmp', filename)

        # Write the BytesIO content to the file
        with open(file_path, 'wb') as file:
            file.write(output.read())

        return file_path, filename
