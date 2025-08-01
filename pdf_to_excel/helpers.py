import os
import re
import pandas as pd
from io import BytesIO
import pdfplumber
import time
import fitz
from pdfminer.pdfdocument import PDFPasswordIncorrect
from collections import defaultdict
import math

class PDFEditor:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.pdf_output_name = pdf_file.name.rstrip(".pdf")
        
    def is_image(self):
        with pdfplumber.open(self.pdf_file) as pdf:
            for i, page in enumerate(pdf.pages):
                print(f"Page {i+1}")
                print("Text:", page.extract_text())
                print("Is Image?", page.to_image() is not None)
                
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
    
    def is_rotated(self):
        
        with pdfplumber.open(self.pdf_file) as pdf:
            for i, page in enumerate(pdf.pages):
                rotation = page.rotation  # This returns 0, 90, 180, or 270
                if rotation > 0:
                    return True
        return False
    
    def extract_page_text(self, pdf, page_num):
        if 0 <= page_num < len(pdf.pages):
            return pdf.pages[page_num].extract_text() or ""
        return ""

    def extract_text(self,start=0, pages=None,password=None):
        """Extract text from specified pages or all pages in the PDF sequentially."""
        
        with pdfplumber.open(self.pdf_file,password=password) as pdf:
            if pages is None:
                pages = range(start,len(pdf.pages))
            else:
                pages = range(pages)
            text_list = []
            for page_num in pages:
                page_text = self.extract_page_text(pdf, page_num)
                text_list.append(page_text)
                # gc.collect()  # Optional: Call garbage collection after each page extraction

        # Join extracted text from all pages.
        return "\n".join(text_list)
    
    def extract_text_delta(self,text,start=0, pages=None):
        """Extract text from specified pages or all pages in the PDF sequentially."""
        print(f"Type of delta text is {type(text)}")
        with pdfplumber.open(text) as pdf:
            if pages is None:
                pages = range(start,len(pdf.pages))
            else:
                pages = range(pages)
            text_list = []
            for page_num in pages:
                page_text = self.extract_page_text(pdf, page_num)
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
    
    def extract_tables_from_pdf(self,password=None):
        all_tables = []
        with pdfplumber.open(self.pdf_file,password=password) as pdf:
            for page in pdf.pages:
                # Extract tables from the page
                tables = page.extract_tables()
                all_tables.extend(tables)
        return all_tables
    
    def extract_tables(self,password=None):
        tables = []  # List to store extracted tables
        with pdfplumber.open(self.pdf_file,password=password) as pdf:
            for page in pdf.pages:  # Iterate through each page
                # Extract tables from the current page
                page_tables = page.extract_tables()
                for table in page_tables:
                    # Convert the table into a DataFrame and append to the list
                    df = pd.DataFrame(table[1:], columns=table[0])  # Skip header
                    tables.append(df)
        return tables
    
    # Uses PyMuPDF
    def extract_text_from_range(self,start_page,end_page=None,password=None):
        # start_time = time.time()
        extracted_text = ""

        self.pdf_file.seek(0) 
        file_content = self.pdf_file.read()
        
        with fitz.open(stream=file_content,filetype="pdf") as pdf:
            if pdf.needs_pass:
                pdf.authenticate(password)
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
        print(text)

        # Extract 'Run Date' and 'Agents' information
        agents = re.findall(r"Subtotals for Agent (\w+)\s+([A-Z,.\s]+)", text)
        run_date = re.findall(r"Run Date:\s(\d{2}/\d{2}/\d{4})", text)

        if not agents:
            return None, None  # Handle invalid case

        new_data, extra_cols = [], []

        for i, agent in enumerate(agents):
            large_table = self.extract_large_table(4, i)
            print(large_table)

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

    def group_words_by_line(self,words, y_tolerance):
        lines = []

        for word in words:
            y = word['top']
            for line in lines:
                if abs(line['y'] - y) <= y_tolerance:
                    line['words'].append(word)
                    break
            else:
                lines.append({'y': y, 'words': [word]})

        # Convert to {y: [words]} dict
        grouped = defaultdict(list)
        for line in lines:
            grouped[round(line['y'], 2)].extend(line['words'])

        return grouped

    def sort_and_join_lines(self,lines):
        sorted_lines = []
        # print(f"Original lines are {lines}")
        for y in sorted(lines):
            # print(f"Lines y is {lines[y][0]["text"]}")
            # for text in lines[y]:
            #     print(text["text"])
            line = sorted(lines[y], key=lambda w: (round(w['top'], 2), w['x0']))


            text_line = ' '.join(w['text'] for w in line)
            # print(f"Text lines are {text_line}")
            
            sorted_lines.append((y, text_line, line))
        # print(f"Sorted lines are {sorted_lines}")
        return sorted_lines

    def smart_join_wrapped_lines(self,sorted_lines, x_tolerance=15):
        final_rows = []
        skip_next = False
        # print(f"Sorted lines are {sorted_lines}")
        for i in range(len(sorted_lines)):
            if skip_next:
                skip_next = False
                continue

            current_y, current_text, current_words = sorted_lines[i]

            # Look ahead to next line
            if i + 1 < len(sorted_lines):
                next_y, next_text, next_words = sorted_lines[i + 1]

                # Check if first word in next line aligns (x) with a word in current
                if abs(next_words[0]['x0'] - current_words[-1]['x0']) < x_tolerance:
                    # Merge lines (name wrap case)
                    merged_text = current_text + ' ' + next_text
                    final_rows.append(merged_text)
                    skip_next = True
                else:
                    final_rows.append(current_text)
            else:
                final_rows.append(current_text)

        return final_rows
    
    def parse_x_based_line(self, words, column_ranges):
        words = sorted(words, key=lambda w: (round(w['top'], 2), w['x0']))
        print(f"words are {words}")

        def get_text_in_range(x_min, x_max):
            return ' '.join(w['text'] for w in words if x_min <= w['x0'] < x_max)

        return [get_text_in_range(x_min, x_max) for (x_min, x_max) in column_ranges]

    def clean_lines_main(self,column_ranges,password=None,y_tolerance=8):
        output_text = ""
        with pdfplumber.open(self.pdf_file,password=password) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # print(f"\n--- Page {page_num + 1} ---\n")
                words = page.extract_words()

                # Step 1: Group by line
                lines = self.group_words_by_line(words,y_tolerance=y_tolerance)

                # Step 2: Sort lines and join text
                sorted_lines = self.sort_and_join_lines(lines)

                for y, text, word_group in sorted_lines:
                    parsed = self.parse_x_based_line(word_group,column_ranges)
                    if parsed:
                        output_text += ', '.join(parsed)
                        output_text += "\n"
                        # print(', '.join(parsed))
        return output_text
                    
    def forester_financial(self):
        column_ranges = [
            (70, 99),     # Producer ID
            (100, 160),   # Producer Name
            (161, 207),   # Insured Name
            (209, 238),   # Policy ID
            (263, 292),   # Plan Code
            (302, 338),   # Date
            (345, 401),   # Transaction Type
            (431, 456),   # Amount
            (457, 487),   # Repl Factor
            (488, 514),   # ?? (might be dup)
            (514, 535),   # ?? (might be dup)
            (536, 560),   # Comp
        ]
        
        commission_types = {
            "FYC": "First Year Commission",
            "FOV": "First Year Override",
            "RYC": "Renewal Commission",
            "ROV": "Renewal Override"
        }

        """Process another type of PDF (define structure here)."""
        output_name = self.pdf_output_name
        data = []
        text = self.clean_lines_main(column_ranges)
            
        tables_pattern = r'balance, forward:, , \$[0-9\.]+\n(.*?)current, balance'
        commissions_pattern = r'(.*?)total for, (\w+)'
        agents_pattern = r'^(\d+), ([a-z \-\,\.]+), ([a-z \-\,\.]+), (\d+), ([a-z0-9 ]+), ([0-9\/]+), ([a-z ]+), ([0-9\.\$]+),([0-9\. ]+),([0-9\. ]+), ([0-9\.]+), ([0-9\.\$]+)$'
        tables = re.findall(tables_pattern,text,re.DOTALL|re.MULTILINE|re.IGNORECASE)
        for table in tables:
            commissions = re.findall(commissions_pattern,table,re.MULTILINE|re.DOTALL|re.IGNORECASE)
            for commission in commissions:
                commission_type = commission[1].upper()
                agents = re.findall(agents_pattern,commission[0],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                for client in agents:
                    data.append(
                        {
                            "Carrier": "Foresters",
                            "type" : commission_types[commission_type],
                            "Writing Producer #": client[0],
                            "Writing Producer Name":client[1],
                            "Insured Name":client[2],
                            "Policy #": client[3],
                            "Plan Code": client[4],
                            "Transaction Date": client[5],
                            "Transaction Type": client[6],
                            "Compensation Basis Amount": client[7],
                            "% Repl. Factor":client[8],
                            "% Share": client[9],
                            "Comp Rate %": client[10],
                            "Amount Due": client[11]
                        }
                    )

        # Add custom column
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)

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

    def bcbs_la_commisions(self,carrier):
        output_name = self.pdf_output_name
        data = []
        text = self.extract_text()

        date = re.search(r"Activity Ending Date: (\d+\/\d+\/\d+)",text,re.DOTALL)
        date = date.group(1)

        statement_type_pattern = r'\w+, \w+, \d+ ([a-z ]+)'
        statement_type = re.search(statement_type_pattern,text,re.DOTALL|re.MULTILINE|re.IGNORECASE)
        
        producers_pattern = r'Writing Producer (\d+)\s(\w+\s\w+)\n(.*?)Total for Writing Producer'
        producers = re.findall(producers_pattern,text,re.MULTILINE|re.DOTALL)
        
        statements_pattern = r'^([a-z ]+)(.*?)total for'
        
        clients_pattern = r'^(\d+)\s([a-zA-Z]+(?:\s[a-zA-Z-]+){0,3})\s(\w+)\s(\d+(?:\s[a-zA-Z]+){0,4})\s(\w+)\s(\w+\s)?(\d+\/\d+\/\d+)\s(\d+\/\d+\/\d+\s)?(\d+\/\d+\/\d+)\s(\d+\/\d+)\s(\d+\s)?(\w+)\s(\(?\$ (?:\d+,)?\d+\.\d+\)?)$'
        
        
        for producer in producers:
            statements = re.findall(statements_pattern,producer[2],re.MULTILINE|re.IGNORECASE|re.DOTALL)
            writing_producer = producer[0]
            writing_id = producer[1]

            for statement in statements:
                statement_type = statement[0]

                clients = re.findall(clients_pattern,statement[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                for client in clients:
                    data.append({
                        "Carrier": carrier,
                        "Statement Type": statement_type,
                        "Date": date,
                        "Type": statement_type,
                        "Writing Producer": writing_producer,
                        "Writing ID": writing_id,
                        "Member ID": client[0],
                        "Name": client[1],
                        "Company": client[2],
                        "Product": client[3],
                        "HICN": client[4],
                        "Override": client[5],
                        "Effective Date": client[6],
                        "Term Date": client[7],
                        "Signed Date": client[8],
                        "Period": client[9],
                        "Cycle Year": client[10],
                        "Retro": client[11],
                        "Amount": client[12], 
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
        
        print(text)

        # agency = re.match(r'^([a-zA-Z ]+)\n\d+',text,re.DOTALL)
        agency = re.search(r'\w+\n([a-zA-Z ,]+?)\nNPN',text,re.DOTALL|re.MULTILINE)
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

        text = self.extract_text()
        # print(text)
        
        period_ending_pattern = r'For Period Ending: ([a-zA-Z]+ [0-9]{,2},? [0-9]{4})'
        state_pattern = r'State: ([a-zA-Z]+)'
        producer_pattern = r'([0-9]+)-\w+-([a-zA-Z0-9 ]+)'
        producers_pattern = r'Producer\/Sub-Producer: (.*)(?:Totals for GA\/Producer|producer statement total)'
           
        try:
            producers = re.findall(producers_pattern,text,re.DOTALL|re.IGNORECASE)
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
            # print(f"Producer is {producer}")
            producer_info = re.search(producer_pattern,producer)
            producer_no = producer_info.group(1)
            producer_name = producer_info.group(2)
            producer_tables = re.findall(r'^(.*?)Totals?(?: for)? Individual ([A-Za-z ]+) [0-9.]+$',producer,re.MULTILINE | re.DOTALL|re.IGNORECASE)
            # print(f"Producer tables are {producer_tables}")
            # print(f"{len(producer_tables)}")
            # print(producer_tables)
            # print(producer)
            for table in producer_tables:
                # print(f"Table is {table}")
                
                commission_type = table[1]
                rows = re.findall(r'([0-9]+) ?([a-zA-Z, ]+) (\d+\/\d+\/\d+) ?(\d+\/\d+\/\d+ )?([a-zA-Z\/\$\%_ ]+) (-?\d+ )?([0-9.]+) (\d+) ([0-9]{2}\/[0-9]{2}\/[0-9]{2}) ?(-?[0-9.]+ )?([0-9.\$\%]+)?(?: healthm)?.*?(\w+) (\w+¬?) ([0-9]{2}\/[0-9]{2}\/[0-9]{2}) (\w+) ([0-9]{2}\/[0-9]{2}\/[0-9]{2}) (-?[0-9.]+) ([0-9.]+ )?([0-9.]+)',table[0],re.DOTALL|re.MULTILINE)
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
        
        # print(text)
        
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
        subscribers_list_pattern = r'(\d+) ([a-zA-Z0-9 &-]+) ([a-zA-Z]{2,4}) ([0-9]{,4}) (\$\d+.\d+) ([0-9$.,]+) (\d+%) (\$\d?,?\d+.\d+)\n([a-zA-Z ]+)?'
        for subscribers in subscribers_table:
            print(subscribers)
            agent, agent_id = re.search(agent_pattern,subscribers,re.DOTALL|re.MULTILINE).groups()

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
        kpif_member_pattern = r'([a-zA-Z, ]+) (\d+\/\d+\/\d+ )?(\$\d?,?\d+.\d+ )?(\d?,?\d+.\d+ % )?(\$\d?,?\d+.\d+ )?(\d+\/\d+\/\d+ )?(\$\d?,?\d+.\d+)'
        for table in blocks:

            agent, agent_id = re.search(agent_pattern,table,re.DOTALL|re.MULTILINE).groups()
            n_st = table.split("\n")
            formatted = ""

            for i in n_st:
                if len(i) > 5 and "$" not in i:
                    i = i + " $0.00 $0.00"
                    formatted = formatted + i + " "
                    continue
                formatted = formatted + i + " "

            subscribers = re.findall(kpif_member_pattern,formatted,re.DOTALL|re.MULTILINE)
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
        
        # rows_pattern = r'(?:[0-9-]+) ([a-zA-Z -]+ )?([A-Z]{,4} [0-9]{4}) (-?\d?,?\d+.\d+) (-?\d?,?\d+.\d+) (-?\d?,?\d+.\d+) (-?\d?,?\d+.?\d?) (-?\$\d?,?\d+.\d+)'
        rows_pattern = r"^(?:[0-9-]+) ([a-zA-Z -']+ )?([A-Z]{,4} [0-9]{4}) (-?\d?,?\d+.\d+) (-?\d?,?\d+.\d+) (-?\d?,?\d+.\d+) (-?\d?,?\d+.?\d?) (-?\$\d?,?\d+.\d+)"
        
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
        
        is_rotated = self.is_rotated()
        if is_rotated:
            rotated_text = self.rotate_pdf()
            text = self.extract_text_delta(text=rotated_text)
        else:
            text = self.extract_text()
        
        try:
            agency_pattern = r'Agency: (.*?) Tax'
            agency = re.search(agency_pattern,text).group(1)
            
            agents_pattern = r'(\d+) (\w+) ([a-zA-Z0-9 ]+) (\$ \d*,?\d+.\d+) (\$ \d*,?\d+ ?\d*.\d+) ([a-zA-Z0-9-]+) ([a-zA-Z0-9 ]+)'
            
            tables_pattern = r'PolicyHolder (?:.*?)\n(.*)'
            tables = re.findall(tables_pattern,text,re.MULTILINE|re.DOTALL)
        except AttributeError:
            return (None,None)
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
        
        agency_pattern = r'\d+\/\d+\/\d+$\n[a-zA-Z0-9 ]+\n([a-zA-Z0-9 &]+)'
        agency = re.search(agency_pattern,text,re.MULTILINE|re.DOTALL).group(1)
        
        statements_pattern = r'(.*?)End of Statement'
        statements = re.findall(statements_pattern,text,re.MULTILINE|re.DOTALL)
    
        
        types_pattern = r'([a-zA-Z0-9 ]+)\nPolicy(.*?)Renewal'
        # clients_pattern = r'(?:^(\w+)\n(?:([A-Z]+, [A-Z]+) ([A-Z0-9 ]+)\n|([A-Z]* ?[A-Z]+,? ?[A-Z]+ ?[a-zA-Z]{0,1}) ((?:AMBR)* [A-Z0-9 ]+)\n|([A-Z]+, [A-Z]+ ?[a-zA-Z]*)\n(.*?)\n)\n?([A-Z0-9-, ]+) (\d+\/\d+\/\d+)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n|^(?:(\w+) ([a-zA-Z]+, [a-zA-Z]+ ?[A-Z]?)|(\w+) ([a-zA-Z,]+ ?-? [a-zA-Z]+ ?[a-zA-Z]{0,1}) ([a-zA-Z0-9 ]+))\n([a-zA-Z0-9 ]+)?\n?^([a-zA-Z0-9, -]+) (\d+\/\d+\/\d+)\n(\d+\/\d+\/\d+)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n)'
        clients_pattern = r'^(\w+)\n([a-z ]+,? [a-z]+)( [a-z0-9- ]+)?\n([a-z0-9- ]+\n)?([a-z0-9-, ]+)( [0-9]{2}\/[0-9]{2}\/[0-9]{4})?\n([0-9]{2}\/[0-9]{2}\/[0-9]{4})\n([0-9]{2}\/[0-9]{2}\/[0-9]{4}\n)?([0-9\$-\.]+)\n([0-9\.%]+)\n([0-9\$-\.]+)\n([0-9\$-\.]+)\n([0-9\$-\.]+)'
        clients_pattern_1 = r'^(\d+)\n(.*?)\n(.*?)\n(.*?)\n(\d+\/\d+\/\d+)\n(\d+\/\d+\/\d+)'
        for statement in statements:
            types = re.findall(types_pattern,statement,re.MULTILINE|re.DOTALL)
            for tpe in types:
                print(tpe[1])
                
                clients = re.findall(clients_pattern,tpe[1],re.DOTALL|re.MULTILINE|re.IGNORECASE)
                diff_pdf_clients = re.findall(clients_pattern_1,tpe[1],re.DOTALL|re.MULTILINE|re.IGNORECASE) 
                
                if clients:

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
                else:
                    for client in diff_pdf_clients:
                        data.append({
                            "FMO": "Peek Performance",
                            "Agency": agency,
                            "Type": tpe[0],
                            "Policy No.": client[0],
                            "Policyholder": client[1],
                            "Product": client[2],
                            "Writing Agent": client[3],
                            "Effective": client[4],
                            "Paid-To": client[5],
                            "Comm. Premium": "",
                            "Rate": "",
                            "Earned Credit": "",
                            "Earned Commission": "",
                            "Escrow Adjustment": "",
                        })
                    
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
    
    def life_shield(self):
            # (98, 161),   # Insured Name
            # (190, 219),   # Policy ID
            # (232, 244), 
        column_ranges = [
            (10, 72),     # Producer ID
            (74, 244),   # Producer Name
            (257, 269),   # Date
            (307, 341),   # Transaction Type
            (349, 378),   # Amount
            (378, 451),   # Repl Factor
            (451, 524),   # ?? (might be dup)
            (524, 593),   # ?? (might be dup)
            (593, 618),
            (618, 691),
            (715, 748),# Comp
        ]
        
        carrier = "LifeShield"
        passwords = ["2646","7964"]

        for password in passwords:
            try:
                test = self.clean_lines_main(column_ranges=column_ranges,password=password)
                break
            except PDFPasswordIncorrect:
                continue
        # print(f"TEST IS {test}")        
        data = []

        output_name = self.pdf_output_name

        if test:
            statements_pattern = r' e d i t s(.*?)Totals'
            tables_pattern = r'(^[a-z0-9- ,]+)\n(.*?)[Life|h] Subtotal[a-z0-9 .*-,]+\n?'
            managers_pattern = r'(.*?)Mgr, ([a-z ,]+) \((\w+)\)'
            agents_pattern = r'(.*?)WAgt, ([a-z ,]+) \((\w+)\)'
            clients_pattern = r'^(\w+), (\d+) ([a-z, ]+) (\d+\/\d+) (\d+), (\d+), ([0-9\.,]+),([0-9\. ,]+),([0-9\. ,]+),([0-9\. ,]+),([0-9\. ,]+),([0-9\. ,]+),([0-9\. ,]+),([a-z0-9\. ,]+)'
            
            statements = re.findall(statements_pattern,test,re.MULTILINE|re.DOTALL)
            agency_pattern = r'C o m m i s s i o n, ([A-Z0-9 &,]+) national'
            agency = re.search(agency_pattern,test,re.IGNORECASE).group(1)
            agency = agency.strip(", ")

            for statement in statements:
                tables = re.findall(tables_pattern,statement,re.MULTILINE|re.DOTALL|re.IGNORECASE)
                for table in tables:
                    category = table[0].strip(", ")
                    managers = re.findall(managers_pattern,table[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                    if managers:
                        for manager in managers:
                            manager_name = manager[1]
                            manager_id = manager[2]
                            agents = re.findall(agents_pattern,manager[0],re.DOTALL|re.IGNORECASE|re.MULTILINE)
                            for agent in agents:
                                w_agent = agent[1]
                                w_agent_id = agent[2]
                                clients = re.findall(clients_pattern,agent[0],re.IGNORECASE|re.MULTILINE|re.DOTALL)
                                for client in clients:
             
                                    data.append({
                                        "Carrier": carrier,
                                        "Agency": agency,
                                        "Category": category,
                                        "Policy Number": client[0],
                                        "Pol/Ridr": client[1],
                                        "Name or Description": client[2],
                                        "Date Paid": client[3],
                                        "Mode Pmts": client[4],
                                        "Pol Yr": client[5],
                                        "WAgt Name": w_agent,
                                        "WAgt ID": w_agent_id,
                                        "Mgr Name": manager_name,
                                        "Mgr ID": manager_id,
                                        "Premium": client[6],
                                        "Rate": client[7],
                                        "First Year": client[8],
                                        "Renewal": client[9],
                                        "Single": client[10],
                                        "Advance or ChgBack": client[11],
                                        "Charges": client[12],
                                        "Credits": client[13],
                                    })
                    else:
                        clients = re.findall(clients_pattern,table[1],re.IGNORECASE|re.MULTILINE|re.DOTALL)
                        for client in clients:
                 
                            data.append({
                                "Carrier": carrier,
                                "Agency": agency,
                                "Category": category,
                                "Policy Number": client[0],
                                "Pol/Ridr": client[1],
                                "Name or Description": client[2],
                                "Date Paid": client[3],
                                "Mode Pmts": client[4],
                                "Pol Yr": client[5],
                                "WAgt Name": "",
                                "WAgt ID": "",
                                "Mgr Name": "",
                                "Mgr ID": "",
                                "Premium": client[6],
                                "Rate": client[7],
                                "First Year": client[8],
                                "Renewal": client[9],
                                "Single": client[10],
                                "Advance or ChgBack": client[11],
                                "Charges": client[12],
                                "Credits": client[13],
                            })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
          
    def libery_bankers(self):
        output_name = self.pdf_output_name
        pw_name = self.pdf_output_name.split()[2]
        pw = pw_name.rstrip("Z")
        print("this is the output name")
        print(output_name)
        passwords = ["WG500","LBL22728"]
        passwords.append(pw)
        is_lbl = False
        for password in passwords:
            is_lbl = False
            try:
                text = self.extract_text(password=password)
                # """Delete this part"""
                # if password == "LBL22728" and output_name == "Agent Statements 22728 LBL22728ZZ 01-02-25 (1)":
                #     tables_from_pdf = self.extract_tables_from_pdf(password=password)
                #     is_lbl = True
                #     """Up to here"""
                break
            except PDFPasswordIncorrect:
                continue
        data = []
        
        carrier = "Liberty Bankers"
        # print(text)
        agency_pattern = r'n\s([a-z, &]+)\s+Liberty Bankers'
        agency = re.search(agency_pattern,text,re.IGNORECASE).group(1)
        
        statements_pattern = r'Beginning Balances.*?\n(.*?)Statement Totals'
        statements = re.findall(statements_pattern,text,re.MULTILINE|re.DOTALL)
 
        tables_pattern = r'([a-z0-9- ]+)\n(.*?)Life Subtotal [a-z0-9 .*-]+\n?'
        managers_pattern = r'(.*?)Mgr ([a-z ,]+) \((\w+)\)?'
        clients_pattern = r'^(\w+) (\d+) ([a-z, ]+) ([a-z0-9]+) (\d+) (\d+\/\d+\/\d+) (-?\d+) (-?\d+) (-?[0-9.,]+) (-?[0-9.,-]+) (-?[0-9.,-]+) (-?[0-9.,-]+) (-?[a-z0-9.,-]+ )?(-?[0-9.,-]+) ?(-?[0-9.,-]+)? ?(-?[0-9.,-]+)?\n'
        agent_pattern = r'(.*?)WAgt ([a-z, ]+) \((\w+)\)'
        wg598_clients_pattern = r'^(\w+) (\d+) ([a-z-, ]+) ([ssimplad|PSIMPL]+) (\d+) (\d+\/\d+\/\d+) (-?\d+) (-?\d+) (-?[0-9.,]+) (-?[0-9.,-]+) (-?[0-9.,-]+) (-?[0-9.,-]+) (-?[a-z0-9.,-]+)( -?[0-9.,-]+)?'
        for statement in statements:
            tables = re.findall(tables_pattern,statement,re.MULTILINE|re.DOTALL|re.IGNORECASE)
            for table in tables:
                # print(f"Table 1 {table[1]}")
                category = table[0]
                managers = re.findall(managers_pattern,table[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                agents = re.findall(agent_pattern,table[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                agents_wg598 = re.findall(wg598_clients_pattern,table[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                if managers:
                    for manager in managers:
                        manager_name = manager[1]
                        manager_id = manager[2]
                        agents = re.findall(agent_pattern,manager[0],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                        for agent in agents:
                            agent_name = agent[1]
                            agent_id = agent[2]
                            clients = re.findall(clients_pattern,agent[0],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                            for client in clients:
                                
                                data.append({
                                    "Carrier": carrier,
                                    "Agency": agency,
                                    "Category": category,
                                    "Policy Number": client[0],
                                    "Ph": client[1],
                                    "Name or Description": client[2],
                                    "Plan": client[3],
                                    "Age": client[4],
                                    "WAgt Name": agent_name,
                                    "WAgt ID": agent_id,
                                    "Mgr Name": manager_name,
                                    "Mgr ID": manager_id,
                                    "Date Paid": client[5],
                                    "Num Pmt": client[6],
                                    "Pol Dur": client[7],
                                    "Premium Paid": client[8],
                                    "Face Amount": client[9],
                                    "Comm Rate": client[10],
                                    "Earned Commission": client[11],
                                    "Advance or ChgBack": client[12],
                                    "Reserve Percent": "",
                                    "Unearned Advance": client[13],
                                    "Reserve Account": "",
                                    "Agent Account": client[14],
                                    "Loan Account": client[15],
                                })
                elif agents:
                    clients_pattern = r'^(\w+) (\d+) ([a-z, ]+) ([a-z0-9]+) (\d+) (\d+\/\d+\/\d+) (-?\d+) (-?\d+) (-?[0-9.,]+) (-?[0-9.,-]+) (-?[0-9.,-]+) (-?[0-9.,-]+) (-?[a-z0-9.,-]+ )?(-?[0-9.,-]+)( -?[a-z0-9.,-]+ )?\n'
                    for agent in agents:
                        agent_name = agent[1]
                        agent_id = agent[2]
                        clients = re.findall(clients_pattern,agent[0],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                        for client in clients:
                            data.append({
                                    "Carrier": carrier,
                                    "Agency": agency,
                                    "Category": category,
                                    "Policy Number": client[0],
                                    "Ph": client[1],
                                    "Name or Description": client[2],
                                    "Plan": client[3],
                                    "Age": client[4],
                                    "WAgt Name": agent_name,
                                    "WAgt ID": agent_id,
                                    "Mgr Name": "",
                                    "Mgr ID": "",
                                    "Date Paid": client[5],
                                    "Num Pmt": client[6],
                                    "Pol Dur": client[7],
                                    "Premium Paid": client[8],
                                    "Face Amount": client[9],
                                    "Comm Rate": client[10],
                                    "Earned Commission": client[11],
                                    "Advance or ChgBack": "",
                                    "Reserve Percent": "",
                                    "Unearned Advance": client[12],
                                    "Reserve Account": "",
                                    "Agent Account": client[13],
                                    "Loan Account": client[14],
                                })
                elif agents_wg598:
                    for client in agents_wg598:
                        data.append({
                            "Carrier": carrier,
                            "Agency": agency,
                            "Category": category,
                            "Policy Number": client[0],
                            "Ph": client[1],
                            "Name or Description": client[2],
                            "Plan": client[3],
                            "Age": client[4],
                            "WAgt Name": "",
                            "WAgt ID": "",
                            "Mgr Name": "",
                            "Mgr ID": "",
                            "Date Paid": client[5],
                            "Num Pmt": client[6],
                            "Pol Dur": client[7],
                            "Premium Paid": client[8],
                            "Face Amount": client[9],
                            "Comm Rate": client[10],
                            "Earned Commission": client[11],
                            "Advance or ChgBack": "",
                            "Reserve Percent": "",
                            "Unearned Advance": "",
                            "Reserve Account": "",
                            "Agent Account": client[12],
                            "Loan Account": client[13],
                        })
                elif is_lbl:
                    clients = re.findall(clients_pattern,table[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                    print(output_name)
                    print(f"Clients {clients}")
                    if not clients:return
                    row = tables_from_pdf[0][5:6]
                    
                    print(row[0][0])
                    print(output_name)
                    print(clients)
                    for client in row:
                        # print(text)
                        print(f"Client is {client}")
                        data.append({
                            "Carrier": carrier,
                            "Agency": agency,
                            "Category": category,
                            "Policy Number": clients[0][0],
                            "Ph": client[1],
                            "Name or Description": client[2],
                            "Plan": client[3],
                            "Age": client[4],
                            "WAgt Name": "",#agent_name,
                            "WAgt ID": "",#agent_id,
                            "Mgr Name": "",#manager_name,
                            "Mgr ID": "",#manager_id,
                            "Date Paid": client[5],
                            "Num Pmt": client[6],
                            "Pol Dur": client[7],
                            "Premium Paid": client[8],
                            "Face Amount": client[9],
                            "Comm Rate": client[10],
                            "Earned Commission": client[11],
                            "Advance or ChgBack": client[12],
                            "Reserve Percent": "",
                            "Unearned Advance": client[14],
                            "Reserve Account": "",
                            "Agent Account": client[16],
                            "Loan Account": client[17] if client[17] is not None else "",
                        })
                    
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
                    
    def baylor_scott(self):
        data_pd = []
        output_name = self.pdf_output_name
        carrier = "Baylor Scott & White"
        text = self.extract_text_from_range(start_page=0)
     
        agency_pattern = r'([a-z0-9 ]+) \((\d+)\)'
        date_pattern = r'Marketplace\n(\d+\/\d+\/\d+)'
        
        date = re.search(date_pattern,text).group(1)
        agency = re.search(agency_pattern,text,re.IGNORECASE)
        agency_name,agency_id = agency.group(1),agency.group(2)
        
        agents_data_pattern = r'Marketplace\n\d+\/\d+\/\d+(.*)Agency:'
        agents_pattern = r'^(?:[0-9]{2}\/[0-9]{2}\/[0-9]{4}\n)?([a-z ]+)\nMember(.*?)Agent: ([a-z ]+) Totals'
        agents_data = re.findall(agents_data_pattern,text,re.MULTILINE|re.DOTALL)
        
        clients_pattern = r'^(\b(?!rate\b)\w+)?\n?(\b(?!Commission Due\b)\w+[a-z ]+)?\n?(\d+\/\d+\/\d+)?\n?([0-9]{2}\/[0-9]{2}\/[0-9]{4} - \d+\/\d+\/\d+)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n'
        clients_pattern2 = r'^(\w+\n)?([a-z ]+\n)?(\w+)\n([a-z -]+)\n(-?\$ (?:\d+)?,?\d+.\d+)\n(-?\$ (?:\d+)?,?\d+.\d+)\n([a-z ]+)'
        for data in agents_data:
            
            agents = re.findall(agents_pattern,data,re.MULTILINE|re.DOTALL|re.IGNORECASE)
            # print(agents[:10000])
            # print(f"Agents {agents[1][:1000]}")
            for agent in agents:
                # print(agent[1][:1000])
                
                clients = re.findall(clients_pattern,agent[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                clients2 = re.findall(clients_pattern2,agent[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                if clients:
                   
                    member_name = ""
                    member_id = ""
                    effective_date = ""
                    for client in clients:
                        member_id = client[0] if client[0] != "" else member_id
                        member_name = client[1] if client[1] != "" else member_name 
                        effective_date = client[2] if client[2] != "" else effective_date
                        data_pd.append({
                            "Carrier": carrier,
                            "Agency": agency_name,
                            "Agency ID": agency_id,
                            "Commission Statement Date": date,
                            "Section": agent[0],
                            "Agent": agent[2],
                            "Member ID": member_id,
                            "Member Name": member_name,
                            "Original Effective Date": effective_date,
                            "Coverage Period": client[3],
                            "LOB": client[4],
                            "Plan Description": client[5],
                            "First or Renewal": client[6],
                            "Premium": client[7],
                            "Month": client[8],
                            "Rate": client[9],
                            "Commission": client[10],
                        })
                if clients2:
                    member_name = ""
                    member_id = ""
                    effective_date = ""
                    for client in clients2:
                        member_id = client[0] if client[0] != "" else member_id
                        member_name = client[1] if client[1] != "" else member_name
                        effective_date = client[2] if client[2] != "" else effective_date
                        data_pd.append({
                            "Carrier": carrier,
                            "Agency": agency_name,
                            "Agency ID": agency_id,
                            "Commission Statement Date": date,
                            "Section": agent[0],
                            "Agent": agent[2],
                            "Member ID": member_id,
                            "Member Name": member_name,
                            "Original Effective Date": "",
                            "Coverage Period": "",
                            "LOB": client[2],
                            "Plan Description": client[3],
                            "First or Renewal": client[6],
                            "Premium": client[4],
                            "Month": "",
                            "Rate": "",
                            "Commission": client[5],
                        })
        # print(data_pd)
        if len(data_pd) >= 1:
            for i in range(1):
                data_pd[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data_pd)
        # print(df)
        return df,output_name 
    def liberty_bankers_life(self):
        data = []
        output_name = self.pdf_output_name
        codes = {
            "1" : "Advance Commission",
            "2" : "Earned Commission",
            "3" : "Advance Recovery",
            "4" : "Advance Chargeback",
            "5" : "Earned Chargeback",
        }
        text = self.extract_text()
        # print(text)
        
        carrier = "Liberty Bankers Ancillary"
        
        agency_pattern = r'([a-z ]+) TRANSACTION CODES'
        agency = re.search(agency_pattern,text,re.IGNORECASE).group(1)
        
        agency_id_pattern = r'(\w+)\n[a-z ]+ TRANSACTION CODES'
        agency_id = re.search(agency_id_pattern,text,re.IGNORECASE).group(1)
        
        agents_data_pattern = r'Writing Agent(.*?)Total Commission'
        agents_data = re.findall(agents_data_pattern,text,re.DOTALL|re.IGNORECASE)
        
        agents_pattern = r'^(\w+) (\w+) ([a-z-, ]+) (\d) (\d+ )?(\d+\/\d+ )?([0-9-]+) ([0-9.,-]+) ([0-9.,-]+) (\w+) ([0-9.,-]+)'
        
        for agent_data in agents_data:
            print(agent_data)
            agents = re.findall(agents_pattern,agent_data,re.DOTALL|re.IGNORECASE|re.MULTILINE)
            for agent in agents:
                data.append({
                    "Carrier": carrier,
                    "Agency": agency,
                    "Agency ID": agency_id,
                    "Writing Agent": agent[0],
                    "Policy No": agent[1],
                    "Description": agent[2],
                    "Code": codes[agent[3]],
                    "Dur": agent[4],
                    "Date Due": agent[5],
                    "Mths.": agent[6],
                    "Premium": agent[7],
                    "Rate": agent[8],
                    "Pymnt Term": agent[9],
                    "Commission": agent[10],
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        # print(df)
        return df,output_name 
        
    def sentara_aca(self):
        data = []
        carrier = "Sentara ACA"
        output_name = self.pdf_output_name
        text = self.extract_tables_from_pdf()
        date_pattern = r'^([a-z]+ \d+, \d+) ([a-z]+ \d+, \d+)'
        clients_pattern = r'(\d+) ([a-z, -]+) (\w+) ([a-z, -]+) (\d+) ([0-9.]+) (\$[0-9.]+) (\$[0-9.]+) ([a-z,-]+ [a-z,-]+ [a-z,-]*) ([a-z]+ )?([0-9.]+) (\d+)'
        # print(text)
        agents = []

        for tables in text:
            for table in tables[2:]:
                for rows in table:
                    print(rows)
                    if rows is None or rows == "":
                        continue
                    agents.append([rows])
                    if rows is not None:
                        date = re.search(date_pattern,rows)
                    if date:
                        date1, date2 = date.group(1),date.group(2)
        for agent in agents:
            date = re.search(date_pattern,agent[0],re.IGNORECASE)
            date1, date2 = date.group(1),date.group(2)
            clients = re.findall(clients_pattern,agent[0],re.DOTALL|re.MULTILINE|re.IGNORECASE)
            for client in clients:
                data.append({
                    "Carrier": carrier,
                    "Check Date": date1,
                    "Last Pay Date": date2,
                    "Agency #": client[0],
                    "AGENCY(group)": client[1],
                    "Client #": client[2],
                    "Client": client[3],
                    "Invoice": client[4],
                    "Comm Rate": client[5],
                    "Premium": client[6],
                    "Net": client[7],
                    "Agent": client[8],
                    "Comm Mkt Seg": client[9],
                    "MS Rate": client[10],
                    "Contract Count": client[11],
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        # print(df)
        return df,output_name 
    
    def stevens_matthews(self):
        output_name = self.pdf_output_name
        fmo = "Stevens Matthews"
        text = self.extract_text()
        data = []
        # print(text)
        
        carrier_date_agency_pattern = r'([a-z ]+)\n(\d+\/\d+\/\d+)\n([a-z ]+)'
        carrier_date_agency = re.search(carrier_date_agency_pattern,text,re.IGNORECASE)
        
        carrier,date,agency = carrier_date_agency.group(1),carrier_date_agency.group(2),carrier_date_agency.group(3)
        
        table_pattern = r'comm\n(.*?)Total Commission'
        agents_pattern = r"^([a-z]+-?[a-z]*,? [a-z]+[ a-z]*.*)"
        agent_pattern = r'^([a-z]+-?[a-z]*,? [a-z]+[ a-z]*)$'

        table = re.search(table_pattern,text,re.MULTILINE|re.DOTALL|re.IGNORECASE).group(1)
        agents = re.search(agents_pattern,table,re.MULTILINE|re.DOTALL|re.IGNORECASE).group(1)
        agents = agents.split("\n")
        
        filtered_agents = {}
        for agent in agents:
            if re.match(agent_pattern,agent,re.IGNORECASE) and agent != "FIRST YEAR":
                agent_name = agent
                filtered_agents[agent] = []
                continue
            filtered_agents[agent_name].append(agent)
            
        info_pattern = r'^([a-z ]+) ((?:silver|gold)[0-9 ]+) ([a-z-0-9]+) ([0-9-]+) ([0-9-]+) (\d+\/\d+\/\d+) (\d+) (\$ [0-9 .]+) (\$ [0-9 .]+)$'
        info_pattern2 = r'^([a-z]+-? ?[a-z]+)-? ?.*? (\w+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) (\$[0-9 .]+) (-?\d+) (\$[0-9 .()]+)$'
        for key,values in filtered_agents.items():
            # print(f"Keys are {key}")
            stm_type = ""
            if values:
                for value in values:        
                    info_match = re.match(info_pattern,value,re.IGNORECASE)
                    info_match2 = re.match(info_pattern2,value,re.IGNORECASE)
                    if info_match2:
                        data.append({
                            "FMO": fmo,
                            "Carrier": carrier,
                            "Statement Date": date,
                            "Agency": agency,
                            "Agent": key,
                            "Type": stm_type,
                            "Policy Holder Name": info_match2.group(1),
                            "Plan Name": "AMBETTER",
                            "Policy #": info_match2.group(2),
                            "Status Date": info_match2.group(4),
                            "Month Paid": info_match2.group(3),
                            "Effective Date": info_match2.group(5),
                            "Per Mem": info_match2.group(7),
                            "Comm": info_match2.group(6),
                            "Total Comm": info_match2.group(8),
                        })
                    else:
                        if not info_match:
                            stm_type = value
                            print(stm_type)
                        else:
                            print(info_match.groups)
                            data.append({
                                "FMO": fmo,
                                "Carrier": carrier,
                                "Statement Date": date,
                                "Agency": agency,
                                "Agent": key,
                                "Type": stm_type,
                                "Policy Holder Name": info_match.group(1),
                                "Plan Name": info_match.group(2),
                                "Policy #": info_match.group(3),
                                "Status Date": info_match.group(4),
                                "Month Paid": info_match.group(5),
                                "Effective Date": info_match.group(6),
                                "Per Mem": info_match.group(7),
                                "Comm": info_match.group(8),
                                "Total Comm": info_match.group(9),
                            })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
    
    def united_american(self):
        output_name = self.pdf_output_name
        data = []
        codes = {
            "7": "AGENT ERROR", 
            "A": "REGISTRATION COMMISSION or ANNUITY COMMISSION",
            "B": "MEDICARE PART D", 
            "C": "CANCELLATION",
            "D": "DECLINED", 
            "G": "DEATH CLAIM IN FIRST YEAR",
            "H": "AGENT HIERACHY M Pending Issue",
            "N": "NEW BUSINESS COMMISSION ADJUSTMENT or NEW BUSINESS ADJUSTMENT", 
            "O": "DOES NOT MEET ADVANCE CRITERIA",
            "P": "INCORRECT PREMIUM BASED ON APPLICATION AS SUBMITTED", 
            "R": "REINSTATED or COMMISSIONABLE RENEWAL PREMIUM",
            "S": "STATUS CHANGED", 
            "U": "UNDERWRITING DECLINED ADVANCE",
            "W": "PARTIAL ADVANCE PENDING INITIAL PAYMENT", 
            "X": "CONVERSION",
            "Z": "BALANCE OF ADVANCE - INITIAL PAYMENT",
            "**": "RETURNED CHECK", 
            "V": "PREVIOUSLY ADVANCED",
        }
        text = self.extract_text()
        
        carrier_date_agency_agency_id_pattern = r'\n(.*?)\n.*?\n(\d+\/\d+\/\d+)\nAgent:.*?\n([a-z ]+),?\s?#?(\w+)$'
        carrier,date,agency,agency_id = re.search(carrier_date_agency_agency_id_pattern,text,re.IGNORECASE|re.MULTILINE|re.DOTALL).groups()
        carrier = "Globe Life" if "Globe Life" in carrier else carrier
        tables_pattern = r'account(.*?)balance forward'
        agents_pattern = r'^([a-z ]+)\n(.*?)total'
        # clients_pattern = r'(\w+) (\d+) ([a-z, ]+) (\d+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) ([0-9.-]+) ([0-9.%-]+) ([0-9.-]+)\s?([a-z*]{,2})?'
        clients_pattern = r"(\w+ )?(\d+) ([a-z, ']+) (\d+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) ([0-9.-]+) ([0-9.%-]+) ([0-9.-]+)\s?([a-z*]{,2})?$"
        
        tables = re.findall(tables_pattern,text,re.IGNORECASE|re.MULTILINE|re.DOTALL)
        
        for table in tables:
            agents = re.findall(agents_pattern,table,re.IGNORECASE|re.DOTALL|re.MULTILINE)
            for agent in agents:
                clients = re.findall(clients_pattern,agent[1],re.IGNORECASE|re.MULTILINE|re.DOTALL)
                writing_agent_id = ""
                for client in clients:
                    writing_agent_id = client[0] if client[0] else writing_agent_id
                        
                    data.append({
                        "Carrier": carrier,
                        "Agency": agency,
                        "Agency ID": agency_id,
                        "Date": date,
                        "Type": agent[0],
                        "Writing Agent ID": writing_agent_id,
                        "Policy Number": client[1],
                        "Insured Name": client[2],
                        "Months Paid": client[3],
                        "Due Date": client[4],
                        "Issue Date": client[5],
                        "Premium Amount": client[6],
                        "Comm Rate": client[7],
                        "Comm Amount": client[8],
                        "Reason Code":codes.get(client[9]),
                    })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name 
    
    def jefferson_health(self):
        data = []
        carrier = "JEFFERSON HEALTH"
        fmo = "Pinnacle"
        
        text = self.extract_text()
        date = re.search(r'(\d+\.\d+\.\d+)',text).group(1)
        clients = re.findall(r'^(\d+) (.*?) (\d+) \w+ ([a-z-,]+ ?[a-z]{0,1}) ([a-z-,]+ ?[a-z]+) ([a-z]{2}) (\d+\/\d+\/\d+)( \d+\/\d+\/\d+)? ([0-9 .$]+)',text,re.IGNORECASE|re.MULTILINE|re.DOTALL)
        for client in clients:
            data.append({
                "Carrier": carrier,
                "FMO": fmo,
                "Date": date,
                "Agent NPN": client[0],
                "Agent Name": client[1],
                "Member ID": client[2],
                "Member HICN": "",
                "Member First": client[3],
                "Member Last": client[4],
                "Member State": client[5],
                "Effective Date": client[6],
                "Cancel Date": client[7],
                "Commission": client[8],
            })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name 
    
    def health_first_fl(self):
        data = []
        carrier = "Health First FL"
        text = self.extract_text()
        agency,agency_id = re.search(r'detail\n(.*?)\n(\d+)',text,re.IGNORECASE).groups()
        tables = re.findall(r'received amount\n(\w+)(.+)total',text,re.IGNORECASE|re.MULTILINE|re.DOTALL)
        categories_pattern = r'^([a-z0-9 ]+)\n([a-z0-9 ]+)\n(.*?)total'
        manual_adjustments_pattern = r'manual adjustments(.*?)total'
        clients_pattern = r'^([a-z, ]+) (\d+) ([a-z, ]+) (\d+\/\d+) ([0-9 \-,.$]+) ([0-9 \-,.$]+) ([0-9]+) ([0-9 \-,.$]+) (\w+) ([0-9 \-.$]+)'
        manual_clients_pattern = r'([a-zA-Z, ]+) (\d+) ([A-Z ]+) ([a-zA-Z ]+) (\d+\/\d+) ([0-9 $.\-]+)'
        for table in tables:
            categories = re.findall(categories_pattern,table[1],re.IGNORECASE|re.DOTALL|re.MULTILINE)
            manual_adjustments = re.findall(manual_adjustments_pattern,table[1],re.IGNORECASE|re.MULTILINE|re.DOTALL)
            for category in categories:
                clients = re.findall(clients_pattern,category[2],re.DOTALL|re.IGNORECASE|re.MULTILINE)
                for client in clients:
                    data.append({
                        "Carrier": carrier,
                        "Agency": agency,
                        "Agency ID": agency_id,
                        "Type": table[0],
                        "Category":category[0],
                        "Description": category[1],
                        "Broker": client[0],
                        "Contract ID": client[1],
                        "Name": client[2],
                        "Prem Period": client[3],
                        "Premium Invoiced": client[4],
                        "Premium Received": client[5],
                        "Count": client[6],
                        "Rate": client[7],
                        "Commission Type": client[8],
                        "Commission Amount": client[9],
                    })
                    # print(client)
            if manual_adjustments:
                for manual in manual_adjustments:
                    manual_clients = re.findall(manual_clients_pattern,manual,re.DOTALL|re.MULTILINE)
                    for manual_client in manual_clients:
                        data.append({
                            "Carrier": carrier,
                            "Agency": agency,
                            "Agency ID": agency_id,
                            "Type": "",
                            "Category": "",
                            "Description": "",
                            "Broker": manual_client[0],
                            "Contract ID": manual_client[1],
                            "Name": manual_client[2],
                            "Prem Period": manual_client[4],
                            "Premium Invoiced": "",
                            "Premium Received": "",
                            "Count": "",
                            "Rate": "",
                            "Commission Type": manual_client[3],
                            "Commission Amount": manual_client[5],
                        })
                        

        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name 
    
    def inshore(self):
        carrier = "Inshore"
        data = []
        text = self.extract_text()
        date = re.search(r'statement date: (\d+\/\d+\/\d+)',text,re.IGNORECASE).group(1)
        agent_no = re.search(r'agent #: (\d+)',text,re.IGNORECASE).group(1)
        agent_name = re.search(r'agent name: ([a-z \.]+)',text,re.IGNORECASE).group(1)
        
        clients_pattern = r'([a-z &]+) (\w+) ([a-z ]+) (\d+\/\d+\/\d+) ([0-9\.\-$]+) (\d+%) ([0-9\.\-$]+)'
        clients = re.findall(clients_pattern,text,re.MULTILINE|re.IGNORECASE|re.DOTALL)
        for client in clients:
            data.append({
                "Carrier": carrier,
                "Statement Date": date,
                "Agent #": agent_no,
                "Agent Name": agent_name,
                "Client Name": client[0],
                "Billing Group Number": client[1],
                "Line of Business": client[2],
                "Coverage Month": client[3],
                "Premium Received": client[4],
                "Comm Rate %": client[5],
                "Commission Earned": client[6],
            })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name 
    
    def nippon_life(self):
        carrier = "Nippon Life"
        data = []
        text = self.extract_text()

        statements_pattern = r'pct amount(.*?)statement total'
        date = re.search(r'period ending (\d+\/\d+\/\d+)',text,re.IGNORECASE).group(1)
        producer_number = re.search(r'producer number ([0-9-]+)',text,re.IGNORECASE).group(1)
        agents_pattern = r'([a-z -]+)(.*?)subtotal'
        clients_pattern = r'(\w+)( [a-z\. ]+)? ([dentalifvsomdcby]+) ([0-9\.]+) (\d+\/\d+) (\d+%)( \d+)? (\d+%) ([0-9\.]+)'
        
        statements = re.findall(statements_pattern,text,re.MULTILINE|re.DOTALL|re.IGNORECASE)
        client_name = ""
        for statement in statements:
            agents = re.findall(agents_pattern,statement,re.DOTALL|re.IGNORECASE|re.MULTILINE)
            for agent in agents:
                clients = re.findall(clients_pattern,agent[1],re.IGNORECASE|re.DOTALL|re.MULTILINE)
                for client in clients:
                    name = client[1]
                    client_name = name if name else client_name
                    data.append({
                        "Carrier": carrier,
                        "For Period Ending": date,
                        "Producer Number": producer_number,
                        "Client No./Unit No.": client[0],
                        "Servicing Agent/Client Name": agent[0],
                        "Client Name": client_name,
                        "Line of Coverage": client[2],
                        "Paid Amount": client[3],
                        "Due for Period Ending": client[4],
                        "Comm Rate": client[5],
                        "Premium Scale Level": client[6],
                        "Split Pct": client[7],
                        "Commission Amount": client[8],
                    })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name 
    
    def kaiser_georgia(self):
        carrier = "Kaiser Permanente Georgia"
        data = []
        text = self.extract_text()
        # print(text)
     
        vendor_vendor_id_agency_date_pattern = r'vendor # (\w+).*?vendor id (\d+)\n(.*?)commission month: ([0-9 -\/]+)'
        agent_clients_pattern = r'mpf: ([a-z, ]+)(.*?)member count'
        clients_pattern = r'^([a-z]+,[a-z]+,[a-z]+|[a-z]+,[a-z]+|[a-z]+ ?-?[a-z]+,[a-z]+)( [a-z]{,2})?( \d+\/\d+\/\d+)?( \d+\/\d+\/\d+)?( [0-9\$\.,]+)?( [0-9\$\.,]+)?( [0-9\$\.,]+)'
        
        vendor_vendor_id_agency_date = re.search(vendor_vendor_id_agency_date_pattern,text,re.IGNORECASE|re.MULTILINE|re.DOTALL)
        vendor = vendor_vendor_id_agency_date.group(1)
        vendor_id = vendor_vendor_id_agency_date.group(2)
        agency = vendor_vendor_id_agency_date.group(3)
        date = vendor_vendor_id_agency_date.group(4)
 
        agent_tables = re.findall(agent_clients_pattern,text,re.IGNORECASE|re.MULTILINE|re.DOTALL)
        for table in agent_tables:
            agent = table[0]
            clients = re.findall(clients_pattern,table[1],re.IGNORECASE|re.MULTILINE|re.DOTALL)
            plan = ""
            eff_date = ""
            payment_trans = ""
            amount_received = ""
            comm_rate = ""
            
            for client in clients:   
                if client[1] != "":
                    plan = client[1]
                    eff_date = client[2]
                    payment_trans = client[3]
                    amount_received = client[4]
                    comm_rate = client[5]
                data.append({
                    "Carrier": carrier,
                    "Agency": agency,
                    "Vendor #": vendor,
                    "AP Vendor ID": vendor_id,
                    "Commission Month": date,
                    "Agent Name": agent,
                    "Member Name":client[0],
                    "Plan":plan,
                    "Member Effective Date": eff_date,
                    "Date of Last Payment Trans.": payment_trans,
                    "Amount Received": amount_received,
                    "Comm Rate": comm_rate,
                    "Commission": client[6],
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name 
    
    def cigna_ms_lisa(self,columns):
        carrier = "Cigna Supplemental"
        data = []
        text = self.clean_lines_main(column_ranges=columns,y_tolerance=6)
        print(f"Text is {text}")
        first_page = self.extract_text_from_range(start_page=0,end_page=1)
        run_date = re.search(r'run date *(\w+ \d+, \d+)',first_page,re.IGNORECASE)
        r_date = run_date.group(1)
        agent_period_ending = re.search(r'(.*?) *period ending *(\w+ \d+, \d+)',first_page,re.IGNORECASE)
        agent,period_ending = agent_period_ending.group(1),agent_period_ending.group(2)
        
        insured_info_pattern = r'^(\w+), ([a-z ,-]+), (\w+), ([0-9\/]+), ([0-9\.]+), ([0-9\.]+), ([0-9\.]+), ([0-9\.]+)?, *([0-9\.]+)?, *([0-9\.]+)?'
        insured_info = re.findall(insured_info_pattern,text,re.IGNORECASE|re.MULTILINE)

        for insured in insured_info:
            
            data.append({
                "Carrier": carrier,
                "Earning Agent Name": agent,
                "Run Date": r_date,
                "Period Ending": period_ending,
                "Policy": insured[0],
                "Insured's Name": insured[1],
                "Plan Code": insured[2],
                "Paid To": insured[3],
                "Premium": insured[4],
                "Per Cent": insured[5],
                "Earned": insured[6],
                "Amt to Pay": insured[7],
                "FICA": insured[8],
                "Apply to Adv": insured[9],
            })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name 
    
    def bcbs_lousiana(self,carrier):
        output_name = self.pdf_output_name
        data = []
        text = self.extract_text()

        tables = self.extract_tables_from_pdf()



        date = re.search(r"Activity Ending Date: (\d+\/\d+\/\d+)",text,re.DOTALL)
        date = date.group(1)

        statement_type_pattern = r'\w+, \w+, \d+ ([a-z ]+)'
        statement_type1 = re.search(statement_type_pattern,text,re.DOTALL|re.MULTILINE|re.IGNORECASE)
        
        producers_pattern = r'Writing Producer (\d+)\s(\w+\s\w+)\n(.*?)Total for Writing Producer'
        producers = re.findall(producers_pattern,text,re.MULTILINE|re.DOTALL)
        
        statements_pattern = r'^([a-z ]+)(.*?)total for'
        
        clients_pattern = r'^(\d+)\s([a-zA-Z]+(?:\s[a-zA-Z-]+){0,3})\s(\w+)\s(\d+(?:\s[a-zA-Z]+){0,4})\s(\w+)\s(\w+\s)?(\d+\/\d+\/\d+)\s(\d+\/\d+\/\d+\s)?(\d+\/\d+\/\d+)\s(\d+\/\d+)\s(\d+\s)?(\w+)\s(\(?\$ (?:\d+,)?\d+\.\d+\)?)$'
        
        for table in tables:
            for t in table:
                if len(t) == 9 and t[-1] and t[-1].startswith('$'):
                    data.append({
                        "Carrier": carrier,
                        "Statement Type": "Manual Adjustment",
                        "Date": date,
                        "Type": t[6],
                        "Writing Producer": t[1],
                        "Writing ID": t[0],
                        "Member ID": t[2],
                        "Name": t[3],
                        "Company": "",
                        "Product": "",
                        "HICN": t[4],
                        "Override": "",
                        "Effective Date": "",
                        "Term Date": "",
                        "Signed Date": "",
                        "Period": t[5],
                        "Cycle Year": "",
                        "Retro": "",
                        "Description":t[7],
                        "Amount": t[8], 
                    })
        
        for producer in producers:
            statements = re.findall(statements_pattern,producer[2],re.MULTILINE|re.IGNORECASE|re.DOTALL)
            writing_producer = producer[1]
            writing_id = producer[0]

            for statement in statements:
                statement_type = statement[0]

                clients = re.findall(clients_pattern,statement[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
                for client in clients:
                    data.append({
                        "Carrier": carrier,
                        "Statement Type": statement_type,
                        "Date": date,
                        "Type": statement_type,
                        "Writing Producer": writing_producer,
                        "Writing ID": writing_id,
                        "Member ID": client[0],
                        "Name": client[1],
                        "Company": client[2],
                        "Product": client[3],
                        "HICN": client[4],
                        "Override": client[5],
                        "Effective Date": client[6],
                        "Term Date": client[7],
                        "Signed Date": client[8],
                        "Period": client[9],
                        "Cycle Year": client[10],
                        "Retro": client[11],
                        "Description":"",
                        "Amount": client[12], 
                    })

        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, output_name
   
    def martins_point(self):
        data = []
        carrier = "Martin's Point"
        text = self.extract_text(password="mphc4apps")
        
        agent_date_pay_to_pattern = r'agent: (.*?) invoice date: ([0-9\/]+)\npay to: (.*?)\n'
        agent_date_pay_to = re.search(agent_date_pay_to_pattern,text,re.MULTILINE|re.IGNORECASE|re.DOTALL)
        
        clients_tables_pattern = r'^([a-z ]+)\nmember(.*?)total'
        clients_tables = re.findall(clients_tables_pattern,text,re.MULTILINE|re.DOTALL|re.IGNORECASE)
        
        clients_pattern = r'^([a-z -]+) (\d+) (\w+) ([a-z -]+)?([0-9\/]+) ([0-9\/]+ )?([0-9\.\$-]+)'
        
        agent,date,pay_to = agent_date_pay_to.groups()
        
        for tables in clients_tables:
            transaction_type = tables[0]
            clients = re.findall(clients_pattern,tables[1],re.MULTILINE|re.DOTALL|re.IGNORECASE)
            for client in clients:
                data.append({
                   "Carrier": carrier,
                   "Agent": agent,
                   "Pay To": pay_to,
                   "Invoice Date": date,
                   "Transaction Type": transaction_type,
                   "Member Name": client[0],
                   "Member ID": client[1],
                   "Status": client[2],
                   "Selling Agent": client[3],
                   "Effective Date": client[4],
                   "Term Date": client[5],
                   "Amount": client[6] 
                })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
    
    def carington(self):
        data = []
        carrier = "Careington"
        text = self.extract_text()

        
        agency_agency_number_pattern = r'([a-z ,\.]+) --- (\w+)'
        agency,agency_number = re.search(agency_agency_number_pattern,text,re.IGNORECASE).groups()
        
        period_ending_pattern = r'eom date: ([0-9\/]+)'
        period_ending = re.search(period_ending_pattern,text,re.IGNORECASE).group(1)
        
        direct_deposit_pattern = r'ref# ([0-9 ]+)'
        direct_deposit = re.search(direct_deposit_pattern,text,re.IGNORECASE).group(1)
        
        enrollments_pattern = r'group : (\w+) ([a-z &]+) agent type : (\w+)(.*?)company'
        enrollments = re.findall(enrollments_pattern,text,re.DOTALL|re.IGNORECASE|re.MULTILINE)
        enrollment_clients_pattern = r'([a-z, ]+) ([0-9a-z,]+) efft : (\d+\/\d+) \w+ pmtdt: (\d+\/\d+\/\d+) cvrg: (\w+) (\w+) prem: ?(\$[0-9\.]+) ([0-9\.]+%) (\$[0-9-\.]+)'
        
        
        for enrollment in enrollments:
            group_number = enrollment[0]
            group_name = enrollment[1]
            agent_type = enrollment[2]
            
            lines = enrollment[3].splitlines()

            # Group them
            groups = []
            current_group = []

            for line in lines:
                if "Enrollment Renewals" in line:
                    # Start new group
                    if current_group:
                        groups.append(current_group)
                    current_group = [line]
                else:
                    current_group.append(line)

            # Add last group
            if current_group:
                groups.append(current_group)
            
            for group in groups:
                if len(group) > 1:
                    category = group[0]
                    group_to_str = "\n".join(group)
                    clients = re.findall(enrollment_clients_pattern,group_to_str,re.IGNORECASE|re.DOTALL|re.MULTILINE)
                    for client in clients:
                        client_id = client[1][-6:]
                        data.append({
                            "Carrier" : carrier,
                            "Period Ending" : period_ending,
                            "Direct Deposit": direct_deposit,
                            "Agency" : agency,
                            "Agency Number" : agency_number,
                            "Group Number" : group_number,
                            "Group Name" : group_name,
                            "Agent Type" : agent_type,
                            "Category" : category,
                            "Client Name" : client[0],
                            "Client ID" : client_id,
                            "Effective Date" : client[2],
                            "Pmt Dat" : client[3],
                            "Cvrg" : client[4],
                            "Mode" : client[5],
                            "Prem" : client[6],
                            "Rate" : client[7],
                            "Commission" : client[8],
                        })

        
        not_paid_pattern = r'group : (\w+) ([a-z &]+) agent type : (\w+)\n(not paid this period)(.*)'
        not_paid = re.findall(not_paid_pattern,text,re.DOTALL|re.IGNORECASE|re.MULTILINE)
        not_paid_clients_pattern = r'([a-z, ]+) ([0-9a-z,]+) efft : (\d+\/\d+) \w+ cvrg: (\w+) (\w+) prem: (\$[0-9\.]+)'
        
        for n_paid in not_paid:
            group_number = n_paid[0]
            group_name = n_paid[1]
            agent_type = n_paid[2]
            category = n_paid[3]

            not_paid_clients = re.findall(not_paid_clients_pattern,n_paid[4],re.DOTALL|re.IGNORECASE|re.MULTILINE)
            for n_p_client in not_paid_clients:
                client_id = n_p_client[1][-6:]
                data.append({
                    "Carrier" : carrier,
                    "Period Ending" : period_ending,
                    "Direct Deposit": direct_deposit,
                    "Agency" : agency,
                    "Agency Number" : agency_number,
                    "Group Number" : group_number,
                    "Group Name" : group_name,
                    "Agent Type" : agent_type,
                    "Category" : category,
                    "Client Name" : n_p_client[0],
                    "Client ID" : client_id,
                    "Effective Date" : n_p_client[2],
                    "Pmt Dat" : "",
                    "Cvrg" : n_p_client[3],
                    "Mode" : n_p_client[4],
                    "Prem" : n_p_client[5],
                    "Rate" : "",
                    "Commission" : "",
                })
        
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
    
    def delta_dental_northeast(self):
        carrier = "Delta Dental Northeast"
        data = []
        text = self.extract_text()
        
        tables_pattern = r'check number(.*?)\n\$'
        headers_pattern = r'([a-z0-9\.]+) ([a-z]+ [a-z]+(?: [a-z]+)?) ([a-z]+ [a-z]+(?: [a-z]+)?) (\d+) ([0-9\/]+) (\w+)'
        clients_pattern = r'(\d+) ([0-9a-z ]+) (\d+\/\d+\/\d+) (\$[0-9\.-]+) (\$[0-9\.-]+)'
        
        tables = re.findall(tables_pattern,text,re.IGNORECASE|re.DOTALL)
        
        for table in tables:
            headers = re.search(headers_pattern,table,re.IGNORECASE).groups()
            clients = re.findall(clients_pattern,table,re.IGNORECASE)
            for client in clients:
                data.append({
                    "Carrier" : carrier,
                    "Vendor ID" : headers[0],
                    "Vendor Name" : headers[1],
                    "Check Name" : headers[2],
                    "Payment Number" : headers[3],
                    "Check Date" : headers[4],
                    "Check Number" : headers[5],
                    "Invoice #" : client[0],
                    "Description" : client[1],
                    "Date" : client[2],
                    "Orig Amnt" : client[3],
                    "Amount Paid" : client[4],
                })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
    
    def general_agent_center(self,column_ranges):
        data = []
        text = self.extract_text(0,1)
        
        plan_types = {
            "VBA-SP" : "Value Benefits of America",
            "VBA-30W" : "Value Benefits of America",
            "NEA-AC" : "NEA-ACCIDENT"
        }

        filtered_text = self.clean_lines_main(column_ranges=column_ranges)
        
        agency_information_pattern = r'([a-z ]+) commission statement ([a-z ]+)\n.*?producer ([0-9a-z-]+).*?period beginning: (\d+\/\d+\/\d+).*?period ending: (\d+\/\d+\/\d+)'
        clients_pattern = r'^([a-z0-9 ]+), (\w+), (\d+), (\d+), (\d+), (\d+), ([a-z, ]+), (\d+), ([a-z0-9-]+), (\d+). ([0-9\.-]+), (\d+), ([0-9a-z\.]+),(.*?),(.*?), ([0-9a-z\.]+)$'
        
        agency_information = re.search(agency_information_pattern,text,re.IGNORECASE|re.DOTALL).groups()
        
        clients = re.findall(clients_pattern,filtered_text,re.DOTALL|re.IGNORECASE|re.MULTILINE)
        
        for client in clients:
            data.append({
                "FMO" : agency_information[0],
                "Carrier" : plan_types.get(client[8],""),
                "Agency" : agency_information[1],
                "Producer ID" : agency_information[2],
                "Period Begginning" : agency_information[3],
                "Period Ending" : agency_information[4],
                "Writing Agent" : client[0],
                "Source Code" : client[1],
                "Tran Date" : client[2],
                "Effect Date" : client[3],
                "Paid From" : client[4],
                "Paid To" : client[5],
                "Name Insured" : client[6],
                "Policy Number" : client[7],
                "Plan Type" : client[8],
                "Curr Mode" : client[9],
                "Premium Collected" : client[10],
                "Fee" : client[11],
                "$ Rate" : client[12],
                "Commission" : client[13],
                "Commissions Retained" : client[14],
                "Misc" : client[15],
            })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name

    def bcbs_sc(self,column_ranges,column_ranges_two):
        data = []
        carrier = "BCBS SC ACA"
  
        filtered_agency_text = self.clean_lines_main(column_ranges)
        agency_pattern = r'^agency, (\w+), ([a-z ,]+)$'
        
        agency = re.search(agency_pattern,filtered_agency_text,re.IGNORECASE|re.MULTILINE).groups()
        
        filtered_text = self.clean_lines_main(column_ranges_two,y_tolerance=6)
        
        clients_pattern = r'([a-z -]+), (\w+), ([a-z ]+), ([0-9\/]+), (\w+), ([0-9\/]+), ([0-9\.,]+), (\$ [0-9\.,]+), ([0-9\.]+),(.*?), (\w+)'
        clients_pattern2 = r'([a-z \-\']+), (\w+), ([a-z ]+), ([0-9\/]+), ?(\w+)?, ([0-9\/]+), (\w+), (\d+), ([0-9\.,]+), (\$ [0-9\.,]+), ([0-9\.]+),(.*?), (\w+)'
        
        clients = re.findall(clients_pattern,filtered_text,re.IGNORECASE|re.DOTALL|re.MULTILINE)
        if not clients:
            clients = re.findall(clients_pattern2,filtered_text,re.IGNORECASE|re.DOTALL|re.MULTILINE)
            

        for client in clients:
            print(len(client))
            data.append({
                "Carrier" : carrier,
                "Agency Name" : agency[1],
                "Agency ID" : agency[0],
                "Subscriber Name" : client[0],
                "Alternate ID" : client[1],
                "Contract Type" : client[2],
                "Eff Date" : client[3],
                "TC" : client[4],
                "Due Date" : client[5],
                "Initial Premium" : client[8] if len(client) > 11 else client[6],
                "Percap Percent" : client[9] if len(client) > 11 else client[7],
                "Commission Amount" : client[10] if len(client) > 11 else client[8],
                "Adj Rea" : client[11] if len(client) > 11 else client[9],
                "Selling Agent" : client[12] if len(client) > 11 else client[10],
            })

        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
    
    def cigna_global(self):
        data = []
        carrier = "Cigna Global"
        text = self.extract_text_from_range(0)
        
        broker_info_pattern = r'commission number:\n(\d+)\naccount name:\n([a-z ,0-9.]+)\nbroker reference number:\n([a-z0-9 ]+)\nstatement date:\n([a-z0-9 ]+)'
        broker_info = re.search(broker_info_pattern,text,re.IGNORECASE)
        commission_number,account_name,broker_reference_number,statement_date = broker_info.groups()
        
        policy_holders_pattern = r'(\d+)\n([a-z ]+\n[a-z ]+\n)(\d+)\n(\d+ \D+ \d+) (\d+ \D+ \d+)\n([a-z0-9. ]+)\n(\d+)\n([a-z0-9 .]+)'
        policy_holders_pattern_1 = r'(\d+)\n([a-z ]+\n)(\d+)\n(\d+ \D+ \d+) (\d+ \D+ \d+)\n([a-z0-9. ]+)\n(\d+)\n([a-z0-9 .]+)'
       
        
        policy_holders = re.findall(policy_holders_pattern,text,re.IGNORECASE|re.MULTILINE|re.DOTALL)
        policy_holders_1 = re.findall(policy_holders_pattern_1,text,re.IGNORECASE|re.MULTILINE|re.DOTALL)

        if policy_holders:
            for policy_holder in policy_holders:
                data.append({
                    "Carrier": carrier,
                    "Commission Number": commission_number,
                    "Account Name": account_name,
                    "Broker Reference Number": broker_reference_number,
                    "Statement Date": statement_date,
                    "Policy Number": policy_holder[0],
                    "Policy Holder": " ".join(policy_holder[1].split("\n")),
                    "Transaction No": policy_holder[2],
                    "Premium Due Date": policy_holder[3],
                    "Premium Paid Date": policy_holder[4],
                    "Premium Paid (Inc Tax)": policy_holder[5],
                    "Commission Percent": policy_holder[6],
                    "Commission Amount": policy_holder[7]
                })
        if policy_holders_1:
            for policy_holder1 in policy_holders_1:
                data.append({
                    "Carrier": carrier,
                    "Commission Number": commission_number,
                    "Account Name": account_name,
                    "Broker Reference Number": broker_reference_number,
                    "Statement Date": statement_date,
                    "Policy Number": policy_holder1[0],
                    "Policy Holder": policy_holder1[1],
                    "Transaction No": policy_holder1[2],
                    "Premium Due Date": policy_holder1[3],
                    "Premium Paid Date": policy_holder1[4],
                    "Premium Paid (Inc Tax)": policy_holder1[5],
                    "Commission Percent": policy_holder1[6],
                    "Commission Amount": policy_holder1[7]
                })
            
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
                    
        

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
