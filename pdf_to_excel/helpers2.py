from .helpers import PDFEditor
import re
import pandas as pd

class PDFS(PDFEditor):
    def __init__(self, pdf_file):
        super().__init__(pdf_file)
        
    def BCBS_Nebraska(self):
        text = self.extract_text()
        
        data = []
        carrier = "BCBS NE"
        
        agency_name_id_pattern = r'agency id# (\w+)\n([a-z0-9 ,\.]+)'
        agency_id, agency_name = re.search(agency_name_id_pattern,text,re.DOTALL|re.IGNORECASE).groups()
        
        clients_info_pattern = r'([a-z0-9-]+)\nbrok(.*?)total for'
        comm_type = re.findall(clients_info_pattern,text,re.DOTALL|re.IGNORECASE)
        
        broker_pattern = r'^er (\w+) ([a-z, ]+)(.*)'
        broker2_pattern = r'(\w+) ([a-z ,]+) \d+ \d+ ([0-9]{4}) (\d+%) [0-9\$\.]+ [0-9\$\.]+ [0-9\$\.]+ [0-9\$\.]+ ([a-z ]+) ([a-z0-9]+) ([0-9\$\.-]+)'
        
        clients_pattern = r'(\d+) ([a-z ]+) ([a-z0-9]+) ([a-z0 ]+) ([a-z]{1}) (\d+\/\d+\/\d+) (\d+) (\w+ )?(\d+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) ([0-9\$\.-]+) (\d+) (\w+) ([0-9\$\.-]+) ([0-9\$\.-]+)'
        
        
        for comm_info in comm_type:
            comm_type = comm_info[0]
            brokers = re.findall(broker_pattern,comm_info[1],re.IGNORECASE|re.DOTALL)
            brokers2 = re.findall(broker2_pattern,comm_info[1],re.IGNORECASE|re.DOTALL)
            
            if brokers:
                for broker in brokers:
                    clients = re.findall(clients_pattern,broker[2],re.IGNORECASE|re.DOTALL)
                    for client in clients:
                        data.append({
                            "Carrier":carrier,
                            "Agency ID": agency_id,
                            "Agency Name": agency_name,
                            "Type": comm_type,
                            "Broker Name" : broker[1],
                            "Broker ID": broker[0],
                            "Customer ID": client[0],
                            "Customer Name": client[1],
                            "Prod Type": client[2],
                            "Coverage Type": client[3],
                            "First Year": client[4],
                            "Eff Date": client[5],
                            "Age at Eff": client[6],
                            "Dissability": client[7],
                            "Member Months": client[8],
                            "Bill Eff Date": client[9],
                            "Paid Thru Date": client[10],
                            "Prem Billed": client[11],
                            "Member Count": client[12],
                            "Comm Type": client[13],
                            "Comm Rate": client[14],
                            "Comm Paid": client[15],
                            "Bonus Year": "",
                            "Agency %": "",
                        })
            if brokers2:
                for broker2 in brokers2:
                    data.append({
                            "Carrier":carrier,
                            "Agency ID": agency_id,
                            "Agency Name": agency_name,
                            "Type": comm_type,
                            "Broker Name" : broker2[1],
                            "Broker ID": broker2[0],
                            "Customer ID": "",
                            "Customer Name": "",
                            "Prod Type": broker2[5],
                            "Coverage Type": broker2[4],
                            "First Year":"",
                            "Eff Date":"",
                            "Age at Eff":"",
                            "Dissability":"",
                            "Member Months":"",
                            "Bill Eff Date":"",
                            "Paid Thru Date": "",
                            "Prem Billed": "",
                            "Member Count": "",
                            "Comm Type": "",
                            "Comm Rate": "",
                            "Comm Paid": broker2[6],
                            "Bonus Year": broker2[2],
                            "Agency %": broker2[3],
                        })
        
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
    
    def kaiser_permanente_northwest(self):
        output_name = self.pdf_output_name
        text = self.extract_text()
        data = []
        
        print(text)
        
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
        
        kpif_pattern = r'KPIF\n(.*?)KPIF Members'
        kpif_table = re.findall(kpif_pattern,text,re.DOTALL|re.MULTILINE)
        
        agent_pattern = r'([a-zA-Z0-9, ]+) \((\w+)\)'
        subscribers_list_pattern = r'(\d+) ([a-zA-Z0-9 &-]+) ([a-zA-Z]{2,4}) ([0-9]{,4}) (\$\d+.\d+) ([0-9$.,]+) (\d+%) (\$\d?,?\d+.\d+)\n([a-zA-Z ]+)?'
        for subscribers in subscribers_table:
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
                
        blocks = []        
                
        for subscribers in kpif_table: 
            subscribers_list_pattern = r'([a-zA-Z0-9, ]+ \(\w+\).+)'

            subscribers_list = re.search(subscribers_list_pattern, subscribers,re.DOTALL|re.MULTILINE).group(1)

            header_pattern = re.compile(r"^[A-Za-z]+,[A-Za-z ]+ \(\w+\)$")
         
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
        kpif_member_pattern = r'([a-z, ]+) ([a-z]{2}) (\d+\/\d+\/\d+ )?(\$\d?,?\d+.\d+ )?(\d?,?\d+.\d+ % )?(\$\d?,?\d+.\d+ )?(\$\d?,?\d+.\d+) ([0-9\$\.]+) ([0-9\. %]+) ([0-9\$\.]+) ([0-9\$\.]+) (\d+\/\d+\/\d+) ([0-9\$\.]+)'

        for table in blocks:

            agent, agent_id = re.search(agent_pattern,table,re.DOTALL|re.MULTILINE).groups()

            subscribers = re.findall(kpif_member_pattern,table,re.DOTALL|re.MULTILINE|re.IGNORECASE)

            for sub in subscribers:

                data.append({
                    "Carrier": "Kaiser Permanente Northwest",
                    "Agency": agency,
                    "Vendor #": vendor,
                    "Commission Month": date,
                    "Writing Agent": agent,
                    "Writing ID": agent_id,
                    "Client Name": sub[0],
                    "St": sub[1],
                    "Payment Receivd Date": sub[2],
                    "Medicare Amount Received": sub[3],
                    "% Med Dues Paid": sub[4],
                    "Med Comm Rate": sub[5],
                    "Medicall Comm": sub[6],
                    "Dental Amount Received": sub[7],
                    "% Dent Dues paid ": sub[8],
                    "Dental Comm Rate": sub[9],
                    "Dental Comm": sub[10],
                    "Commission Last Paid": sub[11],
                    "Total Commission": sub[12],
                })
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,output_name
            
    def pivot_health(self,column_ranges):
        data = []
        carrier = "Pivot Health"
        
        text = self.extract_text(start=0,pages=2)
        print(text)

        filtered_text = self.clean_lines_main(column_ranges,y_tolerance=11)
        
        agency_info_pattern = r'\d+\n([a-z0-9 ]+)\n.*?statement date ([0-9-]+)\n.*?statement (\d+)\n.*?agent id (\d+)'
        clients_pattern = r'([0-9-]+), (\w+), ([a-z ,]+), ([0-9-]+), ([0-9-]+), ([0-9\.-]+), ([0-9\.-]+), ([a-z, ]+)'
        
        agency, statement_date, statement, agent_id = re.search(agency_info_pattern,text,re.IGNORECASE|re.DOTALL).groups()
        
        clients = re.findall(clients_pattern,filtered_text,re.IGNORECASE|re.MULTILINE)
        for client in clients:
            data.append({
                "Carrier": carrier,
                "Agency": agency,
                "Statement Date": statement_date,
                "Statement": statement,
                "Agent ID": agent_id,
                "Case Number": client[0],
                "External ID": client[1],
                "Description": client[2],
                "Coverage Month": client[3],
                "Lives": client[4],
                "Paid Premium": client[5],
                "Compensation": client[6],
                "Writing Agent": client[7],
            })
        
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name
    
    def sons_of_norway(self):
        text = self.extract_text(start=0,pages=1)
        data = []
        carrier = "Sons of Norway"
        
        period_ending_pattern = r'period ending - ([0-9\/]+)'
        period_ending = re.search(period_ending_pattern,text,re.IGNORECASE).group(1)
            
        column_ranges = [
                        (0,38),
                        (38,115),
                        (116,154),
                        (155,192),
                        (194,230),
                        (256,293),
                        (293,322),
                        (367,379),
                        (390,436),
                        (436,520),
                        (530,557),
                        (570,600),
                    ]
            
        filtered = self.clean_lines_main(column_ranges=column_ranges,y_tolerance=6)
        
                
        agents_pattern = r', (\d+), ([a-z ,]+),(.*?)totals'
        agents = re.findall(agents_pattern,filtered,re.IGNORECASE|re.DOTALL|re.MULTILINE)
        
        clients_pattern = r'(\d+), ([a-z -,]+), ([0-9\/]+), ([0-9\/]+), ([0-9\/]+), ([0-9\.-]+), ([0-9\.-]+), ([a-z]{1}),(.*?),(.*?),(.*?),( \d+)?'
        for agent in agents:
            clients = re.findall(clients_pattern,agent[2],re.IGNORECASE|re.MULTILINE|re.DOTALL)
            for client in clients:
                data.append({
                    "Carrier": carrier,
                    "Period Ending": period_ending,
                    "Rep#": agent[0],
                    "Agent": agent[1],
                    "Cert#": client[0],
                    "Holder": client[1],
                    "Effect Date": client[2],
                    "Paid From": client[3],
                    "Paid To": client[4],
                    "Premium Paid": client[5],
                    "Comm FCTR": client[6],
                    "SRC Tea": client[7],
                    "1st Comm": client[8],
                    "1st Excs": client[9],
                    "Ren Comm": client[10],
                    "Ref Rep": client[11],
                })
        
        if len(data) >= 1:
            for i in range(1):
                data[i]["Converted from .pdf by"] = ""
        df = pd.DataFrame(data)
        return df,self.pdf_output_name
    
    def royal_neighbors(self):
        """Process the first type of PDF (with 'Run Date' and 'Agents')."""
        output_name = self.pdf_output_name

        # Extract text from page 4
        text = self.extract_text()


        # Extract 'Run Date' and 'Agents' information
        agents = re.findall(r"(.*?)Subtotals for Agent (\w+)\s+([A-Z,.\s]+)", text,re.IGNORECASE|re.DOTALL)
        run_date = re.findall(r"Run Date:\s(\d{,2}\/\d{,2}\/\d{4})", text)
        clients_pattern = r'([a-z, -]+ )?([0-9]+)( \w+)? ([0-9]{2}\/[0-9]{2}\/[0-9]{4}) ([a-z]+) ([0-9]{2}\/[0-9]{2}\/[0-9]{4}) ([a-z]+) ([0-9\.]+) ([0-9\$\.\-]+) ([0-9\.]+) ([0-9\$\.\-]+) ([0-9\$\.\-]+) ([0-9\$\.\-\(\)]+)'
        extra_cols_pattern = r'balance\n(.*?)earned'
        extra_rows_pattern = r'([0-9\$\.]+)( [a-z ]+)?'
            
        if not agents:
            return None, None  # Handle invalid case

        new_data = []

        for agent in agents:
            print(agent[0])
            clients = re.findall(clients_pattern,agent[0],re.IGNORECASE|re.DOTALL)

            for client in clients:

                new_data.append({
                    "Run Date": run_date[0],
                    "Carrier": "Royal Neighbors",
                    "Agent Name": agent[2],
                    "Agent ID": agent[1],
                    "Insured's Name": client[0],
                    "Certificate": client[1],
                    "Prod ID": client[2],
                    "Issue Date": client[3],
                    "Mode": client[4],
                    "Paid To Date": client[5],
                    "1st Yr Rnwl": client[6],
                    "Split%": client[7],
                    "Prem": client[8],
                    "Comm%": client[9],
                    "Earned": client[10],
                    "Applied To Advance": client[11],
                    "Amt To Pay": client[12],
                }) 
        extra_cols = re.findall(extra_cols_pattern,text,re.IGNORECASE|re.DOTALL)

        for cols in extra_cols:
            rows = re.findall(extra_rows_pattern,cols,re.IGNORECASE)
            for d, e in zip(new_data, rows):
                d["Cert Adv Balance"] = e[0]
                d["Comment"] = e[1]

        # for d, e in zip(new_data, extra_cols):
        #     d["Cert Adv Balance"] = e[0]
        #     d["Comment"] = e[1]

        # Add custom column
        if len(new_data) >= 1:
            for i in range(1):
                new_data[i]["Converted from .pdf by"] = ""

        # Convert to pandas DataFrame
        df = pd.DataFrame(new_data)
        
        return df, output_name
    
    def providence2(self):
        carrier = "Providence"
        data = []
        text = self.extract_text()

        agency_info_pattern = r'([a-z ,\-]+)\n([a-z ]+)\n[a-z0-9 ]+\n[a-z0-9 ,]+\n([a-z ]+\n)?commission period: (\d+-\d+-\d+)'
        agency_info = re.search(agency_info_pattern,text,re.IGNORECASE)
        
        agency = agency_info.group(1)
        document_type = agency_info.group(2) + " " + agency_info.group(3)
        commission_period = agency_info.group(4)
        
        producers_pattern = r'producer (\d+) ([a-z ,]+)\n(.*?)total producer'
        producers = re.findall(producers_pattern,text,re.IGNORECASE|re.DOTALL)
        
        commission_types_pattern = r'^([a-z ]+)\n(.*?)subtotal'
        clients_pattern = r'(\d+) ([a-z ,]+) ([0-9\$\.,-]+) ([0-9\$\.,-]+) (\d+) (\d+) (\d+\/\d+) ([a-z]+) ([0-9\.]+) ([a-z]+) (\w+ )?([0-9\$\.-]+)'
        
        
        for producer in producers:
            commission_types = re.findall(commission_types_pattern,producer[2],re.IGNORECASE|re.DOTALL|re.MULTILINE)
            for commission in commission_types:

                clients = re.findall(clients_pattern,commission[1],re.IGNORECASE|re.DOTALL)
                for client in clients:
                    data.append({
                        "Carrier": carrier,
                        "Agency": agency,
                        "Document Type": document_type.strip("\n"),
                        "Commission Period": commission_period,
                        "Producer Name": producer[1],
                        "Producer ID": producer[0],
                        "Transaction Type": commission[0],
                        "Group ID": client[0],
                        "Group Name": client[1],
                        "Total Premium": client[2],
                        "Premium Paid": client[3],
                        "Total Members": client[4],
                        "Paid Members": client[5],
                        "Prem Month": client[6],
                        "Med/Den": client[7],
                        "Comm Scale": client[8],
                        "Retro": client[9],
                        "Exchange": client[10],
                        "Commission Amount": client[11],
                    })
        
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name
    
    def bcbs_kc(self):
        carrier = "BCBS KC"
        data = []
        text = self.extract_text()
        
        agency_type_pattern = r'([a-z, ]+) (statement of commissions)'
        agency_type = re.search(agency_type_pattern,text,re.IGNORECASE)
        
        producers_tables_pattern = r'blue kc producer number ([a-z0-9 -]+) ([a-z ,]+)(.*?)total for'
        clients_pattern = r'^(\d+) (\w+) ([a-z ,]+) ([a-z]+) (\d+\/\d+\/\d+) (\d+\/\d+\/\d+) (\d+) (\$? ?[0-9\.]+) (\d+\%?) (\$? ?[0-9\.]+)'
        
        if not agency_type:
            return None, None
        agency,commission_type = agency_type.groups()
            
        producers_tables = re.findall(producers_tables_pattern,text,re.IGNORECASE|re.MULTILINE|re.DOTALL) 
        
        for producers in producers_tables:
            clients = re.findall(clients_pattern,producers[2],re.IGNORECASE|re.MULTILINE)
            for client in clients:
                data.append({
                    "Carrier": carrier,
                    "Agency": agency,
                    "Type": commission_type,
                    "Producer Number": producers[0],
                    "Producer Name": producers[1],
                    "Group Number": client[0],
                    "Subscriber ID": client[1],
                    "Name": client[2],
                    "Product": client[3],
                    "Effective Date": client[4],
                    "Premium Due Date": client[5],
                    "Membership Year": client[6],
                    "Premium Amount": client[7],
                    "Rate": client[8],
                    "Commission Amount": client[9],
                })
        for i in range(1):
            data[i]["Converted from .pdf by"] = ""
        
        df = pd.DataFrame(data)
        return df, self.pdf_output_name