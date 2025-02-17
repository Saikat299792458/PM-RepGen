from bijoy2unicode import converter
import re

class extractors:
    def __init__(self):
        self.test = converter.Unicode()
    
    def refresh(self):
        self.data = {
        "PM No": "",
        "PM Date": "",
        "Case No":"",
        "Ref": "",
        "Date": "",
        "Name":"",
        "Sex": "",
        "Age": "",
        "ManNat": "",
        "ManSui": "",
        "ManHom": "",
        "ManAcc": "",
        "Built": "",
        "RMAll":"",
        "RMPart":"",
        "Decomposition":"",
        "ChemRoutine":"",
        "ChemOthers":"",
        "Histo":"",
        "HVS":"",
        "MOthers":"",
        "DNA":"",
        "VNO":"Not Sent",
        "Bladder": "",
        "GenM":"",
        "Uterus":"",
        "Vagina":"",
        "Hymen":"",
        }

    def table1(self, table,x):
        # Iterate through the rows of the table
        for row in table.rows:
            for cell in row.cells:
                # Convert ANSI-based Bangla text to Unicode
                textE = cell.text.strip()
                textB = self.test.convertBijoyToUnicode(cell.text.strip())

                # Extract PM No and PM Date
                if "PM No" in textE:
                    parts = textE.split("Date-")
                    if len(parts) == 2:
                        self.data["PM No"] = parts[0].replace("PM No-", "").strip()
                        self.data["PM Date"] = parts[1].strip().rstrip(".")
                # Extract PM No and PM Date
                if "ময়না তদন্ত্ম নং-" in textB:
                    parts = textB.split("তাং-")
                    self.data["PM No"] = parts[0].replace("ময়না তদন্ত্ম নং-", "").strip()
                    #data["PM No"] = int(re.findall(r"\d+",parts[0])[0])
                    if len(parts) == 2:
                        self.data["PM Date"] = parts[1].strip()

                # Extract Ref
                if "Ref:" in textE:
                    self.data["Ref"] = textE.replace("Ref:-", "").strip()
                # Extract Ref (সূত্র)
                if "সূত্র" in textB:
                    self.data["Ref"] = textB.replace("সূত্র:-", "").strip()

                # Extract Date (তাং)
                if "Case No-" in textE:
                    parts = textE.split("Date-")
                    if len(parts) == 2:
                        self.data["Case No"] = int(re.findall(r"\d+",parts[0])[0])
                        self.data["Date"] = parts[1].strip().rstrip(".")
                if "মামলা নং" in textB:
                    parts = textB.split("তাং-")
                    self.data["Case No"] = int(re.findall(r"\d+",parts[0])[0])
                    if len(parts) == 2:
                        self.data["Date"] = parts[1].strip()

    def table2(self, table,x):
        cell = table.rows[0].cells[2].text
        parts = cell.split("Female")
        if len(parts) == 2:
            self.data["Sex"] = "Female"
        else:
            parts = cell.split("Male")
            if len(parts) == 2:
                self.data["Sex"] = "Male"
        if len(parts) == 2:
            self.data["Name"] = parts[0].rstrip("., ")
        else:
            self.data["Name"] = cell.split(",")[0]
        match = re.search(r"(\d+)\s*[Yy]ears", cell)
        self.data["Age"] = int(match.group(1)) if match else ""

    def table3(self, table,x):
        # Regex pattern
        pattern_built = r"Built:\s*([A-Za-z]+)"  # Extracts 'Average', 'Thin', etc.
        pattern_rm = r"RM: \s*Present(?:\s*all over)?"  # Captures 'Present' and 'Present all over'
        cell = table.rows[0].cells[2].text
        built_match = re.search(pattern_built, cell)
        rm_match = re.search(pattern_rm, cell)

        self.data["Built"] = built_match.group(1) if built_match else ""
        rm = rm_match.group(0).lstrip("RM: ") if rm_match else ""
        self.data["RMAll"] = "All Over" if rm == "Present all over" else ""
        self.data["RMPart"] = "Partial" if rm == "Present" else ""
        if "decompos" in cell:
            self.data["Decomposition"] = "Decomposed"


    def table4(self, table, x):
        ## Extract Viscera Sent Information
        cell = table.rows[-1].cells[-1].text.lower()
        if "chemical" in cell:
            self.data["VNO"] = ""
            if "viscera" in cell:
                self.data["ChemRoutine"] = "Sent"
            else:
                self.data["ChemOthers"] = "Sent"
        if "histopatho" in cell:
            self.data["VNO"] = ""
            self.data["Histo"] = "Sent"
        if "microscopic" in cell:
            self.data["VNO"] = ""
            if "hvs" in cell or "high vaginal" in cell:
                self.data["HVS"] = "Sent"
            else:
                self.data["MOthers"] = "Sent"
        if "dna" in cell:
            self.data["VNO"] = ""
            self.data["DNA"] = "Sent"
        
        ## Extract Bladder and genitalia
        self.data["Bladder"] = table.rows[9].cells[2].text
        gen = table.rows[10].cells[2].text
        if self.data["Sex"] == "Male":
            self.data["GenM"] == gen
        else:
            org = ["Uterus", "Vagina", "Hymen", "Genitalia"] #Empty+ Healthy, -Intact.+ 
            for i in org:  # Iterate over all terms
                parts = gen.split(i, 1)  # Split only at first occurrence
                if len(parts) == 2:
                    value = parts[1]
                    gen = parts[0]
                    for j in org:  # Ensure we extract only up to the next keyword
                        pt2 = value.split(j, 1)
                        if len(pt2) == 2:
                            value = pt2[0]
                            gen = gen + j + pt2[1]

                    if i == "Genitalia":
                        self.data["Vagina"] = value.strip("-: ,=")
                    else:
                        self.data[i] = value.strip("-: ,=")
            if self.data["Vagina"] == "":
                self.data["Vagina"] = gen.strip("-: ,=")


    def table5(self, table,x):
        for row in table.rows[0:2]:
            cell = row.cells[0].text
            if "suicid" in cell:
                self.data["ManSui"] = "Suicide"
            elif "homicid" in cell:
                self.data["ManHom"] = "Homicide"
            elif "accident" in cell:
                self.data["ManAcc"] = "Accident"
            elif "natural" in cell:
                self.data["ManNat"] = "Natural"