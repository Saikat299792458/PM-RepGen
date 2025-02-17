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
        "Built": "",
        "RMAll":"",
        "RMPart":"",
        "Decomposition":"",
        "Tattoo Marks":"",
        "Bladder": "",
        "GenM":"",
        "Uterus":"",
        "Vagina":"",
        "Hymen":"",
        "ChemRoutine":"",
        "ChemOthers":"",
        "Histo":"",
        "HVS":"",
        "MOthers":"",
        "DNA":"",
        "VNO":"Not Sent",
        "Mode":"",
        "ManNat": "",
        "ManSui": "",
        "ManHom": "",
        "ManAcc": "",
        "Hanging":"",
        "Strangulation":"",
        "Suffocation":"",
        "Drowning":"",
        "Poison":"",
        "Burn":"",
        "Electrocution":"",
        "Abrasion":"",
        "Bruise":"",
        "Laceration":"",
        "Fracture":"",
        "Dislocation":"",
        "IncisedW":"",
        "ChopW":"",
        "StabW":"",
        "SelfInfW":"",
        "FabW":"",
        "DefW":"",
        "SurW":"",
        "GSW":"",
        "Fall":"",
        "RoadAcc":"",
        }
    
    def extInjury(self, cell):
        "Sent lowercase string here"
        ## Extract injury type
        if "surgical wound" in cell:
            self.data["SurW"] = "Surgical Wound"
        if "abras" in cell:
            self.data["Abrasion"] = "Abrasion"
        if "bruise" in cell:
            self.data["Bruise"] = "Bruise"
        if "lacerat" in cell:
            self.data["Laceration"] = "Laceration"
        if "fracture" in cell:
            self.data["Fracture"] = "Fracture"
        if "dislocat" in cell:
            self.data["Dislocation"] = "Dislocation"
        if "incised" in cell:
            self.data["IncisedW"] = "Incised Wound"
        if "chop" in cell:
            self.data["ChopW"] = "Chop Wound"
        if "stab" in cell:
            self.data["StabW"] = "Stab Wound"
        

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
        
        ## Extract more Injury info
        cell = table.columns[2].cells[1].text.lower()
        self.extInjury(cell)


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
        
        ## Extract more Injury info
        cell = table.columns[2].cells[1].text.lower()
        self.extInjury(cell)
        
        ## Extract Bladder and genitalia
        self.data["Bladder"] = table.rows[9].cells[2].text
        gen = table.rows[10].cells[2].text
        if self.data["Sex"] == "Male":
            self.data["GenM"] = gen
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
        "Assumes that the cell has the word Death in it."
        cell = table.rows[0].cells[0] if "death" in table.rows[0].cells[0].text.lower() else table.rows[1].cells[0]
        cell = cell.text.lower()
        ## Extract Mode of Death
        mode = []
        if "asph" in cell:
            mode.append("Asphyxia")
        if "hemorr" in cell:
            if "cranial" in cell:
                mode.append("Intracranial Hemorrhage")
            else:
                mode.append("Hemorrhage")
        if "shock" in cell:
                mode.append("Shock")
        if "opc" in cell:
            mode.append("Antemortem OPC Poisoning")
        if "septisemia" in cell:
            mode.append("Septisemia")
        self.data["Mode"] = " and ".join(mode)
        ## Extract manner of death
        if "suicid" in cell:
            self.data["ManSui"] = "Suicide"
        elif "homicid" in cell:
            self.data["ManHom"] = "Homicide"
        elif "accident" in cell:
            self.data["ManAcc"] = "Accident"
        elif "natural" in cell:
            self.data["ManNat"] = "Natural"
        
        ## Extract Cause of Death
        if "hang" in cell:
            self.data["Hanging"] = "Hanging"
        if "strangulat" in cell:
            self.data["Strangulation"] = "Strangulation"
        if "suffocat" in cell:
            self.data["Suffocation"] = "Suffocation"
        if "drown" in cell:
            self.data["Drowning"] = "Drowning"
        if "poison" in cell:
            self.data["Poison"] = "Poison"
        if "burn" in cell:
            self.data["Burn"] = "Burn"
        if "electro" in cell:
            self.data["Electrocution"] = "Electrocution"
        if "fall" in cell:
            self.data["Fall"] = "Fall from Height"
        if "road traffic" in cell:
            self.data["RoadAcc"] = "Road Traffic Accident"
        if "injur" in cell:
            self.extInjury(cell)
