import re

text1 = "Built: Average, RM: Present all over."
text2 = "Built: Average, RM: Present."
text3 = "Built: Thin, RM: Present all over."

# Regex pattern
pattern_built = r"Built:\s*([A-Za-z]+)"  # Extracts 'Average', 'Thin', etc.
pattern_rm = r"RM:\s*Present(?:\s*all over)?"  # Captures 'Present' and 'Present all over'

# Function to extract data
def extract_built_rm(text):
    built_match = re.search(pattern_built, text)
    rm_match = re.search(pattern_rm, text)

    built_value = built_match.group(1) if built_match else ""
    rm_value = rm_match.group(0) if rm_match else ""

    return built_value, rm_value

# Testing the function
print(extract_built_rm(text1))  # ('Average', 'RM: Present all over')
print(extract_built_rm(text2))  # ('Average', 'RM: Present')
print(extract_built_rm(text3))  # ('Thin', 'RM: Present all over')
