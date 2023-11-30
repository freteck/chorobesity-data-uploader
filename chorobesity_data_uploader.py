import openpyxl as op
import statistics
import requests
import math
import json
import sys
    
def is_state(id):
    return 'name' in id.keys()

def clean(str):
    if str is None:
        return "NONE"
    return str.lower().replace(" ", "_").replace("'", '')

hdr = {"Content-Type": "application/json"}

if len(sys.argv) != 3:
    print("Incorrect usage. Should be python3 chorobesity_data_uploader.py <backend_base_url> <data_filename>")
    exit(1)

base_url = sys.argv[1]
filename = sys.argv[2]

if base_url.endswith('/'):
    base_url = base_url[:-1]


if ".xlsx" not in filename:
    print(f"Invalid filename: \"{filename}\". Please ensure your file was saved with a .xlsx extension was obtained from the \"2023 County Health Rankings National Data\" download link located at https://www.countyhealthrankings.org/explore-health-rankings/rankings-data-documentation")
    exit(1)

print(f"Running Step 1/1: Opening {filename}...")
try:
    wb = op.load_workbook(filename)
    sheet1 = wb['Additional Measure Data']
    sheet2 = wb['Ranked Measure Data']
except:
    print("File formatting invalid. Please ensure your file was obtained from the \"2023 County Health Rankings National Data\" download link located at https://www.countyhealthrankings.org/explore-health-rankings/rankings-data-documentation")
    exit(1)

print("Step 1/3 Complete.")
try:
    national_obesity_data = []
    national_diabetes_data = []

    state_obesity_data = {
        'alabama': [],
        'alaska': [],
        'arizona': [],
        'arkansas': [],
        'california': [],
        'colorado': [],
        'connecticut': [],
        'delaware': [],
        'florida': [],
        'georgia': [],
        'hawaii': [],
        'idaho': [],
        'illinois': [],
        'indiana': [],
        'iowa': [],
        'kansas': [],
        'kentucky': [],
        'louisiana': [],
        'maine': [],
        'maryland': [],
        'massachusetts': [],
        'michigan': [],
        'minnesota': [],
        'mississippi': [],
        'missouri': [],
        'montana': [],
        'nebraska': [],
        'nevada': [],
        'new_hampshire': [],
        'new_jersey': [],
        'new_mexico': [],
        'new_york': [],
        'north_carolina': [],
        'north_dakota': [],
        'ohio': [],
        'oklahoma': [],
        'oregon': [],
        'pennsylvania': [],
        'rhode_island': [],
        'south_carolina': [],
        'south_dakota': [],
        'tennessee': [],
        'texas': [],
        'utah': [],
        'vermont': [],
        'virginia': [],
        'washington': [],
        'west_virginia': [],
        'wisconsin': [],
        'wyoming': [],
        'district_of_columbia': []
    }

    state_diabetes_data = {
        'alabama': [],
        'alaska': [],
        'arizona': [],
        'arkansas': [],
        'california': [],
        'colorado': [],
        'connecticut': [],
        'delaware': [],
        'florida': [],
        'georgia': [],
        'hawaii': [],
        'idaho': [],
        'illinois': [],
        'indiana': [],
        'iowa': [],
        'kansas': [],
        'kentucky': [],
        'louisiana': [],
        'maine': [],
        'maryland': [],
        'massachusetts': [],
        'michigan': [],
        'minnesota': [],
        'mississippi': [],
        'missouri': [],
        'montana': [],
        'nebraska': [],
        'nevada': [],
        'new_hampshire': [],
        'new_jersey': [],
        'new_mexico': [],
        'new_york': [],
        'north_carolina': [],
        'north_dakota': [],
        'ohio': [],
        'oklahoma': [],
        'oregon': [],
        'pennsylvania': [],
        'rhode_island': [],
        'south_carolina': [],
        'south_dakota': [],
        'tennessee': [],
        'texas': [],
        'utah': [],
        'vermont': [],
        'virginia': [],
        'washington': [],
        'west_virginia': [],
        'wisconsin': [],
        'wyoming': [],
        'district_of_columbia': []
    }

    data = []
    print("Running Step 2/3: Data Wrangling...")
    for row in range(3, sheet1.max_row + 1):
        state = clean(sheet1[f'B{row}'].value)
        county = clean(sheet1[f'C{row}'].value)
        population = sheet1[f'JG{row}'].value
        dpa = sheet1[f'CG{row}'].value
        opa = sheet2[f'BN{row}'].value
        obj_valid = dpa is not None \
                    and opa is not None \
                    and population is not None
        if (obj_valid):
            state_obesity_data[state].append(math.floor(opa))
            state_diabetes_data[state].append(math.floor(dpa))

        if county == "NONE":
            if (obj_valid):
                national_obesity_data.append(math.floor(opa))
                national_diabetes_data.append(math.floor(dpa))
            obj = {
                "name": f"{state}",
                "valid_data": obj_valid,
                "population": population if obj_valid else -1,
                "diabetes_percentage_afflicted": math.floor(dpa) if obj_valid else -1,
                "obesity_percentage_afflicted": math.floor(opa) if obj_valid else -1,
                "diabetes_population_afflicted": math.floor(population * dpa) if obj_valid else -1,
                "obesity_population_afflicted": math.floor(population * opa) if obj_valid else -1,
                "mean_obesity_percentage": round(sum(state_obesity_data[state]) / len(state_obesity_data[state]), 2) if obj_valid else -1,
                "mean_diabetes_percentage": round(sum(state_diabetes_data[state]) / len(state_diabetes_data[state]), 2) if obj_valid else -1,
                "std_obesity_percentage": 0,
                "std_diabetes_percentage": 0
            }
        else:
            obj = {
                "id": f"{state}+{county}",
                "state": f"{state}",
                "county": county,
                "valid_data": obj_valid,
                "population": population if obj_valid else -1,
                "diabetes_percentage_afflicted": math.floor(dpa) if obj_valid else -1,
                "obesity_percentage_afflicted": math.floor(opa) if obj_valid else -1,
                "diabetes_population_afflicted": math.floor(population * dpa) if obj_valid else -1,
                "obesity_population_afflicted": math.floor(population * opa) if obj_valid else -1
            }
        data.append(obj)
except:
    print("Step 2/3: Data collection FAILED")
    exit(1)

print("Step 2/3 Complete.")

try:
    print("Running Step 3/3: Data Upload...")
    requests.get(f"{base_url}/api/states/flush/")
    for obj in data:
        if is_state(obj):
            obj["std_obesity_percentage"] = round(statistics.pstdev(state_obesity_data[obj['name']]), 2)
            obj["std_diabetes_percentage"] = round(statistics.pstdev(state_diabetes_data[obj['name']]), 2)
            requests.post(f"{base_url}/api/states/", headers=hdr, data=json.dumps(obj).replace("'", '"'))
        else:
            requests.post(f"{base_url}/api/states/?s={obj['state']}", headers=hdr, data=json.dumps(obj).replace("'", '"'))

    obj = {
        "name": "national",
        "valid_data": True,
        "population": -1,
        "diabetes_percentage_afflicted": -1,
        "obesity_percentage_afflicted": -1,
        "diabetes_population_afflicted": -1,
        "obesity_population_afflicted": -1,
        "mean_obesity_percentage": round(sum(national_obesity_data) / len(national_obesity_data), 2),
        "mean_diabetes_percentage": round(sum(national_diabetes_data) / len(national_diabetes_data), 2),
        "std_obesity_percentage": round(statistics.pstdev(national_obesity_data), 2),
        "std_diabetes_percentage": round(statistics.pstdev(national_obesity_data), 2)
    }
    requests.post(f"{base_url}/api/states/", headers=hdr, data=json.dumps(obj).replace("'", '"'))
except:
    print(f"Step 3/3: Data upload FAILED. Please ensure the chorobesity backend is running at \"{base_url}\"")
    exit(1)
print("Step 3/3 Complete...")
print(f"Data Upload Complete! Updated database can be viewed at \"{base_url}/api/states/\"")
wb.close()
