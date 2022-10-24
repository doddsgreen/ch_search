import requests
import pandas as pd
import xlrd

# Set up df and API key
company_df = pd.DataFrame(index=["Company name","Date of Incorporation", "Company Type", "Address line 1", "City", "Postcode", "Country"])
officer_df = pd.DataFrame(columns=['Officer name'])
company_number = input("Enter company code: ")
api_key = "3d1a93e2-ed53-4b03-b1cc-ea83646e0f9e"

# Company name
url = f"https://api.company-information.service.gov.uk/company/{company_number}"
company_profile=requests.get(url, auth=(api_key,""))
company_profile_data = company_profile.json()
company_name = company_profile_data["company_name"]
company_name_format = company_name.title()

# Company incorporation
incorporation_url = f"https://api.company-information.service.gov.uk/search/companies"
incorporation_profile =requests.get(url, auth=(api_key,""))
incorporation_profile_data = incorporation_profile.json()
incorporation = incorporation_profile_data["date_of_creation"]

# Company type
company_type = company_profile_data["type"]

# Company address
address = company_profile_data["registered_office_address"]
address_line_1 = address["address_line_1"]
city = address["locality"]
post_code = address["postal_code"]
country = address["country"]

# Parse df
company_df.loc[["Company name"], 0]=company_name_format
company_df.loc[["Date of Incorporation"], 0]=incorporation
company_df.loc[["Company Type"], 0]=company_type
company_df.loc[["Address line 1"], 0]=address_line_1
company_df.loc[["City"], 0]=city
company_df.loc[["Postcode"], 0]=post_code
company_df.loc[["Country"], 0]=country


# Officers
officer_url = f"https://api.company-information.service.gov.uk/company/{company_number}/officers"
paramaters = {"register_type": "directors", "register_view": "false"}
officer_resp=requests.get(officer_url, auth=(api_key,""), params=paramaters)
data = officer_resp.json()
officers = data["items"]
count_officers = 0
for officer in officers:
    if "resigned_on" not in officer:
        officer_full_name = (officer["name"])
        officer_given_names = officer_full_name.rsplit(',', 1)[1]
        officer_surname = officer_full_name.rsplit(',', 1)[0]
        officer_surname_format = officer_surname.title()
        officer_name = officer_given_names + " " + officer_surname_format
        officer_df.loc[count_officers] = [officer_name]
        count_officers = count_officers+1
    else:
        pass

with pd.ExcelWriter("Companies_House_Extract.xlsx") as writer:
    company_df.to_excel(writer, sheet_name = "Company Details")
    officer_df.to_excel(writer, sheet_name = "Officers")
