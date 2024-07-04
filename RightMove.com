import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import json

def extract_property_data(driver):
    """
    Extracts property data for a given borough from the JSON model using Selenium.
    """
    script_content = driver.execute_script("return JSON.stringify(window.jsonModel);")
    if script_content:
        json_model = json.loads(script_content)
        return json_model
    return None

BOROUGHS = {
    "S6 3BR": "5E1213229",
    "B21 8NL": "5E1297903",
    "WF10 4PD": "5E940753",
    "L22 5QD": "5E445072",
    "CH41 6JB": "5E3760542",
    "CH42 3UN": "5E4051190",
    "CH46 1QQ": "5E4179405",
    "CH42 1QB": "5E1032603",
    "L21 4LY": "5E444773",
    "OL2 8EP": "5E4114305",
    "NR1 1ND": "5E3608016",
    "NR2 4SD": "5E3871325",
    "IP4 2RS": "5E412377",
    "LS15 0AR": "5E491828",
    "WS11 1LA": "5E4185577",
    "WS11 1JW": "5E4665902",
    "DE1 1NL": "5E221583",
    "LS12 3BW": "5E490480",
    "ST13 8EP": "5E823438",
    "PR1 1UT": "5E687393",
    "WS2 8AL": "5E4172057",
    "WS2 8DE": "5E4097312",
    "IP3 8NZ": "5E410734",
    "IP3 8NL": "5E410725",
    "IP3 8NW": "5E410732",
    "CH42 3QA": "5E4188279",
    "CH43 1TJ": "5E1676696",
    "CH43 4SY": "5E1032782",
    "M40 7NS": "5E1143971",
    "CA1 1SR": "5E4055261",
    "DH1 2DS": "5E3889332",
    "L8 3SQ": "5E452024",
    "L8 1TH": "5E451970",
    "TS20 2JD": "5E4059357",
    "SR6 0HJ": "5E4163760"
}


def main():
    # Setup Selenium WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    # Create an Excel writer object to save data into a single Excel file
    excel_writer = pd.ExcelWriter(r"C:\Users\VikasRahar\Desktop\RightmoveData.xlsx", engine="xlsxwriter")

    for borough, borough_code in BOROUGHS.items():
        index = 0
        print(f"We are scraping the borough: {borough} with code: {borough_code}")
        
        all_links = []
        all_addresses = []
        all_titles = []
        all_descriptions = []
        all_monthly_rents = []
        all_weekly_rents = []
        all_bedrooms = []
        all_bathrooms = []
        all_property_statuses = []

        while True:
            url = f"https://www.rightmove.co.uk/property-to-rent/find.html?locationIdentifier=POSTCODE%{borough_code}&radius=0.25&index={index}&propertyTypes=&includeLetAgreed=true&mustHave=&dontShow=&furnishTypes=&keywords="
            driver.get(url)
            
            # Wait for JavaScript to load the content
            time.sleep(5)

            # Extract JSON model using Selenium
            json_model = extract_property_data(driver)
            if json_model:
                properties = json_model.get('properties', [])
                if not properties:
                    print("No more properties found.")
                    break
                
                for property in properties:
                    monthly_rent = "N/A"
                    weekly_rent = "N/A"
                    link = f"https://www.rightmove.co.uk/properties/{property['id']}"
                    all_links.append(link)

                    address = property.get('displayAddress', 'N/A')
                    all_addresses.append(address)

                    title = property.get('propertyTypeFullDescription', 'N/A')
                    all_titles.append(title)

                    description = property.get('summary', 'N/A')
                    all_descriptions.append(description)

                     # Check if displayPrices is present and has at least one price info
                    # Check if price information is available
                    if 'price' in property:
                        price_info = property['price']
                        # Check the frequency and assign the primary price accordingly
                        if price_info.get('frequency') == 'monthly':
                            monthly_rent = f"£{price_info.get('amount', 0)} pcm"
                        elif price_info.get('frequency') == 'weekly':
                            weekly_rent = f"£{price_info.get('amount', 0)} pw"
                        
                        # Extract additional display prices if available
                        for display_price in price_info.get('displayPrices', []):
                            display_price_value = display_price.get('displayPrice', '')
                            if 'pcm' in display_price_value.lower() and monthly_rent == "N/A":
                                monthly_rent = display_price_value
                            elif 'pw' in display_price_value.lower() and weekly_rent == "N/A":
                                weekly_rent = display_price_value

                    # Append the extracted prices to their respective lists
                    all_monthly_rents.append(monthly_rent)
                    all_weekly_rents.append(weekly_rent)


                    bedrooms = property.get('bedrooms', 'N/A')
                    all_bedrooms.append(bedrooms)

                    bathrooms = property.get('bathrooms', 'N/A')
                    all_bathrooms.append(bathrooms)

                    property_status = property.get('displayStatus', 'N/A')  # Extracting property status
                    all_property_statuses.append(property_status)

            else:
                print("JSON model not found or couldn't be extracted.")

            index += 24

        data = {
            "Link": all_links,
            "Address": all_addresses,
            "Title": all_titles,
            "MonthlyRent": all_monthly_rents,
            "WeeklyRent": all_weekly_rents,
            "Bedrooms": all_bedrooms,
            "Bathrooms": all_bathrooms,
            "Property Status": all_property_statuses, 
            "Description": all_descriptions,
        }

        # Create a DataFrame and save it as a new sheet within the Excel file with the borough name as the sheet name
        df = pd.DataFrame(data)
        df.to_excel(excel_writer, sheet_name=borough, index=False)  # Removed the encoding argument
        print(f"Data for {borough} has been successfully saved to the Excel file as a separate sheet. Total properties scraped: {len(all_links)}")


    # Save and close the Excel writer
    #excel_writer.save()
    excel_writer._save()  # Add this line to save the Excel file

    driver.quit()

if __name__ == "__main__":
    main()

