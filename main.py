import pandas as pd
from DrissionPage import ChromiumPage
import openpyxl

# Create a new Excel workbook
workbook = openpyxl.Workbook()

# Select the first sheet
sheet = workbook.active

page = ChromiumPage()

# Read Excel file into a DataFrame
df = pd.read_excel('company_info.xlsx')

num_rows, num_cols = df.shape


for i in range(num_rows):
    if i % 100 == 0:
        workbook = openpyxl.Workbook()
        # Select the first sheet
        sheet = workbook.active

        # Write data to cells
        sheet['A1'] = 'No'
        sheet['B1'] = 'Company Name'
        sheet['C1'] = 'Website'
        sheet['D1'] = 'Email'
        sheet['E1'] = 'Phone'
        sheet['F1'] = 'Address'
        sheet['G1'] = 'Countries/Regions operating in'
        sheet['H1'] = 'Exhibitors who supply to the following industries'

    companyLink = df.iloc[i, 2]
    companyName = df.iloc[i, 1]
    site = ""
    email = ""
    phone = ""
    address = ""
    country = ""
    countries = ""
    exhibitors = ""
    page.get(companyLink)
    try:
        contactInfo = page.ele(".exhibitor-details-contact-us-links").eles('tag:a')
        contactText = page.ele(".exhibitor-details-contact-us-links").text
        if len(contactInfo) == 3:
            site = contactInfo[0].text
            email = contactInfo[1].text
            phone = contactInfo[2].text
        elif len(contactInfo) == 2:
            if "http" not in contactText:
                email = contactInfo[0].text
                phone = contactInfo[1].text
            elif "@" not in contactText:
                site = contactInfo[0].text
                phone = contactInfo[1].text
            else:
                site = contactInfo[0].text
                email = contactInfo[1].text
        else:
            if "http" in contactText:
                site = contactInfo[0].text
            elif "@" in contactText:
                email = contactInfo[0].text
            else:
                phone = contactInfo[0].text
    except:
        pass
    
    try:
        address = page.ele("#exhibitor_details_address").ele("tag:p").text
    except:
        pass

    categories = page.eles("@data-testid=category")
    for category_ele in categories:
        categoryTitle = category_ele.ele("tag:h4").text.lower()
        newCategoryText = ""
        spans = category_ele.eles("tag:span")
        for index, span in enumerate(spans):
            newCategoryText += span.text
            if index < len(spans) - 1:
                newCategoryText += ", "
        newCategoryText.removesuffix(", ")

        if "countries" in categoryTitle:
            countries = newCategoryText
        elif "exhibitors" in categoryTitle:
            exhibitors = newCategoryText
        else:
            pass

    sheet[f'A{i + 2}'] = i + 1
    sheet[f'B{i + 2}'] = companyName
    sheet[f'C{i + 2}'] = site
    sheet[f'D{i + 2}'] = email
    sheet[f'E{i + 2}'] = phone
    sheet[f'F{i + 2}'] = address
    sheet[f'G{i + 2}'] = countries
    sheet[f'H{i + 2}'] = exhibitors
    
    print(f"-------- {i} Company --------")
    print("name:", companyName)
    print("site:", site)
    print("email:", email)
    print("phone:", phone)
    print("address:", address)
    print("countries:", countries)
    print("exhibitors:", exhibitors)

    if i % 100 == 99 or i == num_rows - 1:
        workbook.save(f'companies_{i + 1 - 99}_{i + 1}.xlsx')
        