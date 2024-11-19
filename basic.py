from DrissionPage import ChromiumPage, ChromiumOptions
import openpyxl

# Create a new Excel workbook
workbook = openpyxl.Workbook()

# Select the first sheet
sheet = workbook.active

# Write data to cells
sheet['A1'] = 'No'
sheet['B1'] = 'Company Name'
sheet['C1'] = 'Link'

page = ChromiumPage()
numberOfCompany = 0

for letter in range(ord('a'), ord('z')+1):
    url = f"https://www.ibtmworld.com/en-gb/exhibitor-directory.html?locale=en-GB&query={chr(letter)}#/"
    page.get(url)
    elesName = page.eles('@data-testid=name-control')
    print(f"-------------- {len(elesName)} Companies --------------")
    for index, eleName in enumerate(elesName):
        numberOfCompany += 1
        companyName = eleName.text
        companyLink = eleName.ele('tag:a').link
        print(f"{chr(letter)} {index} - {companyName}")
        sheet[f'A{numberOfCompany + 1}'] = numberOfCompany
        sheet[f'B{numberOfCompany + 1}'] = companyName
        sheet[f'C{numberOfCompany + 1}'] = companyLink

# Save the workbook
workbook.save('company_info.xlsx')

# for letter in range(ord('a'), ord('z')+1):
#     url = f"https://www.ibtmworld.com/en-gb/exhibitor-directory.html?locale=en-GB&query={chr(letter)}#/"
#     page.get(url)
#     elesDesktopRow = page.eles("@data-testid=row-desktop")
#     for index, eleDesktopRow in enumerate(elesDesktopRow):
#         companyName = eleDesktopRow.ele('tag:h3').text
#         phoneNumber = ""
#         email = ""
#         category = ""
#         try:
#             email = eleDesktopRow.ele('@title=Email').link
#         except:
#             pass
#         try:
#             phoneNumber = eleDesktopRow.ele('@title=Phone').link
#         except:
#             pass
#         try:
#             category = eleDesktopRow.ele('.pps-tags').text
#         except:
#             pass
#         print(index, "----", companyName, email, phoneNumber, category)

#     elesDesktopBronze = page.eles("@data-testid=exh-bronze-desktop")
#     for index, eleDesktopBronze in enumerate(elesDesktopBronze):
#         companyName = eleDesktopBronze.ele('tag:h3').text
#         phoneNumber = ""
#         email = ""
#         category = ""
#         # try:
#         #     email = eleDesktopBronze.ele('@title=Email').link
#         # except:
#         #     pass
#         # try:
#         #     phoneNumber = eleDesktopBronze.ele('@title=Phone').link
#         # except:
#         #     pass
#         try:
#             category = eleDesktopBronze.ele('.pps-tags').text
#         except:
#             pass
#         print(index, "----", companyName, email, phoneNumber, category)
#     break