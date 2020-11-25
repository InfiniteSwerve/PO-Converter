import pandas as pd
import time
import numpy as np
from datetime import date


def convert(from_file):
    # Track how long the program takes to run
    start_time = time.time()
    Path="covn\\"
    POFile = from_file
    # POFile = 'PO Details.xls'  # Small Test File
    # POFile = Path+'PO Details 8.31.20.xls' # Big Test File
    TemplateFile = Path+'Sample_Order_Import_File.xlsx'
    PNFile = Path+"Part_Number.xlsx"
    ADDRESSFile = Path+"Lam Addresses.xlsx"
    LEADFile = Path+"Part Number Table.xlsx"
    PO = pd.read_excel(POFile)
    PN = pd.read_excel(PNFile)
    OI = pd.read_excel(TemplateFile)
    MOI = pd.read_excel(POFile)
    AD = pd.read_excel(ADDRESSFile)
    LD = pd.read_excel(LEADFile)

    # Move PO info over to Order Import
    OI['PO_Number'] = PO['PO Number']
    OI['Order_Date'] = PO['PO Create Date']
    OI['Shipping_Address_Contact_Name'] = PO['Buyer name']
    OI['Revision_Level'] = PO['Part Rev']
    OI['Part_Description'] = PO['Desc']
    OI['Status'] = 'Firm'
    OI['Unit_Price'] = PO['Unit Price']
    OI['Pricing_Unit_Of_Measure'] = PO['UOM']
    OI['Order_Release_Quantity'] = PO['Qty Due']
    OI['Order_Release_Date'] = PO['Due Date']
    OI['EDI_Document_Number'] = 850
    OI['Order_Release_Release_Number'] = PO['Schd Line']



    # Find list of all full part numbers in E2. Find correct parts, and flag parts if there is a mismatch.
    # We're assuming that no part numbers contain whitespace in the literal part number. This allows us to account
    # for things like (p) parts
    part_column = pd.Series(PN['Part_Number'])
    rev_column = pd.Series(PN['Revision_Level'])
    # Create a buffer we can give part numbers to
    part_queue = OI['Part_Number'].copy()
    multi_count = 0
    # TODO loops are slow in pandas, when everything is settled I'll see about refactoring this using vectorization
    for i in range(len(PO)):
        part = PO['Material'][i]
        rev = PO['Part Rev'][i]

        # find all the parts that match the part number column and revision column
        cache = part_column[((part_column == part) | (part_column == part + " REV " + str(rev))) & (rev_column == rev)]

        # Flag parts with more than one matching version, or no version at al
        if len(cache) == 1:
            part_queue[i] = cache.iloc[0]
        elif len(cache) > 1:
            part_queue[i] = part
            OI.at[i, 'EDI_Status'] = 'Review_Multiple_PN\'s'
        elif len(cache) == 0:
            part_queue[i] = part
            OI.at[i, 'EDI_Status'] = 'Review_Missing_PN\'s'
        else:
            part_queue[i] = part
            OI.at[i, 'EDI_Status'] = 'Review'

    # Load the part column (with flags) into the new PO
    OI['Part_Number'] = part_queue

    # Create a buffer we can give addresses to
    address_queue = pd.DataFrame(np.random.randn(len(PO),5))

    # Find matching address in PO
    zip_column = pd.Series(AD['Postal_Code'])
    house_column = pd.DataFrame(AD[['Street_Address','City','State_Code','Postal_Code','Country_Code']])

    for i in range(len(PO)):
        zip = PO['Zip'][i]
        house_num = PO['House #'][i]

        # in_db is a boole column with True where the addresses line up
        in_db = (zip_column == zip) & house_column['Street_Address'].str.contains(fr'{house_num}')
        # If the address is present then put it in the address buffer, otherwise flag for row removal
        if in_db.any():
            # Get all the shipping address info. We use a list comprehension at the end because pandas won't set values based on house_column[in_db] for an unknown reason.
            address_queue.iloc[i] = [x for x in house_column[in_db].iloc[0]]
        else:
            address_queue.iloc[i] = ""
            OI.at[i, 'EDI_Status'] = 'Review_address'

    # Put all the address information into the Import Order file
    OI[['Shipping_Address_Street_Address','Shipping_Address_City','Shipping_Address_State_Code',
        'Shipping_Address_Postal_Code','Shipping_Address_Country_Code']] = address_queue

    # Hardcode in the billing address because we don't expect the LAM billing address to change
    OI[['Billing_Address_Street_Address','Billing_Address_City','Billing_Address_State_Code','Billing_Address_Postal_Code',
        'Billing_Address_Country_Code']] = ["Accounts Payable Dept. #C3A11155 SW Leveton Drive",'Tualatin','OR','97062',
                                            'United States']


    # Make a list indexed by part numbers with values of lead times for said parts.
    lead_times = part_queue.copy()
    lead_times = lead_times.map(pd.Series(LD['Lead Time'].values,index=LD['Part_Number']).to_dict(), na_action=0).replace(np.nan,0)


    # Make a copy of all our dates so we can modify it
    date_column = PO[['Prom Date', 'Due Date', 'Need Date', 'PO Create Date']].copy()

    # Add lead time to PO creation date
    date_column['PO Create Date'] = (pd.to_datetime(date_column['PO Create Date']) + pd.to_timedelta(lead_times, 'days'))
    # Set the order date to the max of due date, promise date, and PO date + lead time
    OI['Order_Date'] = pd.to_datetime(date_column.stack(), errors='coerce').unstack().max(axis=1).dt.strftime('%m/%d/%Y')



    OI['PO_Number'] = '\'' + OI['PO_Number'].astype(str)

    # Add current date to track order entry date
    OI['Header_User_Date1'] =date.today().strftime('%m/%d/%Y')

    # Flag part if either:
    # o	PO Type (Column “A”) is “R” or “C”
    # o	FA (column “AA”) is marked “Y”
    manual_rows = (PO['T'] == 'R') | (PO['FA'] == 'Y') | (PO['Desc'].str.contains('DNP'))
    auto_rows = ~manual_rows
    MOI = MOI[manual_rows]
    OI = OI[auto_rows]

    # Export to two sheets.
    OI.to_excel(r'Orders\Order_Import.xlsx', index=False)
    MOI.to_excel(r'Orders\Manual_Order_Import.xlsx', index=False)



    # Stop time tracking once program is over
    print("time elapsed: {:.2f}s".format(time.time() - start_time))

# convert(0)
