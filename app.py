# Dummy Data (modified to include duplicate SFIDs)
sf_ids = ['001abc123', '001def456', '001ghi789']
sfdc_data = [
    {'SFID': '001abc123', 'Opportunity_Name': 'Acme Corp - Deal A (First)', 'Account_Name': 'Acme Corporation'},
    {'SFID': '001abc123', 'Opportunity_Name': 'Acme Corp - Deal A (Second)', 'Account_Name': 'Acme Corporation'},  # Duplicate SFID
    {'SFID': '002xyz987', 'Opportunity_Name': 'Beta Inc - Project X', 'Account_Name': 'Beta Incorporated'},
    {'SFID': '001def456', 'Opportunity_Name': 'Gamma Ltd - Contract 2', 'Account_Name': 'Gamma Limited'},
    {'SFID': '001ghi789', 'Opportunity_Name': 'Delta Co - Renewal', 'Account_Name': 'Delta Company'}
]

# Template data (same as before)
weekly_template_data = [
    {'SFID': '001abc123', 'Opportunity Name': '', 'Account Name': ''},
    {'SFID': '001def456', 'Opportunity Name': '', 'Account Name': ''},
    {'SFID': '001jkl000', 'Opportunity Name': '', 'Account Name': ''}
]

# --- Logic Starts Here ---

print("Starting the simple automation...")
print(f"\nSFIDs from weekly report: {sf_ids}")

# 1. Filter Salesforce Data
filtered_sfdc_data = [record for record in sfdc_data if record['SFID'] in sf_ids]
print(f"\nFiltered Salesforce Data: {filtered_sfdc_data}")

# 2. Map Data to Template
updated_template_data = []
for template_row in weekly_template_data:
    if template_row['SFID'] in sf_ids:
        for sfdc_record in filtered_sfdc_data:
            if template_row['SFID'] == sfdc_record['SFID']:
                updated_template_row = template_row.copy()
                updated_template_row['Opportunity Name'] = sfdc_record['Opportunity_Name']
                updated_template_row['Account Name'] = sfdc_record['Account_Name']
                updated_template_data.append(updated_template_row)
                break
    else:
        updated_template_data.append(template_row)

# 3. Print the Updated Template Data
print("\nUpdated Template Data:")
for row in updated_template_data:
    print(row)

print("\nSimple automation complete!")