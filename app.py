import os
import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import datetime, timedelta


def generate_fake_data(num_records=10, seed=None):

    """Generates consistent data for two excel files using Faker and Pandas library.

    Args:
        num_records (int, optional): The number of records to generate. Defaults to 10.
        seed (int, optional):  Seed for random data generation for reproducibility. Defaults to None.

    Returns:
        tuple: A tuple containing the pandas dataframe for both the excel files: (sfid_df, sfdc_dump_df).
    """

    if seed is not None:
        Faker.seed(seed)
        random.seed(seed)


    fake = Faker()

    # Create consistent account names
    account_names = [fake.company() for _ in range(num_records)]

    # Create a map between SFID and Opp ID
    sfid_opp_id_map = {fake.uuid4()[:8].upper(): fake.uuid4()[:8].upper() for _ in range(num_records)}

    # Create SFID Data
    sfid_data = {
        "Type": [fake.random_element(["New", "Existing"]) for _ in range(num_records)],
        "SBU": [fake.random_element(["SBU1", "SBU2", "SBU3"]) for _ in range(num_records)],
        "Account Name": account_names,
        "SFID": list(sfid_opp_id_map.keys()),
        "Opportunity Name": [fake.bs() for _ in range(num_records)],
        "$ Value (M)": [round(fake.random_number(digits=2, fix_len=True) / 10, 2) for _ in range(num_records)],
        "Opportunity Description": [fake.sentence() for _ in range(num_records)],
        "Client Partner": [fake.name() for _ in range(num_records)],
        "Partner Details": [fake.sentence() for _ in range(num_records)],
        "Status/ Next Steps": [fake.sentence() for _ in range(num_records)],
        "Due Date": [fake.date_this_year() for _ in range(num_records)],
        "Activity Type": [fake.word() for _ in range(num_records)],
        "Bid Manager": [fake.name() for _ in range(num_records)],
        "Proposal Writer": [fake.name() for _ in range(num_records)],
        "Orals SPOC": [fake.name() for _ in range(num_records)],
        "Solution SPOCs": [fake.name() for _ in range(num_records)],
        "Delivery Lead": [fake.name() for _ in range(num_records)],
        "Deal Status": [fake.random_element(["Won", "Lost", "In Progress"]) for _ in range(num_records)],
         "Close Date": [fake.date_between(start_date='-1y', end_date='today') for _ in range(num_records)],
        "Deal Stage": [fake.random_element(["Stage 1", "Stage 2", "Stage 3"]) for _ in range(num_records)],
        "DSC Status": [fake.random_element(["Approved", "Pending", "Rejected"]) for _ in range(num_records)],
    }

    # Create SFDC Dump Data
    sfdc_dump_data = {
        "Opportunity ID": list(sfid_opp_id_map.values()),
        "Practice": [fake.random_element(["Practice1", "Practice2", "Practice3"]) for _ in range(num_records)],
        "Description": [fake.text(max_nb_chars=20) for _ in range(num_records)],
        "Opportunity Name": [fake.bs() for _ in range(num_records)],
        "Type": [fake.random_element(["Type1", "Type2", "Type3"]) for _ in range(num_records)],
        "Lead Source": [fake.word() for _ in range(num_records)],
        "SBU": [fake.random_element(["SBU1", "SBU2", "SBU3"]) for _ in range(num_records)],
        "Billing Country": [fake.country() for _ in range(num_records)],
        "Amount": [round(fake.random_number(digits=5), 2) for _ in range(num_records)],
        "Expected Revenue": [round(fake.random_number(digits=4), 2) for _ in range(num_records)],
        "Stage": [fake.random_element(["Prospecting", "Negotiation", "Closed Won"]) for _ in range(num_records)],
        "Probability (%)": [fake.random_int(min=10, max=100) for _ in range(num_records)],
        "Fiscal Period": [fake.random_element(["2024-Q1", "2024-Q2", "2024-Q3"]) for _ in range(num_records)],
         "Created Date": [fake.date_between(start_date='-1y', end_date='today') for _ in range(num_records)],
        "Opportunity Owner": [fake.name() for _ in range(num_records)],
        "Account Name": account_names,
    }
    
    sfid_df = pd.DataFrame(sfid_data)
    sfdc_dump_df = pd.DataFrame(sfdc_dump_data)
    
    return sfid_df, sfdc_dump_df

def main():
    
    # Create directories if they don't exist
    os.makedirs("input", exist_ok=True)
    os.makedirs("output", exist_ok=True)
    
    num_records = 20 # Set num records as per requirement
    sfid_df, sfdc_dump_df = generate_fake_data(num_records=num_records, seed=42)

    try:
        # Save data to Excel files
        sfid_df.to_excel("input/SFID_file.xlsx", index=False)
        sfdc_dump_df.to_excel("input/SFDC_dump.xlsx", index=False)
        print("Random test data has been generated and saved.")

    except Exception as e:
        print(f"Error generating or saving excel file: {e}")

if __name__ == "__main__":
    main()