import os
import pandas as pd
from faker import Faker
import random
from datetime import datetime, timedelta
import openpyxl


def generate_fake_data(num_records=10, seed=None):
    """Generates consistent data for two excel files using Faker and Pandas library.

    Args:
        num_records (int, optional): The number of records to generate. Defaults to 10.
        seed (int, optional): Seed for random data generation for reproducibility. Defaults to None.

    Returns:
        tuple: A tuple containing the pandas dataframe for both the excel files: (sfid_df, sfdc_dump_df).
    """

    if seed is not None:
        Faker.seed(seed)
        random.seed(seed)

    fake = Faker()

    # Create consistent account names
    account_names = [fake.company() for _ in range(num_records)]

    # Create a map between SFID and Opportunity ID
    sfid_opp_id_map = {fake.uuid4()[:8].upper(): fake.uuid4()[:8].upper() for _ in range(num_records)}
    sfids = list(sfid_opp_id_map.keys())

    # --- Data Generation ---
    sfid_dataframe_data = []
    sfdc_dataframe_data = []

    for account_name in account_names:
        sfid = sfids[account_names.index(account_name)]
        opportunity_id = sfids[account_names.index(account_name)]
        opportunity_name = fake.bs()
        close_date = fake.date_between(start_date='-1y', end_date='today')
        created_date = fake.date_between(start_date='-2y', end_date=close_date)

        # Generate random amount
        random_amount = round(random.uniform(1_000_000, 50_000_000), 2)  # Random number between 1M and 50M

        # SFID Entry
        sfid_entry = {
            "Type": fake.random_element(["New", "Existing"]),
            "SBU": fake.random_element(["SBU1", "SBU2", "SBU3"]),
            "Account Name": account_name,
            "SFID": sfid,
            "Opportunity Name": opportunity_name,
            "$ Value (M)": round(random_amount / 1_000_000, 2),  # Convert to millions
            "Opportunity Description": fake.sentence(),
            "Client Partner": fake.name(),
            "Partner Details": fake.sentence(),
            "Status/ Next Steps": fake.sentence(),
            "Due Date": fake.date_between(start_date='today', end_date='+1y'),
            "Activity Type": fake.word(),
            "Bid Manager": fake.name(),
            "Proposal Writer": fake.name(),
            "Orals SPOC": fake.name(),
            "Solution SPOCs": fake.name(),
            "Delivery Lead": fake.name(),
            "Deal Status": fake.random_element(["Won", "Lost", "In Progress"]),
            "Close Date": close_date,
            "Deal Stage": fake.random_element(["Stage 1", "Stage 2", "Stage 3"]),
            "DSC Status": fake.random_element(["Approved", "Pending", "Rejected"]),
        }

        # SFDC Entry
        sfdc_entry = {
            "Opportunity ID": opportunity_id,
            "Practice": fake.random_element(["Practice1", "Practice2", "Practice3"]),
            "Description": fake.text(max_nb_chars=50),
            "Opportunity Name": opportunity_name,
            "Type": fake.random_element(["Type1", "Type2", "Type3"]),
            "Lead Source": fake.word(),
            "SBU": fake.random_element(["SBU1", "SBU2", "SBU3"]),
            "Billing Country": fake.country(),
            "Amount Currency": "USD",
            "Amount": random_amount,  # Randomly generated amount
            "Expected Revenue Currency": "USD",
            "Expected Revenue": round(random.uniform(1_000, 1_000_000), 2),  # Random revenue
            "Competitor Details Old": fake.sentence(),
            "Close Date": close_date,
            "Next Step": fake.sentence(),
            "Stage": fake.random_element(["Prospecting", "Negotiation", "Closed Won"]),
            "Probability (%)": fake.random_int(min=10, max=100),
            "Fiscal Period": fake.random_element(["2024-Q1", "2024-Q2", "2024-Q3"]),
            "Age": fake.random_int(min=1, max=365),
            "Created Date": created_date,
            "Opportunity Owner": fake.name(),
            "Owner Role": fake.random_element(["Manager", "Director", "VP"]),
            "Account Name": account_name,
            "Project Type": fake.random_element(["Type A", "Type B"]),
            "Technology/Skills": fake.random_element(["Skill1", "Skill2", "Skill3"]),
            "IT Lifecycle": fake.random_element(["Development", "Maintenance"]),
            "Service Offering": fake.word(),
            "Service Category": fake.random_element(["Category1", "Category2"]),
            "Loss Stage": fake.random_element(["Stage1", "Stage2"]),
            "Loss Notes": fake.sentence(),
            "Lost Reason": fake.sentence(),
            "Group SBU": fake.random_element(["Group1", "Group2"]),
            "Vertical Practice": fake.random_element(["Practice1", "Practice2"]),
            "Segment": fake.random_element(["Segment1", "Segment2"]),
            "Created By": fake.name(),
            "Deal Type": fake.random_element(["Type X", "Type Y"]),
            "Industry Solutions": fake.random_element(["Solution1", "Solution2"]),
            "Amount (converted) Currency": "USD",
            "Amount (converted)": random_amount,  # Randomly generated converted amount
            "Virtusa/Polaris": fake.random_element(["Virtusa", "Polaris"]),
            "Quality of Revenue": fake.random_element(["High", "Medium", "Low"]),
            "Proposal Type": fake.random_element(["Type P", "Type Q"]),
            "BOLT Details": fake.sentence(),
            "BOLT Status": fake.random_element(["Open", "Closed"]),
        }

        sfid_dataframe_data.append(sfid_entry)
        sfdc_dataframe_data.append(sfdc_entry)

    sfid_df = pd.DataFrame(sfid_dataframe_data)
    sfdc_dump_df = pd.DataFrame(sfdc_dataframe_data)

    return sfid_df, sfdc_dump_df
