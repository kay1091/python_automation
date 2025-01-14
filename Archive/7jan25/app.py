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

        # SFID Entry
        sfid_entry = {
            "Type": fake.random_element(["New", "Existing"]),
            "SBU": fake.random_element(["SBU1", "SBU2", "SBU3"]),
            "Account Name": account_name,
            "SFID": sfid,
            "Opportunity Name": opportunity_name,
            "$ Value (M)": round(fake.random_number(digits=2, fix_len=True) / 10, 2),
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
            "Amount": round(fake.random_number(digits=5), 2),
            "Expected Revenue Currency": "USD",
            "Expected Revenue": round(fake.random_number(digits=4), 2),
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
            "Amount (converted)": round(fake.random_number(digits=5), 2),
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


def main():
    # Create directories if they don't exist
    os.makedirs("input", exist_ok=True)
    os.makedirs("output", exist_ok=True)

    num_records = 20  # Set num records as per requirement
    sfid_df, sfdc_dump_df = generate_fake_data(num_records=num_records, seed=42)


    try:
        # Save data to Excel files
        sfid_df.to_excel("input/SFID_file.xlsx", index=False)
        sfdc_dump_df.to_excel("input/SFDC_dump.xlsx", index=False)
        
         #Create empty template with just a header:
        template_df = pd.DataFrame(columns=[
            "Account Name", "SFID", "Created Date", "Opportunity Name", "Opportunity Description",
            "Group SBU", "Created By", "Stage", "Est Deal Value in USD", "Opp Type", "Vertical Practice",
            "Tech. Practice", "Service Offering", "Engagement Type", "Probability", "Close Date",
            "Next Steps", "Loss Stage", "Lost Reason", "Age", "BOLT", "Doc. Recvd. Date", "Category",
            "Partner Details", "Proposed Sub. Date", "Domain Practice", "Tech Practice", "Solution SPOCs",
            "Delivery SPOC", "Proposal Owner", "Allocation% Proposal Owner 1", "Proposal Owner 2",
            "Allocation% Proposal Owner 2", "Proposal Writer", "Allocation% Proposal Writer 1",
            "Proposal Writer 2", "Allocation% Proposal Writer 2", "Orals SPOC", "Bid Director",
            "Proposal Updates", "Proposal Status", "Actual Sub. Date", "Commercial Value", "DSC",
            "Opportunity Stage", "Post Sub. Activity", "PSA Activity Status", "PSA Activity Cls. Date",
            "PSA Activity Update", "Est. Deal Value", "Large Deal", "SBU Mapping", "Created in Week",
            "Submitted in Week", "Closing in Week", "PSA Comp. in Week", "Sub. Dt. Mapping",
            "Cl. Dt. Mapping", "PSA Comp. Dt. Mapping", "Opp. Status", "Sb. FY", "Sb. Qtr.",
             "Cl. FY", "Cl. QTR", "TCV Brk. Up", "BFS/ETS", "Direct/ Related",
            "OB Related Status", "TCV Related Status"
         ])
        
        #Save template to excel
        template_file = "input/Weekly_Template.xlsx"
        writer = pd.ExcelWriter(template_file, engine = 'xlsxwriter')
        template_df.to_excel(writer, sheet_name="SFDC",index=False)
        writer.close()


        print("Random test data has been generated and saved.")

    except Exception as e:
        print(f"Error generating or saving excel file: {e}")


if __name__ == "__main__":
    main()