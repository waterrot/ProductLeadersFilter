import pandas as pd

def filter_job_titles(input_file, trash_file, results_file, to_check):
    # Read the input file and reset index
    df = pd.read_excel(input_file).reset_index(drop=True)

    # Create empty output lists
    trash_list = []
    results_list = []

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        job_title = str(row['Job Title']).lower()
        if any(keyword.lower() in job_title for keyword in to_check):
            trash_list.append(row)
        else:
            results_list.append(row)

    # Convert lists to DataFrames
    trash_df = pd.DataFrame(trash_list)
    results_df = pd.DataFrame(results_list)

    # Write the output to separate files
    trash_df.to_excel(trash_file, index=False)
    results_df.to_excel(results_file, index=False)

# Define the strings to check
to_check = ["chemicals", "product owner", "project risk", "Assistant", \
    "project manager", "project owner", "urban communications",  "head of marketing", "product care", \
    "head of finance", "product cyber security", "product policy", "head of innovation", \
    "head of media product", "product introduction", "head of people", "Pedagogisch", \
    "Administrateur", "logistics", "Chief Privacy Officer", "Chief Procurement Officer", \
    "Supply chain", "Coördinerend Privacy Officer", "Procurement", "Domeinmanager", "Dep. PO Sensors", \
    "N/A", "Interior", "Merchandising", "Head of Product Design", "Head of Design & Product development", \
    "Head of design", "Head of Air Product", "Head of Buying and Product Development", "Head of Consumer", \
    "Head of Commercial", "Head of Design", "Head of Food Product", "Head of Governance", "HR", \
    "Head of New Product Development", "Head of Product and Customer Experience", ]

# Define the output file names
input_file = "6trash_leads_no_duplicates.xlsx"
trash_file = "7trash.xlsx"
results_file = "8leads_without_trash.xlsx"

# Call the filtering function
filter_job_titles(input_file, trash_file, results_file, to_check)

