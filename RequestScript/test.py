import csv

# Subject name
subject_name = "Subj1"

# Header name for the single block
header = ["Total"]

# Sample data for the two blocks in each row
data = [
    {"Block1": 80, "Block2": 90},
    {"Block1": 85, "Block2": 92},
    # Add more rows as needed
]

# CSV file path
csv_file_path = f"{subject_name}_grades.csv"

# Writing to CSV file
with open(csv_file_path, mode='w', newline='') as csv_file:
    writer = csv.DictWriter(csv_file, fieldnames=header + [f"{subject_name}1", f"{subject_name}2"])

    # Write the header row
    writer.writeheader()

    # Write data rows
    for row in data:
        # Combine the header block with the data blocks
        writer.writerow({"Total": subject_name, f"{subject_name}1": row["Block1"], f"{subject_name}2": row["Block2"]})

print(f"CSV file '{csv_file_path}' created successfully.")
