import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

INPUT_FOLDER = "input_data"
OUTPUT_FOLDER = "output"
CHARTS_FOLDER = "output/charts"
os.makedirs(INPUT_FOLDER,exist_ok=True)
os.makedirs(OUTPUT_FOLDER,exist_ok=True)
print("Folders Ready")

file_name = input("Enter file (Excel or CSV): ")

file_path = os.path.join(INPUT_FOLDER, file_name)

if not os.path.exists(file_path):
    raise FileNotFoundError("File not found in input folder")

if file_name.endswith(".xlsx"):
    df_raw = pd.read_excel(file_path)
elif file_name.endswith(".csv"):
    df_raw = pd.read_csv(file_path)
else:
    raise ValueError("Please enter an Excel or CSV file")

df = df_raw.copy()
df.head()

missing_summary = df.isnull().sum()
print("missing summary per column:")
print(missing_summary)
df.fillna(df.median(numeric_only=True),inplace=True)

rows_before = len(df)
duplicates_count = df.duplicated().sum()
df.drop_duplicates(inplace=True)
rows_after = len(df)
print(f"Rows Before {rows_before}, Rows After {rows_after} , Duplicates Coumt {duplicates_count}")

numeric_cols = df.select_dtypes(include=np.number).columns
for col in numeric_cols:
    Q1= df[col].quantile(0.25)
    Q3= df[col].quantile(0.75)
    IQR= Q3 - Q1
    Lower = Q1 - 1.5 *IQR
    Upper = Q1 + 1.5 * IQR
    df[col] = np.where(df[col] < Lower , Lower , np.where(df[col] > Upper , Upper , df[col]))
print("Outlier apped using IQR Method")

base_name, ext = os.path.splitext(file_name)

if ext == ".csv":
    cleaned_file = f"cleaned_{base_name}.csv"
    df.to_csv(os.path.join(OUTPUT_FOLDER, cleaned_file), index=False)

elif ext == ".xlsx":
    cleaned_file = f"cleaned_{base_name}.xlsx"
    df.to_excel(os.path.join(OUTPUT_FOLDER, cleaned_file), index=False)

else:
    raise ValueError("Unsupported file format")

print("Cleaned file saved successfully")

os.makedirs(CHARTS_FOLDER, exist_ok=True)

if missing_summary.sum() > 0:
    missing_summary[missing_summary > 0].plot(kind="bar", title="Missing Values per Column")
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_FOLDER, f"{base_name}_missing_values.png"))
    plt.show()
    plt.close()

if len(numeric_cols) > 0:
    df[numeric_cols].boxplot()
    plt.title("Outlier Distribution After Cleaning")
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_FOLDER, f"{base_name}_outliers_boxplot.png"))
    plt.show()
    plt.close()

if len(numeric_cols) > 0:
    main_col = numeric_cols[0]
    df[main_col].hist()
    plt.title(f"{main_col} Distribution")
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_FOLDER, f"{base_name}_{main_col}_distribution.png"))
    plt.show()
    plt.close()

pd.Series([rows_before, rows_after], index=["Before Cleaning", "After Cleaning"])\
    .plot(kind="bar", title="Duplicate Records Impact")
plt.tight_layout()
plt.savefig(os.path.join(CHARTS_FOLDER, f"{base_name}_duplicates_impact.png"))
plt.show()
plt.close()

with pd.ExcelWriter(f"{OUTPUT_FOLDER}/report.xlsx", engine="openpyxl") as writer:
    df.describe().to_excel(writer, sheet_name="Summary Stats")
    missing_summary.to_excel(writer, sheet_name="Missing Values")
    pd.DataFrame({
        "Metric": ["Rows Before", "Rows After", "Duplicates Removed"],
        "Value": [rows_before, rows_after, duplicates_count]
    }).to_excel(writer, sheet_name="Cleaning Log", index=False)

print("Report generated successfully")