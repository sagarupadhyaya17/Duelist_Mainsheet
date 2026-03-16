import pandas as pd
import numpy as np
import io

# ==== File Paths ====
duelist_file = r"Input_Files/duelist_dump_march_11.xlsx"
duelist_main_file = r"Z:/1.Reports Repository/FY 2082.83/1. Duelist/9.Chaitra/Duelist 10th March, 2026.xlsb"
insurance = r"Z:/1.Reports Repository/FY 2082.83/1. Duelist/9.Chaitra/Chaitra Insurance 2082.xlsx"
output_file = r"Output_Files/updated_duelist_march_11.xlsx"

print("Processing... Please wait ⏳")

insurance = pd.read_excel(insurance, dtype={"MainCode": str})

if duelist_file.endswith(".xlsb"):
    duelist = pd.read_excel(
        duelist_file,
        dtype={"MainCode": str, "AcCodeForChg": str, "Nominee": str},
        engine="pyxlsb"
)
else:
    duelist = pd.read_excel(
        duelist_file,
        dtype={"MainCode": str, "AcCodeForChg": str, "Nominee": str},
        engine="openpyxl"
    )  

if duelist_main_file.endswith(".xlsb"):
    duelist_main = pd.read_excel(
        duelist_main_file,
        dtype={"MainCode": str, "AcCodeForChg": str, "Nominee": str},
        engine="pyxlsb"
    )
else:
    duelist_main = pd.read_excel(
        duelist_main_file,
        dtype={"MainCode": str, "AcCodeForChg": str, "Nominee": str},
        engine="openpyxl"
    )

print(len(duelist))
duelist.head(2)

def clean_columns(df):
    df.columns = df.columns.str.strip()
    return df

duelist = clean_columns(duelist)
duelist_main = clean_columns(duelist_main)
insurance = clean_columns(insurance)

insurance.dtypes

insurance["date"] = pd.to_datetime(insurance["date"])

yesterday = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)

insurance["InsurancePremium"] = np.where(
    yesterday < insurance["Date"],
    insurance["InsPremium"],
    0
)

duelist = duelist.copy()
duelist = duelist[(duelist['ClientCode']!="~~~~~")]
print(len(duelist))
duelist.head(2)

print(len(duelist))
print(duelist["MainCode"].nunique())
print(duelist["MainCode"].duplicated().sum())
print("\n")

print(len(duelist_main))
print(duelist_main["MainCode"].nunique())
print(duelist_main["MainCode"].duplicated().sum())
print("\n")

print(len(insurance))
print(insurance["MainCode"].nunique())
print(insurance["MainCode"].duplicated().sum())

insurance[insurance.duplicated("MainCode", keep=False)].head(5)

duelist = duelist.drop_duplicates(subset="MainCode", keep = "last")
duelist_main = duelist_main.drop_duplicates(subset="MainCode", keep = "last")
insurance = insurance.drop_duplicates(subset="MainCode", keep = "last")

print(duelist['AgeingDays'].dtypes)

duelist['AgeingDays'] = pd.to_numeric(duelist['AgeingDays'], errors='coerce')
duelist = duelist.sort_values(by='AgeingDays', ascending=False)
duelist.head(2)

pos = duelist.columns.get_loc("Name") + 1
duelist.insert(pos, "Ageing", np.nan)
duelist.insert(pos, "Loan Type", np.nan)
duelist.insert(pos, "OfficerName", np.nan)

pos = duelist.columns.get_loc("TotCharge") + 1
duelist.insert(pos, "OvDueWithInsurance", np.nan)
duelist.insert(pos, "InsurancePremium", np.nan)
duelist.insert(pos, "TotOvDue", np.nan)

pos = duelist.columns.get_loc("AgeingDays") + 1
duelist.insert(pos, "Bucket", np.nan)

pos = duelist.columns.get_loc("BranchName") + 1
duelist.insert(pos, "Dealer Name", np.nan)

# clean column names
duelist.columns = duelist.columns.str.strip()
duelist_main.columns = duelist_main.columns.str.strip()

# clean MainCode
duelist["MainCode"] = duelist["MainCode"].astype(str).str.strip()
duelist_main["MainCode"] = duelist_main["MainCode"].astype(str).str.strip()

cols = ["OfficerName", "Loan Type", "Dealer Name"]

df = duelist.merge(
    duelist_main[["MainCode"] + cols],
    on="MainCode",
    how="left",
    suffixes=("", "_ref")
)

for col in cols:
    df[col] = df[col].astype(str).str.strip()
    df[col] = df[col].replace(r'^(|None|nan|\s+)$', pd.NA, regex=True)
    df[col] = df[col].fillna(df[f"{col}_ref"])

df.drop(columns=[c + "_ref" for c in cols], inplace=True)

print(len(df))
df.head(2)

# insurance = insurance.rename(columns={
#     "prem":"InsurancePremium"
# })

# merge
df = df.merge(
    insurance[["MainCode", "InsurancePremium"]],
    on="MainCode",
    how="left",
    suffixes=("", "_ref")
)
# fill missing values
df["InsurancePremium"] = df["InsurancePremium"].fillna(df["InsurancePremium_ref"])

# remove extra columns
df.drop(columns=["InsurancePremium_ref"], inplace=True)

# fill remaining empty with 0
df["InsurancePremium"] = df["InsurancePremium"].fillna(0)

print(len(df))
print(df["MainCode"].nunique())
print(df["MainCode"].duplicated().sum())
df.head(2)

age = df['AgeingDays']

conditions1 = [
    age.isna() | (age == 0),
    age <= 30,
    age <= 60,
    age <= 90,
    age <= 120,
    age <= 180,
    age <= 365,
    age > 365
]

choices1 = [
    "Regular",
    "1-30 Days",
    "31-60 Days",
    "61-90 Days",
    "91-120 Days",
    "121-180 Days",
    "181-365 Days",
    "Above 365 Days"
]

df['Ageing'] = np.select(conditions1, choices1, default="Unknown")

conditions2 = [
    age.isna() | (age == 0),
    age <= 90,
    age <= 180,
    age <= 365,
    age > 365
]

choices2 = [
    "Regular",
    "1-90 Days",
    "91-180 Days",
    "181-365 Days",
    "Above 365 Days"
]

df['Bucket'] = np.select(conditions2, choices2, default="Unknown")

cols = ['OutstandingBaln','IntDrAmt','PenalIntAmt','IntOnInt',
        'OvDuePrin','PastDuedInt','TotCharge']
# df[cols].dtypes
df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')
df[cols].dtypes

df['TotOvDue'] = np.where(
    df['Remarks'] == "Expired",
    -df['OutstandingBaln'] + df['IntDrAmt'] + df['TotCharge'],
    df['PenalIntAmt'] + df['IntOnInt'] + df['OvDuePrin'] + df['PastDuedInt'] + df['TotCharge']
)

df['OvDueWithInsurance'] = df['TotOvDue'] + df['InsurancePremium']

keywords = ["sold out", "court case"]

df.loc[
    df["OfficerName"].str.contains("|".join(keywords), case=False, na=False),
    "Bucket"
] = "Above 365 Days"

# df.to_excel(output_file, index=False, sheet_name="Mainsheet")

cols = ["OfficerName", "Loan Type", "Dealer Name"]

# Prepare reference table: one row per combination
df_ref = duelist_main[["AcTypeDesc", "BranchName"] + cols].drop_duplicates(
    subset=["AcTypeDesc", "BranchName"], keep="last"
)

# Merge with df
final_df = df.merge(
    df_ref,
    on=["AcTypeDesc", "BranchName"],
    how="left",
    suffixes=("", "_ref")
)

# Fill missing values
for col in cols:
    final_df[col] = final_df[col].astype(str).str.strip().replace(r'^(|None|nan|\s+)$', pd.NA, regex=True)
    final_df[col] = final_df[col].fillna(final_df[f"{col}_ref"])

# Drop temporary reference columns
final_df.drop(columns=[f"{col}_ref" for col in cols], inplace=True)

final_df.to_excel(output_file, index=False, sheet_name="Mainsheet")

input("Press Enter to exit...")