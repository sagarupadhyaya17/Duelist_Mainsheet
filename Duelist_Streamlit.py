
import streamlit as st
import pandas as pd
import numpy as np

st.title("Duelist Data Processor")

st.write("Upload the required files to process the Duelist report.")

duelist_file = st.file_uploader("Upload Duelist Dump File", type=["xlsx","xlsb"])
duelist_main_file = st.file_uploader("Upload Duelist Main File", type=["xlsx","xlsb"])
insurance_file = st.file_uploader("Upload Insurance File", type=["xlsx"])

if duelist_file and duelist_main_file and insurance_file:

    st.success("Files uploaded successfully. Processing...")

    # ==== READ FILES ====

    insurance = pd.read_excel(insurance_file, dtype={"MainCode": str})

    if duelist_file.name.endswith(".xlsb"):
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

    if duelist_main_file.name.endswith(".xlsb"):
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

    # ==== CLEAN COLUMNS ====

    def clean_columns(df):
        df.columns = df.columns.str.strip()
        return df

    duelist = clean_columns(duelist)
    duelist_main = clean_columns(duelist_main)
    insurance = clean_columns(insurance)

    # ==== INSURANCE LOGIC ====

    insurance["date"] = pd.to_datetime(insurance["date"])

    yesterday = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)

    insurance["InsurancePremium"] = np.where(
        yesterday < insurance["date"],
        insurance["InsPremium"],
        0
    )

    # ==== FILTER CLIENT CODE ====

    duelist = duelist.copy()
    duelist = duelist[(duelist['ClientCode']!="~~~~~")]

    # ==== REMOVE DUPLICATES ====

    duelist = duelist.drop_duplicates(subset="MainCode", keep="last")
    duelist_main = duelist_main.drop_duplicates(subset="MainCode", keep="last")
    insurance = insurance.drop_duplicates(subset="MainCode", keep="last")

    # ==== AGEING DAYS CLEAN ====

    duelist['AgeingDays'] = pd.to_numeric(duelist['AgeingDays'], errors='coerce')

    duelist = duelist.sort_values(by='AgeingDays', ascending=False)

    # ==== INSERT COLUMNS ====

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

    # ==== CLEAN MAINCODE ====

    duelist.columns = duelist.columns.str.strip()
    duelist_main.columns = duelist_main.columns.str.strip()

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

    # ==== MERGE INSURANCE ====

    df = df.merge(
        insurance[["MainCode", "InsurancePremium"]],
        on="MainCode",
        how="left",
        suffixes=("", "_ref")
    )

    df["InsurancePremium"] = df["InsurancePremium"].fillna(df["InsurancePremium_ref"])

    df.drop(columns=["InsurancePremium_ref"], inplace=True)

    df["InsurancePremium"] = df["InsurancePremium"].fillna(0)

    # ==== AGEING BUCKETS ====

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

    # ==== NUMERIC COLUMNS ====

    cols = ['OutstandingBaln','IntDrAmt','PenalIntAmt','IntOnInt',
            'OvDuePrin','PastDuedInt','TotCharge']

    df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')

    # ==== TOTAL OVERDUE ====

    df['TotOvDue'] = np.where(
        df['Remarks'] == "Expired",
        -df['OutstandingBaln'] + df['IntDrAmt'] + df['TotCharge'],
        df['PenalIntAmt'] + df['IntOnInt'] + df['OvDuePrin'] + df['PastDuedInt'] + df['TotCharge']
    )

    df['OvDueWithInsurance'] = df['TotOvDue'] + df['InsurancePremium']

    # ==== KEYWORD BUCKET LOGIC ====

    keywords = ["sold out", "court case"]

    df.loc[
        df["OfficerName"].str.contains("|".join(keywords), case=False, na=False),
        "Bucket"
    ] = "Above 365 Days"

    # ==== SECOND MERGE ====

    cols = ["OfficerName", "Loan Type", "Dealer Name"]

    df_ref = duelist_main[["AcTypeDesc", "BranchName"] + cols].drop_duplicates(
        subset=["AcTypeDesc", "BranchName"], keep="last"
    )

    final_df = df.merge(
        df_ref,
        on=["AcTypeDesc", "BranchName"],
        how="left",
        suffixes=("", "_ref")
    )

    for col in cols:
        final_df[col] = final_df[col].astype(str).str.strip().replace(
            r'^(|None|nan|\s+)$', pd.NA, regex=True
        )
        final_df[col] = final_df[col].fillna(final_df[f"{col}_ref"])

    final_df.drop(columns=[f"{col}_ref" for col in cols], inplace=True)

    # ==== DOWNLOAD OUTPUT ====

    output = final_df.to_excel(index=False, engine='openpyxl')

    st.success("Processing completed!")

    st.dataframe(final_df.head(20))

    st.download_button(
        label="Download Processed File",
        data=final_df.to_excel(index=False, engine="openpyxl"),
        file_name="updated_duelist.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )