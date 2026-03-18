import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Duelist Processor", layout="wide")

st.title("🏦 Duelist Data Processing Tool")

st.markdown("Upload the required files to generate the **Updated Duelist Report**.")

col1, col2, col3 = st.columns(3)

with col1:
    duelist_file = st.file_uploader("Upload Duelist Dump", type=["xlsx","xlsb"])

with col2:
    duelist_main_file = st.file_uploader("Upload Duelist Main File", type=["xlsx","xlsb"])

with col3:
    insurance_file = st.file_uploader("Upload Insurance File", type=["xlsx"])

process = st.button("🚀 Process Files")

if process:

    if duelist_file is None or duelist_main_file is None or insurance_file is None:
        st.error("Please upload all three files.")
        st.stop()

    with st.spinner("Processing data..."):

        # ==========================
        # READ FILES
        # ==========================

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

        # ==========================
        # CLEAN COLUMNS
        # ==========================

        def clean_columns(df):
            df.columns = df.columns.str.strip()
            return df

        duelist = clean_columns(duelist)
        duelist_main = clean_columns(duelist_main)
        insurance = clean_columns(insurance)

        # ==========================
        # INSURANCE LOGIC
        # ==========================

        insurance["Date"] = pd.to_datetime(insurance["Date"])

        yesterday = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)

        insurance["InsurancePremium"] = np.where(
            yesterday < insurance["Date"],
            insurance["InsPremium"],
            0
        )

        insurance_sum = insurance.groupby("MainCode")["InsurancePremium"].sum().reset_index()

        # ==========================
        # FILTER
        # ==========================

        duelist = duelist.copy()
        duelist = duelist[(duelist['ClientCode']!="~~~~~")]

        # ==========================
        # REMOVE DUPLICATES
        # ==========================

        duelist = duelist.drop_duplicates(subset="MainCode", keep="last")
        duelist_main = duelist_main.drop_duplicates(subset="MainCode", keep="last")
        insurance_sum = insurance_sum.drop_duplicates(subset="MainCode", keep="last")

        # ==========================
        # AGEING DAYS
        # ==========================

        duelist['AgeingDays'] = pd.to_numeric(duelist['AgeingDays'], errors='coerce')
        duelist = duelist.sort_values(by='AgeingDays', ascending=False)

        # ==========================
        # INSERT COLUMNS
        # ==========================

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

        # ==========================
        # INSURANCE MERGE
        # ==========================

        df = df.merge(
            insurance_sum[["MainCode", "InsurancePremium"]],
            on="MainCode",
            how="left",
            suffixes=("", "_ref")
        )

        df["InsurancePremium"] = df["InsurancePremium"].fillna(df["InsurancePremium_ref"])
        df.drop(columns=["InsurancePremium_ref"], inplace=True)
        df["InsurancePremium"] = df["InsurancePremium"].fillna(0)

        # ==========================
        # AGEING BUCKET
        # ==========================

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

        # ==========================
        # NUMERIC COLUMNS
        # ==========================

        cols = ['OutstandingBaln','IntDrAmt','PenalIntAmt','IntOnInt',
                'OvDuePrin','PastDuedInt','TotCharge']

        df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')

        # ==========================
        # TOTAL OVERDUE
        # ==========================

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

        final_df = df.copy()

        # cols = ["OfficerName", "Loan Type", "Dealer Name"]

        # df_ref = duelist_main[["AcTypeDesc", "BranchName"] + cols].drop_duplicates(
        #     subset=["AcTypeDesc", "BranchName"], keep="last"
        # )

        # final_df = df.merge(
        #     df_ref,
        #     on=["AcTypeDesc", "BranchName"],
        #     how="left",
        #     suffixes=("", "_ref")
        # )

        # for col in cols:
        #     final_df[col] = final_df[col].astype(str).str.strip().replace(
        #         r'^(|None|nan|\s+)$', pd.NA, regex=True
        #     )
        #     final_df[col] = final_df[col].fillna(final_df[f"{col}_ref"])

        # final_df.drop(columns=[f"{col}_ref" for col in cols], inplace=True)

        # # Merge only on BranchName for remaining missing values ---
        # df_ref_branch = duelist_main[["BranchName"] + cols].drop_duplicates(subset=["BranchName"], keep="last")

        # final_df = final_df.merge(
        #     df_ref_branch,
        #     on="BranchName",
        #     how="left",
        #     suffixes=("", "_branch")
        # )

        # # Fill missing values from BranchName merge
        # for col in cols:
        #     final_df[col] = final_df[col].fillna(final_df[f"{col}_branch"])

        # # Drop temporary branch reference columns
        # final_df.drop(columns=[f"{col}_branch" for col in cols], inplace=True)

        # =====================================================
        # SHOW DATA
        # =====================================================
        tab1, tab2 = st.tabs(["📊 Processed Data", "⚠ Unmatched"])

        with tab1:
            st.dataframe(final_df, use_container_width=True)

        unmatched = final_df[final_df["Loan Type"].isna()]

        with tab2:
            st.write(f"Unmatched rows: {len(unmatched)}")
            st.dataframe(unmatched, use_container_width=True)

        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mainsheet")

            buffer.seek(0)

        st.download_button(
            label="📥 Download Updated Duelist File",
            data=buffer,
            file_name="updated_duelist.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ Processing Completed")