import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

st.set_page_config(page_title="EMS vs Payroll Validator", layout="wide")

st.title("EMS vs Payroll Hours Validation Tool")

ems_file = st.file_uploader("Upload EMS Monitoring Sheet", type=["xlsx"])
payroll_file = st.file_uploader("Upload Payroll Workbook", type=["xlsx"])


# -------- NAME CLEANING FUNCTION --------
def clean_name(name):
    name = str(name).lower()
    name = re.sub(r"[^\w\s]", "", name)
    name = name.replace(",", "")

    parts = name.split()

    if len(parts) >= 2:
        return parts[0] + parts[1]

    return name


if ems_file and payroll_file:

    # ================= EMS FILE =================
    ems_df = pd.read_excel(ems_file, header=[2, 3])

    planned = ems_df[("Planned", "Duration")]
    actual = ems_df[("Actual", "Duration")]
    employee = ems_df[("Actual", "Employee")]

    calc_df = pd.DataFrame({
        "Employee": employee,
        "Planned": planned,
        "Actual": actual
    })

    calc_df["Chosen Hours"] = calc_df[["Planned", "Actual"]].min(axis=1)

    ems_hours = (
        calc_df.groupby("Employee")["Chosen Hours"]
        .sum()
        .reset_index()
    )

    ems_hours.rename(columns={"Chosen Hours": "EMS Hours"}, inplace=True)

    ems_hours["key"] = ems_hours["Employee"].apply(clean_name)

    st.subheader("EMS Calculated Hours")
    st.dataframe(ems_hours)


    # ================= PAYROLL FILE =================
    xl = pd.ExcelFile(payroll_file)

    payroll_records = []
    last_employee = None  # <-- IMPORTANT for continuation sheets

    for sheet in xl.sheet_names:

        # Skip summary sheets
        if "total" in sheet.lower():
            continue

        df = xl.parse(sheet, header=None)

        employee_name = None

        # -------- FIND EMPLOYEE NAME --------
        for r in range(len(df)):
            for c in range(len(df.columns)):

                if "Employee Address" in str(df.iloc[r, c]):

                    name_text = str(df.iloc[r + 1, c]).strip()

                    parts = name_text.split()
                    titles = ["mr", "mrs", "ms", "miss"]

                    if parts and parts[0].lower().replace(".", "") in titles:
                        parts = parts[1:]

                    if len(parts) >= 2:
                        first = parts[0]
                        last = parts[1]
                        employee_name = f"{last}, {first}"

                    break

            if employee_name:
                break

        # -------- HANDLE CONTINUATION SHEETS --------
        if employee_name:
            last_employee = employee_name
        else:
            employee_name = last_employee

        # If still no employee, skip
        if not employee_name:
            continue

        # -------- SKIP IF CANCELLATION SECTION EXISTS --------
        cancellation_found = False

        for r in range(len(df)):
            if "cancellation" in str(df.iloc[r, 0]).lower():
                cancellation_found = True
                break

        if cancellation_found:
            continue

        # -------- FIND SERVICE DETAIL --------
        service_row = None

        for r in range(len(df)):
            if "Service Detail" in str(df.iloc[r, 0]):
                service_row = r
                break

        if service_row is None:
            continue

        header_row = service_row + 1
        headers = df.iloc[header_row]

        hours_col = None

        for i, val in enumerate(headers):
            if "hour" in str(val).lower():
                hours_col = i
                break

        if hours_col is None:
            continue

        # -------- EXTRACT HOURS --------
        hours = pd.to_numeric(
            df.iloc[header_row + 1:, hours_col],
            errors="coerce"
        )

        total_hours = hours.sum()

        payroll_records.append({
            "Employee": employee_name,
            "Payroll Hours": total_hours
        })


    # -------- CREATE PAYROLL DF --------
    payroll_df = pd.DataFrame(payroll_records)

    if not payroll_df.empty:
        payroll_df = payroll_df.groupby("Employee", as_index=False)["Payroll Hours"].sum()
        payroll_df["key"] = payroll_df["Employee"].apply(clean_name)

    st.subheader("Detected Payroll Hours")
    st.dataframe(payroll_df)


    # ================= MERGE =================
    if payroll_df.empty:

        st.error("No payroll records detected")

        result = ems_hours.copy()
        result["Payroll Hours"] = np.nan

    else:

        result = pd.merge(
            ems_hours,
            payroll_df,
            on="key",
            how="left",
            suffixes=("_EMS", "_Payroll")
        )

    result["Difference"] = result["EMS Hours"] - result["Payroll Hours"]

    result["Match"] = np.where(
        abs(result["Difference"]) < 0.01,
        "MATCH",
        "MISMATCH"
    )

    result = result.rename(columns={
        "Employee_EMS": "Employee"
    })

    result = result[[
        "Employee",
        "EMS Hours",
        "Payroll Hours",
        "Difference",
        "Match"
    ]]

    st.subheader("Validation Result")
    st.dataframe(result)


    # ================= EXPORT =================
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Validation Report", index=False)

    output.seek(0)

    st.download_button(
        label="Download Validation Report",
        data=output,
        file_name="Payroll_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )