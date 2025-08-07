import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

def clean_amount(amount):
    try:
        return float(amount.replace("BRL", "").replace(",", "").replace("\xa0", "").strip())
    except:
        return None

def map_payroll_to_batch(payroll_df, template_df):
    output_rows = []

    grouped = payroll_df.groupby(['Email', 'Last name (legal)', 'First name (legal)', 'Currency'])

    for (email, last, first, currency), group in grouped:
        total_amount = group['Amount'].apply(clean_amount).sum()
        full_name = f"{first} {last}"

        # Define transfer details based on currency or custom rules
        if currency == 'BRL':
            country = 'Brazil'
            method = 'SWIFT'
        elif currency == 'PKR':
            country = 'Pakistan'
            method = 'SWIFT'
        elif currency == 'USD':
            country = 'United States of America'
            method = 'ACH'
        elif currency == 'THB':
            country = 'Thailand'
            method = 'SWIFT'
        else:
            continue  # Skip unsupported currencies

        row = {
            'Transfer to': country,
            'Transfer method': method,
            'Currency recipient gets': currency,
            'Transfer amount in currency recipient gets': total_amount,
            'Currency you pay': 'AUD',
            'SWIFT fee option': 'OUR' if method == 'SWIFT' else '',
            'Fee paid by': 'Payer',
            'Account name': full_name,
            'Transfer purpose': 'Payroll',
            'Reference': f"Payroll - {first} {last}",
            'Recipient type': 'Business',
            'Country / region': country
        }
        output_rows.append(row)

    output_df = pd.DataFrame(output_rows)
    for col in template_df.columns:
        if col not in output_df.columns:
            output_df[col] = ''

    return output_df[template_df.columns]

st.title("Payroll to Batch Transfer Converter")
st.write("Upload your payroll file (.xlsx) and download the Airwallex batch transfer template.")

payroll_file = st.file_uploader("Upload Payroll Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Batch Transfer Template (Airwallex)", type=["xlsx"])

if payroll_file and template_file:
    payroll_xlsx = pd.ExcelFile(payroll_file)
    template_xlsx = pd.ExcelFile(template_file)

    try:
        payroll_df = payroll_xlsx.parse('Salary data July')
        template_df = template_xlsx.parse('Airwallex batch transfer')
    except Exception as e:
        st.error(f"Error reading files: {e}")
    else:
        result_df = map_payroll_to_batch(payroll_df, template_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name='Airwallex batch transfer', index=False)
        st.success("Conversion complete! Download your batch file below.")
        st.download_button(
            label="Download Batch Transfer File",
            data=output.getvalue(),
            file_name="batch_transfer_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
