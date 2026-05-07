import pandas as pd

# === File paths ===
input_file = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\CustomGeneralLedger_Master.xlsx"
output_file = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\CustomGeneralLedger_Master_formatted.xlsx"
cashflow_input_file = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\CashFlowStatement.xlsx"
cashflow_output_file = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\CashFlowStatement_formatted.xlsx"
Debtors_Creditors_Balances_input_file = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Debtors_Creditors_Balances.xlsx"
Debtors_Creditors_Balances_output_file = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Debtors_Creditors_Balances_formatted.xlsx"


# === Step 1: Read Excel file, skipping first 6 header rows ===
df = pd.read_excel(input_file, skiprows=6)

# === Step 2: Trim data until the first occurrence of '400000 - TRADE DEBTORS' ===
stop_index = None
for i, row in df.iterrows():
    if row.astype(str).str.contains("400000 - TRADE DEBTORS", case=False, na=False).any():
        stop_index = i
        break

if stop_index is not None:
    df = df.iloc[:stop_index]

# === Step 3: Remove '500000 - CASH AND BANK BALANCES' and fill Account column ===
if "Account" in df.columns:
    df = df[~df["Account"].astype(str).str.contains("500000 - CASH AND BANK BALANCES", case=False, na=False)]
    df["Account"] = df["Account"].fillna(method="ffill")

# === Step 4: Create new column 'transaction_amount_GL' after 'Transaction Number' ===
# Ensure numeric conversion for Debit and Credit
for col in ["Debit", "Credit"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# Initialize the new column
df["transaction_amount_GL"] = 0.0

# Apply logic based on 'Type'
if "Type" in df.columns:
    df.loc[df["Type"].str.lower() == "journal", "transaction_amount_GL"] = df["Debit"] - df["Credit"]
    df.loc[df["Type"].str.lower() == "payment", "transaction_amount_GL"] = df["Debit"]
    df.loc[df["Type"].str.lower() == "bill payment", "transaction_amount_GL"] = df["Credit"]
    df.loc[df["Type"].str.lower() == "customer refund", "transaction_amount_GL"] = df["Credit"]
    df.loc[df["Type"].str.lower() == "transfer", "transaction_amount_GL"] = df["Debit"] - df["Credit"]
    df.loc[df["Type"].str.lower() == "currency revaluation", "transaction_amount_GL"] = df["Debit"] - df["Credit"]

# Insert the column right after 'Transaction Number' (if it exists)
if "Transaction Number" in df.columns:
    cols = list(df.columns)
    tn_index = cols.index("Transaction Number")
    cols.insert(tn_index + 1, cols.pop(cols.index("transaction_amount_GL")))
    df = df[cols]

# === Step 5: Save cleaned data ===
df.to_excel(output_file, index=False)

print(f"✅ Cleaned data saved to {output_file}")
print("💡 Added 'transaction_amount_GL' column with type-based logic applied.")


import pandas as pd

#cashflow_input_file = r"/mnt/data/CashFlowStatement547.xls"   # your uploaded file

# === Step 1: Read the raw Excel file ====
df = pd.read_excel(cashflow_input_file, header=None)

# === Step 2: Extract month name (line after "Cash Flow Statement") ===
month_value = None
for i in range(len(df)):
    row_text = " ".join(df.loc[i].astype(str))
    if "cash flow statement" in row_text.lower():
        # Next row after the title contains the month name
        next_row = df.loc[i + 1].dropna().astype(str)
        month_value = next_row.iloc[0]
        break

if not month_value:
    month_value = "Unknown"

# === Step 3: Find row where column 0 = Financial Row ===
financial_row_idx = df[df[0].astype(str).str.contains("Financial Row", case=False, na=False)].index

if len(financial_row_idx) == 0:
    raise ValueError("Could not find 'Financial Row' in column A.")
    
financial_row_idx = financial_row_idx[0]

# Extract company names from columns B onward
company_names = df.loc[financial_row_idx, 1:].tolist()
company_names = [str(c).strip() for c in company_names]

# === Step 4: Find row containing "Cash at End of Period" ===
cash_row_idx = df[df[0].astype(str).str.contains("Cash at End of Period", case=False, na=False)].index

if len(cash_row_idx) == 0:
    raise ValueError("Could not find 'Cash at End of Period' row.")

cash_row_idx = cash_row_idx[0]

# Extract cash values for each company (columns B onward)
cash_values = df.loc[cash_row_idx, 1:].tolist()

# === Step 5: Build output table ===
output = []
for company, cash in zip(company_names, cash_values):
    output.append({
        "Company_Name": company,
        "Month": month_value,
        "Cash_At_End_Of_Period": cash
    })

result_df = pd.DataFrame(output)
result_df.to_excel(cashflow_output_file, index=False)

#print(result_df)

##

import pandas as pd

# Read Excel file, skip first 9 rows

#df = pd.read_excel(cashflow_input_file, header=None)
df = pd.read_excel(
    Debtors_Creditors_Balances_input_file,
    skiprows=9,
    header=None
)

# Assign column names (first 3 columns)
df.columns = ["party_name", "balance", "subsidiary_name"]

# Create Account column
df["Account"] = None

current_account = None

for idx, value in df["party_name"].astype(str).items():

    val = value.strip()

    # Start Account 400000
    if val.startswith("400000"):
        current_account = "400000"

    # Stop Account 400000
    elif val.startswith("Total - 400000"):
        current_account = None

    # Start Account 200000
    elif val.startswith("200000"):
        current_account = "200000"

    # Stop Account 200000
    elif val.startswith("Total - 200000"):
        current_account = None

# Start Account 16000A
    elif val.startswith("16000A"):
        current_account = "16000A"

    # Stop Account 16000A
    elif val.startswith("Total - 16000A"):
        current_account = None

    # Start Account 16000B
    elif val.startswith("16000B"):
        current_account = "16000B"

    # Stop Account 16000B
    elif val.startswith("Total - 16000B"):
        current_account = None
# Start Account 160300
    elif val.startswith("160300"):
        current_account = "160300"

    # Stop Account 160300
    elif val.startswith("Total - 160300"):
        current_account = None

    # Start Account 160400
    elif val.startswith("160400"):
        current_account = "160400"

    # Stop Account 160400
    elif val.startswith("Total - 160400"):
        current_account = None
# Start Account 400999
    elif val.startswith("400999"):
        current_account = "400999"

    # Stop Account 400999
    elif val.startswith("Total - 400999"):
        current_account = None

    # Start Account 600402
    elif val.startswith("600402"):
        current_account = "600402"

    # Stop Account 600402
    elif val.startswith("Total - 600402"):
        current_account = None


df.at[idx, "Account"] = current_account
#current_account = None

ACCOUNT_CODES = [
    "400000",
    "200000",
    "16000A",
    "16000B",
    "160300",
    "160400",
    "400999",
    "600402"
]

accounts = []
current_account = None

for val in df["party_name"].astype(str):

    if current_account and val.startswith(f"Total - {current_account}"):
        current_account = None

    for acc in ACCOUNT_CODES:
        if val.startswith(acc):
            current_account = acc
            break

    accounts.append(current_account)

df["Account"] = accounts



df["party_name"] = df["party_name"].astype(str).str.strip()
df.loc[df["party_name"].isin(["", "nan", "None"]), "party_name"] = pd.NA


current_party = None

for idx, value in df["party_name"].items():

    if pd.notna(value) and value.startswith("Total"):
        current_party = None
        continue

    if pd.notna(value):
        current_party = value
    else:
        df.at[idx, "party_name"] = current_party



df["Period"] = "Dec 2022"

df.to_excel(Debtors_Creditors_Balances_output_file, index=False)




