import pandas as pd
import sys
import os
import argparse

if len(sys.argv) < 2:
    raise ValueError("Please pass run parameter. Example: python script.py Batch2")

run_param = sys.argv[1]
batch_input = run_param.strip().lower()
print(f"Running for parameter: {batch_input}")


from pathlib import Path

BASE_DIR_INBOUND = Path(r"C:\Broad_field_holdings\Net_suite\JR\Inbound")
BASE_DIR_INBOUND1 = Path(r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Merge_GL")
BASE_DIR_OUTBOUND = Path(r"C:\Broad_field_holdings\Net_suite\JR\Outbound")

#input_file = BASE_DIR / "Outbound" / f"Trade_creditors_Debitors_{run_param}.csv"
#balance_file = BASE_DIR / "Inbound" / "Merge_GL" / f"Debtors_Creditors_Balances_formatted_{run_param}.xlsx"


# === File paths ===
ledger_file = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Merge_GL\CustomGeneralLedger_Master_formatted.xlsx"


#bills_file = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Merge_GL\Bill_ETL_Transaction_Master_file_transformed.xlsx"
bills_file = os.path.join(
    BASE_DIR_INBOUND1,
    f"Bill_ETL_Transaction_Master_file_transformed_{batch_input}.xlsx"
)

#output_file = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Bill_ETL_file_updated_balance.csv"

output_file = os.path.join(
    BASE_DIR_OUTBOUND,
    f"Bill_ETL_file_updated_balance_{batch_input}.csv"
)

#trade_output_file = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Trade_creditors_Debitors.csv"

trade_output_file = os.path.join(
    BASE_DIR_OUTBOUND,
    f"Trade_creditors_Debitors_{batch_input}.csv"
)

#cashflow_output_file = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Aggegated_cashflow_categerization_file.csv"

cashflow_output_file = os.path.join(
    BASE_DIR_OUTBOUND,
    f"Aggegated_cashflow_categerization_file_{batch_input}.csv"
)

lookup_file_cashflowclosing = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Merge_GL\CashFlowStatement_formatted.xlsx"
Debtors_Creditors_Balances_formatted = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Merge_GL\Debtors_Creditors_Balances_formatted.xlsx"
#output_file_partywise = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Partywise_Balances_Monthly.csv"

output_file_partywise = os.path.join(
    BASE_DIR_OUTBOUND,
    f"Partywise_Balances_Monthly_{batch_input}.csv"
)




# === Step 1: Read both Excel files ====
ledger_df = pd.read_excel(ledger_file)
bills_df = pd.read_excel(bills_file)

# === Step 2: Clean column names ===
ledger_df.columns = ledger_df.columns.str.strip()
bills_df.columns = bills_df.columns.str.strip()

# === Step 3: Prepare cleaned versions for lookup ===
bills_df["Payment_Account_clean"] = (
    bills_df["Payment_Account"]
    .astype(str)
    .str.replace("CASH AND BANK BALANCES :", "-", regex=False)
    .str.replace("BANK RECONCILIATION CONTROL ACCOUNT :", "-", regex=False)
    .str.strip()
)


ledger_df["Account_clean"] = ledger_df["Account"].astype(str).str.strip()

# === Step 4: Create composite key in ledger ===
ledger_df["lookup_key"] = (
    ledger_df["Transaction Number"].astype(str) + "|" + ledger_df["Account_clean"]
)

# === Step 5: Create lookup dictionaries ===
lookup_balance = (
    ledger_df.drop_duplicates(subset=["lookup_key"])
    .set_index("lookup_key")["Balance"]
    .to_dict()
)
lookup_transaction_amt = (
    ledger_df.drop_duplicates(subset=["lookup_key"])
    .set_index("lookup_key")["transaction_amount_GL"]
    .to_dict()
)

# === Step 6: Perform lookup only for key1 == 'bankflow' and non-null Vendor_payment_number & Payment_Account ===
mask = (
    (bills_df["key1"] == "bankflow") &
    bills_df["Vendor_payment_number"].notna() &
    bills_df["Payment_Account"].notna()
)

bills_df.loc[mask, "lookup_key"] = (
    bills_df.loc[mask, "Vendor_payment_number"].astype(str)
    + "|"
    + bills_df.loc[mask, "Payment_Account_clean"]
)

# Perform lookup for filtered records only
bills_df.loc[mask, "Balance_from_GL"] = bills_df.loc[mask, "lookup_key"].map(lookup_balance)
bills_df.loc[mask, "transaction_amount_GL"] = bills_df.loc[mask, "lookup_key"].map(lookup_transaction_amt)

# === Step 7: Drop temporary columns before saving ===
bills_df.drop(columns=["lookup_key", "Payment_Account_clean"], inplace=True, errors="ignore")


# === Step 8: Add 'bankflow_final' column ===
print("🏦 Creating 'bankflow_final' flag for eligible records...")

# Insert the new column right after 'key1'
if "key1" in bills_df.columns:
    key1_index = bills_df.columns.get_loc("key1")
    bills_df.insert(key1_index + 1, "bankflow_final", "")

    # Define eligible rows
    mask_bankflow_final = (
        bills_df["Type"].astype(str).str.strip().isin(["Bill", "Expense Report", "Journal", "Invoice", "Customer Refund", "Transfer", "Currency Revaluation"]) &
        (bills_df["key"].astype(str).str.strip().str.lower() == "bankflow") &
        (bills_df["key1"].astype(str).str.strip().str.lower() == "bankflow")
    )

    bills_df.loc[mask_bankflow_final, "bankflow_final"] = "YES"
    print(f"✅ Marked {mask_bankflow_final.sum()} records as 'Yes' in 'bankflow_final'.")
else:
    print("⚠️ Column 'key1' not found — skipping 'bankflow_final' creation.")


# === Step 9: Add 'remarks' column for missing Invoices ===

# Ensure the column exists (in case you rerun)
if "remarks" not in bills_df.columns:
    bills_df["remarks"] = ""

# Get all Vendor_payment_number values for Invoice records
invoice_vendors = set(
    bills_df.loc[
        bills_df["Type"].astype(str).str.lower() == "invoice",
        "Vendor_payment_number"
    ].dropna().astype(str)
)

# Identify Payment records missing a matching Invoice
mask_payment_no_invoice = (
    (bills_df["Type"].astype(str).str.lower() == "payment") &
    #(bills_df["bankflow_final"].astype(str).str.strip().str.lower() == "yes") &
    (~bills_df["Vendor_Bill_number"].astype(str).isin(invoice_vendors))
)

# Update remarks for those
bills_df.loc[mask_payment_no_invoice, "remarks"] = "Invoice not found"

print(f"🧾 Added remarks for {mask_payment_no_invoice.sum()} Payment records with no matching Invoice.")

# === Additional logic: Clear 'Invoice not found' if matching Journal exists ===

# Collect Vendor_payment_number values from Journal records
journal_vendors = set(
    bills_df.loc[
        bills_df["Type"].astype(str).str.strip().str.lower() == "journal",
        "Vendor_payment_number"
    ].dropna().astype(str)
)

# Identify Payment records with 'Invoice not found' but having matching Journal
mask_clear_remarks = (
    (bills_df["Type"].astype(str).str.strip().str.lower() == "payment") &
    (bills_df["remarks"].astype(str).str.strip().str.lower() == "invoice not found") &
    (bills_df["Vendor_payment_number"].astype(str).isin(journal_vendors))
)

# Clear remarks
bills_df.loc[mask_clear_remarks, "remarks"] = ""

print(f"🧹 Cleared remarks for {mask_clear_remarks.sum()} Payment records due to matching Journal.")







# === Additional logic for Payment with Invoice not found for bankflow_final field ===

# === Additional logic for Payment with Invoice not found and Payment_Account like 500 ===

payment_mask = (
    bills_df["Type"].astype(str).str.strip().str.lower() == "payment"
) & (
    bills_df["remarks"].astype(str).str.strip().str.lower() == "invoice not found"
) & (
    bills_df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)
)

# Update the flag
bills_df.loc[payment_mask, "bankflow_final"] = "YES"






# === Step 10: Remove Journal records with Status = Pending Approval (FINAL STEP) ===
initial_count = len(bills_df)
bills_df = bills_df[
    ~(
        (bills_df["Type"].astype(str).str.lower() == "journal") &
        (bills_df["Status"].astype(str).str.lower().isin(["pending approval", "rejected"]))
    )
].copy()

removed_count = initial_count - len(bills_df)
print(f"🗑️ Removed {removed_count} Journal records with Status = Pending Approval before saving output.")

# === Step 9B: Add 'BS_PNL_Flag' column ===
print("🏷️ Creating 'BS_PNL_Flag' column for specific record types...")

# Insert the column right after 'Subcategory' if it exists
#if "Subcategory" in bills_df.columns:
   # subcat_index = bills_df.columns.get_loc("Subcategory")
   # bills_df.insert(subcat_index + 1, "BS_PNL_Flag", "")
#else:
    # If Subcategory column not found, append at end
    #bills_df["BS_PNL_Flag"] = ""

# Normalize Type and Remarks
bills_df["Type_clean"] = bills_df["Type"].astype(str).str.strip().str.title()
bills_df["remarks_clean"] = bills_df["remarks"].astype(str).str.strip().str.lower()

# Define eligible types (always mark YES)
always_yes_types = [
    "Bill", "Invoice", "Journal",
    "Expense Report", "Credit Memo", "Bill Credit"
]

# Create masks
mask_always_yes = bills_df["Type_clean"].isin(always_yes_types)
mask_payment_yes = (
    (bills_df["Type_clean"] == "Payment") &
    (bills_df["remarks_clean"] == "invoice not found")
)

# Apply logic
#bills_df.loc[mask_always_yes | mask_payment_yes, "BS_PNL_Flag"] = "YES"
bills_df.loc[mask_payment_yes, "BS_PNL_Flag_Final"] = "YES"

# --- Additional logic for specific subsidiary & expense account ---

mask_montage_payment = (
    (bills_df["Subsidiary_Name"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.contains("AGENSI PEKERJAAN MONTAGE HUMAN CAPITAL SDN BHD", na=False)
    ) &
    (bills_df["Type"]
        .astype(str)
        .str.strip()
        .str.lower() == "payment"
    ) &
    (bills_df["Expense_Account"]
        .astype(str)
        .str.contains("16000B", na=False)
    )
)

bills_df.loc[mask_montage_payment, "BS_PNL_Flag_Final"] = "YES"



# Drop helper columns
bills_df.drop(columns=["Type_clean", "remarks_clean"], inplace=True, errors="ignore")

#print(f"✅ Marked {(mask_always_yes | mask_payment_yes).sum()} records as 'YES' in BS_PNL_Flag.")

# === Additional logic for null Cashflow_Category_first ===
bills_df['Cashflow_Category_first'] = bills_df['Cashflow_Category_first'].fillna('').astype(str)

bills_df.loc[
    (bills_df['Cashflow_Category_first'].str.strip() == '') &
    (bills_df['bankflow_final'].astype(str).str.upper() == 'YES'),
    'Cashflow_Category_first'
] = 'Operating'


# === Step X: Add Cashflow_Type column based on amount_bankflow ===

# Ensure amount_bankflow is numeric
bills_df["amount_bankflow"] = pd.to_numeric(bills_df["amount_bankflow"], errors="coerce")

# Create new column
bills_df["Cashflow_inflowtype"] = ""
bills_df["Cashflow_outflowtype"] = ""
bills_df["Cashflow_type"] = ""

# Apply inflow / outflow logic
bills_df.loc[bills_df["amount_bankflow"] > 0, "Cashflow_inflowtype"] = "Inflow"
bills_df.loc[bills_df["amount_bankflow"] < 0, "Cashflow_outflowtype"] = "Outflow"
bills_df.loc[bills_df["amount_bankflow"] > 0, "Cashflow_type"] = "Inflow"
bills_df.loc[bills_df["amount_bankflow"] < 0, "Cashflow_type"] = "Outflow"


# === Step X: Move Cashflow_Category and Cashflow_Category_first to the end ===

cols = list(bills_df.columns)

# Columns to move
move_cols = ["Cashflow_Category", "Cashflow_Category_first", "cashflow_first_subcategory", "PNL_STMT_MAIN_CATEGORY" ,"PNL_STMT_CATEGORY", "PNL_STMT_SUBCATEGORY"]

# Keep only columns that actually exist (safety check)
move_cols = [c for c in move_cols if c in cols]

# Reorder: keep all other columns first
remaining_cols = [c for c in cols if c not in move_cols]

# Final column order
bills_df = bills_df[remaining_cols + move_cols]



# === Additional logic: Set cashflow_first_subcategory for missing invoices ===
mask_subcat_missing_invoice = (
    (bills_df["bankflow_final"].astype(str).str.upper() == "YES") &
    (bills_df["remarks"].astype(str).str.strip().str.lower() == "invoice not found")
)

bills_df.loc[mask_subcat_missing_invoice, "cashflow_first_subcategory"] = "INVOICE NOT FOUND"

# === Additional logic: Default missing subcategory to 'OTHERS' for bankflow_final = YES ===
mask_subcat_others = (
    (bills_df["bankflow_final"].astype(str).str.upper() == "YES") &
    (
        bills_df["cashflow_first_subcategory"].isna() |
        (bills_df["cashflow_first_subcategory"].astype(str).str.strip() == "")
    )
)

bills_df.loc[mask_subcat_others, "cashflow_first_subcategory"] = "OTHERS"

# === Additional logic: Map DEPRECIATION to OTHERS ===
mask_depreciation = (
    bills_df["cashflow_first_subcategory"]
    .astype(str)
    .str.strip()
    .str.upper()
    == "DEPRECIATION"
)

bills_df.loc[mask_depreciation, "cashflow_first_subcategory"] = "OTHERS"


# === Additional logic: Extract value after last ':' in Subsidiary_Name ===
#19-Feb-2026 logic fix for subscidary 
#bills_df["Subsidiary_Name"] = (
    #bills_df["Subsidiary_Name"]
    #.astype(str)
    #.str.rsplit(":", n=1)
    #.str[-1]
    #.str.strip()
#)

special_cases = [
    "Headquarters : G3 Healthcare Sdn Bhd (FKA:Bestinet Healthcare)",
    "Headquarters : Bio Clinic Sdn Bhd (FKA:Pengerang  Technology)"
]

def clean_subsidiary(name):
    name = str(name).strip()

    if name in special_cases:
        # Take value after FIRST colon only
        return name.split(":", 1)[1].strip()
    else:
        # Default behavior
        #return name.split(":", 1)[-1].strip()
        return name.split(":")[-1].strip()

bills_df["Subsidiary_Name"] = (
    bills_df["Subsidiary_Name"]
    .apply(clean_subsidiary)
)


# === Fix encoding issue for Krypton company name ===
#bills_df["Subsidiary_Name"] = bills_df["Subsidiary_Name"].replace(
    #"KryptonÂ Global Network (M) Sdn Bhd", ===KryptonÂ Global Network (M) Sdn Bhd
    #"Krypton Global Network (M) Sdn Bhd"
#)

bills_df["Subsidiary_Name"] = (
    bills_df["Subsidiary_Name"]
    .str.replace("\u00A0", " ", regex=False)   # non-breaking space
    .str.replace("Â", "", regex=False)         # stray encoding character
    .str.replace(r"\s+", " ", regex=True)      # normalize spaces
    .str.strip()
)



# === Additional logic: Align Paid_date with posting_period month ===

# === Additional logic: Align Paid_date with correct month (posting_period / period) ===

# Ensure Paid_date is datetime
bills_df["Paid_date"] = pd.to_datetime(bills_df["Paid_date"], errors="coerce")

# Convert posting_period and period to datetime
bills_df["posting_period_dt"] = pd.to_datetime(
    bills_df["posting_period"].astype(str),
    errors="coerce",
    format="mixed"
)

bills_df["period_dt"] = pd.to_datetime(
    bills_df["Period"].astype(str),
    errors="coerce",
    format="mixed"
)

# Extract Paid_date month
paid_ym = bills_df["Paid_date"].dt.to_period("M")

# Extract comparison month based on Type
compare_ym = pd.Series(index=bills_df.index, dtype="period[M]")

# For Journal → use period
# === Journal mask with additional condition: posting_period is NULL ===

mask_journal = (
    (bills_df["Type"].astype(str).str.strip().str.lower() == "journal") &
    (bills_df["posting_period"].isna() | 
     (bills_df["posting_period"].astype(str).str.strip() == ""))
)

compare_ym.loc[mask_journal] = bills_df.loc[mask_journal, "period_dt"].dt.to_period("M")

# For non-Journal → use posting_period
compare_ym.loc[~mask_journal] = bills_df.loc[~mask_journal, "posting_period_dt"].dt.to_period("M")

# Identify month mismatch (exclude Customer Refund)
mask_month_mismatch = (
    bills_df["Paid_date"].notna() &
    compare_ym.notna() &
    (paid_ym != compare_ym) &
    (bills_df["Type"].astype(str).str.strip().str.lower() != "customer refund")
)

# Update remarks
bills_df.loc[mask_month_mismatch, "remarks"] = (
    bills_df.loc[mask_month_mismatch, "Paid_date"]
    .dt.strftime("%Y-%m-%d")
    + " - Posting month is different from Paid_date"
)

# Update Paid_date to first date of correct month
bills_df.loc[mask_month_mismatch & mask_journal, "Paid_date"] = (
    bills_df.loc[mask_month_mismatch & mask_journal, "period_dt"]
    .dt.to_period("M")
    .dt.to_timestamp()
)

bills_df.loc[mask_month_mismatch & ~mask_journal, "Paid_date"] = (
    bills_df.loc[mask_month_mismatch & ~mask_journal, "posting_period_dt"]
    .dt.to_period("M")
    .dt.to_timestamp()
)

# Optional cleanup
bills_df.drop(columns=["posting_period_dt", "period_dt"], inplace=True, errors="ignore")

print(f"📅 Updated Paid_date and remarks for {mask_month_mismatch.sum()} records where posting period differs.")


# === Configurable list of Journal IDs requiring amount reversal ===
REVERSE_JOURNAL_IDS = [
    "JOURNAL_BT_011348",
    # "JOURNAL_BT_011349",
    # "JOURNAL_BT_011350",
]


# === Custom logic: Reverse amount_bankflow for configured Journal IDs ===

mask_reverse_amount = (
    (bills_df["Type"].astype(str).str.strip().str.lower() == "journal") &
    (bills_df["Vendor_Bill_number"].astype(str).str.strip().isin(REVERSE_JOURNAL_IDS)) &
    (bills_df["Vendor_payment_number"].astype(str).str.strip().isin(REVERSE_JOURNAL_IDS))
)

# Ensure amount is numeric and reverse
bills_df.loc[mask_reverse_amount, "amount_bankflow"] = (
    pd.to_numeric(
        bills_df.loc[mask_reverse_amount, "amount_bankflow"],
        errors="coerce"
    ) * -1
)

print(f"🔁 Reversed amount_bankflow for {mask_reverse_amount.sum()} journal record(s).")


## Exclusive Period & Date allignment logic for Besti 

# Ensure Date is datetime
bills_df["Date"] = pd.to_datetime(bills_df["Date"], errors="coerce")

# Convert Period to datetime
bills_df["period_dt"] = pd.to_datetime(
    bills_df["Period"].astype(str), errors="coerce", infer_datetime_format=True
)

# Apply only for Subsidiary_Name = "Besti"
mask_besti = bills_df["Subsidiary_Name"].astype(str).str.strip() == "Bestinet Sdn Bhd"

# Extract month from Date and Period
paid_ym = bills_df.loc[mask_besti, "Date"].dt.to_period("M")
period_ym = bills_df.loc[mask_besti, "period_dt"].dt.to_period("M")

# Identify month mismatch (exclude Customer Refund)
mask_month_mismatch = (
    paid_ym.notna()
    & period_ym.notna()
    & (paid_ym != period_ym)
    & (bills_df.loc[mask_besti, "Type"].astype(str).str.strip().str.lower() != "customer refund")
)

# Update remarks
bills_df.loc[mask_besti, "remarks"] = bills_df.loc[mask_besti, "remarks"].astype(str)  # ensure string
bills_df.loc[mask_besti & mask_month_mismatch, "remarks"] = (
    bills_df.loc[mask_besti & mask_month_mismatch, "Date"]
    .dt.strftime("%Y-%m-%d")
    + " - Posting month is different from Period"
)

# Update Date to first date of Period month
bills_df.loc[mask_besti & mask_month_mismatch, "Date"] = (
    bills_df.loc[mask_besti & mask_month_mismatch, "period_dt"]
    .dt.to_period("M")
    .dt.to_timestamp()
)

# Cleanup
bills_df.drop(columns=["period_dt"], inplace=True, errors="ignore")

# Apply only for Subsidiary_Name = Besti
# Normalize Period for reliable matching
period_norm = (
    bills_df["Period"]
    .astype(str)
    .str.lower()
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)

# Mask: Besti + Adjust 2024 + 12/31
mask_besti_adjust = (
    bills_df["Subsidiary_Name"].astype(str).str.strip().str.lower() == "Bestinet Sdn Bhd"
) & (
    period_norm.str.contains("adjust 2024", na=False)
) & (
    period_norm.str.contains("12/31", na=False)
)

# Update Period
bills_df.loc[mask_besti_adjust, "Period"] = "Dec 24"

mask_bestinet_payment_account = (
    bills_df["Subsidiary_Name"].astype(str).str.strip().str.lower() == "bestinet sdn bhd"
) & (
    bills_df["Type"].astype(str).str.strip().str.lower() == "payment"
) & (
    bills_df["Account"].astype(str).str.strip() == "600111"
)

bills_df.loc[mask_bestinet_payment_account, "BS_PNL_Flag_Final"] = "YES"

#19 Feb-2026 , handling AR Golden Way Sdn Bhd mismatches 

import numpy as np

# Normalize Subsidiary_Name (safety)
bills_df["Subsidiary_Name"] = (
    bills_df["Subsidiary_Name"]
    .astype(str)
    .str.strip()
)

# Condition commented on 2-March 2026 to include besti
cond = (
    (bills_df["Subsidiary_Name"] == "AR Golden Way Sdn Bhd") &
    (bills_df["amount_paid"].abs() == bills_df["transaction_amount_GL"].abs()) &
    (np.sign(bills_df["amount_paid"]) != np.sign(bills_df["transaction_amount_GL"]))
)

#cond = (
    #bills_df["Subsidiary_Name"].isin([
        #"AR Golden Way Sdn Bhd",
        #"Bestinet Sdn Bhd"
    #]) &
    #(bills_df["amount_paid"].abs() == bills_df["transaction_amount_GL"].abs()) &
    #(np.sign(bills_df["amount_paid"]) != np.sign(bills_df["transaction_amount_GL"]))
#)

# Update sign of amount_bankflow
bills_df.loc[cond, "amount_bankflow"] = (
    bills_df.loc[cond, "amount_bankflow"].abs() *
    np.sign(bills_df.loc[cond, "transaction_amount_GL"])
)

#19 Feb 2026 handling Greenfield Hills Sdn Bhd issues 
# Ensure date is in datetime format
bills_df["Paid_date"] = pd.to_datetime(bills_df["Paid_date"], errors="coerce")

cond = (
    (bills_df["Subsidiary_Name"] == "Greenfield Hills Sdn Bhd") &
    (bills_df["bankflow_final"] == "YES") &
    (bills_df["Paid_date"] == pd.Timestamp("2023-02-14")) &
    (bills_df["Payment_Account"] == "500551 GREENFIELD - RECONCILIATION CONTROL ACCOUNT")
)

# Update bankflow_final to null
bills_df.loc[cond, "bankflow_final"] = np.nan


#19 Feb 2026 handling Tass Tech (Malaysia) Sdn Bhd issues 
# Ensure date is in datetime format
bills_df["Paid_date"] = pd.to_datetime(bills_df["Paid_date"], errors="coerce")

cond = (
    (bills_df["Subsidiary_Name"] == "Tass Tech (Malaysia) Sdn Bhd") &
    (bills_df["bankflow_final"] == "YES") &
    (bills_df["Paid_date"] == pd.Timestamp("2023-03-07")) &
    (bills_df["Payment_Account"] == "500517 BANK RECONCILIATION CONTROL ACCOUNT : TASSTECH -RECONCILIATION CONTROL ACCOUNT")
)

# Update bankflow_final to null
bills_df.loc[cond, "bankflow_final"] = np.nan

bills_df["Subsidiary_Name"] = bills_df["Subsidiary_Name"].replace(
    "Hexacloud Sdn Bhd - New",
    "Hexacloud Sdn Bhd  - New"
)


# === Step 11: Save output ===
bills_df.to_csv(output_file, index=False)

# === Step 12: Print summary ===
matched = bills_df["Balance_from_GL"].notna().sum()
total = len(bills_df)
print(f"✅ Updated file saved: {output_file}")
print(f"💡 Matched {matched} out of {total} records.")

# Code for generating Trade creditors and debitors file 
# === Column definitions ===
set1_cols = [
    "Subsidiary_Name", "Date", "Type", "Document Number", "Name",
    "Expense_Account", "Amount", "Item", "TRANSACTION APPROVAL STATUS",
    "Status", "Vendor_Bill_number", "Vendor_payment_number", "source"
]

set2_cols = [
    "Subsidiary_Name", "Paid_date", "Type", "Document Number", "Name",
    "Expense_Account", "amount_bankflow", "Item", "TRANSACTION APPROVAL STATUS",
    "Status", "Vendor_Bill_number", "Vendor_payment_number", "source"
]

# === Set 1 ===
set1 = bills_df[bills_df["BS_PNL_Flag_Final"].astype(str).str.upper() == "YES"].copy()
# Add source column
set1["source"] = "Bill"

# Keep only existing columns
existing_set1_cols = [c for c in set1_cols if c in set1.columns]
set1 = set1[existing_set1_cols]

# Filter Expense_Account = 200000 or 400000
set1 = set1[set1["Expense_Account"].astype(str).str.contains("200000|400000", na=False)]

# ---- NEW LOGIC: Amount * -1 when Expense_Account contains 400000 ----
mask_400k = set1["Expense_Account"].astype(str).str.contains("400000", na=False)
set1.loc[mask_400k, "Amount"] = set1.loc[mask_400k, "Amount"].astype(float) #* -1

# === Set 2 ===
set2 = bills_df[bills_df["bankflow_final"].astype(str).str.upper() == "YES"].copy()

mask_additional_bestinet = (
    bills_df["Subsidiary_Name"]
    .astype(str)
    .str.lower()
    .str.contains("bestinet sdn bhd", na=False)
) & (
    bills_df["Account"].astype(str) == "400000"
) & (    
    bills_df["Payment_Account"]
    .astype(str)
    .str.contains("600111", na=False)
)




#set2_additional = (
    #bills_df.loc[mask_additional_bestinet]
    #.copy()
    #.drop_duplicates()
#)

set2_additional = (
    bills_df.loc[mask_additional_bestinet]
    .copy()
    .drop_duplicates(
        subset=[
            "Subsidiary_Name",
            "Account",
            "Payment_Account",
            "amount_bankflow"
        ]
    )
)


# Multiply amount_bankflow by -1
set2_additional["amount_bankflow"] = (
    set2_additional["amount_bankflow"].astype(float) * -1
)



set2 = pd.concat([set2, set2_additional], ignore_index=True)
# Add source column
set2["source"] = "Payment"

existing_set2_cols = [c for c in set2_cols if c in set2.columns]
set2 = set2[existing_set2_cols]

# Rename
set2 = set2.rename(columns={"Paid_date": "Date", "amount_bankflow": "Amount"})

# Apply same Expense_Account filter
set2 = set2[set2["Expense_Account"].astype(str).str.contains("200000|400000", na=False)]
mask_400k2 = set2["Expense_Account"].astype(str).str.contains("400000", na=False)
set2.loc[mask_400k2, "Amount"] = set2.loc[mask_400k2, "Amount"].astype(float) * -1

# === Concatenate ===
trade_final = pd.concat([set1, set2], ignore_index=True)

# besti special logic for May 2023 entry where user has done wrong entries
trade_final = trade_final[
    ~(
        (trade_final["Vendor_payment_number"] == "CUSTPMT_BT_014781") &
        (trade_final["source"] == "Bill")
    )
]


# === Save ===
trade_final.to_csv(trade_output_file, index=False)
print(f"✅ Trade_creditors_Debitors created with {len(trade_final)} records.")

##### Generating Cashflow summary file

# === Step XX: Aggregation by Subsidiary, Month, Cashflow_Category_first ===
print("📊 Creating aggregated bankflow summary using Cashflow_Category_first...")

# Use only bankflow_final = YES
set2 = bills_df[bills_df["bankflow_final"].astype(str).str.upper() == "YES"].copy()

# Convert Paid_date → datetime
set2['Paid_date'] = pd.to_datetime(set2['Paid_date'], errors='coerce', dayfirst=True)

# Derive Month (YYYY-MM)
set2['Month'] = set2['Paid_date'].dt.to_period('M')

# Clean category column
set2['Cashflow_Category_first'] = set2['Cashflow_Category_first'].fillna("").astype(str).str.strip()

#  NEW: Extract string after last colon from Subsidiary_Name
#  19 Feb -apply the fix for G3 and Bio clinic
#set2['Subsidiary_Name'] = (
    #set2['Subsidiary_Name']
    #.astype(str)
    #.apply(lambda x: x.split(':')[-1].strip())
#)

special_cases = [
    "G3 Healthcare Sdn Bhd (FKA:Bestinet Healthcare)",
    "Bio Clinic Sdn Bhd (FKA:Pengerang Technology)"
]


def clean_subsidiary(name):
    name = str(name).strip()

    if name in special_cases:
        # Do NOT apply any logic
        return name
    else:
        # Apply existing split logic
        return name.split(":", 1)[-1].strip()

set2["Subsidiary_Name"] = (
    set2["Subsidiary_Name"]
    .astype(str)
    .apply(clean_subsidiary)
)


distinct_subsidiaries = set2["Subsidiary_Name"].dropna().unique()
print(distinct_subsidiaries)


# Drop rows with missing Month or numeric issues
set2 = set2.dropna(subset=['Month', 'amount_bankflow', 'Cashflow_Category_first'])

# Ensure amount is numeric
set2['amount_bankflow'] = pd.to_numeric(set2['amount_bankflow'], errors='coerce')

# Group by Subsidiary, Month, Cashflow Category
agg_df = (
    set2.groupby(['Subsidiary_Name', 'Month', 'Cashflow_Category_first'])['amount_bankflow']
    .sum()
    .unstack(fill_value=0)
)

# Rename to final output names
agg_df = agg_df.rename(columns={
    'Operating': 'sum_Operating',
    'Financing': 'sum_Financing',
    'Investing': 'sum_Investing'
})

# Add Total column
for col in ['sum_Operating', 'sum_Financing', 'sum_Investing']:
    if col not in agg_df.columns:
        agg_df[col] = 0

agg_df['Total_Nett'] = (
    agg_df[['sum_Operating', 'sum_Financing', 'sum_Investing']]
    .sum(axis=1)
)

#####13_Jan_2026_Insert Zero values for not existing months 

# Ensure Month is Period type before filling gaps
#agg_df['Month'] = pd.PeriodIndex(agg_df['Month'], freq='M')

agg_df = agg_df.reset_index()

agg_df['Month'] = pd.to_datetime(agg_df['Month'].astype(str)).dt.to_period('M')



# Columns to zero-fill
value_cols = ['sum_Operating', 'sum_Financing', 'sum_Investing', 'Total_Nett']

final_dfs = []

# Process each Subsidiary independently
for subsidiary, grp in agg_df.groupby('Subsidiary_Name'):
    grp = grp.set_index('Month').sort_index()

    # Create full month range
    full_month_range = pd.period_range(
        start=grp.index.min(),
        end=grp.index.max(),
        freq='M'
    )

    # Reindex to include missing months
    grp = grp.reindex(full_month_range)

    # Restore Subsidiary_Name
    grp['Subsidiary_Name'] = subsidiary

    # Fill missing numeric values with zero
    grp[value_cols] = grp[value_cols].fillna(0)

    # Reset index
    grp = grp.reset_index().rename(columns={'index': 'Month'})

    final_dfs.append(grp)

# Combine all subsidiaries back
agg_df = pd.concat(final_dfs, ignore_index=True)

# Convert Month to string YYYY-MM for reporting
agg_df['Month'] = agg_df['Month'].astype(str)

# Final sort
agg_df = agg_df.sort_values(['Subsidiary_Name', 'Month'])

#####13_Jan_2026_Insert Zero values for not existing months 
# Final tidy output
final_df = agg_df.reset_index()



lookup_df = pd.read_excel(lookup_file_cashflowclosing)

# Ensure Month column converted to YYYY-MM period
lookup_df['Month'] = pd.to_datetime(lookup_df['Month'], errors='coerce').dt.to_period('M')

# Expected lookup fields:
#   Company_Name
#   Month
#   Cash_At_End_Of_Period

# Create Month_next = Month + 1 month
lookup_df['Month_next'] = lookup_df['Month'] + 1

lookup_df = lookup_df[['Company_Name', 'Month_next', 'Cash_At_End_Of_Period']].copy()

# Rename so it merges correctly
lookup_df = lookup_df.rename(columns={
    "Company_Name": "Subsidiary_Name",
    "Month_next": "Month"
})

# Convert Month to Period[M] on both DataFrames
final_df["Month"] = pd.to_datetime(
    final_df["Month"].astype(str), errors="coerce"
).dt.to_period("M")

lookup_df["Month"] = pd.to_datetime(
    lookup_df["Month"].astype(str), errors="coerce"
).dt.to_period("M")

# Merge to get opening balance
final_df = final_df.merge(
    lookup_df,
    how='left',
    on=['Subsidiary_Name', 'Month']
)

# Rename Cash_At_End_Of_Period → opening_balance
final_df = final_df.rename(columns={"Cash_At_End_Of_Period": "opening_balance"})

# Move opening_balance to appear AFTER Total
cols = list(final_df.columns)
total_idx = cols.index("Total_Nett")
new_order = cols[:total_idx+1] + ["opening_balance"] + cols[total_idx+1:-1]

final_df = final_df[new_order]

###9-Feb Logic addition Begin

# Convert Period month to timestamp
final_df['Month_dt'] = final_df['Month'].dt.to_timestamp()

# Cutoff date
cutoff_date = pd.to_datetime('2023-01-01')

# Find first month per Subsidiary
first_month_df = (
    final_df
    .groupby('Subsidiary_Name')['Month_dt']
    .min()
    .reset_index()
    .rename(columns={'Month_dt': 'first_month'})
)

# Merge first month back
final_df = final_df.merge(first_month_df, on='Subsidiary_Name', how='left')

# Set opening_balance = 0 for first month if first month > Jan-2023
final_df.loc[
    (final_df['Month_dt'] == final_df['first_month']) &
    (final_df['first_month'] > cutoff_date),
    'opening_balance'
] = 0

# Cleanup helper columns
final_df.drop(columns=['Month_dt', 'first_month'], inplace=True)


####09-Feb 2026 END logic

# Calculate closing balance
final_df['closing_balance'] = final_df['opening_balance'] + final_df['Total_Nett']

cols = list(final_df.columns)

# Remove closing_balance temporarily
cols.remove("closing_balance")

# Insert after opening_balance
open_idx = cols.index("opening_balance")
cols.insert(open_idx + 1, "closing_balance")

# Apply the order
final_df = final_df[cols]


# === APPLY ROLLING LOGIC FOR OPENING & CLOSING BALANCE ===

# Ensure Month is treated as a sortable period
final_df['Month'] = final_df['Month'].astype('period[M]')

# Sort properly for rolling logic
final_df = final_df.sort_values(['Subsidiary_Name', 'Month']).reset_index(drop=True)

# Iterate subsidiary-by-subsidiary
subs_list = final_df['Subsidiary_Name'].unique()

for sub in subs_list:
    sub_df_idx = final_df[final_df['Subsidiary_Name'] == sub].index
    
    started = False
    prev_close = None
    
    for i in sub_df_idx:
        
        ob = final_df.at[i, 'opening_balance']
        tot = final_df.at[i, 'Total_Nett']
        
        # Start rolling from the FIRST NON-NULL closing_balance record
        if not started:
            if pd.notnull(final_df.at[i, 'closing_balance']):
                started = True
                prev_close = final_df.at[i, 'closing_balance']
            continue
        
        # If rolling has started:
        if started:
            # Set opening_balance = previous closing balance
            final_df.at[i, 'opening_balance'] = prev_close
            
            # Calculate new closing_balance
            final_df.at[i, 'closing_balance'] = prev_close + tot
            
            # Update previous closing
            prev_close = final_df.at[i, 'closing_balance']

# === CLEAN SPECIAL CHARACTERS (REMOVE SINGLE QUOTE ') ===

cols_to_clean = [
    'sum_Operating', 'sum_Financing', 'sum_Investing', 'opening_balance' ,'closing_balance'
]

for col in cols_to_clean:
    if col in final_df.columns:
        final_df[col] = (
            final_df[col]
            .astype(str)
            .str.replace("'", "", regex=False)
            .str.strip()
        )


# === ROUND AMOUNT COLUMNS TO 4 DECIMAL PLACES ===
cols_to_round = ['sum_Operating', 'sum_Financing', 'sum_Investing','opening_balance' ,'closing_balance']

for col in cols_to_round:
    if col in final_df.columns:
        final_df[col] = final_df[col].astype(float).round(4)



# Save file
final_df.to_csv(cashflow_output_file, index=False)

print("✅ Aggregation with Opening Balance completed:", cashflow_output_file)


##### Generating partywise summary file

# === Step 1: Read CSV file ===
input_file_COA = r"C:\Broad_field_holdings\Net_suite\JR\inbound\Net_suite_Account_Mapping.xlsx"

acct_map_df = pd.read_excel(input_file_COA)


acct_map_df["Account"] = (
    acct_map_df["Account"]
    .astype(str)
    .str.strip()
)

acct_map_df["Category"] = (
    acct_map_df["Category"]
    .astype(str)
    .str.strip()
)

acct_map_df["Subcategory"] = (
    acct_map_df["Subcategory"]
    .astype(str)
    .str.strip()
)


#df = pd.read_csv(input_file)
df = trade_final

#bills_df

account_codes = [
    "16000B",
    "400999",
    "600402",
    "16000A",
    "160300",
    "160400"
]

# Ensure Expense_Account is string
bills_df["Expense_Account"] = bills_df["Expense_Account"].astype(str)

# Build regex pattern like: 16000B|400999|...
pattern = "|".join(account_codes)

# Apply filters
filtered_bills_df = bills_df[
    bills_df["Expense_Account"].str.contains(pattern, na=False)
    & (bills_df["BS_PNL_Flag_Final"].str.upper() == "YES")
].copy()

# Select required columns
filtered_bills_df = filtered_bills_df[
    [
        "Subsidiary_Name",
        "Date",
        "Type",
        "Document Number",
        "Name",
        "Expense_Account",
        "Amount",
        "Item",
        "TRANSACTION APPROVAL STATUS",
        "Status",
        "Vendor_Bill_number",
        "Vendor_payment_number",
    ]
]

# Append to existing df
df = pd.concat(
    [df, filtered_bills_df],
    ignore_index=True
)

# === Step 2: Convert Date column to datetime ===
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# === Step 3: Derive Month-Year from Date (YYYY-MM) ===
df["Month_Year"] = df["Date"].dt.to_period("M").astype(str)

# === Step 4: Ensure Amount is numeric ===
df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

# === Step 5: Group and aggregate ===
grouped_df = (
    df.groupby(
        ["Subsidiary_Name", "Name", "Expense_Account", "Month_Year"],
        as_index=False
    )["Amount"]
    .sum()
)

# === Step 6: (Optional) Sort for readability ===
grouped_df = grouped_df.sort_values(
    ["Subsidiary_Name", "Name", "Expense_Account", "Month_Year"]
)

grouped_df["Subsidiary_Name"] = (
    grouped_df["Subsidiary_Name"].astype(str).str.strip().str.upper()
)

grouped_df["Name"] = (
    grouped_df["Name"].astype(str).str.strip().str.upper()
)

grouped_df["Expense_Account"] = (
    grouped_df["Expense_Account"].astype(str).str.strip()
)

grouped_df["Month_Year"] = (
    pd.to_datetime(grouped_df["Month_Year"], errors="coerce")
    .dt.to_period("M")
    .astype(str)
)


grouped_df["Expense_Account"] = (
    grouped_df["Expense_Account"]
    .astype(str)
    .str.strip()
    .str.split(" ", n=1)
    .str[0]
)


bal_df = pd.read_excel(Debtors_Creditors_Balances_formatted)

bal_df["subsidiary_name"] = (
    bal_df["subsidiary_name"].astype(str).str.strip().str.upper()
)

bal_df["party_name"] = (
    bal_df["party_name"].astype(str).str.strip().str.upper()
)

bal_df["Account"] = (
    bal_df["Account"].astype(str).str.strip()
)

bal_df["Month_Year"] = (
    pd.to_datetime(bal_df["Period"], errors="coerce", format="mixed")
    .dt.to_period("M")
    .astype(str)
)

#bal_df["Account"] = (
    #pd.to_numeric(bal_df["Account"], errors="coerce")
    #.astype("Int64")
#)



# Normalize text columns
for col in ["Subsidiary_Name", "Name", "Expense_Account", "Month_Year"]:
    grouped_df[col] = grouped_df[col].astype(str).str.strip()

for col in ["subsidiary_name", "party_name", "Account", "Period"]:
    bal_df[col] = bal_df[col].astype(str).str.strip()

# Normalize Period to Month_Year format (YYYY-MM)
bal_df["Month_Year"] = (
    pd.to_datetime(bal_df["Period"], errors="coerce")
    .dt.to_period("M")
    .astype(str)
)


bal_df["Month_Year_shifted"] = (
    pd.to_datetime(bal_df["Period"], errors="coerce")
    .dt.to_period("M")
    .add(1)
    .astype(str)
)

print(grouped_df.columns.tolist())

print(bal_df.columns.tolist())

merged_df = grouped_df.merge(
    bal_df[
        ["subsidiary_name", "party_name", "Account", "Month_Year_shifted", "balance"]
    ],
    left_on=["Subsidiary_Name", "Name", "Expense_Account", "Month_Year"],
    right_on=["subsidiary_name", "party_name", "Account", "Month_Year_shifted"],
    how="outer",
    indicator=True
)


final_df = merged_df[merged_df["_merge"].isin(["both", "left_only"])].copy()

final_df["Opening balance"] = final_df["balance"]

bal_only_df = merged_df[merged_df["_merge"] == "right_only"].copy()

bal_only_df["Subsidiary_Name"] = bal_only_df["subsidiary_name"]
bal_only_df["Name"] = bal_only_df["party_name"]
bal_only_df["Expense_Account"] = bal_only_df["Account"]
bal_only_df["Month_Year"] = bal_only_df["Month_Year_shifted"]

bal_only_df["Opening balance"] = bal_only_df["balance"]
bal_only_df["Nett"] = 0.0
bal_only_df["Closing balance"] = bal_only_df["Opening balance"]


bal_only_df = bal_only_df[
    final_df.columns
]

final_df = pd.concat(
    [final_df, bal_only_df],
    ignore_index=True
)

final_df.drop(
    columns=[
        "subsidiary_name",
        "party_name",
        "Account",
        "Month_Year_shifted",
        "balance",
        "_merge"
    ],
    inplace=True,
    errors="ignore"
)


print(final_df.columns.tolist())



#final_df["Opening balance"] = final_df["balance"]



final_df["Opening balance"] = final_df["Opening balance"].fillna(0)


final_df.rename(columns={"Amount": "Nett"}, inplace=True)

final_df["Nett"] = (final_df["Nett"]).round(4)

final_df["Month_Year_dt"] = pd.to_datetime(
    final_df["Month_Year"], format="%Y-%m"
)

#final_df = final_df.sort_values(
    #["Subsidiary_Name", "Name", "Expense_Account", "Month_Year_dt"]
#)

cutoff = pd.Timestamp("2023-01-01")

global_end = final_df["Month_Year_dt"].max()

keys = ["Subsidiary_Name", "Name", "Expense_Account"]

expanded_groups = []

for _, g in final_df.groupby(keys):
    start = pd.Timestamp("2023-01-01")
    end = global_end

    months = pd.period_range(
        start=start,
        end=end,
        freq="M"
    ).astype(str)

    base = g[keys].iloc[0].to_dict()

    expanded = pd.DataFrame({
        **base,
        "Month_Year": months
    })

    expanded_groups.append(expanded)

expanded_df = pd.concat(expanded_groups, ignore_index=True)

final_df = expanded_df.merge(
    final_df,
    on=keys + ["Month_Year"],
    how="left"
)

# Normalize Subsidiary_Name in both dataframes
final_df["Subsidiary_Name"] = (
    final_df["Subsidiary_Name"]
    .astype(str)
    .str.strip()
    .str.upper()
)

bills_df["Subsidiary_Name"] = (
    bills_df["Subsidiary_Name"]
    .astype(str)
    .str.strip()
    .str.upper()
)

# Keep only subsidiaries present in bills_df
final_df = final_df[
    final_df["Subsidiary_Name"].isin(
        bills_df["Subsidiary_Name"].unique()
    )
].copy()



final_df["Nett"] = pd.to_numeric(
    final_df["Nett"], errors="coerce"
).fillna(0.0)



#final_df["Closing balance"] = pd.to_numeric(
    #final_df["Closing balance"], errors="coerce"
#).fillna(0.0)



final_df["Opening balance"] = pd.to_numeric(
    final_df["Opening balance"], errors="coerce"
)

final_df["Opening balance"] = pd.to_numeric(
    final_df["Opening balance"], errors="coerce"
).fillna(0.0)



final_df["Month_Year_dt"] = pd.to_datetime(
    final_df["Month_Year"], format="%Y-%m"
)

final_df = final_df.sort_values(
    keys + ["Month_Year_dt"]
)



def roll_forward(group):
    group = group.copy()

    for i in range(len(group)):
        if i == 0:
            opening = group.iloc[i]["Opening balance"] or 0
        else:
            opening = group.iloc[i - 1]["Closing balance"]

        group.iloc[i, group.columns.get_loc("Opening balance")] = opening
        group.iloc[i, group.columns.get_loc("Closing balance")] = (
            opening + group.iloc[i]["Nett"]
        )

    return group

final_df["Closing balance"] = 0.0

final_df = (
    final_df
    .groupby(keys, group_keys=False)
    .apply(roll_forward)
)

final_df.drop(columns=["Month_Year_dt"], inplace=True)


final_df["Opening balance"] = pd.to_numeric(
    final_df["Opening balance"], errors="coerce"
).astype(float)

final_df["Nett"] = pd.to_numeric(
    final_df["Nett"], errors="coerce"
).astype(float)

final_df["Closing balance"] = 0.0



final_df = (
    final_df
    .groupby(
        ["Subsidiary_Name", "Name", "Expense_Account"],
        group_keys=False
    )
    .apply(roll_forward)
)


final_df["Closing balance"] = (final_df["Closing balance"]).round(4)


final_df["Expense_Account"] = (
    final_df["Expense_Account"]
    .astype(str)
    .str.strip()
)

final_df = final_df.merge(
    acct_map_df[["Account", "Category", "Subcategory"]],
    left_on="Expense_Account",
    right_on="Account",
    how="left"
)

final_df.drop(columns=["Account"], inplace=True, errors="ignore")

#bills_df

#final_df.drop(columns=["Month_Year_dt"], inplace=True)


# Optional: Save to CSV
#output_file_partywise = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Trade_creditors_Debitors_partywise_Monthly.csv"
final_df.to_csv(output_file_partywise, index=False)


