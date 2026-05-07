import pandas as pd
import os
import argparse
import sys



# === Step 1: Define input/output paths ===
input_excel = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Bill_ETL_Transaction_Master_file.xlsx"
output_excel = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Bill_ETL_Transaction_Master_file_transformed.xlsx"
output_csv = r"C:\Broad_field_holdings\Net_suite\JR\Outbound\Bill_ETL_Transaction_Master_file_transformed.csv"
company_list_file = r"C:\Broad_field_holdings\Net_suite\JR\Company_list.xlsx"



# === Step 2: Ensure output folder exists ===
os.makedirs(os.path.dirname(output_excel), exist_ok=True)

print("========================================")
print("🚀 Running ETL Transformation for Bill / Expense Report / Invoice  File")
print("========================================")

# === Step 3: Load Excel ===
df = pd.read_excel(input_excel)

parser = argparse.ArgumentParser(description="Run script for a batch of companies")
parser.add_argument(
    "--batch",
    required=True,
    help="Batch number (SET) from Company_list file"
)

args = parser.parse_args()
batch_input = args.batch.strip().lower()

base_output_dir = r"C:\Broad_field_holdings\Net_suite\JR\Outbound"
os.makedirs(base_output_dir, exist_ok=True)

output_csv = os.path.join(
    base_output_dir,
    f"Bill_ETL_Transaction_Master_file_transformed_{batch_input}.csv"
)

output_excel = os.path.join(
    base_output_dir,
    f"Bill_ETL_Transaction_Master_file_transformed_{batch_input}.xlsx"
)





company_df = pd.read_excel(company_list_file)

company_df["Batch_clean"] = (
    company_df["Set"]
    .astype(str)
    .str.strip()
    .str.lower()
)

company_df["Company_clean"] = (
    company_df["Company_Name"]
    .astype(str)
    .str.strip()
    .str.lower()
)


batch_companies = (
    company_df.loc[
        company_df["Batch_clean"] == batch_input,
        "Company_clean"
    ]
    .dropna()
    .unique()
    .tolist()
)

if not batch_companies:
    print(f"❌ No companies found for batch: {args.batch}")
    sys.exit(1)

print(f"✅ Batch {args.batch}contains companies: {batch_companies}")

# Logic fix on 19 Feb 2026 
#df["Subsidiary_clean"] = (
    #df["Subsidiary_Name"]
    #.astype(str)
    #.str.split(":")
    #.str[-1]
    #.str.strip()
    #.str.lower()
#)



# Normalize Subsidiary_Name
df["Subsidiary_Name"] = df["Subsidiary_Name"].astype(str).str.strip()

special_cases = [
    "Headquarters : G3 Healthcare Sdn Bhd (FKA:Bestinet Healthcare)",
    "Headquarters : Bio Clinic Sdn Bhd (FKA:Pengerang  Technology)"
]

def clean_subsidiary(name):
    if name in special_cases:
        # Take value after FIRST colon
        return name.split(":", 1)[1].strip().lower()
    else:
        # Existing logic (unchanged)
        #return name.split(":", 1)[-1].strip().lower()
        return name.split(":")[-1].strip().lower()

df["Subsidiary_clean"] = df["Subsidiary_Name"].apply(clean_subsidiary)


df = df[
    df["Subsidiary_clean"].isin(batch_companies)
].copy()

print(df["Subsidiary_clean"].unique())
print(batch_companies)

# === Step 4: Clean column names ===
df.columns = df.columns.astype(str).str.strip()

# === Step 5: Validate required columns ===
required_cols = [
    "Payment_Account", "amount_paid", "Vendor_payment_number",
    "Amount", "Expense_Account", "Description", "Type"
]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise SystemExit(f"❌ Missing required columns: {missing}")

# === Step 6: Filter only Bill, Expense Report, and Invoice ===
df = df[df["Type"].astype(str).str.strip().isin(["Bill", "Expense Report", "Invoice", "Journal", "Payment" ,"Credit Memo", "Bill Credit", "Customer Refund", "Transfer", "Currency Revaluation"])].copy()

# === Step 7: Add new columns ===
df["amount_final"] = None
expense_index = df.columns.get_loc("Expense_Account")
df.insert(expense_index + 1, "Expense_Account_category", None)

# === Step 8: RULE 1 — Payment_Account LIKE '%500%' ===
mask_500 = df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)



# ---- For Bill / Expense Report → amount_final = amount_paid * -1
mask_bill_exp = df["Type"].astype(str).str.strip().isin(["Bill", "Expense Report"])
df.loc[mask_500 & mask_bill_exp, "amount_final"] = df.loc[mask_500 & mask_bill_exp, "amount_paid"] * -1
df.loc[mask_500 & mask_bill_exp, "amount_paid"] = df.loc[mask_500 & mask_bill_exp, "amount_paid"] * -1

# ---- For Invoice → amount_final = amount_paid
mask_invoice = df["Type"].astype(str).str.strip().eq("Invoice")
df.loc[mask_500 & mask_invoice, "amount_final"] = df.loc[mask_500 & mask_invoice, "amount_paid"] * -1
df.loc[mask_500 & mask_invoice, "amount_paid"] = df.loc[mask_500 & mask_invoice, "amount_paid"] * -1

# === Step 9: RULE 2 — Amount / Vendor logic ===

# ---- For Bill/Expense Report: duplicate Vendor_payment_number → amount_final = Amount
vendor_col = df["Vendor_payment_number"].astype(str).str.strip()
dup_mask = vendor_col.duplicated(keep=False)
valid_dup_mask = mask_bill_exp & mask_500 & dup_mask & vendor_col.notna() & (vendor_col != "")
df.loc[valid_dup_mask, "amount_final"] = df.loc[valid_dup_mask, "Amount"]

# ---- For Invoice: if Amount < amount_paid → amount_final = Amount
invoice_condition = mask_invoice & mask_500 & (df["Amount"] < df["amount_paid"])
df.loc[invoice_condition, "amount_final"] = df.loc[invoice_condition, "Amount"]

# === Step 10: RULE 3 — Expense_Account_category ===

# ---- For Bill/Expense Report: copy from Vendor Bill mapping
if "Vendor_Bill_number" in df.columns:
    vendor_bills = df[
        (df["Name"].isna()) &
        (~df["Description"].astype(str).str.contains("SERVICE TAX ON PURCHASES", case=False, na=False))
    ]
    vendor_expense_map = (
        vendor_bills[["Vendor_Bill_number", "Expense_Account"]]
        .dropna()
        .drop_duplicates(subset=["Vendor_Bill_number"])
        .set_index("Vendor_Bill_number")["Expense_Account"]
        .to_dict()
    )

    for idx in df[mask_500 & mask_bill_exp].index:
        vendor_no = df.at[idx, "Vendor_Bill_number"]
        if pd.notna(vendor_no) and vendor_no in vendor_expense_map:
            df.at[idx, "Expense_Account_category"] = vendor_expense_map[vendor_no]
else:
    print("⚠️ Column 'Vendor_Bill_number' not found — skipping Bill/Expense mapping.")

# ---- For Invoice: memo not like 'SOCSO' or 'EPF' and Description not like 'SERVICE TAX ON PURCHASES'

if "Vendor_Bill_number" in df.columns:
    # Define filter for Invoice-specific condition
    mask_invoice_rule3 = df[
        (df["Type"].astype(str).str.strip() == "Invoice") &
        (~df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)) &
        (~df["Description"].astype(str).str.contains("SERVICE TAX ON PURCHASES", case=False, na=False))
&
        #(~df["Memo"].astype(str).str.contains("SOCSO", case=False, na=False)) 
#&
        #(~df["Memo"].astype(str).str.contains("EPF", case=False, na=False))
(df["Expense_Account"].astype(str).str.contains("700", case=False, na=False))
    ]


    # Extract Vendor Bill–type records as reference (distinct by Vendor_Bill_number)
    vendor_expense_map_invoice = (
        mask_invoice_rule3[["Vendor_Bill_number", "Expense_Account"]]
        .dropna()
        .drop_duplicates(subset=["Vendor_Bill_number"])
        .set_index("Vendor_Bill_number")["Expense_Account"]
        .to_dict()
    )

    # Build mapping: Vendor_Bill_number → Expense_Account
   # vendor_expense_map_invoice = vendor_bills_for_invoice.set_index("Vendor_Bill_number")["Expense_Account"].to_dict()

    # Apply mapping to qualifying Invoice rows
    for idx in df[mask_500 & mask_invoice].index:
        vendor_no = df.at[idx, "Vendor_Bill_number"]
        if pd.notna(vendor_no) and vendor_no in vendor_expense_map_invoice:
            df.at[idx, "Expense_Account_category"] = vendor_expense_map_invoice[vendor_no]
else:
    print("⚠️ Column 'memo' not found — skipping Invoice Expense_Account_category logic.")

# === Step 11: RULE 4 — Journal Handling ===
mask_journal = df["Type"].astype(str).str.strip().eq("Journal")
mask_500_exp = df["Expense_Account"].astype(str).str.contains("500", case=False, na=False)



# For journals where Expense_Account like '500' → copy values
df.loc[mask_journal & mask_500_exp, "Payment_Account"] = df.loc[mask_journal & mask_500_exp, "Expense_Account"]
df.loc[mask_journal & mask_500_exp, "amount_final"] = df.loc[mask_journal & mask_500_exp, "Amount"]
df.loc[mask_journal & mask_500_exp, "Vendor_payment_number"] = df.loc[mask_journal & mask_500_exp, "Vendor_Bill_number"]

empty_paid_mask = df["Paid_date"].isna() | (df["Paid_date"].astype(str).str.strip() == "")
df.loc[mask_journal & empty_paid_mask, "Paid_date"] = df.loc[mask_journal & empty_paid_mask, "Date"]

# === Step 12: Journal Aggregation Logic ===
print("🧮 Applying Journal Aggregation Logic...")

journal_df = df[mask_journal & mask_500_exp].copy()
if not journal_df.empty:
    grouped = (
        journal_df.groupby(["Vendor_payment_number", "Payment_Account"], dropna=False)["amount_final"]
        .sum()
        .reset_index()
    )

    for _, row in grouped.iterrows():
        vendor_no = row["Vendor_payment_number"]
        pay_acc = row["Payment_Account"]
        total_amt = row["amount_final"]

        # Find the first matching record in the main df
        match_idx = df[
            (df["Type"].astype(str).str.strip() == "Journal") &
            (df["Vendor_payment_number"] == vendor_no) &
            (df["Payment_Account"] == pay_acc)
        ].index.min()

        if pd.notna(match_idx):
            df.at[match_idx, "amount_paid"] = total_amt
# ===remove this condition if need to put back the aggregation logic
#df.loc[mask_journal & mask_500_exp, "amount_paid"] = df.loc[mask_journal & mask_500_exp, "Amount"] 
          
            # === Step 13: RULE 5 — Payment Handling ===
mask_payment = df["Type"].astype(str).str.strip().eq("Payment")
mask_500_exp = df["Expense_Account"].astype(str).str.contains("500", case=False, na=False)

# For journals where Expense_Account like '500' → copy values
df.loc[mask_payment & mask_500_exp, "Payment_Account"] = df.loc[mask_payment & mask_500_exp, "Expense_Account"]
df.loc[mask_payment & mask_500_exp, "amount_final"] = df.loc[mask_payment & mask_500_exp, "Amount"]
df.loc[mask_payment & mask_500_exp, "amount_paid"] = df.loc[mask_payment & mask_500_exp, "Amount"]
df.loc[mask_payment & mask_500_exp, "Vendor_payment_number"] = df.loc[mask_payment & mask_500_exp, "Vendor_Bill_number"]

#empty_paid_mask = df["Paid_date"].isna() | (df["Paid_date"].astype(str).str.strip() == "")
df.loc[mask_payment , "Paid_date"] = df.loc[mask_payment , "Date"]



# --- Step 1: Handling Customer Refund records where Expense_Account contains 500 ---
#cust_refund_src = df[
    #(df["Type"].astype(str).str.upper() == "CUSTOMER REFUND") &
    #(df["Expense_Account"].astype(str).str.contains("500"))
#].copy()

# Build mapping: Vendor_Bill_number → Expense_Account 
#refund_expense_map = (
    #cust_refund_src.groupby("Vendor_Bill_number")["Expense_Account"]
    #.first()
    #.to_dict()
#)

8# --- Step 1: Build source mapping from Vendor Bill (Expense Account starts with 500) ---
refund_expense_map = (
    df[
        (df["Type"].astype(str).str.strip().str.upper() == "CUSTOMER REFUND") &
        (df["Expense_Account"].astype(str).str.strip().str.startswith("500")) &
        (df["Vendor_Bill_number"].notna()) &
        (df["Vendor_Bill_number"] != "")
    ]
    .groupby("Vendor_Bill_number")["Expense_Account"]
    .first()
    .to_dict()
)

# --- Step 2: Identify Customer Refund records with Vendor_payment_number not null ---
#mask_target = (
    #(df["Type"].astype(str).str.upper() == "CUSTOMER REFUND") &
    #(df["Vendor_payment_number"].notna()) &
    #(df["Vendor_payment_number"].astype(str).str.strip() != "") &
    #(df["Vendor_payment_number"].astype(str).str.startswith("400000"))  ##added condition to mask only for 40000
#)

mask_target = (
    (df["Type"].astype(str).str.strip().str.upper() == "CUSTOMER REFUND") &
    (df["Vendor_Bill_number"].notna()) &
    (df["Vendor_Bill_number"].astype(str).str.strip() != "") &
    (df["Vendor_payment_number"].notna()) &
    (df["Payment_Account"]
        .astype(str)
        .str.strip()
        .str.contains("400000", na=False)
    )
)




# --- Step 3: Copy Expense_Account → Payment_Account ---
df.loc[mask_target, "Payment_Account"] = df.loc[mask_target, "Vendor_Bill_number"].map(refund_expense_map)

mask_target1 = (
    (df["Type"].astype(str).str.strip().str.upper() == "CUSTOMER REFUND") 
)
# --- Step 4: Copy Amount → amount_paid ---
df.loc[mask_target1, "amount_paid"] = df.loc[mask_target1, "Amount"]
df.loc[mask_target1, "Paid_date"] = df.loc[mask_target1, "Date"]


# === Logic for type = "Transfer" ===
mask_transfer = df["Type"].astype(str).str.strip().eq("Transfer")

df.loc[mask_transfer, "amount_paid"] = df.loc[mask_transfer, "Amount"]
df.loc[mask_transfer, "Paid_date"] = df.loc[mask_transfer, "Date"]
df.loc[mask_transfer, "Vendor_payment_number"] = df.loc[mask_transfer, "Vendor_Bill_number"]
df.loc[mask_transfer, "Payment_Account"] = df.loc[mask_transfer, "Expense_Account"]
###########

mask_CR = df["Type"].astype(str).str.strip().eq("Currency Revaluation")


df.loc[mask_CR, "Vendor_payment_number"] = df.loc[mask_CR, "Vendor_Bill_number"]
df.loc[mask_CR & mask_500_exp, "Payment_Account"] = df.loc[mask_CR & mask_500_exp, "Expense_Account"]





# === Step 14: Add 'key' column after 'Paid_date' and populate 'bankflow' ===

#mask_500 = df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)
#mask_500 = df["Payment_Account"].astype(str).str.startswith("500")
mask_500 = (
    df["Payment_Account"].astype(str).str.startswith("500") &
    ~df["Payment_Account"].astype(str).str.startswith("500009")
)

if "Paid_date" in df.columns:
    paid_date_index = df.columns.get_loc("Paid_date")
    df.insert(paid_date_index + 1, "key", None)
else:
    df["key"] = None
df.loc[mask_500, "key"] = "bankflow"




# === Step 16: Update Expense_Account_category for SOCSO-related Invoices ===
print("🔄 Updating Expense_Account_category for SOCSO-related Invoice records...")

mask_socso_update = (
    (df["Type"].astype(str).str.strip().str.lower() == "invoice") &
    (df["Expense_Account_category"].isna() | (df["Expense_Account_category"].astype(str).str.strip() == "")) &
    (df["key"].astype(str).str.lower() == "bankflow") &
    (df["Memo"].astype(str).str.contains("SOCSO", case=False, na=False))
)

updated_count = mask_socso_update.sum()
df.loc[mask_socso_update, "Expense_Account_category"] = "160104 STATUTORY PAYABLES : SOCSO PAYABLE"


# === Step 14A: Add 'key1' for first unique (Vendor_payment_number, Payment_Account) combo with nonzero payment, only where key='bankflow' ===
print("🔑 Adding 'key1' column for first unique (Vendor_payment_number, Payment_Account) combo among key='bankflow' records...")

# Ensure required columns exist
required_cols = {"Vendor_payment_number", "Payment_Account", "amount_paid", "key"}
if required_cols.issubset(df.columns):
    # Insert 'key1' column right after 'key'
    key_index = df.columns.get_loc("key")
    df.insert(key_index + 1, "key1", None)

    # Filter only records where key='bankflow' and amount_paid ≠ 0
    #mask = (df["key"] == "bankflow") & (df["amount_paid"] != 0)
    mask = (
        (df["key"] == "bankflow") &
        (df["amount_paid"] != 0) &
        (df["Type"].astype(str).str.strip().str.lower() != "payment")
    )

    # Get first occurrence of each (Vendor_payment_number, Payment_Account)
    first_indices = (
        df[mask]
        .groupby(["Vendor_payment_number", "Payment_Account"], dropna=False)
        .head(1)
        .index
    )

    # Assign 'bankflow' to key1 for those rows
    df.loc[first_indices, "key1"] = "bankflow"

    print(f"✅ 'key1' assigned to {df['key1'].eq('bankflow').sum()} rows (within key='bankflow').")
else:
    print("⚠️ Missing required columns for key1 generation — skipping this step.")

# Mask for Customer Refund + key1 == "bankflow"
mask_key_copy = (
    (df["Type"].astype(str).str.upper() == "CUSTOMER REFUND") &
    (df["key"].astype(str).str.lower() == "bankflow")
)

# Copy key1 → key
df.loc[mask_key_copy, "key1"] = df.loc[mask_key_copy, "key"]


# === Step 15: Additional Journal logic — Copy Expense_Account for 2-line journal entries ===
print("🧾 Applying additional Journal logic for 2-line entries (copy Expense_Account)...")

# Filter only Journal records
journal_subset = df[df["Type"].astype(str).str.strip() == "Journal"].copy()

# Find Vendor_Bill_number groups with exactly 2 records
journal_pairs = (
    journal_subset.groupby("Vendor_Bill_number")
    .filter(lambda g: len(g) == 2)
    .copy()
)

# Process each 2-line journal group
for vb_no, group in journal_pairs.groupby("Vendor_Bill_number"):
    if vb_no is None or str(vb_no).strip() == "":
        continue

    # Identify rows
    idx_500 = group[group["Payment_Account"].astype(str).str.contains("500", na=False)].index
    idx_non500 = group[~group["Payment_Account"].astype(str).str.contains("500", na=False)].index

    # Proceed only if both 500 and non-500 entries exist
    if not idx_500.empty and not idx_non500.empty:
        expense_value = df.loc[idx_non500[0], "Expense_Account"]
        df.loc[idx_500, "Expense_Account_category"] = expense_value

print("✅ Completed Journal 2-line copy logic.")

# === Step 12C: Salary-based Expense Category update for Journal entries with key='bankflow' ===
print("💼 Applying 'SALARY' tagging logic for Journal entries with key='bankflow'...")

mask_salary_update = (
    (df["Type"].astype(str).str.strip().str.upper() == "JOURNAL")
    & (df["key"].astype(str).str.strip().str.lower() == "bankflow")
    & df["Memo"].astype(str).str.upper().str.contains("SALARY", na=False)
    & (df["Expense_Account_category"].isna() | (df["Expense_Account_category"].astype(str).str.strip() == ""))
)

df.loc[mask_salary_update, "Expense_Account_category"] = (
    "900102 EXPENSES : STAFF COSTS - SALARY"
)

print(f"✅ Updated {mask_salary_update.sum()} Journal records with key='bankflow' where Memo contains 'SALARY' and Expense_Account_category was null.")


# === Step X: Journal-only logic — copy single Expense_Account to Expense_Account_category for groups > 2 ===
print("🔧 Starting Journal-only group logic (group_count > 2, exclude Payment_Account containing '500')...")

# Safety: required columns
required = ["Type", "key", "Vendor_Bill_number", "Payment_Account", "Expense_Account", "Expense_Account_category"]
if not all(col in df.columns for col in required):
    print("⚠️ Required columns missing; skipping this block:", [c for c in required if c not in df.columns])
else:
    # Normalize helpers for safe matching
    df["__Type_norm"] = df["Type"].astype(str).str.strip().str.lower()
    df["__key_norm"] = df["key"].astype(str).str.strip().str.lower()
    # work on the subset of interest: Journal & bankflow
    mask_journal_bankflow = (df["__Type_norm"] == "journal") 
#& (df["__key_norm"] == "bankflow")
   #mask_journal_bankflow1 = (df["__Type_norm"] == "journal") & (df["__key_norm"] == "bankflow")

    # group only on that subset
    grouped = df[mask_journal_bankflow].groupby("Vendor_Bill_number")

    updated_groups = []
    for bill_no, grp in grouped:
        #if bill_no is None or str(bill_no).strip() == "":
           # continue

        if len(grp) <= 2:
            # skip groups that are not > 2 in size
            continue

        # Exclude rows where Payment_Account contains '500' (string match)
        non500_subset = grp[~grp["Payment_Account"].astype(str).str.contains("500", na=False)]

        # Distinct Expense_Account values in this non500 subset
        distinct_exp = pd.Series(non500_subset["Expense_Account"].dropna().astype(str).str.strip().unique())

        # Debug: show what we found
        #print(f"DEBUG: bill_no={bill_no}, group_size={len(grp)}, non500_count={len(non500_subset)}, distinct_exp={distinct_exp.tolist()}")

        if len(distinct_exp) == 1:
            value_to_copy = distinct_exp.iloc[0]
            # Apply only to Journal & bankflow rows for this Vendor_Bill_number
            mask_apply = (
                (df["Vendor_Bill_number"] == bill_no) &
                (df["__Type_norm"] == "journal") &
                (df["__key_norm"] == "bankflow")
            )
            df.loc[mask_apply, "Expense_Account_category"] = value_to_copy
            updated_groups.append((bill_no, value_to_copy, len(grp), len(non500_subset)))

    # drop helper cols
    df.drop(columns=["__Type_norm", "__key_norm"], inplace=True)

    # Log summary
    #if updated_groups:
       # print("✅ Updated Expense_Account_category for the following Vendor_Bill_number groups:")
        #for bill_no, val, grp_sz, non500_sz in updated_groups:
            #print(f"   - {bill_no}: copied '{val}' (group_size={grp_sz}, non500_count={non500_sz})")
    #else:
       # print("ℹ️ No groups matched the condition (no updates applied).")


# === Step X: Copy key → key1 for Payment records ===
if all(col in df.columns for col in ["Type", "key", "key1"]):
    mask_payment = df["Type"].astype(str).str.lower() == "payment"
    df.loc[mask_payment, "key1"] = df.loc[mask_payment, "key"]

# === Step 13: Lookup from NetSuite Account Mapping file ===
mapping_file = r"C:\Broad_field_holdings\Net_suite\JR\Inbound\Net_suite_Account_Mapping.xlsx"  # adjust as needed


if os.path.exists(mapping_file):
    print("🔍 Loading NetSuite Account Mapping file...")

    # Load and clean mapping file
    mapping_df = pd.read_excel(mapping_file)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    # Validate required columns
    required_map_cols = [
        "Account", "Account_Description", "Item", "Item_type",
        "Item_Category", "Category", "Subcategory", "Cashflow_Category", "Cashflow_Subcategory", "PNL_STMT_CATEGORY", "PNL_STMT_SUBCATEGORY", "PNL_STMT_MAIN_CATEGORY"
    ]
    missing_cols = [c for c in required_map_cols if c not in mapping_df.columns]
    if missing_cols:
        raise SystemExit(f"❌ Missing columns in mapping file: {missing_cols}")

    # === Extract prefix (characters before the first space) from Expense_Account ===
    df["Expense_Account_prefix"] = (
        df["Expense_Account"]
        .astype(str)
        .str.strip()
        .str.extract(r"^(\S+)", expand=False)
    )

    # === Build lookup dictionaries ===
    lookup_dicts = {
        "Account_Description": mapping_df.set_index("Account")["Account_Description"].to_dict(),
        "Item": mapping_df.set_index("Account")["Item"].to_dict(),
        "Item_type": mapping_df.set_index("Account")["Item_type"].to_dict(),
        "Item_Category": mapping_df.set_index("Account")["Item_Category"].to_dict(),
        "Category": mapping_df.set_index("Account")["Category"].to_dict(),
        "Subcategory": mapping_df.set_index("Account")["Subcategory"].to_dict(),
        "Cashflow_Category": mapping_df.set_index("Account")["Cashflow_Category"].to_dict(),
        "Cashflow_Subcategory": mapping_df.set_index("Account")["Cashflow_Subcategory"].to_dict(),
        "PNL_STMT_CATEGORY": mapping_df.set_index("Account")["PNL_STMT_CATEGORY"].to_dict(),
        "PNL_STMT_SUBCATEGORY": mapping_df.set_index("Account")["PNL_STMT_SUBCATEGORY"].to_dict(),
        "PNL_STMT_MAIN_CATEGORY": mapping_df.set_index("Account")["PNL_STMT_MAIN_CATEGORY"].to_dict()
    }

    # === Perform lookups ===
    for col, lookup in lookup_dicts.items():
        df[col] = df["Expense_Account_prefix"].map(lookup)

    # === Add derived 'Account' column (characters before first space) ===
    df["Account"] = df["Expense_Account_prefix"]

    # === Insert new columns right after 'Amount' (in the correct order) ===
    insert_cols = [
        "Account_Description", "Account", "Item", "Item_type",
        "Item_Category", "Category", "Subcategory", "Cashflow_Category", "Cashflow_Subcategory", "PNL_STMT_CATEGORY", "PNL_STMT_SUBCATEGORY", "PNL_STMT_MAIN_CATEGORY"
    ]

    if "Amount" in df.columns:
        amount_idx = df.columns.get_loc("Amount")
        for i, col in enumerate(insert_cols):
            # Pop and reinsert to preserve order
            series = df.pop(col)
            df.insert(amount_idx + 1 + i, col, series)
    else:
        print("⚠️ Column 'Amount' not found — appended lookup columns at the end.")

    # === Drop helper column ===
    df.drop(columns=["Expense_Account_prefix"], inplace=True, errors="ignore")
    print(f"✅ Successfully performed lookup and added derived 'Account' column for {len(mapping_df)} account mappings.")

else:
    print(f"⚠️ Mapping file not found at: {mapping_file} — skipping lookup.")



if "amount_paid" in df.columns and "Type" in df.columns:
    paid_idx = df.columns.get_loc("amount_paid")
    df.insert(paid_idx + 1, "amount_bankflow", df["amount_paid"].copy())

    # Multiply by -1 for Bill records
    #mask_bill = df["Type"].astype(str).str.strip().str.lower() == "bill"
    mask_bill = df["Type"].astype(str).str.strip().str.lower().isin(["bill", "expense report", "customer refund"])
    df.loc[mask_bill, "amount_bankflow"] = df.loc[mask_bill, "amount_bankflow"] * -1

    print(f"✅ Added 'amount_bankflow' column (copied from 'amount_paid', "
          f"and multiplied by -1 for {mask_bill.sum()} Bill rows).")
else:
    print("⚠️ Required columns ('amount_paid', 'Type') not found — skipping 'amount_bankflow' creation.")

# === Step 14C: Handle Credit Memo and Bill Credit records ===
print("🧾 Applying Credit Memo and Bill Credit logic...")

mask_credit = df["Type"].astype(str).str.strip().isin(["Credit Memo", "Bill Credit"])

if not mask_credit.any():
    print("ℹ️ No Credit Memo or Bill Credit records found.")
else:
    if "Date" in df.columns and "Paid_date" in df.columns:
        df.loc[mask_credit, "Paid_date"] = df.loc[mask_credit, "Date"]

    if "Vendor_Bill_number" in df.columns and "Vendor_payment_number" in df.columns:
        df.loc[mask_credit, "Vendor_payment_number"] = df.loc[mask_credit, "Vendor_Bill_number"]

    print(f"✅ Updated {mask_credit.sum()} Credit Memo/Bill Credit records:")
    print("   - Copied 'Date' → 'Paid_date'")
    print("   - Copied 'Vendor_Bill_number' → 'Vendor_payment_number'")


# ============== INSERT NEW COLUMNS ==============
# Insert Expense_Account_Bankflow after Expense_Account_category
if "Expense_Account_category" in df.columns:
    idx = df.columns.get_loc("Expense_Account_category")
    df.insert(idx + 1, "Expense_Account_Bankflow", None)
else:
    df["Expense_Account_Bankflow"] = None

# Insert Expense_Account_Bankflow_short after Expense_Account_Bankflow
idx2 = df.columns.get_loc("Expense_Account_Bankflow")
df.insert(idx2 + 1, "Expense_Account_Bankflow_short", None)


# ============== FILTER DATA FOR ELIGIBLE TYPES ==============
valid_types = ["bill", "invoice", "expense report"]
filtered_df = df[df["Type"].astype(str).str.strip().str.lower().isin(valid_types)]


# ============== SHORTENING FUNCTION ==============
def extract_short(val):
   
    if not isinstance(val, str):
        return val

    val = val.strip()
    if not val:
        return val

    # Extract prefix (before first space)
    prefix = val.split(" ")[0].strip()

    # Extract part after colon if exists
    after_colon = None
    if ":" in val:
        after_colon = val.split(":")[-1].strip()

    # Combine logic
    if after_colon:
        return f"{prefix} {after_colon}"
    else:
        return prefix



# ============== BUILD LOOKUPS ==============
bankflow_map = (
    filtered_df[
        (~filtered_df["Description"].astype(str).str.contains("SERVICE TAX ON PURCHASES", case=False, na=False)) &
        (~filtered_df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)) &
        (filtered_df["Amount"].fillna(0) != 0) &
        (~filtered_df["Expense_Account"].astype(str).str.contains("200000", case=False, na=False)) &
        (~filtered_df["Expense_Account"].astype(str).str.contains("400000", case=False, na=False))
    ]
    .sort_values(["Vendor_Bill_number", "Amount"], ascending=[True, False])  # ⬅️ Highest Amount First
    .groupby("Vendor_Bill_number")["Expense_Account"]
    .apply(lambda x: " && ".join(sorted(x.dropna().astype(str).str.strip().unique())))
    .to_dict()
)

bankflow_short_map = (
    filtered_df[
        (~filtered_df["Description"].astype(str).str.contains("SERVICE TAX ON PURCHASES", case=False, na=False)) &
        (~filtered_df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)) &
        (filtered_df["Amount"].fillna(0) != 0) &
        (~filtered_df["Expense_Account"].astype(str).str.contains("200000", case=False, na=False)) &
        (~filtered_df["Expense_Account"].astype(str).str.contains("400000", case=False, na=False))
    ]
    .sort_values(["Vendor_Bill_number", "Amount"], ascending=[True, False])  # ⬅️ Highest Amount First
    .groupby("Vendor_Bill_number")["Expense_Account"]
    .apply(lambda x: " && ".join(sorted(
        extract_short(v) for v in x.dropna().astype(str).str.strip().unique()
        if extract_short(v) not in (None, "", "nan")  # prevents blank outputs
    )))
    .to_dict()
)



# ============== APPLY VALUES TO DF ==============
mask_to_update = (
    df["Type"].astype(str).str.strip().str.lower().isin(valid_types) &
    df["Payment_Account"].astype(str).str.contains("500", case=False, na=False)
)

for row in df[mask_to_update].index:
    vbn = df.at[row, "Vendor_Bill_number"]
    if pd.notna(vbn):
        if vbn in bankflow_map:
            df.at[row, "Expense_Account_Bankflow"] = bankflow_map[vbn]
        if vbn in bankflow_short_map:
            df.at[row, "Expense_Account_Bankflow_short"] = bankflow_short_map[vbn]



print(f"✅ Completed Expense_Account_Bankflow population for {df['Expense_Account_Bankflow'].notna().sum()} rows.")

# -------------- Journal Processing Logic ----------------
df["Type_clean"] = df["Type"].astype(str).str.strip().str.lower()

# Direct unconditional copy for Journal records (full version)
df.loc[df["Type_clean"] == "journal", "Expense_Account_Bankflow"] = \
    df.loc[df["Type_clean"] == "journal", "Expense_Account_category"]

df.loc[df["Type_clean"] == "journal", "Expense_Account_Bankflow_short"] = \
    df.loc[df["Type_clean"] == "journal", "Expense_Account_category"]

# -------------- Short Value Helper ----------------
def shorten_expense_value(val):
    if not isinstance(val, str) or not val.strip():
        return None

    val = val.strip()

    # extract prefix (before first space)
    prefix = val.split(" ")[0].strip()

    # extract part after colon if present
    after_colon = val.split(":")[-1].strip() if ":" in val else None
   #after_colon = val.split(":")[-1].strip() if ":" in val else None 

    if after_colon:
        # return BOTH values separated by space
        return f"{prefix} {after_colon}".strip()

    return prefix  # fallback, always non-empty


# Loop each Vendor_Bill_Number
for vbn in df.loc[df["Type"] == "Journal", "Vendor_Bill_number"].dropna().unique():

    # Filter only Journal & Expense_Account_category NULL rows
    sub = df[
        (df["Vendor_Bill_number"] == vbn) &
        (df["Type_clean"] == "journal") &
        (df["Expense_Account_category"].isna()) &
        (~df["Expense_Account"].astype(str).str.contains("200000", case=False, na=False)) &
        (~df["Expense_Account"].astype(str).str.contains("400000", case=False, na=False))
    ]

    if sub.empty:
        full_concat = None
        short_concat = None
        continue

    # Source rows: Payment_Account NOT containing "500"
    source = sub[~sub["Payment_Account"].astype(str).str.contains("500", case=False, na=False)]

    if not source.empty:
        # ============================
        # FULL CONCAT — Max Amount Sort
        # ============================
        full_df = source[["Expense_Account", "Amount"]].dropna(subset=["Expense_Account"]).copy()
        full_df = full_df.groupby("Expense_Account", as_index=False)["Amount"].max().sort_values("Amount", ascending=False)
        full_vals = full_df["Expense_Account"].astype(str).str.strip().tolist()
        full_concat = " && ".join(full_vals) if full_vals else None

        # ============================
        # SHORT CONCAT — Shortened values
        # ============================
        short_df = source[["Expense_Account", "Amount"]].dropna(subset=["Expense_Account"]).copy()
        short_df = short_df.groupby("Expense_Account", as_index=False)["Amount"].max().sort_values("Amount", ascending=False)
        short_vals = []
        for acc in short_df["Expense_Account"]:
            sv = shorten_expense_value(str(acc))
            if sv and sv not in short_vals:
                short_vals.append(sv)
        short_concat = " && ".join(short_vals) if short_vals else None
    else:
        full_concat = None
        short_concat = None

    # TARGET rows: Payment_Account containing "500"
    target_mask = (
        (df["Vendor_Bill_number"] == vbn) &
        (df["Type_clean"] == "journal") &
        (df["Expense_Account_category"].isna()) &
        (df["Payment_Account"].astype(str).str.contains("500", case=False, na=False))
    )

    if full_concat:
        df.loc[target_mask, "Expense_Account_Bankflow"] = full_concat
    if short_concat:
        df.loc[target_mask, "Expense_Account_Bankflow_short"] = short_concat

print(f"✅ Completed Expense_Account_Bankflow population for Journal {df['Expense_Account_Bankflow'].notna().sum()} rows.")
# ============================================================
# Step: Create Final Output Columns
# ============================================================

# Insert 2 new columns AFTER Expense_Account_Bankflow_short
if "Expense_Account_Bankflow_short" in df.columns:
    idx = df.columns.get_loc("Expense_Account_Bankflow_short")
    df.insert(idx + 1, "Expense_Account_Bankflow_final", None)
    df.insert(idx + 2, "Expense_Account_Bankflow_short_final", None)
else:
    # fallback if previous column missing
    df["Expense_Account_Bankflow_final"] = None
    df["Expense_Account_Bankflow_short_final"] = None
    print("Did not find Expense_Account_Bankflow_short; columns appended at end.")


# ============================================================
# Build mappings only using records where key = bankflow
# ============================================================

eligible_source = df[df["key"].astype(str).str.lower() == "bankflow"]

# Build concatenation maps for both full + short
bankflow_final_map = (
    eligible_source.groupby("Vendor_payment_number")["Expense_Account_Bankflow"]
    .apply(lambda x: " && ".join(sorted(x.dropna().astype(str).str.strip().unique())))
    .to_dict()
)

bankflow_short_final_map = (
    eligible_source.groupby("Vendor_payment_number")["Expense_Account_Bankflow_short"]
    .apply(lambda x: " && ".join(sorted(x.dropna().astype(str).str.strip().unique())))
    .to_dict()
)

# ============================================================
# Populate values in records where key1 = bankflow
# ============================================================

target_mask = df["key1"].astype(str).str.lower() == "bankflow"

for idx in df[target_mask].index:
    vpn = df.at[idx, "Vendor_payment_number"]
    
    if pd.notna(vpn):
        # Full version update
        if vpn in bankflow_final_map and bankflow_final_map[vpn]:
            df.at[idx, "Expense_Account_Bankflow_final"] = bankflow_final_map[vpn]

        # Short version update
        if vpn in bankflow_short_final_map and bankflow_short_final_map[vpn]:
            df.at[idx, "Expense_Account_Bankflow_short_final"] = bankflow_short_final_map[vpn]


print(" Successfully populated final level Bankflow columns:")
print(f"   → Expense_Account_Bankflow_final updated rows: {df['Expense_Account_Bankflow_final'].notna().sum()}")
print(f"   → Expense_Account_Bankflow_short_final updated rows: {df['Expense_Account_Bankflow_short_final'].notna().sum()}")



# Insert new column after Vendor_Bill_number if not already present
if "Vendor_Bill_number" in df.columns:
    insert_idx = df.columns.get_loc("Vendor_Bill_number")
    if "Vendor_Bill_number_multi" not in df.columns:
        df.insert(insert_idx + 1, "Vendor_Bill_number_multi", None)
else:
    df["Vendor_Bill_number_multi"] = None
    print("⚠️ Column 'Vendor_Bill_number' not found. Added at end.")

# Build mapping: Vendor_payment_number → concatenated distinct Vendor_Bill_number
vb_map = (
    df[~df["Vendor_Bill_number"].isna() & (df["Type"].astype(str).str.strip().str.lower() != "payment")] # 21-Dec -added condition to exclude payment type
    .groupby("Vendor_payment_number")["Vendor_Bill_number"]
    .apply(lambda x: " && ".join(sorted(set(str(i).strip() for i in x))))
    .to_dict()
)

# Update ONLY records where key1 == 'bankflow'
mask = df["key1"].astype(str).str.strip().str.lower() == "bankflow"

for idx in df[mask].index:
    vp = df.at[idx, "Vendor_payment_number"]
    if pd.notna(vp) and vp in vb_map:
        df.at[idx, "Vendor_Bill_number_multi"] = vb_map[vp]

print(" Completed Vendor_Bill_number_multi population.")

# Insert new column after Vendor_Bill_number if not already present
if "Vendor_Bill_number" in df.columns:
    insert_idx = df.columns.get_loc("Vendor_Bill_number")
    if "Vendor_payment_number_multi" not in df.columns:
        df.insert(insert_idx + 1, "Vendor_payment_number_multi", None)
else:
    df["Vendor_payment_number_multi"] = None
    print("⚠️ Column 'Vendor_Bill_number' not found — added at end.")

# Build mapping: Vendor_Bill_number → concatenated distinct Vendor_payment_number
vp_map = (
    df[~df["Vendor_payment_number"].isna()]
    .groupby("Vendor_Bill_number")["Vendor_payment_number"]
    .apply(lambda x: " && ".join(sorted(set(str(i).strip() for i in x))))
    .to_dict()
)

# Condition: update ONLY rows where key1 = "bankflow"
mask = df["key1"].astype(str).str.strip().str.lower() == "bankflow"

for idx in df[mask].index:
    vb = df.at[idx, "Vendor_Bill_number"]
    if pd.notna(vb) and vb in vp_map:
        df.at[idx, "Vendor_payment_number_multi"] = vp_map[vb]

print(" Completed Vendor_payment_number_multi population.")

# --- Create column before `Item` --- #
if "BS_PNL_Flag_Final" in df.columns:
    df.drop(columns=["BS_PNL_Flag_Final"], inplace=True)

if "Item" in df.columns:
    insert_idx = df.columns.get_loc("Item")
else:
    insert_idx = len(df.columns)

df.insert(insert_idx, "BS_PNL_Flag_Final", None)

# Normalize fields
df["Type_clean"] = df["Type"].astype(str).str.strip().str.lower()
df["Status_clean"] = df["Status"].astype(str).str.strip().str.lower()
df["Expense_Account_str"] = df["Expense_Account"].astype(str).str.strip().str.lower()

# Apply only for Bills and Expense Reports
#target_mask = df["Type_clean"].isin(["bill", "expense report"])
#df_target = df[target_mask]

target_mask = df["Type_clean"].isin(["bill", "expense report"])

eligible_status_mask = df["Status_clean"].isin(
    ["open", "paid in full", "approved by accounting"]
)

contains_200000 = df["Expense_Account_str"].str.contains("200000", na=False)


# Group by Expense_Account + Amount + Vendor_Bill_number
group_cols = ["Expense_Account_str", "Amount", "Vendor_Bill_number"]

#for _, sub_df in df_target.groupby(group_cols):
    # Eligible statuses
    #eligible_records = sub_df[sub_df["Status_clean"].isin(["open", "paid in full", "approved by accounting"])]
    #if eligible_records.empty:
        #continue

    #exp_acc = sub_df["Expense_Account_str"].iloc[0]

    #if "200000" in exp_acc:
        # Contains 200000 → Only 1st eligible record marked YES
        #df.at[eligible_records.index[0], "BS_PNL_Flag_Final"] = "YES"
    #else:
        # Does NOT contain 200000 → ALL eligible records marked YES
       # df.loc[eligible_records.index, "BS_PNL_Flag_Final"] = "YES"

mask_all_yes = (
    target_mask &
    eligible_status_mask &
    ~contains_200000
)

df.loc[mask_all_yes, "BS_PNL_Flag_Final"] = "YES"


mask_200000 = (
    target_mask &
    eligible_status_mask &
    contains_200000
)

first_idx = (
    df.loc[mask_200000]
      .groupby(group_cols, sort=False)
      .head(1)
      .index
)

df.loc[first_idx, "BS_PNL_Flag_Final"] = "YES"

print(" Completed BS_PNL_Flag_Final = YES population for bill and expense report.")

# --- Invoice BS_PNL_Flag_Final logic --- #
invoice_mask = df["Type_clean"] == "invoice"
df_invoice = df[invoice_mask]

eligible_status_mask = df["Status_clean"].isin(("open", "paid in full"))

contains_400000 = df["Expense_Account_str"].str.contains(
    "400000", regex=False, na=False
)

# Group by Expense_Account + Amount + Vendor_Bill_number
group_cols_invoice = ["Expense_Account_str", "Amount", "Vendor_Bill_number"]

#for _, sub_df in df_invoice.groupby(group_cols_invoice):
    # Eligible statuses
   # eligible_records = sub_df[sub_df["Status_clean"].isin(["open", "paid in full"])]
    #if eligible_records.empty:
       # continue

    #exp_acc = sub_df["Expense_Account_str"].iloc[0]

    #if "400000" in exp_acc:
        # Contains 400000 → Only 1st eligible record marked YES
        #df.at[eligible_records.index[0], "BS_PNL_Flag_Final"] = "YES"
    #else:
        # Does NOT contain 400000 → ALL eligible records marked YES
       # df.loc[eligible_records.index, "BS_PNL_Flag_Final"] = "YES"

mask_all_yes = (
    invoice_mask &
    eligible_status_mask &
    ~contains_400000
)

df.loc[mask_all_yes, "BS_PNL_Flag_Final"] = "YES"


mask_first_only = (
    invoice_mask &
    eligible_status_mask &
    contains_400000
)

first_idx = (
    df.loc[mask_first_only]
      .groupby(group_cols_invoice, sort=False, observed=True)
      .head(1)
      .index
)

df.loc[first_idx, "BS_PNL_Flag_Final"] = "YES"

print(" Completed BS_PNL_Flag_Final = YES population for Invoice records.")

# --- Journal BS_PNL_Flag_Final logic (Aligned with Bills / Expense Reports) ---

# Ensure output column exists
if "BS_PNL_Flag_Final" not in df.columns:
    df["BS_PNL_Flag_Final"] = ""

# Identify Journal records
journal_mask = df["Type_clean"] == "journal"

# Eligible statuses for journals
eligible_journal_status = ~df["Status_clean"].isin([
    "pending approval",
    "rejected"
])

#commented on 25-Feb-2026
#contains_200000 = df["Expense_Account_str"].str.contains(
    #"200000", regex=False, na=False
#)
# 25-Feb-2026 added additional logic for Besti 40000
contains_200000 = (df["Expense_Account_str"].str.contains(
    "200000", regex=False, na=False
) |
df["Expense_Account_str"].str.contains(
    "400000", regex=False, na=False
))

is_krypton = df["Subsidiary_Name"] \
    .astype(str) \
    .str.strip() \
    .str.lower() \
    .str.contains("krypton", na=False)


# Work only on eligible journal rows
#df_journal = df[journal_mask & eligible_journal_status]

# Grouping logic (same as Bills/Expense)
group_cols = ["Expense_Account_str", "Amount", "Vendor_Bill_number"]


#for _, sub_df in df_journal.groupby(group_cols):

    #if sub_df.empty:
        #continue

    #exp_acc = sub_df["Expense_Account_str"].iloc[0]

    #if "200000" in exp_acc:
        # Only first record flagged
        #df.at[sub_df.index[0], "BS_PNL_Flag_Final"] = "YES"
    #else:
        # All records flagged
        #df.loc[sub_df.index, "BS_PNL_Flag_Final"] = "YES"

mask_all_yes = (
    journal_mask &
    eligible_journal_status &
    ~contains_200000
)

df.loc[mask_all_yes, "BS_PNL_Flag_Final"] = "YES"

print(" Completed BS_PNL_Flag_Final = YES population for Journal records.")


mask_first_only = (
    journal_mask &
    eligible_journal_status &
    contains_200000
)

first_idx = (
    df.loc[mask_first_only]
      .groupby(group_cols, sort=False, observed=True)
      .head(1)
      .index
)

df.loc[first_idx, "BS_PNL_Flag_Final"] = "YES"

#is_krypton = (
    #df["Subsidiary_Name"]
    #.astype(str)
    #.str.strip()
    #.str.lower()
    #.str.contains("krypton", na=False) #Krypton
#)

is_krypton = (
    df["Subsidiary_Name"]
    .astype(str)
    .str.strip()
    .str.lower()
    .str.contains("krypton|bestinet sdn bhd", na=False)
)



base_group_cols = ["Expense_Account_str", "Amount", "Vendor_Bill_number"]

mask_krypton = (
    journal_mask &
    eligible_journal_status &
    is_krypton
)

#krypton_group_cols = base_group_cols + ["Memo"]
krypton_group_cols = base_group_cols + ["Memo", "Name"]


mask_krypton_all_yes = mask_krypton & ~contains_200000
df.loc[mask_krypton_all_yes, "BS_PNL_Flag_Final"] = "YES"


# First only where contains 200000
mask_krypton_first_only = mask_krypton & contains_200000

first_idx_krypton = (
    df.loc[mask_krypton_first_only]
      .groupby(krypton_group_cols, sort=False, observed=True)
      .head(1)
      .index
)

df.loc[first_idx_krypton, "BS_PNL_Flag_Final"] = "YES"



print(" Completed BS_PNL_Flag_Final = YES population for Kryptoon Journal records.")

# --- Update BS_PNL_Flag_Final for Bill Credit and Credit Memo --- #

# Bill Credit: mark ALL records as YES
bill_credit_mask = df["Type"].astype(str).str.strip().str.lower() == "bill credit"
df.loc[bill_credit_mask, "BS_PNL_Flag_Final"] = "YES"

# Credit Memo: mark YES only for Status in Open or Fully Applied
credit_memo_mask = df["Type"].astype(str).str.strip().str.lower() == "credit memo"
eligible_status = ["open", "fully applied"]
df.loc[credit_memo_mask & df["Status"].astype(str).str.strip().str.lower().isin(eligible_status), "BS_PNL_Flag_Final"] = "YES"


# === Logic for type = "Currency Revaluation" & updating BS_PNL_Flag_Final ===
mask_CR = df["Type"].astype(str).str.strip().eq("Currency Revaluation")

df.loc[mask_CR, "amount_paid"] = df.loc[mask_CR, "Amount"]
df.loc[mask_CR, "Paid_date"] = df.loc[mask_CR, "Date"]
df.loc[mask_CR, "Vendor_payment_number"] = df.loc[mask_CR, "Vendor_Bill_number"]
df.loc[mask_CR & mask_500_exp, "Payment_Account"] = df.loc[mask_CR & mask_500_exp, "Expense_Account"]
df.loc[mask_CR, "BS_PNL_Flag_Final"] = "YES"
df.loc[mask_CR, "amount_bankflow"] = df.loc[mask_CR, "amount_paid"]
df.loc[mask_CR, "key1"] = df.loc[mask_CR, "key"]
###########

# ---------- Create New Column After Expense_Account_Bankflow_short_final ----------
new_col = "expense_account_bankflow_max"

if new_col in df.columns:
    df.drop(columns=[new_col], inplace=True)

if "Expense_Account_Bankflow_short_final" in df.columns:
    insert_idx = df.columns.get_loc("Expense_Account_Bankflow_short_final")
    df.insert(insert_idx + 1, new_col, None)
else:
    df[new_col] = None

# Normalize comparison fields
df["Expense_Account_str"] = df["Expense_Account"].astype(str).str.strip().str.lower()
df["Payment_Account_str"] = df["Payment_Account"].astype(str).str.strip().str.lower()

# ---------- Processing per Vendor_Bill_number ----------
for vbn, grp in df.groupby("Vendor_Bill_number"):

    # Eligible source records (excluding unwanted values + non-zero amount)
    src = grp[
        (~grp["Payment_Account_str"].str.contains("500", na=False)) &
        (~grp["Expense_Account_str"].str.contains("200000", na=False)) &
        (~grp["Expense_Account_str"].str.contains("400000", na=False)) &
        (grp["Amount"].fillna(0) != 0)
    ]

    if src.empty:
        continue

    # Determine sign of amounts
    if (src["Amount"] > 0).all():
        # All positive → take max per Expense_Account
        max_map = src.groupby("Expense_Account")["Amount"].max()
        top_expense_account = max_map.idxmax()
        top_amount = max_map.max()
    elif (src["Amount"] < 0).all():
        # All negative → take min per Expense_Account
        min_map = src.groupby("Expense_Account")["Amount"].min()
        top_expense_account = min_map.idxmin()
        top_amount = min_map.min()
    else:
        # Mixed signs → optionally skip or take the max by absolute value
        abs_map = src.groupby("Expense_Account")["Amount"].apply(lambda x: x.abs().max())
        top_expense_account = abs_map.idxmax()
        # Use actual signed amount corresponding to that Expense_Account
        top_amount = src.loc[src["Expense_Account"] == top_expense_account, "Amount"].iloc[0]

    # Combine into desired concatenated format
    concat_value = f"{top_expense_account}|{top_amount}"

    # Populate ONLY for rows where key1 = 'bankflow'
    mask_target = (
        (df["Vendor_Bill_number"] == vbn) &
        (df["key1"].astype(str).str.lower() == "bankflow")
    )

    df.loc[mask_target, new_col] = concat_value

print("✅ Completed — field populated with Expense_Account + Amount considering positive/negative logic.")


# --- Remove working/helper columns before saving --- #
for col in ["Type_clean", "Status_clean", "Expense_Account_str", "Payment_Account_str"]:
    if col in df.columns:
        df.drop(columns=[col], inplace=True)

# === Step 15: Copying multiple Expense_Account fields from Invoice → Payment where Vendor_payment_number matches 

print("🔄 Copying Expense_Account fields from Invoice → Payment...")

# Columns to map from Invoice to Payment
columns_to_copy = [
    "Expense_Account_category",
    "Expense_Account_Bankflow",
    "Expense_Account_Bankflow_short",
    "Expense_Account_Bankflow_final",
    "Expense_Account_Bankflow_short_final",
    "expense_account_bankflow_max"
]

# Create mapping dictionary for each column
invoice_maps = {}

invoice_rows = df[df["Type"].astype(str).str.strip() == "Invoice"]

for col in columns_to_copy:
    invoice_maps[col] = (
        invoice_rows
        .dropna(subset=["Vendor_payment_number", col])
        .drop_duplicates(subset=["Vendor_payment_number"])
        .set_index("Vendor_payment_number")[col]
        .to_dict()
    )

# Apply mappings to Payment rows
mask_payment = df["Type"].astype(str).str.strip() == "Payment"

for idx in df[mask_payment].index:
    vendor_no = df.at[idx, "Vendor_payment_number"]

    if pd.notna(vendor_no):
        # Loop through each column and copy if available
        for col in columns_to_copy:
            if vendor_no in invoice_maps[col]:
                df.at[idx, col] = invoice_maps[col][vendor_no]


# === Step 1: Cashflow_Category_first population Prepare mapping dictionary ===
# Ensure Account column in mapping file is clean
mapping_df["Account"] = mapping_df["Account"].astype(str).str.strip()

# Create lookup dictionary: Account → Cashflow_Category
account_lookup = mapping_df.set_index("Account")["Cashflow_Category"].to_dict()

# Dictionary for subcategory
subcategory_lookup = mapping_df.set_index("Account")["Cashflow_Subcategory"].to_dict()


# === Step 2: Extract prefix from source column ===
#df["Bankflow_prefix"] = df["Expense_Account_Bankflow_short_final"].astype(str).str.strip().str.extract(r"^(\S+)", expand=False)
df["Bankflow_prefix"] = df["expense_account_bankflow_max"].astype(str).str.strip().str.extract(r"^(\S+)", expand=False)

# === Step 3: Map Cashflow_Category using the prefix ===
df["Cashflow_Category_first"] = df["Bankflow_prefix"].map(account_lookup)

# === Step D: Populate cashflow_first_subcategory ===
df["cashflow_first_subcategory"] = df["Bankflow_prefix"].map(subcategory_lookup).fillna("")

# === Step 4: (Optional) Drop helper column if not needed ===
df.drop(columns=["Bankflow_prefix"], inplace=True, errors="ignore")


# === Step 16: Summary ===
print("========================================")
print(f"Total rows processed: {len(df)}")
print(f"Rows updated (Payment_Account like '%500%' - Bill/Expense): {(mask_500 & mask_bill_exp).sum()}")
print(f"Rows updated (Payment_Account like '%500%' - Invoice): {(mask_500 & mask_invoice).sum()}")
print(f"Rows updated (duplicate Vendor_payment_number - Bill/Expense): {valid_dup_mask.sum()}")
print(f"Rows updated (Invoice Amount < amount_paid): {invoice_condition.sum()}")
print(f"Rows updated (Expense_Account_category populated): {df['Expense_Account_category'].notna().sum()}")
print(f"Rows updated (Journal - Expense_Account like '500'): {(mask_journal & mask_500_exp).sum()}")
print(f"Rows tagged with 'bankflow' in key column: {df['key'].eq('bankflow').sum()}")
print(f"✅ Copied Expense_Account for {(mask_payment & df['Vendor_payment_number'].isin(invoice_maps.keys())).sum()} Payment rows.")
print(f"✅ Updated {updated_count} records where Type='Invoice', key='bankflow', and Memo like 'SOCSO'.")
print("========================================")

# === Step 17: Save outputs ===
df.to_excel(output_excel, index=False)
df.to_csv(output_csv, index=False, encoding='utf-8-sig')

print(f"✅ ETL complete. Files saved to:\n{output_excel}\n{output_csv}")
print("========================================")
