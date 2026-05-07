import os
import glob
import pandas as pd

# === File paths ===
#input_file = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Inbound\2026_ledger.xlsx"

# === Folder path ===
input_folder = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Inbound"

output_file = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Outbound\Bill_ETL_file_updated_balance_CCT.csv"

# Step 1: Read file
#df = pd.read_excel(input_file, header=None, skiprows=10)

# ==============================
# Step 1: Read ALL ledger_CCT files
# ==============================
file_pattern = os.path.join(input_folder, "*ledger_CCT*.xlsx")
files = sorted(glob.glob(file_pattern))

if not files:
    raise ValueError("❌ No files found with 'ledger_CCT' in the folder")

print("📂 Files detected:")
for f in files:
    print(f)

df_list = []

for file in files:
    try:
        temp_df = pd.read_excel(file, header=None, skiprows=10)
        temp_df["Source_File"] = os.path.basename(file)  # optional (debug trace)
        df_list.append(temp_df)
    except Exception as e:
        print(f"⚠️ Error reading {file}: {e}")

# Combine all files
df = pd.concat(df_list, ignore_index=True, sort=False)

print(f"✅ Total rows loaded: {len(df)}")


# Step 2: Assign headers
df.columns = [f"Column{i+1}" for i in range(len(df.columns))]

# Step 3: Drop empty columns
df = df.dropna(axis=1, how='all')

# Step 4: Detect DATE vs NON-DATE in Column2
df["is_date"] = pd.to_datetime(df["Column2"], errors='coerce').notna()

# Step 5: Create account_code ONLY from NON-DATE rows
df["account_code"] = df["Column2"].where(~df["is_date"])

# Step 6: Forward fill account_code
df["account_code"] = df["account_code"].ffill()

# Step 7: Remove account_code for header rows (non-date rows)
df.loc[~df["is_date"], "account_code"] = None

# Step 8: Drop helper column
df = df.drop(columns=["is_date"])

# Step 9: Move account_code after Column1
cols = df.columns.tolist()
cols.insert(1, cols.pop(cols.index("account_code")))
df = df[cols]

# Step 8: Drop empty rows
df = df.dropna(how='all')

# Step 10: Remove Column6 and Column11 (if they exist)
cols_to_drop = [col for col in ["Column6", "Column11", "Column22"] if col in df.columns]
df = df.drop(columns=cols_to_drop)



print(len(df.columns))
print(df.columns)
print(df.head(5))

# Step 10: Rename columns
new_columns = [
    "Transaction_Date",
    "Account_code",
    "Bill_number",
    "Account_Description",
    "Name",
    "GST_Type",
    "Debit_amount",
    "Credit_amount",
    "Dummy",
    "Balance"
]

df.columns = new_columns

# Step 11: Fill Account_Description downward
df["Account_Description"] = df["Account_Description"].ffill()

# Step 12: Remove Dummy column if exists
if "Dummy" in df.columns:
    df = df.drop(columns=["Dummy"])

# Step 13: Remove records where Transaction_Date is null
df = df[df["Transaction_Date"].notna()]

# Step 14: Remove records where Account_code is null
df = df[df["Account_code"].notna()]


# Step 16: Read COA file
coa_file = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Inbound\CCT_COA.xlsx"
df_coa = pd.read_excel(coa_file)

# Step 17: Standardize join columns (avoid mismatch issues)
df["Account_code"] = df["Account_code"].astype(str).str.strip()
df_coa["Account"] = df_coa["Account"].astype(str).str.strip()

# Step 18: Select required columns from COA
coa_columns = [
    "Account",
    "Item",
    "Item_type",
    "Category",
    "Subcategory",
    "Cashflow_Category",
    "Cashflow_Subcategory",
    "PNL_STMT_CATEGORY",
    "PNL_STMT_SUBCATEGORY",
    "PNL_STMT_MAIN_CATEGORY",
    "Project"
]

df_coa = df_coa[coa_columns]

# Step 19: Perform LEFT JOIN
df = df.merge(
    df_coa,
    how="left",
    left_on="Account_code",
    right_on="Account"
)

# Step 20: Drop duplicate join column
df = df.drop(columns=["Account"], errors="ignore")

# Step 21: Create BS_PNL_Flag_Final column
df["BS_PNL_Flag_Final"] = None

#df.loc[
    #(df["Item"] == "PNL") & (df["Name"].astype(str).str.strip().str.upper() != "OPEN"),
   # "BS_PNL_Flag_Final"
#] = "YES"

df.loc[
    (df["Item"].astype(str).str.strip().str.upper().isin(["PNL", "BS"])) &
    (df["Name"].astype(str).str.strip().str.upper() != "OPEN"),
    "BS_PNL_Flag_Final"
] = "YES"


df["Debit_amount"] = pd.to_numeric(df["Debit_amount"], errors="coerce").fillna(0)
df["Credit_amount"] = pd.to_numeric(df["Credit_amount"], errors="coerce").fillna(0)


# Step 22: Create Amount column
df["Amount"] = None

# Condition for PNL records with YES flag
condition = (
    (df["Item"].astype(str).str.strip().str.upper() == "PNL") &
    (df["BS_PNL_Flag_Final"] == "YES")
)

# Income logic → Credit - Debit
df.loc[
    condition & (df["Item_type"].astype(str).str.strip().str.upper() == "INCOME"),
    "Amount"
] = df["Credit_amount"] - df["Debit_amount"]

# Expense logic → Debit - Credit
df.loc[
    condition & (df["Item_type"].astype(str).str.strip().str.upper() == "EXPENSE"),
    "Amount"
] = df["Debit_amount"] - df["Credit_amount"]

# Step 24: Create final column mapping

final_mapping = {
    "Subsidiary_Name": None,
    "Order Type": None,
    "Date": "Transaction_Date",
    "As-Of Date": None,
    "Period": None,
    "Tax Period": None,
    "Type": None,
    "Document Number": None,
    "Name": "Name",
    "Expense_Account": "Account_code",
    "Expense_Account_category": None,
    "Expense_Account_Bankflow": None,
    "Expense_Account_Bankflow_short": None,
    "Expense_Account_Bankflow_final": None,
    "Expense_Account_Bankflow_short_final": None,
    "expense_account_bankflow_max": None,
    "Memo": None,
    "Amount": "Amount",
    "Account_Description": "Account_Description",
    "Account": "Account_code",
    "BS_PNL_Flag_Final": "BS_PNL_Flag_Final",
    "Item": "Item",
    "Item_type": "Item_type",
    "Item_Category": "Item_Category",
    "Category": "Category",
    "Subcategory": "Subcategory",
    "Cashflow_Subcategory": "Cashflow_Subcategory",
    "TRANSACTION APPROVAL STATUS": None,
    "ONLINE PAYMENT REF": None,
    "Description": None,
    "Status": None,
    "Vendor_Bill_number": "Bill_number",
    "Vendor_payment_number_multi": None,
    "Vendor_Bill_number_multi": None,
    "Vendor_payment_number": None,
    "Payment_Account": None,
    "Paid_date": None,
    "key": None,
    "key1": None,
    "bankflow_final": None,
    "amount_paid": None,
    "amount_bankflow": None,
    "Project Class": "Project",
    "posting_period": None,
    "Subsidiary_clean": None,
    "amount_final": None,
    "Balance_from_GL": "Balance",
    "transaction_amount_GL": None,
    "remarks": None,
    "Cashflow_inflowtype": None,
    "Cashflow_outflowtype": None,
    "Cashflow_type": None,
    "Cashflow_Category": "Cashflow_Category",
    "Cashflow_Category_first": "Cashflow_Category",
    "cashflow_first_subcategory": "Cashflow_Subcategory",
    "PNL_STMT_MAIN_CATEGORY": "PNL_STMT_MAIN_CATEGORY",
    "PNL_STMT_CATEGORY": "PNL_STMT_CATEGORY",
    "PNL_STMT_SUBCATEGORY": "PNL_STMT_SUBCATEGORY"
}

# Create final dataframe
final_df = pd.DataFrame()

for target_col, source_col in final_mapping.items():
    if source_col and source_col in df.columns:
        final_df[target_col] = df[source_col]
    else:
        final_df[target_col] = None

# Add remaining columns (not in mapping) at the end
remaining_cols = [col for col in df.columns if col not in final_mapping.values()]

for col in remaining_cols:
    final_df[col] = df[col]

# Step 26: Populate Subsidiary_Name
final_df["Subsidiary_Name"] = "CCT"

# Convert Transaction_Date to date (remove time part)
final_df["Date"] = pd.to_datetime(final_df["Date"], errors="coerce").dt.date

# Ensure numeric conversion (recommended)
final_df["Debit_amount"] = pd.to_numeric(final_df["Debit_amount"], errors="coerce").fillna(0)
final_df["Credit_amount"] = pd.to_numeric(final_df["Credit_amount"], errors="coerce").fillna(0)
final_df["Balance_from_GL"] = pd.to_numeric(final_df["Balance_from_GL"], errors="coerce").fillna(0)

# Condition for BS records
bs_condition = final_df["Item"].astype(str).str.strip().str.upper() == "BS"

# Case 1: BS and Flag = YES → Credit - Debit
final_df.loc[
    bs_condition & (final_df["BS_PNL_Flag_Final"] == "YES"),
    "Amount"
] = final_df["Debit_amount"] - final_df["Credit_amount"]

# Case 2: BS and Flag = NULL → Balance_from_GL (Balance)
final_df.loc[
    bs_condition & (final_df["BS_PNL_Flag_Final"].isna()),
    "Amount"
] = final_df["Balance_from_GL"]


final_df.loc[
    (final_df["Subcategory"].astype(str).str.strip().str.upper() == "CASH AT BANK") &
    (final_df["Name"].astype(str).str.strip().str.upper() != "OPEN"),
    "bankflow_final"
] = "YES"


# Ensure consistent format for matching
final_df["Vendor_Bill_number"] = final_df["Vendor_Bill_number"].astype(str).str.strip()

# Source: records where bankflow_final != YES
source_df = final_df[
    final_df["bankflow_final"].astype(str).str.strip().str.upper() != "YES"
][[
    "Vendor_Bill_number",
    "Cashflow_Category_first",
    "cashflow_first_subcategory"
]].dropna(subset=["Vendor_Bill_number"])

# Remove duplicates (keep first occurrence)
source_df = source_df.drop_duplicates(subset=["Vendor_Bill_number"])

# Merge into final_df (only to enrich)
final_df = final_df.merge(
    source_df,
    on="Vendor_Bill_number",
    how="left",
    suffixes=("", "_src")
)

# Update ONLY for bankflow_final = YES
condition = final_df["bankflow_final"].astype(str).str.strip().str.upper() == "YES"

final_df.loc[condition, "Cashflow_Category_first"] = final_df.loc[
    condition, "Cashflow_Category_first_src"
]

final_df.loc[condition, "cashflow_first_subcategory"] = final_df.loc[
    condition, "cashflow_first_subcategory_src"
]

# Drop helper columns
final_df = final_df.drop(columns=[
    "Cashflow_Category_first_src",
    "cashflow_first_subcategory_src"
], errors="ignore")

print("✅ Bankflow records enriched using Vendor_Bill_number mapping")


final_df["Amount"] = pd.to_numeric(final_df["Amount"], errors="coerce").fillna(0)

# Copy Amount → amount_bankflow where bankflow_final = YES
final_df.loc[
    final_df["bankflow_final"].astype(str).str.strip().str.upper() == "YES",
    "amount_bankflow"
] = final_df["Amount"]




# ==============================
# Step 28: Create Balance sheet File
# ==============================

# ==============================
# Step 1: Prepare Date + Month
# ==============================
final_df["Date"] = pd.to_datetime(final_df["Date"], errors="coerce")
final_df["Month"] = final_df["Date"].dt.to_period("M").astype(str)

# ==============================
# Step 2: Movement Data (exclude OPEN)
# ==============================
movement_df = final_df[
    (final_df["Item"].astype(str).str.strip().str.upper() == "BS") &
    (final_df["Name"].astype(str).str.strip().str.upper() != "OPEN")
].copy()

movement_df["Month"] = movement_df["Date"].dt.to_period("M").astype(str)

movement_df = movement_df.groupby(["Expense_Account", "Month"]).agg(
    Net=("Amount", "sum"),
    Inflow=("Amount", lambda x: x[x > 0].sum()),
    Outflow=("Amount", lambda x: x[x < 0].sum())
).reset_index()

movement_df["Inflow"] = movement_df["Inflow"].fillna(0)
movement_df["Outflow"] = movement_df["Outflow"].fillna(0)

# ==============================
# Step 3: Opening per Account
# ==============================
opening_df = final_df[
    (final_df["Item"].astype(str).str.strip().str.upper() == "BS") &
    (final_df["Name"].astype(str).str.strip().str.upper() == "OPEN")
].copy()

opening_df = opening_df.sort_values(by=["Expense_Account", "Date"])
opening_df = opening_df.groupby("Expense_Account", as_index=False).first()

opening_df["Month"] = opening_df["Date"].dt.to_period("M").astype(str)

opening_df = opening_df[[
    "Expense_Account", "Category", "Subcategory",
    "Month", "Balance_from_GL"
]]

opening_df = opening_df.rename(columns={
    "Balance_from_GL": "Opening"
})

# ==============================
# Step 4: Build FULL month range per account
# ==============================
all_accounts = final_df["Expense_Account"].dropna().unique()

all_months = pd.period_range(
    start=final_df["Date"].min().to_period("M"),
    end=final_df["Date"].max().to_period("M"),
    freq="M"
).astype(str)

full_grid = pd.MultiIndex.from_product(
    [all_accounts, all_months],
    names=["Account", "Month"]
).to_frame(index=False)

opening_df["Month"] = opening_df["Month"].astype(str)
movement_df["Month"] = movement_df["Month"].astype(str)






# ==============================
# Step 5: Merge Opening
# ==============================
opening_df = opening_df.rename(columns={"Month": "Opening_Month"})

full_df = full_grid.merge(
    opening_df,
    left_on=["Account"],
    right_on=["Expense_Account"],
    how="left"
)

full_df = full_df.drop(columns=["Expense_Account"])

# keep only valid starting month onward per account
#full_df = full_df[full_df["Month"] >= full_df["Month"]]

# fill missing category
full_df["Category"] = full_df["Category"].ffill()
full_df["Subcategory"] = full_df["Subcategory"].ffill()

#movement_df["Date"] = pd.to_datetime(movement_df["Date"], errors="coerce")
#movement_df["Month"] = movement_df["Date"].dt.to_period("M").astype(str)

print("full_df columns:", full_df.columns)
print("movement_df columns:", movement_df.columns)

# ==============================
# Step 6: Merge movement data
# ==============================
full_df = full_df.merge(
    movement_df[["Expense_Account", "Month", "Net", "Inflow", "Outflow"]],
    left_on=["Account", "Month"],
    right_on=["Expense_Account", "Month"],
    how="left"
)

full_df = full_df.drop(columns=["Expense_Account"])

full_df["Net"] = full_df["Net"].fillna(0)
full_df["Inflow"] = full_df["Inflow"].fillna(0)
full_df["Outflow"] = full_df["Outflow"].fillna(0)

# ==============================
# Step 7: Sort for logic
# ==============================
#full_df = full_df.sort_values(["Account", "Month"])


full_df["Month"] = pd.to_datetime(full_df["Month"], format="%Y-%m")

full_df = full_df.sort_values(by=["Account", "Month"])


#numeric_cols = ["Opening", "Closing", "Inflow", "Outflow", "Net"]
numeric_cols = ["Opening", "Inflow", "Outflow", "Net"]

for col in numeric_cols:
    full_df[col] = pd.to_numeric(full_df[col], errors="coerce").fillna(0.0)
    full_df[col] = full_df[col].astype(float).round(4)



def roll_forward(group):
    group = group.sort_values("Month").copy().reset_index(drop=True)

    # Keep original opening from first row
    opening = group.loc[0, "Opening"]

    for i in range(len(group)):
        if i == 0:
            group.loc[i, "Opening"] = opening
        else:
            #group.loc[i, "Opening"] = group.loc[i-1, "Closing"]
            group.loc[i, "Opening"] = round(float(group.loc[i-1, "Closing"]), 4)

        group.loc[i, "Closing"] = group.loc[i, "Opening"] + group.loc[i, "Net"]
        #group.loc[i, "Closing"] = round(Closing, 4) # additional logic 

    return group

full_df = full_df.groupby("Account", group_keys=False).apply(roll_forward)

full_df["Month"] = full_df["Month"].dt.strftime("%Y-%m")



# Output file path
opening_file = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Outbound\BalanceSheet_CCT.csv"

# ==============================
# Step 9: Final structure
# ==============================
full_df.insert(0, "Company", "CCT")

required_cols = [
    "Company",
    "Account",
    "Category",
    "Subcategory",
    "Month",
    "Opening",
    "Closing",
    "Inflow",
    "Outflow",
    "Net",
    
]



final_output = full_df.reindex(columns=required_cols)



# Save file
final_output.to_csv(opening_file, index=False)

print("✅ Opening balance file created at:", final_output)




# Step 29: Save
final_df.to_csv(output_file, index=False)

print("✅ File processed successfully. Output saved at:", output_file)


# ==============================
# Step 30: Create BankFlow_CCT.csv
# ==============================

# Filter only CASH AT BANK
bankflow_df = final_output[
    final_output["Subcategory"].astype(str).str.strip().str.upper() == "CASH AT BANK"
].copy()

# Select required columns
bankflow_df = bankflow_df[[
    "Company",
    "Account",
    "Month",
    "Opening",
    "Closing",
    "Inflow",
    "Outflow",
    "Net"
]]

# Output path
bankflow_file = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Outbound\BankFlow_CCT.csv"

# Save file
bankflow_df.to_csv(bankflow_file, index=False)

print("✅ BankFlow file created at:", bankflow_file)


# === Step XX: Aggregation by Subsidiary, Month, Cashflow_Category_first ===
print("📊 Creating aggregated bankflow summary using Cashflow_Category_first...")

# Use only bankflow_final = YES
set2 = final_df[final_df["bankflow_final"].astype(str).str.upper() == "YES"].copy()

# Convert Date → datetime
set2['Date'] = pd.to_datetime(set2['Date'], errors='coerce', dayfirst=True)

# Derive Month (YYYY-MM)
set2['Month'] = set2['Date'].dt.to_period('M')

# Clean category column
set2['Cashflow_Category_first'] = set2['Cashflow_Category_first'].fillna("").astype(str).str.strip()



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

cashflow_output_file = r"C:\Broad_field_holdings\Non_Netsuite\CCT\Outbound\Aggegated_cashflow_categerization_file_CCT.csv"


# Save file
final_df.to_csv(cashflow_output_file, index=False)

print("✅ Aggregated bankflow summary created: aggregated_bankflow_with_total.csv")

