import pandas as pd
import numpy as np
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


# Sheet1 formatting copied from the attached expected Auto Pricing workbook.
# Only output presentation is changed; pricing/action calculations stay untouched.
SHEET1_BASE_WIDTHS = {
    'A': 9.0,
    'B': 6.28515625,
    'C': 5.5703125,
    'D': 9.85546875,
    'E': 11.140625,
    'F': 5.5703125,
    'G': 4.5703125,
    'H': 3.85546875,
    'I': 4.7109375,
    'J': 13.0,
    'K': 5.28515625,
    'L': 6.0,
    'M': 7.7109375,
    'N': 10.0,
    'O': 22.5703125,
    'P': 7.7109375,
    'Q': 13.0,
    'R': 33.5703125,
    'S': 20.28515625,
    'T': 30.7109375,
    'U': 10.5703125,
    'V': 12.7109375,
    'W': 16.85546875,
}


def apply_sheet1_base_format(ws):
    """Apply the same base table format as the expected workbook's Sheet1."""
    thin = Side(style='thin', color='000000')
    table_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(name='Calibri', size=11, bold=True)
    body_font = Font(name='Calibri', size=11, bold=False)
    header_alignment = Alignment(horizontal='center', vertical='top')
    body_alignment = Alignment(horizontal='center')

    # Match the expected workbook widths for the first 23 requested columns.
    for col_letter, width in SHEET1_BASE_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # Extra output columns must remain behind the 23 requested columns.
    # Give them the same table format and a practical content-based width.
    for col_idx in range(24, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        header = ws.cell(row=1, column=col_idx).value
        max_len = len(str(header)) if header is not None else 0

        # Sample rows for width calculation to avoid slowing very large files.
        sample_last_row = min(ws.max_row, 1000)
        for row_idx in range(2, sample_last_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is not None:
                max_len = max(max_len, len(str(value)))

        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)

    # Header formatting.
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = table_border

    # Data formatting.
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = body_font
            cell.alignment = body_alignment
            cell.border = table_border


def format_weight(val):
    """Ensures carat weight is formatted to exactly two decimal places."""
    try:
        return f"{float(val):.2f}"
    except (ValueError, TypeError):
        return val


def normalize_base_columns(df):
    """Standardizes core column names to prevent errors from trailing spaces or aliases."""
    col_map = {}
    for col in df.columns:
        clean = str(col).strip().lower()
        if clean == 'tag':
            col_map[col] = 'Tag'
        elif clean in ['col', 'color', 'color name']:
            col_map[col] = 'Color name'
        elif clean in ['carat_grp', 'carat grp']:
            col_map[col] = 'Carat_Grp'
        elif clean == 'shp':
            col_map[col] = 'Shp'
        elif clean == 'qua':
            col_map[col] = 'Qua'
        elif clean == 'lab':
            col_map[col] = 'Lab'
        elif clean == 'flo':
            col_map[col] = 'Flo'
        elif clean == 'expected':
            col_map[col] = 'Expected'

    df.rename(columns=col_map, inplace=True)
    return df


def ensure_tag_column(df, filename):
    """Checks if 'Tag' exists; if not, safely reconstructs it or halts with clear instructions."""
    if 'Tag' not in df.columns:
        req_cols = ['Shp', 'Carat_Grp', 'Color name', 'Qua', 'Lab', 'Flo']
        missing = [c for c in req_cols if c not in df.columns]

        if missing:
            raise KeyError(
                f"\n❌ STOPPING: '{filename}' is missing the 'Tag' column.\n"
                f"The script tried to build it automatically, but it is also missing: {missing}\n"
                f"Please ensure your export includes the 'Tag' column "
                f"(or 'Carat_Grp', 'Color name', etc.) and run again.\n"
            )

        df['Tag'] = (
            df['Shp'].astype(str) + '-' +
            df['Carat_Grp'].astype(str) + '-' +
            df['Color name'].astype(str) + '-' +
            df['Qua'].astype(str) + '-' +
            df['Lab'].astype(str) + '-' +
            df['Flo'].astype(str)
        )

    return df


def is_valid_party(row):
    if pd.isna(row['Action_Party_Name']) or pd.isna(row['Master_Sold_Party_List']):
        return False

    action_party = str(row['Action_Party_Name']).strip()
    sold_parties_list = [
        p.strip()
        for p in str(row['Master_Sold_Party_List']).split(',')
    ]

    return action_party in sold_parties_list


def process_stock_upgrade(
    stock_file,
    sales_180_file,
    sales_120_file,
    action_files_dict
):

    # 1. Load Data
    df_stock = pd.read_excel(stock_file)
    df_180 = pd.read_excel(sales_180_file)
    df_120 = pd.read_excel(sales_120_file)

    # Keep untouched source data ONLY for the final multi-sheet output workbook.
    # The processing dataframes below continue to use the existing logic unchanged.
    df_180_output = df_180.copy()
    df_120_output = df_120.copy()
    action_output_sheets = {}

    # 1B. Normalize Columns and Ensure 'Tag' Exists
    df_stock = normalize_base_columns(df_stock)
    df_180 = normalize_base_columns(df_180)
    df_120 = normalize_base_columns(df_120)

    df_stock = ensure_tag_column(df_stock, "Stock File")
    df_180 = ensure_tag_column(df_180, "180-Day Sales File")
    df_120 = ensure_tag_column(df_120, "120-Day Sales File")

    if 'Carat' in df_stock.columns:
        df_stock['Carat'] = df_stock['Carat'].apply(format_weight)

    # 2. Process 180-Day Sales Data (Master Buyer List)
    agg_180 = df_180.groupby('Tag').agg(
        sale_count=('Tag', 'count'),
        Unique_Buyer_Count=('Sold Party', 'nunique'),
        Master_Sold_Party_List=(
            'Sold Party',
            lambda x: ', '.join(x.dropna().astype(str).unique())
        )
    ).reset_index()

    # 3. Process 120-Day Sales Data (Weighted Average Logic)
    if 'Carat' in df_120.columns and 'Amt' in df_120.columns:

        df_120['Carat_Num'] = pd.to_numeric(
            df_120['Carat'],
            errors='coerce'
        ).fillna(0)

        df_120['Amt_Num'] = pd.to_numeric(
            df_120['Amt'],
            errors='coerce'
        ).fillna(0)

        agg_120 = df_120.groupby('Tag').agg(
            Total_Carat=('Carat_Num', 'sum'),
            Total_Amt=('Amt_Num', 'sum')
        ).reset_index()

        agg_120['Sold_Avg_Rate_120days'] = np.where(
            agg_120['Total_Carat'] > 0,
            agg_120['Total_Amt'] / agg_120['Total_Carat'],
            0
        )

        agg_120 = agg_120[
            ['Tag', 'Sold_Avg_Rate_120days']
        ]

    else:
        print(
            "Warning: 'Carat' or 'Amt' column not found in 120-day file."
        )

        agg_120 = pd.DataFrame(
            columns=['Tag', 'Sold_Avg_Rate_120days']
        )

    # 4. Merge Initial Sales Data into Stock
    df_upgraded = pd.merge(
        df_stock,
        agg_180,
        on='Tag',
        how='left'
    )

    df_upgraded = pd.merge(
        df_upgraded,
        agg_120,
        on='Tag',
        how='left'
    )

    # 5. Build Cross-Verification DataFrame base
    verify_cols = [
        'Serial No',
        'Tag',
        'sale_count',
        'Master_Sold_Party_List',
        'Unique_Buyer_Count'
    ]

    available_cols = [
        col for col in verify_cols
        if col in df_upgraded.columns
    ]

    df_verify = df_upgraded[available_cols].copy()

    # GUARANTEE COLUMNS EXIST
    df_upgraded['Action_count'] = 0
    df_upgraded['ALL_Action_count'] = 0

    df_verify['Action_count'] = 0
    df_verify['ALL_Action_count'] = 0

    for action_name in action_files_dict.keys():

        df_upgraded[f"{action_name}_Count"] = 0
        df_upgraded[f"{action_name}_All_Count"] = 0

        df_verify[f"{action_name}_Count"] = 0
        df_verify[f"{action_name}_All_Count"] = 0

    master_actions_accumulator = []

    # 6. Process Action Files
    for action_name, file_path in action_files_dict.items():

        col_matched = f"{action_name}_Count"
        col_all = f"{action_name}_All_Count"

        try:
            df_action = pd.read_excel(file_path)

            # Keep the original action-file columns for the output workbook tab.
            # This copy is not used by the calculation logic.
            action_output_sheets[action_name] = df_action.copy()

            # --- DYNAMIC FUZZY COLUMN MATCHER ---
            action_party_col = None
            action_serial_col = None

            for col in df_action.columns:

                clean_col = (
                    str(col)
                    .lower()
                    .replace(' ', '')
                    .replace('_', '')
                )

                if clean_col in [
                    'soldparty',
                    'party',
                    'PARTY_NAME',
                    'PARTY_COMPANY_NAME'
                ]:
                    action_party_col = col

                elif clean_col in [
                    'serialno',
                    'stoneid',
                    'stockno',
                    'id',
                    'STONE_ID'
                ]:
                    action_serial_col = col

            if not action_party_col or not action_serial_col:
                print(
                    f"Warning: Missing required columns in "
                    f"'{action_name}'. Skipping."
                )
                continue

            # Force the action file to use standard names
            # instead of changing the master files
            df_action.rename(
                columns={
                    action_party_col: 'Action_Party_Name',
                    action_serial_col: 'Serial No'
                },
                inplace=True
            )

            # Save these actions to the master accumulator for later
            master_actions_accumulator.append(
                df_action[
                    ['Serial No', 'Action_Party_Name']
                ].copy()
            )

            # Calculate "ALL COUNT" for this specific file
            agg_action_all = (
                df_action
                .groupby('Serial No')['Action_Party_Name']
                .nunique()
            )

            df_upgraded[col_all] = (
                df_upgraded['Serial No']
                .map(agg_action_all)
                .fillna(0)
                .astype(int)
            )

            df_verify[col_all] = (
                df_verify['Serial No']
                .map(agg_action_all)
                .fillna(0)
                .astype(int)
            )

            # Calculate "MATCHED COUNT" for this specific file
            df_action_merged = pd.merge(
                df_action[
                    ['Serial No', 'Action_Party_Name']
                ],
                df_upgraded[
                    ['Serial No', 'Master_Sold_Party_List']
                ],
                on='Serial No',
                how='inner'
            )

            if not df_action_merged.empty:

                valid_mask = df_action_merged.apply(
                    is_valid_party,
                    axis=1
                )

                valid_actions = df_action_merged[
                    valid_mask
                ]

                if not valid_actions.empty:

                    agg_action_matched = (
                        valid_actions
                        .groupby('Serial No')['Action_Party_Name']
                        .nunique()
                    )

                    df_upgraded[col_matched] = (
                        df_upgraded['Serial No']
                        .map(agg_action_matched)
                        .fillna(0)
                        .astype(int)
                    )

                    df_verify[col_matched] = (
                        df_verify['Serial No']
                        .map(agg_action_matched)
                        .fillna(0)
                        .astype(int)
                    )

        except Exception as e:
            print(
                f"Error processing '{action_name}' file: {e}"
            )

    # =====================================================================
    # 7. CROSS-ACTION COMBINED TOTALS
    # =====================================================================

    if master_actions_accumulator:

        # Mash every action file together into one huge dataframe
        combined_actions_df = pd.concat(
            master_actions_accumulator,
            ignore_index=True
        )

        combined_actions_df.dropna(
            subset=['Serial No', 'Action_Party_Name'],
            inplace=True
        )

        # 7A. ALL_Action_count
        agg_combined_all = (
            combined_actions_df
            .groupby('Serial No')['Action_Party_Name']
            .nunique()
        )

        df_upgraded['ALL_Action_count'] = (
            df_upgraded['Serial No']
            .map(agg_combined_all)
            .fillna(0)
            .astype(int)
        )

        df_verify['ALL_Action_count'] = (
            df_verify['Serial No']
            .map(agg_combined_all)
            .fillna(0)
            .astype(int)
        )

        # 7B. Action_count
        combined_merged = pd.merge(
            combined_actions_df,
            df_upgraded[
                ['Serial No', 'Master_Sold_Party_List']
            ],
            on='Serial No',
            how='inner'
        )

        if not combined_merged.empty:

            valid_mask_combined = combined_merged.apply(
                is_valid_party,
                axis=1
            )

            valid_combined_actions = combined_merged[
                valid_mask_combined
            ]

            if not valid_combined_actions.empty:

                agg_combined_matched = (
                    valid_combined_actions
                    .groupby('Serial No')['Action_Party_Name']
                    .nunique()
                )

                df_upgraded['Action_count'] = (
                    df_upgraded['Serial No']
                    .map(agg_combined_matched)
                    .fillna(0)
                    .astype(int)
                )

                df_verify['Action_count'] = (
                    df_verify['Serial No']
                    .map(agg_combined_matched)
                    .fillna(0)
                    .astype(int)
                )

    # Clean up Master List name
    df_upgraded.rename(
        columns={
            'Master_Sold_Party_List': 'Sold_Party'
        },
        inplace=True
    )

    df_verify.rename(
        columns={
            'Master_Sold_Party_List': 'Sold_Party'
        },
        inplace=True
    )

    # --- Pricing Logic: Calculate 'Diff' and 'Remark' ---
    df_upgraded['Sold_Avg_Rate_120days'] = (
        df_upgraded['Sold_Avg_Rate_120days']
        .fillna(0)
    )

    if 'Expected' in df_upgraded.columns:

        df_upgraded['Expected_Num'] = pd.to_numeric(
            df_upgraded['Expected'],
            errors='coerce'
        ).fillna(0)

        # Diff
        df_upgraded['Diff'] = np.where(
            df_upgraded['Expected_Num'] > 0,
            100 - (
                (
                    df_upgraded['Sold_Avg_Rate_120days']
                    / df_upgraded['Expected_Num']
                ) * 100
            ),
            0
        ).round(2)

        # Remark: Nested np.where to replicate nested Excel IF statement
        df_upgraded['Remark'] = np.where(
            df_upgraded['Sold_Avg_Rate_120days'] == 0,
            'No',
            np.where(
                df_upgraded['Expected_Num']
                > df_upgraded['Sold_Avg_Rate_120days'],
                'Yes',
                'No'
            )
        )

        df_upgraded.drop(
            columns=['Expected_Num'],
            inplace=True
        )

    else:
        print(
            "\nWarning: Could not find 'Expected' column "
            "to calculate Diff and Remark."
        )

    # Fill NaNs for standard numerical columns
    df_upgraded['sale_count'] = (
        df_upgraded['sale_count']
        .fillna(0)
        .astype(int)
    )

    df_upgraded['Unique_Buyer_Count'] = (
        df_upgraded['Unique_Buyer_Count']
        .fillna(0)
        .astype(int)
    )

    # =====================================================================
    # 8. OUTPUT COLUMN SEQUENCE
    # =====================================================================

    priority_columns = [
        'Serial No',
        'Shp',
        'Carat',
        'Carat_Grp',
        'Color name',
        'Qua',
        'Cut',
        'Pol',
        'Sym',
        'Lab',
        'Flo',
        'GRap',
        'GPer',
        'Expected',
        'Sold_Avg_Rate_120days',
        'Diff',
        'Remark',
        'Tag',
        'Unique_Buyer_Count',
        'Sold_Party',
        'sale_count',
        'Action_count',
        'ALL_Action_count'
    ]

    # Keep requested columns first, but only if they exist
    existing_priority_columns = [
        col
        for col in priority_columns
        if col in df_upgraded.columns
    ]

    # Keep all other existing columns after the requested columns
    remaining_columns = [
        col
        for col in df_upgraded.columns
        if col not in existing_priority_columns
    ]

    # Apply final column sequence
    df_upgraded = df_upgraded[
        existing_priority_columns + remaining_columns
    ]

    # 9. Export Files
    # Auto Pricing output is now ONE workbook matching the requested structure:
    # Sheet1, Sale_segment, Avg_Rate, Hold, Wishlist, View, Offer, Cart.
    # Media can still be used in the existing calculation logic, but it is not
    # added as an output tab because it is not part of the expected workbook.
    output_file = r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Auto pricing revision.xlsx"
    cross_verify_file = r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Cross_Verify_Sales_Actions.xlsx"

    # Use an explicit ExcelWriter instead of a `with` block.
    # This avoids the indentation/parser issue reported by the app while
    # preserving the exact same workbook output.
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    # Main processed output. The 23 requested columns are already first,
    # and every other existing output column is kept after them.
    df_upgraded.to_excel(writer, sheet_name='Sheet1', index=False)

    # Apply the same base formatting/layout as the attached expected file.
    apply_sheet1_base_format(writer.book['Sheet1'])

    # IMPORTANT SOURCE MAPPING:
    # Sale_segment = untouched rows/columns from sales_180_file.xlsx
    # Avg_Rate     = untouched rows/columns from sales_120_file.xlsx
    df_180_output.to_excel(writer, sheet_name='Sale_segment', index=False)
    df_120_output.to_excel(writer, sheet_name='Avg_Rate', index=False)

    requested_action_sheet_order = [
        'Hold',
        'Wishlist',
        'View',
        'Offer',
        'Cart'
    ]

    for sheet_name in requested_action_sheet_order:
        if sheet_name in action_output_sheets:
            action_output_sheets[sheet_name].to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    # Explicitly save and close the workbook.
    writer.close()

    # Keep the existing cross-verification output as a separate file.
    df_verify.to_excel(cross_verify_file, index=False)

    print(
        "\n✅ Processing complete. Files saved: "
        "'Auto pricing revision.xlsx' and 'Cross_Verify_Sales_Actions.xlsx'"
    )


# --- Execution ---
action_files = {
    "View": r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\View.xlsx",
    "Hold": r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Hold.xlsx",
    "Offer": r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Offer.xlsx",
    "Wishlist": r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Wishlist.xlsx",
    "Cart": r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Cart.xlsx",
    "Media": r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\media.xlsx" # Assuming you added a Media file based on your error log!
}

process_stock_upgrade(
    r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\Stock file.xlsx",
    r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\sales_180_file.xlsx",
    r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026\sales_120_file.xlsx",
    action_files
)
