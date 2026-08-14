import os
from io import BytesIO

import pandas as pd
import numpy as np
from openpyxl.styles import Alignment, Border, Font, Side


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
    action_files_dict,
    output_file=None,
    cross_verify_file=None
):

    # Backward compatibility with the original 4-argument call.
    # If no output path is supplied, return the generated workbook in memory.
    output_buffer = BytesIO() if output_file is None else None

    # 1. Load Data
    # Keep untouched copies for the source tabs in the final workbook.
    df_stock_source = pd.read_excel(stock_file)
    df_180_source = pd.read_excel(sales_180_file)
    df_120_source = pd.read_excel(sales_120_file)

    # Work on copies so normalization does not change source-sheet headings.
    df_stock = df_stock_source.copy()
    df_180 = df_180_source.copy()
    df_120 = df_120_source.copy()

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
    raw_action_sheets = {}

    # 6. Process Action Files
    for action_name, file_path in action_files_dict.items():

        col_matched = f"{action_name}_Count"
        col_all = f"{action_name}_All_Count"

        try:
            df_action = pd.read_excel(file_path)
            raw_action_sheets[action_name] = df_action.copy()

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
                    'partyname',
                    'partycompanyname'
                ]:
                    action_party_col = col

                elif clean_col in [
                    'serialno',
                    'stoneid',
                    'stockno',
                    'id'
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

    # The reference output contains exactly these 23 columns.
    # Do NOT append temporary per-action helper columns to Sheet1.
    df_upgraded = df_upgraded[existing_priority_columns]

    # =====================================================================
    # 9. EXPORT - MATCH THE REFERENCE WORKBOOK STRUCTURE
    # =====================================================================

    if output_file is not None:
        os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)

    # Reference workbook sheet order:
    # Sheet1, Sale_segment, Avg_Rate, Hold, Wishlist, View, Offer, Cart
    output_target = output_buffer if output_buffer is not None else output_file
    writer = pd.ExcelWriter(output_target, engine='openpyxl')
    df_upgraded.to_excel(writer, sheet_name='Sheet1', index=False)
    df_180_source.to_excel(writer, sheet_name='Sale_segment', index=False)
    df_120_source.to_excel(writer, sheet_name='Avg_Rate', index=False)

    reference_action_order = [
        'Hold',
        'Wishlist',
        'View',
        'Offer',
        'Cart'
    ]

    for action_name in reference_action_order:
        if action_name in raw_action_sheets:
            raw_action_sheets[action_name].to_excel(
                writer,
                sheet_name=action_name,
                index=False
            )

    # Format Sheet1 to match the attached reference output.
    ws = writer.book['Sheet1']
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    reference_widths = {
        'A': 9.0, 'B': 6.29, 'C': 5.57, 'D': 9.86, 'E': 11.14,
        'F': 5.57, 'G': 4.57, 'H': 3.86, 'I': 4.71, 'J': 13.0,
        'K': 5.29, 'L': 6.0, 'M': 7.71, 'N': 10.0, 'O': 22.57,
        'P': 7.71, 'Q': 13.0, 'R': 33.57, 'S': 20.29, 'T': 30.71,
        'U': 10.57, 'V': 12.71, 'W': 16.86
    }

    for col_letter, width in reference_widths.items():
        ws.column_dimensions[col_letter].width = width

    header_font = Font(name='Calibri', size=11, bold=True)
    body_font = Font(name='Calibri', size=11)
    header_alignment = Alignment(horizontal='center', vertical='top')
    body_alignment = Alignment(horizontal='center')

    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.alignment = body_alignment
            cell.border = border
    writer.close()
    # Keep the cross-verification workbook as a separate optional output.
    if cross_verify_file:
        os.makedirs(os.path.dirname(cross_verify_file) or ".", exist_ok=True)
        df_verify.to_excel(cross_verify_file, index=False)

    if output_buffer is not None:
        output_buffer.seek(0)
        return output_buffer

    print(f"\n✅ Processing complete. Main output saved: {output_file}")
    if cross_verify_file:
        print(f"✅ Cross verification saved: {cross_verify_file}")

    return output_file


if __name__ == '__main__':
    # --- Execution / Paths ---
    base_dir = (
        r"C:\Users\bhoja\OneDrive - SJW SHIVAM JEWELS PRIVATE LIMITED"
        r"\Pricing Regular Related Demo 02-05-2026\Input 07-08-2026"
    )

    action_files = {
        'Hold': os.path.join(base_dir, 'Hold.xlsx'),
        'Wishlist': os.path.join(base_dir, 'Wishlist.xlsx'),
        'View': os.path.join(base_dir, 'View.xlsx'),
        'Offer': os.path.join(base_dir, 'Offer.xlsx'),
        'Cart': os.path.join(base_dir, 'Cart.xlsx')
    }

    process_stock_upgrade(
        os.path.join(base_dir, 'Stock file.xlsx'),
        os.path.join(base_dir, 'sales_180_file.xlsx'),
        os.path.join(base_dir, 'sales_120_file.xlsx'),
        action_files,
        output_file=os.path.join(
            base_dir,
            'Auto pricing revision 14-08-2026.xlsx'
        ),
        cross_verify_file=os.path.join(
            base_dir,
            'Cross_Verify_Sales_Actions.xlsx'
        )
    )
