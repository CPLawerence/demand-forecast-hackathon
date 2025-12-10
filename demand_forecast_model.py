import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import warnings
warnings.filterwarnings('ignore')

# File paths
base_path = r'C:\Users\JaeheeKim\Burrow\Burrow - 19. INVENTORY\06_Analysis + Reporting\_Jaehee\Hackathon'
catalog_file = f'{base_path}\\catalog_2025-12-09-1340.csv'
inventory_file = f'{base_path}\\on hand inventory_2025-12-09-1341.csv'
sales_file = f'{base_path}\\sku sales_2025-12-09-1347.csv'
on_order_file = f'{base_path}\\CZ On Order Sample Data.xlsx'
output_file = f'{base_path}\\Demand_Forecast_Inventory_Model.xlsx'

print("Loading data files...")

# Load catalog
catalog = pd.read_csv(catalog_file)
catalog = catalog.rename(columns={'SKU': 'SKU'})

# Load inventory
inventory = pd.read_csv(inventory_file)

# Load sales
sales = pd.read_csv(sales_file)
sales = sales.rename(columns={'COMPONENT_SKU': 'SKU'})
sales['ORDER_DATE'] = pd.to_datetime(sales['ORDER_DATE'])
sales['ORDER_MONTH'] = pd.to_datetime(sales['ORDER_MONTH'])

# Load on-order data
on_order = pd.read_excel(on_order_file)
print(f"On-order columns: {on_order.columns.tolist()}")

# Determine current date and forecast horizon
current_date = datetime(2025, 12, 1)
history_start = datetime(2023, 1, 1)
forecast_months = 6  # Forecast 6 months ahead

# Create month range for historical + forecast
all_months = pd.date_range(start=history_start, end=current_date + relativedelta(months=forecast_months), freq='MS')
historical_months = [m for m in all_months if m <= current_date]
forecast_months_list = [m for m in all_months if m > current_date]

print(f"Historical months: {len(historical_months)}, Forecast months: {len(forecast_months_list)}")

# Aggregate sales by SKU and month
sales_agg = sales.groupby(['SKU', 'ORDER_MONTH'])['UNITS_SOLD'].sum().reset_index()
sales_agg = sales_agg.rename(columns={'ORDER_MONTH': 'MONTH', 'UNITS_SOLD': 'SALES_DEMAND'})

# Create pivot table for sales
sales_pivot = sales_agg.pivot(index='SKU', columns='MONTH', values='SALES_DEMAND').fillna(0)

# Get all unique SKUs from catalog
all_skus = catalog['SKU'].unique()

# Calculate average monthly sales for each SKU (for forecasting)
def calculate_forecast(sku_sales_series, forecast_periods=6):
    """Calculate forecast using simple moving average of last 6 months with data"""
    # Get non-zero values
    non_zero_sales = sku_sales_series[sku_sales_series > 0]
    if len(non_zero_sales) >= 3:
        # Use average of last 6 months with sales
        avg_sales = non_zero_sales.tail(6).mean()
    elif len(non_zero_sales) > 0:
        avg_sales = non_zero_sales.mean()
    else:
        avg_sales = 0
    return avg_sales

# Merge catalog with inventory
sku_master = catalog.merge(inventory, on='SKU', how='left')
sku_master = sku_master.fillna(0)

# Process on-order data - identify the date/quantity columns
print("\nOn-order data sample:")
print(on_order.head())
print(f"\nOn-order dtypes:\n{on_order.dtypes}")

# Try to identify SKU and receipt columns in on-order data
on_order_cols = on_order.columns.tolist()
sku_col_candidates = [c for c in on_order_cols if 'SKU' in c.upper() or 'ITEM' in c.upper() or 'PRODUCT' in c.upper()]
date_col_candidates = [c for c in on_order_cols if 'DATE' in c.upper() or 'ETA' in c.upper() or 'RECEIPT' in c.upper() or 'ARRIVAL' in c.upper()]
qty_col_candidates = [c for c in on_order_cols if 'QTY' in c.upper() or 'QUANTITY' in c.upper() or 'UNITS' in c.upper() or 'ORDER' in c.upper()]

print(f"\nIdentified columns - SKU: {sku_col_candidates}, Date: {date_col_candidates}, Qty: {qty_col_candidates}")

# Process on-order data based on actual columns
on_order_by_sku_month = {}
if len(sku_col_candidates) > 0 and len(qty_col_candidates) > 0:
    sku_col = sku_col_candidates[0]
    qty_col = [c for c in qty_col_candidates if 'QTY' in c.upper()][0] if any('QTY' in c.upper() for c in qty_col_candidates) else qty_col_candidates[0]

    if len(date_col_candidates) > 0:
        # Prefer 'Land Date' for receipt timing
        date_col = next((c for c in date_col_candidates if 'LAND' in c.upper()), date_col_candidates[0])
        on_order['RECEIPT_MONTH'] = pd.to_datetime(on_order[date_col], errors='coerce')
        on_order['RECEIPT_MONTH'] = on_order['RECEIPT_MONTH'].dt.to_period('M').dt.to_timestamp()

        on_order_agg = on_order.groupby([sku_col, 'RECEIPT_MONTH'])[qty_col].sum().reset_index()
        on_order_agg = on_order_agg.rename(columns={sku_col: 'SKU', qty_col: 'ON_ORDER_QTY'})

        for _, row in on_order_agg.iterrows():
            key = (row['SKU'], row['RECEIPT_MONTH'])
            on_order_by_sku_month[key] = row['ON_ORDER_QTY']

    print(f"Processed {len(on_order_by_sku_month)} on-order records by SKU/month")

# Get unique planning categories
planning_categories = sku_master['PLANNING_CATEGORY'].dropna().unique()
planning_categories = [pc for pc in planning_categories if pc and str(pc).strip()]

print(f"\nFound {len(planning_categories)} planning categories")

# Create Excel writer
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

# Define formats
header_format = workbook.add_format({
    'bold': True,
    'bg_color': '#4472C4',
    'font_color': 'white',
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'
})

sku_header_format = workbook.add_format({
    'bold': True,
    'bg_color': '#4472C4',
    'font_color': 'white',
    'border': 1,
    'align': 'left',
    'valign': 'vcenter'
})

row_label_format = workbook.add_format({
    'bold': True,
    'bg_color': '#D9E2F3',
    'border': 1,
    'align': 'left'
})

number_format = workbook.add_format({
    'num_format': '#,##0',
    'border': 1,
    'align': 'center'
})

negative_format = workbook.add_format({
    'num_format': '#,##0',
    'border': 1,
    'align': 'center',
    'font_color': 'red',
    'bold': True
})

forecast_format = workbook.add_format({
    'num_format': '#,##0',
    'border': 1,
    'align': 'center',
    'bg_color': '#FFF2CC'  # Light yellow for forecasted values
})

forecast_negative_format = workbook.add_format({
    'num_format': '#,##0',
    'border': 1,
    'align': 'center',
    'bg_color': '#FFF2CC',
    'font_color': 'red',
    'bold': True
})

text_format = workbook.add_format({
    'border': 1,
    'align': 'left'
})

# Process each planning category
for category in planning_categories:
    print(f"\nProcessing category: {category}")

    # Filter SKUs for this category
    category_skus = sku_master[sku_master['PLANNING_CATEGORY'] == category].copy()

    if len(category_skus) == 0:
        continue

    # Calculate average sales for each SKU to sort
    sku_avg_sales = {}
    for sku in category_skus['SKU'].unique():
        if sku in sales_pivot.index:
            sku_sales = sales_pivot.loc[sku]
            avg = sku_sales[sku_sales > 0].mean() if len(sku_sales[sku_sales > 0]) > 0 else 0
        else:
            avg = 0
        sku_avg_sales[sku] = avg

    category_skus['AVG_SALES'] = category_skus['SKU'].map(sku_avg_sales)
    category_skus = category_skus.sort_values('AVG_SALES', ascending=False)

    # Create worksheet (truncate name if too long)
    sheet_name = str(category)[:31] if len(str(category)) > 31 else str(category)
    sheet_name = sheet_name.replace('/', '-').replace('\\', '-').replace('*', '').replace('?', '').replace('[', '').replace(']', '')

    worksheet = workbook.add_worksheet(sheet_name)

    # Define column structure
    sku_detail_cols = ['CATEGORY', 'SUB_CATEGORY', 'COLLECTION', 'SKU', 'SKU_DESCRIPTION', 'COLOR_NAME', 'SIZE']
    row_types = ['Sales Demand', 'Committed Qty', 'Backorder Qty', 'EOM Inventory', 'Projected EOM Inventory', 'Receipts (On-Order)']

    # Write headers
    # SKU detail headers
    for col_idx, col_name in enumerate(sku_detail_cols):
        worksheet.write(0, col_idx, col_name, sku_header_format)

    # Month headers
    month_start_col = len(sku_detail_cols)
    for month_idx, month in enumerate(all_months):
        month_str = month.strftime('%b %Y')
        is_forecast = month > current_date
        fmt = forecast_format if is_forecast else header_format
        worksheet.write(0, month_start_col + month_idx, month_str, header_format)

    # Set column widths
    worksheet.set_column(0, 0, 15)  # Category
    worksheet.set_column(1, 1, 15)  # Sub-category
    worksheet.set_column(2, 2, 15)  # Collection
    worksheet.set_column(3, 3, 25)  # SKU
    worksheet.set_column(4, 4, 40)  # Description
    worksheet.set_column(5, 5, 12)  # Color
    worksheet.set_column(6, 6, 12)  # Size
    worksheet.set_column(month_start_col, month_start_col + len(all_months), 10)  # Month columns

    current_row = 1

    # Process each SKU
    for _, sku_row in category_skus.iterrows():
        sku = sku_row['SKU']

        # Get SKU details
        sku_details = [
            sku_row.get('CATEGORY', ''),
            sku_row.get('SUB_CATEGORY', ''),
            sku_row.get('COLLECTION', ''),
            sku,
            sku_row.get('SKU_DESCRIPTION', ''),
            sku_row.get('COLOR_NAME', ''),
            sku_row.get('SIZE', '')
        ]

        # Get current inventory data
        current_on_hand = sku_row.get('AVAILABLE_ON_HAND_QTY', 0)
        committed_qty = sku_row.get('QTY_COMMITTED', 0)
        backorder_qty = sku_row.get('QTY_BACKORDERED', 0)

        # Get historical sales
        if sku in sales_pivot.index:
            sku_sales = sales_pivot.loc[sku]
        else:
            sku_sales = pd.Series(0, index=all_months)

        # Calculate forecast
        forecast_value = calculate_forecast(sku_sales)

        # Build data for each row type
        sales_demand_data = []
        committed_data = []
        backorder_data = []
        eom_inventory_data = []
        projected_eom_data = []
        receipts_data = []

        running_inventory = current_on_hand

        for month in all_months:
            is_forecast_month = month > current_date
            is_current_month = month.year == current_date.year and month.month == current_date.month

            # Sales Demand
            if is_forecast_month:
                demand = round(forecast_value)
            else:
                demand = sku_sales.get(month, 0) if month in sku_sales.index else 0
            sales_demand_data.append(demand)

            # Committed (only for current/past months)
            if is_current_month:
                committed_data.append(committed_qty)
            elif not is_forecast_month:
                committed_data.append(0)
            else:
                committed_data.append(0)

            # Backorder (only for current/past months)
            if is_current_month:
                backorder_data.append(backorder_qty)
            elif not is_forecast_month:
                backorder_data.append(0)
            else:
                backorder_data.append(0)

            # Receipts / On-Order
            receipt = on_order_by_sku_month.get((sku, month), 0)
            receipts_data.append(receipt)

            # EOM Inventory (actual for past months)
            # Projected EOM Inventory (for current and future)
            if is_forecast_month or is_current_month:
                # Calculate projected: previous inventory + receipts - demand
                running_inventory = running_inventory + receipt - demand
                eom_inventory_data.append(None)  # No actual EOM for future
                projected_eom_data.append(running_inventory)
            else:
                # For past months, we'd need actual EOM data - using calculated for now
                eom_inventory_data.append(None)
                projected_eom_data.append(None)

        # Recalculate projected EOM from beginning for accuracy
        projected_eom_data = []
        running_inv = current_on_hand
        for month_idx, month in enumerate(all_months):
            is_current_or_future = month >= current_date
            demand = sales_demand_data[month_idx]
            receipt = receipts_data[month_idx]

            if is_current_or_future:
                running_inv = running_inv + receipt - demand
                projected_eom_data.append(running_inv)
            else:
                projected_eom_data.append(None)

        # Write SKU details (first row of this SKU block)
        for col_idx, detail in enumerate(sku_details):
            worksheet.write(current_row, col_idx, detail if pd.notna(detail) else '', text_format)

        # Write row label and data for Sales Demand
        worksheet.write(current_row, len(sku_detail_cols) - 1, 'Sales Demand', row_label_format)
        for month_idx, value in enumerate(sales_demand_data):
            is_forecast_month = all_months[month_idx] > current_date
            fmt = forecast_format if is_forecast_month else number_format
            worksheet.write(current_row, month_start_col + month_idx, value, fmt)
        current_row += 1

        # Committed Qty row
        for col_idx in range(len(sku_detail_cols) - 1):
            worksheet.write(current_row, col_idx, '', text_format)
        worksheet.write(current_row, len(sku_detail_cols) - 1, 'Committed Qty', row_label_format)
        for month_idx, value in enumerate(committed_data):
            worksheet.write(current_row, month_start_col + month_idx, value, number_format)
        current_row += 1

        # Backorder Qty row
        for col_idx in range(len(sku_detail_cols) - 1):
            worksheet.write(current_row, col_idx, '', text_format)
        worksheet.write(current_row, len(sku_detail_cols) - 1, 'Backorder Qty', row_label_format)
        for month_idx, value in enumerate(backorder_data):
            worksheet.write(current_row, month_start_col + month_idx, value, number_format)
        current_row += 1

        # EOM Inventory row (actual - only for historical)
        for col_idx in range(len(sku_detail_cols) - 1):
            worksheet.write(current_row, col_idx, '', text_format)
        worksheet.write(current_row, len(sku_detail_cols) - 1, 'EOM Inventory', row_label_format)
        for month_idx, value in enumerate(eom_inventory_data):
            if value is not None:
                fmt = negative_format if value < 0 else number_format
                worksheet.write(current_row, month_start_col + month_idx, value, fmt)
            else:
                worksheet.write(current_row, month_start_col + month_idx, '', number_format)
        current_row += 1

        # Projected EOM Inventory row
        for col_idx in range(len(sku_detail_cols) - 1):
            worksheet.write(current_row, col_idx, '', text_format)
        worksheet.write(current_row, len(sku_detail_cols) - 1, 'Projected EOM Inv', row_label_format)
        for month_idx, value in enumerate(projected_eom_data):
            if value is not None:
                is_forecast_month = all_months[month_idx] > current_date
                if value < 0:
                    fmt = forecast_negative_format if is_forecast_month else negative_format
                else:
                    fmt = forecast_format if is_forecast_month else number_format
                worksheet.write(current_row, month_start_col + month_idx, value, fmt)
            else:
                worksheet.write(current_row, month_start_col + month_idx, '', number_format)
        current_row += 1

        # Receipts row
        for col_idx in range(len(sku_detail_cols) - 1):
            worksheet.write(current_row, col_idx, '', text_format)
        worksheet.write(current_row, len(sku_detail_cols) - 1, 'Receipts (On-Order)', row_label_format)
        for month_idx, value in enumerate(receipts_data):
            is_forecast_month = all_months[month_idx] > current_date
            fmt = forecast_format if is_forecast_month else number_format
            worksheet.write(current_row, month_start_col + month_idx, value, fmt)
        current_row += 1

        # Add empty row between SKUs for readability
        current_row += 1

    print(f"  - Written {len(category_skus)} SKUs to sheet '{sheet_name}'")

# Close the workbook
writer.close()
print(f"\n{'='*50}")
print(f"Model completed! Output saved to:")
print(f"{output_file}")
print(f"{'='*50}")
