import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Route & Order Automation Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Route-Wise Pending Orders & Dashboard Automation")
st.markdown(
    "Apni **INPUT FILE.xlsx** yahan upload karein. App automatically saari Master sheets, Route sheets, Stock vs Order report aur Route Index dashboard generate karke download link de dega."
)

uploaded_file = st.file_uploader(
    "Upload INPUT FILE (.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    # Read input file
    df_in = pd.read_excel(uploaded_file)
    st.success("File successfully uploaded and read!")

    if st.button("🚀 Generate Output File"):
      with st.spinner("Generating sheets, dashboard, and formulas..."):

        # In-memory output workbook using openpyxl
        output_io = io.BytesIO()
        wb = openpyxl.Workbook()
        # Remove default sheet
        wb.remove(wb.active)

        # 1. Setup Master Dataframes from input
        # Standard columns for category sheets
        master_cols = [
            "Status",
            "Sales Organization",
            "Plant",
            "Plant Name",
            "Sales Order",
            "Sales Order Date",
            "Distribution Channel Description",
            "Sold to Party Name",
            "Route",
            "Ship to Party Name",
            "Product",
            "Product Description",
            "Base Unit of Measure",
            "Order Quantity",
            "InvoicedQuantity",
            "Open Quantity",
            "Sales Person Name",
            "Lookup Key",
        ]

        # Helper to map status
        def clean_status(val):
          val_str = str(val)
          if "In Process" in val_str:
            return "Pending"
          return val_str

        # Prepare processed raw records
        records = []
        for idx, row in df_in.iterrows():
          status_val = clean_status(row.get("Overall Status", "Pending"))
          dist_chan = str(row.get("Distribution Channel Description", ""))
          records.append({
              "Status": status_val,
              "Sales Organization": row.get("Sales Organization"),
              "Plant": row.get("Plant"),
              "Plant Name": row.get("Plant Name"),
              "Sales Order": row.get("Sales Order"),
              "Sales Order Date": row.get("Sales Order Date"),
              "Distribution Channel Description": dist_chan,
              "Sold to Party Name": row.get("Sold to Party Name"),
              "Route": row.get("Route"),
              "Ship to Party Name": row.get("Ship to Party Name"),
              "Product": row.get("Product"),
              "Product Description": row.get("Product Description"),
              "Base Unit of Measure": row.get("Base Unit of Measure"),
              "Order Quantity": row.get("Order Quantity"),
              "InvoicedQuantity": row.get("InvoicedQuantity"),
              "Open Quantity": row.get("Open Quantity"),
              "Sales Person Name": row.get("Sales Person Name"),
              "Lookup Key": f"{row.get('Sales Order')}_{row.get('Product')}",
          })

        df_processed = pd.DataFrame(records)

        # --- A. General Trade Sheet ---
        ws_gt = wb.create_sheet(title="General Trade")
        df_gt_data = df_processed[
            df_processed["Distribution Channel Description"].str.contains(
                "General Trade", case=False, na=False
            )
        ]
        ws_gt.append(master_cols)
        for _, r in df_gt_data.iterrows():
          ws_gt.append([r[c] for c in master_cols])

        # --- B. Nestle Sheet ---
        ws_nestle = wb.create_sheet(title="Nestle")
        df_nestle_data = df_processed[
            df_processed["Distribution Channel Description"] == "Nestle"
        ]
        ws_nestle.append(master_cols)
        for _, r in df_nestle_data.iterrows():
          ws_nestle.append([r[c] for c in master_cols])

        # --- C. Nestle Open Sheet ---
        ws_n_open = wb.create_sheet(title="Nestle Open")
        df_nopen_data = df_processed[
            df_processed["Distribution Channel Description"] == "Nestle Open"
        ]
        ws_n_open.append(master_cols)
        for _, r in df_nopen_data.iterrows():
          ws_n_open.append([r[c] for c in master_cols])

        # --- D. Route Sheets ---
        unique_routes = df_processed["Route"].dropna().unique()
        # Sort routes nicely if possible
        sorted_routes = sorted(
            [str(r) for r in unique_routes],
            key=lambda x: (
                0 if x.isdigit() else 1,
                int(x) if x.isdigit() else str(x),
            ),
        )

        route_summary_data = []

        for r_val in sorted_routes:
          # Format sheet name
          sheet_name = (
              f"Route {r_val}" if not r_val.startswith("Route") else r_val
          )
          if len(r_val) == 1 and r_val.isdigit():
            sheet_name = f"Route 0{r_val}"
          if "/" in sheet_name:
            sheet_name = sheet_name.replace("/", "_")

          ws_route = wb.create_sheet(title=sheet_name)

          # Original input columns for route sheets
          input_cols = list(df_in.columns)
          ws_route.append(input_cols)

          df_r_data = df_in[df_in["Route"].astype(str) == r_val]

          fill_yellow = PatternFill(
              start_color="FEEFC3", end_color="FEEFC3", fill_type="solid"
          )
          fill_blue = PatternFill(
              start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
          )
          fill_green = PatternFill(
              start_color="E2F0D9", end_color="E2F0D9", fill_type="solid"
          )

          row_idx = 2
          for _, raw_row in df_r_data.iterrows():
            row_vals = [raw_row[col] for col in input_cols]
            ws_route.append(row_vals)

            # Apply conditional color on status
            status_txt = str(raw_row.get("Overall Status", ""))
            if "Pending" in status_txt or "Process" in status_txt:
              for col_num in range(1, len(input_cols) + 1):
                ws_route.cell(row=row_idx, column=col_num).fill = fill_yellow
            elif "Planning" in status_txt:
              for col_num in range(1, len(input_cols) + 1):
                ws_route.cell(row=row_idx, column=col_num).fill = fill_blue
            elif "Dispatch" in status_txt:
              for col_num in range(1, len(input_cols) + 1):
                ws_route.cell(row=row_idx, column=col_num).fill = fill_green
            row_idx += 1

          # Total Row
          ws_route.cell(row=row_idx, column=1, value="Total").font = Font(
              bold=True
          )
          open_qty_col_idx = (
              input_cols.index("Open Quantity") + 1
              if "Open Quantity" in input_cols
              else len(input_cols)
          )
          col_letter = get_column_letter(open_qty_col_idx)
          ws_route.cell(
              row=row_idx,
              column=open_qty_col_idx,
              value=f"=SUM({col_letter}2:{col_letter}{row_idx-1})",
          ).font = Font(bold=True)
          ws_route.cell(row=row_idx, column=open_qty_col_idx).fill = (
              PatternFill(
                  start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
              )
          )

          total_pending_items = len(df_r_data)
          # We'll compute total open qty via formula or python value for index
          total_open_qty_val = df_r_data["Open Quantity"].sum()

          # Channels for this route
          channels_list = (
              df_r_data["Distribution Channel Description"]
              .unique()
              .tolist()
          )
          channels_str = ", ".join([str(c) for c in channels_list])

          route_summary_data.append({
              "Route No.": r_val,
              "Total Pending Items": total_pending_items,
              "Total Open Quantity": total_open_qty_val,
              "Distribution Channels": channels_str,
              "SheetName": sheet_name,
              "LastRow": row_idx,
              "OpenColLetter": col_letter,
          })

        # --- E. Stock vs Order Sheet ---
        ws_stock = wb.create_sheet(title="Stock vs Order")
        ws_stock.append(["STOCK VS ORDER & MATERIAL CATEGORY REPORT"])
        ws_stock.append([
            "Note: Available Stock Qty column me apna current stock enter"
            " karein. Shortage / Surplus automatically calculate hoga."
        ])
        ws_stock.append([])
        ws_stock.append([
            "Product Code",
            "Product Description",
            "Material Category",
            "Base UOM",
            "Total Pending Order Qty",
            "Available Stock Qty",
            "Shortage / Surplus",
        ])

        # Group by Product
        stock_agg = (
            df_processed.groupby(
                ["Product", "Product Description", "Base Unit of Measure"]
            )["Open Quantity"]
            .sum()
            .reset_index()
        )
        stock_agg = stock_agg.sort_values(by="Open Quantity", ascending=False)

        stock_row_start = 5
        for idx, row in stock_agg.iterrows():
          p_code = row["Product"]
          p_desc = row["Product Description"]
          p_uom = row["Base Unit of Measure"]
          p_qty = row["Open Quantity"]
          cat = (
              "Nestle"
              if "NESTLE" in str(p_desc).upper() or p_code.startswith("FG50002")
              else "Retail"
          )

          ws_stock.append([
              p_code,
              p_desc,
              cat,
              p_uom,
              p_qty,
              0,
              f"=F{stock_row_start}-E{stock_row_start}",
          ])
          stock_row_start += 1

        # --- F. Route Index Sheet ---
        ws_index = wb.create_sheet(title="Route Index", index=0)
        ws_index.append(["ROUTE-WISE PENDING ORDERS INDEX & DASHBOARD"])
        ws_index.append([
            "Neeche diye gaye route links par click karke aap seedhe us route"
            " ki sheet par ja sakte hain. Sabhi status changes, dropdowns aur"
            " formatting fully synchronized hain."
        ])
        ws_index.append([])
        ws_index.append([
            "Route No.",
            "Total Pending Items",
            "Total Open Quantity",
            "Distribution Channels",
            "Route Sheet Navigation Link",
        ])

        # Style Index Header
        header_fill = PatternFill(
            start_color="1F4E79", end_color="1F4E79", fill_type="solid"
        )
        for col_num in range(1, 6):
          cell = ws_index.cell(row=4, column=col_num)
          cell.fill = header_fill
          cell.font = Font(color="FFFFFF", bold=True)

        idx_row = 5
        for item in route_summary_data:
          s_name = item["SheetName"]
          l_row = item["LastRow"]
          o_col = item["OpenColLetter"]

          ws_index.cell(row=idx_row, column=1, value=item["Route No."])
          ws_index.cell(
              row=idx_row,
              column=2,
              value=f"=COUNTA('{s_name}'!A2:A{l_row-1})",
          )
          ws_index.cell(
              row=idx_row, column=3, value=f"='{s_name}'!{o_col}{l_row}"
          )
          ws_index.cell(row=idx_row, column=4, value=item["Distribution Channels"])
          ws_index.cell(
              row=idx_row,
              column=5,
              value=f'=HYPERLINK("#\'{s_name}\'!A1", "Open {s_name} Sheet")',
          )
          idx_row += 1

        # Format master sheets headers
        for ws in [ws_gt, ws_nestle, ws_n_open]:
          for col_num in range(1, len(master_cols) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = PatternFill(
                start_color="333F50", end_color="333F50", fill_type="solid"
            )
            cell.font = Font(color="FFFFFF", bold=True)

        # Save workbook to bytes
        wb.save(output_io)
        output_io.seek(0)

        st.success(
            "🎉 Output File successfully generated with all Route sheets,"
            " Dashboard, and Formulas!"
        )

        st.download_button(
            label="📥 Download Output File (.xlsx)",
            data=output_io,
            file_name="Output_File_Generated.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

  except Exception as e:
    st.error(f"Error processing file: {e}")
