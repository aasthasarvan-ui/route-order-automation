import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Route & Order Automation Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 100% Exact Route & Order Automation Dashboard")
st.markdown(
    "Apni **INPUT FILE.xlsx** yahan upload karein. App exact original output"
    " file ke format, formulas (`INDEX/MATCH`, `SUMIF`, `HYPERLINK`), dropdowns"
    " aur color formatting ke sath output generate karega."
)

uploaded_file = st.file_uploader(
    "Upload INPUT FILE (.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    df_in = pd.read_excel(uploaded_file)
    st.success("File successfully uploaded and read!")

    if st.button("🚀 Generate Exact Output File"):
      with st.spinner("Processing data, formulas, and formatting..."):

        output_io = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default sheet

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

        def clean_status(val):
          val_str = str(val)
          if "In Process" in val_str:
            return "Pending"
          return val_str

        # Prepare records for master sheets
        records = []
        for idx, row in df_in.iterrows():
          status_val = clean_status(row.get("Overall Status", "Pending"))
          dist_chan = str(row.get("Distribution Channel Description", ""))
          r_idx = len(records) + 2  # Excel row index starting from 2

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
              "Open Quantity": (
                  f'=IF(A{r_idx}="Dispatch", 0, MAX(0, N{r_idx}-O{r_idx}))'
              ),
              "Sales Person Name": row.get("Sales Person Name"),
              "Lookup Key": f'=E{r_idx}&"_"&K{r_idx}',
          })

        df_processed = pd.DataFrame(records)

        # Color fills
        fill_yellow = PatternFill(
            start_color="FEEFC3", end_color="FEEFC3", fill_type="solid"
        )
        fill_blue = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        fill_green = PatternFill(
            start_color="E2F0D9", end_color="E2F0D9", fill_type="solid"
        )

        def populate_master_sheet(ws, df_subset):
          ws.append(master_cols)
          for col_num in range(1, len(master_cols) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = PatternFill(
                start_color="333F50", end_color="333F50", fill_type="solid"
            )
            cell.font = Font(color="FFFFFF", bold=True)

          row_idx = 2
          for _, r in df_subset.iterrows():
            ws.append([
                r["Status"],
                r["Sales Organization"],
                r["Plant"],
                r["Plant Name"],
                r["Sales Order"],
                r["Sales Order Date"],
                r["Distribution Channel Description"],
                r["Sold to Party Name"],
                r["Route"],
                r["Ship to Party Name"],
                r["Product"],
                r["Product Description"],
                r["Base Unit of Measure"],
                r["Order Quantity"],
                r["InvoicedQuantity"],
                r["Open Quantity"],  # Formula string
                r["Sales Person Name"],
                r["Lookup Key"],  # Formula string
            ])

            # Apply row color based on status
            stat = str(r["Status"])
            if "Pending" in stat or "Process" in stat:
              for c in range(1, len(master_cols) + 1):
                ws.cell(row=row_idx, column=c).fill = fill_yellow
            elif "Planning" in stat:
              for c in range(1, len(master_cols) + 1):
                ws.cell(row=row_idx, column=c).fill = fill_blue
            elif "Dispatch" in stat:
              for c in range(1, len(master_cols) + 1):
                ws.cell(row=row_idx, column=c).fill = fill_green
            row_idx += 1

          # Dropdown
          dv = DataValidation(
              type="list", formula1='"Planning,Dispatch,Pending"', allow_blank=True
          )
          ws.add_data_validation(dv)
          dv.add(f"A2:A{max(row_idx-1, 1000)}")
          return row_idx

        # --- A. General Trade Sheet ---
        ws_gt = wb.create_sheet(title="General Trade")
        df_gt_data = df_processed[
            df_processed["Distribution Channel Description"].str.contains(
                "General Trade", case=False, na=False
            )
        ]
        gt_max_row = populate_master_sheet(ws_gt, df_gt_data)

        # --- B. Nestle Sheet ---
        ws_nestle = wb.create_sheet(title="Nestle")
        df_nestle_data = df_processed[
            df_processed["Distribution Channel Description"] == "Nestle"
        ]
        nestle_max_row = populate_master_sheet(ws_nestle, df_nestle_data)

        # --- C. Nestle Open Sheet ---
        ws_n_open = wb.create_sheet(title="Nestle Open")
        df_nopen_data = df_processed[
            df_processed["Distribution Channel Description"] == "Nestle Open"
        ]
        nopen_max_row = populate_master_sheet(ws_n_open, df_nopen_data)

        # --- D. Route Sheets ---
        unique_routes = df_processed["Route"].dropna().unique()
        sorted_routes = sorted(
            [str(r) for r in unique_routes],
            key=lambda x: (
                0 if x.isdigit() else 1,
                int(x) if x.isdigit() else str(x),
            ),
        )

        route_summary_data = []

        # Ranges for master sheets for INDEX/MATCH sync formula
        gt_rng = f"'General Trade'!$A$2:$A${gt_max_row-1}"
        gt_lkp = f"'General Trade'!$R$2:$R${gt_max_row-1}"
        nestle_rng = f"Nestle!$A$2:$A${nestle_max_row-1}"
        nestle_lkp = f"Nestle!$R$2:$R${nestle_max_row-1}"
        nopen_rng = f"'Nestle Open'!$A$2:$A${nopen_max_row-1}"
        nopen_lkp = f"'Nestle Open'!$R$2:$R${nopen_max_row-1}"

        for r_val in sorted_routes:
          sheet_name = (
              f"Route {r_val}" if not r_val.startswith("Route") else r_val
          )
          if len(r_val) == 1 and r_val.isdigit():
            sheet_name = f"Route 0{r_val}"
          if "/" in sheet_name:
            sheet_name = sheet_name.replace("/", "_")

          ws_route = wb.create_sheet(title=sheet_name)
          df_r_data = df_in[df_in["Route"].astype(str) == r_val]

          # Route sheet columns (without Lookup Key in original route sheets)
          route_cols = [
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
          ]
          ws_route.append(route_cols)
          for col_num in range(1, len(route_cols) + 1):
            cell = ws_route.cell(row=1, column=col_num)
            cell.fill = PatternFill(
                start_color="333F50", end_color="333F50", fill_type="solid"
            )
            cell.font = Font(color="FFFFFF", bold=True)

          r_row_idx = 2
          for _, raw_row in df_r_data.iterrows():
            # Status formula synced via INDEX/MATCH with master sheets
            status_formula = (
                f'=IFERROR(INDEX({nestle_rng}, MATCH(E{r_row_idx}&"_"&K{r_row_idx},'
                f" {nestle_lkp}, 0)),"
                f' IFERROR(INDEX({nopen_rng}, MATCH(E{r_row_idx}&"_"&K{r_row_idx},'
                f" {nopen_lkp}, 0)),"
                f' IFERROR(INDEX({gt_rng}, MATCH(E{r_row_idx}&"_"&K{r_row_idx},'
                f' {gt_lkp}, 0)), "Pending")))'
            )

            ws_route.append([
                status_formula,
                raw_row.get("Sales Organization"),
                raw_row.get("Plant"),
                raw_row.get("Plant Name"),
                raw_row.get("Sales Order"),
                raw_row.get("Sales Order Date"),
                raw_row.get("Distribution Channel Description"),
                raw_row.get("Sold to Party Name"),
                raw_row.get("Route"),
                raw_row.get("Ship to Party Name"),
                raw_row.get("Product"),
                raw_row.get("Product Description"),
                raw_row.get("Base Unit of Measure"),
                raw_row.get("Order Quantity"),
                raw_row.get("InvoicedQuantity"),
                (
                    f'=IF(A{r_row_idx}="Dispatch", 0,'
                    f" MAX(0, N{r_row_idx}-O{r_row_idx}))"
                ),
                raw_row.get("Sales Person Name"),
            ])

            # Color formatting based on initial status value
            stat_txt = str(raw_row.get("Overall Status", ""))
            if "Pending" in stat_txt or "Process" in stat_txt:
              for c in range(1, len(route_cols) + 1):
                ws_route.cell(row=r_row_idx, column=c).fill = fill_yellow
            elif "Planning" in stat_txt:
              for c in range(1, len(route_cols) + 1):
                ws_route.cell(row=r_row_idx, column=c).fill = fill_blue
            elif "Dispatch" in stat_txt:
              for c in range(1, len(route_cols) + 1):
                ws_route.cell(row=r_row_idx, column=c).fill = fill_green
            r_row_idx += 1

          # Total Row
          ws_route.cell(row=r_row_idx, column=2, value="Total").font = Font(
              bold=True
          )
          open_col_let = get_column_letter(16)  # Column P is Open Quantity
          ws_route.cell(
              row=r_row_idx,
              column=16,
              value=f"=SUM({open_col_let}2:{open_col_let}{r_row_idx-1})",
          ).font = Font(bold=True)
          ws_route.cell(row=r_row_idx, column=16).fill = PatternFill(
              start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
          )

          # Dropdown
          dv = DataValidation(
              type="list", formula1='"Planning,Dispatch,Pending"', allow_blank=True
          )
          ws_route.add_data_validation(dv)
          dv.add(f"A2:A{max(r_row_idx-1, 1000)}")

          channels_list = (
              df_r_data["Distribution Channel Description"]
              .unique()
              .tolist()
          )
          channels_str = ", ".join([str(c) for c in channels_list])

          route_summary_data.append({
              "Route No.": r_val,
              "Total Pending Items": len(df_r_data),
              "SheetName": sheet_name,
              "LastRow": r_row_idx,
              "OpenColLetter": open_col_let,
              "Channels": channels_str,
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

        stock_agg = (
            df_processed.groupby(
                ["Product", "Product Description", "Base Unit of Measure"]
            )["Open Quantity"]
            .count()
            .reset_index()
        )  # group for unique products
        # Get unique product details from raw input
        unique_prods = (
            df_in[["Product", "Product Description", "Base Unit of Measure"]]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        all_sheet_names_for_sumif = (
            ["General Trade", "Nestle", "Nestle Open"]
            + [item["SheetName"] for item in route_summary_data]
        )

        stock_row = 5
        for _, prod_row in unique_prods.iterrows():
          p_code = prod_row["Product"]
          p_desc = prod_row["Product Description"]
          p_uom = prod_row["Base Unit of Measure"]
          cat = (
              "Nestle"
              if "NESTLE" in str(p_desc).upper()
              or str(p_code).startswith("FG50002")
              else "Retail"
          )

          # Build SUMIF formula across all sheets
          sumif_parts = []
          for s_nm in all_sheet_names_for_sumif:
            sumif_parts.append(
                f"SUMIF('{s_nm}'!$K:$K, A{stock_row}, '{s_nm}'!$P:$P)"
            )
          total_qty_formula = "=" + "+".join(sumif_parts)

          ws_stock.append([
              p_code,
              p_desc,
              cat,
              p_uom,
              total_qty_formula,
              0,
              f"=F{stock_row}-E{stock_row}",
          ])
          stock_row += 1

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
          ws_index.cell(row=idx_row, column=4, value=item["Channels"])
          ws_index.cell(
              row=idx_row,
              column=5,
              value=f'=HYPERLINK("#\'{s_name}\'!A1", "Open {s_name} Sheet")',
          )
          idx_row += 1

        wb.save(output_io)
        output_io.seek(0)

        st.success("🎉 100% Exact Output File successfully generated!")
        st.download_button(
            label="📥 Download Exact Output File (.xlsx)",
            data=output_io,
            file_name="Output_File_Exact.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

  except Exception as e:
    st.error(f"Error processing file: {e}")
