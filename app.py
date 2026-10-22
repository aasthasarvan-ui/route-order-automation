import io
import openpyxl
from openpyxl.formatting.rule import FormulaRule
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

st.title("📊 100% Fully Updated Route & Order Automation Dashboard")
st.markdown(
    "Apni **INPUT FILE.xlsx** yahan upload karein. App exact formatting,"
    " formulas (`INDEX/MATCH`, `SUMIF`, `HYPERLINK`), dropdowns aur proper row"
    " conditional formatting ke sath output generate karega."
)

uploaded_file = st.file_uploader(
    "Upload INPUT FILE (.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    df_in = pd.read_excel(uploaded_file)
    st.success("File successfully uploaded and read!")

    if st.button("🚀 Generate Fully Updated Output File"):
      with st.spinner(
          "Processing data, formulas, conditional formatting, and"
          " layouts..."
      ):

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

        records = []
        for idx, row in df_in.iterrows():
          status_val = clean_status(row.get("Overall Status", "Pending"))
          dist_chan = str(row.get("Distribution Channel Description", ""))
          r_idx = len(records) + 2

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

        # Color Fills Definition
        fill_yellow = PatternFill(
            start_color="FEEFC3", end_color="FEEFC3", fill_type="solid"
        )
        fill_blue = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        fill_green = PatternFill(
            start_color="E2F0D9", end_color="E2F0D9", fill_type="solid"
        )

        def apply_conditional_formatting(ws, max_row, max_col_letter="Q"):
          # Formula rules referencing $A2 (without locking column A so entire row highlights)
          rule_planning = FormulaRule(
              formula=['$A2="Planning"'], stopIfTrue=True, fill=fill_blue
          )
          rule_dispatch = FormulaRule(
              formula=['$A2="Dispatch"'], stopIfTrue=True, fill=fill_green
          )
          rule_pending = FormulaRule(
              formula=['$A2="Pending"'], stopIfTrue=True, fill=fill_yellow
          )

          range_str = f"A2:{max_col_letter}{max_row}"
          ws.conditional_formatting.add(range_str, rule_planning)
          ws.conditional_formatting.add(range_str, rule_dispatch)
          ws.conditional_formatting.add(range_str, rule_pending)

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
                r["Open Quantity"],
                r["Sales Person Name"],
                r["Lookup Key"],
            ])
            row_idx += 1

          max_r = max(row_idx - 1, 2)
          apply_conditional_formatting(ws, max_r, "R")

          # Data Validation Dropdown for Status column
          dv = DataValidation(
              type="list", formula1='"Planning,Dispatch,Pending"', allow_blank=True
          )
          ws.add_data_validation(dv)
          dv.add(f"A2:A{max(max_r, 1000)}")
          return max_r + 1

        # --- A. General Trade Sheet ---
        ws_gt = wb.create_sheet(title="General Trade")
        gt_max_row = populate_master_sheet(
            ws_gt,
            df_processed[
                df_processed["Distribution Channel Description"].str.contains(
                    "General Trade", case=False, na=False
                )
            ],
        )

        # --- B. Nestle Sheet ---
        ws_nestle = wb.create_sheet(title="Nestle")
        nestle_max_row = populate_master_sheet(
            ws_nestle,
            df_processed[
                df_processed["Distribution Channel Description"] == "Nestle"
            ],
        )

        # --- C. Nestle Open Sheet ---
        ws_n_open = wb.create_sheet(title="Nestle Open")
        nopen_max_row = populate_master_sheet(
            ws_n_open,
            df_processed[
                df_processed["Distribution Channel Description"]
                == "Nestle Open"
            ],
        )

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
            r_row_idx += 1

          max_r_route = r_row_idx - 1
          apply_conditional_formatting(ws_route, max_r_route, "Q")

          # Total Row
          total_row_idx = r_row_idx
          ws_route.cell(row=total_row_idx, column=2, value="Total").font = Font(
              bold=True
          )
          open_col_let = get_column_letter(16)
          ws_route.cell(
              row=total_row_idx,
              column=16,
              value=f"=SUM({open_col_let}2:{open_col_let}{max_r_route})",
          ).font = Font(bold=True)
          ws_route.cell(row=total_row_idx, column=16).fill = PatternFill(
              start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
          )

          dv = DataValidation(
              type="list", formula1='"Planning,Dispatch,Pending"', allow_blank=True
          )
          ws_route.add_data_validation(dv)
          dv.add(f"A2:A{max(max_r_route, 1000)}")

          channels_str = ", ".join(
              [
                  str(c)
                  for c in df_r_data["Distribution Channel Description"]
                  .unique()
                  .tolist()
              ]
          )
          route_summary_data.append({
              "Route No.": r_val,
              "Total Pending Items": len(df_r_data),
              "SheetName": sheet_name,
              "LastRow": total_row_idx,
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

        unique_prods = (
            df_in[["Product", "Product Description", "Base Unit of Measure"]]
            .drop_duplicates()
            .reset_index(drop=True)
        )
        all_sheets_sumif = (
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

          sumif_parts = [
              f"SUMIF('{s_nm}'!$K:$K, A{stock_row}, '{s_nm}'!$P:$P)"
              for s_nm in all_sheets_sumif
          ]
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

        st.success(
            "🎉 100% Fully Updated Output File successfully generated with"
            " Conditional Formatting!"
        )
        st.download_button(
            label="📥 Download Fully Updated Output File (.xlsx)",
            data=output_io,
            file_name="Output_File_Updated.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

  except Exception as e:
    st.error(f"Error processing file: {e}")
