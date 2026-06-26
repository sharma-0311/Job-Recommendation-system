import os
import openpyxl
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

class BatchBuilder:
    """
    Constructs a lightweight runtime_batch.xlsx Excel file containing only
    the user's selected records. Injects hidden in-memory tracking fields.
    """
    @staticmethod
    def generate_batch_id() -> str:
        """Generates a hidden runtime Batch ID based on current timestamp."""
        now = datetime.now()
        return f"BATCH_{now.strftime('%Y%m%d_%H%M%S')}"

    @staticmethod
    def build_runtime_excel(
        master_excel_path: Path,
        selected_reg_ids: List[str],
        batch_id: str,
        output_excel_path: Path,
        operator: str = "Operator"
    ) -> List[Dict[str, Any]]:
        """
        Reads the validator's master workbook, filters by selected GST_Reg_ID values,
        adds hidden runtime metadata, and writes out a specialized runtime_batch.xlsx file.
        Returns a list of parsed record dictionaries representing the selected rows.
        """
        if not master_excel_path.exists():
            raise FileNotFoundError(f"Master validation workbook not found at: {master_excel_path}")

        # Load master workbook using openpyxl
        wb_master = openpyxl.load_workbook(master_excel_path, data_only=True)
        if "GST_Registrations" not in wb_master.sheetnames:
            wb_master.close()
            raise ValueError("Sheet 'GST_Registrations' not found in master workbook.")

        ws_master = wb_master["GST_Registrations"]
        rows = list(ws_master.iter_rows(values_only=False))
        if len(rows) < 2:
            wb_master.close()
            raise ValueError("No data found in master 'GST_Registrations' sheet.")

        # Extract headers
        headers = [cell.value for cell in rows[0]]
        headers = [str(h).strip() if h is not None else f"Column_{i}" for i, h in enumerate(headers)]

        # Find column indices for Reg ID and PPOB address
        reg_id_idx = -1
        pob_addr_idx = -1
        doc_path_idx = -1
        legal_name_idx = -1
        state_idx = -1
        possession_idx = -1
        
        for idx, h in enumerate(headers):
            hl = h.lower()
            if hl == "gst_reg_id":
                reg_id_idx = idx
            elif "address" in hl and "ppob" in hl:
                pob_addr_idx = idx
            elif "total documents" in hl:
                doc_path_idx = idx
            elif "legal name" in hl:
                legal_name_idx = idx
            elif "state" in hl:
                state_idx = idx
            elif "ownership proof type" in hl:
                possession_idx = idx

        if reg_id_idx == -1:
            wb_master.close()
            raise ValueError("Could not find required 'GST_Reg_ID' column in workbook.")

        # Filter rows
        selected_set = set(selected_reg_ids)
        selected_rows = []
        parsed_records = []

        # Create new runtime workbook
        wb_runtime = openpyxl.Workbook()
        ws_runtime = wb_runtime.active
        ws_runtime.title = "GST_Registrations"

        # Copy original headers
        for col_idx, h in enumerate(headers, 1):
            ws_runtime.cell(row=1, column=col_idx, value=h)

        # Append new metadata columns to headers list and sheet headers
        meta_headers = [
            "Batch_ID",
            "Batch_Size",
            "Run_ID",
            "Timestamp",
            "Operator",
            "Internal_Only_Flag"
        ]
        meta_start_col = len(headers) + 1
        for col_offset, mh in enumerate(meta_headers):
            ws_runtime.cell(row=1, column=meta_start_col + col_offset, value=mh)

        # Populate selected data
        run_id = f"RUN_{datetime.now().strftime('%Y%m%d')}_{os.getpid()}"
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        batch_size = len(selected_set)
        
        row_counter = 2
        for r_idx in range(1, len(rows)):
            row_cells = rows[r_idx]
            reg_id_val = str(row_cells[reg_id_idx].value or "").strip()
            
            if reg_id_val in selected_set:
                selected_rows.append(row_cells)
                
                # Write standard fields, preserving cell values & hyperlinks
                record_dict = {}
                for col_idx, cell in enumerate(row_cells, 1):
                    new_cell = ws_runtime.cell(row=row_counter, column=col_idx, value=cell.value)
                    header = headers[col_idx - 1]
                    record_dict[header] = cell.value
                    
                    if cell.hyperlink:
                        new_cell.hyperlink = cell.hyperlink
                        record_dict[f"{header}_hyperlink"] = cell.hyperlink.target
                
                # Map standard keys required by amend_req.py
                # Note: GSTIN will map to GST_Reg_ID in the grid, but if user enters/edits it, we preserve it.
                # If there's an explicit 'GSTIN' column already, we map it, else fallback to GST_Reg_ID
                gstin_val = record_dict.get("GSTIN") or reg_id_val
                record_dict["GSTIN"] = gstin_val
                
                if legal_name_idx != -1:
                    record_dict["TradeName"] = record_dict.get(headers[legal_name_idx])
                if state_idx != -1:
                    record_dict["State"] = record_dict.get(headers[state_idx])
                if possession_idx != -1:
                    record_dict["PossessionNature"] = record_dict.get(headers[possession_idx])
                
                # Map documents
                if doc_path_idx != -1:
                    doc_h = headers[doc_path_idx]
                    record_dict["DocumentPath"] = record_dict.get(doc_h)
                    if f"{doc_h}_hyperlink" in record_dict:
                        record_dict["DocumentPath_hyperlink"] = record_dict[f"{doc_h}_hyperlink"]

                # Write metadata fields to the runtime row
                meta_values = [
                    batch_id,
                    batch_size,
                    run_id,
                    timestamp_str,
                    operator,
                    "TRUE"
                ]
                for col_offset, val in enumerate(meta_values):
                    ws_runtime.cell(row=row_counter, column=meta_start_col + col_offset, value=val)

                # Store metadata in record dict
                for mh, mv in zip(meta_headers, meta_values):
                    record_dict[mh] = mv

                parsed_records.append(record_dict)
                row_counter += 1

        wb_runtime.save(output_excel_path)
        wb_runtime.close()
        wb_master.close()

        return parsed_records
