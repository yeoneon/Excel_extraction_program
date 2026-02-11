import os
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, XDRPositiveSize2D
from datetime import datetime
import random
from logger import logger

class ExcelHandler:
    def __init__(self, source_path, form_path, signature_dir, api_handler):
        self.source_path = source_path
        self.form_path = form_path
        self.signature_dir = signature_dir
        self.api_handler = api_handler
        logger.info("ExcelHandler initialized")

    def _safe_write(self, ws, coordinate, value):
        """Writes value to a cell, handling merged cells by finding the top-left coordinate."""
        from openpyxl.cell.cell import MergedCell
        cell = ws[coordinate]
        if isinstance(cell, MergedCell):
            logger.debug(f"Cell {coordinate} is a MergedCell. Finding root...")
            for range_ in ws.merged_cells.ranges:
                if coordinate in range_:
                    root_coord = range_.coord.split(':')[0]
                    logger.debug(f"Found root for {coordinate}: {root_coord}")
                    ws[root_coord].value = value
                    return
        cell.value = value

    def _format_date(self, raw_date):
        if isinstance(raw_date, (datetime, pd.Timestamp)):
            return raw_date.strftime("%Y-%m-%d"), raw_date.strftime("%Y%m%d")
        d_str = str(raw_date).replace(".", "-").replace("/", "-")
        return d_str, d_str.replace("-", "")

    def _add_signature(self, ws):
        try:
            sig_files = [f for f in os.listdir(self.signature_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            if sig_files:
                img_path = os.path.join(self.signature_dir, random.choice(sig_files))
                img = XLImage(img_path)
                img.width, img.height = 200, 75
                
                # Precise positioning: 10pt right and 10pt down from E22 marker
                # Column E is index 4, Row 22 is index 21 (0-indexed)
                # 1pt = 12700 EMUs, 1px = 9525 EMUs
                marker = AnchorMarker(col=4, colOff=10 * 12700, row=21, rowOff=3 * 12700)
                
                # We must define the size (Extent) in EMUs for OneCellAnchor
                size = XDRPositiveSize2D(cx=img.width * 9525, cy=img.height * 9525)
                img.anchor = OneCellAnchor(_from=marker, ext=size)
                
                ws.add_image(img)
                logger.debug(f"Added signature with OneCellAnchor (10pt offset): {img_path}")
            else:
                logger.warning("No signature files found in directory.")
        except Exception as e:
            logger.error(f"Failed to add signature: {e}")

    def process(self):
        """Main processing loop for Excel rows."""
        try:
            today_str = datetime.now().strftime("%Y-%m-%d")
            output_dir = os.path.join(os.getcwd(), today_str)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                logger.info(f"Created output directory: {output_dir}")

            logger.info(f"Reading source file: {self.source_path}")
            df = pd.read_excel(self.source_path, header=5)
            logger.info(f"Source file loaded. Columns: {df.columns.tolist()}")
            logger.info(f"Total rows to process: {len(df)}")

            processed_count = 0
            for index, row in df.iterrows():
                try:
                    logger.info(f"Processing row {index + 1}...")
                    
                    # Extraction (Adjusted based on logs: B=1, C=2, D=3, E=4, G=6)
                    raw_date = row.iloc[1]
                    date_val, date_filename = self._format_date(raw_date)
                    
                    representative = str(row.iloc[2]).strip()
                    company_name = str(row.iloc[3]).strip()
                    address_ko = str(row.iloc[4]).strip()
                    weight = str(row.iloc[6]).strip()

                    # Console logs for debugging
                    logger.info(f"--- Row Details ---")
                    logger.info(f"Date (Col B? Index 2): {date_val}")
                    logger.info(f"Representative (Col C? Index 3): {representative}")
                    logger.info(f"Store Name (Col D? Index 4): {company_name}")
                    logger.info(f"Address (Col E? Index 5): {address_ko}")
                    logger.info(f"Weight (Col G? Index 7): {weight}")
                    logger.info(f"-------------------")

                    logger.debug(f"Row data: Company={company_name}, Date={date_val}, Weight={weight}")

                    # Enrichment
                    company_name_en = self.api_handler.get_romanized_text(company_name, is_company=True)
                    phone, zip_code, address_en_fetched = self.api_handler.get_enriched_data(address_ko, company_name)
                    
                    # Use fetched English address if available, otherwise Romanize
                    if address_en_fetched:
                        address_en = address_en_fetched
                        logger.info(f"Using actual English address from NCP: {address_en}")
                    else:
                        address_en = self.api_handler.get_romanized_text(address_ko)
                        logger.info(f"Falling back to Romanized address: {address_en}")
                    
                    logger.info(f"Enriched Data -> Phone: {phone}, Zip: {zip_code}")

                    # Loading Template
                    wb = openpyxl.load_workbook(self.form_path)
                    ws = wb["CORSIA"] if "CORSIA" in wb.sheetnames else wb.active
                    logger.debug(f"Template loaded. active sheet: {ws.title}")

                    # Filling via _safe_write (Handles Merged Cells)
                    self._safe_write(ws, 'C4', f"{company_name_en} / {company_name}")
                    self._safe_write(ws, 'C5', f"{address_en} / {address_ko}")
                    self._safe_write(ws, 'C7', zip_code)
                    self._safe_write(ws, 'C8', phone)
                    self._safe_write(ws, 'B12', f"수거일 : {date_val}")
                    self._safe_write(ws, 'C12', f"수거량 : {weight} kg")
                    self._safe_write(ws, 'B13', f"DATE : {date_val}")
                    self._safe_write(ws, 'C13', f"Quantity collected: {weight} kg")
                    self._safe_write(ws, 'A22', f"{company_name}/{date_filename}")
                    self._safe_write(ws, 'C22', representative)

                    # Signature
                    self._add_signature(ws)

                    # Save
                    template_name = os.path.splitext(os.path.basename(self.form_path))[0]
                    save_filename = f"{template_name}_{company_name}.xlsx"
                    save_path = os.path.join(output_dir, save_filename)
                    wb.save(save_path)
                    
                    logger.info(f"Saved: {save_path}")
                    processed_count += 1
                    
                    # Debugging: Stop after the first row as requested
                    logger.info("Debug mode: Stopping after the first row.")
                    if processed_count == 2:
                        break

                except Exception as row_error:
                    logger.error(f"Error in row {index + 1}: {row_error}", exc_info=True)
                    continue

            logger.info(f"Processing complete. {processed_count} files generated.")
            return processed_count, today_str

        except Exception as e:
            logger.critical(f"Critical error in ExcelHandler: {e}", exc_info=True)
            raise e
