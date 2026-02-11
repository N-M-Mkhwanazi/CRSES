import sys
import json
import pandas as pd
import openpyxl
import logging
import difflib
import re
from copy import copy
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QSizePolicy,
    QLineEdit, QProgressBar, QDoubleSpinBox, QSpacerItem,)
from PyQt6.QtCore import Qt
from openpyxl.utils import column_index_from_string
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment, Protection
from openpyxl.cell.cell import Cell

with open("C:\\Users\\NMMKHWANAZI\\Documents\\Code\\Appending Alpha\\Map.json", 'r') as f:
    column_match_mapping = json.load(f)

columns_for_flexible_validation = [""]

# --- Logger Setup ---
logging.basicConfig(
    filename='appender.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def find_fuzzy_match(target_name, available_names, min_score=0.7):
    def _norm(name):
        return re.sub(r'\W+', '', str(name).lower())

    target_name_normalized = _norm(target_name)
    available_names_normalized = [_norm(n) for n in available_names]
    
    matches = difflib.get_close_matches(target_name_normalized, available_names_normalized, n=1, cutoff=min_score)
    
    if matches:
        matched_name_lower = matches[0]
        for name in available_names:
            if _norm(name) == matched_name_lower:
                return name
    return None

def percent_within(a, b, tol_percent):
    #Checks if a is within a certain percentage tolerance of b.
    try:
        a_val = float(str(a).replace(',', '.'))
        b_val = float(str(b).replace(',', '.'))
        # Handle zero-division case
        if abs(b_val) < 1e-9:
            return abs(a_val) < 1e-9
        return abs(a_val - b_val) <= (tol_percent / 100.0) * abs(b_val)
    except (ValueError, TypeError):
        return False

def copy_cell_style(source_cell: Cell, target_cell: Cell):

    #Safely clones style attributes fromn a source cell to a target cell.  # the last row in alpha stats
    if source_cell.has_style:
        if source_cell.font:
            target_cell.font = copy(source_cell.font)
        if source_cell.fill:
            target_cell.fill = copy(source_cell.fill)
        if source_cell.border:
            target_cell.border = copy(source_cell.border)
        if source_cell.alignment:
            target_cell.alignment = copy(source_cell.alignment)
        if source_cell.protection:
            target_cell.protection = copy(source_cell.protection)

        target_cell.number_format = source_cell.number_format

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Data Appender")
        self.resize(600, 400)
        self.setStyleSheet("background-color: #FFFFFF;")
        
        self.alpha_file_path = ""
        self.eskom_file_path = ""

        self.initUI()
        self.run_button.setEnabled(False)

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # CRSES Logo Header
        logo_layout = QHBoxLayout()
        logo_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        sun_motif = QLabel("⬤")
        sun_motif.setStyleSheet("color: #FFB81C; font-size: 40px; font-weight: bold;")
        logo_layout.addWidget(sun_motif)
        
        title_label = QLabel("CRSES")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #7A003C; font-size: 36px; font-weight: bold;")
        logo_layout.addWidget(title_label)

        main_layout.addLayout(logo_layout)
        
        subtitle_label = QLabel("Excel Data Appender")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle_label.setStyleSheet("color: #4D4D4D; font-size: 18px;")
        main_layout.addWidget(subtitle_label)

        self.status_label = QLabel("Ready: Please select your files.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("background-color: #F8F8F8; color: #4D4D4D; padding: 10px; border-radius: 8px; border: 1px solid #E0E0E0;")
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        main_layout.addWidget(self.status_label)

        # File Selection Layout
        file_selection_layout = QVBoxLayout()

        alpha_layout = QHBoxLayout()
        self.alpha_label = QLineEdit()
        self.alpha_label.setPlaceholderText("Select primary Excel file")
        self.alpha_label.setStyleSheet("color: #4D4D4D; padding: 5px;")
        self.alpha_label.setReadOnly(True)
        alpha_button = QPushButton("Browse Alpha File")
        alpha_button.setStyleSheet("QPushButton { background-color: #D3D3D3; color: #4D4D4D; border-radius: 10px; }"
                                 "QPushButton:hover { background-color: #808080; }"
                                 "QPushButton:pressed { background-color: #CC9417; }")
        alpha_button.clicked.connect(self.select_alpha_file)
        alpha_layout.addWidget(self.alpha_label)
        alpha_layout.addWidget(alpha_button)
        file_selection_layout.addLayout(alpha_layout)

        eskom_layout = QHBoxLayout()
        self.eskom_label = QLineEdit()
        self.eskom_label.setPlaceholderText("Select secondary Excel file")
        self.eskom_label.setStyleSheet("color: #4D4D4D; padding: 5px;")
        self.eskom_label.setReadOnly(True)
        eskom_button = QPushButton("Browse Eskom File")
        eskom_button.setStyleSheet("QPushButton { background-color: #D3D3D3; color: #4D4D4D; border-radius: 10px; }"
                                 "QPushButton:hover { background-color: #808080; }"
                                 "QPushButton:pressed { background-color: #CC9417; }")
        eskom_button.clicked.connect(self.select_eskom_file)
        eskom_layout.addWidget(self.eskom_label)
        eskom_layout.addWidget(eskom_button)
        file_selection_layout.addLayout(eskom_layout)
        
        main_layout.addLayout(file_selection_layout)

        # Tolerance & Progress Layout
        options_layout = QHBoxLayout()
        options_layout.addWidget(QLabel("Percent Tolerance:"))
        self.tolerance_spinbox = QDoubleSpinBox()
        self.tolerance_spinbox.setRange(0.0, 100.0)
        self.tolerance_spinbox.setValue(10.0)
        self.tolerance_spinbox.setSingleStep(1.0)
        self.tolerance_spinbox.setSingleStep(-1.0)
        self.tolerance_spinbox.setStyleSheet("padding: 5px; color: #4D4D4D;")
        options_layout.addWidget(self.tolerance_spinbox)
        options_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        main_layout.addLayout(options_layout)

        # Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("QProgressBar { text-align: center; color: #4D4D4D; }"
                                     "QProgressBar::chunk { background-color: #FFB81C; border-radius: 5px; }")
        main_layout.addWidget(self.progress_bar)

        # Run Button
        self.run_button = QPushButton("Run Appender")
        self.run_button.setFixedSize(200, 50)
        self.run_button.setStyleSheet(
            "QPushButton { background-color: #7A003C; color: white; border-radius: 10px; }"
            "QPushButton:hover { background-color: #A60052; }"
            "QPushButton:disabled { background-color: #D3D3D3; }"
        )
        self.run_button.clicked.connect(self.run_append)
        
        run_button_layout = QHBoxLayout()
        run_button_layout.addStretch()
        run_button_layout.addWidget(self.run_button)
        run_button_layout.addStretch()
        main_layout.addLayout(run_button_layout)

    def select_alpha_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Alpha Stats Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.alpha_file_path = file_name
            self.alpha_label.setText(file_name)
            self.status_label.setText("Alpha file selected.")
            self.check_run_button_state()

    def select_eskom_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Eskom Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.eskom_file_path = file_name
            self.eskom_label.setText(file_name)
            self.status_label.setText("Eskom file selected.")
            self.check_run_button_state()

    def check_run_button_state(self):
        if self.alpha_file_path and self.eskom_file_path:
            self.run_button.setEnabled(True)

    def run_append(self):
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.status_label.setText("Loading data...")
            self.progress_bar.setValue(10)

            tolerance = self.tolerance_spinbox.value()
            eskom_df, alpha_workbook, alpha_worksheet, last_data_row, alpha_stats_header_map = self._load_files()
            
            if eskom_df is None or alpha_workbook is None:
                self.progress_bar.setVisible(False)
                return
            
            self.progress_bar.setValue(30)
            self.status_label.setText("Validating data...")
            
            matching_row_index = self._validate_data(alpha_worksheet, eskom_df, last_data_row, alpha_stats_header_map, tolerance)
            
            if matching_row_index is None:
                self.progress_bar.setVisible(False)
                return
            
            self.progress_bar.setValue(60)
            self.status_label.setText("Appending new rows...")
            
            self._append_and_save(alpha_workbook, alpha_worksheet, eskom_df, matching_row_index, last_data_row, alpha_stats_header_map)
            
            self.progress_bar.setValue(100)
            self.status_label.setText("Appending complete!")
            
            QMessageBox.information(self, "Success", "Data appended successfully!")
            logging.info("Data appending process completed successfully.")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            logging.error("An unexpected error occurred during the append process.", exc_info=True)
        finally:
            self.progress_bar.setVisible(False)

    def _load_files(self):
        try:
            eskom_df = pd.read_excel(self.eskom_file_path, engine='openpyxl')
        except FileNotFoundError:
            QMessageBox.critical(self, "Error", "Eskom file not found.")
            logging.error("Eskom file not found.")
            return None, None, None, None, None
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error reading Eskom file: {e}")
            logging.error("Error reading Eskom file.", exc_info=True)
            return None, None, None, None, None

        eskom_date_col_name = find_fuzzy_match('date time hour beginning', eskom_df.columns.tolist())
        if eskom_date_col_name:
            try:
                eskom_df[eskom_date_col_name] = pd.to_datetime(eskom_df[eskom_date_col_name], errors='coerce').dt.floor('H')
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error converting Eskom date column to datetime: {e}")
                logging.error("Error converting Eskom date column.", exc_info=True)
                return None, None, None, None, None
        else:
            QMessageBox.critical(self, "Error", "Could not locate 'Date time hour beginning' column in Eskom file.")
            return None, None, None, None, None

        try:
            alpha_workbook = openpyxl.load_workbook(self.alpha_file_path)
        except FileNotFoundError:
            QMessageBox.critical(self, "Error", "Alpha file not found.")
            logging.error("Alpha file not found.")
            return None, None, None, None, None
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error reading Alpha file: {e}")
            logging.error("Error reading Alpha file.", exc_info=True)
            return None, None, None, None, None

        alpha_worksheet = alpha_workbook.active
        
        alpha_stats_header_map = {}
        
        for idx, cell in enumerate(alpha_worksheet[1], start=1):
            header_name = str(cell.value).strip().lower() if cell.value else ""
            if header_name:
                if header_name in alpha_stats_header_map:
                    alpha_stats_header_map[header_name].append(idx)
                else:
                    alpha_stats_header_map[header_name] = [idx]
        
        last_data_row = alpha_worksheet.max_row
        while last_data_row > 1:
            row_data = [cell.value for cell in alpha_worksheet[last_data_row]]
            if any(val is not None and str(val).strip() != "" for val in row_data):
                break
            last_data_row -= 1
        
        return eskom_df, alpha_workbook, alpha_worksheet, last_data_row, alpha_stats_header_map

    def _get_alpha_column_index(self, mapping_tuple, alpha_stats_header_map):
        alpha_column = mapping_tuple[1]
        alpha_column_normalized = alpha_column.strip().lower()
        
        if alpha_column_normalized in alpha_stats_header_map:
            return alpha_stats_header_map[alpha_column_normalized][0]
        
        fuzzy_match = find_fuzzy_match(alpha_column_normalized, list(alpha_stats_header_map.keys()))
        if fuzzy_match:
            return alpha_stats_header_map[fuzzy_match][0]
        
        return None

    def _validate_data(self, alpha_worksheet, eskom_data_frame, last_data_row, alpha_stats_header_map, tolerance):
        alpha_date_cols = [col.lower() for col in ["Year", "Month", "Day", "Hour"]]
        
        if not all(col in alpha_stats_header_map for col in alpha_date_cols):
            QMessageBox.critical(self, "Error", "Missing one or more date/time columns (Year, Month, Day, Hour) in Alpha Stats.")
            logging.error("Missing date/time columns in Alpha Stats file.")
            return None

        try:
            alpha_last_datetime = pd.to_datetime(
                f"{alpha_worksheet.cell(row=last_data_row, column=alpha_stats_header_map['year'][0]).value}-"
                f"{alpha_worksheet.cell(row=last_data_row, column=alpha_stats_header_map['month'][0]).value}-"
                f"{alpha_worksheet.cell(row=last_data_row, column=alpha_stats_header_map['day'][0]).value} "
                f"{alpha_worksheet.cell(row=last_data_row, column=alpha_stats_header_map['hour'][0]).value}:00:00"
            ).floor('H')
        except (ValueError, TypeError) as e:
            QMessageBox.critical(self, "Error", f"Error parsing date from Alpha Stats last row: {e}")
            logging.error("Failed to parse date from Alpha Stats last row.", exc_info=True)
            return None

        eskom_date_col_name = find_fuzzy_match('date time hour beginning', eskom_data_frame.columns.tolist())
        if not eskom_date_col_name:
            QMessageBox.critical(self, "Error", "Could not find 'date time hour beginning' column in Eskom file.")
            return None
        
        matching_indices = eskom_data_frame.index[eskom_data_frame[eskom_date_col_name] == alpha_last_datetime]

        if matching_indices.empty:
            QMessageBox.warning(self, "Warning", f"No matching timestamp ('{alpha_last_datetime}') found in Eskom data.")
            logging.warning(f"No matching timestamp '{alpha_last_datetime}' found in Eskom data.")
            return None

        matching_row_index = matching_indices[0]
        validation_is_ok = True

        for mapping_tuple in column_match_mapping[1:]:
            eskom_column = mapping_tuple[0]
            alpha_column = mapping_tuple[1]
            
            eskom_column_found = find_fuzzy_match(eskom_column, eskom_data_frame.columns.tolist())
            logging.info(f"Fuzzy matched '{eskom_column}' -> '{eskom_column_found}'")

            if not eskom_column_found:
                QMessageBox.warning(self, "Warning", f"Column '{eskom_column}' is missing from the Eskom file. Skipping validation for this column.")
                logging.warning(f"Column '{eskom_column}' is missing in the Eskom file, skipping validation.")
                continue

            alpha_column_index = self._get_alpha_column_index(mapping_tuple, alpha_stats_header_map)
            logging.debug(f"Looking up Alpha column '{alpha_column}' -> found index '{alpha_column_index}'")
            if not alpha_column_index:
                QMessageBox.warning(self, "Warning", f"Column '{alpha_column}' is missing from the Alpha file. Skipping validation for this column.")
                logging.warning(f"Column '{alpha_column}' is missing in the Alpha file, skipping validation.")
                continue

            alpha_value = alpha_worksheet.cell(row=last_data_row, column=alpha_column_index).value
            eskom_value = eskom_data_frame.loc[matching_row_index, eskom_column_found]
            
            if not percent_within(alpha_value, eskom_value, tolerance):
                QMessageBox.critical(self, "Mismatch", f"Mismatch found in '{alpha_column}': Alpha value is '{alpha_value}', Eskom value is '{eskom_value}' (Exceeds tolerance)")
                logging.error(f"Percentage validation failed for '{alpha_column}': Alpha='{alpha_value}', Eskom='{eskom_value}'")
                validation_is_ok = False
                break
            
        return matching_row_index if validation_is_ok else None

    def _append_and_save(self, alpha_workbook, alpha_worksheet, eskom_data_frame, matching_row_index, last_data_row, alpha_stats_header_map):
        pos = eskom_data_frame.index.get_loc(matching_row_index)
        new_rows_to_append = eskom_data_frame.iloc[pos + 1:].copy()
        
        if new_rows_to_append.empty:
            QMessageBox.information(self, "Info", "There are no new rows to append.")
            return

        next_append_row = last_data_row + 1
        
        eskom_date_col_name = find_fuzzy_match('date time hour beginning', eskom_data_frame.columns.tolist())
        if not eskom_date_col_name:
            raise ValueError("Could not find date time column for appending.")

        # Store the entire cell object for style cloning
        last_row_cells = {}
        for col_idx in range(1, alpha_worksheet.max_column + 1):
            last_row_cells[col_idx] = alpha_worksheet.cell(row=last_data_row, column=col_idx)
        
        for _, new_row in new_rows_to_append.iterrows():
            eskom_datetime = pd.to_datetime(new_row[eskom_date_col_name])
            
            alpha_date_cols = ["Year", "Week", "Month", "Day", "Hour"]
            alpha_date_values = [
                eskom_datetime.year,
                eskom_datetime.isocalendar().week,
                eskom_datetime.month,
                eskom_datetime.day,
                eskom_datetime.hour
            ]
            
            for alpha_col_name, alpha_value in zip(alpha_date_cols, alpha_date_values):
                alpha_col_indices = alpha_stats_header_map.get(alpha_col_name.lower())
                if alpha_col_indices:
                    alpha_worksheet.cell(row=next_append_row, column=alpha_col_indices[0]).value = alpha_value
            
            for mapping_tuple in column_match_mapping[1:]:
                eskom_column = mapping_tuple[0]
                alpha_column = mapping_tuple[1]
                
                eskom_column_found = find_fuzzy_match(eskom_column, new_row.index.tolist())
                logging.debug(f"Appending - attempting to fuzzy match Eskom column '{eskom_column}' -> found '{eskom_column_found}'")
                
                if eskom_column_found:
                    alpha_col_index = self._get_alpha_column_index(mapping_tuple, alpha_stats_header_map)
                    logging.debug(f"Appending - looking up Alpha column '{alpha_column}' -> found index '{alpha_col_index}'")
                    if alpha_col_index:
                        value_to_append = new_row[eskom_column_found]
                        if isinstance(value_to_append, str):
                            value_to_append = value_to_append.replace(',', '.')
                        
                        new_cell = alpha_worksheet.cell(row=next_append_row, column=alpha_col_index)
                        new_cell.value = value_to_append
                        
                        # Copy styles from the last row
                        if alpha_col_index in last_row_cells:
                            copy_cell_style(last_row_cells[alpha_col_index], new_cell)
                else:
                    logging.warning(f"Column '{eskom_column}' not found in new row from Eskom file. Skipping.")
            next_append_row += 1

        alpha_workbook.save(self.alpha_file_path)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()