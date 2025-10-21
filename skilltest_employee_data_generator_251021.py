import sys
import random
from datetime import datetime, timedelta
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog
)
from PySide6.QtCore import Qt
import pandas as pd


class EmployeeGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_folder = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Employee Data Generator")
        self.setGeometry(100, 100, 500, 300)

        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)

        # Number of employees input
        emp_layout = QHBoxLayout()
        emp_label = QLabel("Number of Employees:")
        self.emp_input = QLineEdit()
        self.emp_input.setPlaceholderText("e.g., 100")
        emp_layout.addWidget(emp_label)
        emp_layout.addWidget(self.emp_input)
        layout.addLayout(emp_layout)

        # Folder selection
        folder_layout = QHBoxLayout()
        self.folder_label = QLabel("No folder selected")
        self.folder_btn = QPushButton("Select Folder")
        self.folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.folder_label)
        folder_layout.addWidget(self.folder_btn)
        layout.addLayout(folder_layout)

        # Generate data button
        self.generate_btn = QPushButton("Generate Data")
        self.generate_btn.clicked.connect(self.generate_data)
        layout.addWidget(self.generate_btn)

        # Export button
        self.export_btn = QPushButton("Export to Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        self.export_btn.setEnabled(False)
        layout.addWidget(self.export_btn)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-weight: bold; color: green;")
        layout.addWidget(self.status_label)

        # Data storage
        self.df = None

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.selected_folder = folder
            self.folder_label.setText(f"Folder: {folder}")
            self.status_label.setText("")

    def generate_data(self):
        try:
            num_employees = int(self.emp_input.text())
            if num_employees <= 0:
                raise ValueError("Must be positive")
        except ValueError:
            self.status_label.setText("⚠️ Enter a valid number of employees")
            self.status_label.setStyleSheet("font-weight: bold; color: red;")
            return

        # Generate synthetic data
        first_names = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", 
                       "Michael", "Linda", "William", "Elizabeth", "David", "Barbara",
                       "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah",
                       "Charles", "Karen", "Christopher", "Nancy", "Daniel", "Lisa"]
        
        last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia",
                      "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez",
                      "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore"]
        
        departments = ["IT", "HR", "Operations", "Administration", "Finance"]
        
        start_date = datetime(2020, 1, 1)
        end_date = datetime.now()
        date_range = (end_date - start_date).days

        data = []
        for i in range(1, num_employees + 1):
            full_name = f"{random.choice(first_names)} {random.choice(last_names)}"
            department = random.choice(departments)
            salary = random.randint(25000, 120000)
            hire_date = start_date + timedelta(days=random.randint(0, date_range))
            
            data.append({
                "emp_id": i,
                "full_name": full_name,
                "department": department,
                "salary": salary,
                "hire_date": hire_date.strftime("%Y-%m-%d")
            })

        self.df = pd.DataFrame(data)
        self.status_label.setText(f"✅ Generated {num_employees} employees")
        self.status_label.setStyleSheet("font-weight: bold; color: green;")
        self.export_btn.setEnabled(True)

    def export_to_excel(self):
        if self.df is None:
            self.status_label.setText("⚠️ Generate data first")
            self.status_label.setStyleSheet("font-weight: bold; color: red;")
            return

        if not self.selected_folder:
            self.status_label.setText("⚠️ Select a folder first")
            self.status_label.setStyleSheet("font-weight: bold; color: red;")
            return

        # File path
        file_path = Path(self.selected_folder) / "employees.xlsx"

        # Create summary data
        summary_df = self.df.groupby("department")["salary"].mean().reset_index()
        summary_df.columns = ["Department", "Average Salary"]
        summary_df["Average Salary"] = summary_df["Average Salary"].round(2)

        # Export timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Write to Excel with multiple sheets
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            self.df.to_excel(writer, sheet_name="Employees", index=False)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            
            # Add timestamp to Summary sheet
            summary_sheet = writer.sheets["Summary"]
            timestamp_row = len(summary_df) + 3
            summary_sheet.cell(row=timestamp_row, column=1, value="Export Timestamp:")
            summary_sheet.cell(row=timestamp_row, column=2, value=timestamp)

        self.status_label.setText(f"✅ File Generated: {file_path}")
        self.status_label.setStyleSheet("font-weight: bold; color: green;")


def main():
    app = QApplication(sys.argv)
    window = EmployeeGeneratorApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()