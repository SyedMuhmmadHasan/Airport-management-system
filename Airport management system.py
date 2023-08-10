import sys
import sqlite3
import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt

class Database:
    def __init__(self, db_name="airport.db"):
        self.connection = sqlite3.connect(db_name)
        self.cursor = self.connection.cursor()
        self.create_tables()

    def create_tables(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS flights (
                id INTEGER PRIMARY KEY,
                flight_number TEXT,
                departure TEXT,
                destination TEXT
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS passengers (
                id INTEGER PRIMARY KEY,
                passenger_name TEXT,
                flight_number TEXT,
                FOREIGN KEY (flight_number) REFERENCES flights(flight_number)
            )
        ''')
        
        self.connection.commit()

    def add_flight(self, flight_number, departure, destination):
        self.cursor.execute('''
            INSERT INTO flights (flight_number, departure, destination) VALUES (?, ?, ?)
        ''', (flight_number, departure, destination))
        self.connection.commit()

    def add_passenger(self, passenger_name, flight_number):
        self.cursor.execute('''
            INSERT INTO passengers (passenger_name, flight_number) VALUES (?, ?)
        ''', (passenger_name, flight_number))
        self.connection.commit()

    def get_passengers(self):
        self.cursor.execute('''
            SELECT passengers.passenger_name, flights.flight_number, flights.departure, flights.destination
            FROM passengers
            JOIN flights ON passengers.flight_number = flights.flight_number
        ''')
        passengers = self.cursor.fetchall()
        return [{'passenger_name': p[0], 'flight_number': p[1], 'departure': p[2], 'destination': p[3]} for p in passengers]

    def reset_database(self):
        self.cursor.execute("DELETE FROM flights")
        self.cursor.execute("DELETE FROM passengers")
        self.connection.commit()

    def close(self):
        self.connection.close()

class AirportManagementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Airport Management System")
        self.setGeometry(100, 100, 800, 600)
        
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        
        self.layout = QVBoxLayout()
        
        self.flight_info_label = QLabel("Flight Information:")
        self.layout.addWidget(self.flight_info_label)
        
        self.flight_number_label = QLabel("Flight Number:")
        self.flight_number_input = QLineEdit()
        self.layout.addWidget(self.flight_number_label)
        self.layout.addWidget(self.flight_number_input)
        
        self.departure_label = QLabel("Departure:")
        self.departure_input = QLineEdit()
        self.layout.addWidget(self.departure_label)
        self.layout.addWidget(self.departure_input)
        
        self.destination_label = QLabel("Destination:")
        self.destination_input = QLineEdit()
        self.layout.addWidget(self.destination_label)
        self.layout.addWidget(self.destination_input)
        
        self.add_flight_button = QPushButton("Add Flight")
        self.layout.addWidget(self.add_flight_button)
        
        self.passenger_info_label = QLabel("Passenger Information:")
        self.layout.addWidget(self.passenger_info_label)
        
        self.passenger_name_label = QLabel("Passenger Name:")
        self.passenger_name_input = QLineEdit()
        self.layout.addWidget(self.passenger_name_label)
        self.layout.addWidget(self.passenger_name_input)
        
        self.flight_number_label_2 = QLabel("Flight Number:")
        self.flight_number_input_2 = QLineEdit()
        self.layout.addWidget(self.flight_number_label_2)
        self.layout.addWidget(self.flight_number_input_2)
        
        self.add_passenger_button = QPushButton("Add Passenger")
        self.layout.addWidget(self.add_passenger_button)
        
        self.log_text = QTextEdit()
        self.layout.addWidget(self.log_text)
        
        self.passenger_table = QTableWidget()
        self.passenger_table.setColumnCount(4)
        self.passenger_table.setHorizontalHeaderLabels(["Passenger Name", "Flight Number", "Departure", "Destination"])
        self.layout.addWidget(self.passenger_table)
        
        self.save_info_button = QPushButton("Save Info")
        self.layout.addWidget(self.save_info_button)
        
        self.remove_info_button = QPushButton("Remove Info")
        self.layout.addWidget(self.remove_info_button)
        
        self.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
            }
            QLineEdit, QTextEdit {
                padding: 5px;
                font-size: 12px;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 3px;
                font-size: 12px;
            }
            QTableWidget {
                margin-top: 10px;
            }
        """)

        self.central_widget.setLayout(self.layout)
        self.add_flight_button.clicked.connect(self.add_new_flight)
        self.add_passenger_button.clicked.connect(self.add_new_passenger)
        self.save_info_button.clicked.connect(self.save_info)
        self.remove_info_button.clicked.connect(self.remove_info)
        self.db = Database()

    def add_new_flight(self):
        flight_number = self.flight_number_input.text()
        departure = self.departure_input.text()
        destination = self.destination_input.text()

        if flight_number and departure and destination:
            self.db.add_flight(flight_number, departure, destination)
            self.log_text.append(f"Flight {flight_number} added: {departure} to {destination}")
            self.flight_number_input.clear()
            self.departure_input.clear()
            self.destination_input.clear()
        else:
            self.log_text.append("Please enter valid flight information.")

    def add_new_passenger(self):
        passenger_name = self.passenger_name_input.text()
        flight_number = self.flight_number_input_2.text()

        if passenger_name and flight_number:
            self.db.add_passenger(passenger_name, flight_number)
            self.log_text.append(f"Passenger {passenger_name} added to flight {flight_number}")
            self.passenger_name_input.clear()
            self.flight_number_input_2.clear()
            self.update_passenger_table()
        else:
            self.log_text.append("Please enter valid passenger information.")

    def update_passenger_table(self):
        passengers = self.db.get_passengers()
        self.passenger_table.setRowCount(len(passengers))
        
        for row, passenger in enumerate(passengers):
            self.passenger_table.setItem(row, 0, QTableWidgetItem(passenger['passenger_name']))
            self.passenger_table.setItem(row, 1, QTableWidgetItem(passenger['flight_number']))
            self.passenger_table.setItem(row, 2, QTableWidgetItem(passenger['departure']))
            self.passenger_table.setItem(row, 3, QTableWidgetItem(passenger['destination']))

    def save_info(self):
        passengers = self.db.get_passengers()
        if passengers:
            excel_file_path, _ = QFileDialog.getSaveFileName(self, "Save Passenger Info", "", "Excel Files (*.xlsx)")
            if excel_file_path:
                self.save_to_excel(passengers, excel_file_path)
                self.log_text.append("Passenger info saved to Excel.")
        else:
            self.log_text.append("No passenger info to save.")

    def save_to_excel(self, passenger_data, excel_file_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Passenger Name", "Flight Number", "Departure", "Destination"])
        for passenger in passenger_data:
            ws.append([passenger['passenger_name'], passenger['flight_number'], passenger['departure'], passenger['destination']])
        wb.save(excel_file_path)

    def remove_info(self):
        confirmation = QMessageBox.question(self, "Confirm Removal", "Are you sure you want to remove all saved info?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if confirmation == QMessageBox.Yes:
            self.passenger_table.clearContents()
            self.passenger_table.setRowCount(0)
            self.db.reset_database()
            self.log_text.append("All saved info removed.")
        else:
            self.log_text.append("Removal canceled.")

    def closeEvent(self, event):
        self.db.close()

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Set Fusion style for consistent look on all platforms
    window = AirportManagementApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
