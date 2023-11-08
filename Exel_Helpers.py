import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QCheckBox, QLineEdit, QVBoxLayout, QWidget, QMessageBox, QFileDialog

class ReportGenerator(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Звіти з комутаторів")
        self.setGeometry(100, 100, 400, 200)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.label = QLabel("Введіть шлях до Excel-файлу:")
        self.layout.addWidget(self.label)

        self.file_path_entry = QLineEdit(self)
        self.layout.addWidget(self.file_path_entry)

        self.problem_checkbox = QCheckBox("Проблеми", self)
        self.layout.addWidget(self.problem_checkbox)

        self.planova_zam_checkbox = QCheckBox("Планова заміна УПС", self)
        self.layout.addWidget(self.planova_zam_checkbox)

        self.process_button = QPushButton("Розпочати обробку", self)
        self.process_button.clicked.connect(self.process_reports)
        self.layout.addWidget(self.process_button)

        self.central_widget.setLayout(self.layout)

    def process_reports(self):
        file_path = self.file_path_entry.text()
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Помилка при зчитуванні файлу: {str(e)}")
            return

        if self.problem_checkbox.isChecked():
            self.problem(df)
        if self.planova_zam_checkbox.isChecked():
            self.planova_zam(df)

        QMessageBox.information(self, "Завершено", "Готово! Звіти збережено.")

    def problem(self, df):
        # Розділити дані на два DataFrames за умовою "sw-230" та "nc-230"
        km = df[df["Unnamed: 2"].str.contains("sw-230|nc-230", case=False, na=False)].copy()
        vn = df[~df["Unnamed: 2"].str.contains("sw-230|nc-230", case=False, na=False)].copy()

        # Змінити формат дати в DataFrame
        km["Unnamed: 0"] = pd.to_datetime(km["Unnamed: 0"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")
        km["Unnamed: 8"] = pd.to_datetime(km["Unnamed: 8"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")

        vn["Unnamed: 0"] = pd.to_datetime(vn["Unnamed: 0"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")
        vn["Unnamed: 8"] = pd.to_datetime(vn["Unnamed: 8"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")

        # Зберегти два DataFrames в окремі Excel-файли
        km.to_excel('Проблеми з комутаторами ХМ.xlsx', index=False)
        vn.to_excel('Проблеми з комутаторами ВН.xlsx', index=False)

    def planova_zam(self, df):
        # Ваш код для генерації звіту "Планова заміна УПС"
        km = df[(df["Unnamed: 2"].str.contains("sw-230|nc-230", case=False, na=False)) & (~df["Unnamed: 12"].isna())].copy()
        vn = df[(~df["Unnamed: 2"].str.contains("sw-230|nc-230", case=False, na=False)) & (~df["Unnamed: 12"].isna())].copy()

        # Змінити формат дати в DataFrame
        km["Unnamed: 0"] = pd.to_datetime(km["Unnamed: 0"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")
        km["Unnamed: 8"] = pd.to_datetime(km["Unnamed: 8"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")

        vn["Unnamed: 0"] = pd.to_datetime(vn["Unnamed: 0"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")
        vn["Unnamed: 8"] = pd.to_datetime(vn["Unnamed: 8"], format="%d.%m.%y").dt.strftime("%Y-%m-%d")

        # Зберегти два DataFrames в окремі Excel-файли
        km.to_excel('Планова заміна УПС ХМ.xlsx', index=False)
        vn.to_excel('Планова заміна УПС ВН.xlsx', index=False)

def main():
    app = QApplication(sys.argv)
    window = ReportGenerator()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
