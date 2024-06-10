import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QLineEdit, QLabel
from src.data_processing import load_excel, save_excel
from src.modify_template import modify_excel_data
from src.compare_files import compare_excel_files

class ExcelMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.source_file = ""
        self.target_file = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel Merger')
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        self.source_label = QLabel('Source File:', self)
        layout.addWidget(self.source_label)

        self.source_line_edit = QLineEdit(self)
        layout.addWidget(self.source_line_edit)

        self.source_btn = QPushButton('Select Source File', self)
        self.source_btn.clicked.connect(self.select_source_file)
        layout.addWidget(self.source_btn)

        self.target_label = QLabel('Target File:', self)
        layout.addWidget(self.target_label)

        self.target_line_edit = QLineEdit(self)
        layout.addWidget(self.target_line_edit)

        self.target_btn = QPushButton('Select Target File', self)
        self.target_btn.clicked.connect(self.select_target_file)
        layout.addWidget(self.target_btn)

        self.merge_btn = QPushButton('Merge Files', self)
        self.merge_btn.clicked.connect(self.handle_merge_files)
        layout.addWidget(self.merge_btn)

        self.modify_btn = QPushButton('Modify Template', self)
        self.modify_btn.clicked.connect(self.modify_template)
        layout.addWidget(self.modify_btn)

        self.compare_btn = QPushButton('Compare Files', self)
        self.compare_btn.clicked.connect(self.compare_files)
        layout.addWidget(self.compare_btn)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_source_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Source File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.source_file = file_name
            self.source_line_edit.setText(file_name)

    def select_target_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Target File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.target_file = file_name
            self.target_line_edit.setText(file_name)

    def handle_merge_files(self):
        if not self.source_file or not self.target_file:
            print("Please select both source and target files.")
            return

        merge_files(self.source_file, self.target_file, "MergedSheet")

    def modify_template(self):
        if not self.source_file or not self.target_file:
            print("Please select both source and target files.")
            return

        modify_excel_data(self.source_file, self.target_file, "ModifiedSheet")

    def compare_files(self):
        hand_file, _ = QFileDialog.getOpenFileName(self, "Select Hand File", "", "Excel Files (*.xlsx);;All Files (*)", options=QFileDialog.Options())
        if not hand_file:
            return
        
        target_file, _ = QFileDialog.getOpenFileName(self, "Select Target File", "", "Excel Files (*.xlsx);;All Files (*)", options=QFileDialog.Options())
        if not target_file:
            return

        compare_excel_files(target_file, hand_file, 'output/output_comparison.xlsx')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelMergerApp()
    ex.show()
    sys.exit(app.exec_())