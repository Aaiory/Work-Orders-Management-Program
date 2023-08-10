import os

import FileArrangement
import sys
from PySide6.QtGui import QFont, QIcon, Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QLabel, QTextEdit, QStatusBar,
    QGridLayout, QToolBar, QStyleFactory, QComboBox, QPushButton,
    QLineEdit, QWidget, QFileDialog, QMessageBox, QTableWidget, QAbstractItemView, QTableWidgetItem, QHeaderView
)
from CacheHandling import *
from FileInfoFill import get_work_hours


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Work Order Manager")
        self.setGeometry(100, 100, 600, 50)  # Updated window size

        # Set program icon
        program_icon = QIcon("icons/program.ico")
        self.setWindowIcon(program_icon)

        main_layout = QVBoxLayout()
        main_widget = QWidget()

        # Create toolbar
        self.toolbar = QToolBar()
        self.addToolBar(self.toolbar)

        # Add theme selection combo box
        theme_combo = QComboBox()
        theme_combo.addItems(QStyleFactory.keys())
        theme_combo.currentIndexChanged.connect(self.select_theme)
        self.toolbar.addWidget(theme_combo)

        # Add font size selection combo box
        font_size_combo = QComboBox()
        font_size_combo.addItems(["10", "12", "14", "16", "18", "20"])
        font_size_combo.currentIndexChanged.connect(self.select_font_size)
        self.toolbar.addWidget(font_size_combo)

        # Add clear folders button
        clear_folders_button = QPushButton("Clear Work/Move Folders")
        clear_folders_button.setToolTip("Clear Work Folder and Move Folder")
        clear_folders_button.clicked.connect(self.clear_folders)
        self.toolbar.addWidget(clear_folders_button)

        # Load the last used theme
        last_used_theme = load_last_used_theme()
        if last_used_theme:
            theme_combo.setCurrentText(last_used_theme)
            QApplication.setStyle(QStyleFactory.create(last_used_theme))

        # Create a grid layout
        grid_layout = QGridLayout()
        grid_layout.setVerticalSpacing(2)  # Set vertical spacing to 2

        # Root folder path search bar layout
        self.path_label = QLabel("Work Folder:")
        self.path_input = QLineEdit()
        self.browse_button = QPushButton()
        self.browse_button.setIcon(QIcon("icons/browse.ico"))
        self.browse_button.setFixedSize(50, 30)
        self.browse_button.clicked.connect(self.open_folder_dialog)
        self.clear_button = QPushButton()
        self.clear_button.setIcon(QIcon("icons/clear.ico"))
        self.clear_button.setFixedSize(50, 30)
        self.clear_button.clicked.connect(self.clear_folder_path)

        grid_layout.addWidget(self.path_label, 0, 0)
        grid_layout.addWidget(self.path_input, 0, 1)
        grid_layout.addWidget(self.browse_button, 0, 2)
        grid_layout.addWidget(self.clear_button, 0, 3)

        # Option buttons layout
        self.option_buttons = []
        options = [
            {"text": "Rename PDF Files to Work Order #", "icon": "icons/rename.ico"},
            {"text": "Create folders for PDF files", "icon": "icons/folder.ico"},
            {"text": "Merge files by folder", "icon": "icons/merge.ico"}
        ]
        for row, option in enumerate(options, start=1):
            button = QPushButton(option["text"])
            button.setIcon(QIcon(option["icon"]))
            button.setToolTip(option["text"])
            button.clicked.connect(self.handle_option)
            self.option_buttons.append(button)
            grid_layout.addWidget(button, row, 0, 1, 4)

        # Root folder path search bar layout
        self.export_label = QLabel("Move Folder:")
        self.export_input = QLineEdit()
        self.export_browse_button = QPushButton()
        self.export_browse_button.setIcon(QIcon("icons/browse.ico"))
        self.export_browse_button.setFixedSize(50, 30)
        self.export_browse_button.clicked.connect(self.open_export_folder_dialog)
        self.export_clear_button = QPushButton()
        self.export_clear_button.setIcon(QIcon("icons/move-files.ico"))
        self.export_clear_button.setFixedSize(50, 30)
        self.export_clear_button.clicked.connect(self.move_merged_files)

        grid_layout.addWidget(self.export_label, len(options) + 1, 0)
        grid_layout.addWidget(self.export_input, len(options) + 1, 1)
        grid_layout.addWidget(self.export_browse_button, len(options) + 1, 2)
        grid_layout.addWidget(self.export_clear_button, len(options) + 1, 3)

        # Content display
        # Replace self.text_edit with a QTableWidget
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(3)  # One column to display the file names
        self.table_widget.setHorizontalHeaderLabels(["Files in directory", "Time in", "Time Out"])
        self.table_widget.setSelectionMode(QAbstractItemView.SingleSelection)  # Allow selecting a single item

        # Set resize mode for all columns to Stretch
        for column in range(self.table_widget.columnCount()):
            self.table_widget.horizontalHeader().setSectionResizeMode(column, QHeaderView.Stretch)

        # Connect the itemDoubleClicked signal to the custom slot open_file
        self.table_widget.itemDoubleClicked.connect(self.open_file)

        grid_layout.addWidget(self.table_widget, 0, 4, len(options) + 2, 1)  # Modify the row span to match the options

        main_layout.addLayout(grid_layout)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # Load the last used font size
        last_used_font_size = load_last_used_font_size()
        if last_used_font_size:
            font_size_combo.setCurrentText(last_used_font_size)
            self.set_font_size(int(last_used_font_size))

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Load the last used directory
        last_used_directory = load_last_used_directory()
        if last_used_directory:
            self.path_input.setText(last_used_directory)
            self.display_content()

        # Load the last used export directory
        last_used_export_directory = load_last_used_export_directory()
        if last_used_export_directory:
            self.export_input.setText(last_used_export_directory)

        # Add template button
        add_template = QPushButton("Add Template")
        add_template.setToolTip("Adds work order fillable template")
        add_template.clicked.connect(lambda: FileArrangement.copy_paste_fillable_template(last_used_directory))
        self.toolbar.addWidget(add_template)

    def clear_folders(self):
        work_folder = self.path_input.text()
        move_folder = self.export_input.text()

        # Confirmation dialog
        confirmation = QMessageBox.question(
            self, "Confirm Clear Folders",
            "Are you sure you want to delete all files in the Work Folder and Move Folder?",
            QMessageBox.Yes | QMessageBox.No
        )

        if confirmation == QMessageBox.Yes:
            # Delete files and folders in work folder
            FileArrangement.delete_files_and_folders(work_folder)

            # Delete files and folders in move folder
            FileArrangement.delete_files_and_folders(move_folder)

            # Update status bar message
            self.display_content()
            self.status_bar.showMessage("Cleared Work Folder and Move Folder.")

    def select_font_size(self):
        selected_font_size = int(self.sender().currentText())
        self.set_font_size(selected_font_size)
        save_last_used_font_size(selected_font_size)

    def set_font_size(self, size):
        font = QFont()
        font.setPointSize(size)

        # Set font for path_label
        self.path_label.setFont(font)

        # Set font for path_input
        self.path_input.setFont(font)

        # Set font for path_label
        self.export_label.setFont(font)

        # Set font for path_input
        self.export_input.setFont(font)

        # Set font for option_buttons
        for button in self.option_buttons:
            button.setFont(font)

    def select_theme(self):
        selected_theme = self.sender().currentText()
        QApplication.setStyle(QStyleFactory.create(selected_theme))
        save_last_used_theme(selected_theme)

    def open_folder_dialog(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        if dialog.exec():
            selected_folder = dialog.selectedUrls()[0].toLocalFile()
            self.path_input.setText(selected_folder)
            self.display_content()
            save_last_used_directory(selected_folder)

    def clear_folder_path(self):
        self.path_input.clear()

    def display_content(self):
        root_folder_path = self.path_input.text()
        content = FileArrangement.display_content_in_path(root_folder_path).splitlines()

        # Clear existing table content
        self.table_widget.clearContents()
        self.table_widget.setRowCount(len(content))

        # Populate the table with the file names
        for row, filename in enumerate(content):
            item = QTableWidgetItem(filename)
            self.table_widget.setItem(row, 0, item)
        # Auto-resize the first column to fit the content
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)

    def open_export_folder_dialog(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        if dialog.exec():
            selected_folder = dialog.selectedUrls()[0].toLocalFile()
            self.export_input.setText(selected_folder)
            save_last_used_export_directory(selected_folder)

    def move_merged_files(self):
        FileArrangement.move_merged_pdfs(self.path_input.text(), self.export_input.text())
        self.status_bar.showMessage("Moved merge files completed.")
        FileArrangement.compress_pdf_files(self.export_input.text())
        self.status_bar.showMessage("File Compression completed.")
        FileArrangement.remove_pre_and_suf(self.export_input.text())

        # Call get_work_hours function for each merged file
        self.update_work_hours_in_table()

    def update_work_hours_in_table(self):
        move_folder = self.export_input.text()
        work_hours_data = []

        for filename in os.listdir(move_folder):
            if filename.lower().endswith('.pdf'):
                pdf_file_path = os.path.join(move_folder, filename)
                work_hours = get_work_hours(pdf_file_path)

                if work_hours:
                    work_hours_data.append([filename] + work_hours)

        # Update the table with the work hours data
        self.update_table_with_work_hours(work_hours_data)

    def update_table_with_work_hours(self, data):
        # Clear existing table content
        self.table_widget.clearContents()
        self.table_widget.setRowCount(len(data))

        # Populate the table with the work hours data
        for row, row_data in enumerate(data):
            for col, cell_value in enumerate(row_data):
                item = QTableWidgetItem(cell_value)
                self.table_widget.setItem(row, col, item)

        # Resize Time In and Time Out columns to fit the content
        self.table_widget.resizeColumnToContents(1)
        self.table_widget.resizeColumnToContents(2)

    def open_file(self, item):
        # Get the selected row and the first column item (filename)
        row = item.row()
        filename_item = self.table_widget.item(row, 0)

        if filename_item:
            filename = filename_item.text()
            file_path = os.path.join(self.export_input.text(), filename)

            # Check if the file exists and open it
            if os.path.exists(file_path):
                os.startfile(file_path)  # Windows only, for other platforms use different methods

    def handle_option(self):
        option = self.option_buttons.index(self.sender())
        root_folder_path = self.path_input.text()

        if option == 0:
            self.status_bar.showMessage("Extracting work order numbers and renaming PDF files...")
            FileArrangement.debug_print("Option 0 selected: Extract work order numbers and rename PDF files")
            FileArrangement.rename_pdf_files(root_folder_path)
            self.display_content()
            self.status_bar.showMessage(f"Extraction and renaming completed.")
        elif option == 1:
            self.status_bar.showMessage("Creating folders for PDF files...")
            FileArrangement.debug_print("Option 1 selected: Create folders for PDF files")
            created_count = FileArrangement.create_folders_and_move_files(root_folder_path)
            self.status_bar.showMessage(f"Created {created_count} folders.")
            self.status_bar.showMessage("Copying PDF Template to each folder...")
            FileArrangement.copy_pdf_to_subfolders(root_folder_path, "Fillable Work order template.pdf")
            self.display_content()
            FileArrangement.fill_pdf_forms(root_folder_path, "Fillable Work order template.pdf")
            self.status_bar.showMessage("Folder creation completed.")
        elif option == 2:
            FileArrangement.debug_print("Option 2 selected: Merge files by folder")
            self.status_bar.showMessage("Creating Job Image Backups...")
            FileArrangement.create_job_images_folders(root_folder_path)
            self.status_bar.showMessage("Merging PDFs and Image Files...")
            FileArrangement.merge_pdfs_and_images(root_folder_path)
            self.status_bar.showMessage("Deleting converted Image files...")
            deleted_count = FileArrangement.delete_converted_pdfs(root_folder_path)
            FileArrangement.move_files_and_delete_folder(root_folder_path)
            self.status_bar.showMessage(f"Merging completed.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
