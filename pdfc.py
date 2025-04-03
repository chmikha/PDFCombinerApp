import os
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout,
                             QWidget, QFileDialog, QListWidget, QLabel, QMessageBox, QLineEdit)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QIcon, QFont
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter

class PDFCombinerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Combiner by chmikha")
        self.setGeometry(100, 100, 700, 600)  # Slightly larger default size
        self.setWindowIcon(QIcon('./resources/pdf_icon.png'))  # Optional: Add an icon
        self.setAcceptDrops(True)

        # Central widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(15)  # Add more spacing between elements

        # Header section
        self.header_layout = QVBoxLayout()
        self.title_label = QLabel("PDF & Image Combiner")
        self.title_label.setFont(QFont('Segoe UI', 20, QFont.Bold))
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: #333; margin-bottom: 10px;")
        self.header_layout.addWidget(self.title_label)

        self.instructions = QLabel("Drag and drop files or use the 'Add Files' button to combine PDFs and images.")
        self.instructions.setFont(QFont('Segoe UI', 10))
        self.instructions.setAlignment(Qt.AlignCenter)
        self.instructions.setStyleSheet("color: #666; margin-bottom: 20px;")
        self.header_layout.addWidget(self.instructions)
        self.main_layout.addLayout(self.header_layout)

        # File list section
        self.file_list_label = QLabel("Files to Combine:")
        self.file_list_label.setFont(QFont('Segoe UI', 12, QFont.Bold))
        self.file_list_label.setStyleSheet("color: #333;")
        self.main_layout.addWidget(self.file_list_label)

        self.file_list = QListWidget()
        self.file_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 5px;
                background-color: #f7f7f7;
                font: 10pt 'Segoe UI';
            }
            QListWidget::item {
                padding: 8px;
            }
            QListWidget::item:selected {
                background-color: #e0f2f7;
                color: #333;
            }
        """)
        self.file_list.setAlternatingRowColors(True)
        self.main_layout.addWidget(self.file_list)

        # File management buttons
        self.file_buttons_layout = QHBoxLayout()
        self.add_files_btn = QPushButton("Add Files")
        self.remove_btn = QPushButton("Remove Selected")
        self.clear_btn = QPushButton("Clear All")

        button_style = """
            QPushButton {
                background-color: #5cb85c;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font: bold 10pt 'Segoe UI';
            }
            QPushButton:hover {
                background-color: #4cae4c;
            }
            QPushButton:disabled {
                background-color: #d9d9d9;
                color: #999;
            }
        """
        self.add_files_btn.setStyleSheet(button_style)
        self.remove_btn.setStyleSheet(button_style.replace("#5cb85c", "#d9534f").replace("#4cae4c", "#c9302c"))
        self.clear_btn.setStyleSheet(button_style.replace("#5cb85c", "#f0ad4e").replace("#4cae4c", "#eea236"))

        self.add_files_btn.clicked.connect(self.add_files)
        self.remove_btn.clicked.connect(self.remove_selected)
        self.clear_btn.clicked.connect(self.clear_all)

        self.file_buttons_layout.addWidget(self.add_files_btn)
        self.file_buttons_layout.addWidget(self.remove_btn)
        self.file_buttons_layout.addWidget(self.clear_btn)
        self.main_layout.addLayout(self.file_buttons_layout)

        # Output settings section
        self.output_group_layout = QVBoxLayout()
        self.output_label = QLabel("Output Settings:")
        self.output_label.setFont(QFont('Segoe UI', 12, QFont.Bold))
        self.output_label.setStyleSheet("color: #333; margin-top: 15px;")
        self.output_group_layout.addWidget(self.output_label)

        self.filename_layout = QHBoxLayout()
        self.filename_label = QLabel("Filename:")
        self.filename_label.setFont(QFont('Segoe UI', 10))
        self.filename_label.setStyleSheet("color: #333; font-weight: bold;")
        self.filename_layout.addWidget(self.filename_label)
        self.filename_input = QLineEdit("combined.pdf")
        self.filename_input.setFont(QFont('Segoe UI', 10))
        self.filename_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 8px;
                font: 10pt 'Segoe UI';
            }
        """)
        self.filename_layout.addWidget(self.filename_input, 1)
        self.output_group_layout.addLayout(self.filename_layout)

        self.output_path_layout = QHBoxLayout()
        self.output_path_label = QLabel("Output Folder:")
        self.output_path_label.setFont(QFont('Segoe UI', 10))
        self.output_path_label.setStyleSheet("color: #333; font-weight: bold;")
        self.output_path_layout.addWidget(self.output_path_label)
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Default: Application Directory")
        self.output_path.setFont(QFont('Segoe UI', 10))
        self.output_path.setStyleSheet("""
            QLineEdit {
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 8px;
                font: 10pt 'Segoe UI';
            }
        """)
        self.output_path_layout.addWidget(self.output_path, 1)
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.setStyleSheet(button_style.replace("#5cb85c", "#007bff").replace("#4cae4c", "#0056b3"))
        self.browse_btn.setFont(QFont('Segoe UI', 10, QFont.Bold))
        self.browse_btn.clicked.connect(self.browse_output)
        self.output_path_layout.addWidget(self.browse_btn)
        self.output_group_layout.addLayout(self.output_path_layout)

        self.main_layout.addLayout(self.output_group_layout)

        # Combine button
        self.combine_btn = QPushButton("Combine Files to PDF")
        self.combine_btn.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 12px 25px;
                border-radius: 5px;
                font: bold 12pt 'Segoe UI';
                margin-top: 25px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:disabled {
                background-color: #d9d9d9;
                color: #999;
            }
        """)
        self.combine_btn.clicked.connect(self.combine_files)
        self.main_layout.addWidget(self.combine_btn)

        # Status bar
        self.statusBar().setFont(QFont('Segoe UI', 9))
        self.statusBar().setStyleSheet("color: #555;")
        self.statusBar().showMessage("Ready")

        # Initial state of buttons
        self.update_button_states()
        self.file_list.itemSelectionChanged.connect(self.update_button_states)
        self.file_list.model().rowsInserted.connect(self.update_button_states)
        self.file_list.model().rowsRemoved.connect(self.update_button_states)

    def update_button_states(self):
        has_items = self.file_list.count() > 0
        has_selection = len(self.file_list.selectedItems()) > 0
        self.remove_btn.setEnabled(has_selection)
        self.clear_btn.setEnabled(has_items)
        self.combine_btn.setEnabled(has_items)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        valid_files = []

        for url in urls:
            file_path = url.toLocalFile()
            if os.path.isfile(file_path):
                ext = os.path.splitext(file_path)[1].lower()
                if ext in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif']:
                    valid_files.append(file_path)

        if valid_files:
            self.add_files_to_list(valid_files)
        else:
            QMessageBox.warning(self, "Invalid Files", "No valid PDF or image files were dropped.")

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF and Image Files", "",
            "Supported Files (*.pdf *.jpg *.jpeg *.png *.bmp *.tiff *.gif);;All Files (*)"
        )

        if files:
            self.add_files_to_list(files)

    def add_files_to_list(self, files):
        for file_path in files:
            self.file_list.addItem(file_path)
        self.statusBar().showMessage(f"Added {len(files)} file(s)")
        self.update_button_states()

    def remove_selected(self):
        selected_items = self.file_list.selectedItems()
        for item in selected_items:
            self.file_list.takeItem(self.file_list.row(item))
        self.statusBar().showMessage(f"Removed {len(selected_items)} file(s)")
        self.update_button_states()

    def clear_all(self):
        self.file_list.clear()
        self.statusBar().showMessage("Cleared all files")
        self.update_button_states()

    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_path.setText(folder)

    def combine_files(self):
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "No Files", "Please add files to combine.")
            return

        filename = self.filename_input.text().strip()
        if not filename:
            QMessageBox.warning(self, "Missing Filename", "Please enter a filename for the output PDF.")
            return
        if not filename.lower().endswith('.pdf'):
            filename += '.pdf'

        output_dir = self.output_path.text().strip()
        if output_dir and not os.path.isdir(output_dir):
            QMessageBox.critical(self, "Invalid Output Folder", "The specified output folder is not valid.")
            return

        if output_dir:
            output_path = os.path.join(output_dir, filename)
        else:
            app_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
            output_path = os.path.join(app_dir, filename)

        if os.path.exists(output_path):
            reply = QMessageBox.question(
                self, "File Exists",
                f"The file '{os.path.basename(output_path)}' already exists. Overwrite?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        try:
            self.statusBar().showMessage("Combining files...", 0)
            QApplication.processEvents()

            pdf_writer = PdfWriter()

            for i in range(self.file_list.count()):
                file_path = self.file_list.item(i).text()
                ext = os.path.splitext(file_path)[1].lower()

                if ext == '.pdf':
                    try:
                        with open(file_path, 'rb') as pdf_file:
                            pdf_reader = PdfReader(pdf_file)
                            for page_num in range(len(pdf_reader.pages)):
                                page = pdf_reader.pages[page_num]
                                pdf_writer.add_page(page)
                    except Exception as pdf_error:
                        QMessageBox.warning(self, "PDF Error", f"Could not process PDF file:\n{file_path}\n\nError: {pdf_error}")
                        continue
                else:
                    try:
                        temp_dir = os.path.dirname(output_path) or os.getcwd()
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{i}.pdf")
                        self.convert_image_to_pdf(file_path, temp_pdf_path)
                        with open(temp_pdf_path, 'rb') as temp_pdf_file:
                            pdf_reader = PdfReader(temp_pdf_file)
                            for page_num in range(len(pdf_reader.pages)):
                                page = pdf_reader.pages[page_num]
                                pdf_writer.add_page(page)
                        os.remove(temp_pdf_path)
                    except Exception as img_error:
                        QMessageBox.warning(self, "Image Error", f"Could not process image file:\n{file_path}\n\nError: {img_error}")
                        continue

            with open(output_path, 'wb') as out_file:
                pdf_writer.write(out_file)

            self.statusBar().showMessage(f"Successfully created: {output_path}", 0)
            QMessageBox.information(self, "Success", f"PDF successfully created at:\n{output_path}")

        except Exception as e:
            self.statusBar().showMessage("Error occurred", 0)
            QMessageBox.critical(self, "Error", f"An error occurred while creating the PDF:\n{str(e)}")

    def convert_image_to_pdf(self, image_path, output_pdf_path):
        """Convert an image file to a PDF page"""
        try:
            img = Image.open(image_path).convert("RGB")
            img_reader = ImageReader(img)
            width, height = img.size
            c = canvas.Canvas(output_pdf_path, pagesize=letter)
            pdf_width, pdf_height = letter
            aspect = width / float(height)
            max_width = pdf_width - 100
            max_height = pdf_height - 100

            if aspect > max_width / max_height:
                draw_width = max_width
                draw_height = draw_width / aspect
            else:
                draw_height = max_height
                draw_width = draw_height * aspect

            x = (pdf_width - draw_width) / 2
            y = (pdf_height - draw_height) / 2
            c.drawImage(img_reader, x, y, draw_width, draw_height)
            c.save()
        except Exception as e:
            print(f"Error converting image {image_path} to PDF: {e}")
            raise

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = PDFCombinerApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()