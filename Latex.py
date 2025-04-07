import sys
import re
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QTextBrowser, QLabel, QFileDialog
)
from PyQt6.QtGui import QPixmap
from pdf2image import convert_from_path
from pylatex import Document, NoEscape


class ReportSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Select Report Components")
        self.setGeometry(100, 100, 600, 500)

        # Layout
        layout = QVBoxLayout()

        # List Widget for components
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.MultiSelection)

        # Load LaTeX file
        self.tex_file = r"C:\Users\rktej\Desktop\Latex\FOSSEE_SUMMER_FELLOWSHIP_SAMPLE_TEX.tex"  # Change as needed
        self.components = self.extract_components(self.tex_file)

        # Add components with generic names (Component 1, Component 2, etc.)
        for i in range(len(self.components)):
            self.list_widget.addItem(f"Component {i+1}")

        # Button to generate report
        self.generate_btn = QPushButton("Generate Report")
        self.generate_btn.clicked.connect(self.generate_report)

        # Preview Text Browser
        self.preview = QTextBrowser()

        # PDF Preview Label
        self.pdf_preview = QLabel()

        # Add Widgets to Layout
        layout.addWidget(self.list_widget)
        layout.addWidget(self.preview)
        layout.addWidget(self.generate_btn)
        layout.addWidget(self.pdf_preview)

        self.setLayout(layout)

        # Store previous selections
        self.selected_text_list = []

    def extract_components(self, file_path):
        """Extract sections, tables, and figures from a LaTeX file."""
        components = []
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()

            # Extract Sections, Tables, and Figures
            matches = re.findall(r'(\\section{.*?}|\\begin{table}.*?\\end{table}|\\begin{figure}.*?\\end{figure})', content, re.DOTALL)
            components.extend(matches)

        return components

    def generate_report(self):
        """Generate a LaTeX report using PyLaTeX with selected components without replacing previous selections."""
        selected_items = self.list_widget.selectedItems()
        selected_indices = [int(item.text().split()[-1]) - 1 for item in selected_items]
        selected_text = "\n".join([self.components[i] for i in selected_indices])

        if not selected_text:
            self.preview.setText("No components selected!")
            return

        # Append newly selected components
        self.selected_text_list.append(selected_text)
        final_text = "\n".join(self.selected_text_list)

        self.preview.setText(final_text)

        # Ask user where to save the report
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Report", "custom_report.pdf", "PDF Files (*.pdf)")
        if save_path:
            save_path = self.get_unique_filename(save_path)  # Ensure unique file names
            self.compile_tex_to_pdf(save_path, final_text)
            self.show_pdf_preview(save_path.replace(".pdf", ".pdf"))  # Ensure correct path

    def compile_tex_to_pdf(self, save_path, selected_text):
        """Use PyLaTeX to compile selected components into a PDF."""
        doc = Document()  # PyLaTeX automatically includes \documentclass and \begin{document}
        doc.append(NoEscape(selected_text))  # Append only selected content
        doc.generate_pdf(save_path.replace(".pdf", ""), clean_tex=False, compiler="pdflatex")


    def show_pdf_preview(self, pdf_path):
        """Convert first page of the generated PDF to an image and display."""
        images = convert_from_path(pdf_path)
        images[0].save("preview.png", "PNG")
        pixmap = QPixmap("preview.png")
        self.pdf_preview.setPixmap(pixmap)

    def get_unique_filename(self, file_path):
        """Generate a unique filename if a file with the same name already exists."""
        base, ext = os.path.splitext(file_path)
        counter = 1
        new_path = file_path

        while os.path.exists(new_path):
            new_path = f"{base}_{counter}{ext}"
            counter += 1

        return new_path


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReportSelector()
    window.show()
    sys.exit(app.exec())
