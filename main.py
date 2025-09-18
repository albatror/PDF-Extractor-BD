import sys
import os
import json
import re
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, QPushButton, 
                             QVBoxLayout, QWidget, QLabel, QTextEdit, QTableWidget,
                             QTableWidgetItem, QTabWidget, QScrollArea, QMessageBox,
                             QProgressBar, QGroupBox, QHBoxLayout, QCheckBox)
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QPixmap, QImage, QPainter, QColor, QPen
import PyPDF2
import pdfplumber
import pytesseract
import cv2
import numpy as np
import pandas as pd
from collections import defaultdict

# Local imports
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

class PDFExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Travel Expense Extractor")
        self.setGeometry(100, 100, 1200, 800)
        
        # Data storage
        self.pdf_files = []
        self.form_data = []
        self.zone_definitions = {}
        self.extracted_data = []
        
        # UI setup
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # File selection
        file_group = QGroupBox("Fichiers PDF")
        file_layout = QHBoxLayout()
        self.file_button = QPushButton("Sélectionner les PDF")
        self.file_button.clicked.connect(self.load_files)
        file_layout.addWidget(self.file_button)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # Tabs for different views
        self.tabs = QTabWidget()
        
        # Preview tab
        self.preview_tab = QWidget()
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel("Aperçu PDF")
        self.preview_label.setAlignment(Qt.AlignCenter)
        preview_layout.addWidget(self.preview_label)
        self.preview_tab.setLayout(preview_layout)
        self.tabs.addTab(self.preview_tab, "Aperçu")
        
        # Data tab
        self.data_tab = QWidget()
        data_layout = QVBoxLayout()
        self.data_table = QTableWidget(0, 9)
        self.data_table.setHorizontalHeaderLabels(["NOM", "PRENOM", "MONTANT 1", "MONTANT 2", "MONTANT 3", "TOTAL", "PAGES", "OBSERVATIONS", "ACTION"])
        data_layout.addWidget(self.data_table)
        self.data_tab.setLayout(data_layout)
        self.tabs.addTab(self.data_tab, "Données")
        
        layout.addWidget(self.tabs)
        
        # Action buttons
        button_group = QGroupBox("Actions")
        button_layout = QHBoxLayout()
        self.extract_button = QPushButton("Extraire les données")
        self.extract_button.clicked.connect(self.extract_data)
        self.export_button = QPushButton("Exporter en Excel")
        self.export_button.clicked.connect(self.export_excel)
        self.export_button.setEnabled(False)
        button_layout.addWidget(self.extract_button)
        button_layout.addWidget(self.export_button)
        button_group.setLayout(button_layout)
        layout.addWidget(button_group)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
    def load_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Sélectionner les fichiers PDF", "", "PDF Files (*.pdf)"
        )
        if files:
            self.pdf_files = files
            self.statusBar().showMessage(f"{len(files)} fichiers sélectionnés")
            self.extract_button.setEnabled(True)
            
    def extract_data(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner des fichiers PDF.")
            return
            
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Extraction en cours...")
        
        try:
            # Process each PDF
            for i, pdf_path in enumerate(self.pdf_files):
                self.progress_bar.setValue(int((i + 1) / len(self.pdf_files) * 100))
                self.process_pdf(pdf_path)
                
            self.update_data_table()
            self.export_button.setEnabled(True)
            self.statusBar().showMessage("Extraction terminée")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'extraction: {str(e)}")
            
    def process_pdf(self, pdf_path):
        try:
            # Detect form type
            form_type = self.detect_form_type(pdf_path)
            
            # Extract text and images
            pages_data = self.extract_pages_data(pdf_path, form_type)
            
            # Process each agent in the PDF
            agents = self.group_agents(pages_data)
            
            for agent_name, agent_data in agents.items():
                # Extract data for this agent
                extracted_agent_data = self.extract_agent_data(agent_data, form_type)
                if extracted_agent_data:
                    self.extracted_data.append(extracted_agent_data)
                    
        except Exception as e:
            print(f"Erreur avec le PDF {pdf_path}: {str(e)}")
            
    def detect_form_type(self, pdf_path):
        """Detect form type based on page structure"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) >= 4:
                    # Check first page for A4 format indicators
                    first_page = pdf.pages[0]
                    text = first_page.extract_text()
                    
                    # Look for specific patterns that indicate form type
                    if "VERIFIE LE" in text and "PAIE" in text:
                        return "A4_4pages"
                    elif "VERIFIE LE" in text and "PAIE" in text:
                        return "A4_4pages"
                elif len(pdf.pages) >= 2:
                    # Check first page for A3 format indicators
                    first_page = pdf.pages[0]
                    text = first_page.extract_text()
                    
                    if "VERIFIE LE" in text and "PAIE" in text:
                        return "A3_2pages"
                        
        except Exception as e:
            print(f"Erreur detection type: {str(e)}")
            
        # Default to A4 4 pages
        return "A4_4pages"
        
    def extract_pages_data(self, pdf_path, form_type):
        """Extract data from all pages of a PDF"""
        pages_data = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_data = {
                        'page_num': i + 1,
                        'text': page.extract_text(),
                        'image': None
                    }
                    
                    # Try to get image if available
                    try:
                        img = page.to_image()
                        page_data['image'] = img.original
                    except:
                        pass
                        
                    pages_data.append(page_data)
                    
        except Exception as e:
            print(f"Erreur extraction pages: {str(e)}")
            
        return pages_data
        
    def group_agents(self, pages_data):
        """Group pages by agent based on page structure"""
        agents = defaultdict(list)
        
        # Simple grouping logic - in a real implementation this would be more sophisticated
        for page_data in pages_data:
            # Look for agent identifiers (this is simplified)
            text = page_data['text']
            
            # Try to find name patterns
            name_match = re.search(r'(?:NOM|NAME)\s*:\s*([A-Z\s]+)', text, re.IGNORECASE)
            if name_match:
                agent_name = name_match.group(1).strip()
                agents[agent_name].append(page_data)
            else:
                # Default to single agent
                agents['Default Agent'].append(page_data)
                
        return agents
        
    def extract_agent_data(self, agent_pages, form_type):
        """Extract data for a specific agent"""
        try:
            # This is a simplified implementation - in reality this would use OCR and zone definitions
            agent_data = {
                'name': '',
                'first_name': '',
                'amounts': [],
                'pages': [],
                'observations': ''
            }
            
            # Extract name from first page
            if agent_pages:
                first_page_text = agent_pages[0]['text']
                
                # Extract name
                name_match = re.search(r'(?:NOM|NAME)\s*:\s*([A-Z\s]+)', first_page_text, re.IGNORECASE)
                if name_match:
                    agent_data['name'] = name_match.group(1).strip()
                    
                # Extract first name  
                first_name_match = re.search(r'(?:PRENOM|FIRST NAME)\s*:\s*([A-Z\s]+)', first_page_text, re.IGNORECASE)
                if first_name_match:
                    agent_data['first_name'] = first_name_match.group(1).strip()
                    
                # Extract amounts
                amount_matches = re.findall(r'(?:TOTAL A PAYER|TOTAL TO PAY)\s*=\s*(\d+(?:[.,]\d+)?)', first_page_text)
                for match in amount_matches:
                    agent_data['amounts'].append(match)
                    
                # Extract page numbers
                agent_data['pages'] = [str(page['page_num']) for page in agent_pages]
                
            return agent_data
            
        except Exception as e:
            print(f"Erreur extraction agent: {str(e)}")
            return None
            
    def update_data_table(self):
        """Update the data table with extracted information"""
        self.data_table.setRowCount(0)
        
        if not self.extracted_data:
            return
            
        for i, agent_data in enumerate(self.extracted_data):
            row_position = self.data_table.rowCount()
            self.data_table.insertRow(row_position)
            
            # Add data to table
            self.data_table.setItem(row_position, 0, QTableWidgetItem(agent_data['name']))
            self.data_table.setItem(row_position, 1, QTableWidgetItem(agent_data['first_name']))
            
            # Add amounts
            for j, amount in enumerate(agent_data['amounts']):
                if j < 6:  # Only show first 6 amounts
                    self.data_table.setItem(row_position, j + 2, QTableWidgetItem(amount))
                    
            # Add total
            total = sum(float(a.replace(',', '.')) for a in agent_data['amounts'] if a)
            self.data_table.setItem(row_position, 8, QTableWidgetItem(f"{total:.2f}"))
            
            # Add pages
            self.data_table.setItem(row_position, 7, QTableWidgetItem('/'.join(agent_data['pages'])))
            
    def export_excel(self):
        """Export data to Excel file"""
        try:
            if not self.extracted_data:
                QMessageBox.warning(self, "Erreur", "Aucune donnée à exporter.")
                return
                
            # Create DataFrame
            df_data = []
            for agent in self.extracted_data:
                row = {
                    'NOM': agent['name'],
                    'PRENOM': agent['first_name'],
                    'PAGES': '/'.join(agent['pages']),
                    'OBSERVATIONS': agent['observations']
                }
                
                # Add amounts
                for i, amount in enumerate(agent['amounts']):
                    row[f'MONTANT {i+1}'] = amount
                    
                # Calculate total
                total = sum(float(a.replace(',', '.')) for a in agent['amounts'] if a)
                row['TOTAL'] = f"{total:.2f}"
                
                df_data.append(row)
                
            df = pd.DataFrame(df_data)
            
            # Save to Excel
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Exporter en Excel", "", "Excel Files (*.xlsx)"
            )
            
            if file_path:
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Succès", f"Fichier exporté avec succès: {file_path}")
                
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'exportation: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = PDFExtractorApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
