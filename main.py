import os
import sys

from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

import pandas as pd
import numpy as np

import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from datetime import datetime

class PolymerSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Polymer Selection Expert")
        self.setGeometry(100, 100, 1400, 800)

        # Load data
        self.load_data()

        # Initialize storage
        self.filtered_df = self.df.copy()
        self.shortlist_df = pd.DataFrame()
        self.scores = {}

        # Setup UI
        self.setup_ui()
        self.create_menu_bar()

    def load_data(self):
        """Load polymer database from Excel"""
        try:
            # Create sample data if file doesn't exist
            self.df = pd.read_excel('polymer_data.xlsx')
        except:
            self.create_sample_data()

    def create_sample_data(self):
        """Generate sample polymer data for testing"""
        data = {
            'Polymer': ['PA66 GF30', 'PEEK', 'PC', 'POM', 'PPS GF40', 'ABS', 'PET', 'PBT',
                       'PEI', 'PPSU', 'PVDF', 'PTFE', 'PA6', 'PMMA', 'PP', 'HDPE'],
            'Family': ['Polyamide', 'PAEK', 'Polycarbonate', 'Polyacetal', 'PPS', 'Styrenic',
                      'Polyester', 'Polyester', 'Polyimide', 'Polysulfone', 'Fluoropolymer',
                      'Fluoropolymer', 'Polyamide', 'Acrylic', 'Polyolefin', 'Polyolefin'],
            'Tensile_MPa': [180, 110, 70, 65, 160, 45, 80, 60, 110, 75, 50, 25, 75, 70, 35, 30],
            'Flexural_GPa': [8.5, 4.0, 2.4, 2.8, 14.0, 2.2, 3.5, 2.5, 3.3, 2.4, 1.8, 0.6, 2.6, 3.0, 1.5, 1.2],
            'Impact_kJm2': [12, 8, 70, 8, 9, 25, 5, 6, 6, 12, 14, 4, 10, 2, 6, 8],
            'HDT_C': [250, 315, 135, 110, 260, 95, 75, 65, 210, 205, 140, 120, 65, 95, 100, 70],
            'MaxTemp_C': [150, 260, 120, 100, 230, 85, 70, 60, 180, 180, 150, 260, 80, 85, 90, 60],
            'ThermalCond_WmK': [0.24, 0.25, 0.21, 0.31, 0.30, 0.17, 0.29, 0.27, 0.22, 0.21, 0.19, 0.25, 0.23, 0.19, 0.22, 0.45],
            'Density_gcm3': [1.37, 1.32, 1.20, 1.41, 1.66, 1.05, 1.38, 1.31, 1.27, 1.29, 1.78, 2.20, 1.13, 1.19, 0.91, 0.96],
            'WaterAbs_pct': [1.2, 0.1, 0.2, 0.2, 0.03, 0.3, 0.2, 0.08, 0.25, 0.3, 0.04, 0.01, 1.5, 0.3, 0.1, 0.01],
            'Cost_Index': [3, 5, 2, 3, 4, 1, 2, 2, 4, 4, 4, 4, 2, 2, 1, 1],
            'UL94': ['HB', 'V-0', 'V-2', 'HB', 'V-0', 'HB', 'HB', 'HB', 'V-0', 'V-0', 'V-0', 'V-0', 'HB', 'HB', 'HB', 'HB']
        }
        self.df = pd.DataFrame(data)
        # Save for future use
        self.df.to_excel('polymer_data.xlsx', index=False)

    def setup_ui(self):
        """Create main UI with tabs"""
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Create tabs
        self.project_tab = QWidget()
        self.filter_tab = QWidget()
        self.ranking_tab = QWidget()
        self.compare_tab = QWidget()
        self.detail_tab = QWidget()
        self.report_tab = QWidget()

        self.tabs.addTab(self.detail_tab, "0. Polymer Details")
        self.tabs.addTab(self.project_tab, "1. Project Definition")
        self.tabs.addTab(self.filter_tab, "2. Constraint Filtering")
        self.tabs.addTab(self.ranking_tab, "3. Scoring & Ranking")
        self.tabs.addTab(self.compare_tab, "4. Comparison")
        self.tabs.addTab(self.report_tab, "5. Report")

        self.setup_project_tab()
        self.setup_filter_tab()
        self.setup_ranking_tab()
        self.setup_compare_tab()
        self.setup_detail_tab()
        self.setup_report_tab()

    def create_menu_bar(self):
        """Create application menu"""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu('File')
        load_action = QAction('Load Dataset', self)
        load_action.triggered.connect(self.load_dataset)
        file_menu.addAction(load_action)

        export_action = QAction('Export Results', self)
        export_action.triggered.connect(self.export_results)
        file_menu.addAction(export_action)

        file_menu.addSeparator()
        exit_action = QAction('Exit', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Help menu
        help_menu = menubar.addMenu('Help')
        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def setup_project_tab(self):
        """Project definition page"""
        layout = QVBoxLayout()

        # Title
        title = QLabel("Project Definition")
        title.setStyleSheet("font-size: 18px; font-weight: bold; margin: 10px;")
        layout.addWidget(title)

        # Form layout
        form_widget = QWidget()
        form_layout = QFormLayout()

        self.project_name = QLineEdit()
        self.project_name.setPlaceholderText("e.g., Tower Packing (Internals)")
        form_layout.addRow("Project Name:", self.project_name)

        self.application = QComboBox()
        self.application.addItems(["Structural", "Thermal Management", "Electrical Insulation",
                                   "Chemical Resistance", "Optical", "Wear Resistance"])
        form_layout.addRow("Primary Application:", self.application)

        self.environment = QComboBox()
        self.environment.addItems(["Indoor", "Outdoor", "High Temperature", "Chemical Exposure",
                                   "Underwater", "Medical"])
        form_layout.addRow("Operating Environment:", self.environment)

        self.process = QComboBox()
        self.process.addItems(["Injection Molding", "Extrusion", "3D Printing",
                              "Compression Molding", "Thermoforming"])
        form_layout.addRow("Manufacturing Process:", self.process)

        self.notes = QTextEdit()
        self.notes.setMaximumHeight(100)
        form_layout.addRow("Additional Notes:", self.notes)

        form_widget.setLayout(form_layout)
        layout.addWidget(form_widget)

        # Save button
        save_btn = QPushButton("Save Project Details")
        save_btn.clicked.connect(self.save_project)
        layout.addWidget(save_btn)

        layout.addStretch()
        self.project_tab.setLayout(layout)

    def setup_filter_tab(self):
        """Constraint filtering page"""
        layout = QVBoxLayout()

        # Filter group
        filter_group = QGroupBox("Material Constraints")
        filter_layout = QGridLayout()

        # Tensile strength
        filter_layout.addWidget(QLabel("Min Tensile Strength (MPa):"), 0, 0)
        self.min_strength = QSpinBox()
        self.min_strength.setRange(0, 500)
        self.min_strength.setValue(0)
        filter_layout.addWidget(self.min_strength, 0, 1)

        # Flexural modulus
        filter_layout.addWidget(QLabel("Min Flexural Modulus (GPa):"), 0, 2)
        self.min_flexural = QDoubleSpinBox()
        self.min_flexural.setRange(0, 50)
        self.min_flexural.setValue(0)
        filter_layout.addWidget(self.min_flexural, 0, 3)

        # Impact strength
        filter_layout.addWidget(QLabel("Min Impact Strength (kJ/m²):"), 1, 0)
        self.min_impact = QSpinBox()
        self.min_impact.setRange(0, 200)
        self.min_impact.setValue(0)
        filter_layout.addWidget(self.min_impact, 1, 1)

        # HDT
        filter_layout.addWidget(QLabel("Min HDT (°C):"), 1, 2)
        self.min_hdt = QSpinBox()
        self.min_hdt.setRange(0, 400)
        self.min_hdt.setValue(0)
        filter_layout.addWidget(self.min_hdt, 1, 3)

        # Max temperature
        filter_layout.addWidget(QLabel("Min Operating Temp (°C):"), 2, 0)
        self.min_temp = QSpinBox()
        self.min_temp.setRange(0, 400)
        self.min_temp.setValue(0)
        filter_layout.addWidget(self.min_temp, 2, 1)

        # Max water absorption
        filter_layout.addWidget(QLabel("Max Water Absorption (%):"), 2, 2)
        self.max_water = QDoubleSpinBox()
        self.max_water.setRange(0, 10)
        self.max_water.setValue(5.0)
        filter_layout.addWidget(self.max_water, 2, 3)

        # Max density
        filter_layout.addWidget(QLabel("Max Density (g/cm³):"), 3, 0)
        self.max_density = QDoubleSpinBox()
        self.max_density.setRange(0, 5)
        self.max_density.setValue(5)
        filter_layout.addWidget(self.max_density, 3, 1)

        # Max cost
        filter_layout.addWidget(QLabel("Max Cost Index:"), 3, 2)
        self.max_cost = QSpinBox()
        self.max_cost.setRange(1, 5)
        self.max_cost.setValue(5)
        filter_layout.addWidget(self.max_cost, 3, 3)

        # UL94 rating
        filter_layout.addWidget(QLabel("UL94 Rating:"), 4, 0)
        self.ul94 = QComboBox()
        self.ul94.addItems(["Any", "HB", "V-2", "V-1", "V-0", "5VA", "5VB"])
        filter_layout.addWidget(self.ul94, 4, 1)

        # Polymer family multiselect
        filter_layout.addWidget(QLabel("Polymer Families:"), 4, 2)
        self.family_list = QListWidget()
        self.family_list.setSelectionMode(QAbstractItemView.MultiSelection)
        families = sorted(self.df['Family'].unique())
        for family in families:
            self.family_list.addItem(family)
        filter_layout.addWidget(self.family_list, 5, 2, 1, 2)

        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)

        # Apply button
        apply_btn = QPushButton("Apply Filters")
        apply_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px;")
        apply_btn.clicked.connect(self.apply_filters)
        layout.addWidget(apply_btn)

        # Results display
        self.filter_results = QTextEdit()
        self.filter_results.setReadOnly(True)
        layout.addWidget(QLabel("Filter Results:"))
        layout.addWidget(self.filter_results)

        self.filter_tab.setLayout(layout)

    def setup_ranking_tab(self):
        """Scoring and ranking page"""
        layout = QVBoxLayout()

        # Weighting section
        weight_group = QGroupBox("Property Weights (Total = 100%)")
        weight_layout = QGridLayout()

        self.weights = {}
        properties = ['Tensile Strength', 'Flexural Modulus', 'Impact Strength',
                     'Heat Resistance', 'Cost']

        row = 0
        for prop in properties:
            weight_layout.addWidget(QLabel(f"{prop} (%):"), row, 0)
            spin = QSpinBox()
            spin.setRange(0, 100)
            spin.setValue(20)
            spin.valueChanged.connect(self.update_weight_total)
            weight_layout.addWidget(spin, row, 1)
            self.weights[prop] = spin
            row += 1

        self.weight_total_label = QLabel("Total: 0%")
        weight_layout.addWidget(self.weight_total_label, row, 0, 1, 2)

        weight_group.setLayout(weight_layout)
        layout.addWidget(weight_group)

        # Calculate button
        calc_btn = QPushButton("Calculate Rankings")
        calc_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 10px;")
        calc_btn.clicked.connect(self.calculate_ranking)
        layout.addWidget(calc_btn)

        # Results table
        layout.addWidget(QLabel("Ranked Polymers:"))
        self.ranking_table = QTableWidget()
        self.ranking_table.setAlternatingRowColors(True)
        layout.addWidget(self.ranking_table)

        # Chart
        self.ranking_figure = Figure(figsize=(8, 4))
        self.ranking_canvas = FigureCanvas(self.ranking_figure)
        layout.addWidget(self.ranking_canvas)

        self.ranking_tab.setLayout(layout)

    def setup_compare_tab(self):
        """Comparison table page"""
        layout = QVBoxLayout()

        # Polymer selection for comparison
        select_group = QGroupBox("Select Polymers to Compare")
        select_layout = QHBoxLayout()

        self.compare_list = QListWidget()
        self.compare_list.setSelectionMode(QAbstractItemView.MultiSelection)
        select_layout.addWidget(self.compare_list)

        btn_layout = QVBoxLayout()
        add_btn = QPushButton("Add to Shortlist →")
        add_btn.clicked.connect(self.add_to_shortlist)
        remove_btn = QPushButton("← Remove")
        remove_btn.clicked.connect(self.remove_from_shortlist)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        select_layout.addLayout(btn_layout)

        self.shortlist_widget = QListWidget()
        select_layout.addWidget(self.shortlist_widget)

        select_group.setLayout(select_layout)
        layout.addWidget(select_group)

        # Comparison table
        layout.addWidget(QLabel("Property Comparison:"))
        self.comparison_table = QTableWidget()
        layout.addWidget(self.comparison_table)

        # Radar chart
        self.radar_figure = Figure(figsize=(6, 6))
        self.radar_canvas = FigureCanvas(self.radar_figure)
        layout.addWidget(self.radar_canvas)

        self.compare_tab.setLayout(layout)

    def setup_detail_tab(self):
        """Polymer detail page"""
        layout = QVBoxLayout()

        # Polymer selector
        select_layout = QHBoxLayout()
        select_layout.addWidget(QLabel("Select Polymer:"))
        self.detail_combo = QComboBox()
        self.detail_combo.addItems(sorted(self.df['Polymer']))
        self.detail_combo.currentTextChanged.connect(self.show_polymer_details)
        select_layout.addWidget(self.detail_combo)
        select_layout.addStretch()
        layout.addLayout(select_layout)

        # Detail display
        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        layout.addWidget(self.detail_text)

        # Property chart
        self.detail_figure = Figure(figsize=(8, 4))
        self.detail_canvas = FigureCanvas(self.detail_figure)
        layout.addWidget(self.detail_canvas)

        self.detail_tab.setLayout(layout)

    def setup_report_tab(self):
        """Report generation page"""
        layout = QVBoxLayout()

        # Report options
        options_group = QGroupBox("Report Options")
        options_layout = QGridLayout()

        options_layout.addWidget(QLabel("Include in Report:"), 0, 0)
        self.include_project = QCheckBox("Project Details")
        self.include_project.setChecked(True)
        options_layout.addWidget(self.include_project, 0, 1)

        self.include_filters = QCheckBox("Filter Criteria")
        self.include_filters.setChecked(True)
        options_layout.addWidget(self.include_filters, 1, 1)

        self.include_ranking = QCheckBox("Ranking Results")
        self.include_ranking.setChecked(True)
        options_layout.addWidget(self.include_ranking, 2, 1)

        self.include_shortlist = QCheckBox("Shortlist Comparison")
        self.include_shortlist.setChecked(True)
        options_layout.addWidget(self.include_shortlist, 3, 1)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Generate button
        generate_btn = QPushButton("Generate Report (CSV)")
        generate_btn.setStyleSheet("background-color: #FF9800; color: white; padding: 10px;")
        generate_btn.clicked.connect(self.generate_report)
        layout.addWidget(generate_btn)

        # Preview
        layout.addWidget(QLabel("Report Preview:"))
        self.report_preview = QTextEdit()
        self.report_preview.setReadOnly(True)
        layout.addWidget(self.report_preview)

        self.report_tab.setLayout(layout)

    def save_project(self):
        """Save project details"""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Project details saved successfully!")
        msg.setWindowTitle("Success")
        msg.exec_()

    def apply_filters(self):
        """Apply constraints to filter polymers"""
        try:
            filtered = self.df.copy()

            # Apply numeric filters
            filtered = filtered[filtered['Tensile_MPa'] >= self.min_strength.value()]
            filtered = filtered[filtered['Flexural_GPa'] >= self.min_flexural.value()]
            filtered = filtered[filtered['Impact_kJm2'] >= self.min_impact.value()]
            filtered = filtered[filtered['HDT_C'] >= self.min_hdt.value()]
            filtered = filtered[filtered['MaxTemp_C'] >= self.min_temp.value()]
            filtered = filtered[filtered['WaterAbs_pct'] <= self.max_water.value()]
            filtered = filtered[filtered['Density_gcm3'] <= self.max_density.value()]
            filtered = filtered[filtered['Cost_Index'] <= self.max_cost.value()]

            # Apply UL94 filter
            ul94_value = self.ul94.currentText()
            if ul94_value != "Any":
                filtered = filtered[filtered['UL94'] == ul94_value]

            # Apply family filter
            selected_families = [item.text() for item in self.family_list.selectedItems()]
            if selected_families:
                filtered = filtered[filtered['Family'].isin(selected_families)]

            self.filtered_df = filtered

            # Update compare list
            self.compare_list.clear()
            for polymer in sorted(self.filtered_df['Polymer']):
                self.compare_list.addItem(polymer)

            # Show results
            result_text = f"✅ Filter applied successfully!\n\n"
            result_text += f"Original polymers: {len(self.df)}\n"
            result_text += f"Polymers after filtering: {len(self.filtered_df)}\n\n"

            if len(self.filtered_df) > 0:
                result_text += "Top 10 polymers by tensile strength:\n"
                top10 = self.filtered_df.nlargest(10, 'Tensile_MPa')[['Polymer', 'Tensile_MPa', 'Family']]
                result_text += top10.to_string(index=False)
            else:
                result_text += "⚠️ No polymers match your constraints. Try relaxing some filters."

            self.filter_results.setText(result_text)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Filter error: {str(e)}")

    def update_weight_total(self):
        """Update total weight display"""
        total = sum(w.value() for w in self.weights.values())
        self.weight_total_label.setText(f"Total: {total}%")
        if total != 100:
            self.weight_total_label.setStyleSheet("color: red;")
        else:
            self.weight_total_label.setStyleSheet("color: green;")

    def calculate_ranking(self):
        """Calculate weighted scores and rank polymers"""
        if len(self.filtered_df) == 0:
            QMessageBox.warning(self, "Warning", "No polymers to rank. Apply filters first.")
            return

        # Check weights sum to 100
        total_weight = sum(w.value() for w in self.weights.values())
        if total_weight != 100:
            QMessageBox.warning(self, "Warning", f"Weights sum to {total_weight}%, not 100%.")
            return

        # Normalize properties and calculate scores
        ranked_df = self.filtered_df.copy()

        # Normalize each property (0-1 scale)
        for prop, weight in [('Tensile_MPa', 'Tensile Strength'),
                            ('Flexural_GPa', 'Flexural Modulus'),
                            ('Impact_kJm2', 'Impact Strength'),
                            ('HDT_C', 'Heat Resistance')]:
            max_val = ranked_df[prop].max()
            min_val = ranked_df[prop].min()
            if max_val > min_val:
                ranked_df[f'{prop}_norm'] = (ranked_df[prop] - min_val) / (max_val - min_val)
            else:
                ranked_df[f'{prop}_norm'] = 0.5

        # Cost (lower is better)
        max_cost = ranked_df['Cost_Index'].max()
        min_cost = ranked_df['Cost_Index'].min()
        if max_cost > min_cost:
            ranked_df['Cost_norm'] = 1 - (ranked_df['Cost_Index'] - min_cost) / (max_cost - min_cost)
        else:
            ranked_df['Cost_norm'] = 0.5

        # Water absorption (lower is better)
        max_water = ranked_df['WaterAbs_pct'].max()
        min_water = ranked_df['WaterAbs_pct'].min()
        if max_water > min_water:
            ranked_df['Water_norm'] = 1 - (ranked_df['WaterAbs_pct'] - min_water) / (max_water - min_water)
        else:
            ranked_df['Water_norm'] = 0.5

        # Calculate total score
        ranked_df['Score'] = (
            ranked_df['Tensile_MPa_norm'] * self.weights['Tensile Strength'].value() +
            ranked_df['Flexural_GPa_norm'] * self.weights['Flexural Modulus'].value() +
            ranked_df['Impact_kJm2_norm'] * self.weights['Impact Strength'].value() +
            ranked_df['HDT_C_norm'] * self.weights['Heat Resistance'].value() +
            ranked_df['Cost_norm'] * self.weights['Cost'].value()
        ) / 100

        # Sort by score
        ranked_df = ranked_df.sort_values('Score', ascending=False)

        # Display in table
        self.ranking_table.setRowCount(min(20, len(ranked_df)))
        self.ranking_table.setColumnCount(4)
        self.ranking_table.setHorizontalHeaderLabels(['Rank', 'Polymer', 'Score', 'Family'])

        for i, (idx, row) in enumerate(ranked_df.head(20).iterrows()):
            self.ranking_table.setItem(i, 0, QTableWidgetItem(str(i+1)))
            self.ranking_table.setItem(i, 1, QTableWidgetItem(row['Polymer']))
            self.ranking_table.setItem(i, 2, QTableWidgetItem(f"{row['Score']:.1f}"))
            self.ranking_table.setItem(i, 3, QTableWidgetItem(row['Family']))

        self.ranking_table.resizeColumnsToContents()

        # Create bar chart
        self.ranking_figure.clear()
        ax = self.ranking_figure.add_subplot(111)
        top10 = ranked_df.head(10)
        bars = ax.bar(range(len(top10)), top10['Score'].values*100)
        ax.set_xticks(range(len(top10)))
        ax.set_xticklabels(top10['Polymer'].values, rotation=45, ha='right')
        ax.set_ylabel('Score')
        ax.set_title('Top 10 Polymers by Weighted Score')
        ax.set_ylim(0, 100)

        # Color bars
        for i, bar in enumerate(bars):
            bar.set_color(plt.cm.viridis(i/10))

        self.ranking_figure.tight_layout()
        self.ranking_canvas.draw()

        # Store for later use
        self.ranked_df = ranked_df

    def add_to_shortlist(self):
        """Add selected polymers to shortlist"""
        selected = self.compare_list.selectedItems()
        for item in selected:
            if self.shortlist_widget.findItems(item.text(), Qt.MatchExactly):
                continue
            self.shortlist_widget.addItem(item.text())

        #if len(self.shortlist_widget) > 0:
        self.update_comparison()

    def remove_from_shortlist(self):
        """Remove polymers from shortlist"""
        for item in self.shortlist_widget.selectedItems():
            self.shortlist_widget.takeItem(self.shortlist_widget.row(item))

        self.update_comparison()

    def update_comparison(self):
        """Update comparison table and radar chart"""
        shortlist = [self.shortlist_widget.item(i).text() for i in range(self.shortlist_widget.count())]

        if len(shortlist) == 0:
            self.comparison_table.setRowCount(0)
            self.radar_figure.clear()
            self.radar_canvas.draw()
            return

        # Get data for shortlisted polymers
        compare_df = self.df[self.df['Polymer'].isin(shortlist)]

        # Create comparison table
        properties = ['Polymer', 'Tensile_MPa', 'Flexural_GPa', 'Impact_kJm2',
                     'HDT_C', 'MaxTemp_C', 'Cost_Index', 'UL94']

        self.comparison_table.setRowCount(len(compare_df))
        self.comparison_table.setColumnCount(len(properties))
        self.comparison_table.setHorizontalHeaderLabels(properties)

        for i, (idx, row) in enumerate(compare_df.iterrows()):
            for j, prop in enumerate(properties):
                self.comparison_table.setItem(i, j, QTableWidgetItem(str(row[prop])))

        self.comparison_table.resizeColumnsToContents()

        # Create radar chart
        self.create_radar_chart(compare_df)

        if len(shortlist) >= 5:  # Radar chart gets crowded with >5
            QMessageBox.warning(self, "Warning", "Radar chart gets crowded with over 5 polymers. \nPlease remove some.")



    def create_radar_chart(self, compare_df):
        """Create radar chart for comparison"""
        self.radar_figure.clear()

        # Normalize properties for radar
        properties = ['Tensile_MPa', 'Flexural_GPa', 'Impact_kJm2', 'HDT_C', 'MaxTemp_C']
        labels = ['Tensile\nStrength', 'Flexural\nModulus', 'Impact\nStrength', 'Heat\nResistance', 'Max\nTemperature']

        # Create angles for radar chart
        angles = np.linspace(0, 2*np.pi, len(properties), endpoint=False).tolist()
        angles += angles[:1]  # Close the loop

        ax = self.radar_figure.add_subplot(111, projection='polar')

        for idx, row in compare_df.iterrows():
            values = [row[prop] for prop in properties]
            # Normalize values
            max_vals = [self.df[prop].max() for prop in properties]
            min_vals = [self.df[prop].min() for prop in properties]
            norm_values = [(v - min_vals[i]) / (max_vals[i] - min_vals[i]) for i, v in enumerate(values)]
            norm_values += norm_values[:1]  # Close the loop

            ax.plot(angles, norm_values, 'o-', linewidth=2, label=row['Polymer'])
            ax.fill(angles, norm_values, alpha=0.25)

        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(labels)
        ax.set_ylim(0, 1)
        ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0))
        self.radar_figure.tight_layout()
        self.radar_canvas.draw()

    def show_polymer_details(self):
        """Display detailed information for selected polymer"""
        polymer = self.detail_combo.currentText()
        if not polymer:
            return

        polymer_data = self.df[self.df['Polymer'] == polymer].iloc[0]

        # Create detail text
        details = f"""
        <h2>{polymer}</h2>
        <hr>
        <table border="0" cellpadding="5">
        <tr><td><b>Family:</b></td><td>{polymer_data['Family']}</td></tr>
        <tr><td><b>Tensile Strength:</b></td><td>{polymer_data['Tensile_MPa']} MPa</td></tr>
        <tr><td><b>Flexural Modulus:</b></td><td>{polymer_data['Flexural_GPa']} GPa</td></tr>
        <tr><td><b>Impact Strength:</b></td><td>{polymer_data['Impact_kJm2']} kJ/m²</td></tr>
        <tr><td><b>HDT:</b></td><td>{polymer_data['HDT_C']} °C</td></tr>
        <tr><td><b>Max Operating Temp:</b></td><td>{polymer_data['MaxTemp_C']} °C</td></tr>
        <tr><td><b>Thermal Conductivity:</b></td><td>{polymer_data['ThermalCond_WmK']} W/m·K</td></tr>
        <tr><td><b>Density:</b></td><td>{polymer_data['Density_gcm3']} g/cm³</td></tr>
        <tr><td><b>Water Absorption:</b></td><td>{polymer_data['WaterAbs_pct']}%</td></tr>
        <tr><td><b>Cost Index:</b></td><td>{'$' * polymer_data['Cost_Index']}</td></tr>
        <tr><td><b>UL94 Rating:</b></td><td>{polymer_data['UL94']}</td></tr>
        </table>
        """

        self.detail_text.setHtml(details)

        # Create property bar chart
        self.detail_figure.clear()
        ax = self.detail_figure.add_subplot(111)

        properties = ['Tensile_MPa', 'Flexural_GPa', 'Impact_kJm2', 'HDT_C', 'MaxTemp_C']
        values = [polymer_data[p] for p in properties]

        # Normalize for comparison with max
        max_values = [self.df[p].max() for p in properties]
        norm_values = [v/max_v for v, max_v in zip(values, max_values)]

        bars = ax.bar(range(len(properties)), norm_values)
        ax.set_xticks(range(len(properties)))
        ax.set_xticklabels(['Tensile\nStrength', 'Flexural\nModulus', 'Impact\nStrength',
                           'HDT', 'Max\nTemp'], rotation=0)
        ax.set_ylabel('Normalized Value (0-1)')
        ax.set_title(f'{polymer} - Property Profile')
        ax.set_ylim(0, 1.1)

        # Color bars based on performance
        for i, bar in enumerate(bars):
            if norm_values[i] > 0.7:
                bar.set_color('green')
            elif norm_values[i] > 0.4:
                bar.set_color('orange')
            else:
                bar.set_color('red')

        # Choose a colormap (examples: 'RdYlGn', 'coolwarm', 'viridis')
        cmap = plt.cm.coolwarm

        # Map your values to colors
        colors = cmap(norm_values)  # norm_values should be in range [0,1]

        # Apply colors to bars
        for bar, color in zip(bars, colors):
            bar.set_color(color)


        self.detail_figure.tight_layout()
        self.detail_canvas.draw()

    def generate_report(self):
        """Generate CSV report"""
        try:
            # Create report data
            report_data = {}

            if self.include_project.isChecked():
                report_data['Project Details'] = {
                    'Project Name': self.project_name.text(),
                    'Application': self.application.currentText(),
                    'Environment': self.environment.currentText(),
                    'Process': self.process.currentText(),
                    'Notes': self.notes.toPlainText()
                }

            if self.include_filters.isChecked() and hasattr(self, 'filtered_df'):
                report_data['Filter Criteria'] = {
                    'Min Tensile Strength (MPa)': self.min_strength.value(),
                    'Min Flexural Modulus (GPa)': self.min_flexural.value(),
                    'Min Impact (kJ/m²)': self.min_impact.value(),
                    'Min HDT (°C)': self.min_hdt.value(),
                    'Min Operating Temp (°C)': self.min_temp.value(),
                    'Max Water Absorption (%)': self.max_water.value(),
                    'Max Density (g/cm³)': self.max_density.value(),
                    'Max Cost Index': self.max_cost.value(),
                    'UL94 Rating': self.ul94.currentText()
                }

            if self.include_ranking.isChecked() and hasattr(self, 'ranked_df'):
                report_data['Ranking Results'] = self.ranked_df.head(20)[['Polymer', 'Score', 'Family',
                                                                          'Tensile_MPa', 'HDT_C', 'Cost_Index']]

            if self.include_shortlist.isChecked():
                shortlist = [self.shortlist_widget.item(i).text() for i in range(self.shortlist_widget.count())]
                if shortlist:
                    report_data['Shortlist'] = self.df[self.df['Polymer'].isin(shortlist)]

            # Save to Excel with multiple sheets
            filename = f"polymer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            with pd.ExcelWriter(filename) as writer:
                for sheet_name, data in report_data.items():
                    if isinstance(data, pd.DataFrame):
                        data.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                    elif isinstance(data, dict):
                        pd.DataFrame([data]).to_excel(writer, sheet_name=sheet_name[:31], index=False)

            # Show preview
            preview = f"✅ Report generated successfully!\n\n"
            preview += f"Filename: {filename}\n"
            preview += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
            preview += f"Included sections:\n"
            for section in report_data.keys():
                preview += f"• {section}\n"

            self.report_preview.setText(preview)

            QMessageBox.information(self, "Success", f"Report saved as {filename}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate report: {str(e)}")

    def load_dataset(self):
        """Load custom dataset"""
        filename, _ = QFileDialog.getOpenFileName(self, "Load Polymer Dataset", "",
                                                  "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)")
        if filename:
            try:
                if filename.endswith('.csv'):
                    self.df = pd.read_csv(filename)
                else:
                    self.df = pd.read_excel(filename)

                QMessageBox.information(self, "Success", f"Loaded {len(self.df)} polymers")

                # Refresh UI elements
                self.detail_combo.clear()
                self.detail_combo.addItems(sorted(self.df['Polymer']))

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")

    def export_results(self):
        """Export current results"""
        if hasattr(self, 'ranked_df'):
            filename, _ = QFileDialog.getSaveFileName(self, "Export Results", "polymer_results.csv",
                                                      "CSV Files (*.csv)")
            if filename:
                self.ranked_df.to_csv(filename, index=False)
                QMessageBox.information(self, "Success", f"Results exported to {filename}")

    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(self, "About Polymer Selector",
                         "Polymer Selection Expert v1.0\n\n"
                         "A comprehensive tool for engineers to select optimal polymers\n"
                         "based on mechanical, thermal, and economic constraints.\n\n"
                         "Features:\n"
                         "• Detailed material profiles\n"
                         "• Constraint-based filtering\n"
                         "• Weighted scoring and ranking\n"
                         "• Multi-polymer comparison\n"
                         "• Report generation\n\n"
                         "© 2026 Polymer Selection Tool")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern look
    window = PolymerSelectorApp()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()