import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
import os
import tempfile
from datetime import datetime
import seaborn as sns
import random


# Configure default English fonts for plots
plt.rcParams["font.family"] = "Arial"
plt.rcParams["axes.unicode_minus"] = False  # Fix minus sign display

# English names for random generation
FIRST_NAMES = ["James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph", "Thomas", "Charles",
               "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", "Barbara", "Susan", "Jessica", "Sarah", "Karen"]

LAST_NAMES = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez",
              "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin"]

# Class list
CLASSES = ["Grade 10(1)", "Grade 10(2)", "Grade 10(3)", "Grade 10(4)", "Grade 10(5)",
           "Grade 11(1)", "Grade 11(2)", "Grade 11(3)", "Grade 11(4)", "Grade 11(5)",
           "Grade 12(1)", "Grade 12(2)", "Grade 12(3)", "Grade 12(4)", "Grade 12(5)"]

# Subject list
SUBJECTS = [
    {"name": "Ethics", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Language", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Mathematics", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Foreign Language", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Physics", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "History", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Biology", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Geography", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Chemistry", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Politics", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Science", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Information Technology", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "Current Affairs", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60}
]

class StudentDataGenerator:
    """Student Data Generator Class"""
    def __init__(self, parent):
        self.parent = parent
        self.window = None
        
        # Variables to store subject settings
        self.subject_vars = {}
        self.subject_min = {}
        self.subject_max = {}
        self.subject_pass = {}
        self.subject_pass_rate = {}
        
        # Store generated data
        self.generated_data = None
        
    def show_generator_window(self):
        """Show data generator window"""
        if self.window is not None:
            self.window.lift()
            return
            
        self.window = tk.Toplevel(self.parent)
        self.window.title("Student Information and Grade Random Generator")
        self.window.geometry("900x700")
        self.window.minsize(800, 600)
        self.window.transient(self.parent)
        
        # Clean up reference when window closes
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        self.create_widgets()
        
    def on_window_close(self):
        """Window close handler"""
        self.window.destroy()
        self.window = None
        
    def create_widgets(self):
        # Create main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top settings area
        settings_frame = ttk.LabelFrame(main_frame, text="Generation Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Student count setting
        ttk.Label(settings_frame, text="Student Count:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.student_count = tk.IntVar(value=50)
        ttk.Spinbox(settings_frame, from_=1, to=1000, textvariable=self.student_count, width=10).grid(
            row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Student ID prefix setting
        ttk.Label(settings_frame, text="ID Prefix:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.id_prefix = tk.StringVar(value="2023")
        ttk.Entry(settings_frame, textvariable=self.id_prefix, width=10).grid(
            row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # Class setting
        ttk.Label(settings_frame, text="Class Range:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.class_range = tk.StringVar(value="Random Assignment")
        class_combo = ttk.Combobox(settings_frame, textvariable=self.class_range, width=15, state="readonly")
        class_combo['values'] = ["Random Assignment"] + CLASSES
        class_combo.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        
        # Subject settings area
        subjects_frame = ttk.LabelFrame(main_frame, text="Subject Settings", padding="10")
        subjects_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create scroll area
        canvas = tk.Canvas(subjects_frame)
        scrollbar = ttk.Scrollbar(subjects_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add subject setting rows
        headers = ["Select", "Subject Name", "Min Score", "Max Score", "Pass Score", "Pass Rate(%)"]
        for col, header in enumerate(headers):
            ttk.Label(scrollable_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=5, pady=5, sticky=tk.W)
        
        # Create settings for each subject
        for row, subject in enumerate(SUBJECTS, 1):
            name = subject["name"]
            
            # Checkbox
            var = tk.BooleanVar(value=True)
            self.subject_vars[name] = var
            ttk.Checkbutton(scrollable_frame, variable=var).grid(
                row=row, column=0, padx=5, pady=5)
            
            # Subject name
            ttk.Label(scrollable_frame, text=name).grid(
                row=row, column=1, padx=5, pady=5, sticky=tk.W)
            
            # Min score
            min_var = tk.IntVar(value=subject["min"])
            self.subject_min[name] = min_var
            ttk.Entry(scrollable_frame, textvariable=min_var, width=8).grid(
                row=row, column=2, padx=5, pady=5)
            
            # Max score
            max_var = tk.IntVar(value=subject["max"])
            self.subject_max[name] = max_var
            ttk.Entry(scrollable_frame, textvariable=max_var, width=8).grid(
                row=row, column=3, padx=5, pady=5)
            
            # Pass score
            pass_var = tk.IntVar(value=subject["pass_score"])
            self.subject_pass[name] = pass_var
            ttk.Entry(scrollable_frame, textvariable=pass_var, width=8).grid(
                row=row, column=4, padx=5, pady=5)
            
            # Pass rate
            pass_rate_var = tk.IntVar(value=subject["pass_rate"])
            self.subject_pass_rate[name] = pass_rate_var
            ttk.Entry(scrollable_frame, textvariable=pass_rate_var, width=8).grid(
                row=row, column=5, padx=5, pady=5)
        
        # Button area
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="Generate Data", command=self.generate_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export Data", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset Settings", command=self.reset_settings).pack(side=tk.RIGHT, padx=5)
        
        # Result preview area
        result_frame = ttk.LabelFrame(main_frame, text="Generated Result Preview", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # Result table
        table_frame = ttk.Frame(result_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        
        # Treeview table
        self.result_tree = ttk.Treeview(
            table_frame,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        scrollbar_x.config(command=self.result_tree.xview)
        scrollbar_y.config(command=self.result_tree.yview)
        
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.pack(fill=tk.BOTH, expand=True)
    
    def generate_name(self):
        """Randomly generate English name"""
        first_name = random.choice(FIRST_NAMES)
        last_name = random.choice(LAST_NAMES)
        return f"{first_name} {last_name}"
    
    def generate_id(self, index):
        """Generate student ID"""
        prefix = self.id_prefix.get()
        # Generate fixed-length serial number
        return f"{prefix}{index + 1:04d}"
    
    def generate_class(self):
        """Generate class"""
        class_range = self.class_range.get()
        if class_range == "Random Assignment":
            return random.choice(CLASSES)
        return class_range
    
    def generate_scores_for_subject(self, subject, student_count):
        """Generate scores for all students for a specific subject, ensuring pass rate meets set value"""
        name = subject["name"]
        min_val = self.subject_min[name].get()
        max_val = self.subject_max[name].get()
        pass_score = self.subject_pass[name].get()
        pass_rate = self.subject_pass_rate[name].get() / 100
        
        # Ensure parameters are valid
        if min_val > max_val:
            min_val, max_val = max_val, min_val
        
        # Calculate minimum required passing scores
        min_pass_count = int(pass_rate * student_count)
        # If result is 0 but pass rate > 0, guarantee at least 1 passing score
        if min_pass_count == 0 and pass_rate > 0:
            min_pass_count = 1
        
        # First generate guaranteed passing scores
        scores = []
        # Add necessary passing scores
        for _ in range(min_pass_count):
            scores.append(random.randint(pass_score, max_val))
        
        # Generate remaining scores (can be pass or fail)
        remaining = student_count - min_pass_count
        for _ in range(remaining):
            # 50% chance to generate passing score, making final pass rate possibly higher than set value
            if random.random() < 0.5:
                scores.append(random.randint(pass_score, max_val))
            else:
                scores.append(random.randint(min_val, pass_score - 1))
        
        # Shuffle scores
        random.shuffle(scores)
        return scores
    
    def generate_data(self):
        """Generate student data"""
        try:
            count = self.student_count.get()
            if count <= 0:
                messagebox.showwarning("Warning", "Student count must be greater than 0")
                return
            
            # Collect selected subjects
            selected_subjects = [subj for subj in SUBJECTS 
                               if self.subject_vars[subj["name"]].get()]
            
            if not selected_subjects:
                messagebox.showwarning("Warning", "Please select at least one subject")
                return
            
            # Prepare data structure
            data = []
            columns = ["No.", "Student ID", "Name", "Class"] + [subj["name"] for subj in selected_subjects]
            
            # Generate basic information
            basic_info = []
            for i in range(count):
                basic_info.append({
                    "No.": i + 1,
                    "Student ID": self.generate_id(i),
                    "Name": self.generate_name(),
                    "Class": self.generate_class()
                })
            
            # Generate scores for each subject
            subject_scores = {}
            for subj in selected_subjects:
                subject_scores[subj["name"]] = self.generate_scores_for_subject(subj, count)
            
            # Combine basic info and scores
            for i in range(count):
                row = basic_info[i].copy()
                for subj in selected_subjects:
                    row[subj["name"]] = subject_scores[subj["name"]][i]
                data.append(row)
            
            # Convert to DataFrame
            self.generated_data = pd.DataFrame(data)
            
            # Update table display
            self.update_result_table()
            
            messagebox.showinfo("Success", f"Successfully generated {count} student records")
            
        except Exception as e:
            messagebox.showerror("Error", f"Data generation failed: {str(e)}")
    
    def update_result_table(self):
        """Update result table"""
        # Clear existing data
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        if self.generated_data is None:
            return
        
        # Set columns
        columns = list(self.generated_data.columns)
        self.result_tree["columns"] = columns
        self.result_tree["show"] = "headings"
        
        # Set column names and widths
        for col in columns:
            self.result_tree.heading(col, text=col)
            width = 80 if col in ["Name", "Class"] else 60
            self.result_tree.column(col, width=width, anchor=tk.CENTER)
        
        # Fill data (only show first 50 rows to avoid huge table)
        display_data = self.generated_data.head(50)
        for _, row in display_data.iterrows():
            self.result_tree.insert("", tk.END, values=list(row))
        
        # If there's more data, show note
        if len(self.generated_data) > 50:
            self.result_tree.insert("", tk.END, values=["...", "Showing first 50 rows only", "...", "..."] + ["..." for _ in range(len(columns)-4)])
    
    def export_data(self):
        """Export data"""
        if self.generated_data is None:
            messagebox.showwarning("Warning", "Please generate data first")
            return
        
        # Get save path
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"Student_Grade_Data_{timestamp}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.csv'):
                self.generated_data.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                self.generated_data.to_excel(file_path, index=False)
            
            messagebox.showinfo("Success", f"Data successfully exported to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def reset_settings(self):
        """Reset settings to defaults"""
        # Reset student count
        self.student_count.set(50)
        self.id_prefix.set("2023")
        self.class_range.set("Random Assignment")
        
        # Reset subject settings
        for subject in SUBJECTS:
            name = subject["name"]
            self.subject_vars[name].set(True)
            self.subject_min[name].set(subject["min"])
            self.subject_max[name].set(subject["max"])
            self.subject_pass[name].set(subject["pass_score"])
            self.subject_pass_rate[name].set(subject["pass_rate"])
        
        # Clear results
        self.generated_data = None
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)


class StudentGradeAnalysisSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple Student Grade Analysis System Based on Python")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Data storage
        self.data = None
        self.analysis_results = None
        self.current_file = None
        
        # Create data generator instance
        self.data_generator = StudentDataGenerator(self.root)
        
        # Create main interface
        self.create_widgets()
        
    def create_widgets(self):
        # Create menu bar
        self.create_menu()
        
        # Create analysis button area (at the top)
        analysis_frame = ttk.Frame(self.root, padding="10")
        analysis_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Analysis buttons
        ttk.Button(
            analysis_frame, 
            text="Basic Statistical Analysis", 
            command=self.perform_basic_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="Subject Comparison Analysis", 
            command=self.perform_subject_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="Grade Distribution Analysis", 
            command=self.perform_distribution_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="Advanced Analysis", 
            command=self.perform_advanced_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="Export Analysis Results", 
            command=self.export_analysis_results
        ).pack(side=tk.RIGHT, padx=5)
        
        # Add Generate PDF Report button
        ttk.Button(
            analysis_frame, 
            text="Generate PDF Report", 
            command=self.generate_pdf_report
        ).pack(side=tk.RIGHT, padx=5)
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create left panel (data display) - fixed width
        left_frame = ttk.LabelFrame(main_frame, text="Grade Data", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.config(width=600)  # Set fixed width as 600 pixels
        left_frame.pack_propagate(False)  # Prevent child widgets from resizing parent
        
        # Create data table
        table_frame = ttk.Frame(left_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        
        # Treeview table
        self.tree = ttk.Treeview(
            table_frame,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        scrollbar_x.config(command=self.tree.xview)
        scrollbar_y.config(command=self.tree.yview)
        
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Create right panel (analysis results) - take remaining space
        right_frame = ttk.LabelFrame(main_frame, text="Analysis Results", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # Analysis tabs
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Statistical analysis tab
        self.stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_frame, text="Statistical Analysis")
        
        # Visualization tab
        self.visual_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.visual_frame, text="Data Visualization")
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Import Data", command=self.import_data)
        file_menu.add_command(label="Save Data", command=self.save_data)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Student Data Generator", command=self.open_data_generator)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="User Guide", command=self.show_help)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def open_data_generator(self):
        """Open data generator window"""
        self.data_generator.show_generator_window()
    
    def import_data(self):
        """Import data file"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            # Choose appropriate reading method based on file extension
            if file_path.endswith('.csv'):
                self.data = pd.read_csv(file_path)
            else:  # Excel files
                self.data = pd.read_excel(file_path)
            
            self.current_file = file_path
            self.update_table()
            messagebox.showinfo("Success", f"Data imported successfully, total {len(self.data)} records")
            
        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {str(e)}")
    
    def save_data(self):
        """Save data"""
        if self.data is None:
            messagebox.showwarning("Warning", "No data to save")
            return
            
        if self.current_file:
            try:
                if self.current_file.endswith('.csv'):
                    self.data.to_csv(self.current_file, index=False)
                else:
                    self.data.to_excel(self.current_file, index=False)
                messagebox.showinfo("Success", "Data saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Save failed: {str(e)}")
        else:
            self.export_data()
    
    def export_data(self):
        """Export data"""
        if self.data is None:
            messagebox.showwarning("Warning", "No data to export")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.endswith('.csv'):
                self.data.to_csv(file_path, index=False)
            else:
                self.data.to_excel(file_path, index=False)
            self.current_file = file_path
            messagebox.showinfo("Success", "Data exported successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def update_table(self):
        """Update table data"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Set columns
        self.tree["columns"] = list(self.data.columns)
        self.tree["show"] = "headings"
        
        # Set column names
        for col in self.data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)
        
        # Fill data
        for idx, row in self.data.iterrows():
            self.tree.insert("", tk.END, values=list(row))
    
    def perform_basic_analysis(self):
        """Perform basic statistical analysis"""
        if self.data is None:
            messagebox.showwarning("Warning", "Please import data first")
            return
            
        # Clear statistical analysis panel
        for widget in self.stats_frame.winfo_children():
            widget.destroy()
        
        # Create scrollable text box to display statistical results
        text_frame = ttk.Frame(self.stats_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        stats_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, padx=10, pady=10)
        stats_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=stats_text.yview)
        
        # Basic statistical information
        stats_text.insert(tk.END, "===== Basic Statistical Information =====\n\n")
        stats_text.insert(tk.END, f"Total Records: {len(self.data)}\n")
        
        # Assuming "Name" column exists
        if "Name" in self.data.columns:
            stats_text.insert(tk.END, f"Student Count: {self.data['Name'].nunique()}\n\n")
        
        # Identify subject columns (assume columns not ID or name are subjects)
        subject_columns = []
        for col in self.data.columns:
            if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]:
                subject_columns.append(col)
        
        # Calculate total score (if not exists)
        if "Total Score" not in self.data.columns and subject_columns:
            self.data["Total Score"] = self.data[subject_columns].sum(axis=1)
        
        # Calculate average score (if not exists)
        if "Average Score" not in self.data.columns and subject_columns:
            self.data["Average Score"] = self.data[subject_columns].mean(axis=1)
        
        # Calculate ranking (if not exists)
        if "Rank" not in self.data.columns and "Total Score" in self.data.columns:
            self.data["Rank"] = self.data["Total Score"].rank(ascending=False, method="min").astype(int)
        
        # Update table
        self.update_table()
        
        # Subject statistics
        stats_text.insert(tk.END, "===== Subject Statistics =====\n\n")
        for subject in subject_columns:
            stats_text.insert(tk.END, f"{subject}:\n")
            stats_text.insert(tk.END, f"  Average Score: {self.data[subject].mean():.2f}\n")
            stats_text.insert(tk.END, f"  Highest Score: {self.data[subject].max()}\n")
            stats_text.insert(tk.END, f"  Lowest Score: {self.data[subject].min()}\n")
            stats_text.insert(tk.END, f"  Pass Rate: {((self.data[subject] >= 60).sum() / len(self.data) * 100):.2f}%\n")
            stats_text.insert(tk.END, f"  Excellent Rate(>=90): {((self.data[subject] >= 90).sum() / len(self.data) * 100):.2f}%\n\n")
        
        # Total score statistics
        if "Total Score" in self.data.columns:
            stats_text.insert(tk.END, "===== Total Score Statistics =====\n\n")
            stats_text.insert(tk.END, f"  Average Score: {self.data['Total Score'].mean():.2f}\n")
            stats_text.insert(tk.END, f"  Highest Score: {self.data['Total Score'].max()}\n")
            stats_text.insert(tk.END, f"  Lowest Score: {self.data['Total Score'].min()}\n\n")
        
        # Add more statistical functions
        stats_text.insert(tk.END, "===== Detailed Subject Statistics =====\n\n")
        for subject in subject_columns:
            stats_text.insert(tk.END, f"{subject}:")
            stats_text.insert(tk.END, f"\n  Average Score: {self.data[subject].mean():.2f}")
            stats_text.insert(tk.END, f"\n  Median: {self.data[subject].median():.2f}")
            stats_text.insert(tk.END, f"\n  Standard Deviation: {self.data[subject].std():.2f}")
            stats_text.insert(tk.END, f"\n  Minimum: {self.data[subject].min()}")
            stats_text.insert(tk.END, f"\n  Maximum: {self.data[subject].max()}")
            stats_text.insert(tk.END, f"\n  25th Percentile: {self.data[subject].quantile(0.25):.2f}")
            stats_text.insert(tk.END, f"\n  75th Percentile: {self.data[subject].quantile(0.75):.2f}\n\n")
        
        # Save analysis results
        self.analysis_results = {
            "basic_stats": self.data.describe(),
            "subject_columns": subject_columns
        }
        
        # Make text box read-only
        stats_text.config(state=tk.DISABLED)
        
        # Switch to statistical analysis tab
        self.notebook.select(self.stats_frame)
    
    def perform_subject_analysis(self):
        """Perform subject comparison analysis"""
        if self.data is None:
            messagebox.showwarning("Warning", "Please import data first")
            return
            
        # Clear visualization panel
        for widget in self.visual_frame.winfo_children():
            widget.destroy()
        
        # Identify subject columns
        subject_columns = []
        for col in self.data.columns:
            if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]:
                subject_columns.append(col)
        
        if not subject_columns:
            messagebox.showwarning("Warning", "No subject columns identified")
            return
        
        # Create a main frame to hold all charts
        main_canvas = tk.Canvas(self.visual_frame)
        scrollbar_v = ttk.Scrollbar(self.visual_frame, orient="vertical", command=main_canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.visual_frame, orient="horizontal", command=main_canvas.xview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # Position horizontal scrollbar at bottom of visualization interface
        main_canvas.pack(side="top", fill="both", expand=True)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store all charts for export
        self.current_figures = []
        self.current_analysis_type = 'subject'  # Mark analysis type
        
        # 1. Average score comparison chart for each subject
        avg_frame = ttk.LabelFrame(scrollable_frame, text="Average Score Comparison by Subject", padding="10")
        avg_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig1, ax1 = plt.subplots(1, 1, figsize=(12, 6))
        avg_scores = [self.data[subject].mean() for subject in subject_columns]
        
        # Use more attractive colors
        colors = plt.cm.Set3(range(len(subject_columns)))
        bars = ax1.bar(subject_columns, avg_scores, color=colors)
        
        # Add value labels
        for bar, score in zip(bars, avg_scores):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{score:.1f}', ha='center', va='bottom', fontweight='bold')
        
        ax1.set_title('Average Score Comparison by Subject', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Average Score')
        ax1.grid(axis='y', linestyle='--', alpha=0.7)
        
        # Auto-adjust x-axis label angle
        if len(subject_columns) > 5:
            plt.setp(ax1.get_xticklabels(), rotation=45, ha='right')
        
        plt.tight_layout()
        
        # Embed chart in frame
        canvas1 = FigureCanvasTkAgg(fig1, master=avg_frame)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig1)
        
        # 2. Score distribution box plot for each subject
        box_frame = ttk.LabelFrame(scrollable_frame, text="Score Distribution Box Plot by Subject", padding="10")
        box_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig2, ax2 = plt.subplots(1, 1, figsize=(12, 6))
        
        # Prepare box plot data
        box_data = [self.data[subject].dropna() for subject in subject_columns]
        
        # Create box plot
        box_plot = ax2.boxplot(box_data, labels=subject_columns, patch_artist=True)
        
        # Beautify box plot
        colors2 = plt.cm.Pastel1(range(len(subject_columns)))
        for patch, color in zip(box_plot['boxes'], colors2):
            patch.set_facecolor(color)
            patch.set_alpha(0.8)
        
        # Set median line color
        for median in box_plot['medians']:
            median.set_color('red')
            median.set_linewidth(2)
        
        ax2.set_title('Score Distribution Box Plot by Subject', fontsize=14, fontweight='bold')
        ax2.set_ylabel('Score')
        ax2.grid(axis='y', linestyle='--', alpha=0.7)
        
        # Auto-adjust x-axis label angle
        if len(subject_columns) > 5:
            plt.setp(ax2.get_xticklabels(), rotation=45, ha='right')
        
        plt.tight_layout()
        
        # Embed chart in frame
        canvas2 = FigureCanvasTkAgg(fig2, master=box_frame)
        canvas2.draw()
        canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig2)
        
        # Bind mouse wheel event
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Save first chart for export (for compatibility)
        if self.current_figures:
            self.current_fig = self.current_figures[0]
        
        # Switch to visualization tab
        self.notebook.select(self.visual_frame)
    
    def perform_distribution_analysis(self):
        """Perform grade distribution analysis"""
        if self.data is None:
            messagebox.showwarning("Warning", "Please import data first")
            return
            
        # Clear visualization panel
        for widget in self.visual_frame.winfo_children():
            widget.destroy()
        
        # Identify subject columns
        subject_columns = []
        for col in self.data.columns:
            if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]:
                subject_columns.append(col)
        
        if not subject_columns:
            messagebox.showwarning("Warning", "No subject columns identified")
            return
        
        # Create a main frame to hold all charts
        main_canvas = tk.Canvas(self.visual_frame)
        scrollbar_v = ttk.Scrollbar(self.visual_frame, orient="vertical", command=main_canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.visual_frame, orient="horizontal", command=main_canvas.xview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # Position horizontal scrollbar at bottom of visualization interface
        main_canvas.pack(side="top", fill="both", expand=True)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store all charts for export
        self.current_figures = []
        self.current_analysis_type = 'distribution'  # Mark analysis type
        
        # 1. If total score exists, first show total score distribution
        if "Total Score" in self.data.columns:
            total_frame = ttk.LabelFrame(scrollable_frame, text="Total Score Distribution Analysis", padding="10")
            total_frame.pack(fill=tk.X, padx=10, pady=5)
            
            fig_total, ax_total = plt.subplots(1, 1, figsize=(12, 6))
            sns.histplot(self.data["Total Score"], bins=15, kde=True, ax=ax_total, color='green')
            ax_total.set_title('Total Score Distribution Histogram', fontsize=14, fontweight='bold')
            ax_total.set_xlabel('Total Score')
            ax_total.set_ylabel('Student Count')
            ax_total.grid(axis='y', linestyle='--', alpha=0.7)
            
            # Add statistical information
            mean_score = self.data["Total Score"].mean()
            max_score = self.data["Total Score"].max()
            min_score = self.data["Total Score"].min()
            ax_total.axvline(mean_score, color='red', linestyle='--', linewidth=2, label=f'Average: {mean_score:.1f}')
            ax_total.legend()
            
            plt.tight_layout()
            
            canvas_total = FigureCanvasTkAgg(fig_total, master=total_frame)
            canvas_total.draw()
            canvas_total.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig_total)
        
        # 2. Create independent analysis charts for each subject
        colors = sns.color_palette('Set2', n_colors=len(subject_columns))
        
        for i, subject in enumerate(subject_columns):
            # Create independent frame for each subject
            subject_frame = ttk.LabelFrame(scrollable_frame, text=f"{subject} Analysis", padding="10")
            subject_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Create chart with two subplots: histogram and pie chart
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            
            # Left: Score distribution histogram
            sns.histplot(self.data[subject], bins=12, kde=True, ax=ax1, color=colors[i])
            ax1.set_title(f'{subject} Score Distribution Histogram', fontsize=12, fontweight='bold')
            ax1.set_xlabel('Score')
            ax1.set_ylabel('Student Count')
            ax1.grid(axis='y', linestyle='--', alpha=0.7)
            
            # Add statistical information to histogram
            mean_score = self.data[subject].mean()
            pass_score = 60 if self.data[subject].max() <= 100 else 90
            ax1.axvline(mean_score, color='red', linestyle='--', linewidth=2, label=f'Average: {mean_score:.1f}')
            ax1.axvline(pass_score, color='orange', linestyle='--', linewidth=2, label=f'Pass Line: {pass_score}')
            ax1.legend()
            
            # Right: Score range percentage pie chart
            subject_max = self.data[subject].max()
            if subject_max <= 100:
                bins = [0, 60, 70, 80, 90, 100]
                labels = ['Fail\n(0-59)', 'Pass\n(60-69)', 'Average\n(70-79)', 'Good\n(80-89)', 'Excellent\n(90-100)']
            elif subject_max <= 150:
                bins = [0, 90, 105, 120, 135, 150]
                labels = ['Fail\n(0-89)', 'Pass\n(90-104)', 'Average\n(105-119)', 'Good\n(120-134)', 'Excellent\n(135-150)']
            else:
                # For other scoring systems, use percentage segmentation
                bins = [0, subject_max*0.6, subject_max*0.7, subject_max*0.8, subject_max*0.9, subject_max]
                labels = ['Fail', 'Pass', 'Average', 'Good', 'Excellent']
            
            score_ranges = pd.cut(self.data[subject], bins=bins, labels=labels, include_lowest=True)
            range_counts = score_ranges.value_counts(normalize=True).reindex(labels) * 100
            
            # Filter out segments with 0%
            non_zero_counts = range_counts[range_counts > 0]
            non_zero_labels = non_zero_counts.index.tolist()
            
            # Use more attractive colors
            pie_colors = sns.color_palette('viridis', n_colors=len(non_zero_labels))
            
            wedges, texts, autotexts = ax2.pie(non_zero_counts, labels=non_zero_labels, autopct='%1.1f%%', 
                                             startangle=90, colors=pie_colors, explode=[0.05]*len(non_zero_labels))
            
            # Beautify pie chart text
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(10)
            
            ax2.set_title(f'{subject} Score Range Percentage', fontsize=12, fontweight='bold')
            
            plt.tight_layout()
            
            # Embed chart in frame
            canvas = FigureCanvasTkAgg(fig, master=subject_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig)
        
        # Bind mouse wheel event
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Save first chart for export (for compatibility)
        if self.current_figures:
            self.current_fig = self.current_figures[0]
        
        # Switch to visualization tab
        self.notebook.select(self.visual_frame)
    
    def export_analysis_results(self):
        """Export analysis results"""
        if self.data is None or self.analysis_results is None:
            messagebox.showwarning("Warning", "Please perform analysis first")
            return
            
        # Create export dialog
        export_window = tk.Toplevel(self.root)
        export_window.title("Export Analysis Results")
        export_window.geometry("400x300")
        export_window.transient(self.root)
        export_window.grab_set()
        
        # Export options
        ttk.Label(export_window, text="Please select content to export:").pack(pady=10)
        
        var_stats = tk.BooleanVar(value=True)
        var_chart = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(export_window, text="Statistical Data", variable=var_stats).pack(anchor=tk.W, padx=20, pady=5)
        ttk.Checkbutton(export_window, text="Charts", variable=var_chart).pack(anchor=tk.W, padx=20, pady=5)
        
        # Format selection
        ttk.Label(export_window, text="Please select export format:").pack(pady=10)
        
        frame_format = ttk.Frame(export_window)
        frame_format.pack(pady=5)
        
        var_format = tk.StringVar(value="excel")
        ttk.Radiobutton(frame_format, text="Excel", variable=var_format, value="excel").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(frame_format, text="CSV", variable=var_format, value="csv").pack(side=tk.LEFT, padx=10)
        
        # Export button
        def do_export():
            if not var_stats.get() and not var_chart.get():
                messagebox.showwarning("Warning", "Please select at least one item to export")
                return
                
            # Get save path
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"Grade_Analysis_Results_{timestamp}"
            
            if var_stats.get():
                file_ext = ".xlsx" if var_format.get() == "excel" else ".csv"
                file_path = filedialog.asksaveasfilename(
                    defaultextension=file_ext,
                    initialfile=default_filename,
                    filetypes=[
                        ("Excel files", "*.xlsx"),
                        ("CSV files", "*.csv"),
                        ("All files", "*.*")
                    ]
                )
                
                if file_path:
                    try:
                        # Create results table
                        result_df = pd.DataFrame()
                        
                        # Add basic statistics
                        stats_df = self.analysis_results["basic_stats"]
                        
                        # Add top 10 students by total score
                        if "Total Score" in self.data.columns:
                            top_students = self.data.sort_values("Total Score", ascending=False).head(10)
                            result_df = pd.concat([result_df, pd.DataFrame(["", "Top 10 Students by Total Score", ""], columns=["Note"])])
                            result_df = pd.concat([result_df, top_students[["Name", "Total Score", "Rank"]]])
                        
                        # Save to file
                        if file_path.endswith('.csv'):
                            result_df.to_csv(file_path, index=False)
                        else:
                            with pd.ExcelWriter(file_path) as writer:
                                stats_df.to_excel(writer, sheet_name='Statistical Data')
                                if "Total Score" in self.data.columns:
                                    top_students.to_excel(writer, sheet_name='Top Students', index=False)
                
                    except Exception as e:
                        messagebox.showerror("Error", f"Export statistical data failed: {str(e)}")
                        return
            
            # Export charts
            if var_chart.get() and hasattr(self, 'current_figures'):
                # If multiple charts, ask user if they want to export separately
                if len(self.current_figures) > 1:
                    export_choice = messagebox.askyesnocancel(
                        "Chart Export", 
                        "Multiple charts detected.\nClick 'Yes' to export each chart separately\nClick 'No' to export as single file\nClick 'Cancel' to skip chart export"
                    )
                    
                    if export_choice is None:  # Cancel
                        pass
                    elif export_choice:  # Export separately
                        base_path = filedialog.askdirectory(title="Select Save Directory")
                        if base_path:
                            try:
                                for i, fig in enumerate(self.current_figures):
                                    if i == 0 and "Total Score" in self.data.columns:
                                        filename = f"{default_filename}_Total_Score_Distribution.png"
                                    elif hasattr(self, 'current_analysis_type'):
                                        # Name based on analysis type
                                        if self.current_analysis_type == 'subject':
                                            if i == 0 or (i == 1 and "Total Score" not in self.data.columns):
                                                filename = f"{default_filename}_Subject_Average_Comparison.png"
                                            elif i == 1 or (i == 2 and "Total Score" in self.data.columns):
                                                filename = f"{default_filename}_Subject_Score_Distribution_Box.png"
                                            else:
                                                filename = f"{default_filename}_Subject_Comparison_Chart{i+1}.png"
                                        else:  # distribution analysis
                                            # Get subject name
                                            subject_columns = [col for col in self.data.columns 
                                                             if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]]
                                            subject_idx = i - (1 if "Total Score" in self.data.columns else 0)
                                            if subject_idx < len(subject_columns):
                                                subject_name = subject_columns[subject_idx]
                                                filename = f"{default_filename}_{subject_name}_Analysis.png"
                                            else:
                                                filename = f"{default_filename}_Chart{i+1}.png"
                                    else:
                                        # Get subject name (compatibility handling)
                                        subject_columns = [col for col in self.data.columns 
                                                         if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]]
                                        subject_idx = i - (1 if "Total Score" in self.data.columns else 0)
                                        if subject_idx < len(subject_columns):
                                            subject_name = subject_columns[subject_idx]
                                            filename = f"{default_filename}_{subject_name}_Analysis.png"
                                        else:
                                            filename = f"{default_filename}_Chart{i+1}.png"
                                    
                                    file_path = os.path.join(base_path, filename)
                                    fig.savefig(file_path, dpi=300, bbox_inches='tight')
                                
                                messagebox.showinfo("Success", f"Exported {len(self.current_figures)} charts to:\n{base_path}")
                            except Exception as e:
                                messagebox.showerror("Error", f"Chart export failed: {str(e)}")
                                return
                    else:  # Export as single file
                        file_path = filedialog.asksaveasfilename(
                            defaultextension=".png",
                            initialfile=default_filename,
                            filetypes=[
                                ("PNG files", "*.png"),
                                ("JPEG files", "*.jpg"),
                                ("PDF files", "*.pdf"),
                                ("All files", "*.*")
                            ]
                        )
                        
                        if file_path:
                            try:
                                # Combine all charts into one large figure
                                from matplotlib.backends.backend_pdf import PdfPages
                                
                                if file_path.endswith('.pdf'):
                                    with PdfPages(file_path) as pdf:
                                        for fig in self.current_figures:
                                            pdf.savefig(fig, bbox_inches='tight')
                                else:
                                    # For image formats, only save first chart
                                    self.current_figures[0].savefig(file_path, dpi=300, bbox_inches='tight')
                                    
                            except Exception as e:
                                messagebox.showerror("Error", f"Chart export failed: {str(e)}")
                                return
                else:
                    # Only one chart
                    file_path = filedialog.asksaveasfilename(
                        defaultextension=".png",
                        initialfile=default_filename,
                        filetypes=[
                            ("PNG files", "*.png"),
                            ("JPEG files", "*.jpg"),
                            ("PDF files", "*.pdf"),
                            ("All files", "*.*")
                        ]
                    )
                    
                    if file_path:
                        try:
                            self.current_figures[0].savefig(file_path, dpi=300, bbox_inches='tight')
                        except Exception as e:
                            messagebox.showerror("Error", f"Chart export failed: {str(e)}")
                            return
            elif var_chart.get() and hasattr(self, 'current_fig'):
                # Compatibility for old version single chart export
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    initialfile=default_filename,
                    filetypes=[
                        ("PNG files", "*.png"),
                        ("JPEG files", "*.jpg"),
                        ("PDF files", "*.pdf"),
                        ("All files", "*.*")
                    ]
                )
                
                if file_path:
                    try:
                        self.current_fig.savefig(file_path, dpi=300, bbox_inches='tight')
                    except Exception as e:
                        messagebox.showerror("Error", f"Chart export failed: {str(e)}")
                        return
        
            messagebox.showinfo("Success", "Analysis results exported successfully")
            export_window.destroy()
    
    def generate_pdf_report(self):
        """Generate PDF report"""
        if self.data is None or self.analysis_results is None:
            messagebox.showwarning("Warning", "Please perform analysis first")
            return

        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        import os
        import tempfile
        import matplotlib.font_manager as fm

        # Font setup helper function
        def set_pdf_font(pdf, size=12, bold=False):
            """Set PDF font"""
            pdf.setFont("Helvetica-Bold", size) if bold else pdf.setFont("Helvetica", size)

        # Get save path
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*")]
        )

        if not file_path:
            return

        try:
            # Create PDF canvas
            pdf = canvas.Canvas(file_path, pagesize=letter)
            width, height = letter

            # Add title
            set_pdf_font(pdf, 16, bold=True)
            pdf.drawString(30, height - 30, "Student Grade Analysis Report")

            # Add statistical data
            y_position = height - 60
            set_pdf_font(pdf, 12)
            pdf.drawString(30, y_position, "===== Basic Statistical Information =====")
            y_position -= 20
            pdf.drawString(30, y_position, f"Total Records: {len(self.data)}")

            if "Name" in self.data.columns:
                y_position -= 20
                pdf.drawString(30, y_position, f"Student Count: {self.data['Name'].nunique()}")

            y_position -= 40
            pdf.drawString(30, y_position, "===== Detailed Subject Statistics =====")
            y_position -= 20

            for subject in self.analysis_results["subject_columns"]:
                if y_position < 50:
                    pdf.showPage()
                    y_position = height - 30

                pdf.drawString(30, y_position, f"{subject}:")
                y_position -= 20
                pdf.drawString(50, y_position, f"Average Score: {self.data[subject].mean():.2f}")
                y_position -= 20
                pdf.drawString(50, y_position, f"Median: {self.data[subject].median():.2f}")
                y_position -= 20
                pdf.drawString(50, y_position, f"Standard Deviation: {self.data[subject].std():.2f}")
                y_position -= 20

            # Add more statistical content
            if "Total Score" in self.data.columns:
                y_position -= 20
                pdf.drawString(30, y_position, "===== Total Score Statistics =====")
                y_position -= 20
                pdf.drawString(50, y_position, f"Average Score: {self.data['Total Score'].mean():.2f}")
                y_position -= 20
                pdf.drawString(50, y_position, f"Highest Score: {self.data['Total Score'].max()}")
                y_position -= 20
                pdf.drawString(50, y_position, f"Lowest Score: {self.data['Total Score'].min()}")
                y_position -= 20
                pdf.drawString(50, y_position, f"Standard Deviation: {self.data['Total Score'].std():.2f}")
                y_position -= 30
                
                # Add top 10 students
                pdf.drawString(30, y_position, "===== Top 10 Students by Score =====")
                y_position -= 20
                top_students = self.data.sort_values("Total Score", ascending=False).head(10)
                for idx, (_, student) in enumerate(top_students.iterrows()):
                    if y_position < 50:
                        pdf.showPage()
                        y_position = height - 30
                        set_pdf_font(pdf, 12)  # Reset font
                    
                    student_name = student['Name']
                    total_score = student['Total Score']
                    student_info = f"{idx+1}. {student_name}: {total_score}"
                    pdf.drawString(50, y_position, student_info)
                    
                    y_position -= 20

            # Add all charts
            if hasattr(self, 'current_figures'):
                for i, fig in enumerate(self.current_figures):
                    # Check if new page needed
                    if y_position < 350:
                        pdf.showPage()
                        y_position = height - 30

                    # Add chart title
                    set_pdf_font(pdf, 14, bold=True)
                    if i == 0 and "Total Score" in self.data.columns:
                        chart_title = "Total Score Distribution Analysis Chart"
                    elif hasattr(self, 'current_analysis_type'):
                        if self.current_analysis_type == 'subject':
                            if i == 0 or (i == 1 and "Total Score" not in self.data.columns):
                                chart_title = "Subject Average Comparison Chart"
                            elif i == 1 or (i == 2 and "Total Score" in self.data.columns):
                                chart_title = "Subject Score Distribution Box Plot"
                            else:
                                chart_title = f"Subject Comparison Chart {i+1}"
                        else:  # distribution analysis
                            subject_columns = [col for col in self.data.columns 
                                             if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]]
                            subject_idx = i - (1 if "Total Score" in self.data.columns else 0)
                            if subject_idx < len(subject_columns):
                                subject_name = subject_columns[subject_idx]
                                chart_title = f"{subject_name} Analysis Chart"
                            else:
                                chart_title = f"Analysis Chart {i+1}"
                    else:
                        chart_title = f"Analysis Chart {i+1}"
                    
                    pdf.drawString(30, y_position, chart_title)
                    y_position -= 30

                    # Save chart as image (using unique filename)
                    import tempfile
                    import os
                    
                    # Create temporary file
                    temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                    temp_file.close()
                    img_path = temp_file.name
                    
                    try:
                        # Save chart
                        fig.savefig(img_path, format="png", dpi=150, bbox_inches='tight')

                        # Insert image into PDF
                        pdf.drawImage(img_path, 30, y_position - 280, width=500, height=280)
                        y_position -= 320
                        
                    finally:
                        # Clean up temporary file
                        if os.path.exists(img_path):
                            os.unlink(img_path)

            # Add analysis conclusions
            if y_position < 150:
                pdf.showPage()
                y_position = height - 30
            
            set_pdf_font(pdf, 14, bold=True)
            pdf.drawString(30, y_position, "===== Analysis Conclusions =====")
            y_position -= 30
            set_pdf_font(pdf, 12)
            
            # Calculate overall performance
            subject_columns = self.analysis_results["subject_columns"]
            overall_avg = sum(self.data[subject].mean() for subject in subject_columns) / len(subject_columns)
            pdf.drawString(50, y_position, f"1. Overall subject average: {overall_avg:.2f}")
            y_position -= 20
            
            # Find best and worst performing subjects
            best_subject = max(subject_columns, key=lambda x: self.data[x].mean())
            worst_subject = min(subject_columns, key=lambda x: self.data[x].mean())
            pdf.drawString(50, y_position, f"2. Best performing subject: {best_subject} ({self.data[best_subject].mean():.2f})")
            y_position -= 20
            pdf.drawString(50, y_position, f"3. Subject needing improvement: {worst_subject} ({self.data[worst_subject].mean():.2f})")
            y_position -= 20
            
            # Overall pass rate
            overall_pass_rate = sum((self.data[subject] >= 60).sum() for subject in subject_columns) / (len(subject_columns) * len(self.data)) * 100
            pdf.drawString(50, y_position, f"4. Overall pass rate: {overall_pass_rate:.2f}%")
            y_position -= 30
            
            # Add generation time
            from datetime import datetime
            pdf.drawString(50, y_position, f"Report generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

            pdf.save()
            messagebox.showinfo("Success", "PDF report generated successfully")

        except Exception as e:
            messagebox.showerror("Error", f"PDF report generation failed: {str(e)}")
    
    def show_about(self):
        """Show about information"""
        messagebox.showinfo(
            "About",
            "Simple Student Grade Analysis System v1.0\n\n"
            "A tool for analyzing student grades, supporting data import, statistical analysis, data visualization and result export.\n\n"
            "Features:\n- Integrated student data generator\n- Can generate random student grade data\n\n"
        )
    
    def show_help(self):
        """Show user guide"""
        help_window = tk.Toplevel(self.root)
        help_window.title("User Guide")
        help_window.geometry("600x500")
        help_window.transient(self.root)
        
        # Create scrollable text box
        scrollbar = ttk.Scrollbar(help_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        help_text = tk.Text(help_window, wrap=tk.WORD, yscrollcommand=scrollbar.set, padx=10, pady=10)
        help_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=help_text.yview)
        
        # Help content
        help_content = """
Simple Student Grade Analysis System User Guide

1. Data Import
   - Click "File" -> "Import Data" in menu bar
   - Supports Excel(.xlsx, .xls) and CSV(.csv) format files
   - Data should contain student information (name, student ID, etc.) and subject grades

2. Data Generation
   - Click "Tools" -> "Student Data Generator" in menu bar
   - Can customize student count, ID prefix, class range
   - Can select subjects and set score range, pass score, pass rate
   - Generated data can be directly exported as Excel or CSV files

3. Basic Operations
   - After importing data, system displays data in left table
   - Can use buttons at bottom for various analyses

4. Analysis Features
   - Basic Statistical Analysis: Calculate average score, highest score, lowest score, pass rate, etc. for each subject
   - Subject Comparison Analysis: Generate subject average comparison charts and score distribution box plots
   - Grade Distribution Analysis: Generate total score distribution histograms and score range percentage pie charts

5. Result Export
   - Click "Export Analysis Results" button
   - Can choose to export statistical data and charts
   - Statistical data supports Excel and CSV formats, charts support PNG, JPG and PDF formats

6. Data Saving
   - Click "File" -> "Save Data" to save current data
   - Click "File" -> "Export Data" to export data as new file
        """
        
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)
    
    def perform_advanced_analysis(self):
        """Perform advanced analysis"""
        if self.data is None:
            messagebox.showwarning("Warning", "Please import data first")
            return
            
        # Clear visualization panel
        for widget in self.visual_frame.winfo_children():
            widget.destroy()
        
        # Identify subject columns
        subject_columns = []
        for col in self.data.columns:
            if col not in ["No.", "Student ID", "Name", "Class", "Total Score", "Rank"]:
                subject_columns.append(col)
        
        if not subject_columns:
            messagebox.showwarning("Warning", "No subject columns identified")
            return
        
        # Create a main frame to hold all charts
        main_canvas = tk.Canvas(self.visual_frame)
        scrollbar_v = ttk.Scrollbar(self.visual_frame, orient="vertical", command=main_canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.visual_frame, orient="horizontal", command=main_canvas.xview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # Position horizontal scrollbar at bottom of visualization interface
        main_canvas.pack(side="top", fill="both", expand=True)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store all charts for export
        self.current_figures = []
        self.current_analysis_type = 'advanced'  # Mark analysis type
        
        # 1. Subject correlation heatmap
        corr_frame = ttk.LabelFrame(scrollable_frame, text="Subject Correlation Analysis", padding="10")
        corr_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig1, ax1 = plt.subplots(1, 1, figsize=(12, 8))
        
        # Calculate correlation matrix
        correlation_matrix = self.data[subject_columns].corr()
        
        # Create heatmap
        sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0,
                   square=True, ax=ax1, fmt='.2f', cbar_kws={'shrink': .8})
        ax1.set_title('Subject Correlation Heatmap', fontsize=14, fontweight='bold')
        plt.tight_layout()
        
        canvas1 = FigureCanvasTkAgg(fig1, master=corr_frame)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig1)
        
        # 2. Class grade comparison analysis (if class information exists)
        if "Class" in self.data.columns:
            class_frame = ttk.LabelFrame(scrollable_frame, text="Class Grade Comparison Analysis", padding="10")
            class_frame.pack(fill=tk.X, padx=10, pady=5)
            
            fig2, ax2 = plt.subplots(1, 1, figsize=(12, 6))
            
            # Calculate average scores for each class
            class_avg = self.data.groupby("Class")[subject_columns].mean()
            
            # Create stacked bar chart
            class_avg.plot(kind='bar', ax=ax2, width=0.8)
            ax2.set_title('Average Score Comparison by Class and Subject', fontsize=14, fontweight='bold')
            ax2.set_ylabel('Average Score')
            ax2.set_xlabel('Class')
            ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.xticks(rotation=45)
            plt.tight_layout()
            
            canvas2 = FigureCanvasTkAgg(fig2, master=class_frame)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig2)
        
        # 3. Grade distribution density plot
        density_frame = ttk.LabelFrame(scrollable_frame, text="Grade Distribution Density Analysis", padding="10")
        density_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig3, ax3 = plt.subplots(1, 1, figsize=(12, 6))
        
        # Draw density curve for each subject
        colors = plt.cm.Set3(range(len(subject_columns)))
        for i, subject in enumerate(subject_columns):
            sns.kdeplot(data=self.data, x=subject, ax=ax3, color=colors[i], label=subject)
        
        ax3.set_title('Grade Distribution Density Plot by Subject', fontsize=14, fontweight='bold')
        ax3.set_xlabel('Score')
        ax3.set_ylabel('Density')
        ax3.legend()
        ax3.grid(True, alpha=0.3)
        plt.tight_layout()
        
        canvas3 = FigureCanvasTkAgg(fig3, master=density_frame)
        canvas3.draw()
        canvas3.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig3)
        
        # 4. Student grade radar chart (top 5 students)
        if "Total Score" in self.data.columns:
            radar_frame = ttk.LabelFrame(scrollable_frame, text="Top Students Grade Radar Chart", padding="10")
            radar_frame.pack(fill=tk.X, padx=10, pady=5)
            
            fig4, ax4 = plt.subplots(1, 1, figsize=(10, 10), subplot_kw=dict(projection='polar'))
            
            # Get top 5 students
            top5_students = self.data.nlargest(5, "Total Score")
            
            # Set radar chart angles
            angles = np.linspace(0, 2 * np.pi, len(subject_columns), endpoint=False)
            angles = np.concatenate((angles, [angles[0]]))  # Close the shape
            
            colors_radar = plt.cm.Set1(range(5))
            
            for i, (_, student) in enumerate(top5_students.iterrows()):
                values = [student[subject] for subject in subject_columns]
                values += [values[0]]  # Close the shape
                
                ax4.plot(angles, values, 'o-', linewidth=2, label=f"{student['Name']}", color=colors_radar[i])
                ax4.fill(angles, values, alpha=0.1, color=colors_radar[i])
            
            ax4.set_xticks(angles[:-1])
            ax4.set_xticklabels(subject_columns)
            ax4.set_title('Top 5 Students Grade Radar Chart', fontsize=14, fontweight='bold', pad=20)
            ax4.legend(loc='upper right', bbox_to_anchor=(1.2, 1.0))
            plt.tight_layout()
            
            canvas4 = FigureCanvasTkAgg(fig4, master=radar_frame)
            canvas4.draw()
            canvas4.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig4)
        
        # 5. Grade trend analysis (scatter plot matrix)
        scatter_frame = ttk.LabelFrame(scrollable_frame, text="Subject Grade Scatter Plot Matrix", padding="10")
        scatter_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Select main subjects for scatter plot analysis (avoid overly complex chart)
        main_subjects = subject_columns[:4] if len(subject_columns) > 4 else subject_columns
        
        if len(main_subjects) > 1:
            fig5, axes = plt.subplots(len(main_subjects), len(main_subjects), figsize=(12, 12))
            
            for i, subject1 in enumerate(main_subjects):
                for j, subject2 in enumerate(main_subjects):
                    ax = axes[i, j] if len(main_subjects) > 1 else axes
                    
                    if i == j:
                        # Diagonal shows histogram
                        ax.hist(self.data[subject1], bins=15, alpha=0.7, color='skyblue')
                        ax.set_title(f'{subject1} Distribution')
                    else:
                        # Off-diagonal shows scatter plot
                        ax.scatter(self.data[subject2], self.data[subject1], alpha=0.6, s=20)
                        
                        # Add trend line
                        z = np.polyfit(self.data[subject2], self.data[subject1], 1)
                        p = np.poly1d(z)
                        ax.plot(self.data[subject2], p(self.data[subject2]), "r--", alpha=0.8)
                    
                    if j == 0:
                        ax.set_ylabel(subject1)
                    if i == len(main_subjects) - 1:
                        ax.set_xlabel(subject2)
            
            plt.suptitle('Main Subject Grade Relationship Scatter Plot Matrix', fontsize=14, fontweight='bold')
            plt.tight_layout()
            
            canvas5 = FigureCanvasTkAgg(fig5, master=scatter_frame)
            canvas5.draw()
            canvas5.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig5)
        
        # Bind mouse wheel event
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Save first chart for export (for compatibility)
        if self.current_figures:
            self.current_fig = self.current_figures[0]
        
        # Switch to visualization tab
        self.notebook.select(self.visual_frame)

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentGradeAnalysisSystem(root)
    root.mainloop()