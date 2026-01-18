Here is the **English version** of your Student Performance Analysis System README:

# Student Performance Analysis System

A comprehensive student performance analysis system built with Python and Tkinter, supporting data generation, import, statistical analysis, visualization, and report generation.

## 🚀 Main Features

### 📊 Data Management

- **Data Import**: Supports Excel (.xlsx, .xls) and CSV format score files
- **Data Generation**: Built-in student data generator with customizable test data
- **Data Export**: Export data and analysis results in multiple formats

### 📈 Statistical Analysis

- **Basic Statistics**: Average score, highest score, lowest score, pass rate, excellence rate per subject
- **Subject Comparison**: Subject average score comparison chart + box plots of score distribution
- **Score Distribution Analysis**: Total score histogram + percentage analysis by score segments per subject
- **Advanced Analysis**: Correlation analysis, class comparison, density plots, radar charts, scatter plot matrix

### 📋 Visualization Features

- **Various Chart Types**: Bar, box, histogram, pie, heatmap, radar, scatter, etc.
- **Interactive Interface**: Intuitive Tkinter-based GUI
- **Real-time Preview**: Analysis results displayed instantly in the visualization panel

### 📄 Report Generation

- **PDF Report**: Automatically generates comprehensive report with charts and statistics
- **Chinese Support**: Full support for Chinese character display
- **Multi-format Export**: Charts can be exported as PNG, JPG, PDF, etc.

## 🛠️ Tech Stack

- **Python** 3.7+
- **GUI Framework**: Tkinter
- **Data Processing**: Pandas, NumPy
- **Visualization**: Matplotlib, Seaborn
- **PDF Generation**: ReportLab
- **Others**: Random (for data generation)

## 📦 Installation

```bash
pip install pandas numpy matplotlib seaborn reportlab
```

## 🎯 Quick Start

### 1. Run the program

```bash
python integrated_system.py
```

### 2. Prepare Data

#### Option A: Import Existing Data

- Menu → File → Import Data
- Select Excel or CSV score file
- File should contain: student info (name, ID, class, etc.) + subject scores

#### Option B: Generate Test Data

- Menu → Tools → Student Data Generator
- Customizable settings:
  - Number of students (1–1000)
  - Student ID prefix
  - Class range
  - Subjects and score ranges
  - Pass mark & expected pass rate
- Generated data can be exported or used directly for analysis

### 3. Analyze Data

#### Basic Statistics

- Click "Basic Statistical Analysis"
- View detailed stats per subject
- Auto-calculates total score, average, and ranking

#### Subject Comparison

- Click "Subject Comparison Analysis"
- View average score comparison bar chart
- View score distribution box plots per subject

#### Score Distribution

- Click "Score Distribution Analysis"
- View total score histogram
- Analyze percentage in each score segment per subject

#### Advanced Analysis

- Click "Advanced Analysis"
- Subject correlation heatmap
- Class-level comparison
- Score density plots
- Top students radar chart
- Subject scores scatter plot matrix

### 4. Export Results

#### Export Analysis Results

- Click "Export Analysis Results"
- Choose content (statistics / charts)
- Select format (Excel / CSV / PNG / JPG / PDF)

#### Generate PDF Report

- Click "Generate PDF Report"
- Automatically creates full report with charts, stats, and conclusions
- Supports Chinese characters and includes timestamp

## 📁 Project Structure

```text
Student-Performance-Analysis-System/
├── integrated_system.py          # Main program
├── SimSun.ttf                    # Chinese font file (for PDF & display)
├── README.md                     # This file
├── LICENSE                       # License file
└── sample_files/
    ├── student_scores_*.xlsx     # Generated sample data
    ├── analysis_charts_*.png     # Exported charts
    └── student_performance_report.pdf   # Generated PDF report
```

## 🎨 Key Highlights

### Data Generator Features

- Realistic Chinese name generation (common surnames + given names)
- Flexible class assignment
- Controlled score distribution (pass rate, score range)
- Supports up to 13 preset subjects (customizable)

### Analysis Features

- Multi-dimensional stats: mean, median, std, quartiles, etc.
- Smart full-mark detection for different subjects
- Automatic total score, ranking, and average calculation
- Robust handling of missing values & outliers

### Visualization Features

- Responsive scrollable view
- Professional color schemes
- Good label rotation & readability
- High-resolution export (300 DPI)

### Report Generation Features

- Chinese font optimization (uses SimSun.ttf when available)
- Smart layout with auto-pagination
- Includes stats tables, charts, and brief conclusions
- Generation timestamp included

## 🔧 Configuration Notes

### Chinese Font Support

1. Place `SimSun.ttf` in the project folder
2. The program auto-detects and prefers local font
3. Also supports system fonts on Windows

### Data Format Requirements

- **Required column**: Name (姓名)
- **Optional columns**: Student ID, Class, Index
- **Subject columns**: Any subject names
- **Encoding**: UTF-8 recommended for Chinese characters

## 🐛 FAQ

**Q: Chinese characters appear as boxes/gibberish?**  
A: Make sure `SimSun.ttf` is in the project directory.

**Q: Import data failed?**  
A: Check that the file has a "Name" column and uses UTF-8 encoding.

**Q: Charts are cut off / not fully visible?**  
A: Use mouse wheel to scroll or resize the window.

**Q: PDF report generation failed?**  
A: Ensure `reportlab` is installed and you have write permission / enough disk space.

## 📄 License

MIT License — see the [LICENSE](LICENSE) file for details.

## 👥 Contributing

Feel free to open Issues or submit Pull Requests!

## 📧 Contact

Questions, suggestions, or bugs? Please open an Issue.

---

**Student Performance Analysis System v2.0**  
Making data analysis easier — helping education decisions become smarter!

Good luck with your project!