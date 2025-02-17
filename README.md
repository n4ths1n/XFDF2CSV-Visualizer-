# XFDF2CSV Visualizer Documentation

## Overview
XFDF2CSV Visualizer is a Python-based application that allows users to convert XFDF files into CSV format and visualize the data using different graphical representations such as bar charts, heatmaps, pie charts, and network graphs.

## Features
- Convert XFDF files to CSV format
- Load and visualize CSV data
- Various visualization options: Bar charts, Heatmaps, Pie charts, Line charts, and Network graphs
- Interactive zooming for network graphs
- User-friendly GUI using Tkinter
- Error handling with pop-up notifications
- Support for batch processing of XFDF files

## Requirements
Before running the application, ensure you have the following dependencies installed:

```bash
pip install tkinter pandas matplotlib seaborn networkx
```

## Usage

### Running the Application
Run the following command to start the application:
```bash
python main.py
```

### Converting XFDF to CSV
1. Click on the "Convert XFDF to CSV" button.
2. Select the folder containing XFDF files.
3. Choose a location to save the output CSV file.
4. Wait for the process to complete and check the saved file.

### Loading a CSV File
1. Click on the "Load CSV" button.
2. Select the CSV file you want to analyze.
3. The data will be displayed for further visualization.

### Selecting a Question for Visualization
1. Use the dropdown menu to select a question.
2. Choose from the following options:
   - Department
   - Q1: Preferred departments to work with
   - Q2: Preferred colleagues outside the department
   - Q3: Departments you prefer not to work with
   - Q4: Departments that might not find value in working with you

### Choosing a Visualization Type
Click on the corresponding button to switch visualization types:
- **Bar Chart**: Displays categorical data distributions
- **Heatmap**: Shows correlations or frequency data
- **Pie Chart**: Represents proportions
- **Line Chart**: Analyzes trends over time
- **Network Graph**: Visualizes relationships and connections

## Zooming in Network Graphs
- **Scroll up** to zoom in.
- **Scroll down** to zoom out.
- **Click and drag** to move around the graph.

## Error Handling
If an error occurs, an error message will be displayed in a pop-up window.

## License
This project is open-source and available for modification and distribution under the MIT License.

