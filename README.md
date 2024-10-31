# Smart Attendance Generator


A streamlined web application that automates attendance sheet generation based on chosen subjects, student batches, and batch sizes, designed to handle diverse student data with flexible configurations.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Technologies Used](#technologies-used)

## Overview
The Attendance Sheet Generator allows users to upload Excel files containing student data divided by division and batch. With the app, users can filter students based on subjects, specify batch sizes, and automatically generate formatted Excel sheets with the filtered information organized by custom batch names. This tool simplifies creating structured attendance lists and arranging students in multiple sheets if batch size limits are reached.

## Features
- **Automated Attendance Sheet Creation**: Easily generate attendance sheets based on selected subjects, batch size, and year level.
- **Multi-Sheet Output for Large Batches**: Automatically divides students across multiple sheets if the batch size is exceeded.
- **Custom Batch Sorting**: Sorts student data by predefined batch order.
- **Custom Subject Mapping**: Recognizes both full names and short forms of subjects.
- **Excel Export**: Generates output in an organized, Excel-compatible format.

## Installation
To set up the Attendance Sheet Generator, follow these steps:

1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/attendance-sheet-generator.git
2. Navigate to the project directory:
   ```bash
   cd smart-attendance-generator
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
4. Run the application:
   ```bash
   python app.py

## Technologies Used:

- **Python**: Core programming language used for application logic.
- **Flask**: Lightweight web framework for handling the backend and user interactions.
- **Pandas**: Data manipulation library used to filter and organize student data.
- **OpenPyXL**: Library to read and write Excel files, enabling structured and formatted output.
- **HTML/CSS**: Basic frontend elements for the form interface in the web application.

