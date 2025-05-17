# Logistic Regression Predictor

A desktop application built with Python and Tkinter for performing logistic regression analysis on diverse datasets. The app supports loading data from various file formats and SQL databases, provides built-in statistical tests, prediction capabilities, and multi-language support.

---

## Overview

This application allows you to:

- Load data from Excel, CSV, JSON, Parquet files, or SQL databases (MySQL, PostgreSQL, SQLite, MSSQL).
- Select independent features and the target variable interactively.
- Handle missing data with options to drop rows or fill missing values with the mean.
- Perform essential statistical tests, including the Hosmer-Lemeshow goodness-of-fit test.
- Build and run logistic regression models for prediction and evaluation.
- Save analysis results as Excel or Word documents.
- Print results directly from the application.
- Switch between English and Arabic UI languages.
- Use a clean, user-friendly interface with undo, copy, and paste functionalities.

---

## Installation and Requirements

Make sure you have Python installed (version 3.7+ recommended). Install the required libraries using:

```bash
pip install pandas scikit-learn statsmodels scipy openpyxl python-docx sqlalchemy pymysql psycopg2-binary pyodbc
