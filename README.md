# Legacy Project – Automated Weekly Cost Reports (2020)

## Overview

This is the first automation project I developed using Python, back in 2020.

The goal was to eliminate a fully manual and repetitive process of generating weekly Excel reports for purchased items and supplier cost analysis.

Although simple by today’s standards, this project marks the starting point of my transition into software development and automation.

## Problem Context

At the time, the process required:
- Manually identifying missing weekly reports
- Collecting multiple .txt files as data sources
- Extracting and consolidating manually data into Excel (partially supported by macros)
- Performing cost calculations and currency conversions
- Standardizing supplier information
- Creating new and mission standard weekly repots

This process was time-consuming, error-prone, and not scalable.

## Solution
I developed a Python script that fully automated the workflow:
- Scans historical directories to identify missing weekly reports
- Reads and processes raw .txt input files
- Extracts and transforms relevant data
- Performs:
  - Cost calculations
  - Volume aggregation
  - Currency conversion
  - Grouping data from different company sites
- Standardizes supplier data
- Generates structured Excel reports (.xlsx) using a predefined template

## Key Features
- Automated gap detection (missing weekly reports)
- Data parsing from semi-structured .txt files
- Business rule implementation for cost calculations
- Excel generation with formatting and structure
- Supplier data normalization
- Batch processing for multiple weeks

## Technologies Used
- Python 3.x
- openpyxl
- File system automation (os, shutil)
- CLI-based execution

## Impact
- Eliminated repetitive manual work
- Reduced report generation time from ~4 hours to ~5 minutes
- Improved consistency and reliability of data
- First step toward larger automation and cost optimization initiatives
- Opened the way to make the process fully automated, running weekly without human intervention

## Limitations (Legacy Context)
- Hardcoded file paths (Windows environment)
- No modular structure
- No error handling standardization
- No interface (CLI only)

## Evolution

This project was the foundation for more advanced solutions involving:
- Data processing pipelines
- Automated cost analysis systems
- Cost reduction and cost avoidance analysis based on high quality processed data

## Note

This repository is kept as a legacy reference, demonstrating a continuous pursuit of new skills and solutions through a problem-solving mindset since the beginning of my journey as a developer 6 years ago.

## .txt file example (anonymized)
![img.png](img.png)

## .xlsx file exlample (anonymized)
![img_1.png](img_1.png)