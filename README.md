
# Asset Management System

A Streamlit-based asset management application designed to make organizational asset records easier to search, manage, update, and monitor.

The project transforms a traditional Excel-based asset register into an interactive web application while keeping Excel as the underlying data source.

## Features

- Asset overview dashboard
- Search assets by:
  - Serial Number
  - Barcode
  - User Accountable
  - Equipment Status
  - Description
- Case-insensitive search
- Detailed asset view
- View all asset information
- Add new assets
- Update existing assets
- Duplicate Serial Number detection
- Duplicate Barcode detection
- Asset status tracking:
  - IN USE
  - IN STORE
  - FAULTY
  - DISPOSED
  - LOST
- Export search results to CSV
- Excel-based data storage
- Preserves the original 24-field asset structure

## Technology Stack

- Python
- Streamlit
- Pandas
- OpenPyXL
- Excel

## Project Structure

```text
asset-management-system/
├── app.py
├── assets_sample.xlsx
├── requirements.txt
├── README.md
└── .gitignore
```

The real organizational asset file (`assets.xlsx`) is excluded from the repository to protect sensitive information.

`assets_sample.xlsx` contains fictional demonstration data and can be used for testing and demonstration purposes.

## Dataset and Column Structure

The column names used in this project are based on the structure of the original organizational asset register that inspired the application.

Some field names and terminology are therefore organization-specific and may differ from those typically used in a standard asset management system.

The public sample dataset preserves this structure to demonstrate how the application can work with an existing organizational asset register. All records included in the sample dataset are fictional and contain no real organizational or personal data.

The application can be adapted to different asset schemas, field names, and organizational requirements.

## Installation

Clone the repository and install the required packages:

```bash
pip install -r requirements.txt
```

## Running the Application

Run the Streamlit application with:

```bash
streamlit run app.py
```

## Data Privacy

The original organizational asset register is not included in this repository.

The real data file is excluded through `.gitignore`, and the publicly available sample Excel file contains fictional demonstration data only.

No real personal, organizational, asset, or internal operational data is included in the public repository.

## Future Improvements

Planned improvements include:

- Advanced asset filters
- Dashboard charts and analytics
- Asset verification alerts
- Improved reporting
- AI-powered natural language asset search
- Database integration for larger deployments

## Purpose

This project demonstrates how an existing Excel-based asset management workflow can be transformed into a more searchable, interactive, and user-friendly application using Python and Streamlit.

It also demonstrates how an application can be designed around an existing organizational data structure while remaining adaptable to different asset management requirements.
