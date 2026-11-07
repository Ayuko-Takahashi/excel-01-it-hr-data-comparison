# Excel IT & HR Data Comparison

An Excel VBA automation tool that compares employee data between IT and HR datasets, identifies discrepancies, and generates an exception-based report for review.

---

## 🚀 Project Overview

This project automates the comparison of employee information maintained separately by **IT and HR**.

The **IT user list is used as the master dataset**, and employee records are matched using the **email address / User Principal Name (UPN)** as the unique identifier.

The VBA macro compares employee position titles, identifies missing users and data discrepancies, and generates a **Comparison Report** containing only records that require review.

The tool identifies issues without automatically modifying the source data, allowing discrepancies to be verified before corrections are made.

---

## 🧩 Business Problem

IT and HR may maintain employee information in separate systems or data exports. Over time, differences can occur between the two datasets, such as:

- Different position titles
- Formatting inconsistencies
- Employees missing from the HR dataset
- Employees appearing in HR but not in the IT master dataset

Manually comparing these records can be time-consuming and makes discrepancies easy to overlook.

This project provides a reusable Excel-based comparison process that highlights exceptions and allows users to focus only on records requiring investigation.

---

## ⚙️ Workflow Process

1. Open the macro-enabled Excel comparison tool (`.xlsm`).
2. Click the **Compare IT & HR** button.
3. Select the **IT user list** as the master dataset.
4. Select the **HR user list** to compare.
5. The VBA macro matches employees using their **email address / UPN**.
6. Position titles are compared between matched records.
7. A reverse check identifies users who exist in HR but not in IT.
8. Detected issues are highlighted in the HR dataset.
9. A **Comparison Report** is generated containing only records requiring review.
10. A summary displays the comparison results.

---

## 📊 Key Features

- Automated comparison of two CSV datasets using **Excel VBA**
- Email / UPN-based employee matching
- Position title comparison
- Detection of exact matches
- Detection of formatting differences
- Detection of genuine title mismatches
- Detection of users missing from HR
- Reverse comparison to identify HR users not found in IT
- 🟡 Yellow highlighting for formatting differences
- 🔴 Pink highlighting for mismatches and missing records
- Exception-only **Comparison Report**
- Automatic summary of comparison results
- Automatic column resizing
- Frozen report headers for easier navigation
- Simple **Compare IT & HR** button for end users
- No automatic correction of source data, allowing human verification before changes are made

---

## 🎯 Skills Demonstrated

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Data comparison and reconciliation
- Data validation
- CSV file processing
- Conditional logic
- Record matching using unique identifiers
- Exception handling and reporting
- Excel report automation
- User-friendly automation design
- Business process improvement
