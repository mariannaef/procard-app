# PROCARD_PURCHASES  
**Completed:** 04/03/2026

---

## What Is It?

### Project Overview
PROCARD_PURCHASES is a `Streamlit` application built for the Athletic Business Office (ABO) to automate the monthly ProCard reconciliation process for athletic purchases. The app is designed to reduce manual formatting, improve reconciliation accuracy, and streamline the monthly processing of ProCard transactions across all athletic business purchases.

The primary users are staff within the Athletic Business Office responsible for reviewing and reconciling ProCard activity.

This tool takes a workflow export and a compressed folder of bank statements, processes and reconciles transaction data between both sources, and outputs a fully formatted File Feed ready for review and business use.

### Purpose
This project was created to streamline ProCard reconciliation by automating the cleanup, matching, and formatting of monthly athletic purchase transaction data.

Before this tool, reconciliation required manual review of workflow submissions against bank statement transactions, along with manual formatting of the final output. This app reduces repetitive work, improves consistency, and lowers the risk of manual reconciliation errors.

### Inputs
The app requires two monthly input files:

1. **Workflow Export**
   - **Source:** Internal workflow export  
   - **Format:** `.csv`  
   - **Contains:** Submitted purchase records, transaction details, FOAPAL information, and business office workflow data  

2. **Bank Statements**
   - **Source:** Regions bank statement exports  
   - **Format:** `.zip` containing statement `.pdf` files  
   - **Contains:** Monthly ProCard transaction records downloaded from Regions  

Both files must be uploaded into the Streamlit app before processing.

> [!IMPORTANT]
> The app assumes these files retain their expected formatting and structure. If source systems change file structure, naming conventions, or column names, the app may fail or produce incorrect output.

### Process
The app performs the following steps:

- Loads the workflow `.csv` export  
- Extracts and reads transaction data from the uploaded Regions bank statement `.pdf` files inside the `.zip` folder  
- Parses transaction data from the bank statements into a structured format  
- Matches bank transactions to workflow submissions  
- Excludes outdated or unmatched bank statement rows that should not be included in the current reconciliation cycle  
- Includes all valid matched transactions in the final dataset  
- Identifies and processes split transactions after matching  
- Splits matched rows where multiple workflow entries correspond to a single bank transaction  
- Checks for missing FOAPAL values in matched rows  
- Allows the user to manually fill in missing FOAPAL values before final output  
- Formats the final output into a File Feed structure used by the ABO  

This removes the need for manual statement review, reduces reconciliation time, and standardizes monthly ProCard processing.

### Deployment / Access
This project is currently deployed in two ways:

1. **Streamlit Community Cloud (Primary Use)**
   - Hosted through Streamlit Community Cloud  
   - Connected to GitHub for deployment  
   - Intended for routine monthly use  

2. **Local Host (Backup / Development)**
   - Can be run locally through VS Code  
   - Useful for testing, debugging, or making updates  

---

## Problems
Currently, there are no known problems.

If the Athletic Business Office identifies an issue, refer to the codebase for troubleshooting and make updates as needed.

---

## Current Status
This project is complete and functioning as intended for current business needs.  

As of **04/03/2026**, there are no known issues.

---

## Moving Forward
If the ABO needs this updated for any reason:

1. Make the necessary code changes  
2. Test locally in VS Code  
3. Push updates to GitHub  
4. Redeploy or reboot in Streamlit  

Future updates will most likely be needed if:

- Workflow export structure changes  
- Regions bank statement formatting changes  
- PDF parsing logic needs adjustment  
- FOAPAL validation requirements change  
- ABO reconciliation rules are updated  
- Additional formatting is required in the final File Feed  