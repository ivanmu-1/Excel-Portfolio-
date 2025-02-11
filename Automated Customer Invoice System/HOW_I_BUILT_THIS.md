# How I Built This Project

## Overview
This document outlines the steps taken to build the Excel-based invoice management system, including setting up the sheets, implementing data validation, and automating processes using VBA macros.

---

## Step 1: Setting Up the Excel Sheets

### 1.1 Creating the Invoice Form
* Designed the main invoice form layout.
* Included Key sections:
  - Invoice Number & PO Number: Auto-increments for new invoices
  - Date & Payment Terms: Predefined Payment terms of 30,60, and 90 days
  - Invoice To: Automatically retreives customer details from the customer database
  - Description Table: List items, quantites, unit prices, VAT, and totals
  - Company & Payment Detail: Company Name, VAT number, Contact Information, etc
  
* Used data validation to create predefined terms such as **30, 60, 90 days** for payment terms.
* Used =VLookUp to automatically retreives customer details from the customer database

### 1.2 Creating the Customer Detail Sheet
* Created a dedicated sheet to store customer details (**name, address, city, state, zip code, phone-number, and email.**).
* Utilized Name Manager to define a structured list named Cust_list for data retrieval using VLookUp 

### 1.3 Automating the Invoice Section
* Set up =VLOOKUP using the Cust_listto pull customer details into the invoice form when selecting a customer.
* Implemented =IFNA Formula to hide unnecessary error message nad ensure cleaner data presentation.

### 1.4 Logging Issued Invoices
* Created an **Invoice Record Sheet** to track:
  * Invoice Number
  * Company
  * Amounts
  * Date Issues / Date Dues
  * If Payment was made
  * Information regarding Invoice file format
  * Information regarding Invoice emailed 
* This serves as a historical record for all transactions.

---

## Step 2: Implementing VBA Macros for Automation

### 2.1 Automated Invoice Creation
* Created a button to generate a new invoice with an automatically assigned invoice number.
* Integrated logic to pull customer details from the **Customer Detail Sheet** when selecting a customer.

**VBA Code:**
```vba
' Sub RecordofInVoice()

    ' Declare variables to store invoice details
    Dim invno As Integer       ' Stores the invoice number
    Dim custname As String     ' Stores the customer's name
    Dim amt As Currency        ' Stores the total invoice amount
    Dim dt_issue As Date       ' Stores the issue date of the invoice
    Dim term As Byte           ' Stores the payment terms in days
    Dim nextrec As Range       ' Stores the next available row in the record sheet

    ' Retrieve values from the invoice sheet (Sheet1)
    invno = Sheet1.Range("C3")      ' Invoice Number from cell C3
    custname = Sheet1.Range("B10")  ' Customer Name from cell B10
    amt = Sheet1.Range("H38")       ' Total Invoice Amount from cell H38
    dt_issue = Sheet1.Range("C5")   ' Invoice Issue Date from cell C5
    term = Sheet1.Range("C6")       ' Payment Terms (e.g., 30, 60, 90 days) from cell C6

    ' Find the next available row in the Invoice Record sheet (Sheet2)
    ' It searches column A (Invoice Numbers) for the last used row and moves one row down
    Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

    ' Record invoice details into the next available row
    nextrec = invno                    ' Store Invoice Number in column A
    nextrec.Offset(0, 1) = custname     ' Store Customer Name in column B
    nextrec.Offset(0, 2) = amt          ' Store Invoice Amount in column C
    nextrec.Offset(0, 3) = dt_issue     ' Store Invoice Issue Date in column D
    nextrec.Offset(0, 4) = dt_issue + term  ' Store Due Date (Issue Date + Payment Terms) in column E

End Sub

```

### 2.2 Recording and Saving Invoices
* Developed a macro to record invoices into the **Invoice Record Sheet**, keeping a log of all generated invoices.
* Enabled users to save invoices as Excel files for future reference.

**VBA Code:**
```vba
' Code will be added here
```

### 2.3 Exporting Invoice as PDF
* Implemented a VBA macro to export invoices as PDFs with a single click.
* Systematically named the PDF using the **invoice number** and **customer's name**.

**VBA Code:**
```vba
' Code will be added here
```

### 2.4 Email Invoice as PDF
* Used VBA macros to automatically attach the generated PDF to an email.
* Configured Outlook to open a new email draft with the invoice attached.
* Included a predefined subject and message for consistency.

**VBA Code:**
```vba
' Code will be added here
```

---

## Next Steps
* Add error handling to VBA macros for robustness.
* Improve user interface for better usability.
* Expand functionalities based on feedback.

---

This document serves as a guide to understanding how the project was structured and automated. The blank sections will be filled with VBA code as the project evolves.

