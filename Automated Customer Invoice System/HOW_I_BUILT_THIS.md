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
* Set up =VLOOKUP using the Cust_list list to pull customer details into the invoice form when selecting a customer.
* Implemented =IFNA Formula to hide unnecessary error message and ensure cleaner data presentation.

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

## Step 2: Implementing Invoice Tracking

### 2.1 Creating the Tracker Sheet
- Developed a dedicated tracker sheet to monitor issued invoices and their statuses.
- The sheet includes the following columns:
  - **Invoice Number**: Unique identifier for each invoice.
  - **Customer Name**: Customer associated with the invoice.
  - **Issue Date**: Date when the invoice was created.
  - **Due Date**: Date by which payment is expected.
  - **Status**: Current status of the invoice (e.g., "Paid", "Overdue", "Pending").

### 2.2 Conditional Formatting for Overdue Invoices
- Applied conditional formatting to highlight overdue invoices for prompt attention.
- Used Excel’s `TODAY()` function to compare the **Due Date** with the current date. If the due date has passed and the invoice status is not "Paid", the cell is highlighted in a specified color (e.g., red) to indicate that the invoice is overdue.
- This allows for immediate identification of overdue invoices, helping prioritize follow-ups with customers.

### 2.3 Tracking Payment Status
- Added a **Status** column to monitor whether invoices are paid, overdue, or pending.
- Set up manual or automatic updates to change the invoice status, depending on payment or due date.

### 2.4 Logging Invoice Email Details
- Implemented a system to track when and to whom the invoice was emailed, including the date and time of sending.
- This ensures a record is maintained of all communications related to the invoice.

---

## Step 3: Implementing VBA Macros for Automation

### 3.1 Automated Invoice Creation
* Created a button to generate a new invoice with an automatically assigned invoice number.
* Integrated logic to clear previous details and pull customer data when a customer is selected.

**VBA Code:**
```vba
Sub CreateNewInvoice()

    ' Declare a variable "invno" to store the current invoice number as a long integer
    Dim invno As Long

    ' Assign the value from cell C3 in Sheet1 to the variable "invno"
    invno = Sheet1.Range("C3")

    ' Clear the contents of the specified ranges: C4:D4, B10, and B19:G32
    ' This is likely to reset parts of the invoice for new data
    Range("C4:D4,B10,B19:G32").ClearContents

    ' Show a message box to inform the user of the next invoice number (current invoice number + 1)
    MsgBox "Your next invoice number is " & invno + 1

    ' Update cell C3 with the next invoice number (current invoice number + 1)
    Range("C3") = invno + 1

    ' Move the selection to cell B10, likely for the user to input the next set of data
    Range("B10").Select

    ' Save the workbook after making changes
    ThisWorkbook.Save

End Sub


```

### 3.2 Recording and Saving Invoices
* Developed a macro to record invoices into the **Invoice Record Sheet**, keeping a log of all generated invoices.

**VBA Code:**
```vba
Sub RecordofInVoice()

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

### 3.4 Exporting Invoice as Excel
* Implemented a VBA macro to export invoices as Excel files with a single click, naming them systematically using invoice numbers and customer names.
* Created a VBA macro to export invoices as xlsx's with systematic naming and saved them to a specified location.


**VBA Code:**
```vba
Sub SaveInvAsExcel()

    ' Declare variables for invoice details
    Dim invno As Long
    Dim custname As String
    Dim amt As Currency
    Dim dt_issue As Date
    Dim term As Byte
    Dim path As String
    Dim fname As String
    Dim nextrec As Range

    ' Assign values from the worksheet to the variables
    invno = Range("C3")              ' Invoice number from cell C3
    custname = Range("B10")          ' Customer name from cell B10
    amt = Range("H38")               ' Amount from cell H38
    dt_issue = Range("C5")           ' Date of issue from cell C5
    term = Range("C6")               ' Terms (e.g., payment terms) from cell C6
    path = "C:\Users\Ivan\Desktop\Invoice\"  ' Folder path to save the invoice
    fname = invno & " - " & custname   ' Filename combining invoice number and customer name

    ' Copy the invoice template (Sheet1) to a new workbook
    Sheet1.Copy

    ' Delete all shapes (buttons, images, etc.) on the copied sheet
    Dim Shpe As Shape
    For Each Shpe In ActiveSheet.Shapes
        Shpe.Delete   ' Delete any shape (button, picture, etc.)
    Next Shpe

    ' Delete all non-picture shapes from the copied sheet
    For Each Shpe In ActiveSheet.Shapes
        If Shpe.Type <> msoPicture Then Shpe.Delete  ' Only delete non-picture shapes (like buttons)
    Next Shpe

    ' Save the newly copied workbook (without shapes) as a new invoice
    With ActiveWorkbook
        .Sheets(1).Name = "Invoice"  ' Name the first sheet as "Invoice"
        .SaveAs Filename:=path & fname, FileFormat:=51  ' Save as an Excel file (.xlsx)
        .Close  ' Close the new workbook after saving
    End With

    ' Input the details of the invoice into the "Invoices" record sheet (Sheet2)
    Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)  ' Find the next available row in Sheet2

    ' Input invoice details into the new row in Sheet2
    nextrec = invno  ' Place the invoice number in column A
    nextrec.Offset(0, 1) = custname  ' Place the customer name in column B
    nextrec.Offset(0, 2) = amt  ' Place the amount in column C
    nextrec.Offset(0, 3) = dt_issue  ' Place the issue date in column D
    nextrec.Offset(0, 4) = dt_issue + term  ' Place the due date (issue date + terms) in column E

    ' Add a hyperlink to the new invoice file in column H of the same row
    Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 7), Address:=path & fname & ".xlsx"  ' Link to the saved invoice file

End Sub

```


### 3.5 Exporting Invoice as PDF
* Implemented a VBA macro to export invoices as PDFs files with a single click, naming them systematically using invoice numbers and customer names.
* Created a VBA macro to export invoices as PDFs with systematic naming and saved them to a specified location.

**VBA Code:**
```vba

 ' Same as Excel Code instead we add extra line of code and modify Excel code by adding an extra line and adjusting the offset from (0,7) to (0,6) to accommodate the different cell categories in the invoice records
Sub SaveasPDF()

Dim invno As Long
Dim custname As String
Dim amt As Currency
Dim dt_issue As Date
Dim term As Byte
Dim path As String
Dim fname As String
Dim nextrec As Range


invno = Range("C3")
custname = Range("B10")
amt = Range("H38")
dt_issue = Range("C5")
term = Range("C6")
path = "C:\Users\Ivan\Desktop\Invoice\"
fname = invno & " - " & custname

' Export the active sheet as a PDF, respecting print areas, and save it to the specified path with the generated filename.
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, Filename:=path & fname

Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = invno
nextrec.Offset(0, 1) = custname
nextrec.Offset(0, 2) = amt
nextrec.Offset(0, 3) = dt_issue
nextrec.Offset(0, 4) = dt_issue + term

 ' Export the active sheet as a PDF to the specified path
Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 6), Address:=path & fname & ".pdf."

End Sub
```

### 3.6 Email Invoice as PDF
* Used VBA macros to automatically attach the generated PDF to an email and configured Outlook to open a new email draft with predefined subject and message for consistency.

**VBA Code:**
```vba
Sub EmailasPDF()

    ' Create an Outlook application object
    Dim EApp As Object
    Set EApp = CreateObject("Outlook.Application")

    ' Create a new email item within Outlook
    Dim EItem As Object
    Set EItem = EApp.CreateItem(0)  ' 0 corresponds to an email item

    ' Declare variables for invoice details
    Dim invno As Long
    Dim custname As String
    Dim amt As Currency
    Dim dt_issue As Date
    Dim term As Byte
    Dim path As String
    Dim fname As String
    Dim nextrec As Range

    ' Assign values from the worksheet to the variables
    invno = Range("C3")              ' Invoice number from cell C3
    custname = Range("B10")          ' Customer name from cell B10
    amt = Range("H38")               ' Amount from cell H38
    dt_issue = Range("C5")           ' Date of issue from cell C5
    term = Range("C6")               ' Terms (e.g., payment terms) from cell C6
    path = "C:\Users\Ivan\Desktop\Invoice\"  ' Folder path to save the invoice
    fname = invno & " - " & custname   ' Filename combining invoice number and customer name

    ' Export the active sheet as a PDF to the specified path
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, Filename:=path & fname

    ' Find the next available row in Sheet2 to record invoice details
    Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

    ' Record invoice details in the next available row of Sheet2
    nextrec = invno  ' Place the invoice number in column A
    nextrec.Offset(0, 1) = custname  ' Place the customer name in column B
    nextrec.Offset(0, 2) = amt  ' Place the amount in column C
    nextrec.Offset(0, 3) = dt_issue  ' Place the issue date in column D
    nextrec.Offset(0, 4) = dt_issue + term  ' Place the due date (issue date + terms) in column E
    nextrec.Offset(0, 8) = Now  ' Record the current date and time in column I

    ' Add a hyperlink to the saved PDF invoice file in column G of the same row
    Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 6), Address:=path & fname & ".pdf."

    ' Prepare and send the email with the PDF attachment
    With EItem
        .To = Range("B16")  ' The recipient's email address from cell B16
        .Subject = "Invoice no: " & invno  ' Subject line includes the invoice number
        .Body = "Please find invoice attached."  ' Email body message
        .Attachments.Add (path & fname & ".pdf")  ' Attach the PDF invoice
        .Display  ' Display the email to the user (ready for manual sending)
    End With

End Sub

```

---



---


