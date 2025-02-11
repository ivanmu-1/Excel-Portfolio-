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

## Step 3: Implementing VBA Macros for Automation

### 3.1 Automated Invoice Creation
* Created a button to generate a new invoice with an automatically assigned invoice number.
* Clear Previous Invoice Details in order to 
* Integrated logic to pull customer details from the **Customer Detail Sheet** when selecting a customer.

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
* Enabled users to save invoices as Excel files for future reference.

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
* Implemented a VBA macro to export invoices as Excels with a single click.
* Systematically named the PDF using the **invoice number** and **customer's name**.

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


### 3.4 Exporting Invoice as PDF
* Implemented a VBA macro to export invoices as PDFs with a single click.
* Systematically named the PDF using the **invoice number** and **customer's name**.

**VBA Code:**
```vba

 ' Same as Excel Code instead we change FixedFormat = PDF
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

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, Filename:=path & fname

Set nextrec = Sheet2.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = invno
nextrec.Offset(0, 1) = custname
nextrec.Offset(0, 2) = amt
nextrec.Offset(0, 3) = dt_issue
nextrec.Offset(0, 4) = dt_issue + term

' Add a hyperlink to the new invoice file in column H of the same row
Sheet2.Hyperlinks.Add anchor:=nextrec.Offset(0, 6), Address:=path & fname & ".pdf."

End Sub
```

### 3.5 Email Invoice as PDF
* Used VBA macros to automatically attach the generated PDF to an email.
* Configured Outlook to open a new email draft with the invoice attached.
* Included a predefined subject and message for consistency.

**VBA Code:**
```vba
' Code will be added here
```

---



---

This document serves as a guide to understanding how the project was structured and automated. The blank sections will be filled with VBA code as the project evolves.

