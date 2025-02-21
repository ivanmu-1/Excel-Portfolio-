# Automated Customer Invoice System (Excel + VBA Macros)

<img width="728" alt="Image" src="https://github.com/user-attachments/assets/e53b1ec8-c816-4984-90cb-cb3fc84b1083" />

## Overview
This project is an **Automated Customer Invoice System** built in **Microsoft Excel** using **VBA Macros**. It simplifies the process of creating, saving, and managing invoices with just a click of a button. The system automates multiple tasks, including:

- **Generating new invoices**
- **Recording invoices** in a dedicated sheet
- **Saving invoices** as an Excel file
- **Exporting invoices** as a PDF
- **Sending invoices via email** as a PDF attachment

## Features
### **1. Automated Invoice Creation**
- A button generates a **new invoice** with an automatically assigned invoice number.
- Pulls customer details from a dedicated sheet when selecting a customer.

### **2. Record and Save Invoices**
- The system allows users to **record invoices** into a log sheet, keeping a history of all generated invoices.
- Users can **save invoices** as Excel files for future reference.

### **3. Export as PDF/Excel**
- Invoices can be exported and saved as **PDF/Excel files** with a single click.
- PDFs and Excel Files are named systematically using the invoice number and customer name.

### **4. Email Invoice as PDF**
- Uses **VBA Macros** to automatically attach the generated PDF to an email.
- Opens an email draft in **Outlook** with the invoice attached and a predefined subject & message.

### **5. Customer and Invoice Records**
- A **Customer Details Sheet** stores all customer information (e.g., Name, Contact, Email, Address, etc.).

<img width="718" alt="Image" src="https://github.com/user-attachments/assets/1c13b8f2-400e-4eb2-97ed-4ce148a59e01" />

- An **Invoice Record Sheet** logs all issued invoices, including amounts, dates, and customer details.

<img width="1079" alt="Image" src="https://github.com/user-attachments/assets/60387809-7266-49e7-9720-0e0b1e80f629" />

### **6. File Saving Records**
- A **A file-saving system** that stores invoices as either Excel or PDF files, using the designated name and invoice number. The system allows users to select a specific file path on their PC for saving invoices efficiently.

insert image here

## How It Works
1. **Open the Excel file** containing the automated invoice system.
2. **Use the buttons** to perform actions:
   - **"New Invoice"** → Generates a blank invoice with a unique number.
   - **"Record Invoice"** → Saves the invoice data in the record sheet.
   - **"Save as Excel"** → Saves the invoice as a standalone Excel file.
   - **"Save as PDF"** → Converts the invoice into a PDF file.
   - **"Email as PDF"** → Automatically attaches the PDF to an Outlook email.
3. **Customer data is auto-filled** from the customer details sheet when selecting a customer.
4. **Review the Invoice Record Sheet** for past invoices and customer transactions.

## Technologies Used
- **Microsoft Excel** for data management and interface.
- **VBA (Visual Basic for Applications)** for automation and macros.
- **Outlook Integration** for email functionality.

## Setup Instructions
1. **Enable Macros** in Excel:
   - Go to **File > Options > Trust Center > Trust Center Settings**.
   - Select **Macro Settings** and enable macros.
2. **Ensure Outlook is Installed & Configured** if using the email feature.
3. **Modify the VBA Code** (if necessary) to adjust email templates or file-saving paths.




