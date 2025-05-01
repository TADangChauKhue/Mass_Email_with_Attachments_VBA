# Mass_Email_With_Attachments_VBA

Automate the generation and sending of personalized rent receipts to multiple tenants using Excel and VBA.

## Introduction

This Excel-VBA project was developed to help landlords and property managers save time and avoid repetitive manual tasks by automating two essential processes:

- **Generating rent receipts in PDF format**
- **Sending personalized emails to tenants with the receipt as an attachment**

The tool uses tenant data and payment status to create PDFs and send emails via Outlook, based on one-click macros integrated into the Excel file.

---

## 1. Business Problem

Each month, rent receipts need to be issued, archived, and emailed. Manually performing these actions is time-consuming and error-prone. This tool automates:

- Dynamic rent receipt generation based on payment status
- PDF export and archiving by tenant name
- Personalized email sending with automatic file attachment

---

## 2. Dataset

The Excel file contains two sheets:

### ➤ `MODELE_QUITTANCE` (Receipt Template)

- A formatted rent receipt automatically filled with tenant information
- Drop-down menu to select the **rental period** (e.g. May 2025)
- Two buttons:
  - **Exporter PDF**: Generates PDFs for tenants who have a payment date
  - **Envoyer Mail**: Sends emails with the correct receipt attached

### ➤ `DONNEES_LOCATAIRES` (Tenant Information)

- Contains all relevant tenant details:
![image](https://github.com/user-attachments/assets/856526f3-857a-454c-b6b3-66bb2c27b60b)


- Only tenants **with a payment date** will be processed.

- The column `Mail template` contains HTML formatting for the message body:

```html
<p>Bonjour [Nom]</p>
<p>Ci-jointe votre quittance de loyer.</p>
<p>Cordialement</p>
<p>Lanlord name</p>


## Method: VBA Automation

The tool includes two main macros:

1. **Export PDF Macro**  
   - Loops through each tenant  
   - Populates the receipt template  
   - Exports the receipt as a PDF  
   - Saves it in a folder named `Quittances`

2. **Send Mail Macro**  
   - Uses Microsoft Outlook via VBA  
   - Sends a personalized email to each tenant  
   - Attaches the correct receipt PDF  
   - Includes dynamic content (name, period, etc.) in the message

## Instructions

1. Open the Excel file and enable macros.
2. Enter tenant payment data in `DONNEES_LOCATAIRES`.
3. Click **Exporter PDF** in `MODELE_QUITTANCE` to create receipts.
4. Click **Envoyer Mail** to send emails to all tenants with their receipt.

## Requirements

- Microsoft Excel with macros enabled
- Microsoft Outlook (configured on your computer)
- A local folder named `Quittances` to store PDF files

## Benefits

- Save time with bulk PDF and email automation
- Avoid repetitive manual tasks
- Ensure consistency in formatting and communication
- Easily archive rent receipts in one place
