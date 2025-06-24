# extractor.py
# This module contains the logic for PDF text extraction and communicating with the Gemini API.
# LOGIC UPDATED to process a comprehensive 85-field structure.

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import os
import sys
import json
import requests # A robust library for making API calls

# --- CRITICAL: Tesseract Path Configuration ---
def get_tesseract_path():
    """Returns the path to the Tesseract executable."""
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, 'Tesseract-OCR', 'tesseract.exe')
    else:
        return r"C:\Program Files\Tesseract-OCR\tesseract.exe"

try:
    pytesseract.pytesseract.tesseract_cmd = get_tesseract_path()
except Exception as e:
    print(f"Error setting Tesseract path: {e}")

# --- PDF Text Extraction Functions ---
def extract_text_from_pdf(pdf_path):
    """Extracts machine-readable text directly from a PDF."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            full_text += page.get_text("text")
        doc.close()
        return full_text
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return ""

def extract_text_with_ocr(pdf_path):
    """Renders PDF pages as images and uses OCR to extract text."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))
            text = pytesseract.image_to_string(image, lang='eng')
            full_text += text + "\n"
        doc.close()
        return full_text
    except Exception as e:
        print(f"Error during OCR for {pdf_path}: {e}")
        return ""

# --- Gemini API Interaction ---
def get_gemini_prompt():
    """Returns the detailed instruction prompt for the Gemini API for all 85 fields."""
    json_schema = {
        "invoiceHeader": {
            "invoiceDate": "NA", "invoiceNo": "NA", "supplierInvoiceNo": "NA", "supplierInvoiceDate": "NA", "voucherType": "Purchase",
            "orderNo": "NA", "orderDate": "NA", "orderDueDate": "NA", "documentType": "Invoice", "subType": "NA",
            "receiptNoteNo": "NA", "receiptNoteDate": "NA"
        },
        "supplierDetails": {
            "name": "NA", "address1": "NA", "address2": "NA", "address3": "NA", "pincode": "NA", "state": "NA",
            "placeOfSupply": "NA", "country": "INDIA", "gstin": "NA", "gstRegistrationType": "Regular"
        },
        "buyerDetails": {
            "name": "NA", "address1": "NA", "address2": "NA", "address3": "NA", "pincode": "NA", "state": "NA",
            "place": "NA", "gstin": "NA"
        },
        "logisticsDetails": {
            "lrNo": "NA", "despatchThrough": "NA", "destination": "NA", "transportMode": "NA", "distance": "NA",
            "transporterName": "NA", "vehicleNumber": "NA", "vehicleType": "NA", "docAirWayBillNo": "NA", "docDate": "NA", "transporterID": "NA"
        },
        "eWayBillDetails": {
            "eWayBillNo": "NA", "eWayBillDate": "NA", "consolidatedEWayBillNo": "NA", "consolidatedEWayDate": "NA", "statusOfEWayBill": "NA"
        },
        "lineItems": [{
            "itemName": "NA", "hsnCode": "NA", "itemDescription": "NA", "taxRate": 0.0, "batchNo": "NA",
            "mfgDate": "NA", "expDate": "NA", "qty": 0, "uom": "NA", "rate": 0.0, "discount": 0.0, "amount": 0.0
        }],
        "summary": {
            "totalAmount": 0.0, "cgstLedger": "CGST", "cgstAmount": 0.0, "sgstLedger": "SGST", "sgstAmount": 0.0,
            "igstLedger": "IGST", "igstAmount": 0.0, "cessLedger": "Cess", "cessAmount": 0.0, "roundOffLedger": "Round-Off", "roundOffAmount": 0.0,
            "narration": "NA", "termsOfPayment": "NA", "otherReference": "NA", "termsOfDelivery": "NA",
            "purchaseLedger": "Purchase Account", "costCenterGodown": "Main Location"
        }
    }
    
    prompt = f"""
    You are an expert AI data extractor for invoices. Your task is to analyze the raw text from an Indian invoice and extract all specified information into a structured JSON format.

    Instructions:
    1.  Thoroughly analyze the provided invoice text.
    2.  Populate all fields in the JSON schema below.
    3.  If information for a field is not found, you MUST use the default value ("NA" for text, 0 for numbers). Do not leave any field blank.
    4.  'lineItems' must be a JSON array. Create one object in the array for each distinct product or service line item found in the invoice table.
    5.  The final output must be ONLY the valid JSON object, with no additional text, explanations, or formatting.

    JSON Schema to populate:
    {json.dumps(json_schema, indent=2)}

    Now, here is the invoice text:
    ---
    """
    return prompt

def extract_data_with_gemini(api_key, invoice_text):
    """Sends the invoice text to the Gemini API and parses the structured response."""
    if not api_key: raise ValueError("Gemini API Key is required.")
    full_prompt = get_gemini_prompt() + invoice_text
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"parts": [{"text": full_prompt}]}],
        "generationConfig": { "responseMimeType": "application/json", "temperature": 0.1 }
    }
    try:
        response = requests.post(api_url, headers=headers, data=json.dumps(payload))
        response.raise_for_status()
        content_text = response.json()['candidates'][0]['content']['parts'][0]['text']
        return json.loads(content_text)
    except requests.exceptions.RequestException as e:
        error_info = response.json().get('error', {})
        raise ConnectionError(f"API Error: {error_info.get('message', 'Failed to connect to Gemini API.')}")
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        raise ValueError(f"Could not parse a valid JSON response from Gemini. Error: {e}")

# --- Main Orchestration Function (LOGIC REVERTED to one row per item) ---
def process_invoice_file(pdf_path, api_key):
    """
    Orchestrates extraction and returns a list of dictionaries, one for each line item.
    """
    print(f"Processing {pdf_path}...")
    
    text = extract_text_from_pdf(pdf_path)
    if len(text.strip()) < 150:
        print("Switching to OCR...")
        text = extract_text_with_ocr(pdf_path)
    if not text:
        raise ValueError("Could not extract any text from the PDF.")

    structured_data = extract_data_with_gemini(api_key, text)

    # --- Reformat the JSON from Gemini into the flat, row-based structure for Excel ---
    all_rows = []
    
    # Unpack all data sections from the JSON response, using .get() for safety
    header = structured_data.get('invoiceHeader', {})
    supplier = structured_data.get('supplierDetails', {})
    buyer = structured_data.get('buyerDetails', {})
    logistics = structured_data.get('logisticsDetails', {})
    eway = structured_data.get('eWayBillDetails', {})
    summary = structured_data.get('summary', {})
    
    # Base data is the same for all items in one invoice
    base_data = {
        'Invoice Date': header.get('invoiceDate'), 'Invoice No': header.get('invoiceNo'),
        'Supplier Invoice No': header.get('supplierInvoiceNo'), 'Supplier Invoice Date': header.get('supplierInvoiceDate'),
        'Voucher Type': header.get('voucherType'), 'Supplier Name': supplier.get('name'),
        'Address 1': supplier.get('address1'), 'Address 2': supplier.get('address2'), 'Address 3': supplier.get('address3'),
        'Supplier Pincode': supplier.get('pincode'), 'State': supplier.get('state'), 'Place of Supply': supplier.get('placeOfSupply'),
        'Country': supplier.get('country'), 'GSTIN/UIN': supplier.get('gstin'),
        'Consignor From Name': buyer.get('name'), 'Consignor From Add 1': buyer.get('address1'), 'Consignor From Add 2': buyer.get('address2'),
        'Consignor From Add 3': buyer.get('address3'), 'Consignor From State': buyer.get('state'), 'Consignor From Place': buyer.get('place'),
        'Consignor From Pincode': buyer.get('pincode'), 'Consignor From GSTIN': buyer.get('gstin'),
        'GST Registration Type': supplier.get('gstRegistrationType'), 'Receipt Note No': header.get('receiptNoteNo'),
        'Receipt Note Date': header.get('receiptNoteDate'), 'Order No': header.get('orderNo'), 'Order Date': header.get('orderDate'),
        'Order Due Date': header.get('orderDueDate'), 'LR No': logistics.get('lrNo'), 'Despatch Through': logistics.get('despatchThrough'),
        'Destination': logistics.get('destination'), 'Term of Payment': summary.get('termsOfPayment'),
        'Other Reference': summary.get('otherReference'), 'Terms of Delivery': summary.get('termsOfDelivery'),
        'Purchase Ledger': summary.get('purchaseLedger'), 'CGST Ledger': summary.get('cgstLedger'), 'CGST Amount': summary.get('cgstAmount'),
        'SGST Ledger': summary.get('sgstLedger'), 'SGST Amount': summary.get('sgstAmount'), 'IGST Ledger': summary.get('igstLedger'),
        'IGST Amount': summary.get('igstAmount'), 'Cess Ledger': summary.get('cessLedger'), 'Cess Amount': summary.get('cessAmount'),
        'Round off Ledger': summary.get('roundOffLedger'), 'Round off Amount': summary.get('roundOffAmount'),
        'Cost Center Godown': summary.get('costCenterGodown'), 'Narration': summary.get('narration'),
        'e-Way Bill No': eway.get('eWayBillNo'), 'e-Way Bill Date': eway.get('eWayBillDate'),
        'Consolidated e-Way Bill No': eway.get('consolidatedEWayBillNo'), 'Consolidated e-Way Date': eway.get('consolidatedEWayDate'),
        'Sub Type': header.get('subType'), 'Document Type': header.get('documentType'),
        'Status of e-Way Bill': eway.get('statusOfEWayBill'), 'Transport Mode': logistics.get('transportMode'),
        'Distance': logistics.get('distance'), 'Transporter Name': logistics.get('transporterName'),
        'Vehical Number': logistics.get('vehicleNumber'), 'Vehical Type': logistics.get('vehicleType'),
        'Doc/AirWay Bill No': logistics.get('docAirWayBillNo'), 'Doc Date': logistics.get('docDate'),
        'Transporter ID': logistics.get('transporterID'),
    }

    line_items = structured_data.get('lineItems', [{}])
    if not line_items: line_items = [{}] # Ensure at least one row is created

    for item in line_items:
        # Combine the main invoice data with the line item data for each row
        full_row = {
            **base_data, # Start with all the header/footer info
            'Item Name': item.get('itemName'), 'HSN Code': item.get('hsnCode'), 'Item Description': item.get('itemDescription'),
            'Tax Rate': item.get('taxRate'), 'Batch No': item.get('batchNo'), 'Mfg Date': item.get('mfgDate'),
            'Exp Date': item.get('expDate'), 'QTY': item.get('qty'), 'UOM': item.get('uom'), 'Rate': item.get('rate'),
            'Discount': item.get('discount'), 'Amount': item.get('amount'),
        }
        all_rows.append(full_row)
        
    print(f"Successfully processed {pdf_path} using Gemini, found {len(all_rows)} items.")
    return all_rows