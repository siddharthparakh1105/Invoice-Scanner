


# AI-Powered Invoice Extractor

This application extracts data from PDF invoices using a combination of direct text extraction, Optical Character Recognition (OCR), and Google's Gemini AI. It features a user-friendly graphical interface to select PDF files, process them, and append the extracted data to an Excel spreadsheet.

## Features

* **Modern GUI**: A clean and modern user interface built with Tkinter.
* **PDF Processing**: Handles both text-based and image-based (scanned) PDFs.
* **Smart Text Extraction**: First attempts to read text directly from the PDF. If that fails or yields minimal text, it automatically switches to OCR.
* **AI-Powered Data Extraction**: Leverages the Google Gemini API to intelligently parse raw text and extract data into a structured format, covering 85 distinct fields.
* **Batch Processing**: Select and process multiple PDF files in one go.
* **Append or Create Excel Files**: Appends extracted data to an existing Excel file or creates a new one if it doesn't exist.
* **Real-time Logging**: An in-app console shows the real-time status of the extraction process.
* **Progress Tracking**: A visual progress bar shows the overall status of the batch operation.

## Prerequisites

Before you begin, ensure you have the following installed:

1.  **Python 3**: [Download Python](https://www.python.org/downloads/)
2.  **Tesseract-OCR**: This is crucial for processing scanned or image-based PDFs.
    * Download and install from the official Tesseract repository: [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
    * **Important**: During installation, make sure to note the installation path. The application defaults to `C:\Program Files\Tesseract-OCR\tesseract.exe`. If you install it elsewhere, you will need to update the path in `extractor.py`.
3.  **Google Gemini API Key**: You need a valid API key from Google AI Studio to use the data extraction feature.
    * Get your key here: [Google AI Studio](https://ai.google.dev/)

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd <your-repository-name>
    ```

2.  **Install the required Python libraries:**
    Open your terminal or command prompt and run the following command to install all dependencies from the `requirements.txt` file:
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

1.  **Tesseract Path (if needed)**:
    If you installed Tesseract in a location other than the default, open the `extractor.py` file and modify the `get_tesseract_path` function to point to your `tesseract.exe`.

2.  **Gemini API Key**:
    Open the `main.py` file and replace the placeholder API key with your actual Gemini API key:
    ```python
    # in main.py, inside the InvoiceApp class
    self.api_key = "YOUR_GEMINI_API_KEY_HERE"
    ```

## How to Use

1.  **Run the application:**
    ```bash
    python main.py
    ```

2.  **Step 1: Upload PDF Files**
    * Click the "Browse Files..." button to select one or more PDF invoice files.
    * The button will update to show the number of files you've selected.

3.  **Step 2: Select Existing Excel File (Optional)**
    * Click the "Browse..." button to select an existing `.xlsx` file.
    * The extracted data will be appended to this file.
    * If you don't select a file, the application will prompt you to choose a location to save a *new* Excel file after the extraction is complete.

4.  **Step 3: Start AI Extraction**
    * Click the "Start Extraction" button.
    * The application will begin processing the files one by one. You can monitor the progress in the "Real-Time Logs" section and the progress bar at the bottom.

5.  **Completion**
    * Once all files are processed, the data will be saved to the specified Excel file.
    * A confirmation message will appear indicating that the export was successful.

## File Descriptions

* **`main.py`**: Contains the main application logic, including the Tkinter GUI, event handling, and thread management for processing.
* **`extractor.py`**: Handles all the backend logic for PDF processing. This includes extracting text, performing OCR with PyTesseract, and making the API call to Google Gemini.
* **`requirements.txt`**: A list of all the Python packages required to run the application.

