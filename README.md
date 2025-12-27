# PDF to DOCX Converter (Mediation Application Form)

This project recreates a legal PDF form as a Microsoft Word document using Python.
The focus of this assignment is on accurately replicating the structure, layout,
spacing, and content of the original PDF file.

## Features
- Upload a PDF file through a simple web interface
- Generate a Word (.docx) document that closely matches the original PDF layout
- Uses python-docx for document creation
- Built using Flask for easy interaction

## Technologies Used
- Python 3
- Flask
- python-docx

## Project Structure
- `app.py` – Flask application for handling file upload and download
- `converter.py` – Core logic to generate the Word document
- `templates/index.html` – Simple HTML interface
- `requirements.txt` – Project dependencies

## How to Run the Project

1. Clone the repository:
git clone https://github.com/abhay-chayal/pdf-to-docx-converter.git
cd pdf-to-docx-converter

2. Install dependencies:
pip install -r requirements.txt

3. Run the application:
python app.py

4. Open the application in your browser:
http://127.0.0.1:5000

5. Upload the PDF file and download the generated Word document:
Select the PDF file, click Convert, and the replicated .docx file will be downloaded automatically.


