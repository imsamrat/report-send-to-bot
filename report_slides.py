
import os
import json
import requests
import fitz  # PyMuPDF
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from dotenv import load_dotenv

load_dotenv()


def get_credentials():
    creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    # Change scope to allow writing (refreshing charts)
    SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations']
    
    if creds_json:
        try:
            creds_dict = json.loads(creds_json)
        except json.JSONDecodeError:
            if os.path.exists(creds_json):
                return service_account.Credentials.from_service_account_file(
                    creds_json, scopes=SCOPES)
            raise Exception("INVALID GOOGLE_CREDENTIALS_JSON format")
        return service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES)
    
    if os.path.exists('Credentials.json'):
        return service_account.Credentials.from_service_account_file(
            'Credentials.json', scopes=SCOPES)
    
    raise Exception("No credentials found.")

def refresh_charts(service, presentation_id):
    """
    Scans the presentation for linked Sheets Charts and refreshes them.
    Note: This only works for Charts. Linked Tables cannot be refreshed via API currently.
    """
    try:
        print("Scanning for linked charts to refresh...")
        presentation = service.presentations().get(presentationId=presentation_id).execute()
        slides = presentation.get('slides', [])
        
        chart_ids = []
        for slide in slides:
            for element in slide.get('pageElements', []):
                if 'sheetsChart' in element:
                    chart_ids.append(element['objectId'])
        
        if not chart_ids:
            print("No linked charts found to refresh. (Note: Linked Tables cannot be auto-refreshed via API)")
            return

        print(f"Found {len(chart_ids)} charts. Refreshing...")
        
        requests = []
        for chart_id in chart_ids:
            requests.append({
                'refreshSheetsChart': {
                    'objectId': chart_id
                }
            })
            
        if requests:
            body = {'requests': requests}
            response = service.presentations().batchUpdate(
                presentationId=presentation_id, body=body).execute()
            print("Charts refreshed successfully.")
            
    except Exception as e:
        print(f"Error refreshing charts: {e}")
        # Turn into warning rather than stopping
        print("Continuing with export without refresh...")

def export_presentation_pdf(service, file_id):
    """
    Export Google Slide to PDF via Drive API
    """
    try:
        request = service.files().export_media(fileId=file_id, mimeType='application/pdf')
        pdf_content = request.execute()
        
        filename = "presentation_export.pdf"
        with open(filename, 'wb') as f:
            f.write(pdf_content)
        return filename
    except Exception as e:
        error_str = str(e)
        if "accessNotConfigured" in error_str and "drive.googleapis.com" in error_str:
            print("\nCRITICAL ERROR: Google Drive API is not enabled.")
            print("Please enable it here: https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=664882132915\n")
        print(f"Error exporting PDF: {e}")
        return None

def convert_pdf_page_to_image(pdf_path, page_num):
    """
    Convert specific 1-based page number to image
    """
    try:
        doc = fitz.open(pdf_path)
        pg_idx = page_num - 1
        if 0 <= pg_idx < len(doc):
            page = doc.load_page(pg_idx)
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72)) # 300 DPI high res
            image_path = f"slide_{page_num}.png"
            pix.save(image_path)
            # Optimize: Check file size?
            return image_path
        else:
            print(f"Page {page_num} out of range in PDF")
            return None
    except Exception as e:
        print(f"Error converting page {page_num}: {e}")
        return None

def send_telegram_photo(bot_token, chat_id, image_path, caption=""):
    url = f"https://api.telegram.org/bot{bot_token}/sendPhoto"
    with open(image_path, 'rb') as f:
        files = {'photo': f}
        data = {'chat_id': chat_id, 'caption': caption}
        response = requests.post(url, data=data, files=files)
        return response.json()

def main():
    slide_id = os.getenv('SLIDE_ID')
    slide_no_str = os.getenv('SLIDE_NO')
    
    if not slide_id or not slide_no_str:
        print("Error: SLIDE_ID or SLIDE_NO not found in .env")
        return

    try:
        slide_indices = [int(x.strip()) for x in slide_no_str.split(',') if x.strip()]
    except ValueError:
        print("Error: SLIDE_NO must be comma-separated integers (e.g. 1,2,3)")
        return

    print(f"Processing Slides {slide_indices} from Presentation {slide_id}")
    
    creds = get_credentials()
    # Refresh if needed
    if not creds.valid:
        creds.refresh(Request())
    
    # 1. Refresh Charts (Requires Slides API)
    slides_service = build('slides', 'v1', credentials=creds)
    refresh_charts(slides_service, slide_id)

    # 2. Export PDF (Requires Drive API)
    drive_service = build('drive', 'v3', credentials=creds)
    
    print("Exporting presentation to PDF...")
    pdf_path = export_presentation_pdf(drive_service, slide_id)
    
    if not pdf_path:
        print("Failed to export PDF.")
        return

    bot_token = os.getenv('TELEGRAM_BOT_TOKEN')
    chat_id = os.getenv('TELEGRAM_CHAT_ID')
    
    for num in slide_indices:
        print(f"Converting Slide {num} to image...")
        image_path = convert_pdf_page_to_image(pdf_path, num)
        
        if image_path and bot_token and chat_id:
            print(f"Sending Slide {num} to Telegram...")
            send_telegram_photo(bot_token, chat_id, image_path, caption=f"Slide {num}")
        else:
            print(f"Skipping Slide {num} (Conversion failed or credentials missing)")


if __name__ == "__main__":
    main()
