
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
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly']
    
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

def get_sheet_gid(service, spreadsheet_id, sheet_name):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in spreadsheet.get('sheets', []):
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    raise Exception(f"Sheet '{sheet_name}' not found.")

def export_range_as_image(spreadsheet_id, range_a1, output_filename="report.jpg"):
    creds = get_credentials()
    if not creds.valid:
        creds.refresh(Request())
    
    token = creds.token
    service = build('sheets', 'v4', credentials=creds)
    
    # Extract Sheet Name and Range
    # Expected format: "SheetName!A1:B2" or just "A1:B2" (default to first sheet?)
    # The .env says SHEET_NAME=Summary, RANGE_1=B5:J23
    # We should construct the full range string or handle gid.
    
    sheet_name = os.getenv('SHEET_NAME', 'Summary')
    gid = get_sheet_gid(service, spreadsheet_id, sheet_name)
    
    url = (f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
           f"?format=pdf"
           f"&gid={gid}"
           f"&range={range_a1}"
           f"&size=A4"
           f"&portrait=false"
           f"&fitw=true"
           f"&gridlines=false"
           f"&printtitle=false"
           f"&sheetnames=false"
           f"&pagenum=UNDEFINED"
           f"&attachment=false")

    print(f"Exporting PDF from Sheet Range: {range_a1}...")
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"Failed to export PDF: {response.text}")
        return None
    
    pdf_path = "temp_export.pdf"
    with open(pdf_path, "wb") as f:
        f.write(response.content)
    
    # Convert PDF to Image
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
    
    temp_png = "temp_export.png"
    pix.save(temp_png)
    
    # Crop whitespace using Pillow
    from PIL import Image, ImageChops
    
    def trim_whitespace(img_path, out_path, padding=40):
        try:
            with Image.open(img_path) as im:
                # Assume top-left pixel color is background (usually white)
                bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
                diff = ImageChops.difference(im, bg)
                diff = ImageChops.add(diff, diff, 2.0, -100)
                bbox = diff.getbbox()
                if bbox:
                    # Expand bbox by padding
                    left, upper, right, lower = bbox
                    width, height = im.size
                    
                    left = max(0, left - padding)
                    upper = max(0, upper - padding)
                    right = min(width, right + padding)
                    lower = min(height, lower + padding)
                    
                    im_crop = im.crop((left, upper, right, lower))
                    im_crop.save(out_path)
                    print(f"Image cropped to {im_crop.size}")
                    return True
                else:
                    return False
        except Exception as e:
            print(f"Error trimming image: {e}")
            return False

    print("Trimming whitespace...")
    if not trim_whitespace(temp_png, output_filename):
        # Fallback if trimming fails or image is empty
        os.replace(temp_png, output_filename)
        
    # Clean up temp files
    try:
        if os.path.exists(pdf_path): os.remove(pdf_path)
        if os.path.exists(temp_png): os.remove(temp_png)
    except: pass
    
    return output_filename

def send_telegram_photo(bot_token, chat_id, image_path, caption=""):
    url = f"https://api.telegram.org/bot{bot_token}/sendPhoto"
    with open(image_path, 'rb') as f:
        files = {'photo': f}
        data = {'chat_id': chat_id, 'caption': caption}
        response = requests.post(url, data=data, files=files)
        return response.json()

def main():
    sheet_id = os.getenv('SHEET_ID')
    
    # Define ranges to process
    ranges = [
        ("Daily Summary", os.getenv('RANGE_1')),
        ("Report 2", os.getenv('RANGE_2')),
        ("Report 3", os.getenv('RANGE_3'))
    ]
    
    bot_token = os.getenv('TELEGRAM_BOT_TOKEN')
    chat_id = os.getenv('TELEGRAM_CHAT_ID')

    for title, range_val in ranges:
        if not range_val:
            continue
            
        print(f"\nProcessing {title}: {range_val}...")
        output_filename = f"report_{title.replace(' ', '_').lower()}.jpg"
        
        image_path = export_range_as_image(sheet_id, range_val, output_filename)
        
        if image_path:
            if bot_token and chat_id:
                print(f"Sending {title} to Telegram...")
                send_telegram_photo(bot_token, chat_id, image_path, caption=title)
                # Cleanup final image after sending
                try:
                    os.remove(image_path)
                    print(f"Removed temp file: {image_path}")
                except Exception as e:
                    print(f"Warning: Could not remove {image_path}: {e}")
            else:
                print("Telegram credentials missing.")
        else:
            print(f"Export failed for {title}.")

if __name__ == "__main__":
    main()
