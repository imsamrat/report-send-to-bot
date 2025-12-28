
import os
import json
import requests

import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

def get_credentials():
    creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        try:
            creds_dict = json.loads(creds_json)
        except json.JSONDecodeError:
            if os.path.exists(creds_json):
                return service_account.Credentials.from_service_account_file(
                    creds_json, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
            raise Exception("INVALID GOOGLE_CREDENTIALS_JSON format")
        return service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    
    if os.path.exists('Credentials.json'):
        return service_account.Credentials.from_service_account_file(
            'Credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    
    raise Exception("No credentials found. Set GOOGLE_CREDENTIALS_JSON or provide Credentials.json")

def fetch_data():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet_id = os.getenv('SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME')
    range_val = os.getenv('RANGE_1')
    range_name = f"{sheet_name}!{range_val}"
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=range_name).execute()
    return result.get('values', [])

def send_telegram_photo(bot_token, chat_id, image_path):
    url = f"https://api.telegram.org/bot{bot_token}/sendPhoto"
    with open(image_path, 'rb') as f:
        files = {'photo': f}
        data = {'chat_id': chat_id}
        response = requests.post(url, data=data, files=files)
        return response.json()

def generate_image(data):
    if not data:
        print("No data to visualize")
        return
    
    # --- Data Cleaning ---
    # Detect if Column 0 is empty
    first_col_empty = all((not row or row[0] == '') for row in data)
    if first_col_empty:
        data = [row[1:] for row in data]

    title_text = "Daily Summary"
    clean_rows = []
    
    found_title = False
    for row in data:
        if not row or not any(row): continue
        row_str = "".join(row).lower()
        if not found_title and ("summary" in row_str or "issued" in row_str) and "zipper" not in row_str:
            title_text = next((x for x in row if x), "Daily Summary")
            # Clean title text: remove newlines or extra spaces
            title_text = title_text.replace("\n", " ").strip()
            found_title = True
            continue
        clean_rows.append(row)

    max_len = 0
    if clean_rows:
        max_len = max(len(r) for r in clean_rows)
    # Pad rows
    final_data = [row + [''] * (max_len - len(row)) for row in clean_rows]
    
    processed_data = [] # Rows
    
    # Colors
    COLOR_BLUE_TITLE = "#5b9bd5"
    COLOR_GREEN = "#a9d08e"
    COLOR_ORANGE = "#f4b084"
    COLOR_TEAL = "#8497b0"
    COLOR_GRAY = "#bfbfbf"
    COLOR_WHITE = "#ffffff"
    COLOR_BORDER = "black"
    
    colors_per_row = []
    font_weights_per_row = [] 
    line_colors_per_row = []
    heights_per_row = []

    NORMAL_HEIGHT = 30
    GAP_HEIGHT = 10 # Narrow gap

    for row in final_data:
        first_cell = row[0].strip() if row else ""
        
        bg_color = COLOR_WHITE
        font_weight = 'normal'
        
        # Determine if we need a gap row BEFORE this row
        if ("MT" in first_cell) or ("Total" in first_cell) or ("Avg" in first_cell):
             if processed_data:
                 processed_data.append([''] * max_len)
                 colors_per_row.append(COLOR_WHITE)
                 font_weights_per_row.append('normal')
                 line_colors_per_row.append(COLOR_WHITE) # Hide border
                 heights_per_row.append(GAP_HEIGHT)
        
        # Style Current Row
        if "Zipper" in first_cell:
            bg_color = COLOR_GREEN
            font_weight = 'bold'
        elif "MT" in first_cell:
            bg_color = COLOR_ORANGE
            font_weight = 'bold'
        elif "Total" in first_cell:
            bg_color = COLOR_TEAL
            font_weight = 'bold'
        elif "Avg" in first_cell:
            bg_color = COLOR_GRAY
            font_weight = 'bold'
        
        processed_data.append(row)
        colors_per_row.append(bg_color)
        font_weights_per_row.append(font_weight)
        line_colors_per_row.append(COLOR_BORDER)
        heights_per_row.append(NORMAL_HEIGHT)

    # Apply bold to Total Values
    for i, color in enumerate(colors_per_row):
        if color == COLOR_TEAL:
            j = i + 1
            while j < len(processed_data):
                if line_colors_per_row[j] == COLOR_WHITE: 
                    break
                font_weights_per_row[j] = 'bold'
                j += 1
                
    # Apply HTML Bold Formatting
    formatted_data = [] 
    for r_idx, row in enumerate(processed_data):
        new_row = []
        is_bold_row = (font_weights_per_row[r_idx] == 'bold')
        for cell in row:
            val = str(cell)
            if is_bold_row and val.strip():
                val = f"<b>{val}</b>"
            new_row.append(val)
        formatted_data.append(new_row)

    cols = list(map(list, zip(*formatted_data)))
    
    # 2D Color/Line Arrays
    fill_color_array = []
    line_color_array = []
    
    for c in range(len(cols)):
        col_colors = []
        col_lines = []
        for r in range(len(processed_data)):
            col_colors.append(colors_per_row[r])
            col_lines.append(line_colors_per_row[r])
        fill_color_array.append(col_colors)
        line_color_array.append(col_lines)

    # --- Dimensions & Layout ---
    # Calculate Table Height
    # Enforce scalar height to avoid Plotly validation errors
    ROW_HEIGHT = 30
    table_pixel_height = len(processed_data) * ROW_HEIGHT
    
    # Margins
    MARGIN_TOP = 60 
    MARGIN_BOTTOM = 20
    
    total_height = table_pixel_height + MARGIN_TOP + MARGIN_BOTTOM
    total_width = 1100

    fig = go.Figure(data=[go.Table(
        header=dict(values=[], height=0),
        cells=dict(
            values=cols,
            fill_color=fill_color_array,
            align='center',
            font=dict(color='black', size=13, family="Arial"),
            height=ROW_HEIGHT, # Strict scalar
            line=dict(color=line_color_array, width=1)
        ),
        columnwidth=[2.5] + [1.5]*(len(cols)-1)
    )])
    
    # Layout using Paper Coordinates
    # Y=1 is top of plot area. Y>1 is inside Top Margin.
    # We need to calculate how much "Paper Units" the Top Margin corresponds to.
    plot_height = total_height - MARGIN_TOP - MARGIN_BOTTOM
    margin_top_paper = MARGIN_TOP / plot_height
    
    # Title Bar Y Range in Paper Coords (approx):
    # From 1.0 (top of plot) to 1.0 + margin_top_paper
    
    fig.update_layout(
        margin=dict(l=20, r=20, t=MARGIN_TOP, b=MARGIN_BOTTOM),
        height=total_height,
        width=total_width,
        shapes=[
            # Outer Frame
            dict(
                type="rect",
                xref="paper", yref="paper", 
                x0=0, y0=0, x1=1, y1=1,
                line=dict(color="black", width=2),
                fillcolor="rgba(0,0,0,0)"
            ),
             # Title Background (Above plot, in margin)
            dict(
                type="rect",
                xref="paper", yref="paper",
                x0=0, x1=1, 
                y0=1, y1=1 + margin_top_paper, 
                fillcolor=COLOR_BLUE_TITLE,
                line=dict(width=0)
            )
        ]
    )
    
    # Place Header Text
    fig.add_annotation(
        text=title_text,
        xref="paper", yref="paper",
        x=0.5, 
        y=1 + (margin_top_paper / 2), # Middle of margin area
        showarrow=False,
        font=dict(color="white", size=18, family="Arial", weight='bold'),
        align="center",
        valign="middle"
    )

    output_file = "report.png"
    # Fallback to scalar height if array fails (robustness)
    try:
        fig.write_image(output_file, engine="kaleido", scale=2)
    except Exception as e:
        print(f"Error with variable height: {e}. Retrying with scalar height.")
        fig.data[0].cells.height = 30
        fig.write_image(output_file, engine="kaleido", scale=2)
        
    return output_file

def main():
    try:
        data = fetch_data()
        image_path = generate_image(data)
        
        bot_token = os.getenv('TELEGRAM_BOT_TOKEN')
        chat_id = os.getenv('TELEGRAM_CHAT_ID')
        
        if bot_token and chat_id and image_path:
            send_telegram_photo(bot_token, chat_id, image_path)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
