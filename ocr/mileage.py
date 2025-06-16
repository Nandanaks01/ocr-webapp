import time
import pandas as pd
from collections import defaultdict
from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from msrest.authentication import CognitiveServicesCredentials
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
import os

# Azure OCR Setup

subscription_key = "CmVkKxEucLIUkS3xa8lAOT3I2dXCW9IlzDvVnDBVNhN3FAY616U6JQQJ99BFACYeBjFXJ3w3AAAFACOGmwr5"
endpoint = "https://ocr-table-reader.cognitiveservices.azure.com/"

client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))

def extract_table(read_result, num_columns=8, row_height=90):
    all_words = []
    for page in read_result.analyze_result.read_results:
        for line in page.lines:
            for word in line.words:
                left = min(word.bounding_box[::2])
                top = min(word.bounding_box[1::2])
                all_words.append({'text': word.text, 'left': left, 'top': top})

    all_words.sort(key=lambda x: (x['top'], x['left']))

    if not all_words:
        raise ValueError("No words detected by OCR")

    min_top = min(word['top'] for word in all_words)
    max_top = max(word['top'] for word in all_words)
    num_rows_est = int((max_top - min_top) / row_height) + 2
    row_anchors = [min_top + i * row_height for i in range(num_rows_est)]

    row_groups = defaultdict(list)
    for word in all_words:
        closest_row = min(row_anchors, key=lambda y: abs(word['top'] - y))
        row_groups[closest_row].append(word)

    sorted_rows = [sorted(row_groups[y], key=lambda w: w['left']) for y in sorted(row_groups)]

    reference_row = None
    for row in sorted_rows:
        digit_count = sum(c.isdigit() for w in row for c in w['text'])
        if len(row) >= num_columns and digit_count > 5:
            reference_row = row
            break

    if reference_row is None:
        raise ValueError("No valid reference row found")

    col_lefts = sorted(w['left'] for w in reference_row[:num_columns])
    col_boundaries = [(col_lefts[i] + col_lefts[i+1]) // 2 for i in range(len(col_lefts)-1)]
    col_boundaries.append(col_lefts[-1] + 200)

    structured_table = []
    for row in sorted_rows:
        aligned_row = ["" for _ in range(num_columns)]
        for word in row:
            for i, bound in enumerate(col_boundaries):
                if word['left'] <= bound:
                    aligned_row[i] += (" " if aligned_row[i] else "") + word['text']
                    break

        row_text = " ".join(aligned_row).upper()
        if any(cell.strip() for cell in aligned_row) and not any(
            keyword in row_text for keyword in [
                "MONTH", "YEAR", "PROJECT", "ODOMETER", "TOTAL",
                "TRIP", "BUSINESS", "MILES", "START", "END", "DAY"]):
            structured_table.append(aligned_row)

    return structured_table

def run(image_path, output_dir):
    with open(image_path, "rb") as image_stream:
        read_response = client.read_in_stream(image_stream, raw=True)

    operation_location = read_response.headers["Operation-Location"]
    operation_id = operation_location.split("/")[-1]

    while True:
        result = client.get_read_result(operation_id)
        if result.status not in ['notStarted', 'running']:
            break
        time.sleep(1)

    if result.status == OperationStatusCodes.succeeded:
        table_data = extract_table(result)
        headers = ["MONTH", "DAY", "YEAR", "PROJECT NO.", "START", "END", "TOTAL", "BUSINESS"]
        df = pd.DataFrame(table_data, columns=headers)
        out_file = os.path.join(output_dir, f"mileage_{os.path.basename(image_path).split('.')[0]}.xlsx")
        df.to_excel(out_file, index=False)
        return out_file
    else:
        raise Exception("Azure OCR failed.")
