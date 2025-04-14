import unittest
from flask import Flask, request, send_file, Response
from unittest.mock import MagicMock
import fitz  # PyMuPDF (PyMuPDF)
from docx import Document, text
from docx.shared import RGBColor
import os
import csv
from datetime import datetime
from openai import AzureOpenAI
import io, json
import pytesseract
from PIL import Image

app = Flask(__name__)

# Azure OpenAI configuration
endpoint = "https://openai-eus-ti-poc-shared-resources.openai.azure.com/"
api_key = "f6f65288e4ec4f89a92387cb877c4b17"  # Replace with your actual API key
deployment_id = "chat-gpt-4o"  # or "chat-gpt-4o-mini"
api_version = "2025-01-01-preview"

client = AzureOpenAI(
    api_version="2025-01-01-preview",
    azure_endpoint=endpoint,
    api_key=api_key,
)


def extract_text_from_image(img_data):
    """Extracts text from an image using OCR."""
    try:
        image = Image.open(io.BytesIO(img_data))
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        print(f"OCR error: {e}")
        return ""


def pdf_to_png(pdf_path):
    """Converts a PDF to a list of PNG images."""
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        image = page.get_pixmap()
        img_data = image.tobytes("png")
        images.append(img_data)
    return images


def extract_changes_from_images(images, client, deployment_id, pdf_filename):
    """Extracts tracked changes from a list of PNG images (a batch) using Azure OpenAI."""
    
    all_text = ""
    for img_data in images:
        all_text += extract_text_from_image(img_data)

    prompt = f"""
    You are an expert document editor. You are given text extracted from a PDF file that was generated from a Word document with track changes enabled.

    Your task is to identify and extract only the paragraphs that have been modified or deleted by track changes. Please focus exclusively on changes such as insertions, deletions, and replacements.  Do not extract paragraphs without modification.
    
    - Only consider paragraphs that begin with a numerical prefix (e.g., "1.", "2.1", "3.a").
    - Identify the type of change for each modified paragraph: 'modified' or 'deleted'.
    - For 'modified' paragraphs, represent formatting changes as follows:
        - Underlined text: <u>text</u>
        - Strikethrough text: <s>text</s>
        - Highlighted text: <highlight>text</highlight>
    - Do not return the paragraphs that have no change.

    Return the output in JSON format, each element should has following info, 'paragraph_number': 'the paragraph number' , 'content':'the paragraph content'
    Here is the text:
    {all_text}
    """
    try:

        messages = [
            {"role": "system", "content": "You are an expert document editor."},
            {"role": "user", "content": prompt},
        ]

        response = client.chat.completions.with_raw_response.create(
            model=deployment_id,
            messages=messages,
            temperature=0,
            top_p=0.95,
            frequency_penalty=0,
            presence_penalty=0
        )

        response_content = response.choices[0].message.content.strip()
        print(f"response_content: {response_content}")

        try:
            extracted_changes = json.loads(response_content)
            print(f"extracted_changes:{extracted_changes}")
        except json.JSONDecodeError as e:
            print(f"JSONDecodeError: {e}")
            extracted_changes = []

        token_usage = response.parse().usage.total_tokens if response.parse().usage.total_tokens else 0
        

        return extracted_changes, token_usage

    except Exception as e:
        print(f"Azure OpenAI API error:{e}, filename:{pdf_filename}")
        return [], 0


def add_formatted_text(paragraph, text):
    """Adds text to a paragraph, applying formatting markers."""
    if text == "":
        return

def fill_word_template(template_path, changes):
    """Fills a Word template with extracted changes."""
    doc = Document(template_path)
    for change in changes:
        if "content" not in change or change["content"]=="":
            continue
        content=change["content"]
        for paragraph in doc.paragraphs:
            if paragraph.text.startswith(change["paragraph_number"]):
                    
                    parts = []
                    current_part = ""
                    in_tag = False
                    tag_name = ""

                    for i in range(len(content)):
                        if content[i] == '<' and (i + 1 < len(content) and content[i+1] !='/'):
                            if current_part!="":
                                parts.append(current_part)
                            current_part = ""
                            in_tag = True
                            tag_name_end_index = content.find(">", i)                            
                            tag_name = content[i + 1 : tag_name_end_index]
                            i = tag_name_end_index
                            parts.append("<"+tag_name+">")
                            
                        elif content[i] == '/':
                            in_tag = False
                            tag_name_end_index = content.find(">", i)                            
                            tag_name = content[i + 1 : tag_name_end_index]
                            i = content.find("</"+tag_name+">", i) + len("</"+tag_name+">")
                            parts.append(content[tag_name_end_index + 1: i])
                            parts.append("</"+tag_name+">")
                        else:
                            current_part += content[i]
                    
                    parts.append(current_part)
                    for part in parts:
                        print(f"part:{part}")
                        add_formatted_text(paragraph, part)
    return doc


def log_api_call(pdf_filename, word_filename, api_info, token_usage=None):
    """Logs API call information to a CSV file."""
    log_data = {
        "timestamp": datetime.now().isoformat(),
        "pdf_filename": pdf_filename,
        "word_filename": word_filename,
        "api_info": api_info,
        "token_usage": token_usage,
    }
    with open("api_log.csv", "a", newline="", encoding="utf-8") as csvfile:
        fieldnames = ["timestamp", "pdf_filename", "word_filename", "api_info", "token_usage"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if csvfile.tell() == 0:
            writer.writeheader()
        writer.writerow(log_data)


def add_formatted_text(paragraph, text):
    """Adds text to a paragraph, applying formatting markers."""
    if text == "":
        return
    
    start_underline = text.find("<u>")
    start_strikethrough = text.find("<s>")
    start_highlight = text.find("<highlight>")

    if start_underline != -1:
        run = paragraph.add_run(text[0:start_underline])
        run_underline = paragraph.add_run(text[start_underline+3:text.find("</u>")])
        run_underline.underline = True
        add_formatted_text(paragraph,text[text.find("</u>")+4:])
    elif start_strikethrough != -1:
        run = paragraph.add_run(text[0:start_strikethrough])
        run_strikethrough = paragraph.add_run(text[start_strikethrough+3:text.find("</s>")])
        run_strikethrough.font.strike = True
        add_formatted_text(paragraph,text[text.find("</s>")+4:])
    elif start_highlight != -1:
        run = paragraph.add_run(text[0:start_highlight])
        run_highlight = paragraph.add_run(text[start_highlight+10:text.find("</highlight>")])
        run_highlight.font.highlight_color = 4  # wdYellow
        add_formatted_text(paragraph,text[text.find("</highlight>")+11:])
    else:
        paragraph.add_run(text)
        
@app.route("/")
@app.route("/convert", methods=["POST"])
def convert_pdf_to_word():
    """API endpoint to convert a PDF with tracked changes to a Word document."""
    if "file" not in request.files or "template" not in request.files:
        return Response("Please provide both a PDF file and a Word template.", status=400)

    pdf_file = request.files["file"]
    template_file = request.files["template"]

    if not pdf_file.filename.endswith(".pdf") or not template_file.filename.endswith(".docx"):
        return Response("Invalid file types. Please provide a PDF file and a Word template.", status=400)

    pdf_filename = pdf_file.filename
    template_filename = template_file.filename

    pdf_path = os.path.join("/tmp", pdf_filename)
    template_path = os.path.join("/tmp", template_filename)
    pdf_file.save(pdf_path)
    template_file.save(template_path)
    try:
        images = pdf_to_png(pdf_path)
        changes = []
        total_token_usage = 0
        batch_size = 10
        num_batches = (len(images) + batch_size - 1) // batch_size

        for i in range(num_batches):
            start_index = i * batch_size
            end_index = min((i + 1) * batch_size, len(images))
            batch_images = images[start_index:end_index]


            print(f"Processing batch {i+1}/{num_batches} (pages {start_index+1}-{end_index})")
            batch_changes, token_usage = extract_changes_from_images(batch_images, client, deployment_id)
            changes.extend(batch_changes)
            total_token_usage += token_usage

        output_doc = fill_word_template(template_path, changes)

        word_filename = f"output_{os.path.splitext(pdf_filename)[0]}.docx"
        output_path = os.path.join("/tmp", word_filename)
        output_doc.save(output_path)

        log_api_call(pdf_filename, word_filename, f"Azure OpenAI API: {deployment_id}", total_token_usage)

        return send_file(output_path, as_attachment=True, download_name=word_filename)    

    except Exception as e:
        print(f"Error: {e}")
        return Response(f"An error occurred: {e}", status=500)

    finally:
        os.remove(pdf_path)
        os.remove(template_path)
        if 'output_path' in locals():
            os.remove(output_path)


if __name__ == "__main__":
    if not os.path.exists("api_log.csv"):
        with open("api_log.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "pdf_filename", "word_filename", "api_info", "token_usage"])
    app.run(debug=False, host="0.0.0.0", port=5000)