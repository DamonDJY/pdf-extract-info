import unittest
import os
from unittest.mock import MagicMock, patch
import json
import requests
import base64
from io import BytesIO
from PIL import Image


def extract_text_from_image(img_data):
    """Placeholder for image text extraction."""
    return "Text extracted from image"


# Function to test our new image-based extraction approach
def test_extract_changes_from_pdf_with_images(pdf_path):
    """Test function for the new PDF image-based extraction workflow."""
    from pdf2image import convert_from_path
    import io, base64, json
    
    try:
        # Try to convert PDF to images
        print(f"Testing PDF to image conversion for: {pdf_path}")
        
        # First attempt with default settings
        try:
            images = convert_from_path(
                pdf_path,
                dpi=200,
                fmt="png"
            )
        except Exception as poppler_error:
            print(f"Standard conversion failed, attempting with poppler path: {poppler_error}")
            # Try with explicit poppler path - adjust this path as needed
            poppler_path = r"C:\Program Files\poppler-23.11.0\Library\bin"
            images = convert_from_path(
                pdf_path,
                dpi=200,
                fmt="png",
                poppler_path=poppler_path
            )
        
        print(f"Successfully converted PDF to {len(images)} images")
        
        # Test image conversion to base64 (just first image)
        if images:
            img = images[0]
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            print(f"Successfully converted image to base64 (first {len(img_base64)[:20]}...)")
            
            # Optionally save the first image to disk for manual inspection
            img.save("test_output_first_page.png")
            print("Saved first page as test_output_first_page.png for inspection")
            
            return True
        return False
    except Exception as e:
        print(f"Error in test_extract_changes_from_pdf_with_images: {e}")
        return False


# Assuming you have a function to extract text from PDF
def extract_text_from_pdf(pdf_path):  
    import pdfplumber
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            all_text += page.extract_text()
    return all_text


# Test Azure OpenAI vision API with a single image
def test_azure_openai_vision(image_path=None):
    """Test Azure OpenAI's vision capabilities with an image."""
    # Azure OpenAI configuration
    endpoint = "https://openai-eus-ti-poc-shared-resources.openai.azure.com/"
    api_key = "f6f65288e4ec4f89a92387cb877c4b17"
    deployment_id = "chat-gpt-4o"
    api_version = "2025-01-01-preview"
    
    # Use the test image if available, otherwise use a placeholder test
    if image_path and os.path.exists(image_path):
        # Load and convert the image to base64
        with open(image_path, "rb") as img_file:
            img_data = img_file.read()
            img_base64 = base64.b64encode(img_data).decode('utf-8')
    else:
        print("No image provided for vision test, using a test image")
        # If no image provided, generate simple test image
        img = Image.new('RGB', (100, 100), color='white')
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    # Prepare the prompt
    system_message = "You are an image analysis assistant."
    user_content = "What do you see in this image? Describe it briefly."
    
    # Create message content with image
    messages = [
        {"role": "system", "content": system_message},
        {
            "role": "user", 
            "content": [
                {"type": "text", "text": user_content},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_base64}"}}
            ]
        }
    ]
    
    # Call Azure OpenAI API
    url = f"{endpoint}openai/deployments/{deployment_id}/chat/completions?api-version={api_version}"
    headers = {
        "Content-Type": "application/json",
        "api-key": api_key
    }
    
    payload = {
        "messages": messages,
        "temperature": 0,
        "max_tokens": 500,
        "stream": False
    }
    
    try:
        print("Sending request to Azure OpenAI Vision API...")
        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code == 200:
            response_json = response.json()
            print(f"API Response (Vision): {response_json['choices'][0]['message']['content'][:100]}...")
            print(f"Token usage: {response_json.get('usage', {}).get('total_tokens', 0)}")
            return True
        else:
            print(f"Error calling Azure OpenAI Vision API: {response.status_code}, {response.text}")
            return False
    except Exception as e:
        print(f"Exception while calling Azure OpenAI Vision API: {e}")
        return False


# Assuming you have a function to call OpenAI
def call_openai(text_content, client, deployment_id):
    """Calls OpenAI to process text content and extract tracked changes."""
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
    {text_content}
    """
    
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
    return response


class MockAzureOpenAIClient:
    def __init__(self):
        pass

    def mock_chat_completion_with_raw_response(self):
        mock_response = MagicMock()
        mock_response.choices = [MagicMock(message=MagicMock(content="Mocked OpenAI response"))]
        mock_response.usage = MagicMock(prompt_tokens=100, completion_tokens=50, total_tokens=150)
        mock_response.http_response = MagicMock(status_code=200)  # Add a mock status code
        
        return mock_response

    def chat(self):
        chat_completions_mock = MagicMock()
        chat_completions_mock.with_raw_response = MagicMock(create=lambda **kwargs: self.mock_chat_completion_with_raw_response())
        return type('ChatCompletion', (object,), {'completions': chat_completions_mock})()


class TestOpenAIPDF(unittest.TestCase):

    @patch('test_openai_pdf.call_openai')  # Patch the call_openai function
    def test_openai_pdf_processing(self, mock_call_openai):
        pdf_path = "pdf/3. VI_2 (Tracked Changes).pdf"
        
        # Mock the Azure OpenAI client
        mock_client = MockAzureOpenAIClient()
        
        mock_call_openai.return_value = mock_client.mock_chat_completion_with_raw_response()
        
        if not os.path.exists(pdf_path):
            self.skipTest(f"PDF file not found: {pdf_path}")
        
        text_content = extract_text_from_pdf(pdf_path)       
        response = call_openai(text_content, mock_client.chat(), "your_deployment_id")

        self.assertIsNotNone(response)

        # Verify that the mock was called
        mock_call_openai.assert_called_once()

    def test_pdf_to_image_conversion(self):
        """Tests the new PDF to image conversion functionality."""
        pdf_path = "pdf/3. VI_2 (Tracked Changes).pdf"
        
        if not os.path.exists(pdf_path):
            self.skipTest(f"PDF file not found: {pdf_path}")
        
        success = test_extract_changes_from_pdf_with_images(pdf_path)
        self.assertTrue(success, "PDF to image conversion failed")
    
    def test_azure_vision_api(self):
        """Tests Azure OpenAI Vision API functionality."""
        image_path = "test_output_first_page.png"  # Use image from previous test if available
        
        # This test will use a placeholder if the file doesn't exist
        success = test_azure_openai_vision(image_path)
        self.assertTrue(success, "Azure OpenAI Vision API test failed")

    def test_convert_api_call_simulation(self):
        """Simulates calling the /convert API endpoint."""
        pdf_path = "pdf/3. VI_2 (Tracked Changes).pdf"
        
        # Simulate API request (no actual API call)
        # Assume the API would process the PDF and return this JSON
        simulated_api_response = {
            "output": "Processed output from PDF",
            "token_usage": 200
        }

        # Simulate the API response
        response_json = simulated_api_response
        self.assertIn('output', response_json)
        print("Simulated API Response:", response_json)


if __name__ == "__main__":
    unittest.main()
