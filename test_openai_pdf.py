import unittest
import os
from unittest.mock import MagicMock, patch
import json




def extract_text_from_image(img_data):
    """Placeholder for image text extraction."""
    return "Text extracted from image"

def extract_images_from_pdf(pdf_path):
    """Placeholder for PDF image extraction."""
    return ["image1", "image2"]


# Assuming you have a function to extract text from PDF
def extract_text_from_pdf(pdf_path):  
    import pdfplumber
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            all_text += page.extract_text()
    return all_text

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
        chat_completions_mock.with_raw_response = MagicMock(return_value=self.mock_chat_completion_with_raw_response())
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

        # Record response and token usage (print for demonstration)
        print("OpenAI Response:", response.value.choices[0].message.content)
        print("Token Usage:", response.value.usage.total_tokens)
        

if __name__ == "__main__":
    unittest.main()
 