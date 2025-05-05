import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

class GeminiProcessor:
    def __init__(self):
        load_dotenv()
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable not set")
        genai.configure(api_key=api_key)
        
        # Use the newer model version
        self.model = genai.GenerativeModel('gemini-1.5-flash')
    
    def process(self, text_content, target_audience, tone, custom_instructions):
        prompt = self._create_prompt(text_content, target_audience, tone, custom_instructions)
        
        try:
            # Generate content using the correct model name
            response = self.model.generate_content(prompt)
            
            if not response.text:
                raise ValueError("Empty response received from Gemini API")
            
            # Clean and format the response text
            response_text = response.text.strip()
            
            # Find the last complete JSON array
            start_idx = response_text.find('[')
            end_idx = response_text.rfind(']')
            
            if start_idx == -1 or end_idx == -1:
                raise ValueError("Response does not contain a JSON array")
            
            # Extract the complete JSON array
            json_text = response_text[start_idx:end_idx + 1]
            
            try:
                parsed_response = json.loads(json_text)
            except json.JSONDecodeError as e:
                print(f"Raw response: {response_text}")
                raise ValueError(f"Failed to parse Gemini response as JSON: {str(e)}")
            
            # Validate response structure
            if not isinstance(parsed_response, list):
                raise ValueError("Response must be a list of slides")
            
            for slide in parsed_response:
                if not isinstance(slide, dict):
                    raise ValueError("Each slide must be a dictionary")
                if "title" not in slide or "bullets" not in slide:
                    raise ValueError("Each slide must have 'title' and 'bullets' fields")
            
            return parsed_response
            
        except Exception as e:
            raise ValueError(f"Error processing Gemini request: {str(e)}")
    
    def _create_prompt(self, text_content, target_audience, tone, custom_instructions):
        return f"""
        Create a presentation outline from the following text content.
        Format your response EXACTLY as a JSON array of slides.
        The response must be a valid JSON array starting with '[' and ending with ']'.
        Do not include any text before or after the JSON array.
        Keep the response concise and within reasonable length.
        
        Target Audience: {target_audience}
        Tone: {tone}
        Additional Instructions: {custom_instructions}
        
        Text Content:
        {text_content}
        
        Response format must be exactly like this:
        [
            {{
                "title": "Slide Title",
                "bullets": ["Point 1", "Point 2"],
                "image_hint": "Page number or context where an image might be relevant"
            }}
        ]
        """