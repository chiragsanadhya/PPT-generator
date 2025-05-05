import streamlit as st
import os
from src.parser import DocumentParser
from src.image_extractor import ImageExtractor
from src.gemini_api import GeminiProcessor
from src.ppt_generator import PPTGenerator

def main():
    st.title("Document to Presentation Converter")
    
    # File upload
    uploaded_file = st.file_uploader("Upload your document", type=['pdf', 'docx', 'txt'])
    
    # User preferences
    target_audience = st.selectbox(
        "Select target audience",
        ["General", "Executive", "Technical"]
    )
    
    tone = st.selectbox(
        "Select presentation tone",
        ["Formal", "Friendly", "Concise"]
    )
    
    custom_instructions = st.text_area(
        "Additional instructions (optional)",
        "Example: Focus on key metrics and include charts"
    )
    
    if uploaded_file and st.button("Generate Presentation"):
        with st.spinner("Processing your document..."):
            # Save uploaded file temporarily
            temp_path = os.path.join("extracted", uploaded_file.name)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                # Parse document
                parser = DocumentParser()
                text_content = parser.parse(temp_path)
                
                # Extract images
                image_extractor = ImageExtractor()
                images = image_extractor.extract(temp_path)
                
                # Process with Gemini
                gemini = GeminiProcessor()
                slides_content = gemini.process(
                    text_content,
                    target_audience,
                    tone,
                    custom_instructions
                )
                
                # Generate PPT
                ppt_gen = PPTGenerator()
                output_path = os.path.join("output", "presentation.pptx")
                ppt_gen.generate(slides_content, images, output_path)
                
                # Provide download link
                with open(output_path, "rb") as f:
                    st.download_button(
                        "Download Presentation",
                        f,
                        file_name="presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                # Optional: Show Gemini response in sidebar
                with st.sidebar:
                    st.subheader("AI Processing Details")
                    st.json(slides_content)
                    
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
            finally:
                # Cleanup
                if os.path.exists(temp_path):
                    os.remove(temp_path)

if __name__ == "__main__":
    main()