import uvicorn
import converter
from fastapi import FastAPI, File, UploadFile, Query, HTTPException
from fastapi.responses import JSONResponse # <-- 1. Import JSONResponse
from io import BytesIO
from fastapi.middleware.cors import CORSMiddleware
import base64 # <-- 2. Import base64

origins = [
    "*", 
]
app = FastAPI(title="Resume Converter API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
@app.post("/convert-resume/")
async def convert_resume_endpoint(
    file: UploadFile = File(...),
    template_id: str = Query(..., enum=["template1", "template2"]) 
):
    
    file_contents = await file.read()
    file_buffer = BytesIO(file_contents)
    
    text = ""
    
    if file.filename.endswith('.pdf'):
        text = converter.extract_text_from_pdf(file_buffer)
    elif file.filename.endswith('.docx'):
        text = converter.extract_text_from_docx(file_buffer)
    else:
        raise HTTPException(status_code=400, detail="Unsupported file type...")
    
    if not text:
        raise HTTPException(status_code=400, detail="Could not extract text...")

    docx_buffer = None
    gemini_text = "" # <-- 3. Create a variable to hold the text
    try:
        if template_id == "template1":
            prompt = converter.get_prompt_for_template_1(text)
            gemini_text = converter.call_gemini_api(prompt) # <-- 4. Store the text
            data = converter.parse_text_for_template_1(gemini_text)
            docx_buffer = converter.build_docx_for_template_1(data)
            
        elif template_id == "template2":
            prompt = converter.get_prompt_for_template_2(text)
            gemini_text = converter.call_gemini_api(prompt) # <-- 5. Store the text
            data = converter.parse_text_for_template_2(gemini_text)
            docx_buffer = converter.build_docx_for_template_2(data)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error during conversion: {str(e)}")

    if not docx_buffer:
        raise HTTPException(status_code=500, detail="Failed to generate DOCX buffer.")

    # --- 6. START: New Return Logic ---
    
    # Get the binary data from the buffer
    docx_bytes = docx_buffer.getvalue()
    
    # Encode the binary data to a Base64 text string
    docx_b64 = base64.b64encode(docx_bytes).decode('utf-8')
    
    # Create the JSON payload
    response_data = {
        "gemini_text": gemini_text,
        "file_data_b64": docx_b64,
        "file_name": f"{template_id}_Formatted_{file.filename.split('.')[0]}.docx"
    }
    
    # Return the JSON
    return JSONResponse(content=response_data)
    # --- END: New Return Logic ---

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
