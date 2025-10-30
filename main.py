import uvicorn
import api.converter as converter  
from fastapi import FastAPI, File, UploadFile, Query, HTTPException
from fastapi.responses import StreamingResponse
from io import BytesIO
from fastapi.middleware.cors import CORSMiddleware
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
    """
    This endpoint receives a resume file and a template_id,
    converts it, and returns the formatted DOCX file.
    """
    
    # 1. Read the uploaded file into a buffer
    file_contents = await file.read()
    file_buffer = BytesIO(file_contents)
    
    text = ""
    
    # 2. Extract text using shared functions from converter
    if file.filename.endswith('.pdf'):
        text = converter.extract_text_from_pdf(file_buffer)
    elif file.filename.endswith('.docx'):
        text = converter.extract_text_from_docx(file_buffer)
    else:
        raise HTTPException(status_code=400, detail="Unsupported file type. Please upload a PDF or DOCX.")
    
    if not text:
        raise HTTPException(status_code=400, detail="Could not extract text from the document.")

    # 3. Process based on the selected template_id
    docx_buffer = None
    try:
        if template_id == "template1":
            # --- Use Template 1 functions ---
            prompt = converter.get_prompt_for_template_1(text)
            gemini_text = converter.call_gemini_api(prompt)
            data = converter.parse_text_for_template_1(gemini_text)
            docx_buffer = converter.build_docx_for_template_1(data)
            
        elif template_id == "template2":
            # --- Use Template 2 functions ---
            prompt = converter.get_prompt_for_template_2(text)
            gemini_text = converter.call_gemini_api(prompt)
            data = converter.parse_text_for_template_2(gemini_text)
            docx_buffer = converter.build_docx_for_template_2(data)

    except Exception as e:
        # Return a server error if anything goes wrong during conversion
        raise HTTPException(status_code=500, detail=f"Error during conversion: {str(e)}")

    if not docx_buffer:
        raise HTTPException(status_code=500, detail="Failed to generate DOCX buffer.")

    # 4. Return the generated file as a stream
    response_filename = f"{template_id}_Formatted_{file.filename.split('.')[0]}.docx"
    
    return StreamingResponse(
        docx_buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={response_filename}"}
    )

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)