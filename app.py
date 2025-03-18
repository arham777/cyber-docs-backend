from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import shutil
import uuid
import tempfile
import uvicorn
from typing import List
import time
from pathlib import Path

# Import your existing document processing functions
from main import apply_branding_to_docx, convert_pdf_to_docx

app = FastAPI(title="Document Template Processor", 
              description="API for processing documents with a branded template")

# Set up CORS to allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Specify your frontend URL in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configuration
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
TEMPLATE_DOCX = "cybergen-template.docx"  # Your branded template

# Create directories if they don't exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Store job statuses
processing_jobs = {}

@app.get("/")
async def root():
    return {"message": "Document Template Processor API"}

@app.post("/upload-files/")
async def upload_files(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...)
):
    """
    Upload multiple documents (DOCX or PDF) for processing.
    Returns a job ID that can be used to check status and download results.
    """
    # Create a unique job ID
    job_id = str(uuid.uuid4())
    job_upload_dir = os.path.join(UPLOAD_DIR, job_id)
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    
    os.makedirs(job_upload_dir, exist_ok=True)
    os.makedirs(job_output_dir, exist_ok=True)
    
    # Save uploaded files
    saved_files = []
    for file in files:
        # Validate file type
        if not (file.filename.endswith('.docx') or file.filename.endswith('.pdf')):
            continue  # Skip unsupported files
            
        # Save file
        file_path = os.path.join(job_upload_dir, file.filename)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        saved_files.append(file.filename)
    
    if not saved_files:
        raise HTTPException(status_code=400, detail="No supported files were uploaded (only .docx and .pdf are supported)")
    
    # Initialize job status
    processing_jobs[job_id] = {
        "status": "queued",
        "files": saved_files,
        "processed_files": [],
        "created_at": time.time(),
        "completed_at": None
    }
    
    # Start processing in the background
    background_tasks.add_task(process_documents, job_id, job_upload_dir, job_output_dir)
    
    return {
        "job_id": job_id,
        "status": "queued",
        "message": f"Processing {len(saved_files)} files",
        "files": saved_files
    }

@app.get("/job-status/{job_id}")
async def job_status(job_id: str):
    """Check the status of a document processing job"""
    if job_id not in processing_jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    return processing_jobs[job_id]

@app.get("/download/{job_id}/{filename}")
async def download_file(job_id: str, filename: str):
    """Download a processed file"""
    if job_id not in processing_jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    file_path = os.path.join(job_output_dir, filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=file_path, 
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/download-all/{job_id}")
async def download_all(job_id: str, background_tasks: BackgroundTasks):
    """Create and download a zip file with all processed documents"""
    if job_id not in processing_jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job_status = processing_jobs[job_id]
    if job_status["status"] != "completed":
        raise HTTPException(status_code=400, detail="Job processing is not yet complete")
    
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    zip_filename = f"{job_id}_documents.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_filename)
    
    # Create a zip file containing all processed documents
    shutil.make_archive(os.path.splitext(zip_path)[0], 'zip', job_output_dir)
    
    # Clean up the zip file after download
    background_tasks.add_task(cleanup_zip, zip_path)
    
    return FileResponse(
        path=f"{os.path.splitext(zip_path)[0]}.zip",
        filename=zip_filename,
        media_type="application/zip"
    )

@app.delete("/job/{job_id}")
async def delete_job(job_id: str):
    """Delete a job and its associated files"""
    if job_id not in processing_jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job_upload_dir = os.path.join(UPLOAD_DIR, job_id)
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    
    # Remove directories
    if os.path.exists(job_upload_dir):
        shutil.rmtree(job_upload_dir)
    
    if os.path.exists(job_output_dir):
        shutil.rmtree(job_output_dir)
    
    # Remove job from tracking
    del processing_jobs[job_id]
    
    return {"message": "Job deleted successfully"}

@app.on_event("startup")
async def startup_event():
    """Clean up any old processing directories on startup"""
    # Clean up old uploads and outputs that might be left from previous runs
    clean_old_dirs(UPLOAD_DIR)
    clean_old_dirs(OUTPUT_DIR)

def clean_old_dirs(parent_dir):
    """Remove directories older than 24 hours"""
    current_time = time.time()
    one_day_in_seconds = 86400  # 24 hours
    
    for item in os.listdir(parent_dir):
        item_path = os.path.join(parent_dir, item)
        if os.path.isdir(item_path):
            # Check if directory is older than 24 hours
            if current_time - os.path.getctime(item_path) > one_day_in_seconds:
                try:
                    shutil.rmtree(item_path)
                except Exception as e:
                    print(f"Error removing old directory {item_path}: {e}")

def process_documents(job_id: str, upload_dir: str, output_dir: str):
    """Process all documents in the upload directory"""
    try:
        # Update job status
        processing_jobs[job_id]["status"] = "processing"
        
        processed_files = []
        
        for filename in os.listdir(upload_dir):
            input_path = os.path.join(upload_dir, filename)
            
            try:
                if filename.endswith(".docx"):
                    # Process DOCX directly
                    output_path = os.path.join(output_dir, filename)
                    apply_branding_to_docx(input_path, output_path)
                    processed_files.append(filename)
                    
                elif filename.endswith(".pdf"):
                    # Convert PDF to DOCX, then apply branding
                    base_name = os.path.splitext(filename)[0]
                    temp_docx = os.path.join(output_dir, f"{base_name}_converted.docx")
                    output_docx = os.path.join(output_dir, f"{base_name}.docx")
                    
                    if convert_pdf_to_docx(input_path, temp_docx):
                        apply_branding_to_docx(temp_docx, output_docx)
                        
                        # Clean up temporary DOCX
                        if os.path.exists(temp_docx):
                            os.remove(temp_docx)
                        
                        processed_files.append(f"{base_name}.docx")
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
                # Continue with other files even if one fails
        
        # Update job status to completed
        processing_jobs[job_id]["status"] = "completed"
        processing_jobs[job_id]["processed_files"] = processed_files
        processing_jobs[job_id]["completed_at"] = time.time()
        
    except Exception as e:
        # Update job status to failed
        processing_jobs[job_id]["status"] = "failed"
        processing_jobs[job_id]["error"] = str(e)
        print(f"Job {job_id} failed: {str(e)}")

def cleanup_zip(zip_path):
    """Clean up the zip file after a delay"""
    time.sleep(300)  # Wait 5 minutes before deleting
    try:
        if os.path.exists(zip_path):
            os.remove(zip_path)
    except Exception as e:
        print(f"Error cleaning up zip file: {e}")

if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)