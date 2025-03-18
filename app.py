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
import logging
import sys

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Import your existing document processing functions
try:
    from main import apply_branding_to_docx, convert_pdf_to_docx
except ImportError as e:
    logger.error(f"Failed to import processing functions: {e}")
    sys.exit(1)

app = FastAPI(title="Document Template Processor", 
              description="API for processing documents with a branded template")

# Set up CORS to allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configuration
try:
    # Use Path for cross-platform compatibility
    BASE_DIR = Path(__file__).resolve().parent
    UPLOAD_DIR = BASE_DIR / "uploads"
    OUTPUT_DIR = BASE_DIR / "output"
    TEMPLATE_DOCX = BASE_DIR / "cybergen-template.docx"

    # Create directories if they don't exist
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Verify template exists
    if not TEMPLATE_DOCX.exists():
        logger.error(f"Template file not found at: {TEMPLATE_DOCX}")
        raise FileNotFoundError(f"Template file not found at: {TEMPLATE_DOCX}")

except Exception as e:
    logger.error(f"Failed to initialize directories: {e}")
    raise

# Store job statuses in a more reliable way
processing_jobs = {}

@app.get("/")
async def root():
    """Root endpoint with enhanced status checking"""
    try:
        template_exists = TEMPLATE_DOCX.exists()
        upload_dir_exists = UPLOAD_DIR.exists()
        output_dir_exists = OUTPUT_DIR.exists()
        
        # Check write permissions
        upload_writable = os.access(UPLOAD_DIR, os.W_OK)
        output_writable = os.access(OUTPUT_DIR, os.W_OK)
        
        return {
            "message": "Document Template Processor API",
            "status": "running",
            "environment": {
                "python_version": sys.version,
                "platform": sys.platform,
                "cwd": str(BASE_DIR)
            },
            "files": {
                "template_exists": template_exists,
                "template_path": str(TEMPLATE_DOCX),
                "upload_dir_exists": upload_dir_exists,
                "upload_dir_writable": upload_writable,
                "output_dir_exists": output_dir_exists,
                "output_dir_writable": output_writable
            }
        }
    except Exception as e:
        logger.error(f"Error in root endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/upload-files/")
async def upload_files(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...)
):
    """Upload multiple documents (DOCX or PDF) for processing."""
    try:
        # Create a unique job ID
        job_id = str(uuid.uuid4())
        job_upload_dir = UPLOAD_DIR / job_id
        job_output_dir = OUTPUT_DIR / job_id
        
        # Create job directories
        job_upload_dir.mkdir(parents=True, exist_ok=True)
        job_output_dir.mkdir(parents=True, exist_ok=True)
        
        # Save uploaded files
        saved_files = []
        for file in files:
            try:
                # Validate file type
                if not (file.filename.endswith('.docx') or file.filename.endswith('.pdf')):
                    logger.warning(f"Skipping unsupported file: {file.filename}")
                    continue
                
                # Ensure safe filename
                safe_filename = Path(file.filename).name
                file_path = job_upload_dir / safe_filename
                
                # Save file using chunks to handle large files
                with file_path.open("wb") as buffer:
                    while chunk := await file.read(8192):
                        buffer.write(chunk)
                
                saved_files.append(safe_filename)
                logger.info(f"Successfully saved file: {safe_filename}")
                
            except Exception as e:
                logger.error(f"Error saving file {file.filename}: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Error saving file: {str(e)}")
        
        if not saved_files:
            raise HTTPException(status_code=400, detail="No supported files were uploaded (only .docx and .pdf are supported)")
        
        # Initialize job status
        processing_jobs[job_id] = {
            "status": "queued",
            "files": saved_files,
            "processed_files": [],
            "created_at": time.time(),
            "completed_at": None,
            "upload_dir": str(job_upload_dir),
            "output_dir": str(job_output_dir)
        }
        
        # Start processing in the background
        background_tasks.add_task(process_documents, job_id, str(job_upload_dir), str(job_output_dir))
        
        return {
            "job_id": job_id,
            "status": "queued",
            "message": f"Processing {len(saved_files)} files",
            "files": saved_files
        }
        
    except Exception as e:
        logger.error(f"Error in upload_files: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

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
        logger.info(f"Starting processing for job {job_id}")
        # Update job status
        processing_jobs[job_id]["status"] = "processing"
        
        processed_files = []
        
        for filename in os.listdir(upload_dir):
            input_path = os.path.join(upload_dir, filename)
            
            try:
                if filename.endswith(".docx"):
                    # Process DOCX directly
                    output_path = os.path.join(output_dir, filename)
                    logger.info(f"Processing DOCX file: {filename}")
                    apply_branding_to_docx(input_path, output_path)
                    processed_files.append(filename)
                    
                elif filename.endswith(".pdf"):
                    # Convert PDF to DOCX, then apply branding
                    base_name = os.path.splitext(filename)[0]
                    temp_docx = os.path.join(output_dir, f"{base_name}_converted.docx")
                    output_docx = os.path.join(output_dir, f"{base_name}.docx")
                    
                    logger.info(f"Converting PDF to DOCX: {filename}")
                    if convert_pdf_to_docx(input_path, temp_docx):
                        logger.info(f"Applying branding to converted file: {filename}")
                        apply_branding_to_docx(temp_docx, output_docx)
                        
                        # Clean up temporary DOCX
                        if os.path.exists(temp_docx):
                            os.remove(temp_docx)
                        
                        processed_files.append(f"{base_name}.docx")
            except Exception as e:
                logger.error(f"Error processing {filename}: {str(e)}")
                # Continue with other files even if one fails
        
        # Update job status to completed
        processing_jobs[job_id]["status"] = "completed"
        processing_jobs[job_id]["processed_files"] = processed_files
        processing_jobs[job_id]["completed_at"] = time.time()
        logger.info(f"Completed processing for job {job_id}")
        
    except Exception as e:
        # Update job status to failed
        logger.error(f"Job {job_id} failed: {str(e)}")
        processing_jobs[job_id]["status"] = "failed"
        processing_jobs[job_id]["error"] = str(e)

def cleanup_zip(zip_path):
    """Clean up the zip file after a delay"""
    time.sleep(300)  # Wait 5 minutes before deleting
    try:
        if os.path.exists(zip_path):
            os.remove(zip_path)
    except Exception as e:
        print(f"Error cleaning up zip file: {e}")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)