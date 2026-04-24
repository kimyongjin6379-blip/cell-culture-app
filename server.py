import os
import tempfile

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from processor import process_file

app = FastAPI(title="Cell Culture Analysis Tool")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def root():
    return FileResponse("static/index.html")


@app.post("/api/process")
async def process(
    file: UploadFile = File(...),
    basal_media: str = Form(""),
    feed_media: str = Form(""),
):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Excel 파일(.xlsx)만 업로드 가능합니다.")
    try:
        contents = await file.read()
        result = process_file(contents, basal_media=basal_media, feed_media=feed_media)
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"처리 중 오류 발생: {str(e)}")


@app.get("/api/download/{file_id}")
async def download(file_id: str):
    # Sanitize file_id to prevent path traversal
    if not all(c.isalnum() or c == "-" for c in file_id):
        raise HTTPException(status_code=400, detail="Invalid file ID")
    path = os.path.join(tempfile.gettempdir(), f"{file_id}.xlsx")
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="cell_culture_results.xlsx",
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
