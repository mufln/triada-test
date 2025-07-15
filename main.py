import json
import uuid

import pandas as pd
import uvicorn
from fastapi import FastAPI, UploadFile
from starlette.requests import Request
from starlette.responses import FileResponse, Response

app = FastAPI()
file = open("index.html", 'r', encoding='utf-8')
html = file.read()

@app.get("/")
async def index(request: Request):
    base_url = request.url
    response = (html.replace("marker-json-to-xlsx", base_url.__str__() + "json-to-xlsx")
                .replace("marker-xlsx-to-json", base_url.__str__() + "xlsx-to-json"))
    return Response(content=response, media_type="text/html")


@app.post("/json-to-xlsx")
async def read_json(file: UploadFile):
    name = uuid.uuid4().hex
    data = json.loads(await file.read())
    if isinstance(data, list):
        if all(isinstance(item, dict) for item in data):
            df = pd.json_normalize(data)
        else:
            df = pd.DataFrame(data, columns=['data'])
    elif isinstance(data, dict):
        df = pd.json_normalize(data)
    else:
        df = pd.DataFrame([data])
    df.to_excel(f"{name}.xlsx")
    return FileResponse(path=f"{name}.xlsx")


@app.post("/xlsx-to-json")
async def read_excel(file: UploadFile):
    name = uuid.uuid4().hex
    df = pd.read_excel(await file.read())
    df = df.to_json(orient="records")
    json.dump(json.loads(df), open(f"{name}.json", "w"))
    return FileResponse(path=f"{name}.json")


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
