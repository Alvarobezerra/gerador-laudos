import sqlite3
import json
from datetime import datetime
from fastapi import FastAPI, HTTPException, Depends
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os

app = FastAPI(title="Laudo API")

# Setup Auth
security = HTTPBasic()
def verify_auth(credentials: HTTPBasicCredentials = Depends(security)):
    if not ((credentials.username.lower() in ["alvaro", "perito"]) and credentials.password in ["123", "icrim123"]):
        raise HTTPException(
            status_code=401,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username

# Allow CORS for development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_PATH = os.path.join(os.path.dirname(__file__), "banco_laudos.sqlite")

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS laudos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mobile_id TEXT UNIQUE,
            data_sincronizacao TEXT,
            dados_json TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

class LaudoPayload(BaseModel):
    mobile_id: str
    dados: dict

@app.post("/api/laudos")
def sync_laudo(payload: LaudoPayload, username: str = Depends(verify_auth)):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Check if exists
    c.execute('SELECT id FROM laudos WHERE mobile_id = ?', (payload.mobile_id,))
    existing = c.fetchone()
    
    now = datetime.now().isoformat()
    json_data = json.dumps(payload.dados)
    
    if existing:
        c.execute('UPDATE laudos SET dados_json = ?, data_sincronizacao = ? WHERE mobile_id = ?',
                 (json_data, now, payload.mobile_id))
    else:
        c.execute('INSERT INTO laudos (mobile_id, data_sincronizacao, dados_json) VALUES (?, ?, ?)',
                 (payload.mobile_id, now, json_data))
                 
    conn.commit()
    conn.close()
    return {"status": "ok", "message": "Sincronizado com sucesso"}

@app.get("/api/laudos")
def get_laudos(username: str = Depends(verify_auth)):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT id, mobile_id, data_sincronizacao, dados_json FROM laudos ORDER BY data_sincronizacao DESC')
    rows = c.fetchall()
    conn.close()
    
    result = []
    for r in rows:
        result.append({
            "id": r["id"],
            "mobile_id": r["mobile_id"],
            "data_sincronizacao": r["data_sincronizacao"],
            "dados": json.loads(r["dados_json"])
        })
    return result

# Serve static files for PWA (fallback to prevent crash if dir doesn't exist yet)
mobile_dir = os.path.join(os.path.dirname(__file__), "mobile")
if not os.path.exists(mobile_dir):
    os.makedirs(mobile_dir)
    
app.mount("/", StaticFiles(directory=mobile_dir, html=True), name="mobile")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True)
