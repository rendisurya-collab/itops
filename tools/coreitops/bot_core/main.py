from fastapi import FastAPI

# Inisialisasi ini wajib ada:
app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "Hello World"}