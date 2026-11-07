
from app import create_app
from waitress import serve

app = create_app()

if __name__ == "__main__":
    print("OM ENTERPRISE ERP is starting...")
    print("Open in browser: http://127.0.0.1:5000")

    serve(app, host="127.0.0.1", port=5000)
