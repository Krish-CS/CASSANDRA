"""
WSGI entry point for production deployments (Render, Gunicorn)
"""

from app import app

if __name__ == "__main__":
    app.run()
