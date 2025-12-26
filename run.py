#!/usr/bin/env python3
"""
Simple script to run the Research Data Collection System
"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run(
        "app.main:app",
        host="127.0.0.1",
        port=8000,
        reload=True
    )
# Updated 2026-01-14
