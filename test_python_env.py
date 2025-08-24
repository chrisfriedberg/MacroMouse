#!/usr/bin/env python3
"""
Test script to check Python environment and google-cloud-storage availability.
"""

import sys
import os

print("=== Python Environment Test ===")
print(f"Python executable: {sys.executable}")
print(f"Python version: {sys.version}")
print(f"Python path: {sys.path[0]}")

print("\n=== Testing google-cloud-storage import ===")
try:
    from google.cloud import storage
    print("google-cloud-storage import successful")
    print(f"   Version: {storage.__version__}")
except ImportError as e:
    print("google-cloud-storage import failed: {e}")
    print("   This means the package is not available in this Python environment")

print("\n=== Testing other imports ===")
try:
    import requests
    print("requests import successful")
except ImportError as e:
    print("requests import failed: {e}")

try:
    import customtkinter
    print("customtkinter import successful")
except ImportError as e:
    print("customtkinter import failed: {e}")

print("\n=== Environment variables ===")
print(f"PYTHONPATH: {os.environ.get('PYTHONPATH', 'Not set')}")
print(f"VIRTUAL_ENV: {os.environ.get('VIRTUAL_ENV', 'Not set')}")

print("\n=== Done ===")
