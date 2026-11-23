"""Path resolver for the shared pipeline scripts.

Resolves paths from this file's location so the scripts work unchanged across
chats (each chat's session slug is different; all chats in the project share
this workspace folder).
"""
import os

# scripts/ → shared/ → _pipeline/ → workspace root
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_DIR  = os.path.dirname(SCRIPTS_DIR)
PIPELINE    = os.path.dirname(SHARED_DIR)
WORKSPACE   = os.path.dirname(PIPELINE)

TXN_DIR     = os.path.join(SHARED_DIR, "transactions")   # Square Transactions CSV exports
DATA_DIR    = os.path.join(SHARED_DIR, "data")
RAW_DIR     = os.path.join(SHARED_DIR, "raw")
