# Prompt 1 of this project


"""
EIS API Forge
──────────────────
Upload JSON spec + Excel test cases · Select EIS / NON-EIS ·
LLM mutates payloads · EIS encryption pipeline · Dashboard + export
"""

import streamlit as st
import pandas as pd
import json, time, re, uuid, requests, warnings
from datetime import datetime
from io import BytesIO

warnings.filterwarnings("ignore")  # suppress SSL warnings in console
