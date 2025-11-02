import base64
import requests
import pandas as pd
from io import BytesIO
import streamlit as st

GITHUB_OWNER = "MKato2361"
GITHUB_REPO = "CI_FILES"

def get_pat() -> str:
    """PATをSecretsから取得（環境変数にも対応可）"""
    try:
        return st.secrets["GITHUB_PAT"]
    except Exception:
        raise ValueError("❌ GitHub PAT が設定されていません。secrets.toml または Streamlit Secrets に GITHUB_PAT を追加してください。")

def list_repo_files(folder_path: str = ""):
    """指定フォルダ内のCSV/XLSX/XLSファイル一覧を返す"""
    token = get_pat()
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{folder_path}"
    headers = {"Authorization": f"token {token}"}

    res = requests.get(url, headers=headers)
    if res.status_code != 200:
        raise Exception(f"GitHubファイル一覧取得エラー: {res.status_code} - {res.text}")

    files = res.json()
    return [f for f in files if f["type"] == "file" and f["name"].lower().endswith((".csv", ".xlsx", ".xls"))]

def load_file_from_github(file_info):
    """GitHubファイルを読み込み pandas.DataFrame で返す"""
    token = get_pat()
    headers = {"Authorization": f"token {token}"}
    res = requests.get(file_info["url"], headers=headers)

    if res.status_code != 200:
        raise Exception("ファイル取得に失敗しました")

    content = res.json()
    file_data = base64.b64decode(content["content"])

    if file_info["name"].lower().endswith(".csv"):
        return pd.read_csv(BytesIO(file_data))
    return pd.read_excel(BytesIO(file_data))
