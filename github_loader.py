# github_loader.py
import base64
from io import BytesIO
from typing import Dict, List

import requests
import streamlit as st

# ====== 設定（必要に応じて変更）======
GITHUB_OWNER = "MKato2361"        # 例: "my-org" or "my-user"
GITHUB_REPO = "CI_FILES"       # 例: "my-private-repo"
GITHUB_API_BASE = "https://api.github.com"

# ====== ユーティリティ ======
def get_pat() -> str:
    try:
        return st.secrets["GITHUB_PAT"]
    except Exception:
        raise ValueError(
            "❌ GitHub PAT が見つかりません。 .streamlit/secrets.toml か デプロイ先のSecretsに GITHUB_PAT を設定してください。"
        )

def is_supported_file(name: str) -> bool:
    name_l = name.lower()
    return name_l.endswith(".csv") or name_l.endswith(".xlsx") or name_l.endswith(".xls")

def _headers() -> Dict[str, str]:
    return {"Authorization": f"token {get_pat()}"}

@st.cache_data(ttl=600)
def list_dir(path: str = "") -> List[Dict]:
    path = path.lstrip("/")  # ✅ 余分な削除のみ
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{path}"
    res = requests.get(url, headers=_headers())
    if res.status_code != 200:
        raise Exception(f"GitHub list_dir エラー {res.status_code}: {res.text}")
    data = res.json()
    return data if isinstance(data, list) else [data]


@st.cache_data(ttl=600)
def walk_repo_tree(base_path: str = "", max_depth: int = 3) -> List[Dict]:
    """
    再帰的に /contents を辿って最大max_depthまでのノードを返す。
    返り値: [{name, path, type, depth}, ...] typeは "dir" or "file"
    """
    nodes: List[Dict] = []

    def _walk(path: str, depth: int):
        if depth > max_depth:
            return
        try:
            items = list_dir(path)
        except Exception as e:
            # 404等はスキップ
            return
        for it in items:
            node = {
                "name": it.get("name", ""),
                "path": it.get("path", ""),
                "type": it.get("type", ""),
                "depth": depth,
            }
            nodes.append(node)
            if it.get("type") == "dir":
                _walk(it.get("path", ""), depth + 1)

    _walk(base_path.strip("/"), 0)
    return nodes

@st.cache_data(ttl=600)
def load_file_bytes_from_github(path: str) -> BytesIO:
    """指定パスのコンテンツ（Base64）を取得して BytesIO で返す。"""
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{path}".rstrip("/")
    res = requests.get(url, headers=_headers())
    if res.status_code != 200:
        raise Exception(f"❌ GitHubファイル取得失敗 {res.status_code}: {res.text}")
    content = res.json()
    file_data = base64.b64decode(content["content"])
    return BytesIO(file_data)
