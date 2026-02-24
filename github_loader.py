# github_loader.py
import base64
from io import BytesIO
from typing import Dict, List, Optional
import requests
import streamlit as st

# ====== 設定（必要に応じて変更）======
GITHUB_OWNER = "MKato2361"
GITHUB_REPO = "CI_FILES"
GITHUB_API_BASE = "https://api.github.com"


# ====== ユーティリティ ======

def get_pat() -> str:
    try:
        return st.secrets["GITHUB_PAT"]
    except Exception:
        raise ValueError(
            "❌ GitHub PAT が見つかりません。 .streamlit/secrets.toml か"
            " デプロイ先のSecretsに GITHUB_PAT を設定してください。"
        )


def is_supported_file(name: str) -> bool:
    name_l = name.lower()
    return name_l.endswith(".csv") or name_l.endswith(".xlsx") or name_l.endswith(".xls")


def _headers() -> Dict[str, str]:
    return {"Authorization": f"token {get_pat()}"}


# ====== 読み取り系 ======

@st.cache_data(ttl=600)
def list_dir(path: str = "") -> List[Dict]:
    path = path.lstrip("/")
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
        except Exception:
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
def walk_repo_tree_with_dates(base_path: str = "", max_depth: int = 3) -> List[Dict]:
    """
    walk_repo_tree と同じツリー構造を返しつつ、
    ファイルノードには "updated" (YYYY-MM-DD) を付加する。
    Commits API はファイルごとに1回だが、結果はまとめて ttl=600 でキャッシュされる。

    返り値: [{name, path, type, depth, updated}, ...]
      - type=="dir"  のとき updated==""
      - type=="file" のとき updated=="2025-01-10" or "-"
    """
    nodes: List[Dict] = []

    def _walk(path: str, depth: int):
        if depth > max_depth:
            return
        try:
            items = list_dir(path)
        except Exception:
            return
        for it in items:
            node = {
                "name": it.get("name", ""),
                "path": it.get("path", ""),
                "type": it.get("type", ""),
                "depth": depth,
                "updated": "",
            }
            if it.get("type") == "file":
                try:
                    url = (
                        f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
                        f"/commits?path={it['path']}&per_page=1"
                    )
                    res = requests.get(url, headers=_headers())
                    if res.status_code == 200 and res.json():
                        raw = res.json()[0]["commit"]["committer"]["date"]
                        node["updated"] = raw[:10]  # "YYYY-MM-DD"
                    else:
                        node["updated"] = "-"
                except Exception:
                    node["updated"] = "-"
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


# ====== Admin UI 向け ======

def list_github_files(base_path: str = "") -> List[Dict]:
    """
    指定ディレクトリ直下のファイル一覧を返す（キャッシュなし）。
    admin UI の一覧表示・削除用。
    """
    clean = base_path.strip().strip("/")
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{clean}"
    res = requests.get(url, headers=_headers())
    if res.status_code == 404:
        return []
    if res.status_code != 200:
        raise Exception(f"GitHub list エラー {res.status_code}: {res.text}")
    data = res.json()
    return data if isinstance(data, list) else [data]


def get_dir_commit_dates(base_path: str = "") -> Dict[str, str]:
    """
    指定ディレクトリ配下の各ファイルの最終コミット日時を一括取得して返す。
    返り値: { "path/to/file.csv": "2025-01-10", ... }

    GitHub の /commits API はディレクトリ単位で絞り込めないため、
    対象ディレクトリの直下ファイルごとに per_page=1 で1件だけ取得する。
    ファイル数が多い場合は並列化を検討。
    """
    clean = base_path.strip().strip("/")
    result: Dict[str, str] = {}

    try:
        items = list_github_files(clean)
        file_paths = [it["path"] for it in items if it.get("type") == "file"]
    except Exception:
        return result

    for path in file_paths:
        try:
            url = (
                f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
                f"/commits?path={path}&per_page=1"
            )
            res = requests.get(url, headers=_headers())
            if res.status_code == 200 and res.json():
                raw = res.json()[0]["commit"]["committer"]["date"]  # ISO8601
                result[path] = raw[:10]  # "YYYY-MM-DD"
            else:
                result[path] = "-"
        except Exception:
            result[path] = "-"

    return result


def get_file_sha(target_path: str) -> Optional[str]:
    """
    指定パスのファイルが既に存在する場合は SHA を返す。
    存在しない場合は None を返す。
    上書きアップロード時に使用。
    """
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{target_path}"
    res = requests.get(url, headers=_headers())
    if res.status_code == 200:
        return res.json().get("sha")
    if res.status_code == 404:
        return None
    raise Exception(f"SHA取得エラー {res.status_code}: {res.text}")


def upload_file_to_github(
    target_path: str,
    content: bytes,
    message: str,
) -> Dict:
    """
    ファイルをアップロードする。
    - 既存ファイルが存在する場合は SHA を自動取得して上書き（PUT）。
    - 存在しない場合は新規作成。
    """
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{target_path}"

    payload: Dict = {
        "message": message,
        "content": base64.b64encode(content).decode("utf-8"),
    }

    # 既存SHAを取得して上書きモードにする
    existing_sha = get_file_sha(target_path)
    if existing_sha:
        payload["sha"] = existing_sha

    res = requests.put(url, headers=_headers(), json=payload)

    if res.status_code not in (200, 201):
        raise Exception(f"アップロードエラー {res.status_code}: {res.text}")

    return res.json()


def delete_file_from_github(
    target_path: str,
    sha: str,
    message: str,
) -> Dict:
    """
    指定パスのファイルを削除する。
    sha は list_github_files で取得した値を渡す。
    """
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{target_path}"

    payload = {
        "message": message,
        "sha": sha,
    }

    res = requests.delete(url, headers=_headers(), json=payload)

    if res.status_code not in (200, 204):
        raise Exception(f"削除エラー {res.status_code}: {res.text}")

    return res.json() if res.content else {}
