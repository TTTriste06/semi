import requests
import base64
import streamlit as st
import pandas as pd
from io import BytesIO

from config import GITHUB_TOKEN_KEY, REPO_NAME, BRANCH


def upload_to_github(file, path_in_repo, commit_message):
    """
    上传文件到 GitHub 指定仓库与路径。
    """
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"

    file.seek(0)
    file_content = file.read()
    encoded_content = base64.b64encode(file_content).decode('utf-8')

    # 获取文件 SHA（如果已存在）
    response = requests.get(api_url, headers={"Authorization": f"token {st.secrets[GITHUB_TOKEN_KEY]}"})
    sha = response.json().get('sha') if response.status_code == 200 else None

    payload = {
        "message": commit_message,
        "content": encoded_content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha

    response = requests.put(api_url, json=payload, headers={"Authorization": f"token {st.secrets[GITHUB_TOKEN_KEY]}"})

    if response.status_code in [200, 201]:
        st.success(f"{path_in_repo} 上传成功！")
    else:
        st.error(f"上传失败: {response.json()}")


def download_excel_from_url(url, token=None):
    """
    从公开或私有的 raw URL 下载 Excel 文件。
    """
    headers = {"Authorization": f"token {token}"} if token else {}
    response = requests.get(url, headers=headers)

    if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' not in response.headers.get('Content-Type', ''):
        raise ValueError("下载的不是 Excel 文件，请检查链接或权限")

    return pd.read_excel(BytesIO(response.content))

def download_excel_from_repo(filename, show_warning=True):
    """
    从 GitHub 仓库中下载 Excel 文件（支持私有 repo），使用 GitHub API。
    """
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{filename}"
    headers = {"Authorization": f"token {st.secrets[GITHUB_TOKEN_KEY]}"}
    response = requests.get(api_url, headers=headers)

    if response.status_code != 200:
        if show_warning:
            st.warning(f"⚠️ 无法下载 {filename}，返回码 {response.status_code}")
        return pd.DataFrame()

    try:
        content = base64.b64decode(response.json()['content'])
        df = pd.read_excel(BytesIO(content))
    except Exception as e:
        if show_warning:
            st.warning(f"⚠️ 解析 {filename} 失败：{e}")
        return pd.DataFrame()

    return df
