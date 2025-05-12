import requests
import base64
import streamlit as st
import pandas as pd
from io import BytesIO

from config import GITHUB_TOKEN_KEY, REPO_NAME, BRANCH, MAPPING_COLUMNS


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


def download_excel_from_github(url, token=None):
    """
    从 GitHub 原始地址下载 Excel 文件。
    """
    headers = {"Authorization": f"token {token}"} if token else {}
    response = requests.get(url, headers=headers)
    content_type = response.headers.get('Content-Type', '')

    if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' not in content_type:
        raise ValueError("下载的不是 Excel 文件，请检查链接或权限")

    return pd.read_excel(BytesIO(response.content))


def download_backup_file(file_name):
    """
    下载 repo 根目录中的 Excel 备份文件（用于缺失处理）。
    """
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{file_name}"
    headers = {"Authorization": f"token {st.secrets[GITHUB_TOKEN_KEY]}"}
    response = requests.get(api_url, headers=headers)

    if response.status_code != 200:
        st.warning(f"⚠️ 无法下载 {file_name}，GitHub 返回码 {response.status_code}")
        return pd.DataFrame()

    content = response.json().get('content')
    if not content:
        st.warning(f"⚠️ {file_name} 文件内容为空或解析失败")
        return pd.DataFrame()

    try:
        df = pd.read_excel(BytesIO(base64.b64decode(content)))
    except Exception as e:
        st.warning(f"⚠️ {file_name} 解析 Excel 失败：{e}")
        return pd.DataFrame()

    return df


def download_mapping_from_github(path_in_repo="mapping_file.xlsx"):
    """
    下载并预处理新旧料号文件。
    """
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"
    response = requests.get(api_url, headers={"Authorization": f"token {st.secrets[GITHUB_TOKEN_KEY]}"})
    if response.status_code != 200:
        st.warning("GitHub 上找不到 mapping_file.xlsx，用默认表或请先上传")
        return pd.DataFrame(columns=MAPPING_COLUMNS)

    content = base64.b64decode(response.json()['content'])
    df = pd.read_excel(BytesIO(content))

    if df.shape[1] >= 6:
        df = df.iloc[:, :6]
        df.columns = MAPPING_COLUMNS
    else:
        st.warning(f"mapping_file.xlsx 列数不足：只发现 {df.shape[1]} 列，需要至少 6 列")
        df = pd.DataFrame(columns=MAPPING_COLUMNS)

    return df
