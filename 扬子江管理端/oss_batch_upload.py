"""
OSS 批量并发上传脚本
原流程：Apifox 手动 10 步 → 每文件串行
本脚本：一键并发上传，N 个文件同时跑
"""
import os
import sys
import json
import time
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

# ============================================================
# 配置区 —— 改这里就行
# ============================================================
BASE_URL = "https://uat-yzj-admin-kb-api.nullht.com"
KB_SESSION_ID = input("请输入 KB_SESSION_ID: ").strip()
AUTHORIZATION = input("请输入 authorization: ").strip()
PROJECT_ID = input("请输入 project_id: ").strip()
LOCAL_DIR = input("请输入文件夹路径（默认 D:\\yzj_test\\）: ").strip() or r"D:\yzj_test"
DIR_ID = input("请输入 dir_id（默认 0）: ").strip() or "0"
MAX_WORKERS = 3  # 并发数，公司内网服务器别太高

# 公共请求头
HEADERS = {
    "Cookie": f"KB_SESSION_ID={KB_SESSION_ID}",
    "Authorization": AUTHORIZATION,
    "Content-Type": "application/json",
    "User-Agent": "OSS-BatchUpload/1.0",
}

# 禁用 SSL 警告（公司内网自签证书场景）
requests.packages.urllib3.disable_warnings()


def log(msg: str, level: str = "INFO"):
    """带时间戳的日志"""
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def api_post(path: str, body: object, timeout: int = 60) -> dict:
    """封装 POST 请求，遇连接重置自动重试"""
    last_error = None
    for retry in range(3):
        try:
            resp = requests.post(
                f"{BASE_URL}{path}",
                headers=HEADERS,
                json=body,
                verify=False,
                timeout=timeout,
            )
            resp.raise_for_status()
            try:
                return resp.json()
            except json.JSONDecodeError:
                return {"raw_status": resp.status_code, "text": resp.text}
        except (requests.ConnectionError, requests.Timeout,
                requests.HTTPError) as e:
            last_error = e
            if retry < 2:
                wait = (retry + 1) * 3
                log(f"  POST {path} 失败 (第{retry+1}次)，{wait}s 后重试: {e}", "WARN")
                time.sleep(wait)
    raise last_error


def scan_files(dir_path: str) -> list:
    """扫描目录，返回文件绝对路径列表，过滤系统文件"""
    files = []
    for f in sorted(os.listdir(dir_path)):
        if f == ".DS_Store":
            continue
        full = os.path.join(dir_path, f)
        if os.path.isfile(full):
            files.append(full)
    return files


def build_file_id(index: int) -> str:
    """生成 fileId，格式与 Apifox 脚本一致"""
    return f"file-{int(time.time() * 1000)}-{index}"


def create_upload_tasks(file_paths: list) -> list:
    """步骤1：批量创建上传任务"""
    log(f"创建 {len(file_paths)} 个上传任务...")
    timestamp = int(time.time() * 1000)
    task_list = []
    for idx, path in enumerate(file_paths):
        name = os.path.basename(path)
        task_list.append({
            "dirId": DIR_ID,
            "fileId": build_file_id(idx),
            "name": name,
            "projectId": PROJECT_ID,
            "status": "IDLE",
            "createdTime": timestamp,
            "fileKey": "",
        })
    result = api_post("/api/uploadTask/list", task_list)
    if result.get("msg") != "success":
        raise RuntimeError(f"创建上传任务失败: {result}")
    log(f"  创建成功，共 {len(task_list)} 个任务")
    return task_list


def upload_single_file(file_path: str, task: dict, index: int) -> dict:
    """上传单个文件的完整流程（步骤2-5），每个文件独立执行"""
    file_name = os.path.basename(file_path)
    log(f"  [{index}] 开始处理: {file_name}")

    try:
        # 步骤2：获取预签名上传链接
        sig_resp = api_post(
            "/api/oss/generate_pre_upload_signature",
            {"filenames": [file_name]},
        )
        sig_data = sig_resp["data"][0]
        upload_url = sig_data["url"]
        file_key = sig_data["file_key"]

        # 步骤3：PUT 文件到 OSS（文件名含特殊字符可能导致连接被拒，加重试）
        log(f"  [{index}] 预签名 URL 前60字符: {upload_url[:60]}...")
        for retry in range(3):
            try:
                with open(file_path, "rb") as fh:
                    resp = requests.put(
                        upload_url,
                        data=fh,
                        headers={"Content-Type": "application/octet-stream"},
                        verify=False,
                        timeout=300,
                    )
                    resp.raise_for_status()
                break
            except Exception as e:
                log(f"  [{index}] OSS PUT 失败 (第{retry+1}次): {e}", "WARN")
                if retry == 2:
                    raise
                time.sleep(3)
        log(f"  [{index}] OSS 上传完成: {file_name}")

        # 步骤4：注册上传文件
        api_post("/api/kb/file/upload", {
            "files": [{
                "file_key": file_key,
                "filename": file_name,
                "type": "FILE",
                "is_replace": False,
                "uploadTaskFileId": task["fileId"],
            }],
            "project_id": PROJECT_ID,
            "dir_id": DIR_ID,
            "source_type": "LOCAL_UPLOAD",
        })

        # 步骤5：更新任务状态
        api_post("/api/uploadTask/status", [{
            "fileId": task["fileId"],
            "status": "SUCCESS",
        }])

        log(f"  [{index}] 注册完成: {file_name}")
        return {"file_name": file_name, "file_key": file_key, "file_id": task["fileId"], "success": True}

    except Exception as e:
        log(f"  [{index}] 失败: {file_name} — {e}", "ERROR")
        return {"file_name": file_name, "success": False, "error": str(e)}


def push_file_status(file_ids: list) -> None:
    """步骤6：推送文件状态（接口响应慢，加长超时 + 重试）"""
    log(f"推送状态，文件数: {len(file_ids)}")
    for attempt in range(3):
        try:
            api_post("/api/kb/status", {
                "source_type": "LOCAL_UPLOAD",
                "file_ids": file_ids,
            }, timeout=120)
            return
        except Exception as e:
            log(f"  推送状态失败 (第{attempt+1}次): {e}", "WARN")
            time.sleep(5)
    log("  推送状态最终失败，但文件已上传成功，可手动推送", "WARN")


def wait_for_conversion(project_id: str, timeout: int = 120) -> str:
    """步骤7-9：轮询等待文件转换完成，返回首个文件的 cover_key"""
    log("等待文件转换完成...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        resp = api_post("/api/kb/file/list", {
            "project_id": project_id,
            "dir_id": DIR_ID,
            "name": "",
            "page_num": 1,
            "page_size": 100000,
            "tag_value_ids": [],
        })
        rows = resp.get("data", {}).get("rows", [])
        if rows:
            status = rows[0].get("convert_status")
            cover = rows[0].get("cover")
            log(f"  转换状态: {status}")
            if status == "COMPLETED":
                log("  全部转换完成!")
                return cover
        time.sleep(3)
    raise TimeoutError("转换超时")


def generate_download_url(cover_key: str) -> str:
    """步骤10：生成预签名下载链接"""
    resp = api_post("/api/oss/generate_pre_download_signature", {
        "file_keys": [cover_key],
    })
    return resp["data"][0]["url"]


# ============================================================
# 主流程
# ============================================================
def main():
    log("=".join([""] * 40))
    log("OSS 批量并发上传启动")

    # 0. 扫描文件
    file_paths = scan_files(LOCAL_DIR)
    if not file_paths:
        log("没有找到文件，请检查路径", "ERROR")
        sys.exit(1)
    log(f"扫描到 {len(file_paths)} 个文件")

    # 1. 创建上传任务
    tasks = create_upload_tasks(file_paths)

    # 2-5. 并发上传
    log(f"开始并发上传，并发数: {MAX_WORKERS}")
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = []
        for i, (fp, task) in enumerate(zip(file_paths, tasks), 1):
            fut = pool.submit(upload_single_file, fp, task, i)
            futures.append(fut)
            time.sleep(0.5)  # 错峰提交，避免同时打服务器
        for future in as_completed(futures):
            results.append(future.result())

    # 统计结果
    success = [r for r in results if r["success"]]
    failed = [r for r in results if not r["success"]]
    log(f"上传完成: 成功 {len(success)}, 失败 {len(failed)}")
    for f in failed:
        log(f"  失败: {f['file_name']} — {f['error']}", "WARN")

    if not success:
        log("全部失败，终止", "ERROR")
        sys.exit(1)

    # 6. 推送文件状态
    push_file_status([r["file_id"] for r in success])

    # 7-9. 等待转换
    cover_key = wait_for_conversion(PROJECT_ID)

    # 10. 生成下载链接
    if cover_key:
        dl_url = generate_download_url(cover_key)
        log(f"下载链接: {dl_url}")

    log("全部完成!")


if __name__ == "__main__":
    main()
