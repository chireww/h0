"""
AZ 批量新建文件审核（dev3-病例分享材料） — 自动化测试脚本
==========================================================
将 Apifox 测试场景 "AZ_批量新建文件审核(dev3-病例分享材料)" 转换为 Python。

功能: 批量上传文件到 AZ 审核系统并创建审核任务，收集所有审核结果。

流程:
  1. 登录 → 初始化结果集 → 扫描文件目录
  2. Loop 遍历每个文件:
     a. 上传文件 → 获取 file_id
     b. 若上传成功 → 创建审核任务 → 获取 task_id
     c. 等待 5 秒（模拟原脚本 delay）
     d. 获取审核详情 → 收集 {id, task_id, file_name}
  3. 输出最终结果 JSON

特点:
  - 不做金标准对比，仅批量创建任务并收集结果
  - product_Ids 支持动态变量 (drug_type)
  - 索引式顺序处理，可断点续传

用法:
  python az_batch_create_audit.py
"""

import os
import sys
import json
import time
import logging
from pathlib import Path
from typing import Any

import requests

# ============================================================
# 配置
# ============================================================
CONFIG = {
    "az_base_url": os.getenv("AZ_BASE_URL", "https://dev-api-v3-az-mlr.nullht.com/api"),
    "az_email": os.getenv("AZ_EMAIL", "1@qq.com"),
    "az_password": os.getenv("AZ_PASSWORD", "123456"),
    "file_dir": os.getenv("AZ_FILE_DIR", r"D:\AZ_test_0"),
    # 药物类型 — 原脚本用 {{drug_type}} 变量
    "drug_type": os.getenv("AZ_DRUG_TYPE", "f7ce1e4f-5ca7-11f0-8e6d-00163e36469b,f7ce0c74-5ca7-11f0-8e6d-00163e36469b"),
    # 审核请求默认参数（dev3-病例分享材料）
    "audit_defaults": {
        "literature_package": [],
        "category": "NS",
        "material_properties": [2],
        "file_type_level_1": [73],
        "file_type_level_2": [66],
        "file_type_level_3": [29],
        "medical_education_sub_categories": [4],
        "target_audience": [21],
    },
    # 创建任务后等待时间（秒），对应原脚本 5s delay
    "audit_delay": 5,
    "verify_ssl": False,
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("az_batch")


# ============================================================
# AZ 客户端
# ============================================================
class AZClient:
    def __init__(self, config: dict):
        self.base_url = config["az_base_url"].rstrip("/")
        self.verify_ssl = config["verify_ssl"]
        self.audit_defaults = config["audit_defaults"]
        self.drug_types = [
            d.strip()
            for d in config["drug_type"].split(",")
            if d.strip()
        ]
        self.session = requests.Session()
        self.session.verify = self.verify_ssl
        self.session.headers.update({
            "User-Agent": "AZ-Batch-Audit/1.0 (Python)",
            "Content-Type": "application/json",
        })
        self.token: str | None = None

    @property
    def auth_header(self) -> dict:
        return {"Authorization": self.token} if self.token else {}

    # ----- 登录 -----
    def login(self, email: str, password: str) -> str:
        url = f"{self.base_url}/auth/email/login"
        body = {"email": email, "password": password}
        log.info(f"🔑 登录: {email}")
        resp = self.session.post(url, json=body)
        resp.raise_for_status()
        data = resp.json()
        token = data.get("data", {}).get("token")
        if not token:
            raise RuntimeError(f"登录失败: {resp.text[:300]}")
        self.token = f"Bearer {token}"
        log.info("✅ 登录成功")
        return self.token

    # ----- 上传文件 -----
    def upload_file(self, file_path: str) -> tuple[str | None, bool]:
        """
        上传文件，返回 (file_id, success)。
        对应原脚本: 响应 code==0 时 upload_success=true
        """
        url = f"{self.base_url}/oss/upload"
        file_name = Path(file_path).name
        log.info(f"📤 上传: {file_name}")
        try:
            with open(file_path, "rb") as f:
                resp = self.session.post(
                    url, headers=self.auth_header,
                    files={"file": (file_name, f)},
                )
            resp.raise_for_status()
            result = resp.json()

            # 原脚本逻辑: code==0 才算成功
            if result.get("code") == 0:
                file_id = result.get("data")
                log.info(f"   ✅ file_id={file_id}")
                return file_id, True
            else:
                log.warning(f"   ⚠️ 上传失败: code={result.get('code')} msg={result.get('message', '')}")
                return None, False
        except requests.RequestException as e:
            log.error(f"   ❌ 上传异常: {e}")
            return None, False

    # ----- 创建审核任务 -----
    def create_audit_task(self, file_id: str, file_name: str) -> str | None:
        """
        创建审核任务，返回 task_id。
        原脚本 product_Ids 使用 {{drug_type}} 变量（可能是逗号分隔的多个 ID）。
        """
        url = f"{self.base_url}/audit/management/add"
        body = {
            "file_id": file_id,
            "file_name": file_name,
            "product_Ids": self.drug_types,
            **self.audit_defaults,
        }
        log.info(f"📝 创建审核: {file_name}")
        try:
            resp = self.session.post(url, json=body, headers=self.auth_header)
            resp.raise_for_status()
            task_id = resp.json().get("data")
            if task_id:
                log.info(f"   ✅ task_id={task_id}")
                return task_id
            log.warning(f"   ⚠️ 空响应: {resp.text[:200]}")
            return None
        except requests.RequestException as e:
            log.error(f"   ❌ 创建失败: {e}")
            return None

    # ----- 获取审核详情 -----
    def get_audit_detail(self, task_id: str) -> dict | None:
        """获取审核详情，返回完整 JSON 响应。"""
        url = f"{self.base_url}/audit/management/detail"
        body = {"id": task_id}
        try:
            resp = self.session.post(url, json=body, headers=self.auth_header)
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException as e:
            log.error(f"   ❌ 获取详情失败 (id={task_id}): {e}")
            return None


# ============================================================
# 文件扫描
# ============================================================
def scan_files(directory: str) -> list[dict]:
    """
    扫描目录，返回 [{"path": ..., "name": ...}, ...]。
    对应原脚本 get_files.py 的返回格式: {success, directory, files, file_count}
    """
    p = Path(directory)
    if not p.exists():
        raise FileNotFoundError(f"目录不存在: {directory}")

    files = []
    for fp in sorted(p.iterdir()):  # 排序保证顺序一致
        if fp.is_file() and fp.name != ".DS_Store":
            files.append({
                "path": str(fp.resolve()),
                "name": fp.name,
            })

    log.info(f"📂 扫描完成: {len(files)} 个文件 (目录: {directory})")
    return files


# ============================================================
# 主流程
# ============================================================
def main():
    cfg = CONFIG

    # ===== Step 1: 初始化 =====
    log.info("=" * 60)
    log.info("AZ 批量新建文件审核 (dev3-病例分享材料)")
    log.info("=" * 60)

    # 原脚本 pre-request: 初始化空结果集
    audit_result_json: list[dict] = []
    log.info("✅ 初始化完成，准备开始收集数据...")

    # 扫描文件
    files = scan_files(cfg["file_dir"])
    if not files:
        log.error("❌ 没有找到待审核文件")
        sys.exit(1)

    total_files = len(files)
    log.info(f"📋 共 {total_files} 个文件待处理")

    # 登录
    az = AZClient(cfg)
    az.login(cfg["az_email"], cfg["az_password"])

    # ===== Step 2: Loop 遍历每个文件 =====
    log.info("=" * 60)
    log.info("开始循环处理文件...")
    log.info("=" * 60)

    for index, file_info in enumerate(files):
        file_path = file_info["path"]
        file_name = file_info["name"]

        log.info(f"\n--- [{index + 1}/{total_files}] {file_name} ---")

        # a. 上传文件
        file_id, upload_success = az.upload_file(file_path)

        # b. 若上传成功 → 创建审核任务
        if not upload_success or not file_id:
            log.warning(f"⏭️ 跳过 {file_name}（上传失败）")
            continue

        task_id = az.create_audit_task(file_id, file_name)
        if not task_id:
            log.warning(f"⏭️ {file_name} 上传成功但创建任务失败")
            continue

        # c. 等待 5 秒（对应原脚本 delay 步骤）
        log.info(f"⏳ 等待 {cfg['audit_delay']}s...")
        time.sleep(cfg["audit_delay"])

        # d. 获取审核详情 → 收集结果
        detail = az.get_audit_detail(task_id)
        if detail and detail.get("data"):
            d = detail["data"]
            result_item = {
                "id": d.get("id"),
                "task_id": d.get("task_id"),
                "file_name": d.get("file_name"),
            }
            audit_result_json.append(result_item)
            log.info(f"📋 已记录: {d.get('file_name')} (id={d.get('id')})")
        else:
            log.warning(f"⚠️ 无法获取 {file_name} 的审核详情")

    # ===== Step 3: 输出最终结果 =====
    log.info("")
    log.info("=" * 60)
    log.info("📊 最终执行结果 (JSON)")
    log.info("=" * 60)
    output = json.dumps(audit_result_json, ensure_ascii=False, indent=4)
    print(output)
    log.info("=" * 60)
    log.info(f"共收集到 {len(audit_result_json)} 条数据。")

    # 保存到文件
    output_path = "batch_audit_results.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(audit_result_json, f, ensure_ascii=False, indent=4)
    log.info(f"结果已保存至: {output_path}")

    sys.exit(0 if audit_result_json else 1)


if __name__ == "__main__":
    main()
