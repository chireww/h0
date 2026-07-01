"""
百济患教 — 文档入库自动化脚本
================================
将 Apifox 测试场景 "文档入库" 转换为 Python。

流程:
  1. 生成 3 个预签名上传 URL（文章 PDF、封面图、参考 PDF）
  2. 使用预签名 URL 上传文件到 OSS
  3. 创建文章记录（关联所有 fileKey 和元数据）

用法:
  python baiji_document_import.py --filename "文章标题" --article-pdf ./doc.pdf --cover ./cover.png --reference-pdf ./ref.pdf --reference-link "https://..." --tags '[1,2,3]'
"""

import os
import sys
import json
import argparse
import logging
from pathlib import Path
from typing import Tuple

import requests

# ============================================================
# 配置
# ============================================================
CONFIG = {
    "base_url": os.getenv("BAIJI_BASE_URL", "https://dev-api-baiji-patient-edu.nullht.com/api"),
    "openid": os.getenv("BAIJI_OPENID", ""),
    "verify_ssl": False,
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("baiji_import")


# ============================================================
# 客户端
# ============================================================
class BaijiClient:
    """百济患教 API 客户端"""

    def __init__(self, config: dict):
        self.base_url = config["base_url"].rstrip("/")
        self.verify_ssl = config["verify_ssl"]
        self.session = requests.Session()
        self.session.verify = self.verify_ssl
        self.session.headers.update({
            "User-Agent": "Baiji-Import/1.0 (Python)",
            "Content-Type": "application/json",
        })

    # ----- Step 1-3: 生成预签名上传 URL -----
    def get_presigned_url(self, file_name: str, file_type: str) -> Tuple[str, str]:
        """
        POST /file/presigned-url
        返回 (upload_url, file_key)

        file_type: "article" | "cover"
        """
        url = f"{self.base_url}/file/presigned-url"
        body = {"fileName": file_name, "fileType": file_type}
        log.info(f"🔗 获取预签名 URL: {file_name} (type={file_type})")
        resp = self.session.post(url, json=body)
        resp.raise_for_status()
        data = resp.json()
        upload_url = data.get("data", {}).get("uploadUrl")
        file_key = data.get("data", {}).get("fileKey")
        if not upload_url or not file_key:
            raise RuntimeError(f"获取预签名 URL 失败: {resp.text[:300]}")
        log.info(f"   ✅ uploadUrl={upload_url[:80]}...")
        log.info(f"   ✅ fileKey={file_key}")
        return upload_url, file_key

    # ----- Step 4-6: 上传文件到 OSS -----
    def upload_to_oss(self, upload_url: str, file_path: str, file_type_label: str) -> None:
        """
        PUT 上传文件到 OSS 预签名 URL。
        二进制上传，不设置 Content-Type 让 OSS 自动处理。
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        file_size = os.path.getsize(file_path)
        log.info(f"📤 上传 {file_type_label}: {Path(file_path).name} ({file_size} bytes)")

        with open(file_path, "rb") as f:
            resp = requests.put(
                upload_url,
                data=f,
                headers={"Content-Type": "application/octet-stream"},
            )

        if resp.status_code in (200, 201, 204):
            log.info(f"   ✅ 上传成功")
        else:
            log.warning(f"   ⚠️ 上传返回 {resp.status_code}: {resp.text[:200]}")

    # ----- Step 7: 创建文章记录 -----
    def create_article(
        self,
        title: str,
        article_name: str,
        article_key: str,
        cover_key: str,
        reference_key: str,
        reference_link: str,
        tag_ids: list,
    ) -> dict:
        """
        POST /articles
        创建文章记录，返回响应 JSON。
        """
        url = f"{self.base_url}/articles"
        body = {
            "title": title,
            "cover_key": cover_key,
            "article_name": article_name,
            "article_key": article_key,
            "reference_link": reference_link,
            "reference_key": reference_key,
            "tag_ids": tag_ids,
        }
        log.info(f"📝 创建文章: {title}")
        log.info(f"   body: {json.dumps(body, ensure_ascii=False, indent=2)}")

        resp = self.session.post(url, json=body)
        resp.raise_for_status()
        result = resp.json()

        if result.get("message") == "success":
            log.info("   ✅ 创建成功")
        else:
            log.warning(f"   ⚠️ 响应 message={result.get('message', '')}")
        return result


# ============================================================
# 主流程
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="百济患教 — 文档入库")
    parser.add_argument("--title", required=True, help="文章标题（也用作文件名前缀）")
    parser.add_argument("--article-pdf", required=True, help="文章 PDF 文件路径")
    parser.add_argument("--cover", required=True, help="封面图片路径 (.png)")
    parser.add_argument("--reference-pdf", required=True, help="参考 PDF 文件路径")
    parser.add_argument("--reference-link", default="", help="公众号原文链接")
    parser.add_argument("--tags", default="[]", help="Tag IDs (JSON 数组), 如 '[1,2,3]'")
    parser.add_argument("--delay", type=float, default=2.0, help="创建文章前等待秒数 (默认 2s)")
    args = parser.parse_args()

    cfg = CONFIG
    client = BaijiClient(cfg)

    title = args.title
    tag_ids = json.loads(args.tags)

    log.info("=" * 60)
    log.info(f"百济患教 — 文档入库: {title}")
    log.info("=" * 60)

    # ===== Step 1: 预签名 URL — 文章 PDF =====
    article_url, article_key = client.get_presigned_url(f"{title}.pdf", "article")

    # ===== Step 2: 预签名 URL — 封面图 =====
    cover_url, cover_key = client.get_presigned_url(f"{title}.png", "cover")

    # ===== Step 3: 预签名 URL — 参考 PDF =====
    reference_url, reference_key = client.get_presigned_url(f"{title}.pdf", "article")

    # ===== Step 4: 上传文章 PDF 到 OSS =====
    client.upload_to_oss(article_url, args.article_pdf, "文章PDF")

    # ===== Step 5: 上传封面到 OSS =====
    client.upload_to_oss(cover_url, args.cover, "封面图")

    # ===== Step 6: 上传参考 PDF 到 OSS =====
    client.upload_to_oss(reference_url, args.reference_pdf, "参考PDF")

    # ===== Step 7: 创建文章（等 2 秒，对应原脚本 setTimeout）=====
    log.info(f"⏳ 等待 {args.delay}s...")
    import time
    time.sleep(args.delay)

    result = client.create_article(
        title=title,
        article_name=f"{title}.pdf",
        article_key=article_key,
        cover_key=cover_key,
        reference_key=reference_key,
        reference_link=args.reference_link,
        tag_ids=tag_ids,
    )

    # 输出结果
    log.info("=" * 60)
    if result.get("message") == "success":
        log.info("✅ 文档入库完成!")
    else:
        log.error(f"❌ 入库失败: {result}")
    sys.exit(0 if result.get("message") == "success" else 1)


if __name__ == "__main__":
    main()
