"""
百济患教 — 文档上架自动化脚本
================================
将 Apifox 测试场景 "文档上架" 转换为 Python。

功能: 批量更新文章发布状态为"已上架"。

原 Apifox 脚本使用 $sequence(62, 1, 21) 生成 62→82 的 ID 序列，
此处用 for 循环替代。

用法:
  # 默认: 发布 ID 62~82 (共 21 篇)
  python baiji_document_publish.py

  # 自定义范围
  python baiji_document_publish.py --start 10 --end 30

  # 单篇文章
  python baiji_document_publish.py --id 62
"""

import os
import sys
import argparse
import logging

import requests

# ============================================================
# 配置
# ============================================================
CONFIG = {
    "base_url": os.getenv("BAIJI_BASE_URL", "https://dev-api-baiji-patient-edu.nullht.com/api"),
    "verify_ssl": False,
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("baiji_publish")


# ============================================================
# 主流程
# ============================================================
def publish_article(base_url: str, article_id: int, verify_ssl: bool) -> bool:
    """
    PUT /articles/{id}/publish-status
    更新单篇文章为已发布状态。
    返回是否成功。
    """
    url = f"{base_url}/articles/{article_id}/publish-status"

    try:
        resp = requests.put(
            url,
            headers={
                "User-Agent": "Baiji-Publish/1.0 (Python)",
                "Content-Type": "application/json",
            },
            verify=verify_ssl,
        )
        resp.raise_for_status()
        log.info(f"  ✅ ID={article_id} 上架成功")
        return True
    except requests.RequestException as e:
        log.error(f"  ❌ ID={article_id} 上架失败: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="百济患教 — 文档上架")
    parser.add_argument("--id", type=int, default=None, help="单篇文章 ID")
    parser.add_argument("--start", type=int, default=62, help="起始 ID（默认 62）")
    parser.add_argument("--end", type=int, default=82, help="结束 ID（默认 82）")
    parser.add_argument("--step", type=int, default=1, help="ID 步长（默认 1）")
    args = parser.parse_args()

    cfg = CONFIG
    base_url = cfg["base_url"].rstrip("/")

    # 确定要发布的 ID 列表
    if args.id is not None:
        article_ids = [args.id]
    else:
        # $sequence(62, 1, 21) → start=62, step=1, count=21 → IDs: 62~82
        article_ids = list(range(args.start, args.end + 1, args.step))

    log.info("=" * 60)
    log.info(f"百济患教 — 文档上架: {len(article_ids)} 篇文章")
    log.info(f"ID 范围: {article_ids[0]} → {article_ids[-1]}")
    log.info("=" * 60)

    success_count = 0
    fail_count = 0

    for article_id in article_ids:
        if publish_article(base_url, article_id, cfg["verify_ssl"]):
            success_count += 1
        else:
            fail_count += 1

    log.info("=" * 60)
    log.info(f"📊 结果: {success_count} 成功 / {fail_count} 失败 / {len(article_ids)} 总计")
    sys.exit(0 if fail_count == 0 else 1)


if __name__ == "__main__":
    main()
