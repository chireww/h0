# 百济患教 — 接口自动化脚本

将 Apifox 测试场景转换为 Python 实现，覆盖百济患教系统的文档入库与上架流程。

## 项目概览

百济患教是一个患者教育材料管理系统，核心流程为：

```
文档入库                         文档上架
┌──────────────────────┐        ┌─────────────────┐
│ 1. 获取预签名上传 URL  │        │ PUT /articles/   │
│ 2. 上传文件到 OSS     │   →    │ {id}/publish-    │
│ 3. 创建文章记录       │        │ status           │
└──────────────────────┘        └─────────────────┘
  3 个文件（PDF+封面+参考）        批量更新发布状态
```

## 脚本说明

| 脚本 | 来源 | 功能 | API 调用数 |
|------|------|------|-----------|
| `baiji_document_import.py` | 文档入库.apifox-cli.json | 上传 3 个文件到 OSS 并创建文章 | 7 (3 GET + 3 PUT + 1 POST) |
| `baiji_document_publish.py` | 文档上架.apifox-cli.json | 批量更新文章发布状态 | N (可配置) |

### `baiji_document_import.py`

```
POST /file/presigned-url  ×3  获取 article/cover/reference 的上传 URL 和 key
PUT  {presigned_url}      ×3  上传文件到 OSS（二进制流）
POST /articles            ×1  创建文章记录（关联所有 key + 标签）
     ↳ 断言 response.message == "success"
     ↳ 创建前等 2s（原脚本 setTimeout）
```

**命令行参数：**

```bash
python baiji_document_import.py \
  --title "药品指南" \
  --article-pdf ./doc.pdf \
  --cover ./cover.png \
  --reference-pdf ./ref.pdf \
  --reference-link "https://mp.weixin.qq.com/..." \
  --tags '[1,2,3]'
```

### `baiji_document_publish.py`

原 Apifox 脚本用 `$sequence(62, 1, 21)` 生成 62→82 的 ID 序列，每次调用 `PUT /articles/{id}/publish-status`。Python 版用 `range()` 替代。

```bash
# 默认范围 62-82
python baiji_document_publish.py

# 自定义范围
python baiji_document_publish.py --start 10 --end 30

# 单篇
python baiji_document_publish.py --id 100
```

## 涉及的接口自动化知识点

### 一、OSS 预签名上传模式

```python
# 1. 从业务服务器获取预签名 URL
POST /file/presigned-url  → { uploadUrl, fileKey }

# 2. 直接 PUT 文件到 OSS（不经过业务服务器）
PUT {uploadUrl}  + 文件二进制流
```

**知识点：**
- **预签名 URL（Presigned URL）**：服务端生成临时授权链接，客户端直接上传到对象存储，减轻服务器压力
- **业务服务器只负责生成 URL**，文件流不经过业务服务器
- `fileKey` 是 OSS 中的唯一标识，后续创建文章时用于关联文件
- 上传 OSS 时 `Content-Type: application/octet-stream`，不指定具体 MIME 类型

### 二、多文件关联上传

同一个文章需要上传 3 个文件，通过 3 组 `{url, key}` 分别上传和关联：

```
article_key  → 文章 PDF 在 OSS 的 key
cover_key    → 封面图片在 OSS 的 key
reference_key → 参考 PDF 在 OSS 的 key
```

**知识点：**
- 3 个预签名 URL 并行获取（无依赖关系），但原脚本串行调用
- 所有 key 在最后一步的 `POST /articles` 中统一提交
- 上传失败时，已获取的 key 会成为孤儿文件（实际项目中需考虑清理）

### 三、Apifox `$sequence` → Python `range()`

```
Apifox: $sequence(62, 1, 21)  →  [62, 63, 64, ..., 82]
Python: range(62, 83, 1)      →  [62, 63, 64, ..., 82]
```

**调用链：** URL 路径 `articles/$sequence(62,1,21)/publish-status` 会被 Apifox 展开为 21 个独立请求。

**知识点：**
- `$sequence(start, step, count)` 是 Apifox 特有的序列生成器
- 转换为 Python 用 `range(start, start + count * step, step)`
- 批量操作应记录成功/失败计数，单个失败不中断

### 四、JavaScript → Python 转换要点

| Apifox / JS | Python |
|-------------|--------|
| `pm.variables.set("key", value)` | 变量赋值 / 返回元组 |
| `$.data.uploadUrl` (JSONPath) | `resp.json()["data"]["uploadUrl"]` |
| `setTimeout(()=>{}, 2000)` | `time.sleep(2)` |
| `pm.expect(value).to.eql(expected)` | `assert value == expected` |
| `$sequence(62, 1, 21)` | `range(62, 83)` |

### 五、OSS PUT 上传注意事项

```python
with open(file_path, "rb") as f:
    requests.put(upload_url, data=f, headers={"Content-Type": "application/octet-stream"})
```

- 用 `rb` 模式读取，确保二进制完整性
- `data=f` 直接传文件对象，`requests` 会流式读取
- 大文件应使用流式上传避免内存溢出
- `PUT` 方法而非 `POST`（OSS 预签名 URL 通常指定了 HTTP 方法）

### 六、命令行参数设计

```python
parser.add_argument("--title", required=True)
parser.add_argument("--article-pdf", required=True)
parser.add_argument("--tags", default="[]")
```

**知识点：**
- `required=True` 标记必填参数
- `--tags '[1,2,3]'` 接受 JSON 字符串 → `json.loads()` 解析
- `default` 提供合理的兜底值
- 比环境变量更适合"每次运行都不同"的参数

### 七、批量操作结果汇总模式

```python
success_count = 0
fail_count = 0
for article_id in article_ids:
    if publish_article(...):
        success_count += 1
    else:
        fail_count += 1
log.info(f"{success_count} 成功 / {fail_count} 失败")
```

**知识点：**
- 单个失败不抛异常，记录后继续
- 最终汇总报告成功/失败计数
- 退出码反映整体结果：`sys.exit(0 if fail_count == 0 else 1)`

## 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `BAIJI_BASE_URL` | API 基础地址 | `https://dev-api-baiji-patient-edu.nullht.com/api` |
| `BAIJI_OPENID` | 用户 openid | — |

## 依赖

```bash
pip install requests
```
