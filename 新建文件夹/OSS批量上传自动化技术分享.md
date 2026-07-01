# 知识库 OSS 批量上传接口自动化实现

## 痛点背景

公司知识库管理后台的文件上传流程依赖阿里云 OSS 预签名上传机制，完整链路涉及 **10 个步骤**：

```
扫描目录 → 创建上传任务 → 获取预签名URL → PUT上传OSS → 注册文件 → 更新任务状态 → 推送状态 → 轮询转换 → 生成下载链接
```

此前该流程在 Apifox 中通过 ForEach 循环 + PUT 请求实现。但 **PUT 方法本身不支持批量请求体**——一次 PUT 只能携带一个文件流，Apifox 的 ForEach 本质上只是串行遍历，30 个文件就要串行走 30 遍，中途任何一步网络抖动都可能导致整个流程中断，只能从头再来。

### 核心矛盾

| 环节 | Apifox 能做到 | Apifox 做不到 |
|------|:-----------:|:----------:|
| 创建上传任务 | 一次性批量提交 | - |
| 获取预签名 URL | 单文件请求 | 批量获取 |
| **PUT 上传 OSS** | 手动单文件 | **并发上传、失败重试** |
| 注册文件 | 单文件注册 | 失败自动重试 |
| 状态推送/轮询 | 可但也容易断 | 智能等待 |

**瓶颈就在 PUT 这一步。** 但要理解为什么，需要拆开看 OSS 的预签名机制和 PUT 的协议限制。

---

## 为什么 PUT 只能 Binary、只能一次一个

这不是 Apifox 功能不足，而是 OSS 预签名机制 + HTTP 协议语义共同决定的。

### OSS 预签名 URL 是"一次一文件"的

生成预签名 URL 的接口请求体里虽然 `filenames` 是数组：

```json
{"filenames": ["奥扎雷纳.ppt", "保密行为准则.pdf"]}
```

看起来可以批量，但阿里云 OSS 返回的每个 URL 里**签名和文件名是绑定的**：

```
https://xxx.oss-cn-shanghai.aliyuncs.com/PRE_SIGN_UPLOAD/xxx_奥扎雷纳.ppt
    ?Expires=1778142281
    &OSSAccessKeyId=LTAI5tPLmicNev4jfEC6qKZp
    &Signature=ZOB0ZWRiIeaxOdF0wg0lxuRJAuQ%3D
```

签名 `Signature` 是对文件名、过期时间、AccessKey 的加密结果。你拿这个 URL 传另一个文件，OSS 验证签名不匹配直接拒绝。**后端接口虽然接受批量传文件名，实质也是循环生成 N 个不同的 URL，每个 URL 只能 PUT 一个文件。**

### PUT 和 POST 的协议语义区别

| | POST | PUT |
|---|---|---|
| HTTP 语义 | 创建资源 | **替换**指定 URL 的完整实体 |
| 请求体格式 | `multipart/form-data`，可含多个文件 | **单一二进制流（binary）** |
| 幂等性 | 不幂等 | 幂等，重复 PUT 同一内容结果不变 |
| 与 OSS 的匹配 | 不合适，签名无法覆盖多个文件 | 语义完美：URL = 资源位置，PUT = 写入 |

OSS 选 PUT 而非 POST 的原因：预签名 URL 本身就是唯一资源标识，PUT 上去就是替换该资源，语义完全匹配。如果用 POST 则需要 `multipart/form-data` + boundary 分隔符，签名机制会更复杂，文件校验也绕。

### PUT 请求的实际样子

```
PUT /PRE_SIGN_UPLOAD/xxx_奥扎雷纳.ppt?Expires=...&Signature=... HTTP/1.1
Host: enterprise-private-yzj-dev.oss-cn-shanghai.aliyuncs.com
Content-Type: application/octet-stream

[文件的全部二进制字节，没有 key=value，没有 boundary，就是赤裸的字节流]
```

这就是为什么 Apifox 里 PUT 请求的 body 只有 **binary** 选项——请求体本身就是文件，不存在任何结构化格式。

### 根本矛盾

```
后端设计：  每个文件 → 独立预签名 URL → 独立 PUT → OSS 一一对应
Apifox 能做的：ForEach 拿到 N 个 URL → 一个一个串行 PUT → 中间断一次全流程报废
Python 能做的：ThreadPool → 3 个 URL 同时 PUT → 失败自动重试 → 全跑完才退出
```

**不是不会用 Apifox，是 HTTP 协议不允许一个 PUT 传多个文件。**

---

## 解决方案

**用 Python 脚本替换 Apifox 的 ForEach 循环**，核心思路：保留原有接口调用逻辑不变，仅在 OSS 上传环节引入 `ThreadPoolExecutor` 并发执行，同时对全部 HTTP 请求加入重试机制。

### 脚本架构

```
main()
├── scan_files()              # 0. 扫描本地目录
├── create_upload_tasks()     # 1. 批量创建上传任务
├── [并发池]                  # 2-5. N个线程同时执行
│   └── upload_single_file()
│       ├── api_post("generate_pre_upload_signature")  # 获取预签名URL
│       ├── requests.put(upload_url, file_binary)       # PUT上传OSS
│       ├── api_post("kb/file/upload")                  # 注册文件
│       └── api_post("uploadTask/status")               # 更新任务状态
├── push_file_status()        # 6. 推送文件状态
├── wait_for_conversion()     # 7-9. 轮询等待转换完成
└── generate_download_url()   # 10. 生成预签名下载链接
```

---

## 关键函数说明

### 1. `api_post` — 统一请求封装（含重试）

```python
def api_post(path: str, body: object, timeout: int = 60) -> dict:
```

封装了请求发送、SSL 证书跳过、状态码校验、JSON 解析。核心特点是 **3 次自动重试**：捕获 `ConnectionError`、`Timeout`、`HTTPError` 三种异常，每次重试间隔递增（3s、6s）。公司内网环境下间歇性连接重置（ConnectionResetError 10054）是常态，没有重试的话 8 个文件能失败 3-4 个，加了之后基本全部通过。

---

### 2. `scan_files` — 目录扫描

```python
def scan_files(dir_path: str) -> list:
```

遍历指定文件夹，过滤 `.DS_Store` 等系统文件，返回绝对路径列表。按文件名排序，保证后续 fileId 生成顺序可预期。

---

### 3. `create_upload_tasks` — 批量创建上传任务

```python
def create_upload_tasks(file_paths: list) -> list:
```

根据文件列表生成 `uploadTaskList` 请求体，**一次性** POST 到 `/api/uploadTask/list`。每个文件生成唯一 `fileId`（格式 `file-{时间戳}-{序号}`），初始状态设为 `IDLE`，这一步和 Apifox 预请求脚本的逻辑完全一致。

---

### 4. `upload_single_file` — 单文件完整上传（并发核心）

```python
def upload_single_file(file_path: str, task: dict, index: int) -> dict:
```

这是整个脚本的核心，每个文件在一个独立线程中串行执行 4 个子步骤：

| 子步骤 | 接口 | 关键处理 |
|--------|------|---------|
| 获取预签名 URL | `POST /api/oss/generate_pre_upload_signature` | 传入文件名，从响应提取 `url` 和 `file_key` |
| **PUT 上传 OSS** | `PUT {预签名URL}` | 以 `application/octet-stream` 二进制流上传，3 次重试 |
| 注册文件 | `POST /api/kb/file/upload` | 关联 file_key 和任务 ID |
| 更新状态 | `POST /api/uploadTask/status` | 标记为 SUCCESS |

OSS PUT 步骤单独加重试的原因：预签名 URL 中的文件名经过 URL 编码，含中文、全角符号（如 `！`）时偶发连接重置，与 `api_post` 的 JSON 请求不同，需独立处理。

返回字典包含 `file_name`、`file_key`、`file_id`、`success`，供主线程统计。

---

### 5. `push_file_status` — 推送文件状态

```python
def push_file_status(file_ids: list) -> None:
```

调用 `/api/kb/status`，将全部成功文件的 `file_id` 批量推送给后端触发文档解析。此接口响应较慢（10-30s），因此超时设为 120s，同样 3 次重试，最终失败不影响上传结果。

---

### 6. `wait_for_conversion` — 轮询转换状态

```python
def wait_for_conversion(project_id: str, timeout: int = 120) -> str:
```

每 3 秒查询一次 `/api/kb/file/list`，检查首个文件的 `convert_status` 字段，变为 `COMPLETED` 时返回 `cover` 字段（用于后续下载链接生成）。超时上限 120 秒。

---

### 7. `generate_download_url` — 生成下载链接

```python
def generate_download_url(cover_key: str) -> str:
```

用 `cover_key` 调用 `/api/oss/generate_pre_download_signature`，返回预签名下载 URL。可用于验证上传文件可正常访问。

---

## 并发实现

```python
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
    for i, (fp, task) in enumerate(zip(file_paths, tasks), 1):
        pool.submit(upload_single_file, fp, task, i)
        time.sleep(0.5)  # 错峰提交，避免同时打爆服务器
```

三个设计细节：
- **并发数 3**：公司内网服务器承载能力有限，5 并发时连接重置率明显上升，3 并发是最优平衡点
- **错峰 0.5s**：线程提交间隔 500ms，避免 3 个请求同时到达同一端点
- **`as_completed` 收集**：谁先完成先收集谁的结果，不阻塞已完成的线程

---

## 效果对比

| 维度 | Apifox 手动 | Python 脚本 |
|------|:---------:|:--------:|
| 8 文件总耗时 | ~5 分钟 | ~40 秒 |
| 失败重试 | 中断后手动排查 | 自动重试，基本不掉 |
| 操作步骤 | 输入配置 + 等待 | 粘贴 3 个参数 + 回车 |
| 可复用性 | 每次重新配置 | 改配置区 3 行即可 |

---

## 演示关键点

分享时可以展示三样东西：

1. **脚本全貌** — 279 行 Python 代码截图（30s）
2. **实际运行** — 终端跑 `python oss_batch_upload.py`，8 个文件从扫描到全部完成的完整日志（2min）
3. **反面对比** — Apifox 手动操作的录屏快放（1min）

---

## 附：环境依赖

```bash
pip install requests
```

Python 版本 ≥ 3.7，`concurrent.futures` 为标准库无需额外安装。
