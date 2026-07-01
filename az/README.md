# AZ 审核系统 — 接口自动化脚本集

将 Apifox 测试场景转换为 Python 实现，涵盖 AZ 审核系统的完整自动化测试链路。

## 脚本总览

| 脚本 | 来源 | 一句话定位 |
|------|------|-----------|
| `az_audit_compare.py` | 统一对比版本 | 最简洁的"上传→审核→对比金标准"标准流程 |
| `az_audit_compare_with_download.py` | 添加下载文件夹 | 全自动：从飞书拉取金标准 Excel 再执行审核对比 |
| `az_audit_compare_layla.py` | layla 调试方案 | 两段式循环：先批量创建任务，再逐个轮询结果对比 |
| `az_batch_create_audit.py` | dev3-病例分享材料 | 纯批量创建：不做金标准对比，仅收集审核结果 |

---

## 各脚本独特之处

### 1. `az_audit_compare.py` — 标准流程

**独特设计：**
- **单 forEach 循环 + 统一对比**：所有文件上传并创建审核任务后，一次性获取全部审核结果，再与金标准集中对比
- **task_queue 数据结构**：`[{id, fileName}]`，携带文件名信息，便于对比时按文件名匹配金标准
- **轮询等待机制**：在所有任务创建完成后统一轮询，而非每个任务单独等待

**适用场景：** 金标准 Excel 已在本地，需要完整对比报告

### 2. `az_audit_compare_with_download.py` — 全自动飞书管线

**独特设计：**
- **Phase 1 飞书 8 步管线**：从飞书云文档自动下载金标准 Excel，无需人工准备
- **双客户端架构**：`FeishuClient`（Cookie 认证）+ `AZClient`（Bearer Token 认证），两种认证方式共存
- **飞书 API 链式调用**：
  ```
  创建文件夹下载 → 轮询完成 → 打包 zip → 轮询完成
  → 下载解压 → 导出 Excel → 轮询完成 → 下载解析
  ```
- **zip 解压 + Excel 提取**：内联实现了原 Apifox 调用的 `extract_zip_and_get_excel_url.py` 逻辑

**适用场景：** 金标准 Excel 存储在飞书云文档中，需要全自动化拉取

### 3. `az_audit_compare_layla.py` — 两段式循环

**独特设计：**
- **两个独立的 forEach 循环**：
  - 第一段：遍历文件 → 上传 → 创建任务 → 收集 `taskIds[]`
  - 第二段：遍历 `taskIds[]` → 逐个获取审核结果 → 实时对比
- **纯 ID 数组**：`taskIds` 只存 ID，不附带文件名，通过 `file_names` 字典做映射
- **审核参数差异**：`material_properties=[2]`, `target_audience=[26]`，与其他版本不同
- **原脚本对比逻辑未实现**：Apifox 版"获取审核结果"步骤为空壳，Python 版补齐了对比+轮询

**适用场景：** 需要分阶段观察中间状态，或任务创建与结果获取需要时间间隔

### 4. `az_batch_create_audit.py` — 批量创建不对比

**独特设计：**
- **Apifox `loop` 循环而非 `forEach`**：通过 `current_file_index` 索引递增实现顺序处理，天然支持断点续传
- **5 秒 delay**：创建审核任务与获取详情之间插入固定等待（原脚本 `delay` 组件）
- **不做金标准对比**：纯粹的批量任务创建 + 结果收集，输出为 `{id, task_id, file_name}` 列表
- **`drug_type` 动态注入**：`product_Ids` 通过环境变量配置，支持不同药物类型
- **上传成败判断**：检查 `response.code == 0` 而非仅看 HTTP 状态码

**适用场景：** 仅需批量创建审核任务，不需要验证审核结果正确性

---

## 涉及的接口自动化知识点

### 一、Apifox → Python 转换模式

| Apifox 概念 | Python 对应 |
|-------------|------------|
| `pm.environment.get/set` | 字典 / `os.getenv` / 实例属性 |
| `pm.variables.get/set` | 局部变量 / 实例属性 |
| `pm.executeAsync("script.py")` | 内联 Python 函数 |
| `pm.response.json()` | `resp.json()` |
| `forEach` 循环 | `for file in files:` |
| `loop` + `current_file_index` | `for index, file in enumerate(files):` |
| `if` 条件分支 | `if upload_success:` |
| `delay` 组件 | `time.sleep(5)` |
| `postProcessors.extractor` (JSONPath) | `resp.json().get("data")` 链式取值 |
| `preProcessors.customScript` | 请求前的预处理函数 |
| 环境变量 `{{var}}` | `os.getenv("VAR")` 或 `CONFIG["var"]` |

### 二、认证方式

| 类型 | 实现 | 出现脚本 |
|------|------|---------|
| **Bearer Token** | 登录后从 `resp.json().data.token` 提取，设置 `Authorization: Bearer xxx` 头 | 全部 4 个 |
| **Cookie 认证** | 飞书 API 需要 `Cookie` + `X-CSRF-TOKEN` + `Referer` 三个头 | `with_download` |

**关键点：** 两种认证方式共存的场景，需要分别管理 Session 和请求头，避免串扰。

### 三、轮询模式

项目中出现了三种轮询策略：

```python
# 1. 固定间隔 + 最大次数（飞书异步任务）
for attempt in range(max_attempts):
    status = api.check()
    if status == target:
        break
    time.sleep(interval)

# 2. 超时控制（长时间异步任务）
start = time.time()
while time.time() - start < timeout:
    ...

# 3. 创建后固定延迟（审核任务）
time.sleep(5)  # 然后直接获取结果
```

**关键点：** 飞书异步任务（文件夹下载、zip 打包、Excel 导出）用轮询；审核任务用固定 delay。

### 四、文件上传（multipart/form-data）

```python
with open(file_path, "rb") as f:
    resp = session.post(url, files={"file": (file_name, f)})
```

**关键点：**
- `requests` 的 `files` 参数自动设置 `Content-Type: multipart/form-data`
- 上传时不设置 `Content-Type: application/json`，否则会冲突
- 上传完成后恢复 JSON Content-Type 头

### 五、Excel 解析（openpyxl）

```python
wb = openpyxl.load_workbook(filepath, data_only=True)
ws = wb[sheet_name]
headers = [cell.value for cell in ws[1]]
rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    rows.append({headers[i]: v for i, v in enumerate(row) if ...})
# 按文件名分组 → merged_result
```

**关键点：**
- `data_only=True` 读取公式计算结果而非公式本身
- 按"文件名"字段将行数据分组为 `{"文件名": [审核点列表]}` 结构
- 通过 `字段名` 列实现审核点到 API 响应字段的灵活映射

### 六、金标准对比模式

```python
# 1. Excel 行 → 文件名分组
gold_standard = {"文件A.pdf": [{"审核点": "分类", "期望值": "NS", "字段名": "category"}, ...]}

# 2. 从 API 响应中提取实际值
actual = _extract_value(audit_data, gold_row)  # 支持 JSONPath + 递归搜索

# 3. 逐项对比
passed = str(actual) == str(expected)
```

**关键点：**
- `字段名` 列定义 API 响应 JSONPath，实现数据驱动的字段映射
- 回退策略：无字段名时递归搜索响应 JSON 中匹配的 key
- 对比结果结构化输出为 JSON 报告

### 七、飞书/Lark Open API 调用链

```
POST /space/api/box/invoke/create/     → invoke_code (创建异步下载任务)
GET  /space/api/box/invoke/check/      → invoke_status (轮询任务状态)
POST /space/api/box/zip/create/        → zip_code     (打包为 zip)
GET  /space/api/box/zip/check/         → zip_status   (轮询打包状态)
GET  (下载 zip)                        → 解压提取 Excel token
POST /space/api/export/create/         → ticket       (创建导出任务)
GET  /space/api/export/result/{ticket} → job_error_msg (轮询导出状态)
GET  (下载 Excel)                      → openpyxl 解析
```

**关键点：**
- 每个异步操作都是 "创建 → 轮询" 对
- 状态值的语义不同：`invoke_status==1` 表示完成，`zip_status==0` 表示完成，`job_error_msg=="success"` 表示完成
- 飞书 API 依赖 Cookie 认证而非 Bearer Token

### 八、循环控制策略对比

| 策略 | 脚本 | 特点 |
|------|------|------|
| `for` 循环（forEach） | `compare`, `with_download` | 一次性遍历，适合已知文件列表的批量操作 |
| 两段 `for` 循环 | `layla` | 阶段分离：创建 → 等待 → 获取，中间可插入逻辑 |
| 索引递增（loop） | `batch_create` | 顺序处理，支持断点续传，`current_file_index` 持久化 |

### 九、错误处理与容错

```python
# 1. 上传失败 → 跳过当前文件继续
if not file_id:
    log.warning(f"跳过 {file_name}")
    continue

# 2. API 调用失败 → 记录错误不中断流程
try:
    detail = client.get_audit_detail(task_id)
except Exception:
    results.append({"status": "error", ...})

# 3. 金标准缺失 → 标记而非崩溃
if not gold_rows:
    results.append({"status": "no_gold_standard", ...})
```

**关键点：** 批量任务中的单个失败不应中断整体流程，错误应被记录并在报告中体现。

### 十、配置管理模式

```python
CONFIG = {
    "az_base_url": os.getenv("AZ_BASE_URL", "https://dev-api-..."),
    "az_email": os.getenv("AZ_EMAIL", ""),
    ...
}
```

**关键点：**
- 环境变量优先 → 默认值兜底
- 敏感信息（密码、token）不进代码，通过环境变量注入
- `audit_defaults` 子字典管理审核请求的固定参数

---

## 使用方式

```bash
# 安装依赖
pip install requests openpyxl

# 设置环境变量
export AZ_EMAIL=your_email
export AZ_PASSWORD=your_password
export AZ_FILE_DIR=D:\test_files\

# 运行（按需选择）
python az_audit_compare.py                      # 标准对比
python az_audit_compare_with_download.py        # 飞书全自动
python az_audit_compare_layla.py                # 两段式
python az_batch_create_audit.py                 # 批量创建
```
