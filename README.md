# onenote_access

独立软件组件：让其他系统能够**读取本机 OneNote 内容**（只读）。

提供三种调用方式，共享同一核心：

| 方式 | 入口 | 适用场景 |
|------|------|----------|
| **REST API** | `python -m onenote_access serve` | 其他系统/服务经 HTTP 调用（推荐） |
| **Python 库** | `from onenote_access import OneNoteClient` | 嵌入你自己的 Python 程序 |
| **CLI** | `python -m onenote_access list\|export` | 人工操作 / 脚本 / 定时任务 |

## 运行要求

- **Windows** + **OneNote 桌面版**（ONENOTE.EXE，非 UWP 版），且已启动并登录
- Python 3.8+
- `pip install pywin32`（唯一第三方依赖，COM 调用必需）

> OneNote UWP 版 / 网页版不暴露 COM 接口，本组件不支持；此场景需走 Microsoft Graph API。

## 快速开始

```bash
pip install pywin32

# 1) 离线自测 (不连 OneNote, 验证解析/转换/HTTP 管线)
python -m onenote_access selftest

# 2) 列出 笔记本/分区/页面 层级
python -m onenote_access list

# 3) 启动 REST API (默认 http://127.0.0.1:8765)
python -m onenote_access serve
```

## REST API

所有响应为 UTF-8 JSON。除 `/health` 外均以 `/api/v1/` 开头。

| 端点 | 方法 | 说明 |
|------|------|------|
| `/health` | GET | 探活 + COM 可用性（`com_available`） |
| `/api/v1/notebooks` | GET | 完整层级树（笔记本/分区组/分区/页面，含页面 ID） |
| `/api/v1/pages` | GET | 页面元数据列表，query 过滤（见下） |
| `/api/v1/pages/{page_id}` | GET | 单页详情 + Markdown（`?images=false` 跳过图片） |
| `/api/v1/export` | POST | 批量导出为 Markdown 文件 |

**`/api/v1/pages` 过滤参数**（均支持 `*` `?` 通配符，不区分大小写）：

| 参数 | 默认 | 说明 |
|------|------|------|
| `notebook` | `*` | 笔记本名 |
| `section` | `*` | 分区名（自动展开分区组） |
| `page` | `*` | 页面标题 |
| `search` | — | 全文筛选：只返回内容包含该文本的页面 |
| `limit` | `20` | 返回条数上限 |

**`/api/v1/export` 请求体**：

```json
{
  "notebook": "工作笔记",
  "section": "会议*",
  "page": "*",
  "search": null,
  "images": true,
  "all": false,
  "out_dir": "backup-2026"
}
```

`out_dir` 被限制在服务端 `--export-root` 之内（防目录穿越）；匹配超过 20 页需
`"all": true`。返回 `{out_dir, exported, failed, index, files}`。

**示例**：

```bash
curl http://127.0.0.1:8765/health
curl "http://127.0.0.1:8765/api/v1/pages?notebook=工作笔记&section=会议*"
curl "http://127.0.0.1:8765/api/v1/pages/{BG-xxxxxxx}?images=false"
curl -X POST http://127.0.0.1:8765/api/v1/export \
     -H "Content-Type: application/json" \
     -d '{"notebook": "工作笔记", "search": "SM9", "all": true}'
```

**错误约定**：`400` 参数非法/目录越界；`401` 未授权；`404` 页面不可读（不存在/加密分区）；
`503` OneNote COM 不可用（附排查提示）；`500` 内部错误。

### 鉴权与安全

- 服务**默认只绑定 `127.0.0.1`**；绑定非回环地址而未配置令牌时会打印警告
- 对外暴露时务必配置 Bearer 令牌：`--token xxx` 或环境变量 `ONENOTE_ACCESS_TOKEN`，
  之后所有请求带 `Authorization: Bearer xxx`
- 全组件只读：仅调用 OneNote COM 的 `GetHierarchy` / `GetPageContent` /
  `GetHyperlinkToObject` 三个只读方法，不修改任何笔记内容
- 加密分区无法读取内容（OneNote 限制），需先在 OneNote 中手动解密

## 作为库调用

```python
from onenote_access import OneNoteClient

client = OneNoteClient()

tree = client.notebooks()                       # 完整层级
pages = client.find_pages(notebook='工作笔记',   # 通配符过滤
                          section='会议*')
page = client.get_page(pages[0]['page_id'])     # 单页 → dict
print(page['markdown'])                         # Markdown 正文
print(page['onenote_link'])                     # onenote:// 跳转链接

hits = client.search_pages('SM9')               # 全文筛选

result = client.export(out_dir=r'D:\backup',    # 批量导出
                       notebook='*', all_pages=True)
print(result['exported'], '页 ->', result['out_dir'])
```

COM 不可用时所有方法抛 `onenote_access.OneNoteComError`（`get_page` 对加密分区返回 `None`）。

## CLI

```bash
python -m onenote_access list
python -m onenote_access export --out-dir D:\backup --notebook "工作*" --all
python -m onenote_access export --out-dir .\out --search "整定" --no-images
python -m onenote_access serve --host 0.0.0.0 --port 8765 --token <secret> --export-root D:\exports
python -m onenote_access selftest        # 离线自测, 不连 OneNote
```

## 输出格式

导出的 Markdown 带 YAML front matter：

```markdown
---
title: 周会 2026-10-10
notebook: 工作笔记
section: 会议记录
created: 2026-10-10T09:00:00.000Z
modified: 2026-10-10T10:00:00.000Z
page_id: "{...}"
onenote_link: "onenote:#section-id=...&page-id=..."
source: OneNote desktop (COM API, read-only)
---

（正文：标题/列表/表格/代码块/图片，图片落同名 .assets/ 目录）
```

目录布局：`<out_dir>\<笔记本>\<分区组…>\<分区>\<页面>.md`，同名页面自动 `(2)` 去重，
根目录生成 `README.md` 索引。

## 项目结构

```
onenote_access/
├── onenote_access/
│   ├── com_backend.py      # OneNote COM 封装 (pywin32, 多调用样式兼容)
│   ├── hierarchy.py        # 层级 XML 解析 / 通配符过滤 / frontmatter
│   ├── exporter.py         # Markdown 导出管线 (含降级路径)
│   ├── onenote_html2md.py  # HTML → Markdown 转换器 (纯标准库)
│   ├── client.py           # OneNoteClient 高层 API
│   ├── server.py           # REST API 服务 (stdlib http.server)
│   └── __main__.py         # CLI
├── tests/test_offline.py   # 离线自测 (33 项, 不需要 OneNote)
├── requirements.txt
└── README.md
```

## 已知限制

- 仅 Windows + OneNote 桌面版；组件本身跨平台，但 COM 层在非 Windows 上会报错
- 加密分区读不了内容（需先解密）；`onenote_link` 获取失败时为 `null`（不影响其他字段）
- 全文筛选 `search` 是逐页拉取后字面匹配，页数很多时较慢（上限 500 页扫描）
- 在受限沙箱环境中 COM 可能被安全策略拦截（组件会把原因写进错误信息）
