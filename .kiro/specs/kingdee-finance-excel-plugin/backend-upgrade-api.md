---
inclusion: manual
---

# 后台升级服务器 API 参考

本文档描述了部署在 `F:\CodeCompany\UpdateServer` 的 Django 升级服务器的 API 接口规范。
客户端开发升级功能时应参考此文档。

## 服务概述

- **技术栈**: Django 4.2 + Django REST Framework + MySQL + Uvicorn (ASGI)
- **默认端口**: 8100
- **认证方式**: API Key（通过 `X-API-Key` 请求头），管理接口需要认证，客户端查询/下载接口无需认证
- **版本号格式**: 语义化版本号 MAJOR.MINOR.PATCH（如 "1.2.3"）

## 客户端无需认证的公开接口

### 1. 获取产品版本清单

```
GET /api/v1/products/{product_id}/manifest
```

**响应 200**:
```json
{
  "product_id": "excel-common-tools",
  "version": "1.2.3",
  "force_update": false,
  "files": {
    "path/to/file.dll": "sha256-hash-64-chars-lowercase-hex"
  },
  "full_package_url": "/api/v1/products/excel-common-tools/packages/1.2.3"
}
```

**响应 404**: 产品不存在或无已发布版本
```json
{"error": "该产品无已发布版本可用"}
```

### 2. 版本检查（简化接口）

```
GET /api/v1/products/{product_id}/check
```

### 3. 下载增量文件

```
GET /api/v1/products/{product_id}/files/{version}/{file_path}
```

- 返回二进制流
- 支持 Range 请求（断点续传）
- 响应头包含 Content-Length

### 4. 下载全量包

```
GET /api/v1/products/{product_id}/packages/{version}
```

- 返回 zip 压缩包二进制流
- 支持 Range 请求

### 5. 分类产品列表

```
GET /api/v1/categories/{category_id}/products
```

**响应 200**:
```json
[
  {
    "product_id": "excel-common-tools",
    "display_name": "Excel通用工具",
    "published_version": "1.2.3"
  }
]
```

### 6. 分类聚合清单

```
GET /api/v1/categories/{category_id}/manifest
```

**响应 200**:
```json
[
  {
    "product_id": "excel-common-tools",
    "manifest": {
      "product_id": "excel-common-tools",
      "version": "1.2.3",
      "force_update": false,
      "files": {...},
      "full_package_url": "..."
    }
  }
]
```

> 公共分类（is_public=true）下的产品会自动合并到任意分类请求的结果中。

### 7. 健康检查

```
GET /health
```

**响应 200**: `{"status": "ok"}`

## 客户端升级流程（推荐）

1. **检查更新**: 请求 `GET /api/v1/products/{product_id}/manifest`
2. **比较版本**: 将 manifest 中的 `version` 与本地版本号比较
3. **判断强制**: 若 `force_update` 为 true，强制升级；否则提示用户
4. **增量更新**: 对比 `files` 中的哈希值与本地文件哈希，仅下载变更文件
5. **全量更新**: 若增量更新不可行，下载 `full_package_url` 指向的全量包
6. **校验完整性**: 下载完成后用 SHA256 校验文件哈希

## 后台源码参考位置

如需查看后台具体实现，关键文件在 `F:\CodeCompany\UpdateServer`：

| 模块 | 路径 |
|------|------|
| 需求文档 | `Readme/requirements.md` |
| 设计文档 | `.kiro/specs/design.md` |
| 产品管理 | `products/views.py`, `products/services.py`, `products/models.py` |
| 版本管理 | `versions/views.py`, `versions/services.py`, `versions/models.py` |
| 文件存储 | `files/views.py`, `files/services.py` |
| 分类管理 | `categories/views.py`, `categories/services.py`, `categories/models.py` |
| URL路由 | `upgrade_server/urls.py` |
| 认证中间件 | `middleware/auth.py` |
| 配置文件 | `config.yaml` |

## 错误响应格式

所有错误响应统一为 JSON 格式:
```json
{"error": "错误描述信息"}
```

常见状态码:
- 400: 请求参数格式错误
- 401: 认证失败（缺少密钥 / 密钥无效）
- 404: 资源不存在
- 416: Range 请求超出范围
- 429: 请求过于频繁（IP 被限流）
- 500: 服务器内部错误
