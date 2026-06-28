# 报告生成系统后端

这是按《前后端接口统一规范文档-二次补充终版》实现的一版可运行后端，使用 Python 标准库实现 HTTP 服务，不依赖 FastAPI/Flask。

## 启动

```powershell
cd C:\Users\hp\Documents\Codex\2026-06-27\ho\report-backend
C:\Users\hp\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -X utf8 app.py
```

默认地址：

```text
http://127.0.0.1:8000
```

接口基础路径：

```text
/api
```

## 默认账号

```text
普通用户：student / 123456
管理员：admin / admin123
超级管理员：super / super123
```

## 已覆盖接口

- 认证：注册、登录、当前用户、退出
- 报告：创建、列表、详情、更新、删除
- 大纲：生成、保存、新增章节、删除章节、重新编号
- 内容：SSE 流式生成、进度、章节状态、表格事件、保存编辑内容、单章节重新生成
- 导出：docx、pdf、md、txt，导出状态、导出历史、下载
- 模板：列表、详情、上传、更新结构、停用删除
- 素材：列表、上传、删除、启用/停用
- 模型配置：查询、保存、测试连接
- 用户管理：超级管理员查询用户、修改角色、启用/停用

## 数据与文件

数据保存在：

```text
report-backend/data/db.json
```

上传和导出文件保存在：

```text
report-backend/storage
```

删除 `data/db.json` 后重启服务，会重新生成默认数据。

