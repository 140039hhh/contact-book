# 联络 · 通讯录

基于 Python 和 SQLite 的网页通讯录，支持联系人管理、收藏、多种联系方式以及 Excel 导入导出。

## 功能

- 新建、编辑、删除联系人，搜索姓名、联系方式和备注。
- 收藏常用联系人，在“我的收藏”中单独查看。
- 每位联系人可以添加多条电话、邮箱、社交账号和地址。
- 社交账号支持 QQ、微信，也可以手动填写其他软件名称。
- 导入和导出 Excel，数据保存在 SQLite 中，重启后仍然保留。

## 技术栈与目录

前端使用 HTML、CSS 和 JavaScript，后端使用 Python，数据库使用 SQLite，Excel 文件由 openpyxl 处理。

| 文件 | 说明 |
| --- | --- |
| server.py | 后端接口、数据校验、数据库及Excel处理 |
| static/index.html | 页面结构 |
| static/style.css | 页面样式 |
| static/app.js | 页面交互 |
| test_app.py | HTTP接口测试 |
| test_core.py | 核心逻辑测试 |
| requirements.txt | Python依赖 |
| Dockerfile、compose.yaml | 容器部署配置 |
| CODE_STYLE.md | 代码规范 |

## 本地运行

需要 Python 3.10 或以上。在项目目录打开 PowerShell，执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe server.py
```

浏览器打开 http://localhost:8000。保持终端运行，按 Ctrl+C 停止服务。

程序自动创建 data/contacts.db。更新程序时保留 data 文件夹；备份时先停止服务，再复制该文件夹。

如需设置访问密码，启动前执行：

```powershell
$env:APP_PASSWORD="设置自己的密码"
.\.venv\Scripts\python.exe server.py
```

在网页点击“访问密码”并输入相同密码。未设置环境变量时，本地访问不需要密码。

## Excel 格式

每位联系人占一行，前三列固定为 name、favorite、notes。

| 列名 | 内容 | 示例 |
| --- | --- | --- |
| name | 姓名，必填 | 测试联系人 |
| favorite | 1表示收藏，0表示不收藏 | 1 |
| notes | 备注 | 同学 |
| phone_1 | 电话 | 00000000001 |
| email_1 | 邮箱 | demo@example.com |
| social_1 | 软件名称和账号 | QQ：demo_qq |
| social_2 | 第二条社交账号 | 微信：demo_wechat |
| address_1 | 地址 | 示例地址 |

同类联系方式可以继续增加编号列，如 phone_2。社交软件名称与账号之间使用中文冒号“：”，保存在同一个单元格。旧版 social_1 与 socialplatform_1 分列格式仍可导入。

建议先导出一份空白 Excel 作为模板。电话号码和账号使用文本格式，避免前导零丢失。

导出包含全部联系人，不受搜索和收藏筛选影响。导入只新增，不覆盖、不去重，整份数据校验通过后才写入。单次导入最多10000位联系人，请求大小不超过10MB。

## 测试

```powershell
.\.venv\Scripts\python.exe -m unittest -v
```

测试使用临时数据库，不修改正式联系人。主要检查联系人增删改查、访问密码、Excel往返、社交软件名称保留和无效文件处理。

只运行不需要网络连接的核心测试：

```powershell
.\.venv\Scripts\python.exe -m unittest test_core -v
```

## Docker 部署

项目包含部署配置；公网运行需要另外准备服务器和域名。

在安装了 Docker 和 Compose 的 Linux 服务器上进入项目目录：

```sh
cp .env.example .env
# 编辑.env并修改APP_PASSWORD
 docker compose up -d --build
 docker compose logs --tail=50
```

容器映射到服务器本机127.0.0.1:8000，数据库使用持久数据卷。可通过 Caddy 配置HTTPS反向代理：

```caddy
contacts.example.com {
    reverse_proxy 127.0.0.1:8000
}
```

替换为实际域名，将域名解析到服务器，并允许80和443端口访问。完成配置后，从其他设备检查公网访问、联系人操作和数据保存。

更新代码后重新执行 docker compose up -d --build。停止使用 docker compose down，不要添加 -v，以免删除数据库数据卷。

## 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | /api/contacts | 查询联系人 |
| POST | /api/contacts | 新增联系人 |
| PUT | /api/contacts/ID | 更新联系人 |
| DELETE | /api/contacts/ID | 删除联系人 |
| GET | /api/export | 导出Excel |
| POST | /api/import | 导入Excel原始二进制文件 |

新增和更新使用JSON。设置访问密码后，请求头需携带 Authorization: Bearer 密码。

## 使用范围

当前为共享通讯录，使用共享访问密码，适合个人或小组演示。尚未实现独立用户账号和数据隔离。

## 在线访问

https://你的实际地址.onrender.com

部署于 Render 免费服务。首次访问可能需要等待唤醒；
SQLite 数据可能在服务休眠、重启或重新部署后重置。
