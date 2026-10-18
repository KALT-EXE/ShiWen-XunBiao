# 发布到 GitHub

这份清单把项目上传、设置仓库、生成 Release 的步骤放在一起。按顺序执行即可。

---

## 1. 上传前检查

不要上传以下内容：

- `.cache/`（OCR 临时文件）
- `.config/`（本机配置和日志）
- `__pycache__/`、`*.pyc`
- `*.log`
- `*.zip`（发布压缩包应放到 GitHub Releases，而不是源码仓库）
- 任何真实姓名、手机号、身份证号、Excel 业务数据

项目开头的 `.gitignore` 已经覆盖这些内容。发布压缩包 `识文寻表-v1.1.0.zip` 位于项目文件夹的上一级，不要放进仓库。

---

## 2. 创建仓库

建议仓库名：

```text
shiwen-xunbiao
```

备选：

```text
ocr-excel-finder
```

在 GitHub 新建仓库时：

- **不要**勾选 “Add a README file”，项目里已经有 `README.md`
- **不要**添加 `.gitignore` 或 License 模板，项目里已经准备好
- 仓库可见性按你的需要选择 Public / Private

---

## 3. 上传项目

### 方式 A：GitHub Desktop（推荐）

本项目已经初始化本地 Git 仓库，并完成了初始提交，可以直接发布。

1. 打开 **GitHub Desktop**，登录 GitHub 账号。
2. 菜单 **File → Add local repository…**。
3. 选择项目文件夹：

   ```text
   D:\works\work dsh\识文寻表
   ```

4. 确认左侧 **Current Repository** 显示 `识文寻表`，分支为 `main`。
5. 点击窗口顶部的 **Publish repository**。
6. 填写：
   - **Name**：`shiwen-xunbiao`
   - **Description**：复制下面的仓库简介
   - 如需公开项目，取消勾选 **Keep this code private**
7. 点击 **Publish repository**。

发布完成后点击 **View on GitHub** 打开仓库。

> GitHub Desktop 只能发布代码，不能创建 Release 或上传 Release 附件。Release 需要继续在网页上操作（见第 5 节）。

### 方式 B：命令行

如果仓库还没有初始化：

```powershell
git init
git branch -M main
git add .
git commit -m "feat: initial release v1.1.0"
git remote add origin https://github.com/<your-name>/shiwen-xunbiao.git
git push -u origin main
```

如果已经完成初始提交，只需要：

```powershell
git remote add origin https://github.com/<your-name>/shiwen-xunbiao.git
git push -u origin main
```

如果还没安装 Git，可从 <https://git-scm.com/download/win> 下载。GitHub Desktop 也自带 Git。

### 方式 C：网页上传

1. 在 GitHub 仓库页面点击 **Add file → Upload files**。
2. 把 `识文寻表` 文件夹内的文件拖进去（保留目录结构）。
3. 确认没有 `.cache`、`.config`、`__pycache__`、`.zip`。
4. Commit changes。

> 网页上传对隐藏文件不友好，建议优先使用 GitHub Desktop。

---

## 4. 仓库 About 信息

在仓库首页右上角 **About** 中填写。

### 一句话简介（推荐）

```text
框选屏幕文字，自动在 Excel 全表中查找并定位。Windows 本地 OCR、离线运行、整行输出、窗口置顶、自定义快捷键。
```

### 英文简介

```text
Select text on screen and instantly find it across Excel workbooks with local Windows OCR. Offline, full-row output, always-on-top.
```

### Topics

逐个添加以下 topics（最多 20 个）：

```text
ocr
excel
windows
python
tkinter
screen-capture
windows-ocr
text-recognition
productivity
desktop-app
offline
xlsx
csv
chinese
```

### 其他设置

- **Website**：暂时留空，或填写你自己的介绍页。
- **Include in the home page**：勾选 Releases、Packages 等按需。
- **Social preview**：上传 `assets/social-preview.png`（1280×640）。

---

## 5. 创建 Release

1. 打开仓库的 **Releases → Draft a new release**。
2. Tag 填 `v1.1.0`，目标分支选 `main`。
3. Release title：

   ```text
   v1.1.0 — 整行输出、窗口置顶、自定义快捷键
   ```

4. 描述直接复制 [docs/RELEASE_NOTES_v1.1.0.md](RELEASE_NOTES_v1.1.0.md) 的内容。
5. 上传附件：`识文寻表-v1.1.0.zip`（位于项目文件夹的上一级）。
6. 点击 **Publish release**。

发布后，README 中的 “Releases 页面下载最新压缩包” 即可正常使用。

---

## 6. 替换占位符

搜索并替换：

| 文件 | 占位符 | 替换为 |
|---|---|---|
| `README.md` | `<repository-url>` | 你的仓库 clone 地址 |
| `README_EN.md` | `<repository-url>` | 你的仓库 clone 地址 |
| `CHANGELOG.md` | `yourname` | 你的 GitHub 用户名 |

```powershell
git grep -n "yourname"
git grep -n "repository-url"
```

如果使用 GitHub 网页编辑，直接在文件中查找替换即可。

---

## 7. 建议的仓库设置

- **Issues**：开启（模板已在 `.github/ISSUE_TEMPLATE/`）。
- **Discussions**：可选，适合收集使用场景。
- **Actions**：保持开启，`CI` 工作流会自动运行核心测试。
- **Wiki**：可选，不需要可以关闭。
- **Releases**：至少发布 `v1.1.0`，附上 zip。
- **Social preview**：建议上传，链接分享时更专业。

---

## 8. 发布后自检

- [ ] README 中的图标、截图、社交预览图能正常显示。
- [ ] `assets/screenshot.png` 与当前版本界面一致。
- [ ] Release 附件能下载并解压运行。
- [ ] CI 全部通过。
- [ ] 没有提交 `.cache`、`.config`、日志或真实业务数据。
- [ ] `LICENSE` 的版权人已改成你的名字或组织。
- [ ] `CHANGELOG.md` 的 `yourname` 已替换。

---

## 9. 后续版本

发布新版本时：

1. 更新 `APP_VERSION`（`app.py`）。
2. 更新 `CHANGELOG.md`。
3. 更新 `README.md` 中的版本与截图（如有变化）。
4. 重新打包 `识文寻表-vX.Y.Z.zip`。
5. 新建同版本 tag 和 Release，附上新的 zip。

```powershell
git add .
git commit -m "feat: v1.1.1"
git tag v1.1.1
git push origin main
git push origin v1.1.1
```
