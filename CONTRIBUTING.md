# 贡献指南

感谢你愿意改进 **识文寻表**。这是一个 Windows 桌面工具，优先保证“框选 → OCR → Excel 查找”这条主链路的稳定和易用。

## 提交 Issue

- 使用仓库中的 **Bug 反馈** 或 **功能建议** 模板。
- Bug 请说明 Windows 版本、Python 版本、Excel 文件格式，以及复现步骤。
- 涉及 OCR 时，请附上**脱敏后的截图**或识别结果；不要上传真实姓名、手机号、身份证号等个人信息。
- 涉及 Excel 时，请尽量提供最小化的示例文件，并先删除敏感数据。

## 开发环境

- Windows 10 / 11
- Python **3.10+**（64 位，带 Tkinter）
- Pillow

```powershell
python -m pip install -r requirements.txt
python app.py
```

运行测试：

```powershell
python -m unittest discover -s tests -v
```

编译检查：

```powershell
python -m py_compile app.py excel_index.py ocr_service.py win_support.py
```

## 项目结构

```text
app.py            主界面、状态管理、快捷键设置
excel_index.py    XLSX / CSV 解析、索引、搜索
ocr_service.py    Windows OCR 封装与图像预处理
win_support.py    全局快捷键、截图、框选浮层
assets/           图标、界面截图、PowerShell 辅助脚本
tests/            不依赖 GUI 的核心单元测试
tools/launcher/   轻量 EXE 启动器源码
```

## 代码约定

- 使用 Python 3.10+ 语法，保持类型标注。
- 面向用户的错误信息使用中文，并尽量说明**下一步怎么做**。
- 不要引入网络请求、遥测或上传用户数据。
- 新增 OCR / Excel 行为时，优先补充 `tests/` 下的纯逻辑测试。
- UI 改动请同时检查高 DPI（125% / 150%）和较小的窗口尺寸。
- 保持现有文件结构；不要提交 `__pycache__`、`.cache/`、`.config/`、日志或本地配置。

## Pull Request 清单

- [ ] 改动聚焦，一个 PR 解决一个问题。
- [ ] 已运行单元测试和 `py_compile`。
- [ ] 涉及 UI 的改动附上截图或说明。
- [ ] 没有提交敏感数据、临时文件或构建产物。
- [ ] 如新增用户可见功能，已更新 `README.md` 和 `CHANGELOG.md`。

## 提交信息

建议使用简短、明确的提交信息，例如：

```text
feat: show full row in result detail
fix: keep table usable on 700px height
docs: add GitHub issue templates
```

## 许可

提交代码即表示你同意以本项目的 [MIT License](LICENSE) 授权你的贡献。
