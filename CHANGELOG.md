# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 格式，版本号使用 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Planned
- 更紧凑的结果详情布局，兼容低分辨率屏幕。
- 可选的多结果整行对比视图。
- 更完善的 OCR 后处理（常见字段自动纠错）。

## [1.1.0] - 2026-10-19

### Added
- 结果详情面板：直接显示整行全部非空字段，支持一键复制整行。
- 窗口置顶：默认始终显示在其他窗口上方，右上角可切换并记住状态。
- 自定义全局快捷键：支持现场录制组合键，内置三组推荐组合。
- 快捷键冲突检测：注册失败时保留原快捷键并给出明确提示。
- 新增 GitHub 项目文档、CI、测试、Issue/PR 模板。

### Changed
- 默认全局快捷键由 `Ctrl + Alt + Q` 改为 `Ctrl + Shift + Q`，降低与系统快捷键冲突的概率。
- 结果表默认自动选中第一条，并直接展示完整整行内容。
- 小窗口下自动收起整行详情，优先保留结果列表可用空间，可手动展开。
- 完善帮助说明与高 DPI 布局。

### Fixed
- 修复部分场景下结果行内容被截断、无法直接复制完整行的问题。
- 修复高 DPI 下窗口顶部控件、帮助页和快捷键设置页的布局问题。

## [1.0.0] - 2026-10-19

### Added
- 首次发布。
- 屏幕框选 + Windows 本地 OCR。
- 支持 `.xlsx`、`.xlsm`、`.csv` 全工作簿检索。
- 智能匹配、精确匹配、包含匹配。
- 双击结果在 Microsoft Excel 中定位单元格。
- 全局快捷键 `Ctrl + Alt + Q`。
- 本地离线运行，无网络请求。

[Unreleased]: https://github.com/yourname/shiwen-xunbiao/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/yourname/shiwen-xunbiao/releases/tag/v1.1.0
[1.0.0]: https://github.com/yourname/shiwen-xunbiao/releases/tag/v1.0.0
