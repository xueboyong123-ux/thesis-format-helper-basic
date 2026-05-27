# Thesis Formatter Pro

> 论文格式辅助排版与检查工具，面向毕业论文、毕业设计说明书、课程论文和常规 Word 文档批量排版。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Thesis Formatter Pro 基于 `cwyalpha/Word-Formatter-Pro` 二次开发。原项目采用 MIT License，本项目保留原作者版权声明和 [LICENSE](LICENSE)。本工具用于辅助统一 Word 文档格式、识别论文结构并生成格式检查报告，不承诺对所有学校模板做到一键完全合规。

## 项目定位

本项目是一个 Word 论文格式辅助排版与检查工具，主要帮助你在排版前后完成常见格式统一、论文结构保护、正文首行缩进处理和格式问题提示。不同学校、学院、专业的论文模板存在差异，排版后仍建议人工打开 Word 或 WPS 复核。

## 二次开发声明

- 原项目：`cwyalpha/Word-Formatter-Pro`
- 原项目协议：MIT License
- 本项目性质：在原项目基础上进行论文格式场景增强与发布整理
- 版权说明：保留原作者版权声明和 LICENSE，不将本项目描述为完全原创项目

## 新增功能

- 正文首行缩进
- 首行缩进 GUI 配置
- 启动时界面缩放
- 格式检查报告
- 论文模式
- 论文结构识别
- 目录、参考文献、图题、表题、公式、表格内文字保护
- 图表题注、编号、位置与正文引用检查

## 适用场景

- 毕业论文
- 毕业设计说明书
- 课程论文
- 常规 Word 文档批量排版

## 环境要求

- Python 3.x
- 依赖见 [requirements.txt](requirements.txt)
- Windows 下如需处理 `.doc`、`.wps`、修订或自动编号转换，建议安装 Microsoft Office 或 WPS Office
- Linux、Kylin、macOS 下如需转换 `.doc`、`.wps`，建议安装 LibreOffice 并确保 `soffice` 可用

## 安装依赖

```bash
python -m pip install -U pip
pip install -r requirements.txt
```

## 使用流程

1. 启动 GUI：

   ```bash
   python wfp.py
   ```

2. 加载 Word 文档。
3. 选择论文模式。
4. 设置正文首行缩进参数。
5. 开始排版。
6. 查看生成的格式检查报告。

## CLI 简要用法

```bash
# 单文件排版
python wfp_cli.py format -i input.docx

# 多文件或目录批量处理
python wfp_cli.py format -i input.docx -i ./documents -o ./formatted_output

# 查看当前配置
python wfp_cli.py show-config

# 运行内置测试
python wfp_cli.py test
```

## 示例配置

通用论文推荐配置见 [sample_configs/thesis_general.json](sample_configs/thesis_general.json)。可根据学校模板要求复制后调整，再通过 GUI 加载配置或通过 CLI 参数使用。

## 测试命令

```bash
python wfp.py --test
python wfp_cli.py test
python -m py_compile wfp_gui.py wfp_config.py wfp_core.py wfp_tests.py wfp.py wfp_cli.py
```

## 已知限制摘要

- 不保证所有学校一键完全合规。
- 不默认自动更新目录。
- 不默认处理复杂页码分节。
- 不自动重写参考文献内容。
- 不自动重写图表编号。
- 不自动同步正文交叉引用。
- 对浮动图片、文本框图片、组合图形识别有限。

完整限制说明见 [KNOWN_LIMITATIONS.md](KNOWN_LIMITATIONS.md)。

## 发布检查

发布前建议按 [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md) 逐项检查，确认测试、LICENSE、二次开发来源、缓存文件和输出文件清理情况。

## 更新记录

版本变化见 [CHANGELOG.md](CHANGELOG.md)。

## License

本项目基于 MIT License 授权。原项目版权声明保留在 [LICENSE](LICENSE) 中。
