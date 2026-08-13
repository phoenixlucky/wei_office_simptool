# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and the project follows Semantic Versioning in a pragmatic way:

- minor releases may include breaking changes before `1.0.0`
- patch releases are reserved for backwards-compatible fixes and documentation-only corrections

## [0.7.0] - 2026-08-13

### Added

- `MySQLDatabase` 支持 `with` 语句上下文管理，退出自动关闭连接
- 新增异常类型：`MySQLDatabaseError`（数据库操作失败统一抛异常）、`MailError`（邮件发送失败）
- `py.typed` 标记 + 公共 API 类型注解（text/utils/mail/database 域），恢复 IDE 补全与静态检查
- CLI 新增 `date`（日期计算）与 `excel info`（工作簿信息）子命令
- 新增 `examples/` 目录：`quickstart.py`、`excel_demo.py`、`chatbot_demo.py`
- 新增 `.github/workflows/release.yml`：打 `v*` tag 自动构建并发布到 PyPI（需配置 `PYPI_TOKEN` secret）
- 测试从 25 增至 38：数据库错误处理、Excel 真实读写、邮件错误路径、CLI 新子命令

### Changed

- 依赖按领域拆分：核心包仅保留 `toml`/`requests`；`pandas`/`openpyxl` 移入 `[excel]` extras，`mysql-connector-python` 移入 `[database]` extras
- 数据库/邮件/Excel/文本模块内部 `print` 改为 `logging`，避免污染调用方输出
- 邮件附件改用 `Path` 拼接 + `MIMEApplication`（UTF-8 文件名），并校验路径与长度
- CI 安装命令更新为 `pip install -e ".[analysis,excel,database]"`

### Removed

- `MySQLDatabase.run_ai_chatbot`：依赖不可用的 `mysql.ai.genai`（MySQL HeatWave 专有模块），已移除。数据库 AI 能力请改用 `wei_data_shu.ai.ChatBot`

### Migration Notes

- `pip install wei-data-shu` 后，使用 Excel/数据库功能需额外安装 `wei-data-shu[excel]` / `wei-data-shu[database]`
- 依赖 `MySQLDatabase` 旧行为（失败静默打印）的代码，请改为捕获 `MySQLDatabaseError`

## [0.6.1] - 2026-03-17

### Added

- added `.github/dependabot.yml` with weekly schedule for pip and GitHub Actions
- added minimum version constraints to all pip dependencies (`pandas>=2.0`, `openpyxl>=3.1`, etc.)

### Changed

- bumped package version to `0.6.1`

## [0.5.3] - 2026-03-17

### Added

- added a `LICENSE` file (GPL-3.0) to the project root, matching existing package metadata
- added GPL-3.0 classifier to `pyproject.toml` and updated all license references in README
- added table of contents, project badges, and cross-reference tables to README
- added `MultipleTrendPredictor` usage example and enriched all domain samples

### Changed

- switched project license from MIT to GPL-3.0 (SPDX: `GPL-3.0-only`)
- rewrote and standardized the entire README: enriched usage examples for all 10 domains, added output samples, improved formatting consistency
- updated project description to bilingual format (Chinese + English)
- bumped package version to `0.5.3`

## [0.5.2] - 2026-03-17

### Added

- added a practical "5-minute quick start" example to the README using `ExcelManager`, `DateFormat`, `search_colors()`, and `generate_password()`

### Changed

- improved README usage documentation structure for faster onboarding
- bumped package version to `0.5.2`

## [0.5.0] - 2026-03-17

### Added

- added a lightweight CLI entrypoint: `wei-data-shu` and `python -m wei_data_shu`
- added `wei_data_shu.utils.generate_password` for readable password generation with ambiguous characters removed
- added searchable color metadata with Chinese display names and `search_colors()` support for English, Chinese, and HEX queries
- added `wei_data_shu.utils` for shared helpers like `fn_timer`, `mav_colors`, and `generate_password`
- added architecture documentation under `docs/plans/`
- added domain-level tests for root package, AI, database, docs, Excel, files, mail, text, and utils

### Changed

- finalized the package architecture around domain packages: `ai`, `database`, `docs`, `excel`, `files`, `mail`, `text`, and `utils`
- restricted root package exports so `wei_data_shu` now exposes domain packages only
- updated README examples to use domain-package imports consistently
- split Excel functionality into dedicated modules under `wei_data_shu.excel`
- reorganized tests into domain-specific test files

### Removed

- removed legacy flat modules such as `SQLManager.py`, `excelManager.py`, `fileManager.py`, `mailManager.py`, `ollamaManager.py`, `stringManager.py`, `textManager.py`, `timingTool.py`, `baseColor.py`, and `chartsManager.py`
- removed support for the old root-level object import style in documentation and public architecture

### Migration Notes

- replace root-level object imports with domain-package imports
- example:

```python
from wei_data_shu.excel import ExcelManager
from wei_data_shu.database import MySQLDatabase
from wei_data_shu.text import DateFormat
```

## [0.4.0] - 2026-02-13

### Added

- introduced a reorganized package layout and modernized package metadata
- improved README coverage for Excel, text analysis, AI chat, and daily email reports

### Changed

- updated the project to publish as `wei-data-shu`
