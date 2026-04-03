## What's New in v1.2.0

### New Features

- **多 App 支持**：二级分类列表从 SKILL.md 中抽离，每个 App 独立存放在 `references/` 目录下
- **App 确认阶段**：执行前先列出已支持的 App，询问用户是哪个 App 的评论
- **新 App 动态扩展**：如果用户的 App 不在列表中，可提供二级分类列表，自动保存为新的 reference 文件

### Changes

- 流程从三阶段改为四阶段：确认 App → 脚本处理 → 询问意图 → 分类
- 首个已支持 App：MP3 Cutter（`references/mp3-cutter.md`）
