# App Review Classifier Skill

将 Google Play 导出的评论 CSV 转换为结构化、带分类的 xlsx 文件。

## 功能

- 读取 Google Play 评论 CSV（UTF-16/UTF-8）
- 数据清洗 + 日期格式化
- 按星级拆分为两个 Sheet（5星 / 1-4星）
- 自动写入 GOOGLETRANSLATE 机翻公式
- 60+ 二级分类标注（广告、好评、崩溃、功能、音频处理等）
- 格式化 xlsx 输出（样式、冻结、筛选）

## 使用方式

### 作为 Claude Code Skill

将 `app-review-classifier.skill` 导入 Claude Code 即可使用。触发词：

- "处理评论" / "分类评论" / "评论转xlsx"
- "Google Play 评论"
- 上传文件名含 `reviews` 的 CSV

### 独立预处理脚本

```bash
pip install pandas openpyxl
python review-classifier-skill/review_csv_to_xlsx.py input.csv [output.xlsx]
```

脚本完成 CSV → xlsx 的数据处理（清洗、拆分、公式、样式），二级分类列留空等待 Claude 填充。

## 构建

推送 `v*` tag 时自动通过 GitHub Actions 构建 `.skill` 文件并创建 Release。

```bash
git tag v1.0.0
git push origin v1.0.0
```
