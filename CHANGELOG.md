## What's New in v1.0.0

### Features

- **CSV 读取**：自动检测 UTF-16/UTF-8 编码，提取 Google Play 评论的 5 个关键列
- **数据清洗**：过滤空评论，日期统一为 `YYYY-MM-DD` 格式
- **分表输出**：5星评论与 1-4星评论自动拆分为独立 Sheet
- **机翻公式**：B 列自动写入 `GOOGLETRANSLATE` 公式，上传 Google Sheets 即可生效
- **二级分类**：覆盖广告、好评、崩溃、功能、音频裁剪/合并/混合、下载、铃声、文件、音质等 13 大类 60+ 子类
- **xlsx 样式**：标题行蓝底白字、列宽预设、冻结首行、自动筛选
- **预处理脚本**：附带 `review_csv_to_xlsx.py`，可独立运行 Step 1-4 + 6 的数据处理，减少 Claude token 消耗
