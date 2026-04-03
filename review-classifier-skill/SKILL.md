---
name: app-review-classifier
description: 处理 Google Play 用户评论 CSV 文件，完成清洗、格式转换、分类，并输出带机翻公式的 xlsx 文件。当用户上传评论 CSV 文件，或提到"处理评论"、"分类评论"、"评论转xlsx"、"Google Play 评论"时，必须使用此 skill。即使用户只说"帮我处理这个CSV"，只要文件名含 reviews 字样也应触发。
---

# App 评论处理 & 分类 Skill

## 概述

将 Google Play 导出的评论 CSV 转换为结构化、带分类的 xlsx 文件，支持5星与1-4星分表、二级分类标注、机翻公式。

---

## Step 1：读取 CSV

Google Play 导出的 CSV 通常为 **UTF-16 编码**，必须用以下方式读取：

```python
import pandas as pd
df = pd.read_csv('file.csv', encoding='utf-16')
```

**目标列**（固定提取这5列）：

| 原始列名 | 输出列名 |
|---|---|
| Review Text | 原文 (Review Text) |
| Star Rating | 评分 (Star Rating) |
| Reviewer Language | 语言代码 (Language) |
| Review Submit Date and Time | 发布时间 (Submit Date) |
| Review Link | Review Link |

**日期格式**：统一转为 `YYYY-MM-DD`：
```python
df['Review Submit Date and Time'] = pd.to_datetime(df['Review Submit Date and Time']).dt.strftime('%Y-%m-%d')
```

---

## Step 2：数据清洗

```python
# 过滤原文为空的行
df = df[df['Review Text'].notna() & (df['Review Text'].str.strip() != '')]
```

---

## Step 3：拆分为两个 Sheet

- **5星评论** Sheet：`Star Rating == 5`
- **1-4星评论** Sheet：`Star Rating < 5`

---

## Step 4：列顺序与机翻公式

最终列顺序（从左到右）：

```
A: 二级分类
B: 机翻 (Translation)
C: 原文 (Review Text)
D: 评分 (Star Rating)
E: 语言代码 (Language)
F: 发布时间 (Submit Date)
G: Review Link
```

**机翻公式**（B列，上传到 Google Sheets 后自动生效）：
```
=IF(C2="","",GOOGLETRANSLATE(C2,E2,"zh-CN"))
```

---

## Step 5：二级分类

### 分类规则

- **看内容，不看星级**：内容表达正面意思就归好评类，不管星级是几星
- **不确定时一律归**：`12-10 表述不清暂无法归类`
- 5星表格的纯好评 → `2-01 5星纯好评`
- 1-4星表格的纯好评 → `2-02 4星及以下纯好评`（包括 good/nice/best 等简短好评，以及其他语言中意思明确的好评词）
- `2-03 4星无意义评论` 仅用于真正意思不明的词（如 "bakbas"、"apk silid"）

### 完整二级分类列表

```
1-01 抱怨有广告
1-02 提及广告(4星及以下)
1-03 提及广告但给5星的
1-04 看广告时间长
1-05 广告很难关闭
1-06 色情广告
1-07 看广告解锁失败
1-08 广告_其他
1-09 希望减少广告
2-01 5星纯好评
2-02 4星及以下纯好评
2-03 4星无意义评论
3-01 崩溃/闪退
3-02 APP停止响应
3-03 程序打不开
3-04 程序卡顿
4-01 期望新功能
4-02 原有功能优化/改进
4-03 功能有限
4-04 功能模块分散/没有联动
5-01 音频裁剪保存失败
5-02 音频裁剪保存后音质差
5-03 音频裁剪不准确
5-04 音频裁剪_其他
5-05 音频裁剪_淡入淡出问题
5-06 音频裁剪_音量问题
5-07 音频裁剪_Mark标记功能
5-08 音频裁剪_无法剪辑
5-09 音频裁剪_无法连续裁剪
5-10 音频裁剪_裁剪杆问题
6-01 音频合并保存失败
6-02 音频合并_其他
6-03 音频合并_合并后播放卡顿
7-01 音频混合保存失败
7-02 音频混合_其他
7-03 音频混合_音质差
8-01 TikTok audio 下载失败
8-02 Instagram audio 下载失败
8-03 TikTok audio下载_其他
8-04 Instagram audio 下载_其他
9-01 铃声/闹铃/通知铃设置失败
9-02 铃声/闹铃/通知铃重置失败
9-03 在线铃声下载失败
9-04 铃声/闹铃/通知铃_其他
10-01 文件无法打开
10-02 文件打开失败
10-03 找不到需要的文件
10-04 外部文件打开失败
10-05 文件格式不支持
10-06 不能从SD卡中选择文件
10-07 选择列表不显示新文件
10-08 文件打开显示时长不一致问题
10-09 无法显示全部本地音乐
11-01 音质差/音质降低
11-02 保存速度慢
11-03 保存失败/转换失败/不能保存
11-04 音量不起作用
11-05 保存/转换时进度条卡住
11-06 保存的音频打不开/无法播放
12-01 不能使用/Don't work
12-02 不喜欢/表示很糟糕
12-03 要求评分
12-04 付费相关/抱怨付费
12-05 不安全/隐私问题/病毒
12-06 无意义/不相关差评
12-07 试一试/测试一下/先用用看
12-08 旧版本更好/不喜欢新版本
12-09 复杂/难以使用
12-10 表述不清暂无法归类
12-11 没有使用指导
12-12 其他
12-13 权限类问题
12-14 音频播放问题
12-15 GDPR弹窗
12-16 保存问题
12-17 没有声音
12-18 平板横屏问题
13-01 不能更改比特率
13-02 mp3和aac不能保存为其他格式
```

### 分类操作方式

数据量 ≤ 200 条时，由 Claude 直接逐条判断分类，无需 API 调用。

数据量 > 200 条时，建议用户使用本地 Python 脚本 + Anthropic API key 批量处理（询问用户是否有 API key）。

---

## Step 6：xlsx 格式要求

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 标题行样式
header_fill = PatternFill('solid', start_color='4472C4')
header_font = Font(bold=True, color='FFFFFF', name='Arial', size=11)

# 列宽参考
widths = [30, 45, 45, 12, 18, 18, 80]  # A到G

# 冻结首行 + 自动筛选
ws.freeze_panes = 'A2'
ws.auto_filter.ref = 'A1:G1'
```

---

## 注意事项

- 机翻公式在 xlsx 中只是字符串，需上传到 **Google Sheets** 才能生效
- 如用户对分类有异议，按用户反馈修正后重新生成文件
- 输出文件保存至 `/mnt/user-data/outputs/`
