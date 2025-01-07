# 小红书爬虫项目

一个基于 Python 的小红书笔记爬虫工具，用于获取和分析鞋类相关的笔记数据。

## 功能特点

- 支持关键词搜索和类目搜索两种方式
- 自动过滤和分析笔记内容
- 识别商业笔记、带货笔记和低粉爆文
- 数据导出为 Excel 格式

## 项目结构
xiaohongshu_spider/
├── src/
│ ├── huitun_v2.py # 主程序文件
│ ├── image_downloader.py # 图片下载模块
│ ├── note_processor.py # 笔记处理模块
│ ├── title_analyzer.py # 标题分析模块
│ └── topic_analyzer.py # 话题分析模块
├── data/ # 数据输出目录
└── README.md

## 使用方法

1. 安装依赖：
bash
pip install requests pandas

2. 配置认证信息：
- 在 `huitun_v2.py` 中设置 `TOKEN` 和 `COOKIES`

3. 运行程序：
bash
python src/huitun_v2.py

## 数据输出

程序会在 `data` 目录下生成两个 Excel 文件：
- 原始数据文件：包含未经过滤的搜索结果
- 处理后数据文件：包含经过筛选和分析的最终数据

## 注意事项

- 需要有效的慧吞 API 认证信息
- 建议遵守小红书的爬虫规范
- 仅用于学习研究使用
