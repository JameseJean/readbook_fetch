# 小红书笔记处理工具

一个用于批量获取小红书笔记内容和话题标签的Python工具。该工具可以从Excel文件中读取笔记URL，自动抓取每个笔记的详细内容和话题标签，并将结果保存到新的Excel文件中。

## 功能特点

- 批量处理小红书笔记URL
- 自动提取笔记详情内容
- 自动提取笔记话题标签
- 支持Excel文件读写
- 完整的错误处理机制
- 详细的日志记录
- 内置请求延时保护

## 环境要求

- Python 3.6+
- pandas
- requests
- beautifulsoup4
- openpyxl

## 安装步骤

1. 克隆项目到本地：
```bash
git clone [项目地址]
```

2. 安装依赖包：
```bash
pip install pandas requests beautifulsoup4 openpyxl
```

## 使用说明

1. 配置文件路径
   - 打开 `xiaohongshu_processor.py`
   - 修改 `EXCEL_PATH` 为输入Excel文件的路径
   - 修改 `OUTPUT_PATH` 为输出Excel文件的路径

2. 配置请求头
   - 在 `HEADERS` 中填入有效的Cookie
   - 根据需要修改User-Agent

3. 准备输入文件
   - Excel文件第一行为标题
   - 第一列为小红书笔记URL

4. 运行程序
```bash
python xiaohongshu_processor.py
```

## 注意事项

1. Cookie配置
   - 需要使用有效的小红书Cookie
   - Cookie可能需要定期更新
   - 建议使用登录状态的Cookie

2. 请求限制
   - 程序内置2秒延时，可根据需要调整
   - 建议不要过于频繁请求
   - 注意请求次数限制

3. 数据处理
   - 确保Excel文件格式正确
   - 注意URL格式的有效性
   - 输出文件如已存在会被覆盖

4. 错误处理
   - 程序会记录详细日志
   - 出错时检查日志文件
   - 网络错误时会自动跳过并继续处理

## 输出结果

程序将在输出Excel文件中添加两列：
- `笔记详情`：包含笔记的详细内容
- `笔记话题`：包含笔记的所有话题标签（以逗号分隔）

## 常见问题

1. 如果遇到请求失败，请检查：
   - Cookie是否有效
   - 网络连接是否正常
   - 请求频率是否过高

2. 如果数据为空，请确认：
   - URL是否有效
   - 页面结构是否变化
   - 是否有访问权限

3. 如果程序报错，请查看：
   - 日志文件内容
   - Excel文件格式
   - Python环境配置