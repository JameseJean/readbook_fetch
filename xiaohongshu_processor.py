import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 配置信息
EXCEL_PATH = "./input_data_xlsx/12月航海小红书演示数据使用.xlsx"  # 替换为实际的Excel文件路径
OUTPUT_PATH = "./output_data_xlsx/12月航海小红书演示数据使用output.xlsx"  # 替换为输出文件路径

# 请求头配置
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Cookie": "YOUR_COOKIE_HERE"  # 请替换为你的Cookie
}

def read_excel():
    """读取Excel文件"""
    try:
        logging.info("开始读取Excel文件...")
        df = pd.read_excel(EXCEL_PATH)
        logging.info(f"成功读取Excel文件，共{len(df)}行数据")
        return df
    except Exception as e:
        logging.error(f"读取Excel文件失败: {str(e)}")
        raise

def get_note_content(url):
    """获取笔记内容和话题标签"""
    try:
        logging.info(f"开始请求URL: {url}")
        # 添加随机延时
        time.sleep(2)
        
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 获取笔记详情
        detail_desc = soup.find(id="detail-desc")
        desc_content = detail_desc.text.strip() if detail_desc else ""
        
        # 获取话题标签
        hash_tags = soup.find_all(id="hash-tag")
        tags_content = ",".join([tag.text.strip() for tag in hash_tags]) if hash_tags else ""
        
        logging.info("成功获取笔记内容和话题标签")
        return desc_content, tags_content
    
    except requests.RequestException as e:
        logging.error(f"请求失败: {str(e)}")
        return "", ""
    except Exception as e:
        logging.error(f"解析内容失败: {str(e)}")
        return "", ""

def process_notes():
    """处理所有笔记"""
    try:
        # 读取Excel文件
        df = read_excel()
        
        # 创建新列
        df['笔记详情'] = ""
        df['笔记话题'] = ""
        
        # 获取第一列的URL
        urls = df.iloc[:, 0]
        
        # 处理每个URL
        for index, url in enumerate(urls):
            logging.info(f"正在处理第{index + 1}条数据...")
            desc, tags = get_note_content(url)
            df.at[index, '笔记详情'] = desc
            df.at[index, '笔记话题'] = tags
        
        # 保存结果
        logging.info("开始保存处理结果...")
        df.to_excel(OUTPUT_PATH, index=False)
        logging.info(f"处理完成，结果已保存至: {OUTPUT_PATH}")
        
    except Exception as e:
        logging.error(f"处理过程发生错误: {str(e)}")
        raise

if __name__ == "__main__":
    process_notes() 