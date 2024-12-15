import pandas as pd
import requests
import os
import logging
import time
from urllib.parse import urlparse
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 配置文件路径
INPUT_FILE = "./input_data_xlsx/12月航海小红书演示数据使用.xlsx"
OUTPUT_DIR = "./output_data_xlsx/image/"

# 请求头配置
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

def create_output_dir():
    """创建输出目录"""
    try:
        Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
        logging.info(f"输出目录已准备: {OUTPUT_DIR}")
    except Exception as e:
        logging.error(f"创建输出目录失败: {str(e)}")
        raise

def get_filtered_data():
    """读取并筛选数据"""
    try:
        logging.info("开始读取Excel文件...")
        df = pd.read_excel(INPUT_FILE)
        
        # 检查必要的列是否存在
        required_columns = ['粉丝数', '互动量', '封面地址']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中未找到'{col}'列")
        
        # 筛选符合条件的数据
        filtered_df = df[
            (df['粉丝数'] < 1000) & 
            (df['互动量'] > 100)
        ]
        
        logging.info(f"筛选出{len(filtered_df)}条符合条件的数据")
        return filtered_df
    
    except Exception as e:
        logging.error(f"读取或筛选数据失败: {str(e)}")
        raise

def get_safe_filename(url):
    """生成安全的文件名"""
    # 从URL中获取文件名
    parsed_url = urlparse(url)
    original_filename = os.path.basename(parsed_url.path)
    
    # 如果URL中没有文件名，使用URL的哈希值
    if not original_filename:
        original_filename = str(hash(url)) + '.jpg'
    
    # 确保文件名唯一
    base_name = os.path.splitext(original_filename)[0]
    extension = os.path.splitext(original_filename)[1] or '.jpg'
    filename = base_name + extension
    counter = 1
    
    while os.path.exists(os.path.join(OUTPUT_DIR, filename)):
        filename = f"{base_name}_{counter}{extension}"
        counter += 1
    
    return filename

def download_image(url, filename):
    """下载单个图片"""
    try:
        # ���加随机延时
        time.sleep(1)
        
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        
        file_path = os.path.join(OUTPUT_DIR, filename)
        with open(file_path, 'wb') as f:
            f.write(response.content)
        
        logging.info(f"成功下载图片: {filename}")
        return True
    
    except requests.RequestException as e:
        logging.error(f"下载图片失败 {url}: {str(e)}")
        return False
    except Exception as e:
        logging.error(f"保存图片失败 {filename}: {str(e)}")
        return False

def process_images():
    """处理所有符合条件的图片"""
    try:
        # 创建输出目录
        create_output_dir()
        
        # 获取筛选后的数据
        df = get_filtered_data()
        
        # 统计信息
        total_images = len(df)
        successful_downloads = 0
        
        # 下载图片
        for index, row in df.iterrows():
            image_url = row['封面地址']
            if pd.isna(image_url):
                logging.warning(f"第{index + 1}行的封面地址为空，跳过")
                continue
                
            filename = get_safe_filename(image_url)
            if download_image(image_url, filename):
                successful_downloads += 1
                
            # 显示进度
            progress = (index + 1) / total_images * 100
            logging.info(f"处理进度: {progress:.2f}% ({index + 1}/{total_images})")
        
        # 输出统计信息
        logging.info(f"图片下载完成: 成功{successful_downloads}个，失败{total_images - successful_downloads}个")
        
    except Exception as e:
        logging.error(f"处理过程发生错误: {str(e)}")
        raise

if __name__ == "__main__":
    process_images() 