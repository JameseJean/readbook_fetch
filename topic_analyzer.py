import pandas as pd
from collections import Counter
import logging
from pathlib import Path
import os

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 配置文件路径
INPUT_FILE = "D:/AImoney/coding/redbook_fetch/output_data_xlsx/12月航海小红书演示数据使用output.xlsx"
OUTPUT_DIR = "D:/AImoney/coding/redbook_fetch/output_data_xlsx"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "top_50_topics.txt")

def read_excel():
    """读取Excel文件"""
    try:
        logging.info("开始读取Excel文件...")
        df = pd.read_excel(INPUT_FILE)
        
        if '笔记话题' not in df.columns:
            raise ValueError("Excel文件中未找到'笔记话题'列")
            
        logging.info(f"成功读取Excel文件，共{len(df)}行数据")
        return df['笔记话题']
    except Exception as e:
        logging.error(f"读取Excel文件失败: {str(e)}")
        raise

def process_topics(topics_series):
    """处理话题数据"""
    logging.info("开始处理话题数据...")
    
    # 存储所有话题
    all_topics = []
    
    # 处理每行话题
    for topics in topics_series:
        if pd.isna(topics):
            continue
            
        # 分割话题
        topic_list = topics.split(',')  # 使用英文逗号分割
        
        # 清理话题格式
        cleaned_topics = [topic.strip() for topic in topic_list]
        all_topics.extend(cleaned_topics)
    
    # 统计话题频次
    topic_counter = Counter(all_topics)
    
    # 获取前50个最常见的话题
    top_50_topics = topic_counter.most_common(50)
    
    logging.info(f"话题处理完成，共发现{len(topic_counter)}个不同话题")
    return top_50_topics

def save_results(top_topics):
    """保存结果到文件"""
    try:
        # 创建输出目录
        Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
        
        logging.info("开始保存分析结果...")
        
        # 写入结果
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write("小红书话题Top50统计结果\n")
            f.write("=" * 50 + "\n\n")
            
            for index, (topic, count) in enumerate(top_topics, 1):
                f.write(f"{index}. {topic}: {count}次\n")
        
        logging.info(f"结果已保存至: {OUTPUT_FILE}")
        
    except Exception as e:
        logging.error(f"保存结果失败: {str(e)}")
        raise

def main():
    """主函数"""
    try:
        # 读取话题数据
        topics_data = read_excel()
        
        # 处理话题
        top_topics = process_topics(topics_data)
        
        # 保存结果
        save_results(top_topics)
        
    except Exception as e:
        logging.error(f"程序执行失败: {str(e)}")
        raise

if __name__ == "__main__":
    main() 