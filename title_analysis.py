import pandas as pd
import jieba
from collections import Counter
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 配置文件路径
INPUT_FILE = "./input_data_xlsx/12月航海小红书演示数据使用.xlsx"
OUTPUT_FILE = "./output_data_xlsx/12月航海小红书演示数据使用output_key_head.xlsx"

# 停用词列表
STOP_WORDS = {'的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
              '之', '用', '于', '把', '等', '去', '给', '到', '又', '在',
              '或', '好', '来', '让', '但', '被', '什么', '这', '那', '你',
              '我', '他', '它', '她', '这个', '那个', '哪个', '谁', '啊',
              '吧', '呢', '么', '啦', '呀', '嘛', '哦', '噢', '哈', '嗯',
              '这样', '那样', '怎样', '如何', '为什么', '怎么', '多少'}

def read_excel():
    """读取Excel文件"""
    try:
        logging.info("开始读取Excel文件...")
        df = pd.read_excel(INPUT_FILE)
        if '笔记标题' not in df.columns:
            raise ValueError("Excel文件中未找到'笔记标题'列")
        logging.info(f"成功读取Excel文件，共{len(df)}行数据")
        return df['笔记标题']
    except Exception as e:
        logging.error(f"读取Excel文件失败: {str(e)}")
        raise

def extract_keywords(title):
    """提取标题中的关键词"""
    # 使用jieba进行分词
    words = jieba.cut(str(title))
    # 过滤停用词和单个字符
    keywords = [word for word in words if word not in STOP_WORDS and len(word) > 1]
    return keywords

def analyze_titles(titles):
    """分析标题关键词"""
    logging.info("开始分析标题关键词...")
    
    # 存储所有关键词及其对应的原标题
    keyword_titles = {}
    # 统计关键词频次
    keyword_counter = Counter()
    
    for title in titles:
        if pd.isna(title):  # 跳过空标题
            continue
            
        keywords = extract_keywords(title)
        # 更新关键词计数
        keyword_counter.update(keywords)
        # 记录关键词对应的原标题
        for keyword in keywords:
            if keyword not in keyword_titles:
                keyword_titles[keyword] = []
            keyword_titles[keyword].append(title)
    
    # 整理结果数据
    results = []
    for keyword, count in keyword_counter.most_common():
        results.append({
            '关键词': keyword,
            '出现次数': count,
            '相关标题': '\n'.join(set(keyword_titles[keyword]))  # 使用set去重
        })
    
    logging.info(f"分析完成，共找到{len(results)}个关键词")
    return results

def save_results(results):
    """保存分析结果到Excel"""
    try:
        logging.info("开始保存分析结果...")
        df = pd.DataFrame(results)
        df.to_excel(OUTPUT_FILE, index=False)
        logging.info(f"结果已保存至: {OUTPUT_FILE}")
    except Exception as e:
        logging.error(f"保存结果失败: {str(e)}")
        raise

def main():
    """主函数"""
    try:
        # 读取标题数据
        titles = read_excel()
        
        # 分析标题
        results = analyze_titles(titles)
        
        # 保存结果
        save_results(results)
        
    except Exception as e:
        logging.error(f"程序执行失败: {str(e)}")
        raise

if __name__ == "__main__":
    main() 