import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import time
import random
from config.base_config import base_config

class NoteProcessor:
    def __init__(self, excel_path=None):
        """初始化处理器"""
        self.excel_path = excel_path or base_config.EXCEL_PATH
        self.headers = base_config.HEADERS
        
    def process_notes(self):
        """处理笔记数据"""
        try:
            print(f"开始读取文件: {self.excel_path}")
            df = pd.read_excel(self.excel_path)
            
            # 添加调试信息
            print("\n数据列名:", df.columns.tolist())
            print(f"共发现 {len(df)} 条笔记")
            
            # 检查空值情况
            empty_details = df['笔记详情'].isna().sum()
            empty_topics = df['笔记话题'].isna().sum()
            print(f"空的笔记详情数量: {empty_details}")
            print(f"空的笔记话题数量: {empty_topics}")
            
            # 处理笔记
            for idx in range(len(df)):
                detail = df.iloc[idx]['笔记详情']
                topics = df.iloc[idx]['笔记话题']
                
                if not pd.isna(detail) or not pd.isna(topics):
                    continue
                
                note_url = df.iloc[idx]['官方笔记地址']
                print(f"\n处理第 {idx+1}/{len(df)} 条笔记")
                print(f"笔记地址: {note_url}")
                
                try:
                    response = requests.get(
                        note_url, 
                        headers=self.headers, 
                        timeout=base_config.REQUEST_TIMEOUT
                    )
                    print(f"请求状态码: {response.status_code}")
                    
                    if response.status_code != 200:
                        print(f"请求失败，状态码: {response.status_code}")
                        continue
                    
                    # 打印响应内容的一部分，用于调试
                    print("响应内容预览:", response.text[:200])
                    
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # 尝试不同的选择器
                    detail_desc = (
                        soup.find(id='detail-desc') or 
                        soup.find('div', class_='content') or
                        soup.find('div', class_='note-content')
                    )
                    
                    hash_tags = (
                        soup.find_all(id='hash-tag') or
                        soup.find_all('span', class_='tag-item') or
                        soup.find_all('a', class_='tag-link')
                    )
                    
                    if detail_desc:
                        print("找到笔记内容元素")
                        full_text = detail_desc.text.strip()
                        if hash_tags:
                            for tag in hash_tags:
                                tag_text = tag.text.strip()
                                full_text = full_text.replace(tag_text, '').strip()
                        df.at[idx, '笔记详情'] = full_text
                        print(f"提取的文本长度: {len(full_text)}")
                    else:
                        print("未找到笔记内容元素")
                    
                    if hash_tags:
                        print("找到话题标签元素")
                        tags = [tag.text.strip() for tag in hash_tags]
                        df.at[idx, '笔记话题'] = ','.join(tags)
                        print(f"提取的话题数量: {len(tags)}")
                    else:
                        print("未找到话题标签元素")
                    
                    # 每处理指定数量的记录保存一次
                    if (idx + 1) % base_config.BATCH_SIZE == 0:
                        df.to_excel(self.excel_path, index=False)
                        print(f"已保存到第 {idx+1} 条记录")
                    
                    time.sleep(random.uniform(*base_config.SLEEP_TIME))
                    
                except Exception as e:
                    print(f"处理笔记时出错: {str(e)}")
                    continue
            
            df.to_excel(self.excel_path, index=False)
            print("\n处理完成")
            
        except Exception as e:
            print(f"处理过程出错: {str(e)}")
            import traceback
            print(traceback.format_exc())

def main():
    """主函数"""
    processor = NoteProcessor()
    processor.process_notes()

if __name__ == '__main__':
    main()