import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
from datetime import datetime
import time
import random

class NoteProcessor:
    def __init__(self, excel_path):
        """初始化处理器"""
        self.excel_path = excel_path
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.data_dir = os.path.join(self.base_dir, 'data')
        
        # 请求头设置
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Cookie': 'a1=18f67f0a967g79wyshll6yx9trgs54f1696alqylq50000213544; webId=4515c2cc4032caa7deee0fee9f2fa8e0; gid=yYiKWi8DJSKqyYiKWi80jjAvKWkWjEvMx11KK1AvCj9Fk028M24iy1888Jyq2448Sdq8djyJ; abRequestId=4515c2cc4032caa7deee0fee9f2fa8e0; xsecappid=xhs-pc-web; acw_tc=0a00dbf117358326346273919ebe69a9438ac140e802e681c89c811db0797b; webBuild=4.53.0; web_session=040069b32076e91226be0efc43354b14a648eb; unread={%22ub%22:%22676aa718000000000b01754b%22%2C%22ue%22:%22675513ae00000000020260ad%22%2C%22uc%22:32}; websectiga=f3d8eaee8a8c63016320d94a1bd00562d516a5417bc43a032a80cbf70f07d5c0; sec_poison_id=194a48a4-8734-4a54-b28f-ed65afd69c93',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Connection': 'keep-alive'
        }
        
    def process_notes(self):
        """处理笔记数据"""
        try:
            # 读取Excel文件
            print(f"开始读取文件: {self.excel_path}")
            df = pd.read_excel(self.excel_path)
            
            if '官方笔记地址' not in df.columns:
                print("错误：Excel文件中没有'官方笔记地址'列！")
                print("现有的列名：", df.columns.tolist())
                return
            
            # 添加新列
            df['笔记详情'] = ''
            df['笔记话题'] = ''
            
            total_notes = len(df)
            print(f"\n共发现 {total_notes} 条笔记待处理")
            
            # 处理每条笔记
            for idx, row in df.iterrows():
                note_url = row['官方笔记地址']
                if pd.isna(note_url):
                    continue
                
                print(f"\n处理第 {idx+1}/{total_notes} 条笔记")
                print(f"笔记地址: {note_url}")
                
                try:
                    # 获取笔记内容
                    response = requests.get(note_url, headers=self.headers, timeout=10)
                    if response.status_code != 200:
                        print(f"请求失败，状态码: {response.status_code}")
                        continue
                    
                    # 解析HTML
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # 获取笔记详情和话题
                    detail_desc = soup.find(id='detail-desc')
                    hash_tags = soup.find_all(id='hash-tag')
                    
                    # 处理笔记详情：先获取完整文本
                    if detail_desc:
                        full_text = detail_desc.text.strip()
                        # 如果有话题标签，从详情中移除
                        if hash_tags:
                            for tag in hash_tags:
                                tag_text = tag.text.strip()
                                full_text = full_text.replace(tag_text, '').strip()
                        df.at[idx, '笔记详情'] = full_text
                    
                    # 处理话题标签
                    if hash_tags:
                        tags = [tag.text.strip() for tag in hash_tags]
                        df.at[idx, '笔记话题'] = ','.join(tags)
                    
                    print("成功获取笔记内容")
                    if detail_desc:
                        print(f"详情长度: {len(full_text)}")
                    if hash_tags:
                        print(f"话题数量: {len(tags)}")
                    
                    # 随机延时，避免请求过快
                    time.sleep(random.uniform(1, 3))
                    
                except Exception as e:
                    print(f"处理笔记时出错: {str(e)}")
                    continue
            
            # 保存结果
            filename = os.path.basename(self.excel_path)
            name_without_ext = os.path.splitext(filename)[0]
            current_time = datetime.now().strftime('%Y%m%d_%H点%M分')
            output_filename = f'{name_without_ext}_详情补充_{current_time}.xlsx'
            output_path = os.path.join(self.data_dir, output_filename)
            
            df.to_excel(output_path, index=False)
            print(f"\n处理完成！")
            print(f"结果已保存至：{output_filename}")
            
        except Exception as e:
            print(f"处理过程出错: {str(e)}")
            import traceback
            print(traceback.format_exc())

def main():
    """主函数"""
    excel_path = r"C:\Users\Administrator\Desktop\code\xiaohongshu_spider\data\xiaohongshu_鞋子_20241223_16点04分 10条 .xlsx"
    
    if not os.path.exists(excel_path):
        print(f"错误：找不到文件 {excel_path}")
        return
        
    processor = NoteProcessor(excel_path)
    processor.process_notes()

if __name__ == '__main__':
    main()