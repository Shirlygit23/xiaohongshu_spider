import pandas as pd
import os
import requests
from datetime import datetime
import re
from urllib.parse import unquote
from concurrent.futures import ThreadPoolExecutor
import time
import random
from config.base_config import base_config

class ImageDownloader:
    def __init__(self, excel_path=None):
        """初始化下载器"""
        self.excel_path = excel_path or base_config.EXCEL_PATH
        self.data_dir = base_config.DATA_DIR
        
    def create_save_dir(self, prefix, keyword):
        """创建保存目录"""
        current_time = datetime.now().strftime('%Y%m%d_%H点%M分')
        dir_name = f'{prefix}_{keyword}_{current_time}'
        save_dir = os.path.join(self.data_dir, dir_name)
        os.makedirs(save_dir, exist_ok=True)
        return save_dir
        
    def download_image(self, url, save_path):
        """下载单个图片"""
        try:
            response = requests.get(
                url, 
                headers=base_config.HEADERS, 
                timeout=base_config.REQUEST_TIMEOUT
            )
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    f.write(response.content)
                return True
            return False
        except Exception as e:
            print(f"下载图片失败: {url}")
            print(f"错误信息: {str(e)}")
            return False
            
    def download_images(self, filtered_df, save_dir):
        """批量下载图片"""
        success_count = 0
        with ThreadPoolExecutor(max_workers=base_config.MAX_CONCURRENT_DOWNLOADS) as executor:
            for idx, row in filtered_df.iterrows():
                cover_url = row['封面地址']
                if pd.isna(cover_url):
                    continue
                    
                # 从URL中提取文件名
                file_name = f"image_{idx+1}.jpg"
                save_path = os.path.join(save_dir, file_name)
                
                # 下载图片
                if self.download_image(cover_url, save_path):
                    success_count += 1
                
                # 添加延时避免请求过快
                time.sleep(random.uniform(*base_config.SLEEP_TIME))
        
        return success_count

    def process_low_fans_data(self):
        """处理低粉丝数高互动数据"""
        try:
            df = pd.read_excel(self.excel_path)
            
            # 从文件名提取关键词
            filename = os.path.basename(self.excel_path)
            keyword_match = re.search(r'xiaohongshu_(.+?)_', filename)
            keyword = keyword_match.group(1) if keyword_match else '未知关键词'
            
            # 创建保存目录
            save_dir = self.create_save_dir('低粉豹纹封面', keyword)
            
            # 使用原有的阈值
            filtered_df = df[
                (df['粉丝数'] < 1000) & 
                (df['互动量'] > 100)
            ].copy()
            
            if filtered_df.empty:
                print("未找到符合低粉丝数高互动的数据")
                return
            
            print(f"\n找到 {len(filtered_df)} 条低粉丝数高互动的数据")
            print("开始下载低粉丝数高互动的封面图片...")
            
            success_count = self.download_images(filtered_df, save_dir)
            
            print(f"\n低粉丝数高互动封面下载完成！")
            print(f"成功下载: {success_count} 张图片")
            print(f"保存位置: {save_dir}")
            
        except Exception as e:
            print(f"处理低粉丝数高互动数据时出错: {str(e)}")
            import traceback
            print(traceback.format_exc())

    def process_viral_data(self):
        """处理爆文数据"""
        try:
            df = pd.read_excel(self.excel_path)
            
            # 从文件名提取关键词
            filename = os.path.basename(self.excel_path)
            keyword_match = re.search(r'xiaohongshu_(.+?)_', filename)
            keyword = keyword_match.group(1) if keyword_match else '未知关键词'
            
            # 创建保存目录
            save_dir = self.create_save_dir('爆文封面', keyword)
            
            # 使用已有的爆款阈值配置
            filtered_df = df[df['点赞数'] > base_config.TITLE_VIRAL_THRESHOLD].copy()
            
            if filtered_df.empty:
                print("未找到爆文数据")
                return
            
            print(f"\n找到 {len(filtered_df)} 条爆文数据")
            print("开始下载爆文封面图片...")
            
            success_count = self.download_images(filtered_df, save_dir)
            
            print(f"\n爆文封面下载完成！")
            print(f"成功下载: {success_count} 张图片")
            print(f"保存位置: {save_dir}")
            
        except Exception as e:
            print(f"处理爆文数据时出错: {str(e)}")
            import traceback
            print(traceback.format_exc())

def main():
    """主函数"""
    downloader = ImageDownloader()
    
    # 下载低粉丝数高互动的封面
    print("\n=== 开始处理低粉丝数高互动数据 ===")
    downloader.process_low_fans_data()
    
    # 下载爆文封面
    print("\n=== 开始处理爆文数据 ===")
    downloader.process_viral_data()
    
    print("\n=== 所有下载任务完成 ===")

if __name__ == '__main__':
    main()