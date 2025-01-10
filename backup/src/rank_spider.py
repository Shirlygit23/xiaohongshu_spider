import os
import time
import random
import pandas as pd
import requests
from datetime import datetime
from src.config.base_config import BaseConfig

class BaseRankSpider(BaseConfig):
    """榜单爬虫基类"""
    def __init__(self, category='鞋靴', days=7, max_pages=20):
        super().__init__()
        self.category = category
        self.days = days
        self.max_pages = max_pages
        self.data_dir = os.path.join(self.base_dir, 'data', 'rank_data')
        os.makedirs(self.data_dir, exist_ok=True)
        
        # 基础 URL
        self.base_url = 'https://xhsapi.huitun.com'
        
        # 添加请求头
        self.headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Cookie': self.COOKIES,
            'Origin': 'https://xhs.huitun.com',
            'Referer': 'https://xhs.huitun.com/',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
            'source': 'web',
            'vs': '141114881556.220172242.101',
            'Authorization': f'Bearer {self.TOKEN}',
            'Host': 'xhsapi.huitun.com',
            'Sec-Ch-Ua': '"Not A(Brand";v="99", "Google Chrome";v="121", "Chromium";v="121"',
            'Sec-Ch-Ua-Mobile': '?0',
            'Sec-Ch-Ua-Platform': '"Windows"',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-site'
        }

    def fetch_data(self, url, params):
        """发送请求获取数据"""
        current_timestamp = int(time.time() * 1000)
        full_url = f"{self.base_url}/{url}?_t={current_timestamp}"
        try:
            response = requests.get(full_url, headers=self.headers, params=params)
            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 0:
                    return data.get('extData', {}).get('list', [])
            print(f"请求失败: {response.status_code}")
            print(f"响应内容: {response.text}")
            return []
        except Exception as e:
            print(f"请求异常: {str(e)}")
            return []

    def get_params(self):
        # 计算近7天的日期
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - pd.Timedelta(days=self.days)).strftime('%Y-%m-%d')
        
        return {
            'page': 1,
            'pageSize': 20,
            'sort': 5,
            'rangeList': '1,2,3,5',
            'dateEnd': end_date,
            'dateStart': start_date,
            'days': self.days,
            'del': 'true',
            'tagList': '78',
            'tagNames': '鞋靴'
        }

    def parse_note(self, note_data):
        return {
            'title': note_data.get('title', ''),
            'user_name': note_data.get('nick', ''),
            'user_level': note_data.get('level', ''),
            'fans_count': note_data.get('fans', 0),
            'like_count': note_data.get('like', 0),
            'collect_count': note_data.get('coll', 0),
            'comment_count': note_data.get('comm', 0),
            'share_count': note_data.get('share', 0),
            'read_count': note_data.get('read', 0),
            'interact_count': note_data.get('stat', 0),
            'publish_time': note_data.get('postTime', '')
        }

    def save_data(self, data, rank_type):
        """保存数据到对应榜单文件"""
        filename = f"{rank_type}_{self.category}_{datetime.now().strftime('%Y%m%d')}.csv"
        filepath = os.path.join(self.data_dir, filename)
        df = pd.DataFrame(data)
        df.to_csv(filepath, index=False, encoding='utf-8-sig')

    def run(self):
        print(f"开始获取{self.__class__.__doc__}数据...")
        params = self.get_params()
        all_notes = []
        
        for page in range(1, self.max_pages + 1):
            print(f"正在获取第{page}页...")
            params['page'] = page
            data = self.fetch_data(self.rank_url, params)
            if not data:
                print(f"第{page}页获取失败或无数据")
                break
                
            notes = [self.parse_note(note) for note in data]
            all_notes.extend(notes)
            print(f"第{page}页获取成功，当前共{len(all_notes)}条数据")
            
            time.sleep(random.uniform(1, 3))
            
        if all_notes:
            self.save_data(all_notes, self.rank_type)
            print(f"数据保存成功，共{len(all_notes)}条")
        else:
            print("未获取到任何数据")


class HotRankSpider(BaseRankSpider):
    """实时热门榜爬虫"""
    def __init__(self, category='鞋靴', days=7, max_pages=20):
        super().__init__(category, days, max_pages)
        self.rank_url = 'note/search'
        self.rank_type = 'hot_rank'


class LowFansSpider(BaseRankSpider):
    """低粉爆文榜爬虫"""
    def __init__(self, category='鞋靴', days=7, max_pages=20):
        super().__init__(category, days, max_pages)
        self.rank_url = 'note/search'
        self.rank_type = 'low_fans_rank'


class CommercialSpider(BaseRankSpider):
    """商业笔记榜爬虫"""
    def __init__(self, category='鞋靴', days=7, max_pages=20):
        super().__init__(category, days, max_pages)
        self.rank_url = 'note/search'
        self.rank_type = 'commercial_rank'


class ShoppingSpider(BaseRankSpider):
    """带货笔记榜爬虫"""
    def __init__(self, category='鞋靴', days=7, max_pages=20):
        super().__init__(category, days, max_pages)
        self.rank_url = 'note/search'
        self.rank_type = 'shopping_rank'


def run_rank_spiders(category='鞋靴', days=7, max_pages=20):
    """运行所有榜单爬虫"""
    spiders = [
        HotRankSpider(category, days, max_pages),
        LowFansSpider(category, days, max_pages),
        CommercialSpider(category, days, max_pages),
        ShoppingSpider(category, days, max_pages)
    ]
    
    for spider in spiders:
        print(f"开始采集{spider.__doc__}")
        spider.run()
        print(f"完成采集{spider.__doc__}")


if __name__ == '__main__':
    # 先测试单个榜单
    spider = HotRankSpider(category='鞋靴', days=7, max_pages=2)
    spider.run()