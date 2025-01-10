# 在文件顶部的导入部分
import requests
import json
import pandas as pd
from datetime import datetime, timedelta  # 添加 timedelta 的导入
import time
import os
import random
import re
from enum import Enum
import logging

# 设置日志
logging.basicConfig(
    filename='spider_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

# 添加榜单类型枚举定义
class RankType(Enum):
    HOT = 'hot'           # 热门笔记榜
    LOW_FANS = 'lowfans'  # 低粉爆文榜
    BUSINESS = 'business' # 商业笔记榜
    SHOPPING = 'shopping' # 带货笔记榜

# 设置基础路径
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')

# 确保data目录存在
os.makedirs(DATA_DIR, exist_ok=True)

# 设置认证信息
# 设置认证信息
TOKEN = '0128d63f-7c77-4a5e-80bb-9474d9a8de53'
               
COOKIES = '_ga=GA1.1.979168562.1734707638; __root_domain_v=.huitun.com; _qddaz=QD.446234707638484; Hm_lvt_51956877b2ac5aabc38d224aa78a05d8=1734707635,1734765602; _ga_JBKBWWH0KV=GS1.1.1734796794.5.1.1734796876.0.0.0; xhsapiToken=0128d63f-7c77-4a5e-80bb-9474d9a8de53; _clck=1plil31%7C2%7Cfsg%7C0%7C1815; _clsk=mr6uza%7C1736485915084%7C2%7C1%7Ck.clarity.ms%2Fcollect; _ga_KNCC792VGY=GS1.1.1736485643.38.1.1736486030.0.0.0'

# 鞋类话题关键词常量
SHOE_TOPIC_KEYWORDS = {
    # 鞋类品类
    '球鞋', '运动鞋', '高跟鞋', '单鞋', '板鞋', '德训鞋', '马丁靴', 
    '雪地靴', '拖鞋', '凉鞋', '帆布鞋', '小白鞋', '老爹鞋', '休闲鞋',
    '增高鞋', '乐福鞋', '豆豆鞋', '平底鞋', '跑鞋', '靴子', '鞋', '靴',
    # 品牌相关
    'nike', 'aj', '耐克', '阿迪达斯', 'adidas', 'newbalance', 'nb', 
    '匡威', 'converse', 'vans', 'puma', 'skechers', '李宁', '安踏', 
    '特步', '彪马', '亚瑟士', 'asics', '锐步', 'reebok',
    # 口语化表达
    'jiojio', '脚脚', '鞋鞋', '小皮鞋', '小鞋子',
    '脚踝', '脚背', '脚掌', '脚型', '脚感',
    # 相关描述
    '穿搭', '搭配', '上脚', '配色', '鞋款',
    '脚踩', '踩踏', '足部', '足底', '足弓'
}

# 鞋类内容关键词
SHOE_CONTENT_KEYWORDS = {
    # 原有关键词
    '鞋码', '鞋型', '鞋垫', '鞋帮', '鞋面', '鞋底', 
    '穿搭', '脚感', '试穿', '上脚', '码数', '尺码',
    '鞋楦', '鞋跟', '鞋带', '内里', '大底', '中底',
    '鞋身', '后跟', '鞋头',
    # 添加口语化表达
    'jiojio', '脚脚', '鞋鞋', '小皮鞋', '小鞋子',
    '脚踝', '脚背', '脚掌', '脚型', '脚感',
    '穿搭', '搭配', '上脚', '配色', '鞋款',
    '脚踩', '踩踏', '足部', '足底', '足弓'
}

def check_shoe_topics(topics):
    """检查话题是否包含鞋类相关内容"""
    for topic in topics:
        topic_lower = topic.lower()
        if any(keyword in topic_lower for keyword in SHOE_TOPIC_KEYWORDS):
            return True
    return False

def check_shoe_keywords(title, keywords):
    """检查标题和关键词是否包含鞋类相关内容"""
    text = (title + ' ' + keywords).lower()
    return any(keyword in text for keyword in SHOE_TOPIC_KEYWORDS)

def get_date_range(rank_type):
    """根据榜单类型获取不同的日期范围"""
    current_date = datetime.now()
    
    if rank_type in [RankType.BUSINESS, RankType.SHOPPING]:
        # 商业榜和带货榜使用固定周期（周榜）
        # 获取上周的周一到周日
        today = current_date.date()
        last_monday = today - timedelta(days=today.weekday() + 7)
        last_sunday = last_monday + timedelta(days=6)
        return last_monday.strftime('%Y-%m-%d'), last_sunday.strftime('%Y-%m-%d')
    else:
        # 热门榜和低粉榜使用近7天
        end_date = current_date.date()
        start_date = end_date - timedelta(days=7)
        return start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')

def get_rank_notes(rank_type):
    """获取指定类型的榜单数据"""
    all_notes = []
    page = 1
    max_pages = 5
    
    start_date_str, end_date_str = get_date_range(rank_type)
    
    while page <= max_pages:
        sleep_time = random.uniform(2, 3)
        time.sleep(sleep_time)
        
        current_timestamp = int(time.time() * 1000)        
    
        # 修复 URL 中的 timeRange 参数（之前有重复的 =7）
        if rank_type == RankType.HOT:
            url = f'https://xhsapi.huitun.com/rank/hotNote?_t={current_timestamp}&gender=&noteType=&page={page}&pageSize=10&sort=0&tagIdList=78&timeRange=7'
        elif rank_type == RankType.LOW_FANS:
            url = f'https://xhsapi.huitun.com/rank/lowPowderExplosionText?_t={current_timestamp}&page={page}&pageSize=10&sort=0&tagIdList=78&timeRange=7'
        elif rank_type == RankType.BUSINESS:
            url = f'https://xhsapi.huitun.com/rank/bussNote?_t={current_timestamp}&page={page}&pageSize=20&sort=0&tagIdList=78&dateStart={start_date_str}&dateEnd={end_date_str}'
        elif rank_type == RankType.SHOPPING:
            url = f'https://xhsapi.huitun.com/rank/goodsNote?_t={current_timestamp}&page={page}&pageSize=20&sort=0&catIds=72&dateStart={start_date_str}&dateEnd={end_date_str}'
        
        logging.info(f"请求URL: {url}")

        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Cookie': COOKIES,
            'Origin': 'https://xhs.huitun.com',
            'Referer': 'https://xhs.huitun.com/',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
            'source': 'web',
            'vs': '141114881556.220172242.101',
            'Authorization': f'Bearer {TOKEN}',
            'Host': 'xhsapi.huitun.com',
            'Sec-Ch-Ua': '"Not A(Brand";v="99", "Google Chrome";v="121", "Chromium";v="121"',
            'Sec-Ch-Ua-Mobile': '?0',
            'Sec-Ch-Ua-Platform': '"Windows"',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-site'
        }
        
        try:
            response = requests.get(url, headers=headers)
            logging.info(f"响应状态码: {response.status_code}")
            
            if response.status_code == 200:
                try:
                    # 记录响应内容，帮助调试
                    logging.info(f"响应内容: {response.text[:200]}...")
                    
                    data = response.json()
                    # 先检查 data 是否为 None
                    if data is None:
                        logging.error("API返回空数据")
                        break
                        
                    if data.get('status') == 0:
                        note_list = data.get('extData', {}).get('list', [])
                        if note_list:
                            all_notes.extend(note_list)
                            page += 1
                        else:
                            break  # 如果没有数据了就退出循环
                    else:
                        logging.error(f"API返回错误状态: {data.get('status')}, 错误信息: {data.get('message', '未知错误')}")
                        break
                except json.JSONDecodeError as e:
                    logging.error(f"JSON解析失败: {str(e)}")
                    logging.error(f"响应内容: {response.text}")
                    break
            else:
                logging.error(f"请求状态码: {response.status_code}")
                logging.error(f"响应内容: {response.text}")
                break
        except Exception as e:
            logging.error(f"请求失败: {str(e)}")
            break
            
    return all_notes

def get_note_detail(note_id):
    """获取笔记详情"""
    current_timestamp = int(time.time() * 1000)
    url = f'https://xhsapi.huitun.com/note/detail?_t={current_timestamp}&noteId={note_id}'
    
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Cookie': COOKIES,
        'Origin': 'https://xhs.huitun.com',
        'Referer': 'https://xhs.huitun.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'source': 'web',
        'vs': '141114881556.220172242.101',
        'Authorization': f'Bearer {TOKEN}',
        'Host': 'xhsapi.huitun.com',
        'Sec-Ch-Ua': '"Not A(Brand";v="99", "Google Chrome";v="121", "Chromium";v="121"',
        'Sec-Ch-Ua-Mobile': '?0',
        'Sec-Ch-Ua-Platform': '"Windows"',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-site'
    }
    
    try:
        response = requests.get(url, headers=headers)
        data = response.json()
        if data.get('status') == 0:
            note_data = data.get('extData', {})
            
            # 获取原始描述文本
            desc = note_data.get('desc', '')
            
            # 提取话题标签和清理内容
            topics = []
            clean_desc = desc
            if desc:
                # 提取话题标签
                topics = re.findall(r'#([^#\[]+?)(?:\[话题\])?#', desc)
                # 清理内容中的话题标签
                clean_desc = re.sub(r'#[^#]+?\[话题\]#', '', desc).strip()
            
            return {
                'noteLink': note_data.get('noteLink', ''),
                'desc': clean_desc,        # 清理后的笔记内容
                'topics': topics,          # 话题列表
                'type': note_data.get('type', 'normal'),
                'title': note_data.get('title', ''),
                'postTime': note_data.get('postTime', '')
            }
        return {}
    except Exception as e:
        logging.error(f"获取笔记详情失败: {str(e)}")
        return {}

def process_notes_data(notes_list):
    """处理笔记数据并转换为DataFrame格式"""
    processed_data = []
    keyword_filtered_count = 0
    topic_filtered_count = 0
    
    print("\n=== 数据分析 ===")
    print(f"获取到 {len(notes_list)} 条笔记，开始过滤...")
    
    # 1. 先进行关键词过滤
    filtered_notes = []
    for note in notes_list:
        title = note.get('title', '').lower()
        keywords = note.get('keyw', '').lower()
        
        if check_shoe_keywords(title, keywords):
            filtered_notes.append(note)
        else:
            print(f"\n关键词过滤掉：")
            print(f"- 标题：{note.get('title')}")
            print(f"- 关键词：{note.get('keyw')}")
            keyword_filtered_count += 1
    
    print(f"关键词过滤后保留 {len(filtered_notes)} 条笔记")
    
    # 2. 获取详情并进行话题过滤
    for note in filtered_notes:
        note_detail = get_note_detail(note.get('noteId', ''))
        content = note_detail.get('desc', '')
        topics = note_detail.get('topics', [])
        
        # 检查话题标签
        if not check_shoe_topics(topics):
            print(f"\n话题过滤掉：")
            print(f"- 标题：{note.get('title')}")
            print(f"- 话题：{', '.join(topics)}")
            topic_filtered_count += 1
            continue
            
        # 获取粉丝数和点赞数
        fans_count = note.get('fans', 0)
        like_count = note.get('like', 0)
        
        processed_note = {
            '官方笔记地址': note_detail.get('noteLink', ''),
            '笔记标题': note.get('title'),
            '笔记内容': content,
            '话题标签': '，'.join(topics),
            '关键词': note.get('keyw'),
            '预估阅读量': note.get('read', 0),
            '互动量': note.get('stat', 0),
            '点赞数': like_count,
            '收藏数': note.get('coll', 0),
            '评论数': note.get('comm', 0),
            '分享数': note.get('share', 0),
            '发布时间': note_detail.get('postTime', ''),
            '笔记类型': '视频笔记' if note_detail.get('type') == 'video' else '图文笔记',
            '封面地址': note.get('imageUrl', ''),
            '达人名称': note.get('nick', ''),
            '小红书号': note.get('redId', ''),
            '粉丝数': fans_count
        }
        
        processed_data.append(processed_note)
        time.sleep(random.uniform(0.5, 1))
    
    print(f"\n过滤完成：")
    print(f"- 原始数据：{len(notes_list)} 条")
    print(f"- 关键词过滤掉：{keyword_filtered_count} 条")
    print(f"- 话题过滤掉：{topic_filtered_count} 条")
    print(f"- 最终保留：{len(processed_data)} 条")
    
    # 最终结果按点赞数排序
    result_df = pd.DataFrame(processed_data)
    result_df = result_df.sort_values(by='点赞数', ascending=False)
    return result_df

def save_to_excel(all_notes, excel_path):
    """保存数据到Excel文件"""
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for rank_type, notes in all_notes.items():
                if notes:
                    df = process_notes_data(notes)
                    if not df.empty:
                        sheet_name = {
                            'hot': '热门榜单',
                            'lowfans': '低粉榜单',
                            'business': '商业榜单',
                            'shopping': '带货榜单'
                        }.get(rank_type, rank_type)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
        logging.info(f"数据已保存到Excel文件: {excel_path}")
        return True
    except Exception as e:
        logging.error(f"保存Excel文件失败: {str(e)}")
        return False

def save_to_csv(all_notes, base_path):
    """保存数据到CSV文件"""
    try:
        for rank_type, notes in all_notes.items():
            if notes:
                df = process_notes_data(notes)
                if not df.empty:
                    csv_path = f"{base_path}_{rank_type}.csv"
                    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    logging.info(f"数据已保存到CSV文件: {csv_path}")
        return True
    except Exception as e:
        logging.error(f"保存CSV文件失败: {str(e)}")
        return False

def main():
    logging.info("=== 程序开始运行 ===")
    all_notes = {}  # 使用字典存储不同榜单的数据
    
    # 要抓取的榜单类型列表
    rank_types = [RankType.HOT, RankType.LOW_FANS, RankType.BUSINESS, RankType.SHOPPING]
    
    try:
        # 遍历每个榜单类型
        for rank_type in rank_types:
            logging.info(f"=== 开始获取{rank_type.value}榜单数据 ===")
            notes = get_rank_notes(rank_type)  # 直接调用get_rank_notes，不需要传page参数
            
            if notes:
                all_notes[rank_type.value] = notes
                logging.info(f"{rank_type.value}榜单成功获取 {len(notes)} 条数据")
            else:
                logging.error(f"{rank_type.value}榜单数据获取失败，请检查TOKEN和COOKIES是否过期")
            
            # 添加榜单之间的间隔时间
            time.sleep(random.uniform(3, 5))

        # 处理数据并保存
        if any(all_notes.values()):
            current_time = datetime.now().strftime('%Y%m%d_%H点%M分')
            file_name = f'xiaohongshu_鞋靴_多榜单_{current_time}'
            excel_path = os.path.join(DATA_DIR, f'{file_name}.xlsx')
            
            # 尝试保存为Excel
            if not save_to_excel(all_notes, excel_path):
                # 如果Excel保存失败，尝试保存为CSV
                csv_base_path = os.path.join(DATA_DIR, file_name)
                if not save_to_csv(all_notes, csv_base_path):
                    logging.error('所有保存尝试都失败了')
        else:
            logging.warning('没有获取到任何数据')
    except Exception as e:
        logging.error(f"程序运行出错: {str(e)}")
    finally:
        logging.info("=== 程序运行结束 ===")

if __name__ == '__main__':
    main()