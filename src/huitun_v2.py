import requests
import json
import pandas as pd
from datetime import datetime
import time
import os
import random
import re

# 设置基础路径
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')

# 确保data目录存在
os.makedirs(DATA_DIR, exist_ok=True)

# 设置认证信息
TOKEN = '4090e538-4fb8-41d3-832b-7d7fc454d457'

COOKIES = '_ga=GA1.1.979168562.1734707638; __root_domain_v=.huitun.com; _qddaz=QD.446234707638484; Hm_lvt_51956877b2ac5aabc38d224aa78a05d8=1734707635,1734765602; _ga_JBKBWWH0KV=GS1.1.1734796794.5.1.1734796876.0.0.0; xhsapiToken=4090e538-4fb8-41d3-832b-7d7fc454d457; _clck=1plil31%7C2%7Cfsc%7C0%7C1815; _clsk=1wi3aqc%7C1736150257038%7C2%7C1%7Cn.clarity.ms%2Fcollect; _ga_KNCC792VGY=GS1.1.1736149626.22.1.1736150630.0.0.0'

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

def check_commercial(content):
    """判断是否是商业笔记"""
    commercial_keywords = [
        '广告', '推广', '#广告', '合作', '赞助', 
        '品牌方', '商务合作', '商业合作', '推荐商家',
        '安利', '种草', '@商家', '商家'
    ]
    return any(keyword in content.lower() for keyword in commercial_keywords)

def check_shopping(content):
    """判断是否是带货笔记"""
    shopping_keywords = [
        '购买链接', '买链接', '购入', '链接在评论', '链接在简介',
        '购买渠道', '购买方式', '下单', '链接在', '店铺',
        '专柜', '专卖店', '旗舰店', '价格', '¥', '￥',
        '代购', '团购', '预售', '优惠券'
    ]
    return any(keyword in content.lower() for keyword in shopping_keywords)

def check_low_fans_hit(fans_count, like_count):
    """判断是否是低粉爆文"""
    return fans_count < 10000 and like_count > 1000

def check_shoe_topics(topics):
    """检查话题是否包含鞋类相关内容"""
    for topic in topics:
        topic_lower = topic.lower()
        if any(keyword in topic_lower for keyword in SHOE_TOPIC_KEYWORDS):
            return True
    return False

def check_shoe_content(content):
    """检查内容是否包含鞋类相关描述"""
    content_lower = content.lower()
    return any(keyword in content_lower for keyword in SHOE_CONTENT_KEYWORDS)

def check_shoe_keywords(title, keywords):
    """检查标题和关键词是否包含鞋类相关内容"""
    text = (title + ' ' + keywords).lower()
    return any(keyword in text for keyword in SHOE_TOPIC_KEYWORDS)

def get_notes_data(page):
    """通过关键词获取指定页码的笔记数据"""
    sleep_time = random.uniform(2, 3)
    time.sleep(sleep_time)
    
    current_timestamp = int(time.time() * 1000)
    
    # 计算近7天的日期
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - pd.Timedelta(days=7)).strftime('%Y-%m-%d')
    
    url = f'https://xhsapi.huitun.com/note/search?_t={current_timestamp}&page={page}&pageSize=20&sort=5&rangeList=1,2,3,5&dateEnd={end_date}&dateStart={start_date}&days=7&del=true&keyword=%E9%9E%8B%E5%AD%90'
    
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
        if response.status_code == 200:
            data = response.json()
            if data.get('status') == 0:
                note_list = data.get('extData', {}).get('list', [])
                if note_list:
                    normal_count = len([n for n in note_list if n.get('type') == 'normal'])
                    video_count = len([n for n in note_list if n.get('type') == 'video'])
                    print(f"关键词搜索获取到 {len(note_list)} 条数据 (图文: {normal_count}, 视频: {video_count})")
                    for note in note_list:
                        note['source'] = 'keyword'  # 标记为关键词搜索
                    return note_list
        print(f"请求状态码: {response.status_code}")
        return []
    except Exception as e:
        print(f"请求失败: {str(e)}")
        return []
    
def get_category_notes(page):
    """通过类目获取指定页码的笔记数据"""
    sleep_time = random.uniform(2, 3)
    time.sleep(sleep_time)
    
    current_timestamp = int(time.time() * 1000)
    
    # 计算近7天的日期
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - pd.Timedelta(days=7)).strftime('%Y-%m-%d')
    
    url = f'https://xhsapi.huitun.com/note/search?_t={current_timestamp}&page={page}&pageSize=20&sort=5&rangeList=1,2,3,5&dateEnd={end_date}&dateStart={start_date}&days=7&del=true&tagList=78&tagNames=%E9%9E%8B%E9%9D%B4'
    
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
        if response.status_code == 200:
            data = response.json()
            if data.get('status') == 0:
                note_list = data.get('extData', {}).get('list', [])
                if note_list:
                    normal_count = len([n for n in note_list if n.get('type') == 'normal'])
                    video_count = len([n for n in note_list if n.get('type') == 'video'])
                    print(f"类目搜索获取到 {len(note_list)} 条数据 (图文: {normal_count}, 视频: {video_count})")
                    for note in note_list:
                        note['source'] = 'category'  # 标记为类目搜索
                    return note_list
        print(f"请求状态码: {response.status_code}")
        return []
    except Exception as e:
        print(f"请求失败: {str(e)}")
        return []

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
        print(f"获取笔记详情失败: {str(e)}")
        return {}
    
def process_notes_data(notes_list):
    """处理笔记数据并转换为DataFrame格式"""
    processed_data = []
    keyword_filtered_count = 0
    topic_filtered_count = 0
    
    # 导出过滤前的原始数据
    original_notes = []
    for note in notes_list:
        note_id = note.get('noteId', '')
        # 获取笔记详情以获取正确的链接
        note_detail = get_note_detail(note_id)
        original_note = {
            '笔记标题': note.get('title'),
            '关键词': note.get('keyw'),
            '笔记ID': note_id,
            '笔记链接': note_detail.get('noteLink', ''),
            '点赞数': note.get('like', 0),
            '粉丝数': note.get('fans', 0),
            '来源': note.get('source', '')  # 添加来源字段
        }
        original_notes.append(original_note)
        time.sleep(random.uniform(0.5, 1))

    # 创建Excel写入器
    timestamp = datetime.now().strftime("%H点%M分")
    original_file = os.path.join(DATA_DIR, f'原始数据_{timestamp}.xlsx')
    
    # 分离关键词搜索和类目搜索的数据
    keyword_notes = [note for note in original_notes if note.get('来源') == 'keyword']
    category_notes = [note for note in original_notes if note.get('来源') == 'category']
    
    # 创建DataFrame并排序
    keyword_df = pd.DataFrame(keyword_notes)
    category_df = pd.DataFrame(category_notes)
    
    if not keyword_df.empty:
        keyword_df = keyword_df.sort_values(by='点赞数', ascending=False)
    if not category_df.empty:
        category_df = category_df.sort_values(by='点赞数', ascending=False)
    
    # 保存到Excel的不同sheet
    with pd.ExcelWriter(original_file, engine='openpyxl') as writer:
        keyword_df.to_excel(writer, sheet_name='关键词搜索', index=False)
        category_df.to_excel(writer, sheet_name='类目搜索', index=False)
    
    print(f"原始数据已保存到: {original_file}")
    
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
            '粉丝数': fans_count,
            '是否商业笔记': '是' if check_commercial(content) else '否',
            '是否带货笔记': '是' if check_shopping(content) else '否',
            '是否低粉爆文': '是' if check_low_fans_hit(fans_count, like_count) else '否'
        }
        
        # 如果是低粉爆文，打印提示
        if check_low_fans_hit(fans_count, like_count):
            print(f"\n发现低粉爆文：")
            print(f"- 标题：{note.get('title')}")
            print(f"- 粉丝数：{fans_count}")
            print(f"- 点赞数：{like_count}")
        
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

def merge_and_deduplicate(keyword_notes, category_notes):
    """合并两种搜索结果并去重"""
    all_notes = []
    note_ids = set()
    
    # 合并并去重
    for note in (keyword_notes + category_notes):
        note_id = note.get('noteId')
        if note_id and note_id not in note_ids:
            all_notes.append(note)
            note_ids.add(note_id)
    
    # 按点赞量排序
    all_notes.sort(key=lambda x: x.get('like', 0), reverse=True)
    
    print(f"\n=== 数据去重 ===")
    print(f"关键词搜索: {len(keyword_notes)} 条")
    print(f"类目搜索: {len(category_notes)} 条")
    print(f"去重后总数: {len(all_notes)} 条")
    
    return all_notes

def main():
    print("=== 程序开始运行 ===")
    all_notes = []
    
    try:
        # 1. 关键词搜索（1页20条）
        keyword_notes = []
        for page in range(1, 2):
            print(f'\n正在获取关键词搜索第{page}页数据...')
            notes = get_notes_data(page)
            if notes:
                keyword_notes.extend(notes)
            else:
                print("关键词搜索获取数据失败，请检查TOKEN和COOKIES是否过期")
        
        # 2. 类目搜索（1页20条）
        category_notes = []
        for page in range(1, 2):
            print(f'\n正在获取类目搜索第{page}页数据...')
            notes = get_category_notes(page)
            if notes:
                category_notes.extend(notes)
            else:
                print("类目搜索获取数据失败，请检查TOKEN和COOKIES是否过期")
        
        # 3. 合并去重
        all_notes = merge_and_deduplicate(keyword_notes, category_notes)
        
        # 4. 处理数据并保存
        if all_notes:
            result_df = process_notes_data(all_notes)
            try:
                current_time = datetime.now().strftime('%Y%m%d_%H点%M分')
                file_name = f'xiaohongshu_鞋靴_关键词和类目_{current_time}.xlsx'
                
                save_path = os.path.join(DATA_DIR, file_name)
                result_df.to_excel(save_path, index=False)
                print(f'数据已保存到: {save_path}')
                print(f'共保存 {len(result_df)} 条数据')
            except Exception as e:
                print(f'保存文件失败: {str(e)}')
        else:
            print('没有获取到任何数据')
    except Exception as e:
        print(f"程序运行出错: {str(e)}")
    finally:
        print("\n=== 程序运行结束 ===")

if __name__ == '__main__':
    main()