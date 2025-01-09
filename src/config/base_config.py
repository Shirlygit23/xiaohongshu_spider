import os

class BaseConfig:
    def __init__(self):
        # 基础路径配置
        self.BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.DATA_DIR = os.path.join(self.BASE_DIR, 'data')
        self.LOG_DIR = os.path.join(self.BASE_DIR, 'logs')
        self.IMAGE_DIR = os.path.join(self.DATA_DIR, 'images')
        
        # Excel文件配置部分
        self.EXCEL_FILENAME = 'xiaohongshu_鞋靴_关键词和类目_20250107_22点51分 488条.xlsx'
        self.EXCEL_PATH = os.path.join(self.DATA_DIR, self.EXCEL_FILENAME)
        
        # HTTP请求配置
        self.HEADERS = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Connection': 'keep-alive'
        }
        
        # 请求配置
        self.REQUEST_TIMEOUT = 10
        self.RETRY_TIMES = 3
        self.SLEEP_TIME = (2, 4)
        
        # 数据处理配置
        self.REQUIRED_COLUMNS = ['官方笔记地址', '笔记详情', '笔记话题']
        self.BATCH_SIZE = 5  # 每处理多少条保存一次
        
        # 在 BaseConfig 类的 __init__ 方法中，话题分析配置部分应该是：

        # 话题分析配置
        self.TOP_N_TOPICS = 50  # 展示前N个热门话题
        self.TOPIC_HEADERS = ['话题', '出现次数', '总互动量', '平均互动量', '最高互动量', '最低互动量']
        self.TOPIC_COLUMN_WIDTHS = {
            'A': 25,  # 话题列
            'B': 15,  # 出现次数列
            'C': 15,  # 总互动量列
            'D': 15,  # 平均互动量列
            'E': 15,  # 最高互动量列
            'F': 15   # 最低互动量列
        }
        
        # 标题分析配置
        self.TITLE_MIN_LENGTH = 5
        self.TITLE_MAX_LENGTH = 100
        self.TITLE_KEYWORDS_COUNT = 10
        self.TITLE_VIRAL_THRESHOLD = 1000  # 爆款标准：点赞数阈值
        self.TITLE_MIN_LIKES = 100  # 低粉爆款最低点赞数
        self.TITLE_SAVE_BATCH = 10  # 标题分析保存批次

        # 标题分析关键词
        self.TITLE_RELEVANT_KEYWORDS = [
            '穿搭', '搭配', '风格', '合集', '教程', '攻略',
            '推荐', '分享', '必入', '必备', '清单', 
            '种草', '安利', '测评', '一周', '搭着',
            '好看', '显瘦', '百搭', '实穿', '舒服',
            '款式', '单品', '流行', '时尚', '潮流'
        ]

        # 标题模板模式
        self.TITLE_PATTERNS = {
            '紧迫感': [
                r'有救了.*',
                r'赶紧存下来.*',
                r'不看后悔.*',
                r'错过.*',
                r'速收藏.*',
                r'.*趁早.*',
                r'.*抓紧.*',
                r'.*必须收藏.*',
            ],
            '数字清单': [
                r'^\d+[个条招式款].*',
                r'.*\d+个秘密.*',
                r'.*\d+条铁律.*',
                r'.*TOP\d+.*',
                r'.*第\d+名.*',
                r'.*\d+款必入.*',
            ],
            '强烈情感': [
                r'.*真的[绝牛太好].*',
                r'.*绝绝子.*',
                r'.*爱了爱了.*',
                r'.*震惊.*',
                r'.*太太太.*',
                r'.*泪推.*',
                r'.*吹爆.*',
                r'.*惊艳.*',
                r'.*好绝.*',
            ],
            '独家发现': [
                r'.*人不知道的.*',
                r'.*终于找到.*',
                r'.*被忽视的.*',
                r'.*秘密都在这.*',
                r'.*私藏.*',
                r'.*小众.*',
                r'.*神秘.*',
                r'.*发现宝藏.*',
            ],
            '解决痛点': [
                r'.*再也不怕.*',
                r'.*解决困扰.*',
                r'.*不用愁.*',
                r'.*教你解决.*',
                r'.*完美解决.*',
                r'.*一次性解决.*',
                r'.*不用担心.*',
                r'.*轻松搞定.*',
            ],
            '热门趋势': [
                r'.*爆火.*',
                r'.*大火.*',
                r'.*火遍.*',
                r'.*刷屏.*',
                r'.*爆款.*',
                r'.*热门.*',
                r'.*流行.*',
                r'.*出圈.*',
                r'.*火爆全网.*',
            ],
            '性价比': [
                r'.*平价.*',
                r'.*便宜.*',
                r'.*白菜价.*',
                r'.*性价比.*',
                r'.*实惠.*',
                r'.*划算.*',
                r'.*省钱.*',
                r'.*超值.*',
                r'.*百元内.*',
            ],
            '季节热点': [
                r'.*新年.*',
                r'.*圣诞.*',
                r'.*冬日.*',
                r'.*秋冬.*',
                r'.*春季.*',
                r'.*夏天.*',
                r'.*过年.*',
                r'.*节日.*',
                r'.*跨年.*',
            ],
            '专业测评': [
                r'.*测评.*',
                r'.*评测.*',
                r'.*对比.*',
                r'.*分析.*',
                r'.*横评.*',
                r'.*实测.*',
                r'.*深度体验.*',
                r'.*使用感受.*',
            ],
            '穿搭指南': [
                r'.*搭配.*',
                r'.*穿搭.*',
                r'.*穿法.*',
                r'.*风格.*',
                r'.*look.*',
                r'.*outfit.*',
                r'.*style.*',
                r'.*混搭.*',
            ]
        }
        
        # 图片下载配置
        self.MAX_CONCURRENT_DOWNLOADS = 5  # 最大并发下载数
        self.IMAGE_FORMATS = ['.jpg', '.jpeg', '.png', '.gif']  # 支持的图片格式
        self.MAX_IMAGE_SIZE = 10 * 1024 * 1024  # 最大图片大小（10MB）
        self.IMAGE_TIMEOUT = 30  # 图片下载超时时间

        # 图片下载筛选阈值
        self.IMAGE_VIRAL_THRESHOLD = 1000  # 爆文点赞数阈值
        self.IMAGE_LOW_FANS_THRESHOLD = 1000  # 低粉丝数阈值
        self.IMAGE_HIGH_INTERACTION_THRESHOLD = 100  # 高互动量阈值
        
        # 日志配置
        self.LOG_FILENAME = 'processor.log'
        self.LOG_PATH = os.path.join(self.LOG_DIR, self.LOG_FILENAME)
        self.LOG_LEVEL = 'INFO'
        self.LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        
        # 错误处理配置
        self.MAX_RETRIES = 3  # 最大重试次数
        self.RETRY_DELAY = 5  # 重试延迟（秒）
        self.ERROR_SLEEP_TIME = 60  # 发生错误时的等待时间
        
        # 确保必要的目录存在
        self._ensure_directories()
        
        # 验证文件存在
        self._validate_files()
    
    def _ensure_directories(self):
        """确保必要的目录存在"""
        directories = [self.DATA_DIR, self.LOG_DIR, self.IMAGE_DIR]
        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"创建目录: {directory}")
    
    def _validate_files(self):
        """验证必要的文件是否存在"""
        if not os.path.exists(self.EXCEL_PATH):
            print(f"警告: Excel文件不存在: {self.EXCEL_PATH}")
            print(f"请确保文件位于正确的位置: {self.DATA_DIR}")
            print(f"当前支持的文件名: {self.EXCEL_FILENAME}")

# 创建全局配置实例
base_config = BaseConfig()

# 如果这个文件被直接运行，执行测试
if __name__ == '__main__':
    print("基础配置信息:")
    print(f"项目根目录: {base_config.BASE_DIR}")
    print(f"数据目录: {base_config.DATA_DIR}")
    print(f"Excel文件路径: {base_config.EXCEL_PATH}")
    print(f"Excel文件是否存在: {os.path.exists(base_config.EXCEL_PATH)}")