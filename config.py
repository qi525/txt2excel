# config.py

# 文件夹名称
HISTORY_FOLDER_NAME = "反推历史记录"
OUTPUT_FOLDER_NAME = "反推记录"
HISTORY_EXCEL_NAME = "scan_history.xlsx"

# 缓存文件夹路径 (请根据您的系统自行修改)
# 注意：使用 r'' 前缀创建原始字符串，避免反斜杠转义问题
CACHE_FOLDER_PATH_STR = r'C:\个人数据\pythonCode\反推图片信息\cache'

# 定义R18相关词汇列表
R18_KEYWORDS = [
    'sex', 'nude', 'pussy', 'penis', 'cum', 'nipples', 'vaginal', 'cum_in_pussy',
    'oral', 'rape', 'fellatio', 'facial', 'anus', 'anal', 'ejaculation',
    'gangbang', 'testicles', 'multiple_penises', 'erection', 'handjob',
    'cumdrip', 'pubic_hair', 'pussy_juice', 'bukkake', 'clitoris',
    'female_ejaculation', 'threesome', 'doggystyle', 'sex_from_behind',
    'cum_on_breasts', 'double_penetration', 'anal_object_insertion',
    'cunnilingus', 'triple_penetration', 'paizuri', 'vaginal_object_insertion',
    'imminent_rape', 'impregnation', 'prone_bone', 'reverse_cowgirl_position',
    'cum_inflation', 'milking_machine', 'cumdump', 'anal_hair', 'futanari',
    'glory_hole', 'penis_on_face', 'licking_penis', 'breast_sucking',
    'breast_squeeze', 'straddling'
]

# 定义需要清洗掉的Tag关键词列表
WORDS_TO_CLEAN_TAGS = [
    'censor', 'monochrome', 'greyscale', 'furry', 'animal_focus', 'no_human', 'background'
]

# 定义敏感词列表，用于标记 'uncensored'
SENSITIVE_KEYWORDS_FOR_UNCENSORED = [
    'censor', 'nipple', 'pussy', 'penis', 'hetero', 'sex', 'anus'
]

# 定义boy类型词汇
BOY_KEYWORDS = ['1boy', '2boys', 'multiple_boys']

# 定义furry类型词汇
FURRY_KEYWORDS = ['furry', 'animal_focus']

# 定义黑白原图类型词汇
MONOCHROME_GREYSCALE_KEYWORDS = ['monochrome', 'greyscale']

# 定义简单背景类型词汇
SIMPLE_BACKGROUND_KEYWORDS = ['background']