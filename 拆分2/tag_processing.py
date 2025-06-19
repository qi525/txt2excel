import re
from typing import Tuple
# --- Data Processor (RESTORED FROM V4.0) ---
def detect_types(line: str, cleaned: str) -> str:
    """
    根据文本内容推断提示词类型。还原自 V4.0 版本。
    Args:
        line (str): 原始的txt文件内容。
        cleaned (str): 清洗后的txt文件内容。
    Returns:
        str: 识别到的提示词类型，用逗号分隔。
    """
    types = []
    lower_line = line.lower()
    # R18相关词汇
    if any(word in lower_line for word in [
        'sex', 'nude', 'pussy', 'penis', 'cum', 'nipples', 'vaginal',
        'cum_in_pussy', 'oral', 'rape', 'fellatio', 'facial', 'anus',
        'anal', 'ejaculation', 'gangbang', 'testicles', 'multiple_penises',
        'erection', 'handjob', 'cumdrip', 'pubic_hair', 'pussy_juice',
        'bukkake', 'clitoris', 'female_ejaculation', 'threesome',
        'doggystyle', 'sex_from_behind', 'cum_on_breasts', 'double_penetration',
        'anal_object_insertion', 'cunnilingus', 'triple_penetration',
        'paizuri', 'vaginal_object_insertion', 'imminent_rape', 'impregnation',
        'prone_bone', 'reverse_cowgirl_position', 'cum_inflation',
        'milking_machine', 'cumdump', 'anal_hair', 'futanari', 'glory_hole',
        'penis_on_face', 'licking_penis', 'breast_sucking', 'breast_squeeze', 'straddling'
    ]):
        types.append('R18')
    # boy类型
    if any(boy_word in lower_line for boy_word in ['1boy', '2boys', 'multiple_boys']):
        types.append('boy')
    # no_human类型
    if 'no_human' in lower_line:
        types.append('no_human')
    # furry 类型
    if any(word in lower_line for word in ['furry', 'animal_focus']):
        types.append('furry')
    # monochrome和greyscale类型
    if any(word in lower_line for word in ['monochrome', 'greyscale']):
        types.append('黑白原图')
    # 新增功能：检测"background"相关词汇并标记为“简单背景”类型
    if 'background' in lower_line:
        types.append('简单背景')
    
    # 如果没有检测到任何类型，返回 "N/A"
    if not types:
        return "N/A"
    return ','.join(types)

def clean_tags(line: str) -> Tuple[str, bool]:
    """
    清洗标签字符串。修改了对'censor'词的清理逻辑和'uncensored'的添加逻辑。
    Args:
        line (str): 原始的标签字符串。
    Returns:
        Tuple[str, bool]: 清洗后的字符串和是否含有敏感词的布尔值。
    """
    tags = [tag.strip() for tag in line.strip().split(',')]
    
    # 定义需要清洗掉的关键词，不包含 'uncensored'
    words_to_clean = ['censor', 'censored', 'monochrome', 'greyscale', 'furry', 'animal_focus', 'no_human', 'background']
    
    # 检查是否含有敏感词 (基于原始标签列表，因为这些词不应该被清洗掉，而是用于标记)
    has_sensitive = any(
        any(word in tag.lower() for word in [
            'nipple', 'pussy', 'penis', 'hetero', 'sex', 'anus', 'naked', 'explicit' # 增加一些常见的敏感词
        ])
        for tag in tags
    )
    
    # 过滤掉需要清洗的关键词
    cleaned_tags = []
    for tag in tags:
        # 如果是 'uncensored'，直接添加，不进行清洗
        if tag.lower() == 'uncensored':
            cleaned_tags.append(tag)
            continue
            
        # 只有当tag不包含任何words_to_clean中的词时才保留
        if not any(word in tag.lower() for word in words_to_clean):
            cleaned_tags.append(tag)

    # 如果检测到敏感词，则添加 'uncensored' 标记 
    # 确保只添加一次
    if has_sensitive and 'uncensored' not in [t.lower() for t in cleaned_tags]:
        cleaned_tags.append('uncensored')
    
    # 过滤掉空字符串，然后用逗号和空格连接
    cleaned_line = ', '.join([tag for tag in cleaned_tags if tag])
    return cleaned_line, has_sensitive

