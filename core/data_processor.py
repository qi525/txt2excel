# core/data_processor.py
from typing import Tuple, List, Dict
from config import (
    R18_KEYWORDS, BOY_KEYWORDS, FURRY_KEYWORDS,
    MONOCHROME_GREYSCALE_KEYWORDS, SIMPLE_BACKGROUND_KEYWORDS,
    WORDS_TO_CLEAN_TAGS, SENSITIVE_KEYWORDS_FOR_UNCENSORED
)

def detect_types(line: str) -> str:
    """
    根据预定义关键词检测并返回提示词类型。
    """
    types = []
    lower_line = line.lower()

    if any(word in lower_line for word in R18_KEYWORDS):
        types.append('R18')
    if any(boy_word in lower_line for boy_word in BOY_KEYWORDS):
        types.append('boy')
    if 'no_human' in lower_line:
        types.append('no_human')
    if any(word in lower_line for word in FURRY_KEYWORDS):
        types.append('furry')
    if any(word in lower_line for word in MONOCHROME_GREYSCALE_KEYWORDS):
        types.append('黑白原图')
    if any(word in lower_line for word in SIMPLE_BACKGROUND_KEYWORDS):
        types.append('简单背景')
        
    return ','.join(types)

def clean_tags(line: str) -> Tuple[str, bool]:
    """
    清洗Tag，移除不需要的关键词，并检测是否包含敏感词。
    返回清洗后的Tag字符串和是否包含敏感词的布尔值。
    """
    tags = [tag.strip() for tag in line.strip().split(',')]
    
    cleaned_tags = [
        tag for tag in tags
        if not any(word in tag.lower() for word in WORDS_TO_CLEAN_TAGS)
    ]

    has_sensitive = any(
        any(word in tag.lower() for word in SENSITIVE_KEYWORDS_FOR_UNCENSORED)
        for tag in tags
    )
    if has_sensitive:
        # 确保uncensored只添加一次，并且不与原有tag重复
        if 'uncensored' not in [t.lower() for t in cleaned_tags]:
            cleaned_tags.append('uncensored')

    # 过滤掉空字符串，并用逗号+空格连接
    cleaned_line = ', '.join(filter(None, cleaned_tags))
    return cleaned_line, has_sensitive