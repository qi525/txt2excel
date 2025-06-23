from typing import Tuple, List, Dict, Set

# --- Global Configuration for Tag Processing ---
# 新增全局配置：定义各种类型检测的规则
# 原理：将类型检测的关键词和对应的类型名称集中管理，提高可维护性和扩展性。
# 这样，添加新的类型检测只需修改此字典，而无需修改detect_types函数内部逻辑。
TAG_DETECTION_RULES: Dict[str, List[str]] = {
    'R18': [
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
    ],
    'boy': ['1boy', '2boys', 'multiple_boys'],
    'no_human': ['no_human'],
    'furry': ['furry', 'animal_focus'],
    '黑白原图': ['monochrome', 'greyscale'],
    '简单背景': ['background']
}

# 新增全局配置：统一管理需要进行敏感词检查的词汇集合
# 原理：将分散的敏感词汇集中管理，确保clean_tags函数中敏感词判断的准确性和一致性。
# 主要改动点：将R18词汇和额外的通用敏感词合并到一个集合中。
SENSITIVE_WORDS_FOR_CHECK: Set[str] = set(TAG_DETECTION_RULES['R18'] + [
    'nipple', 'pussy', 'penis', 'hetero', 'sex', 'anus', 'naked', 'explicit'
])

# --- Data Processor (RESTORED FROM V4.0) ---
def detect_types(line: str, cleaned: str) -> str:
    """
    根据文本内容推断提示词类型。还原自 V4.0 版本。
    优化原理：通过遍历预定义的TAG_DETECTION_RULES字典，动态地检测各种类型。
             这使得添加或修改类型检测规则时，只需修改TAG_DETECTION_RULES字典，
             而无需修改函数内部逻辑，极大提高了代码的通用性和可维护性。
    Args:
        line (str): 原始的txt文件内容。
        cleaned (str): 清洗后的txt文件内容。 (在此函数中未使用cleaned，但保留参数签名以兼容原有接口)
    Returns:
        str: 识别到的提示词类型，用逗号分隔。
    """
    types: List[str] = []
    lower_line: str = line.lower()

    # 主要改动点：迭代TAG_DETECTION_RULES，动态检测类型
    for type_name, keywords in TAG_DETECTION_RULES.items():
        if any(word in lower_line for word in keywords):
            types.append(type_name)

    # 如果没有检测到任何类型，返回 "N/A"
    if not types:
        return "N/A"
    return ','.join(types)

def clean_tags(line: str) -> Tuple[str, bool]:
    """
    清洗标签字符串。修改了对'censor'词的清理逻辑和'uncensored'的添加逻辑。
    优化原理：统一管理需要清洗的关键词和敏感词，减少重复定义，提高代码一致性。
    Args:
        line (str): 原始的标签字符串。
    Returns:
        Tuple[str, bool]: 清洗后的字符串和是否含有敏感词的布尔值。
    """
    tags: List[str] = [tag.strip() for tag in line.strip().split(',')]

    # 主要改动点：根据TAG_DETECTION_RULES动态生成需要清洗的关键词
    # 不包含 'uncensored'，因为uncensored不是一个需要被清洗的通用tag，而是标记
    # 同时排除R18类型词汇，因为这些词在clean_tags中不应被直接清洗，而是用于判断敏感性。
    words_to_clean: Set[str] = set()
    for type_name, keywords in TAG_DETECTION_RULES.items():
        if type_name != 'R18': # R18关键词用于敏感词判断，而不是直接清洗
            words_to_clean.update(keywords)
    # 额外添加'censor', 'censored'，确保这些通用清理词也被包含
    words_to_clean.update(['censor', 'censored'])


    # 主要改动点：使用统一的SENSITIVE_WORDS_FOR_CHECK集合判断是否含有敏感词
    # 检查是否含有敏感词 (基于原始标签列表，因为这些词不应该被清洗掉，而是用于标记)
    has_sensitive: bool = False
    for tag in tags:
        if any(word in tag.lower() for word in SENSITIVE_WORDS_FOR_CHECK):
            has_sensitive = True
            break
    
    cleaned_tags: List[str] = []
    for tag in tags:
        lower_tag = tag.lower()
        
        # 如果是 'uncensored'，直接添加，不进行清洗
        if lower_tag == 'uncensored':
            cleaned_tags.append(tag)
            continue
            
        # 只有当tag不包含任何words_to_clean中的词时才保留
        # 确保完整的tag不包含清理词，而不是部分包含
        if not any(word in lower_tag for word in words_to_clean):
            cleaned_tags.append(tag)

    # 如果检测到敏感词，则添加 'uncensored' 标记
    # 确保只添加一次
    if has_sensitive and 'uncensored' not in [t.lower() for t in cleaned_tags]:
        cleaned_tags.append('uncensored')
    
    # 过滤掉空字符串，然后用逗号和空格连接
    cleaned_line: str = ', '.join([tag for tag in cleaned_tags if tag])
    return cleaned_line, has_sensitive