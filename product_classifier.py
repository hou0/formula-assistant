"""
产品分类判定与毒理学路径选择 (PRD 5.8)
======================================
根据产品类别自动选择评估路径和毒理学终点
"""

from dataclasses import dataclass, field


# ── 产品分类定义 ──

@dataclass
class ProductCategory:
    """产品类别定义"""
    code: str                              # 内部代码
    name_zh: str                           # 中文名称
    category_group: str                    # 大类: skincare/haircare/color/cosmetic/mouth/body
    product_type: str                      # 普通/特殊
    is_rinsed: bool = True                 # 淋洗型
    is_professional: bool = False          # 专业用
    application_site: str = ''             # 典型使用部位
    description: str = ''                  # 描述


# ── 预定义产品类别 ──

_PRODUCT_CATEGORIES: dict[str, ProductCategory] = {
    # 护肤类
    'toner': ProductCategory('toner', '化妆水', 'skincare', '普通', False, application_site='面部'),
    'lotion': ProductCategory('lotion', '乳液', 'skincare', '普通', False, application_site='面部'),
    'cream': ProductCategory('cream', '面霜/日霜', 'skincare', '普通', False, application_site='面部'),
    'essence': ProductCategory('essence', '精华液', 'skincare', '普通', False, application_site='面部'),
    'cleanser': ProductCategory('cleanser', '洗面奶', 'skincare', '普通', True, application_site='面部'),
    'makeup_remover': ProductCategory('makeup_remover', '卸妆液/油', 'skincare', '普通', True, application_site='面部'),
    'mask': ProductCategory('mask', '面膜', 'skincare', '普通', False, application_site='面部'),
    'sunscreen': ProductCategory('sunscreen', '防晒霜', 'skincare', '特殊', False, application_site='全身皮肤'),
    'eye_cream': ProductCategory('eye_cream', '眼霜/眼胶', 'skincare', '普通', False, application_site='眼部'),
    'hand_cream': ProductCategory('hand_cream', '护手霜', 'skincare', '普通', False, application_site='手部/足部'),
    'body_lotion': ProductCategory('body_lotion', '体乳/体霜', 'skincare', '普通', False, application_site='全身皮肤'),
    'body_wash': ProductCategory('body_wash', '沐浴露', 'skincare', '普通', True, application_site='全身皮肤'),

    # 发用类
    'shampoo': ProductCategory('shampoo', '洗发水', 'haircare', '普通', True, application_site='头发'),
    'conditioner': ProductCategory('conditioner', '护发素', 'haircare', '普通', True, application_site='头发'),
    'hair_mask': ProductCategory('hair_mask', '发膜', 'haircare', '普通', False, application_site='头发'),
    'hair_oil': ProductCategory('hair_oil', '护发精油', 'haircare', '普通', False, application_site='头发'),
    'hair_dye_oxidative': ProductCategory('hair_dye_oxidative', '染发剂（氧化型）', 'haircare', '特殊', False, application_site='头发'),
    'hair_dye_nonoxidative': ProductCategory('hair_dye_nonoxidative', '染发剂（非氧化型）', 'haircare', '特殊', False, application_site='头发'),

    # 彩妆类
    'foundation': ProductCategory('foundation', '粉底粉', 'color', '普通', False, application_site='面部'),
    'loose_powder': ProductCategory('loose_powder', '散粉/蜜粉', 'color', '普通', False, application_site='面部'),
    'lipstick': ProductCategory('lipstick', '口红', 'color', '普通', False, application_site='嘴部'),
    'lip_gloss': ProductCategory('lip_gloss', '唇彩/唇釉', 'color', '普通', False, application_site='嘴部'),
    'eye_shadow': ProductCategory('eye_shadow', '眼影', 'color', '普通', False, application_site='眼部'),
    'eyeliner': ProductCategory('eyeliner', '眼线笔/液', 'color', '普通', False, application_site='眼部'),
    'mascara': ProductCategory('mascara', '睫毛膏', 'color', '普通', False, application_site='眼部'),
    'eyebrow': ProductCategory('eyebrow', '眉笔/眉粉', 'color', '普通', False, application_site='眼部'),

    # 口腔类
    'toothpaste_adult': ProductCategory('toothpaste_adult', '牙膏（成人）', 'mouth', '普通', True, application_site='口腔'),
    'toothpaste_child': ProductCategory('toothpaste_child', '牙膏（儿童）', 'mouth', '普通', True, application_site='口腔'),

    # 身体护理
    'soap': ProductCategory('soap', '香皂', 'body', '普通', True, application_site='全身皮肤'),
    'deodorant': ProductCategory('deodorant', '止汗露/香体露', 'body', '普通', False, application_site='全身皮肤'),
}


def classify_product(category_text: str) -> ProductCategory:
    """Classify product by Chinese name, return best match or default."""
    category_text = category_text.strip().lower()

    # Direct match
    for code, cat in _PRODUCT_CATEGORIES.items():
        if cat.name_zh.lower() == category_text:
            return cat

    # Fuzzy match
    for code, cat in _PRODUCT_CATEGORIES.items():
        if any(char in category_text for char in cat.name_zh):
            return cat

    # Default
    return ProductCategory('unknown', category_text, 'skincare', '普通', False, application_site='')


def get_categories_by_group(group: str) -> list[ProductCategory]:
    """Get all categories in a group."""
    return [c for c in _PRODUCT_CATEGORIES.values() if c.category_group == group]


def guess_application_site(category_name: str) -> str:
    """Guess application site from product category name."""
    cat = classify_product(category_name)
    if cat.application_site:
        return cat.application_site

    site_keywords = {
        '面': '面部',
        '眼': '眼部',
        '唇': '嘴部',
        '口': '口腔',
        '手': '手部/足部',
        '足': '手部/足部',
        '体': '全身皮肤',
        '身': '全身皮肤',
        '发': '头发',
        '眉': '眼部',
    }
    for kw, site in site_keywords.items():
        if kw in category_name:
            return site
    return '面部'


# ── 毒理学路径选择 ──

@dataclass
class ToxicologyPath:
    """毒理学评估路径"""
    name: str
    required_endpoints: list[str] = field(default_factory=list)
    description: str = ''
    priority: int = 0


_TOX_PATHS: dict[str, ToxicologyPath] = {
    'full': ToxicologyPath(
        '完整毒理学评估',
        ['急性毒性', '刺激性/腐蚀性', '皮肤致敏', '光毒性', '光致敏性',
         '重复剂量毒性', '遗传毒性', '生殖发育毒性', '致癌性', '毒代动力学'],
        '适用于新原料或缺少参考数据的情况',
        priority=1,
    ),
    'threshold': ToxicologyPath(
        '阈值法评估',
        ['急性毒性', '刺激性/腐蚀性', '皮肤致敏', '遗传毒性'],
        '适用于低暴露量且已有参考数据的原料',
        priority=2,
    ),
    'read_across': ToxicologyPath(
        '交叉参照评估',
        ['急性毒性', '刺激性/腐蚀性', '皮肤致敏'],
        '适用于结构类似物数据充分的情况',
        priority=3,
    ),
    'historical': ToxicologyPath(
        '历史使用评估',
        ['刺激性/腐蚀性', '皮肤致敏'],
        '适用于有长期安全使用历史的原料（已上市目录）',
        priority=4,
    ),
}


def select_toxicology_path(
    contains_novel_ingredient: bool = False,
    has_usage_data: bool = True,
    has_tox_data: bool = False,
    is_high_exposure: bool = False,
) -> ToxicologyPath:
    """Select the appropriate toxicology assessment path based on product characteristics.

    Args:
        contains_novel_ingredient: Whether the formula contains a novel ingredient
        has_usage_data: Whether usage reference data is available
        has_tox_data: Whether toxicology data is available
        is_high_exposure: Whether the product has high exposure

    Returns:
        Appropriate ToxicologyPath
    """
    if contains_novel_ingredient:
        return _TOX_PATHS['full']
    if is_high_exposure or not has_usage_data:
        if has_tox_data:
            return _TOX_PATHS['threshold']
        return _TOX_PATHS['full']
    if has_tox_data:
        return _TOX_PATHS['read_across']
    return _TOX_PATHS['historical']


def get_required_endpoints_for_product(category: ProductCategory) -> list[str]:
    """Get required toxicological endpoints for a product category.

    Based on 《化妆品安全评估技术导则》 endpoint requirements.
    """
    base_endpoints = ['急性毒性', '刺激性/腐蚀性', '皮肤致敏']

    if not category.is_rinsed:
        base_endpoints.append('光毒性')
        base_endpoints.append('光致敏性')

    if category.product_type == '特殊':
        base_endpoints.extend(['重复剂量毒性', '遗传毒性', '生殖发育毒性'])

    if category.category_group in ('mouth',):
        base_endpoints.append('口腔黏膜刺激')

    if category.application_site == '眼部' or '眼' in (category.application_site or ''):
        base_endpoints.append('眼刺激性')

    return base_endpoints


def suggest_assessment_approach(category_text: str) -> dict:
    """Suggest assessment approach based on product category.

    Returns a dict with suggested path, endpoints, and notes.
    """
    cat = classify_product(category_text)
    endpoints = get_required_endpoints_for_product(cat)
    path = select_toxicology_path(
        contains_novel_ingredient=False,
        has_usage_data=True,
        has_tox_data=False,
        is_high_exposure=not cat.is_rinsed,
    )

    return {
        'product_category': cat,
        'suggested_path': path,
        'required_endpoints': endpoints,
        'is_rinsed': cat.is_rinsed,
        'is_special': cat.product_type == '特殊',
        'notes': (
            f"该产品为{cat.product_type}用途{cat.category_group}类产品，"
            f"{'淋洗型' if cat.is_rinsed else '驻留型'}。"
            f"建议采用{path.name}进行毒理学评估。"
        ),
    }
