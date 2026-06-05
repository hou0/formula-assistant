"""
多国市场配置模块 (PRD 5.6)
==========================
管理不同国家/地区的法规要求和安全评估标准
"""

from dataclasses import dataclass


@dataclass
class MarketConfig:
    """市场配置"""
    code: str
    name_zh: str
    name_en: str
    default_body_weight_kg: float = 60.0
    default_language: str = 'zh-CN'
    regulatory_framework: str = ''
    moe_threshold: float = 100.0          # MOE安全阈值
    requires_moe: bool = True              # 是否需要MOE计算
    requires_product_notification: bool = True
    requires_safety_report: bool = True
    requires_stability_test: bool = True
    requires_microbiology_test: bool = True
    requires_heavy_metal_test: bool = True
    label_language: str = 'zh-CN'
    notes: str = ''


# ── 预定义市场配置 ──

_MARKETS: dict[str, MarketConfig] = {
    'CN': MarketConfig(
        code='CN',
        name_zh='中国',
        name_en='China',
        default_body_weight_kg=60.0,
        default_language='zh-CN',
        regulatory_framework='《化妆品监督管理条例》《化妆品安全技术规范》（2015年版）',
        moe_threshold=100.0,
        requires_moe=True,
        requires_product_notification=True,
        requires_safety_report=True,
        requires_stability_test=True,
        requires_microbiology_test=True,
        requires_heavy_metal_test=True,
        label_language='zh-CN',
        notes='国产普通化妆品备案制，特殊化妆品注册制',
    ),
    'EU': MarketConfig(
        code='EU',
        name_zh='欧盟',
        name_en='European Union',
        default_body_weight_kg=60.0,
        default_language='en',
        regulatory_framework='EU Cosmetics Regulation (EC) No 1223/2009',
        moe_threshold=100.0,
        requires_moe=True,
        requires_product_notification=True,
        requires_safety_report=True,
        requires_stability_test=True,
        requires_microbiology_test=True,
        requires_heavy_metal_test=True,
        label_language='local',
        notes='CPNP通报，Safety Report mandatory',
    ),
    'US': MarketConfig(
        code='US',
        name_zh='美国',
        name_en='United States',
        default_body_weight_kg=60.0,
        default_language='en',
        regulatory_framework='FD&C Act / MoCRA',
        moe_threshold=100.0,
        requires_moe=True,
        requires_product_notification=False,
        requires_safety_report=False,
        requires_stability_test=True,
        requires_microbiology_test=True,
        requires_heavy_metal_test=False,
        label_language='en',
        notes='MoCRA逐步实施安全报告要求，FDA注册',
    ),
    'JP': MarketConfig(
        code='JP',
        name_zh='日本',
        name_en='Japan',
        default_body_weight_kg=50.0,
        default_language='ja',
        regulatory_framework='医薬品医療機器等法',
        moe_threshold=100.0,
        requires_moe=True,
        requires_product_notification=True,
        requires_safety_report=True,
        requires_stability_test=True,
        requires_microbiology_test=True,
        requires_heavy_metal_test=True,
        label_language='ja',
        notes='日本体重标准50kg，许可制+备案制',
    ),
    'KR': MarketConfig(
        code='KR',
        name_zh='韩国',
        name_en='South Korea',
        default_body_weight_kg=55.0,
        default_language='ko',
        regulatory_framework='化妆品法（화장품법）',
        moe_threshold=100.0,
        requires_moe=True,
        requires_product_notification=True,
        requires_safety_report=True,
        requires_stability_test=True,
        requires_microbiology_test=True,
        requires_heavy_metal_test=True,
        label_language='ko',
        notes='韩国体重标准55kg',
    ),
    'ASEAN': MarketConfig(
        code='ASEAN',
        name_zh='东盟',
        name_en='ASEAN',
        default_body_weight_kg=60.0,
        default_language='en',
        regulatory_framework='ASEAN Cosmetic Directive',
        moe_threshold=100.0,
        requires_moe=True,
        requires_product_notification=True,
        requires_safety_report=True,
        requires_stability_test=False,
        requires_microbiology_test=True,
        requires_heavy_metal_test=True,
        label_language='en',
        notes='东盟统一化妆品法规',
    ),
}


def get_market_config(code: str) -> MarketConfig:
    """Get market configuration by code."""
    if code in _MARKETS:
        return _MARKETS[code]
    return _MARKETS['CN']  # Default to China


def get_all_markets() -> dict[str, MarketConfig]:
    """Get all available market configurations."""
    return dict(_MARKETS)


def get_market_names() -> list[dict]:
    """Get market list for UI display."""
    return [
        {'code': m.code, 'name_zh': m.name_zh, 'name_en': m.name_en}
        for m in _MARKETS.values()
    ]


def get_body_weight_for_market(market_code: str) -> float:
    """Get default body weight for a market."""
    return get_market_config(market_code).default_body_weight_kg


def get_required_tests(market_code: str) -> dict:
    """Get required tests for a market."""
    config = get_market_config(market_code)
    return {
        'stability': config.requires_stability_test,
        'microbiology': config.requires_microbiology_test,
        'heavy_metal': config.requires_heavy_metal_test,
    }
