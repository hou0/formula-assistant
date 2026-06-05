"""
安全评估引擎 (PRD Section 5.3 / 5.5)
====================================
四层合规检查 + 暴露量评估 (SED/MOE)
"""

import json
from typing import Any
import os

import database as db


# ── 配置加载 ──

def _load_config():
    """Load safety assessment configuration from JSON file."""
    config_path = os.path.join(os.path.dirname(__file__), 'userdata', 'safety_config.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Failed to load safety_config.json: {e}")
        return _get_default_config()

def _get_default_config():
    """Default configuration if JSON file fails to load."""
    return {
        'constants': {'default_body_weight_kg': 60.0},
        'regulatory_templates': {
            'banned': {
                'table1': '《化妆品安全技术规范》（2015年版）化妆品禁用组分（表1）',
                'table2': '《化妆品安全技术规范》（2015年版）化妆品禁用植（动）物组分（表2）'
            },
            'restricted': {'table3': '《化妆品安全技术规范》（2015年版）化妆品限用组分（表3）'},
            'allowed': {
                'table4': '《化妆品安全技术规范》（2015年版）化妆品准用防腐剂（表4）',
                'table5': '《化妆品安全技术规范》（2015年版）化妆品准用防晒剂（表5）',
                'table6': '《化妆品安全技术规范》（2015年版）化妆品准用着色剂（表6）',
                'table7': '《化妆品安全技术规范》（2015年版）化妆品准用染发剂（表7）'
            }
        }
    }

# Load configuration
_CONFIG = _load_config()
DEFAULT_BW = _CONFIG['constants'].get('default_body_weight_kg', 60.0)
REGULATORY_TEMPLATES = _CONFIG.get('regulatory_templates', {})


# ── 风险物质识别规则（从JSON文件加载） ──

def _load_risk_substances():
    """Load risk substance rules from JSON file."""
    json_path = os.path.join(os.path.dirname(__file__), 'userdata', 'risk_substances.json')
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get('general_risk_substances', []), data.get('child_risk_substances', [])
    except Exception as e:
        print(f"Failed to load risk_substances.json: {e}")
        return [], []

# Load rules from JSON file
RISK_SUBSTANCE_RULES, CHILD_RISK_SUBSTANCE_RULES = _load_risk_substances()


def identify_risk_substances(ingredient_name: str, is_child_product: bool = False) -> list:
    """Identify potential risk substances based on ingredient name and product type.
    
    Based on 《化妆品风险物质识别与评估技术指导原则》
    
    Args:
        ingredient_name: Chinese name of the ingredient
        is_child_product: whether this is a children's product
    
    Returns:
        List of risk substance names that may be present
    """
    identified = []
    
    # Check general risk substance rules
    for rule in RISK_SUBSTANCE_RULES:
        if rule['trigger_all']:
            # Always include this risk substance
            identified.append(rule['name'])
        else:
            # Check if any trigger keyword is in the ingredient name
            for keyword in rule['trigger_keywords']:
                if keyword.lower() in ingredient_name.lower():
                    identified.append(rule['name'])
                    break
    
    # Check children-specific risk substances
    if is_child_product:
        for rule in CHILD_RISK_SUBSTANCE_RULES:
            if rule['trigger_all']:
                identified.append(rule['name'])
            else:
                for keyword in rule['trigger_keywords']:
                    if keyword.lower() in ingredient_name.lower():
                        identified.append(rule['name'])
                        break
    
    return list(set(identified))  # Remove duplicates


def _parse_percent(val: Any) -> float | None:
    """Parse a percentage value from text (e.g. '0.25%', '2.5', '3 %')."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    val = str(val).strip().replace('%', '').replace(' ', '')
    try:
        return float(val)
    except ValueError:
        return None


# ═══════════════════════════════════════════════
# Layer 1: 禁用组分检查
# ═══════════════════════════════════════════════

def check_banned(name_zh: str, name_inci: str = '', cas_no: str = '') -> dict:
    """Check if ingredient is banned.

    Priority matching: CAS > INCI > Chinese name (exact match > partial match)

    Returns:
        {'banned': True/False, 'matches': [...], 'summary': str}
    """
    hits = db.check_banned(name_zh, name_inci, cas_no)

    if not hits:
        return {
            'banned': False,
            'matches': [],
            'summary': '未在禁用列表中',
        }

    # Generate detailed summary for banned ingredients
    match_details = []
    banned_templates = REGULATORY_TEMPLATES.get('banned', {})
    
    for h in hits:
        table_type = h.get('table_type', '')
        seq = h.get('seq', '')

        if table_type == 'botanical':
            detail = banned_templates.get('table2', "《化妆品安全技术规范》（2015年版）化妆品禁用植（动）物组分（表2）")
        else:
            detail = banned_templates.get('table1', "《化妆品安全技术规范》（2015年版）化妆品禁用组分（表1）")

        if seq:
            detail += f"序号{seq}"

        match_details.append(detail)

    return {
        'banned': True,
        'matches': hits,
        'summary': '; '.join(match_details),
    }


# ═══════════════════════════════════════════════
# Layer 2: 限用组分检查
# ═══════════════════════════════════════════════

def check_restricted(name_zh: str, name_inci: str = '',
                     concentration: float | None = None) -> dict:
    """Check if ingredient is restricted, and validate concentration.

    Returns:
        {'restricted': True/False, 'matches': [...], 'summary': str,
         'concentration_ok': True/False/None}
    """
    hits = db.check_restricted(name_zh, name_inci)
    if not hits:
        return {
            'restricted': False,
            'matches': [],
            'summary': '未在列表中',
            'concentration_ok': None,
        }

    concentration_ok = True
    violations = []
    match_details = []
    restricted_templates = REGULATORY_TEMPLATES.get('restricted', {})

    for h in hits:
        seq = h.get('seq', '')
        _max_conc = h.get('max_concentration', '')
        _scope = h.get('scope_of_use', '')
        _restrictions = h.get('restrictions', '')
        _label = h.get('label_requirements', '')
        name_zh_val = h.get('name_zh', '')

        detail = restricted_templates.get('table3', "《化妆品安全技术规范》（2015年版）化妆品限用组分（表3）")
        if seq:
            detail += f"序号{seq}"
        if _max_conc:
            detail += f"；最大浓度: {_max_conc}"
        if _scope:
            detail += f"；使用范围: {_scope}"
        if _restrictions:
            detail += f"；限制: {_restrictions}"
        if _label:
            detail += f"；标签要求: {_label}"

        match_details.append(detail)

        max_conc_val = _parse_percent(h.get('max_concentration'))
        if max_conc_val is not None and concentration is not None:
            if concentration > max_conc_val:
                concentration_ok = False
                violations.append(
                    f"{name_zh_val}: 添加量 {concentration}% > 限值 {max_conc_val}%"
                )

    return {
        'restricted': True,
        'matches': hits,
        'concentration_ok': concentration_ok,
        'violations': violations,
        'summary': '; '.join(match_details) + ('；浓度超标!' if not concentration_ok and violations else ''),
    }


# ═══════════════════════════════════════════════
# Layer 3: 准用组分检查
# ═══════════════════════════════════════════════

def _check_allowed_generic(name_zh: str, name_inci: str,
                           check_func, table_label: str,
                           table_number: int,
                           concentration: float | None = None) -> dict:
    """Generic check for allowed-lists (preservatives, sunscreens, etc.).

    table_number: the table number in the Cosmetic Safety Technical Specification
    """
    hits = check_func(name_zh, name_inci)
    
    allowed_templates = REGULATORY_TEMPLATES.get('allowed', {})
    table_key = f'table{table_number}'
    default_template = f"《化妆品安全技术规范》（2015年版）化妆品准用{table_label}（表{table_number}）"
    template = allowed_templates.get(table_key, default_template)
    
    if not hits:
        return {
            'in_allowed_list': False,
            'matches': [],
            'summary': f'未在{template}列表中',
            'concentration_ok': None,
        }

    concentration_ok = True
    violations = []
    match_details = []

    for h in hits:
        seq = h.get('seq', '')
        name_val = h.get('name_zh', '')
        _max_conc = h.get('max_concentration', '')
        _scope_and_conditions = h.get('scope_and_conditions', '')
        _restrictions = h.get('restrictions', '')
        _label_requirements = h.get('label_requirements', '')

        detail = template
        if seq:
            detail += f"序号{seq}"

        # Add hair dye specific details
        max_conc_oxidative = h.get('max_conc_oxidative', '')
        max_conc_non_oxidative = h.get('max_conc_non_oxidative', '')
        if max_conc_oxidative:
            detail += f"；氧化型染发: {max_conc_oxidative}"
        if max_conc_non_oxidative:
            detail += f"；非氧化型染发: {max_conc_non_oxidative}"

        if _max_conc:
            detail += f"；最大浓度: {_max_conc}"
        if _scope_and_conditions:
            detail += f"；使用范围和条件: {_scope_and_conditions}"
        if _restrictions:
            detail += f"；限制: {_restrictions}"
        if _label_requirements:
            detail += f"；标签要求: {_label_requirements}"

        match_details.append(detail)

        # Check concentration limits
        max_conc_val = _parse_percent(h.get('max_concentration'))
        if max_conc_val is not None and concentration is not None:
            if concentration > max_conc_val:
                concentration_ok = False
                violations.append(
                    f"{name_val}: 添加量 {concentration}% > 限值 {max_conc_val}%"
                )

    # Also check hair dye specific concentration fields
    for h in hits:
        if concentration is not None:
            max_conc_oxidative_val = _parse_percent(h.get('max_conc_oxidative'))
            max_conc_non_oxidative_val = _parse_percent(h.get('max_conc_non_oxidative'))

            if max_conc_oxidative_val is not None:
                if concentration > max_conc_oxidative_val:
                    concentration_ok = False
                    violations.append(
                        f"{h.get('name_zh', '')}: 氧化性染发添加量 {concentration}% > 限值 {max_conc_oxidative_val}%"
                    )

            if max_conc_non_oxidative_val is not None:
                if concentration > max_conc_non_oxidative_val:
                    concentration_ok = False
                    violations.append(
                        f"{h.get('name_zh', '')}: 非氧化性染发添加量 {concentration}% > 限值 {max_conc_non_oxidative_val}%"
                    )

    return {
        'in_allowed_list': True,
        'matches': hits,
        'concentration_ok': concentration_ok,
        'violations': violations,
        'summary': '; '.join(match_details) + ('；浓度超标!' if not concentration_ok and violations else ''),
    }


def check_allowed_preservative(name_zh: str, name_inci: str = '',
                               concentration: float | None = None) -> dict:
    return _check_allowed_generic(name_zh, name_inci,
                                  db.check_allowed_preservative, '防腐剂', 4, concentration)


def check_allowed_sunscreen(name_zh: str, name_inci: str = '',
                            concentration: float | None = None) -> dict:
    return _check_allowed_generic(name_zh, name_inci,
                                  db.check_allowed_sunscreen, '防晒剂', 5, concentration)


def check_allowed_colorant(name_zh: str = '', name_inci: str = '') -> dict:
    """Check if colorant is in allowed list (表6).

    Fields in database: seq, name_zh, color_index, ci_generic_name,
    color, scope_1~scope_4, restrictions
    """
    hits = db.check_allowed_colorant(name_zh)
    
    allowed_templates = REGULATORY_TEMPLATES.get('allowed', {})
    template = allowed_templates.get('table6', "《化妆品安全技术规范》（2015年版）化妆品准用着色剂（表6）")
    
    if not hits:
        return {
            'in_allowed_list': False,
            'matches': [],
            'concentration_ok': None,
            'summary': f'未在{template}列表中',
        }

    match_details = []
    for h in hits:
        seq = h.get('seq', '')
        _restrictions = h.get('restrictions', '')

        detail = template
        if seq:
            detail += f"序号{seq}"

        # Add scope info (scope_1~scope_4 are boolean values)
        scopes = []
        if h.get('scope_1'):
            scopes.append('I类')
        if h.get('scope_2'):
            scopes.append('II类')
        if h.get('scope_3'):
            scopes.append('III类')
        if h.get('scope_4'):
            scopes.append('IV类')
        if scopes:
            detail += f"；适用范围: {'、'.join(scopes)}"

        if _restrictions:
            detail += f"；限制: {_restrictions}"

        match_details.append(detail)

    return {
        'in_allowed_list': True,
        'matches': hits,
        'concentration_ok': None,
        'summary': '; '.join(match_details),
    }


def check_allowed_hair_dye(name_zh: str, name_inci: str = '',
                           concentration: float | None = None) -> dict:
    return _check_allowed_generic(name_zh, name_inci,
                                  db.check_allowed_hair_dye, '染发剂', 7, concentration)


# ═══════════════════════════════════════════════
# Layer 4: 使用信息查询
# ═══════════════════════════════════════════════

def query_usage_reference(name_zh: str, name_inci: str = '',
                          concentration: float | None = None) -> dict:
    """Query usage data from market and international sources.

    Returns:
        {'records': [...], 'max_usage_found': float/None,
         'within_range': True/False/None, 'summary': str}
    """
    records = db.query_usage_data(name_zh, name_inci)

    # Parse all max_percent values
    max_values = []
    for r in records:
        v = _parse_percent(r.get('max_percent'))
        if v is not None:
            max_values.append(v)

    overall_max = max(max_values) if max_values else None
    within_range = None
    if overall_max is not None and concentration is not None:
        within_range = concentration <= overall_max

    return {
        'records': records,
        'max_usage_found': overall_max,
        'within_range': within_range,
        'summary': (
            f'查到 {len(records)} 条使用参考'
            + (f'，最高使用量 {overall_max}%' if overall_max else '，无浓度数据')
            + (f'，当前含量 {concentration}%'
               + (' (在安全范围内)' if within_range
                  else ' (超出参考范围!)' if within_range is False
                  else '')
               if concentration is not None else '')
        ),
    }


# ═══════════════════════════════════════════════
# Layer 5: 暴露量计算 (SED)
# ═══════════════════════════════════════════════

def calc_sed(concentration_pct: float,
             daily_amount_g: float,
             retention_factor: float,
             body_weight_kg: float = DEFAULT_BW) -> float:
    """Calculate Systemic Exposure Dosage (SED).

    SED = C × A × F / BW
      C = concentration in formulation (% as decimal e.g. 0.01 for 1%)
      A = daily usage amount (g/day)
      F = retention factor
      BW = body weight (kg)
    """
    c_decimal = concentration_pct / 100.0
    return c_decimal * daily_amount_g * retention_factor / body_weight_kg


def query_exposure_for_assessment(product_category: str,
                                  application_site: str = '') -> dict:
    """Query exposure data and return best match.

    Returns:
        {'records': [...], 'best_match': dict/None, 'summary': str}
    """
    records = db.query_exposure(product_category, application_site)
    if not records:
        # Try broader match - just by site
        if application_site:
            records = db.query_exposure(application_site=application_site)
    return {
        'records': records,
        'best_match': records[0] if records else None,
        'summary': f'查到 {len(records)} 条暴露量数据' if records else '未找到匹配的暴露量数据',
    }


def estimate_sed_for_product(
        concentration_pct: float,
        product_category: str = '',
        application_site: str = '',
        daily_amount_override: float | None = None,
        retention_factor_override: float | None = None,
        data_source_override: str | None = None,
        population: str = '成人',
        body_weight_override: float | None = None
) -> dict:
    """Full SED estimation for an ingredient in a given product type.

    daily_amount_override / retention_factor_override / data_source_override:
      product-level overrides set via UI input boxes. When provided, they
      take precedence over the DB lookup values.

    population / body_weight_override: use population label for DB lookup,
      or body_weight_override to directly override body weight.

    Returns detailed result with all intermediate values.
    """
    # Get body weight
    if body_weight_override is not None and body_weight_override > 0:
        body_weight = body_weight_override
    else:
        bw = db.get_body_weight(population)
        body_weight = bw.get('default_weight_kg', DEFAULT_BW)

    # Get exposure data (DB lookup)
    exposure = query_exposure_for_assessment(product_category, application_site)
    best = exposure['best_match']

    # Apply overrides: UI input > DB value > None
    daily_g = daily_amount_override if daily_amount_override is not None else (
        best['daily_amount_g'] if best else None)
    retention = retention_factor_override if retention_factor_override is not None else (
        best['retention_factor'] if best else None)
    source = data_source_override if data_source_override else (
        best.get('source', '') if best else '')

    if daily_g is None or retention is None:
        return {
            'sed': None,
            'daily_amount_g': daily_g,
            'retention_factor': retention,
            'body_weight_kg': body_weight,
            'product_category': product_category,
            'concentration_pct': concentration_pct,
            'source': source,
            'error': '未找到匹配的暴露量数据',
            'summary': '无法计算 SED：缺少暴露量数据',
        }

    sed = calc_sed(concentration_pct, daily_g, retention, body_weight)
    sed_formula_str = (
        f'SED = {concentration_pct}% / 100 * {daily_g}g * {retention} / {body_weight}kg'
        if daily_amount_override is not None or retention_factor_override is not None
        else f'SED = {concentration_pct}% / 100 * {daily_g}g * {retention} / {body_weight}kg'
    )

    return {
        'sed': round(sed, 6),
        'daily_amount_g': daily_g,
        'retention_factor': retention,
        'body_weight_kg': body_weight,
        'product_category': best.get('product_category', product_category) if best else product_category,
        'application_site': best.get('application_site', application_site) if best else application_site,
        'source': source,
        'concentration_pct': concentration_pct,
        'sed_formula': sed_formula_str,
        'summary': f'SED = {round(sed, 6)} mg/kg bw/day',
        'overridden': daily_amount_override is not None or retention_factor_override is not None,
    }


def calc_moe(noael: float, sed: float) -> float:
    """Calculate Margin of Exposure (MOE = NOAEL / SED)."""
    if not sed or sed == 0:
        return float('inf')
    return round(noael / sed, 2)


# ═══════════════════════════════════════════════
# 综合评估: 对配方进行完整安全评估
# ═══════════════════════════════════════════════

def assess_ingredient(
    name_zh: str,
    name_inci: str = '',
    std_name: str = '',
    concentration: float | None = None,
    product_category: str = '',
    application_site: str = '',
    purpose: str = '',
    daily_amount_override: float | None = None,
    retention_factor_override: float | None = None,
    data_source_override: str | None = None,
    population: str = '成人',
    body_weight_override: float | None = None,
    cas_no: str = '',
    is_child_product: bool = False,
) -> dict:
    """Run ALL checks on a single ingredient and return a comprehensive result.

    Returns a dict with keys:
        ingredient, std_name, concentration, purpose,
        banned, restricted, allowed_preservative, allowed_sunscreen,
        allowed_colorant, allowed_hair_dye,
        usage_reference, exposure, sed,
        risk_substances,
        summary, overall_pass
    """
    result = {
        'ingredient': name_zh,
        'std_name': std_name,
        'inci_name': name_inci,
        'cas_no': cas_no,
        'concentration': concentration,
        'purpose': purpose,
    }

    # Layer 1: Banned check (priority: CAS > INCI > Chinese name)
    result['banned'] = check_banned(name_zh, name_inci, cas_no)

    # Layer 2: Restricted
    result['restricted'] = check_restricted(name_zh, name_inci, concentration)

    # Layer 3: Allowed lists
    result['allowed_preservative'] = check_allowed_preservative(
        name_zh, name_inci, concentration)
    result['allowed_sunscreen'] = check_allowed_sunscreen(
        name_zh, name_inci, concentration)
    result['allowed_colorant'] = check_allowed_colorant(name_zh, name_inci)
    result['allowed_hair_dye'] = check_allowed_hair_dye(
        name_zh, name_inci, concentration)

    # Layer 4: Usage reference
    result['usage_reference'] = query_usage_reference(
        name_zh, name_inci, concentration)

    # Layer 5: SED estimation (only if concentration is available)
    if concentration is not None:
        result['exposure'] = estimate_sed_for_product(
            concentration, product_category, application_site,
            daily_amount_override, retention_factor_override, data_source_override,
            population, body_weight_override)
    else:
        result['exposure'] = {'sed': None, 'summary': '未提供浓度，无法计算 SED'}

    # Layer 6: Risk substance identification (based on 《化妆品风险物质识别与评估技术指导原则》)
    result['risk_substances'] = identify_risk_substances(name_zh, is_child_product)

    # Overall assessment summary
    issues = []
    if result['banned']['banned']:
        issues.append(f"禁用: {result['banned']['summary']}")
    if result['restricted'].get('violations'):
        issues.append(f"限用超标: {'; '.join(result['restricted']['violations'])}")
    if result['usage_reference'].get('within_range') is False:
        issues.append('使用浓度超出参考范围')

    result['overall_pass'] = len(issues) == 0
    result['issues'] = issues
    result['summary'] = '✓ 通过' if result['overall_pass'] else f'✗ 不通过: {"; ".join(issues)}'

    return result


def assess_formula(
    formula_items: list[dict],
    product_category: str = '',
    application_site: str = '',
    daily_amount_override: float | None = None,
    retention_factor_override: float | None = None,
    data_source_override: str | None = None,
    population: str = '成人',
    body_weight_override: float | None = None,
    is_child_product: bool = False,
) -> dict:
    """Run full safety assessment on an entire formula.

    formula_items: list of dicts from db.get_current_formula()
                   each needs: name_zh, name_inci, composition (JSON),
                   added_percent, purpose_override/default_purpose

    daily_amount_override / retention_factor_override / data_source_override:
      product-level overrides from UI input boxes; applied to ALL ingredients.
    population / body_weight_override: overrides for population and weight.
    is_child_product: whether this is a children's product (affects risk substance identification)
    """
    results = []
    for item in formula_items:
        # Parse concentration
        base_concentration = item.get('added_percent')
        if base_concentration is not None:
            base_concentration = float(base_concentration)

        # Determine purpose
        purpose = (item.get('purpose_override') or
                   item.get('default_purpose') or '')

        # Get INCI from name_inci field
        name_inci = item.get('name_inci', '') or ''

        # Get composition
        comps = []
        composition = item.get('composition', '')
        if composition:
            try:
                comps = json.loads(composition)
            except Exception:
                pass

        # If no composition data, use the raw material itself as a single component
        if not comps:
            comps = [{
                'name': item.get('name_zh', ''),
                'percent': 100,
                'inci': name_inci
            }]

        # Evaluate each component separately
        for idx, comp in enumerate(comps):
            comp_name = comp.get('name', '')
            comp_percent = comp.get('percent', 100)
            comp_inci = comp.get('inci', '') or name_inci
            comp_cas = comp.get('cas', '') or item.get('cas_no', '') or ''

            # Calculate actual concentration for this component
            actual_concentration = base_concentration * comp_percent / 100 if base_concentration else None

            # Use component INCI if available, otherwise fall back to the raw material INCI
            result_inci = comp_inci

            result = assess_ingredient(
                name_zh=comp_name,
                name_inci=result_inci,
                std_name=comp_name,
                concentration=actual_concentration,
                product_category=product_category,
                application_site=application_site,
                purpose=purpose,
                daily_amount_override=daily_amount_override,
                retention_factor_override=retention_factor_override,
                data_source_override=data_source_override,
                population=population,
                body_weight_override=body_weight_override,
                cas_no=comp_cas,
                is_child_product=is_child_product,
            )
            result['formula_id'] = item.get('id')
            result['component_index'] = idx
            result['component_percent'] = comp_percent
            result['base_concentration'] = base_concentration
            result['raw_material_name'] = item.get('name_zh', '')
            results.append(result)

    # Summary statistics
    total = len(results)
    passed = sum(1 for r in results if r['overall_pass'])
    failed_ingredients = [r for r in results if not r['overall_pass']]

    return {
        'results': results,
        'product_category': product_category,
        'application_site': application_site,
        'total_ingredients': total,
        'passed': passed,
        'failed': total - passed,
        'failed_ingredients': failed_ingredients,
        'all_pass': passed == total,
        'summary': f'配方安全评估: {passed}/{total} 通过'
                   + (', 全部通过 ✓' if passed == total
                      else f', {total - passed} 个原料存在问题 ✗'),
    }
