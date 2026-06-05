"""
从 PubChem + COSING 数据库爬取原料 CAS 号数据
数据来源:
  - PubChem API (主数据源, ~0.3s/个)
  - COSING HTTP API (备选, ~0.12s/个, 逆向分析所得)
    POST https://api.tech.ec.europa.eu/search-api/prod/rest/search
"""
import os
import sys
import json
import time
import re
import requests
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed


def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_userdata_dir():
    base = get_base_path()
    userdata = os.path.join(base, "userdata")
    os.makedirs(userdata, exist_ok=True)
    return userdata


def normalize_quotes(name):
    """
    归一化各种引号字符，统一为标准直撇号 '
    解决 4,4'- 和 4,4'- 等弯引号与直引号匹配不上的问题
    """
    if not name:
        return name
    # 弯引号 → 直撇号
    return (name
            .replace('\u2018', "'")  # ' LEFT SINGLE QUOTATION MARK
            .replace('\u2019', "'")  # ' RIGHT SINGLE QUOTATION MARK
            .replace('\u201C', '"')  # " LEFT DOUBLE QUOTATION MARK
            .replace('\u201D', '"')  # " RIGHT DOUBLE QUOTATION MARK
            .replace('\u02BC', "'")  # ʼ MODIFIER LETTER APOSTROPHE
            .replace('\u02C8', "'")  # ˈ MODIFIER LETTER VERTICAL LINE
            .replace('\u2032', "'")  # ′ PRIME
            .replace('\u2033', '"')  # ″ DOUBLE PRIME
            )


def fetch_cas_from_pubchem(name):
    """
    根据物质名称从 PubChem 查询 CAS 号（主数据源）
    CAS 号在 synonyms 列表中（格式: 7732-18-5）
    """
    url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{name}/synonyms/JSON"
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            info_list = data.get('InformationList', {}).get('Information', [])
            if info_list:
                synonyms = info_list[0].get('Synonym', [])
                cas_pattern = re.compile(r'^\d{2,7}-\d{2}-\d$')
                for syn in synonyms:
                    if cas_pattern.match(syn):
                        return syn
    except:
        pass
    return ''


def _normalize_name(text):
    """归一化 INCI 名称中的特殊字符"""
    char_map = {
        '\u2018': "'", '\u2019': "'",  # 弯引号 -> 直撇
        '\u201C': '"', '\u201D': '"',  # 弯双引号 -> 直双引号
        '\u2013': '-', '\u2014': '-',  # 短/长破折号 -> 连字符
        '\u00A0': ' ',                  # 不换行空格 -> 普通空格
    }
    for k, v in char_map.items():
        text = text.replace(k, v)
    return text


def _cosing_search(name, field):
    """底层 COSING API 搜索调用，返回原始结果列表"""
    url = "https://api.tech.ec.europa.eu/search-api/prod/rest/search"
    api_key = "285a77fd-1257-4271-8507-f0c6b2961203"

    query = {
        "bool": {
            "must": [
                {
                    "text": {
                        "query": name,
                        "fields": [field],
                        "defaultOperator": "AND"
                    }
                }
            ]
        }
    }

    params = {
        "apiKey": api_key,
        "text": name,
        "pageSize": 10,
        "pageNumber": 1
    }

    files = {
        "query": ("query.json", json.dumps(query), "application/json")
    }

    try:
        resp = requests.post(url, params=params, files=files, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            return data.get("results", [])
    except Exception:
        pass
    return []


def _extract_info_from_results(results, name):
    """从搜索结果中提取最匹配的 (CAS号, EC号)"""
    # 归一化名称用于比较
    name_norm = _normalize_name(name.lower().strip())

    def _first_val(meta, key):
        vals = meta.get(key, [])
        if vals and vals[0] not in ('-', ''):
            return vals[0]
        return ''

    # 优先: 精确名称匹配 + Active 状态
    for r in results:
        meta = r.get("metadata", {})
        inci_names = [_normalize_name(n.lower().strip()) for n in meta.get("inciName", [])]
        status = meta.get("status", [])

        if name_norm in inci_names and "Active" in status:
            cas = _first_val(meta, "casNo")
            ec = _first_val(meta, "ecNo")
            if cas:
                return cas, ec

    # 备选: 名称前缀/包含匹配
    for r in results:
        meta = r.get("metadata", {})
        inci_names = meta.get("inciName", [])
        has_match = any(name_norm in _normalize_name(n.lower().strip())
                        or _normalize_name(n.lower().strip()) in name_norm
                        for n in inci_names)
        if has_match:
            cas = _first_val(meta, "casNo")
            if cas:
                ec = _first_val(meta, "ecNo")
                return cas, ec

    # 最后: 返回第一条有 CAS 的结果
    for r in results:
        meta = r.get("metadata", {})
        cas = _first_val(meta, "casNo")
        if cas:
            ec = _first_val(meta, "ecNo")
            return cas, ec

    return '', ''


def _gen_query_variants(name):
    """生成名称的多个查询变体，应对不同分词策略"""
    variants = [name]  # 原始归一化名称

    # 将 / 替换为空格（ES 全文检索处理 / 分词异常）
    if '/' in name:
        variants.append(name.replace('/', ' '))

    # 将 - 替换为空格
    if '-' in name:
        variants.append(name.replace('-', ' '))

    # 去除所有引号
    if "'" in name or '"' in name:
        variants.append(name.replace("'", '').replace('"', ''))

    # 全替换: 特殊字符变空格
    import re as _re
    cleaned = _re.sub(r"[/'\"-]", ' ', name)
    cleaned = _re.sub(r'\s+', ' ', cleaned).strip()
    if cleaned != name and cleaned not in variants:
        variants.append(cleaned)

    return variants


def fetch_cas_ec_from_cosing_api(name):
    """
    使用 COSING HTTP API 查询 (CAS号, EC号)
    基于逆向分析的内置 API:
      POST https://api.tech.ec.europa.eu/search-api/prod/rest/search
    响应时间 ~120ms/个
    """
    normalized = _normalize_name(name)
    # 生成多个查询变体
    search_terms = [normalized] + _gen_query_variants(normalized)
    # 去重
    seen = set()
    search_terms = [x for x in search_terms if not (x in seen or seen.add(x))]

    for term in search_terms:
        # 策略1: 精确字段搜索 (inciName.exact)
        results = _cosing_search(term, "inciName.exact")
        cas, ec = _extract_info_from_results(results, name)
        if cas:
            return cas, ec

        # 策略2: 全文检索 (inciName)
        results = _cosing_search(term, "inciName")
        cas, ec = _extract_info_from_results(results, name)
        if cas:
            return cas, ec

    return '', ''


def fetch_cas_from_cosing_api(name):
    """兼容包装：只返回 CAS 号"""
    cas, _ = fetch_cas_ec_from_cosing_api(name)
    return cas


def _migrate_cosing_data(data):
    """自动迁移旧格式 {name: cas_str} 到新格式 {name: {cas: str, ec: str}}"""
    migrated = {}
    for k, v in data.items():
        if isinstance(v, str):
            migrated[k] = {"cas": v, "ec": ""}
        elif isinstance(v, dict):
            migrated[k] = {"cas": v.get("cas", ""), "ec": v.get("ec", "")}
        else:
            migrated[k] = {"cas": "", "ec": ""}
    return migrated


def _load_existing(output_path):
    """加载已有的 CAS 数据，对键名做引号归一化，自动迁移旧格式"""
    existing = {}
    if os.path.exists(output_path):
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            existing = {normalize_quotes(k): v for k, v in raw.items()}
            existing = _migrate_cosing_data(existing)
        except:
            pass
    return existing


def _save_output(output_path, result):
    """保存 CAS 数据到文件"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)


def _download_cosing_chunk(names, existing_count, result, progress_callback=None, max_workers=15):
    """并发查询一批物质，直接写入 result 字典"""
    total = len(names)
    done = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {}
        for name in names:
            future = executor.submit(fetch_cas_ec_from_cosing_api, name)
            norm = normalize_quotes(name)
            futures[future] = (name, norm)

        for future in as_completed(futures):
            name, norm = futures[future]
            cas, ec = future.result()
            if cas:
                ec_str = f" EC:{ec}" if ec else ""
                print(f"  {name} -> {cas}{ec_str}")
                result[norm] = {"cas": cas, "ec": ec}
            else:
                print(f"  {name} -> COSING 未找到")

            done += 1
            if progress_callback:
                progress_callback(done, total, len(result) - existing_count)


def download_cosing_only(names, output_path=None, progress_callback=None, incremental=True, max_workers=15, chunk_size=100):
    """
    使用 COSING API 并发查询 CAS 号（~0.3-1.5s/个，并发15线程）
    COSING 服务器在欧盟，延迟 ~300ms，通过并发平摊延迟
    建议让用户选择模式后挂机运行
    """
    if output_path is None:
        output_path = os.path.join(get_userdata_dir(), "cosing_data.json")

    existing = _load_existing(output_path)

    if incremental:
        normalized_existing = set(existing.keys())
        names = [n for n in names if normalize_quotes(n) not in normalized_existing]
        if not names:
            print("所有物质已有 CAS 数据，无需爬取")
            return existing

    total = len(names)
    result = dict(existing)
    existing_count = len(existing)

    # 分块：每 chunk_size 个保存一次，防止中途失败丢失进度
    chunks = [names[i:i + chunk_size] for i in range(0, len(names), chunk_size)]
    print(f"COSING API 并发查询启动（{max_workers} 线程），共 {total} 个物质，分 {len(chunks)} 批保存...\n")

    for cidx, chunk in enumerate(chunks):
        print(f"--- 批 {cidx + 1}/{len(chunks)}（{len(chunk)} 个物质）---")
        _download_cosing_chunk(chunk, existing_count, result, progress_callback, max_workers)
        # 每批保存一次
        _save_output(output_path, result)
        new_count = len(result) - existing_count
        print(f"  已保存，累计 {len(result)} 条（本次新增 {new_count} 条）\n")

    new_count = len(result) - existing_count
    print(f"COSING 查询完成！共 {len(result)} 条数据（本次新增 {new_count} 条）")
    return result


def _download_pubchem_chunk(names, existing_count, result, progress_callback=None, max_workers=10):
    """并发查询一批 PubChem 物质，直接写入 result 字典"""
    total = len(names)
    done = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {}
        for name in names:
            future = executor.submit(fetch_cas_from_pubchem, name)
            norm = normalize_quotes(name)
            futures[future] = (name, norm)

        for future in as_completed(futures):
            name, norm = futures[future]
            cas = future.result()
            if cas:
                print(f"  {name} -> {cas} (PubChem)")
                result[norm] = {"cas": cas, "ec": ""}
            else:
                print(f"  {name} -> PubChem 未找到")

            done += 1
            if progress_callback:
                progress_callback(done, total, len(result) - existing_count)


def download_pubchem_fallback(names, output_path=None, progress_callback=None, incremental=True, max_workers=10, chunk_size=100):
    """
    对 COSING 未查到的物质使用 PubChem 并发补查
    PubChem CDN 快，默认 10 并发，分块保存
    """
    if output_path is None:
        output_path = os.path.join(get_userdata_dir(), "cosing_data.json")

    existing = _load_existing(output_path)

    if incremental:
        normalized_existing = set(existing.keys())
        names = [n for n in names if normalize_quotes(n) not in normalized_existing]
        if not names:
            print("所有物质已有 CAS 数据，无需补查")
            return existing

    total = len(names)
    result = dict(existing)
    existing_count = len(existing)

    chunks = [names[i:i + chunk_size] for i in range(0, len(names), chunk_size)]
    print(f"PubChem 并发补查启动（{max_workers} 线程），共 {total} 个物质，分 {len(chunks)} 批保存...\n")

    for cidx, chunk in enumerate(chunks):
        print(f"--- 批 {cidx + 1}/{len(chunks)}（{len(chunk)} 个物质）---")
        _download_pubchem_chunk(chunk, existing_count, result, progress_callback, max_workers)
        _save_output(output_path, result)
        new_count = len(result) - existing_count
        print(f"  已保存，累计 {len(result)} 条（本次新增 {new_count} 条）\n")

    new_count = len(result) - existing_count
    print(f"PubChem 补查完成！共获取 {len(result)} 条数据（本次新增 {new_count} 条）")
    return result


def load_cosing_data(input_path=None):
    """加载 CAS/EC 号数据（自动迁移旧格式）"""
    if input_path is None:
        input_path = os.path.join(get_userdata_dir(), "cosing_data.json")
    if not os.path.exists(input_path):
        return {}
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            raw = json.load(f)
        return _migrate_cosing_data(raw)
    except:
        return {}


def _get_cas_val(entry):
    """从新旧格式的 entry 中提取 CAS 号"""
    if isinstance(entry, str):
        return entry
    if isinstance(entry, dict):
        return entry.get("cas", "")
    return ""


def _get_ec_val(entry):
    """从新格式 entry 中提取 EC 号"""
    if isinstance(entry, dict):
        return entry.get("ec", "")
    return ""


def match_cas(name, cosing_data):
    """根据名称匹配 CAS 号（支持引号归一化，兼容新旧格式）"""
    if not name or not cosing_data:
        return ''
    name = name.strip()
    norm = normalize_quotes(name)
    # 精确匹配
    if norm in cosing_data:
        return _get_cas_val(cosing_data[norm])
    # 忽略大小写匹配
    name_lower = norm.lower().replace(' ', '')
    for key, entry in cosing_data.items():
        if key.lower().replace(' ', '') == name_lower:
            return _get_cas_val(entry)
    return ''


def match_ec(name, cosing_data):
    """根据名称匹配 EC 号（需要新格式数据）"""
    if not name or not cosing_data:
        return ''
    name = name.strip()
    norm = normalize_quotes(name)
    # 精确匹配
    if norm in cosing_data:
        return _get_ec_val(cosing_data[norm])
    # 忽略大小写匹配
    name_lower = norm.lower().replace(' ', '')
    for key, entry in cosing_data.items():
        if key.lower().replace(' ', '') == name_lower:
            return _get_ec_val(entry)
    return ''


if __name__ == "__main__":
    import pandas as pd

    catalog_path = os.path.join(get_userdata_dir(), "ingredient_catalog.json")
    if not os.path.exists(catalog_path):
        print("目录 I 文件不存在，请先运行目录爬取")
        sys.exit(1)

    catalog = pd.read_json(catalog_path)
    names = catalog['inci_name'].dropna().tolist()

    print(f"将从 COSING + PubChem 查询 {len(names)} 个物质的 CAS 号...")
    print("支持增量爬取，自动跳过已查询的物质")
    print()

    # 先跑 COSING（快，~0.12s/个）
    download_cosing_only(names, incremental=True)
    # 再对 COSING 未查到的用 PubChem 补
    download_pubchem_fallback(names, incremental=True)
