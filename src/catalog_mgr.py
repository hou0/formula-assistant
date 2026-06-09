import os
import json
import sys

USERDATA_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'userdata'
)
if getattr(sys, 'frozen', False):
    USERDATA_DIR = os.path.join(
        os.path.dirname(sys.executable),
        'userdata'
    )
CATALOG_FILE = os.path.join(USERDATA_DIR, 'ingredient_catalog.json')
CATALOG_FILE_II = os.path.join(
    USERDATA_DIR,
    'ingredient_catalog_ii.json'
)
COSING_FILE = os.path.join(USERDATA_DIR, 'cosing_data.json')


class CatalogManager:
    def __init__(self):
        self.catalog_i = []
        self.catalog_ii = []
        self.catalog_i_dict = {}
        self.catalog_ii_dict = {}
        self.name_lower_index = {}
        self.inci_lower_index = {}
        self.catalog_lookup = {}
        self.cosing_data = {}

        os.makedirs(USERDATA_DIR, exist_ok=True)
        self.load_all()
        self.build_indexes()
        self.load_cosing()

    @staticmethod
    def _get_cas_val(entry):
        """从新旧格式的 entry 中提取 CAS 号"""
        if isinstance(entry, str):
            return entry
        if isinstance(entry, dict):
            return entry.get("cas", "")
        return ""

    @staticmethod
    def _get_ec_val(entry):
        """从新格式 entry 中提取 EC 号"""
        if isinstance(entry, dict):
            return entry.get("ec", "")
        return ""

    @staticmethod
    def _migrate_entry(entry):
        """统一成 {cas, ec} 格式"""
        if isinstance(entry, str):
            return {"cas": entry, "ec": ""}
        if isinstance(entry, dict):
            return {"cas": entry.get("cas", ""), "ec": entry.get("ec", "")}
        return {"cas": "", "ec": ""}

    def load_cosing(self):
        """加载 COSING CAS/EC 号数据（自动迁移旧格式）"""
        self.cosing_data = {}
        if os.path.exists(COSING_FILE):
            try:
                with open(COSING_FILE, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                self.cosing_data = {
                    k: self._migrate_entry(v)
                    for k, v in raw.items()
                }
            except Exception as e:
                print(f"加载COSING数据失败: {e}")

    def match_cas(self, name):
        """
        根据原料INCI名称或中文名称从COSING数据中匹配CAS号
        支持引号归一化（弯引号 ↔ 直撇号）
        """
        if not name or not self.cosing_data:
            return ''

        def _normalize(n):
            return (n.replace('\u2018', "'").replace('\u2019', "'")
                     .replace('\u02BC', "'").replace('\u2032', "'"))

        name = _normalize(name.strip())

        # 精确匹配 INCI 名称
        if name in self.cosing_data:
            return self._get_cas_val(self.cosing_data[name])

        # 忽略大小写精确匹配
        name_lower = name.lower().replace(' ', '')
        for key, entry in self.cosing_data.items():
            if key.lower().replace(' ', '') == name_lower:
                return self._get_cas_val(entry)

        # 中文名称反向匹配：先找对应的INCI，再匹配CAS
        if name in self.catalog_lookup:
            inci = self.catalog_lookup[name].get('inci_name', '')
            if inci:
                inci_norm = _normalize(inci)
                if inci_norm in self.cosing_data:
                    return self._get_cas_val(self.cosing_data[inci_norm])
                inci_lower = inci_norm.lower().replace(' ', '')
                for key, entry in self.cosing_data.items():
                    if key.lower().replace(' ', '') == inci_lower:
                        return self._get_cas_val(entry)

        return ''

    def match_ec(self, name):
        """
        根据原料INCI名称从COSING数据中匹配EC号
        需要新格式数据（旧格式返回空）
        """
        if not name or not self.cosing_data:
            return ''

        def _normalize(n):
            return (n.replace('\u2018', "'").replace('\u2019', "'")
                     .replace('\u02BC', "'").replace('\u2032', "'"))

        name = _normalize(name.strip())

        # 精确匹配 INCI 名称
        if name in self.cosing_data:
            return self._get_ec_val(self.cosing_data[name])

        # 忽略大小写精确匹配
        name_lower = name.lower().replace(' ', '')
        for key, entry in self.cosing_data.items():
            if key.lower().replace(' ', '') == name_lower:
                return self._get_ec_val(entry)

        # 中文名称反向匹配
        if name in self.catalog_lookup:
            inci = self.catalog_lookup[name].get('inci_name', '')
            if inci:
                inci_norm = _normalize(inci)
                if inci_norm in self.cosing_data:
                    return self._get_ec_val(self.cosing_data[inci_norm])
                inci_lower = inci_norm.lower().replace(' ', '')
                for key, entry in self.cosing_data.items():
                    if key.lower().replace(' ', '') == inci_lower:
                        return self._get_ec_val(entry)

        return ''

    def load_all(self):
        self.load_catalog_i()
        self.load_catalog_ii()

    def load_catalog_i(self):
        self.catalog_i = []
        self.catalog_i_dict = {}
        if os.path.exists(CATALOG_FILE):
            try:
                with open(CATALOG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    self.catalog_i = data
                    for item in data:
                        cn = item.get('cn_name', '')
                        if cn and cn not in self.catalog_i_dict:
                            self.catalog_i_dict[cn] = item
                else:
                    self.catalog_i = list(data.values())
                    self.catalog_i_dict = data
            except Exception as e:
                print(f"加载目录I失败: {e}")

    def load_catalog_ii(self):
        self.catalog_ii = []
        self.catalog_ii_dict = {}
        if os.path.exists(CATALOG_FILE_II):
            try:
                with open(CATALOG_FILE_II, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    self.catalog_ii = data
                    for item in data:
                        cn = item.get('cn_name', '')
                        if cn and cn not in self.catalog_ii_dict:
                            self.catalog_ii_dict[cn] = item
                else:
                    self.catalog_ii = list(data.values())
                    self.catalog_ii_dict = data
            except Exception as e:
                print(f"加载目录II失败: {e}")

    def build_indexes(self):
        self.name_lower_index = {}
        self.inci_lower_index = {}
        self.catalog_lookup = {}

        for item in self.catalog_i:
            cn = item.get('cn_name', '')
            if cn:
                nk = cn.lower().replace(' ', '')
                if nk not in self.name_lower_index:
                    self.name_lower_index[nk] = cn
                if cn not in self.catalog_lookup:
                    self.catalog_lookup[cn] = item
                inci = item.get('inci_name', '')
                if inci:
                    ik = inci.lower().replace(' ', '')
                    if ik not in self.inci_lower_index:
                        self.inci_lower_index[ik] = cn

        for item in self.catalog_ii:
            cn = item.get('cn_name', '')
            if cn:
                nk = cn.lower().replace(' ', '')
                if nk not in self.name_lower_index:
                    self.name_lower_index[nk] = cn
                if cn not in self.catalog_lookup:
                    self.catalog_lookup[cn] = item
                inci = item.get('inci_name', '')
                if inci:
                    ik = inci.lower().replace(' ', '')
                    if ik not in self.inci_lower_index:
                        self.inci_lower_index[ik] = cn

    def match_inci(self, name):
        if not name:
            return ''
        name = name.strip()
        if name in self.catalog_lookup:
            return self.catalog_lookup[name].get('inci_name', '')
        nl = name.lower().replace(' ', '')
        if nl in self.name_lower_index:
            cn = self.name_lower_index[nl]
            return self.catalog_lookup[cn].get('inci_name', '')
        if nl in self.inci_lower_index:
            cn = self.inci_lower_index[nl]
            return self.catalog_lookup[cn].get('inci_name', '')
        return ''

    def in_catalog_i(self, name):
        if not name:
            return False
        name = name.strip()
        if name in self.catalog_i_dict:
            return True
        nl = name.lower().replace(' ', '')
        for cn in self.catalog_i_dict:
            if cn.lower().replace(' ', '') == nl:
                return True
        for item in self.catalog_i:
            inci = item.get('inci_name', '')
            if inci and inci.lower().replace(' ', '') == nl:
                return True
        return False

    def get_catalog_source(self, name):
        if not name:
            return None
        name = name.strip()
        nl = name.lower().replace(' ', '')
        if name in self.catalog_i_dict:
            return 'catalog_i'
        for cn in self.catalog_i_dict:
            if cn.lower().replace(' ', '') == nl:
                return 'catalog_i'
        if name in self.catalog_ii_dict:
            return 'catalog_ii'
        for cn in self.catalog_ii_dict:
            if cn.lower().replace(' ', '') == nl:
                return 'catalog_ii'
        for item in self.catalog_i:
            inci = item.get('inci_name', '')
            if inci and inci.lower().replace(' ', '') == nl:
                return 'catalog_i'
        for item in self.catalog_ii:
            inci = item.get('inci_name', '')
            if inci and inci.lower().replace(' ', '') == nl:
                return 'catalog_ii'
        return None

    def search_cn_name(self, keyword):
        results = []
        if not keyword:
            return results
        if keyword in self.catalog_i_dict:
            info = self.catalog_i_dict[keyword]
            results.append({
                        'cn_name': info['cn_name'],
                        'inci_name': info['inci_name']
                    })
        else:
            seen = set()
            for item in self.catalog_i:
                cn = item.get('cn_name', '')
                if keyword in cn and cn not in seen:
                    seen.add(cn)
                    results.append({
                        'cn_name': cn,
                        'inci_name': item.get('inci_name', '')
                    })
                    if len(results) >= 10:
                        break
        return results

    def get_catalog_info(self):
        def count_records(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                return len(d) if isinstance(d, list) else len(d)
            except Exception:
                return 0
        return {
            'catalog_i': count_records(CATALOG_FILE),
            'catalog_ii': count_records(CATALOG_FILE_II),
            'catalog_i_unique': len(self.catalog_i_dict),
            'catalog_ii_unique': len(self.catalog_ii_dict),
        }

    def save_catalog_i(self, data):
        os.makedirs(os.path.dirname(CATALOG_FILE), exist_ok=True)
        with open(CATALOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.load_catalog_i()
        self.build_indexes()
