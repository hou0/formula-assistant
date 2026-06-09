import json

def clean_risk_data():
    """Clean up the risk substances JSON file."""
    with open('risk_substances.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Clean general risk substances
    cleaned_general = []
    for r in data['general_risk_substances']:
        name = r['name'].strip()
        
        # Remove invalid entries
        if name == '无' or not name:
            print(f"移除无效条目: {name}")
            continue
        
        # Fix 二噁烷 -> 二𫫇烷
        if name == '二噁烷':
            r['name'] = '二𫫇烷'
            print(f"修正名称: 二噁烷 -> 二𫫇烷")
        
        # Clean special characters
        if '#' in name:
            # Remove # and everything after
            cleaned_name = name.split('#')[0].strip()
            if cleaned_name:
                r['name'] = cleaned_name
                print(f"清理特殊字符: {name} -> {cleaned_name}")
            else:
                print(f"移除无效条目: {name}")
                continue
        
        # Remove entries with special symbols that don't make sense
        if name.startswith('★'):
            print(f"移除特殊符号条目: {name}")
            continue
        
        cleaned_general.append(r)
    
    data['general_risk_substances'] = cleaned_general
    
    # Remove duplicates by name
    seen = set()
    unique_general = []
    for r in data['general_risk_substances']:
        if r['name'] not in seen:
            seen.add(r['name'])
            unique_general.append(r)
        else:
            print(f"移除重复条目: {r['name']}")
    
    data['general_risk_substances'] = unique_general
    
    # Save cleaned data
    with open('risk_substances.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n清理完成!")
    print(f"通用风险物质: {len(data['general_risk_substances'])}种")
    print(f"儿童风险物质: {len(data['child_risk_substances'])}种")
    
    print("\n最终风险物质列表:")
    for r in data['general_risk_substances']:
        print(f"  - {r['name']}")

if __name__ == '__main__':
    clean_risk_data()