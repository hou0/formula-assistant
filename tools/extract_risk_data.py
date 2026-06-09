import pandas as pd
import json

def extract_risk_data_from_excel(excel_path):
    """Extract risk substance data from Excel file."""
    try:
        # Read all sheets
        xls = pd.ExcelFile(excel_path)
        sheets = xls.sheet_names
        print(f"Excel文件包含 {len(sheets)} 个工作表: {sheets}")
        
        all_risk_data = {}
        
        for sheet_name in sheets:
            print(f"\n=== 工作表: {sheet_name} ===")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Print column names
            print(f"列名: {list(df.columns)}")
            
            # Show first few rows
            print("\n前5行数据:")
            print(df.head(5).to_string())
            
            # Check if this sheet contains risk substance data
            risk_columns = ['风险物质', '风险物质名称', '物质名称', '原料名称', 
                           '触发原料', '关键词', '限值', '限量', '浓度',
                           'CAS', '参考依据', '描述', '说明']
            
            has_risk_data = any(col in df.columns for col in risk_columns)
            if has_risk_data:
                print("\n✓ 发现风险物质数据")
                all_risk_data[sheet_name] = df.to_dict('records')
            else:
                print("\n- 未发现风险物质数据")
        
        return all_risk_data
    
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None

def convert_to_json_format(risk_data):
    """Convert extracted data to the risk_substances.json format."""
    general_rules = []
    child_rules = []
    
    for sheet_name, records in risk_data.items():
        for record in records:
            # Try to extract fields from different possible column names
            name = record.get('风险物质') or record.get('风险物质名称') or record.get('物质名称') or record.get('原料名称')
            keywords = record.get('触发原料') or record.get('关键词') or record.get('触发关键词')
            limit = record.get('限值') or record.get('限量') or record.get('浓度') or record.get('最大浓度')
            reference = record.get('参考依据') or record.get('参考文献') or record.get('标准')
            description = record.get('描述') or record.get('说明') or record.get('备注')
            
            if not name:
                continue
            
            # Parse keywords
            trigger_keywords = []
            if keywords:
                if isinstance(keywords, str):
                    # Split by common separators
                    separators = ['、', '，', ',', ';', '；', ' ', '\n']
                    for sep in separators:
                        if sep in keywords:
                            trigger_keywords = [k.strip() for k in keywords.split(sep) if k.strip()]
                            break
                    if not trigger_keywords:
                        trigger_keywords = [keywords.strip()]
                elif isinstance(keywords, list):
                    trigger_keywords = keywords
            
            rule = {
                'name': str(name).strip(),
                'trigger_keywords': trigger_keywords,
                'trigger_all': len(trigger_keywords) == 0,
                'limit': str(limit).strip() if limit else '',
                'reference': str(reference).strip() if reference else '',
                'description': str(description).strip() if description else ''
            }
            
            # Determine if it's child-specific
            if '儿童' in sheet_name or 'child' in sheet_name.lower():
                child_rules.append(rule)
            else:
                general_rules.append(rule)
    
    return {
        'general_risk_substances': general_rules,
        'child_risk_substances': child_rules
    }

def merge_with_existing_json(new_data, existing_json_path):
    """Merge new risk data with existing JSON file."""
    try:
        with open(existing_json_path, 'r', encoding='utf-8') as f:
            existing = json.load(f)
    except:
        existing = {
            'version': '1.0',
            'last_updated': '2026-06-02',
            'reference': '《化妆品风险物质识别与评估技术指导原则》',
            'general_risk_substances': [],
            'child_risk_substances': []
        }
    
    # Merge general risk substances
    existing_names = set(r['name'] for r in existing['general_risk_substances'])
    for rule in new_data['general_risk_substances']:
        if rule['name'] not in existing_names:
            existing['general_risk_substances'].append(rule)
            print(f"新增通用风险物质: {rule['name']}")
    
    # Merge child risk substances
    child_names = set(r['name'] for r in existing['child_risk_substances'])
    for rule in new_data['child_risk_substances']:
        if rule['name'] not in child_names:
            existing['child_risk_substances'].append(rule)
            print(f"新增儿童风险物质: {rule['name']}")
    
    # Update version and date
    existing['version'] = str(float(existing.get('version', '1.0')) + 0.1)
    existing['last_updated'] = '2026-06-02'
    
    return existing

# Main execution
if __name__ == '__main__':
    excel_path = r'C:\Users\houping\Nutstore\1\我的坚果云\化妆品安全评估\化妆品安全评估工具\AP_DATA.xlsx'
    json_path = r'c:\Users\houping\Desktop\配方表助手\risk_substances.json'
    
    print("="*60)
    print("从Excel提取风险物质数据")
    print("="*60)
    
    # Extract data from Excel
    risk_data = extract_risk_data_from_excel(excel_path)
    
    if risk_data:
        # Convert to JSON format
        print("\n" + "="*60)
        print("转换为JSON格式")
        print("="*60)
        new_data = convert_to_json_format(risk_data)
        
        # Merge with existing JSON
        print("\n" + "="*60)
        print("合并到现有JSON文件")
        print("="*60)
        merged_data = merge_with_existing_json(new_data, json_path)
        
        # Save merged data
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(merged_data, f, ensure_ascii=False, indent=2)
        
        print(f"\n✓ 成功更新 {json_path}")
        print(f"通用风险物质数量: {len(merged_data['general_risk_substances'])}")
        print(f"儿童风险物质数量: {len(merged_data['child_risk_substances'])}")
    else:
        print("\n✗ 未提取到风险物质数据")