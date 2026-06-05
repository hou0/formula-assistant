import json

with open('risk_substances.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print(f"版本: {data['version']}")
print(f"更新日期: {data['last_updated']}")
print(f"通用风险物质: {len(data['general_risk_substances'])}种")
print(f"儿童风险物质: {len(data['child_risk_substances'])}种")

print("\n通用风险物质列表:")
for r in data['general_risk_substances']:
    print(f"  - {r['name']}")

print("\n儿童风险物质列表:")
for r in data['child_risk_substances']:
    print(f"  - {r['name']}")