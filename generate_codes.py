#!/usr/bin/env python3
"""万能办公助手 - 激活码批量生成器
生成200个通用激活码，导出CSV供面包多(mbd.pub)上传为卡密
"""
import hmac, hashlib, json, csv
from pathlib import Path

# ⚠️ 这个密钥必须和 office_assistant.py 中的 _ACTIVATION_KEY 完全一致！
SECRET = b"YTQJ2025_OFFICE_PRO_V6"
OUTPUT_FILE = Path.home() / "Desktop" / "万能办公助手_激活码_200个.csv"

def generate_code(cid: int) -> str:
    """生成第cid个激活码"""
    raw = hmac.new(SECRET, str(cid).encode(), hashlib.sha256).hexdigest()[:16].upper()
    return f"{raw[:4]}-{raw[4:8]}-{raw[8:12]}-{raw[12:16]}"

def verify_code(code: str) -> bool:
    """验证激活码是否有效"""
    raw = code.replace("-", "").upper()
    if len(raw) != 16:
        return False
    for cid in range(1, 2001):
        expected = hmac.new(SECRET, str(cid).encode(), hashlib.sha256).hexdigest()[:16].upper()
        if raw == expected:
            return True
    return False

# 生成200个激活码
codes = []
print("=" * 60)
print("万能办公助手 - 激活码生成器")
print("=" * 60)
print(f"\n密钥指纹: {hashlib.sha256(SECRET).hexdigest()[:16]}")
print(f"生成数量: 200个")
print(f"有效期:   1年（从首次激活起算）")
print(f"售价:     ¥99/年")
print()

for i in range(1, 201):
    code = generate_code(i)
    codes.append({
        "序号": i,
        "激活码": code,
        "状态": "未使用",
        "备注": f"批号#1 序号{i:03d}"
    })

# 验证前几个
print("前5个激活码验证:")
for c in codes[:5]:
    ok = "✅" if verify_code(c["激活码"]) else "❌"
    print(f"  {ok} #{c['序号']:03d}: {c['激活码']}")

# 验证最后1个
last = codes[-1]
ok = "✅" if verify_code(last["激活码"]) else "❌"
print(f"  {ok} #{last['序号']:03d}: {last['激活码']}")

# 验证一个伪造码
fake_code = "AAAA-BBBB-CCCC-DDDD"
print(f"  ❌ 伪造码({fake_code}): {'误通过❌' if verify_code(fake_code) else '正确拒绝✅'}")

# 导出CSV（面包多卡密格式）
with open(OUTPUT_FILE, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(["激活码", "序号", "备注"])
    for c in codes:
        writer.writerow([c["激活码"], c["序号"], c["备注"]])

print(f"\n📄 CSV已导出: {OUTPUT_FILE}")
print(f"   共 {len(codes)} 个激活码")
print(f"\n{'='*60}")
print("面包多上传步骤:")
print("1. 打开 https://mbd.pub 注册/登录")
print("2. 创建商品 → 虚拟商品 → 卡密发货")
print("3. 上传此CSV文件作为卡密库")
print("4. 设置价格 ¥99/年")
print("5. 上架！用户付款后自动发码")
print("=" * 60)
input("\n按Enter退出...")
