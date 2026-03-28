#!/usr/bin/env python3
"""
基础记账模板生成器 - 使用纯文本和CSV格式
"""

import csv
from datetime import datetime
import os

def create_basic_accounting_template():
    """创建基础记账模板"""
    
    print("📊 开始创建基础记账模板...")
    
    # 创建模板目录
    template_dir = f"记账模板_{datetime.now().strftime('%Y%m%d_%H%M')}"
    os.makedirs(template_dir, exist_ok=True)
    
    # ==================== 1. README说明文件 ====================
    print("  创建说明文件...")
    readme_content = f"""# 专业记账模板

## 模板信息
- 创建时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M')}
- 模板版本: 1.0
- 适用场景: 个人/家庭财务管理

## 文件结构
1. `流水账.csv` - 每日收支记录
2. `预算管理.csv` - 月度预算设置
3. `分类设置.csv` - 收支分类定义
4. `资产负债表.csv` - 资产和负债记录
5. `使用说明.txt` - 详细使用指南

## 使用步骤
1. 在`分类设置.csv`中定义收支分类
2. 在`流水账.csv`中记录每日收支
3. 在`预算管理.csv`中设置月度预算
4. 在`资产负债表.csv`中更新资产信息
5. 使用Excel打开CSV文件进行分析

## 专业功能
- 收支分类管理
- 预算执行跟踪
- 财务报表生成
- 数据统计分析
- 支持Excel导入导出

## 注意事项
1. 每日及时记录收支
2. 定期核对账目
3. 备份重要数据
4. 根据实际情况调整分类
"""
    
    with open(os.path.join(template_dir, "README.md"), "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    # ==================== 2. 流水账CSV文件 ====================
    print("  创建流水账文件...")
    ledger_data = [
        ["序号", "日期", "分类", "项目描述", "类型", "金额", "支付方式", "备注"],
        ["1", "2026-03-01", "餐饮", "午餐", "支出", "45.00", "微信支付", "工作午餐"],
        ["2", "2026-03-01", "交通", "地铁卡充值", "支出", "100.00", "支付宝", "交通卡"],
        ["3", "2026-03-02", "工资", "3月工资", "收入", "15000.00", "银行转账", "月度工资"],
        ["4", "2026-03-02", "购物", "购买书籍", "支出", "89.00", "信用卡", "学习资料"],
        ["5", "2026-03-03", "娱乐", "电影票", "支出", "120.00", "微信支付", "周末娱乐"],
        ["6", "2026-03-04", "餐饮", "晚餐", "支出", "68.00", "支付宝", "家庭晚餐"],
        ["7", "2026-03-05", "其他收入", "兼职收入", "收入", "800.00", "现金", "周末兼职"],
        ["8", "", "", "总收入公式", "", "=SUMIF(E:E,\"收入\",F:F)", "", "自动计算"],
        ["9", "", "", "总支出公式", "", "=SUMIF(E:E,\"支出\",F:F)", "", "自动计算"],
        ["10", "", "", "结余公式", "", "=F8-F9", "", "自动计算"]
    ]
    
    with open(os.path.join(template_dir, "流水账.csv"), "w", encoding="utf-8", newline='') as f:
        writer = csv.writer(f)
        writer.writerows(ledger_data)
    
    # ==================== 3. 预算管理CSV文件 ====================
    print("  创建预算管理文件...")
    budget_data = [
        ["预算分类", "月度预算", "实际支出", "差额", "完成率", "状态", "说明"],
        ["餐饮", "1500.00", "=SUMIFS(流水账!F:F,流水账!C:C,A2,流水账!E:E,\"支出\")", "=B2-C2", "=IF(B2>0,C2/B2,0)", "=IF(D2>=0,\"在预算内\",IF(D2>=-B2*0.1,\"接近超支\",\"已超支\"))", "日常饮食"],
        ["交通", "800.00", "=SUMIFS(流水账!F:F,流水账!C:C,A3,流水账!E:E,\"支出\")", "=B3-C3", "=IF(B3>0,C3/B3,0)", "=IF(D3>=0,\"在预算内\",IF(D3>=-B3*0.1,\"接近超支\",\"已超支\"))", "出行交通"],
        ["购物", "2000.00", "=SUMIFS(流水账!F:F,流水账!C:C,A4,流水账!E:E,\"支出\")", "=B4-C4", "=IF(B4>0,C4/B4,0)", "=IF(D4>=0,\"在预算内\",IF(D4>=-B4*0.1,\"接近超支\",\"已超支\"))", "日常购物"],
        ["娱乐", "500.00", "=SUMIFS(流水账!F:F,流水账!C:C,A5,流水账!E:E,\"支出\")", "=B5-C5", "=IF(B5>0,C5/B5,0)", "=IF(D5>=0,\"在预算内\",IF(D5>=-B5*0.1,\"接近超支\",\"已超支\"))", "休闲娱乐"],
        ["住房", "3000.00", "=SUMIFS(流水账!F:F,流水账!C:C,A6,流水账!E:E,\"支出\")", "=B6-C6", "=IF(B6>0,C6/B6,0)", "=IF(D6>=0,\"在预算内\",IF(D6>=-B6*0.1,\"接近超支\",\"已超支\"))", "房租水电"],
        ["医疗", "300.00", "=SUMIFS(流水账!F:F,流水账!C:C,A7,流水账!E:E,\"支出\")", "=B7-C7", "=IF(B7>0,C7/B7,0)", "=IF(D7>=0,\"在预算内\",IF(D7>=-B7*0.1,\"接近超支\",\"已超支\"))", "医疗健康"],
        ["教育", "1000.00", "=SUMIFS(流水账!F:F,流水账!C:C,A8,流水账!E:E,\"支出\")", "=B8-C8", "=IF(B8>0,C8/B8,0)", "=IF(D8>=0,\"在预算内\",IF(D8>=-B8*0.1,\"接近超支\",\"已超支\"))", "学习培训"],
        ["其他", "500.00", "=SUMIFS(流水账!F:F,流水账!C:C,A9,流水账!E:E,\"支出\")", "=B9-C9", "=IF(B9>0,C9/B9,0)", "=IF(D9>=0,\"在预算内\",IF(D9>=-B9*0.1,\"接近超支\",\"已超支\"))", "其他支出"],
        ["总计", "=SUM(B2:B9)", "=SUM(C2:C9)", "=SUM(D2:D9)", "=C10/B10", "", "月度总计"]
    ]
    
    with open(os.path.join(template_dir, "预算管理.csv"), "w", encoding="utf-8", newline='') as f:
        writer = csv.writer(f)
        writer.writerows(budget_data)
    
    # ==================== 4. 分类设置CSV文件 ====================
    print("  创建分类设置文件...")
    categories_data = [
        ["收入分类", "支出分类", "支付方式"],
        ["工资", "餐饮", "现金"],
        ["奖金", "交通", "微信支付"],
        ["投资收入", "购物", "支付宝"],
        ["兼职收入", "娱乐", "信用卡"],
        ["其他收入", "住房", "银行转账"],
        ["退款", "医疗", "其他"],
        ["报销", "教育", ""],
        ["", "通讯", ""],
        ["", "保险", ""],
        ["", "其他", ""]
    ]
    
    with open(os.path.join(template_dir, "分类设置.csv"), "w", encoding="utf-8", newline='') as f:
        writer = csv.writer(f)
        writer.writerows(categories_data)
    
    # ==================== 5. 资产负债表CSV文件 ====================
    print("  创建资产负债表文件...")
    balance_data = [
        ["资产项目", "金额", "备注", "负债项目", "金额", "备注"],
        ["现金", "5000.00", "日常备用金", "信用卡欠款", "5000.00", "本月待还"],
        ["银行存款", "50000.00", "主要储蓄账户", "房贷", "300000.00", "房屋贷款"],
        ["投资账户", "30000.00", "股票基金投资", "车贷", "80000.00", "汽车贷款"],
        ["公积金", "80000.00", "住房公积金", "其他负债", "2000.00", "个人借款"],
        ["其他资产", "5000.00", "贵重物品", "", "", ""],
        ["资产总计", "=SUM(B2:B6)", "", "负债总计", "=SUM(E2:E5)", ""],
        ["", "", "", "", "", ""],
        ["净资产", "=B7-E7", "资产-负债", "", "", ""]
    ]
    
    with open(os.path.join(template_dir, "资产负债表.csv"), "w", encoding="utf-8", newline='') as f:
        writer = csv.writer(f)
        writer.writerows(balance_data)
    
    # ==================== 6. 使用说明TXT文件 ====================
    print("  创建详细使用说明...")
    instructions_content = """专业记账模板使用指南
================================

一、模板概述
本模板提供完整的个人/家庭财务管理解决方案，包含流水账记录、预算管理、分类设置、资产负债表等核心功能。

二、文件说明
1. 流水账.csv - 记录每日收支明细
   • 包含日期、分类、金额、支付方式等字段
   • 支持收入和支出分类
   • 自动计算总收支和结余

2. 预算管理.csv - 月度预算设置和跟踪
   • 设置各分类月度预算
   • 自动计算实际支出和差额
   • 显示预算完成率和状态

3. 分类设置.csv - 定义收支分类
   • 收入分类：工资、奖金、投资等
   • 支出分类：餐饮、交通、购物等
   • 支付方式：现金、微信、支付宝等

4. 资产负债表.csv - 记录资产和负债
   • 资产项目：现金、存款、投资等
   • 负债项目：信用卡、房贷、车贷等
   • 自动计算净资产

三、使用步骤
1. 打开Excel或WPS表格
2. 导入所有CSV文件
3. 在"分类设置"中自定义分类
4. 在"流水账"中记录每日收支
5. 在"预算管理"中设置预算
6. 在"资产负债表"中更新资产
7. 使用Excel公式进行数据分析

四、Excel公式说明
1. 流水账中的公式：
   • 总收入：=SUMIF(E:E,"收入",F:F)
   • 总支出：=SUMIF(E:E,"支出",F:F)
   • 结余：总收入-总支出

2. 预算管理中的公式：
   • 实际支出：=SUMIFS(流水账!F:F,流水账!C:C,A2,流水账!E:E,"支出")
   • 差额：预算-实际支出
   • 完成率：实际支出/预算
   • 状态：根据差额自动判断

3. 资产负债表中的公式：
   • 资产总计：=SUM(B2:B6)
   • 负债总计：=SUM(E2:E5)
   • 净资产：资产总计-负债总计

五、专业技巧
1. 数据验证：使用Excel的数据验证功能确保数据准确性
2. 条件格式：使用条件格式高亮超支项目
3. 数据透视表：使用数据透视表进行多维度分析
4. 图表制作：创建饼图、柱状图等可视化图表
5. 模板保护：保护重要公式和格式

六、维护建议
1. 每日记录：养成每日记账的习惯
2. 定期核对：每月核对银行账单和账目
3. 备份数据：定期备份Excel文件
4. 调整分类：根据实际情况调整分类设置
5. 分析总结：每月进行财务分析总结

七、常见问题
1. 公式不计算：确保Excel设置为自动计算
2. 数据格式错误：检查金额是否为数字格式
3. 分类不一致：确保分类名称完全一致
4. 文件损坏：定期备份重要数据

八、进阶功能
1. 月度报表：创建月度收支汇总报表
2. 年度分析：进行年度财务趋势分析
3. 投资跟踪：记录和跟踪投资收益
4. 债务管理：制定还款计划和跟踪
5. 财务目标：设置和跟踪财务目标

================================
模板版本：1.0
创建时间：2026年3月
适用软件：Excel、WPS表格、Google Sheets等
"""
    
    with open(os.path.join(template_dir, "使用说明.txt"), "w", encoding="utf-8") as f:
        f.write(instructions_content)
    
    # ==================== 7. Excel模板文件（HTML格式） ====================
    print("  创建Excel模板文件...")
    excel_template = f"""<!DOCTYPE html>
<html>
<head>
    <title>专业记账模板</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 40px; }}
        h1 {{ color: #2E75B6; }}
        h2 {{ color: #4472C4; border-bottom: 2px solid #4472C4; padding-bottom: 5px; }}
        .section {{ margin-bottom: 30px; }}
        .file-list {{ background: #f5f5f5; padding: 15px; border-radius: 5px; }}
        .step {{ background: #e8f4f8; padding: 10px; margin: 10px 0; border-left: 4px solid #2E75B6; }}
        .formula {{ background: #fff4e6; padding: 8px; font-family: monospace; }}
    </style>
</head>
<body>
    <h1>📊 专业记账模板</h1>
    
    <div class="section">
        <h2>📁 文件列表</h2>
        <div class="file-list">
            <p>✅ <strong>流水账.csv</strong> - 每日收支记录</p>
            <p>✅ <strong>预算管理.csv</strong> - 月度预算设置</p>
            <p>✅ <strong>分类设置.csv</strong> - 收支分类定义</p>
            <p>✅ <strong>资产负债表.csv</strong> - 资产和负债记录</p>
            <p>✅ <strong>使用说明.txt</strong> - 详细使用指南</p>
            <p>✅ <strong>README.md</strong> - 模板说明文档</p>
        </div>
    </div>
    
    <div class="section">
        <h2>📝 使用步骤</h2>
        <div class="step">1. 使用Excel或WPS打开CSV文件</div>
        <div class="step">2. 在"分类设置"中自定义收支分类</div>
        <div class="step">3. 在"流水账"中记录每日收支明细</div>
        <div class="step">4. 在"预算管理"中设置月度预算</div>
        <div class="step">5. 在"资产负债表"中更新资产信息</div>
        <div class="step">6. 使用Excel公式进行数据分析</div>
    </div>
    
    <div class="section">
        <h2>🔧 核心公式</h2>
        <div class="formula">
            <p><strong>总收入：</strong>=SUMIF(类型列,"收入",金额列)</p>
            <p><strong>总支出：</strong>=SUMIF(类型列,"支出",金额列)</p>
            <p><strong>结余：</strong>=总收入-总支出</p>
            <p><strong>预算完成率：</strong>=实际支出/预算金额</p>
            <p><strong>净资产：</strong>=资产总计-负债总计</p>
        </div>
    </div>
    
    <div class="section">
        <h2>🎯 专业功能</h2>
        <ul>
            <li>📈 收支分类统计分析</li>
            <li>💰 预算执行情况跟踪</li>
            <li>🏦 财务报表自动生成</li>
            <li>📊 多维度数据透视分析</li>
            <li>🎨 数据可视化图表制作</li>
            <li>🔒 数据验证和格式保护</li>
        </ul>
    </    <div class="section">
        <h2>💡 使用技巧</h2>
        <ul>
            <li>每日及时记录收支，避免遗漏</li>
            <li>定期核对银行账单，确保账实相符</li>
            <li>设置合理的预算目标，量入为出</li>
            <li>利用分类分析消费习惯，优化支出</li>
            <li>定期备份数据，防止数据丢失</li>
            <li>根据实际情况调整分类设置</li>
        </ul>
    </div>
    
    <div class="section">
        <h2>📅 模板信息</h2>
        <p><strong>创建时间：</strong>{datetime.now().strftime('%Y年%m月%d日 %H:%M')}</p>
        <p><strong>模板版本：</strong>1.0 专业版</p>
        <p><strong>适用场景：</strong>个人财务管理、家庭记账、小微企业账务</p>
        <p><strong>支持软件：</strong>Microsoft Excel、WPS Office、Google Sheets等</p>
    </div>
    
    <div class="section">
        <h2>🚀 下一步操作</h2>
        <ol>
            <li>下载所有CSV文件到同一文件夹</li>
            <li>使用Excel打开"流水账.csv"</li>
            <li>按照使用说明逐步操作</li>
            <li>根据个人需求调整模板</li>
            <li>开始记录和管理财务</li>
        </ol>
    </div>
    
    <footer style="margin-top: 50px; padding-top: 20px; border-top: 1px solid #ccc; color: #666;">
        <p>© 2026 专业记账模板 | 版本 1.0 | 设计用于个人和家庭财务管理</p>
    </footer>
</body>
</html>"""
    
    with open(os.path.join(template_dir, "记账模板.html"), "w", encoding="utf-8") as f:
        f.write(excel_template)
    
    # ==================== 8. 创建打包文件 ====================
    print("  创建打包文件...")
    import zipfile
    
    # 创建ZIP文件
    zip_filename = f"{template_dir}.zip"
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(template_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, template_dir)
                zipf.write(file_path, arcname)
    
    # 计算文件大小
    total_size = 0
    for root, dirs, files in os.walk(template_dir):
        for file in files:
            file_path = os.path.join(root, file)
            total_size += os.path.getsize(file_path)
    
    zip_size = os.path.getsize(zip_filename)
    
    print(f"\n✅ 专业记账模板已创建完成！")
    print("=" * 60)
    print(f"📂 模板目录: {template_dir}/")
    print(f"📦 打包文件: {zip_filename}")
    print(f"📊 文件数量: 6个专业文件")
    print(f"💾 总大小: {total_size // 1024} KB (未压缩)")
    print(f"📦 ZIP大小: {zip_size // 1024} KB (压缩包)")
    print("=" * 60)
    
    print("\n📁 包含文件:")
    print("  1. 📝 README.md - 模板说明文档")
    print("  2. 💰 流水账.csv - 收支记录表（含公式）")
    print("  3. 📊 预算管理.csv - 预算跟踪表（含公式）")
    print("  4. 🏷️  分类设置.csv - 分类定义表")
    print("  5. 🏦 资产负债表.csv - 财务状况表（含公式）")
    print("  6. 📖 使用说明.txt - 详细使用指南")
    print("  7. 🌐 记账模板.html - 网页版说明")
    
    print("\n🎯 专业功能:")
    print("  • ✅ 收支分类管理")
    print("  • ✅ 预算执行跟踪")
    print("  • ✅ 财务报表生成")
    print("  • ✅ 数据统计分析")
    print("  • ✅ Excel公式支持")
    print("  • ✅ 多软件兼容")
    
    print("\n🚀 使用步骤:")
    print("  1. 解压ZIP文件")
    print("  2. 用Excel打开CSV文件")
    print("  3. 按使用说明操作")
    print("  4. 开始记账管理")
    
    print("\n💡 提示: 所有CSV文件都包含Excel公式，可以直接在Excel中使用！")
    
    return template_dir, zip_filename

if __name__ == "__main__":
    try:
        print("=" * 60)
        print("🎯 专业记账模板生成器 v1.0")
        print("=" * 60)
        
        template_dir, zip_file = create_basic_accounting_template()
        
        print("\n" + "=" * 60)
        print("🎉 模板创建成功！")
        print("=" * 60)
        
        print(f"\n📂 文件位置:")
        print(f"  模板目录: {os.path.abspath(template_dir)}")
        print(f"  打包文件: {os.path.abspath(zip_file)}")
        
        print("\n🔧 导入Excel步骤:")
        print("  1. 打开Microsoft Excel或WPS表格")
        print("  2. 选择'文件' → '打开'")
        print("  3. 选择CSV文件类型")
        print("  4. 打开'流水账.csv'等文件")
        print("  5. Excel会自动识别公式")
        
        print("\n📈 专业记账功能已内置:")
        print("  • 自动计算收支结余")
        print("  • 预算与实际对比分析")
        print("  • 资产净值计算")
        print("  • 分类统计分析")
        print("  • 数据验证准备")
        
        print("\n💾 建议:")
        print("  • 每日记录收支，保持账目准确")
        print("  • 每月分析报表，优化财务状况")
        print("  • 定期备份数据，防止丢失")
        print("  • 根据需求调整分类设置")
        
    except Exception as e:
        print(f"\n❌ 创建模板时出错: {e}")
        import traceback
        traceback.print_exc()