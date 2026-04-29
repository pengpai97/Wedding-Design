#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
婚礼座次表一键更新脚本
用法: python update_seating.py
"""

import pandas as pd
import json
import os
import sys

# Windows控制台编码修复
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def read_excel_data(excel_path):
    """读取Excel数据"""
    df = pd.read_excel(excel_path)
    print(f"✅ 读取Excel成功，共 {len(df)} 条记录")

    result = {}
    # 转换为普通Python整数
    table_numbers = [int(x) for x in sorted(df['桌次'].unique())]
    print(f"📊 涉及桌次: {table_numbers}")

    for table_num in table_numbers:
        table_data = df[df['桌次'] == table_num]
        guests = []
        for _, row in table_data.iterrows():
            guests.append({
                'source': str(row['来源']),
                'name': str(row['姓名']),
                'count': int(row['人数'])
            })
        result[table_num] = {
            'guests': guests,
            'total_people': int(table_data['人数'].sum()),
            'total_groups': len(guests)
        }

    return result

# 桌号对应的显示类别
TABLE_CATEGORIES = {
    1: '男方亲友',
    2: '新人同学',
    3: '新人同学',
    4: '新人同学',
    5: '新人同学',
    6: '男方亲友',
    7: '新人同学',
    8: '新人同学',
    9: '新人同学',
    10: '新人同学',
    11: '峨庄亲友',
    12: '峨庄亲友',
    13: '临淄亲友',
    14: '事务所',
    15: '女方泳友',
    16: '峨庄亲友',
    17: '临淄亲友',
    18: '临淄亲友',
    19: '女方同学',
    20: '机动',
}

def generate_table_html(table_num, table_data):
    """生成单个桌的HTML"""
    category = TABLE_CATEGORIES.get(table_num, '')
    if table_data:
        count_text = f"{table_data['total_people']}人"
    else:
        count_text = "10人"  # 默认10人
    return f'''                <div class="table" data-table="{table_num}" onclick="showTableDetail({table_num})">
                    <div class="table-number">{table_num}</div>
                    <div class="table-source">{category}</div>
                    <div class="table-count">{count_text}</div>
                </div>'''

def generate_js_data(table_data):
    """生成JavaScript数据"""
    json_str = json.dumps(table_data, ensure_ascii=False, indent=12)
    # 调整缩进格式
    lines = json_str.split('\n')
    indented_lines = []
    for line in lines:
        if line.strip() == '{' or line.strip() == '}':
            indented_lines.append('        ' + line)
        else:
            indented_lines.append('        ' + line)
    return '\n'.join(indented_lines)

def generate_seating_html(table_data):
    """生成婚礼座次布局图HTML"""
    # 生成4列x5行的座位布局
    columns = [[], [], [], []]
    for i in range(1, 21):
        col_idx = (i - 1) // 5
        columns[col_idx].append(generate_table_html(i, table_data.get(i)))

    html_template = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>婚礼宴席座次表</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Microsoft YaHei', 'SimHei', sans-serif;
            background: linear-gradient(135deg, #fff5f5 0%, #ffe4e6 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
        }

        h1 {
            text-align: center;
            color: #be123c;
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(190, 18, 60, 0.2);
        }

        .subtitle {
            text-align: center;
            color: #9f1239;
            font-size: 1.2em;
            margin-bottom: 30px;
        }

        .stage {
            background: linear-gradient(135deg, #be123c 0%, #9f1239 100%);
            color: white;
            text-align: center;
            padding: 20px;
            border-radius: 10px 10px 0 0;
            font-size: 1.3em;
            font-weight: bold;
            margin-bottom: 35px;
            box-shadow: 0 6px 18px rgba(190, 18, 60, 0.3);
            position: relative;
            max-width: 850px;
            margin-left: auto;
            margin-right: auto;
        }

        .stage::after {
            content: '';
            position: absolute;
            bottom: -10px;
            left: 50%;
            transform: translateX(-50%);
            width: 100px;
            height: 4px;
            background: #fda4af;
            border-radius: 2px;
        }

        .stage-icon {
            font-size: 1.5em;
            margin-bottom: 8px;
        }

        .seating-area {
            display: flex;
            justify-content: center;
            gap: 65px;
            flex-wrap: wrap;
        }

        .column {
            display: flex;
            flex-direction: column;
            gap: 28px;
        }

        .table {
            width: 130px;
            height: 130px;
            background: linear-gradient(145deg, #ffffff 0%, #fef2f2 100%);
            border: 3px solid #fb7185;
            border-radius: 50%;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(251, 113, 133, 0.2);
            position: relative;
        }

        .table.has-data {
            border-color: #be123c;
            background: linear-gradient(145deg, #fff1f2 0%, #ffe4e6 100%);
        }

        .table:hover {
            transform: scale(1.1);
            box-shadow: 0 8px 25px rgba(251, 113, 133, 0.4);
            border-color: #be123c;
        }

        .table-number {
            font-size: 1.5em;
            font-weight: bold;
            color: #be123c;
            margin-bottom: 5px;
        }

        .table-source {
            font-size: 0.8em;
            color: #db2777;
            font-weight: 500;
        }

        .table-count {
            font-size: 0.75em;
            color: #f43f5e;
            margin-top: 3px;
            background: #fecdd3;
            padding: 2px 8px;
            border-radius: 10px;
        }

        .table.empty .table-number {
            color: #9ca3af;
        }

        .table.empty {
            border-color: #d1d5db;
            background: linear-gradient(145deg, #f9fafb 0%, #f3f4f6 100%);
        }

        .table.empty .table-count {
            background: #e5e7eb;
            color: #6b7280;
        }

        .modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }

        .modal-overlay.active {
            display: flex;
        }

        .modal {
            background: white;
            border-radius: 15px;
            padding: 30px;
            max-width: 500px;
            width: 90%;
            max-height: 80vh;
            overflow-y: auto;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            animation: modalIn 0.3s ease;
        }

        @keyframes modalIn {
            from {
                opacity: 0;
                transform: scale(0.9);
            }
            to {
                opacity: 1;
                transform: scale(1);
            }
        }

        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #fecdd3;
        }

        .modal-title {
            font-size: 1.5em;
            color: #be123c;
            font-weight: bold;
        }

        .modal-close {
            width: 35px;
            height: 35px;
            border-radius: 50%;
            border: none;
            background: #fecdd3;
            color: #be123c;
            font-size: 1.2em;
            cursor: pointer;
            transition: all 0.2s;
        }

        .modal-close:hover {
            background: #be123c;
            color: white;
        }

        .modal-summary {
            background: linear-gradient(135deg, #fff1f2 0%, #ffe4e6 100%);
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
        }

        .modal-summary p {
            color: #9f1239;
            margin: 5px 0;
        }

        .guest-list {
            list-style: none;
        }

        .guest-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 15px;
            border-bottom: 1px solid #f3f4f6;
            transition: background 0.2s;
        }

        .guest-item:hover {
            background: #f9fafb;
        }

        .guest-item:last-child {
            border-bottom: none;
        }

        .guest-info {
            display: flex;
            flex-direction: column;
        }

        .guest-name {
            font-weight: bold;
            color: #1f2937;
            font-size: 1.05em;
        }

        .guest-source {
            font-size: 0.85em;
            color: #6b7280;
            margin-top: 3px;
        }

        .guest-count {
            background: #be123c;
            color: white;
            padding: 5px 12px;
            border-radius: 15px;
            font-size: 0.9em;
            font-weight: bold;
        }

        .empty-notice {
            text-align: center;
            color: #9ca3af;
            padding: 40px 20px;
        }

        .empty-notice-icon {
            font-size: 3em;
            margin-bottom: 15px;
        }

        .legend {
            margin-top: 50px;
            padding: 20px;
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }

        .legend h3 {
            color: #be123c;
            margin-bottom: 15px;
            text-align: center;
        }

        .legend-items {
            display: flex;
            justify-content: center;
            gap: 40px;
            flex-wrap: wrap;
        }

        .legend-item {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .legend-dot {
            width: 20px;
            height: 20px;
            border-radius: 50%;
        }

        .legend-dot.has-data {
            border: 2px solid #be123c;
            background: linear-gradient(145deg, #fff1f2 0%, #ffe4e6 100%);
        }

        .legend-dot.empty {
            border: 2px solid #d1d5db;
            background: linear-gradient(145deg, #f9fafb 0%, #f3f4f6 100%);
        }

        .footer {
            text-align: center;
            margin-top: 40px;
            color: #9f1239;
            font-size: 0.9em;
        }

        @media (max-width: 768px) {
            .seating-area {
                gap: 20px;
            }

            .table {
                width: 100px;
                height: 100px;
            }

            .table-number {
                font-size: 1.2em;
            }

            h1 {
                font-size: 1.8em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>💒 婚礼宴席座次表</h1>
        <p class="subtitle">Wedding Seating Plan</p>

        <div class="stage">
            <div class="stage-icon">💕</div>
            <div>舞台 / STAGE</div>
        </div>

        <div class="seating-area" id="seatingArea">
'''

    # 添加4列
    for col in columns:
        html_template += '            <div class="column">\n'
        html_template += '\n'.join(col)
        html_template += '\n            </div>\n\n'

    html_template += '''        </div>

        <div class="legend">
            <h3>📋 图例说明</h3>
            <div class="legend-items">
                <div class="legend-item">
                    <div class="legend-dot has-data"></div>
                    <span>已安排宾客</span>
                </div>
                <div class="legend-item">
                    <div class="legend-dot empty"></div>
                    <span>空桌 / 备用桌</span>
                </div>
            </div>
        </div>

        <div class="footer">
            <p>👆 点击任意桌号查看详细宾客名单</p>
        </div>
    </div>

    <div class="modal-overlay" id="modalOverlay" onclick="closeModal(event)">
        <div class="modal" onclick="event.stopPropagation()">
            <div class="modal-header">
                <div class="modal-title" id="modalTitle">第 1 桌</div>
                <button class="modal-close" onclick="closeModal()">×</button>
            </div>
            <div id="modalContent">
            </div>
        </div>
    </div>

    <script>
        // Excel数据
        const tableData = '''

    # 添加数据
    html_template += generate_js_data(table_data)

    html_template += ''';

        function initTables() {
            document.querySelectorAll('.table').forEach(table => {
                const tableNum = parseInt(table.dataset.table);
                if (tableData[tableNum]) {
                    table.classList.add('has-data');
                } else {
                    table.classList.add('empty');
                }
            });
        }

        function showTableDetail(tableNum) {
            const modalOverlay = document.getElementById('modalOverlay');
            const modalTitle = document.getElementById('modalTitle');
            const modalContent = document.getElementById('modalContent');

            modalTitle.textContent = `第 ${tableNum} 桌`;

            const data = tableData[tableNum];

            if (data) {
                const sources = [...new Set(data.guests.map(g => g.source))].join(' / ');

                modalContent.innerHTML = `
                    <div class="modal-summary">
                        <p><strong>宾客来源：</strong>${sources}</p>
                        <p><strong>宾客组数：</strong>${data.total_groups} 组</p>
                        <p><strong>总人数：</strong>${data.total_people} 人</p>
                    </div>
                    <ul class="guest-list">
                        ${data.guests.map(guest => `
                            <li class="guest-item">
                                <div class="guest-info">
                                    <span class="guest-name">${guest.name}</span>
                                    <span class="guest-source">${guest.source}</span>
                                </div>
                                <span class="guest-count">${guest.count} 人</span>
                            </li>
                        `).join('')}
                    </ul>
                `;
            } else {
                modalContent.innerHTML = `
                    <div class="modal-summary">
                        <p><strong>宾客来源：</strong>待安排</p>
                        <p><strong>宾客组数：</strong>待安排</p>
                        <p><strong>总人数：</strong>10 人</p>
                    </div>
                    <div class="empty-notice">
                        <p>此桌宾客名单待录入</p>
                    </div>
                `;
            }

            modalOverlay.classList.add('active');
        }

        function closeModal(event) {
            if (event && event.target !== event.currentTarget) return;
            document.getElementById('modalOverlay').classList.remove('active');
        }

        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                closeModal();
            }
        });

        initTables();
    </script>
</body>
</html>'''

    return html_template


def generate_guest_list_html(table_data):
    """生成宾客指引名单HTML"""
    html_template = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>婚礼宾客指引名单</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Microsoft YaHei', 'SimHei', sans-serif;
            background: #fff;
            padding: 20px;
            line-height: 1.6;
        }

        .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 3px solid #be123c;
        }

        .header h1 {
            color: #be123c;
            font-size: 2.5em;
            margin-bottom: 10px;
        }

        .header p {
            color: #666;
            font-size: 1.2em;
        }

        .print-btn {
            background: #be123c;
            color: white;
            border: none;
            padding: 12px 30px;
            font-size: 1em;
            border-radius: 5px;
            cursor: pointer;
            margin-top: 15px;
            transition: background 0.3s;
        }

        .print-btn:hover {
            background: #9f1239;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        .table-section {
            margin-bottom: 30px;
            page-break-inside: avoid;
        }

        .table-header {
            background: linear-gradient(135deg, #be123c 0%, #db2777 100%);
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            font-size: 1.3em;
            font-weight: bold;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }

        .table-category {
            font-size: 0.8em;
            opacity: 0.9;
            font-weight: normal;
        }

        .guest-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
            gap: 10px;
        }

        .guest-card {
            background: #fff5f5;
            border: 1px solid #fecdd3;
            border-radius: 6px;
            padding: 10px 15px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .guest-name {
            font-weight: bold;
            color: #1f2937;
            font-size: 1.05em;
        }

        .guest-count {
            background: #be123c;
            color: white;
            padding: 2px 8px;
            border-radius: 10px;
            font-size: 0.85em;
            min-width: 35px;
            text-align: center;
        }

        .empty-notice {
            color: #9ca3af;
            font-style: italic;
            padding: 10px 20px;
        }

        .summary {
            background: #fef2f2;
            border: 2px solid #fecdd3;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 30px;
            display: flex;
            justify-content: center;
            gap: 50px;
            flex-wrap: wrap;
        }

        .summary-item {
            text-align: center;
        }

        .summary-number {
            font-size: 2.5em;
            font-weight: bold;
            color: #be123c;
        }

        .summary-label {
            color: #666;
            font-size: 0.95em;
        }

        .footer {
            text-align: center;
            margin-top: 50px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #999;
            font-size: 0.9em;
        }

        @media print {
            .print-btn {
                display: none;
            }

            body {
                padding: 0;
            }

            .header h1 {
                font-size: 2em;
            }

            .table-header {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }

            .guest-card {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
        }

        @media (max-width: 768px) {
            .guest-grid {
                grid-template-columns: repeat(2, 1fr);
            }

            .header h1 {
                font-size: 1.8em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>💒 婚礼宾客指引名单</h1>
            <p>请找到您的姓名，查看对应的桌号就座</p>
            <button class="print-btn" onclick="window.print()">🖨️ 打印名单</button>
        </div>

        <div class="summary">
            <div class="summary-item">
                <div class="summary-number" id="totalTables">0</div>
                <div class="summary-label">已安排桌数</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" id="totalGuests">0</div>
                <div class="summary-label">宾客组数</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" id="totalPeople">0</div>
                <div class="summary-label">总人数</div>
            </div>
        </div>

        <div id="guestList">
            <!-- 动态生成内容 -->
        </div>

        <div class="footer">
            <p>🎊 祝您用餐愉快 🎊</p>
        </div>
    </div>

    <script>
        // 桌号对应的显示类别
        const TABLE_CATEGORIES = '''

    # 添加桌号类别
    html_template += json.dumps(TABLE_CATEGORIES, ensure_ascii=False, indent=8)

    html_template += ''';

        // Excel数据
        const tableData = '''

    # 添加桌号数据
    html_template += generate_js_data(table_data)

    html_template += ''';

        // 生成宾客名单
        function generateGuestList() {
            const container = document.getElementById('guestList');
            let html = '';

            // 按桌号排序
            const tableNumbers = Object.keys(tableData).map(Number).sort((a, b) => a - b);

            // 统计数据
            let totalGuests = 0;
            let totalPeople = 0;

            tableNumbers.forEach(tableNum => {
                const data = tableData[tableNum];
                const category = TABLE_CATEGORIES[tableNum] || '';

                totalGuests += data.total_groups;
                totalPeople += data.total_people;

                html += `
                    <div class="table-section">
                        <div class="table-header">
                            <span>第 ${tableNum} 桌</span>
                            <span class="table-category">${category} · 共${data.total_people}人</span>
                        </div>
                        <div class="guest-grid">
                `;

                data.guests.forEach(guest => {
                    html += `
                        <div class="guest-card">
                            <span class="guest-name">${guest.name}</span>
                            <span class="guest-count">${guest.count}人</span>
                        </div>
                    `;
                });

                html += `
                        </div>
                    </div>
                `;
            });

            container.innerHTML = html;

            // 更新统计
            document.getElementById('totalTables').textContent = tableNumbers.length;
            document.getElementById('totalGuests').textContent = totalGuests;
            document.getElementById('totalPeople').textContent = totalPeople;
        }

        // 初始化
        generateGuestList();
    </script>
</body>
</html>'''

    return html_template

def main():
    print("=" * 50)
    print("💒 婚礼座次表一键更新工具")
    print("=" * 50)

    # 文件路径
    excel_path = "婚礼座次.xlsx"
    seating_html_path = "婚礼座次表.html"
    guest_list_html_path = "宾客指引名单.html"

    if not os.path.exists(excel_path):
        print(f"❌ 找不到文件: {excel_path}")
        return

    try:
        # 读取数据
        table_data = read_excel_data(excel_path)

        # 统计信息
        total_guests = sum(d['total_people'] for d in table_data.values())
        total_groups = sum(d['total_groups'] for d in table_data.values())
        print(f"📈 总计: {total_groups} 组宾客, 共 {total_guests} 人")

        # 生成座次布局图HTML
        seating_html = generate_seating_html(table_data)
        with open(seating_html_path, 'w', encoding='utf-8') as f:
            f.write(seating_html)
        print(f"✅ 已生成: {seating_html_path}")

        # 生成宾客指引名单HTML
        guest_list_html = generate_guest_list_html(table_data)
        with open(guest_list_html_path, 'w', encoding='utf-8') as f:
            f.write(guest_list_html)
        print(f"✅ 已生成: {guest_list_html_path}")

        print("=" * 50)
        print("🎉 更新完成！请在浏览器中打开HTML文件查看")

    except Exception as e:
        print(f"❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
