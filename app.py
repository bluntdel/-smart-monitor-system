from flask import Flask, render_template, request, flash, redirect, jsonify, Response,url_for,  send_file
import pandas as pd
import sqlite3
import os
import json
from dotenv import load_dotenv
from datetime import date

load_dotenv()

app = Flask(__name__)
app.secret_key = 'your-secret-key'

# 必须配置上传文件夹
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# SQLite 数据库文件路径（从环境变量读取，默认为 app.db）
DB_PATH = os.getenv("DB_PATH", "app.db")


def get_db():
    """返回 SQLite 数据库连接，设置 row_factory 使返回字典风格的行"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row  # 允许通过列名访问，如 row['column_name']
    return conn


def init_db():
    """初始化数据库表（如果不存在）"""
    conn = get_db()
    cursor = conn.cursor()

    # 创建 model_config 表（新增 jgmc, jgbm, sjrq）
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS model_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL UNIQUE,
            jgmc TEXT,
            jgbm TEXT,
            sjrq TEXT,
            field1 TEXT, field2 TEXT, field3 TEXT, field4 TEXT, field5 TEXT,
            field6 TEXT, field7 TEXT, field8 TEXT, field9 TEXT, field10 TEXT,
            field11 TEXT, field12 TEXT, field13 TEXT, field14 TEXT, field15 TEXT,
            field16 TEXT, field17 TEXT, field18 TEXT, field19 TEXT, field20 TEXT
        )
    ''')

    # 创建 model_data 表（新增 jgmc, jgbm, sjrq）
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS model_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL,
            jgmc TEXT,
            jgbm TEXT,
            sjrq TEXT,
            field1 TEXT, field2 TEXT, field3 TEXT, field4 TEXT, field5 TEXT,
            field6 TEXT, field7 TEXT, field8 TEXT, field9 TEXT, field10 TEXT,
            field11 TEXT, field12 TEXT, field13 TEXT, field14 TEXT, field15 TEXT,
            field16 TEXT, field17 TEXT, field18 TEXT, field19 TEXT, field20 TEXT,
            create_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 创建 model_type 表（不变）
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS model_type (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL,
            type_code INTEGER
        )
    ''')

    conn.commit()
    conn.close()

# 应用启动时初始化数据库
with app.app_context():
    init_db()


def get_val(arr, idx):
    try:
        return str(arr[idx]).strip() if idx < len(arr) else ""
    except:
        return ""


# ---------- 独立导入接口 ----------
@app.route('/import', methods=['POST'])
def import_excel():
    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        flash('请上传正确的 Excel 文件', 'danger')
        return redirect(url_for('dashboard'))

    # 使用 app.config['UPLOAD_FOLDER']
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    df = pd.read_excel(file_path, engine='openpyxl', header=None)
    data = df.values.tolist()

    if len(data) < 2:
        flash('文件至少需要两行数据', 'danger')
        return redirect(url_for('dashboard'))

    first_row = data[0]
    model_name = get_val(data[1], 0)

    # 根据 Excel 列映射：第2列 -> jgbm，第3列 -> jgmc
    jgbm_title = get_val(first_row, 1)
    jgmc_title = get_val(first_row, 2)
    sjrq_title = get_val(first_row, 3)

    field_titles = [get_val(first_row, i) for i in range(4, 24)]

    conn = get_db()
    cursor = conn.cursor()

    if model_name:
        cursor.execute("SELECT id FROM model_config WHERE model_name = ?", (model_name,))
        if not cursor.fetchone():
            cursor.execute('''
                INSERT INTO model_config (
                    model_name, jgbm, jgmc, sjrq,
                    field1, field2, field3, field4, field5,
                    field6, field7, field8, field9, field10,
                    field11, field12, field13, field14, field15,
                    field16, field17, field18, field19, field20
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                model_name,
                jgbm_title, jgmc_title, sjrq_title,
                *field_titles
            ))

    for row in data[1:]:
        jgbm_val = get_val(row, 1)
        jgmc_val = get_val(row, 2)
        sjrq_val = get_val(row, 3)

        cursor.execute('''
            INSERT INTO model_data (
                model_name, jgbm, jgmc, sjrq,
                field1, field2, field3, field4, field5,
                field6, field7, field8, field9, field10,
                field11, field12, field13, field14, field15,
                field16, field17, field18, field19, field20
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            get_val(row, 0),
            jgbm_val,
            jgmc_val,
            sjrq_val,
            get_val(row, 4), get_val(row, 5), get_val(row, 6), get_val(row, 7), get_val(row, 8),
            get_val(row, 9), get_val(row, 10), get_val(row, 11), get_val(row, 12), get_val(row, 13),
            get_val(row, 14), get_val(row, 15), get_val(row, 16), get_val(row, 17), get_val(row, 18),
            get_val(row, 19), get_val(row, 20), get_val(row, 21), get_val(row, 22), get_val(row, 23)
        ))

    conn.commit()
    cursor.close()
    conn.close()

    flash(f"导入成功！模型名称：{model_name}", "success")
    return redirect(url_for('dashboard'))

# ---------- 驾驶舱页面（仅 GET） ----------
@app.route('/')
def dashboard():
    """驾驶舱首页，展示统计图表"""
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) AS cnt FROM model_config")
    totalModel = cursor.fetchone()['cnt']

    cursor.execute("SELECT COUNT(*) AS cnt FROM model_data")
    totalData = cursor.fetchone()['cnt']

    cursor.execute("SELECT COUNT(DISTINCT model_name) AS cnt FROM model_data")
    totalOrg = cursor.fetchone()['cnt']

    today = date.today()
    cursor.execute("SELECT COUNT(*) AS cnt FROM model_data WHERE DATE(create_time)=?", (today,))
    todayImport = cursor.fetchone()['cnt']

    cursor.execute("""
        SELECT model_name, COUNT(*) AS cnt 
        FROM model_data 
        GROUP BY model_name 
        ORDER BY cnt DESC 
        LIMIT 5
    """)
    top5Models = cursor.fetchall()

    cursor.execute("""
        SELECT model_name, COUNT(*) AS count 
        FROM model_data 
        GROUP BY model_name 
        ORDER BY count DESC
    """)
    orgData = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template("index.html",
                           totalModel=totalModel,
                           totalData=totalData,
                           totalOrg=totalOrg,
                           todayImport=todayImport,
                           top5Models=top5Models,
                           orgData=orgData)


@app.route('/api/chart-data-all')
def api_chart_data():
    month = request.args.get('month')
    if not month:
        return jsonify({"error": "缺少必传参数：month!"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 1. 查询各类型在指定月份的实际统计数量
        sql_stats = """
            SELECT mt.type_code, COUNT(*) AS count
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE SUBSTR(md.sjrq, 1, 6) = ?
            GROUP BY mt.type_code
        """
        cursor.execute(sql_stats, (month,))
        rows = cursor.fetchall()
        stats_map = {row["type_code"]: row["count"] for row in rows}

        # 2. 获取所有类型及其描述（从 model_type 表查询去重）
        sql_types = """
            SELECT DISTINCT type_code, type_des
            FROM model_type
            WHERE type_code IN (1,2,3,4,5)
            ORDER BY type_code
        """
        cursor.execute(sql_types)
        type_rows = cursor.fetchall()
        # 如果表没有 type_des 字段，或者查询为空，则使用默认映射
        if not type_rows:
            # 默认描述映射
            type_desc_map = {1: "类型1", 2: "类型2", 3: "类型3", 4: "类型4", 5: "类型5"}
        else:
            type_desc_map = {row["type_code"]: row["type_des"] for row in type_rows}

        # 3. 定义所有类型（1~5），补全缺失类型
        all_types = [1, 2, 3, 4, 5]
        result = []
        for type_code in all_types:
            count = stats_map.get(type_code, 0)
            type_des = type_desc_map.get(type_code, f"类型{type_code}")
            result.append({"type_code": type_code, "typeDes": type_des, "count": count})

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("查询错误：", e)
        return jsonify([])

@app.route('/api/chart-data-monthly-all')
def chart_data_monthly_type():
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate 和 endDate"}), 400

    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 1. 查询类型描述映射
        sql_types = """
            SELECT DISTINCT type_code, type_des
            FROM model_type
            WHERE type_code IN (1,2,3,4,5)
            ORDER BY type_code
        """
        cursor.execute(sql_types)
        type_rows = cursor.fetchall()
        if type_rows:
            type_desc_map = {row["type_code"]: row["type_des"] for row in type_rows}
        else:
            type_desc_map = {1: "类型1", 2: "类型2", 3: "类型3", 4: "类型4", 5: "类型5"}

        # 2. 查询实际数据：按月份和类型分组统计
        sql = """
            SELECT SUBSTR(md.sjrq, 1, 6) AS month, mt.type_code, COUNT(*) AS count
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE md.sjrq BETWEEN ? AND ?
            GROUP BY SUBSTR(md.sjrq, 1, 6), mt.type_code
            ORDER BY month, mt.type_code
        """
        cursor.execute(sql, (start_date, end_date))
        rows = cursor.fetchall()

        # 转换为字典 {(month, type_code): count}
        data_map = {}
        for row in rows:
            key = (row["month"], row["type_code"])
            data_map[key] = row["count"]

        # 3. 生成完整月份列表
        start_year = int(start_date[:4])
        start_month = int(start_date[4:6])
        end_year = int(end_date[:4])
        end_month = int(end_date[4:6])

        months = []
        year, month = start_year, start_month
        while (year < end_year) or (year == end_year and month <= end_month):
            months.append(f"{year}{month:02d}")
            if month == 12:
                month = 1
                year += 1
            else:
                month += 1

        all_types = [1, 2, 3, 4, 5]

        # 4. 构建嵌套结果：每个月份下 data 为对象数组
        result = []
        for month in months:
            data_list = []
            for tc in all_types:
                count = data_map.get((month, tc), 0)
                type_des = type_desc_map.get(tc, f"类型{tc}")
                data_list.append({
                    "type_code": tc,
                    "typeDes": type_des,
                    "count": count
                })
            formatted_month = f"{month[:4]}-{month[4:]}"
            result.append({
                "month": formatted_month,
                "data": data_list
            })

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("按月类型统计查询错误：", e)
        return jsonify({"error": str(e)}), 500


@app.route('/api/chart-data')
def api_chart_data1():
    type_code = request.args.get('typeCode')
    month = request.args.get('month')
    if not type_code:
        return jsonify({"error": "缺少必传参数：typeCode!"}), 400
    if not month:
        return jsonify({"error": "缺少必传参数：month!"}), 400
    sql = """
           SELECT t1.model_name, COUNT(t1.model_name) AS count
           FROM model_data t1
           LEFT JOIN model_type t2 ON t1.model_name = t2.model_name
           WHERE t2.type_code = ? and SUBSTR(t1.sjrq, 1, 6)=?

           GROUP BY t1.model_name
       """
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute(sql, (type_code,month))
        rows = cursor.fetchall()
        data = [{"model_name": row["model_name"], "count": row["count"]} for row in rows]
        cursor.close()
        conn.close()
        return Response(
            json.dumps(data, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("查询错误：", e)
        return jsonify([])


@app.route('/api/chart-data-detail1')
def chart_data_detail1():
    model_name = request.args.get('modelName')
    month = request.args.get('month')
    jgbm = request.args.get('jgbm')          # 新增：机构编码筛选（可选）

    # 参数校验：model_name 和 sjrq 必须存在
    if not model_name or not month:
        return Response(json.dumps({"error": "缺少 modelName 或 month 参数", "code": 400}, ensure_ascii=False),
                    mimetype='application/json')

    data2 = []  # config 字段值列表（包含所有字段）
    data = []   # data 多条记录的值列表（每行所有字段）

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor2 = conn.cursor()

        # ---------- 动态构建 model_data 查询 ----------
        sql = """
            SELECT t1.id, t1.model_name,
                   t1.jgmc, t1.jgbm, t1.sjrq,
                   t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
                   t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
                   t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
                   t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
            FROM model_data t1
            WHERE t1.model_name = ? AND SUBSTR(t1.sjrq, 1, 6)= ?
        """
        params = [model_name, month]

        # 如果提供了机构编码，则添加筛选条件
        if jgbm:
            sql += " AND t1.jgbm = ?"
            params.append(jgbm)

        cursor.execute(sql, params)
        rows = cursor.fetchall()

        # model_config 查询（不需要机构筛选）
        sql2 = """
            SELECT t1.model_name, t1.jgmc, t1.jgbm, t1.sjrq,
                   t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
                   t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
                   t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
                   t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
            FROM model_config t1
            WHERE t1.model_name = ?
            LIMIT 1
        """
        cursor2.execute(sql2, (model_name,))
        rows2 = cursor2.fetchall()

        # 处理 model_data 数据
        for row in rows:
            data.append(list(row))

        # 处理 model_config 数据
        if rows2:
            data2 = list(rows2[0])

        cursor.close()
        cursor2.close()
        conn.close()

    except Exception as e:
        print("查询错误：", e)
        data2 = []
        data = []

    # 返回格式：[config_fields, data_rows]
    res = [data2, data]
    return Response(json.dumps(res, ensure_ascii=False), mimetype="application/json")

from flask import request, send_file, jsonify
import io
import openpyxl
from openpyxl.styles import Font, Alignment

from flask import request, send_file, jsonify
import io
import openpyxl
from openpyxl.styles import Font, Alignment

@app.route('/api/export-excel')
def export_excel():
    model_name = request.args.get('modelName')
    month = request.args.get('month')
    jgbm = request.args.get('jgbm')          # 可选

    if not model_name or not month:
        return jsonify({"error": "缺少 modelName 或 month 参数"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # ---------- 1. 查询 model_config，获取固定字段的值（如 model_name, jgmc, jgbm, sjrq） ----------
        #    这里只取配置用于动态标题，查询顺序不影响最终结果
        sql_config = """
            SELECT model_name, jgbm, jgmc, sjrq,
                   field1, field2, field3, field4, field5,
                   field6, field7, field8, field9, field10,
                   field11, field12, field13, field14, field15,
                   field16, field17, field18, field19, field20
            FROM model_config
            WHERE model_name = ?
            LIMIT 1
        """
        cursor.execute(sql_config, (model_name,))
        config_row = cursor.fetchone()
        if not config_row:
            return jsonify({"error": "未找到模型配置"}), 404

        # 固定字段（Excel中的顺序：模型名称、机构编码、机构名称、数据日期）
        fixed_headers = ['模型名称', '机构编码', '机构名称', '数据日期']

        # 动态字段标题：从 config_row 中提取 field1~field20
        dynamic_raw = config_row[4:]   # 从第5个元素（field1）开始
        dynamic_raw = list(dynamic_raw) + [None] * (20 - len(dynamic_raw))
        dynamic_headers = []
        for idx, val in enumerate(dynamic_raw, start=1):
            if val and str(val).strip():
                dynamic_headers.append(str(val).strip())
            else:
                dynamic_headers.append(f"列{idx}")

        all_headers = fixed_headers + dynamic_headers

        # ---------- 2. 查询 model_data 明细数据 ----------
        # 关键修改：SELECT 子句中固定字段的顺序必须与 all_headers 的固定部分一致
        # 原顺序：model_name, jgmc, jgbm, sjrq
        # 新顺序：model_name, jgbm, jgmc, sjrq
        sql_data = """
            SELECT model_name, jgbm, jgmc, sjrq,
                   field1, field2, field3, field4, field5,
                   field6, field7, field8, field9, field10,
                   field11, field12, field13, field14, field15,
                   field16, field17, field18, field19, field20
            FROM model_data
            WHERE model_name = ? AND SUBSTR(sjrq, 1, 6) = ?
        """
        params = [model_name, month]
        if jgbm:
            sql_data += " AND jgbm = ?"
            params.append(jgbm)

        cursor.execute(sql_data, params)
        data_rows = cursor.fetchall()
        conn.close()

    except Exception as e:
        print("导出失败：", e)
        return jsonify({"error": str(e)}), 500

    # ---------- 3. 生成 Excel ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{model_name}_{month}"

    # 写入表头
    for col_idx, header in enumerate(all_headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # 写入数据行（此时 row_data 的列顺序已与 all_headers 完全一致）
    for row_idx, row_data in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # 自动调整列宽
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        adjusted_width = min(max_len + 2, 40)
        ws.column_dimensions[col_letter].width = adjusted_width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"{model_name}_{month}_{jgbm if jgbm else 'all'}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )
@app.route('/api/org-list')
def org_list():
    model_name = request.args.get('modelName')
    if not model_name:
        return Response(json.dumps({"error": "缺少 modelName 参数","code":400}, ensure_ascii=False),mimetype='application/json')

    try:
        conn = get_db()
        cursor = conn.cursor()
        sql = """
            SELECT DISTINCT jgmc, jgbm
            FROM model_data
            WHERE model_name = ? AND jgmc IS NOT NULL AND jgbm IS NOT NULL
            ORDER BY jgmc
        """
        cursor.execute(sql, (model_name,))
        rows = cursor.fetchall()
        result = [{"jgmc": row["jgmc"], "jgbm": row["jgbm"]} for row in rows]
        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("查询机构列表错误：", e)
        return jsonify({"error": str(e)}), 500


@app.route('/api/model-list')
def model_list():
    """
    获取所有可用的模型名称列表（基于 model_data 表，关联 model_type 支持按 typeCode 过滤）
    可选参数：typeCode - 类型编码
    返回格式：[{"model_name": "模型名称"}, ...]
    """
    type_code = request.args.get('typeCode')

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 查询所有不重复的 model_name，关联 model_type 支持过滤
        sql = """
            SELECT DISTINCT md.model_name
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE md.model_name IS NOT NULL AND md.model_name != ''
        """
        params = []

        if type_code:
            sql += " AND mt.type_code = ?"
            params.append(type_code)

        sql += " ORDER BY md.model_name"

        cursor.execute(sql, params)
        rows = cursor.fetchall()
        result = [{"model_name": row["model_name"]} for row in rows]

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("查询模型名称列表错误：", e)
        return jsonify({"error": str(e)}), 500
@app.route('/api/org-type-stats')
def org_type_stats():
    # 必传参数：startDate, endDate
    month = request.args.get('month')

    # 校验必传参数
    if not month:
        return jsonify({"error": "缺少必传参数：month"}), 400

    # 可选参数：jgmc（机构名称）
    jgmc = request.args.get('jgmc')

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 基础 SQL：统计每个机构的各模型类型记录数，限定日期范围
        sql = """
            SELECT 
                md.jgmc AS 机构名称,
                md.jgbm AS 机构编码,
                COUNT(CASE WHEN mt.type_code = 1 THEN 1 END) AS 模型类型1记录数,
                COUNT(CASE WHEN mt.type_code = 2 THEN 1 END) AS 模型类型2记录数,
                COUNT(CASE WHEN mt.type_code = 3 THEN 1 END) AS 模型类型3记录数,
                COUNT(CASE WHEN mt.type_code = 4 THEN 1 END) AS 模型类型4记录数,
                COUNT(CASE WHEN mt.type_code = 5 THEN 1 END) AS 模型类型5记录数
            FROM model_data md
            LEFT JOIN model_type mt ON md.model_name = mt.model_name
            WHERE SUBSTR(md.sjrq, 1, 6)= ?
        """
        params = [month]

        # 可选：按机构名称筛选
        if jgmc:
            sql += " AND md.jgmc = ?"
            params.append(jgmc)

        sql += " GROUP BY md.jgmc, md.jgbm ORDER BY md.jgmc"

        cursor.execute(sql, params)
        rows = cursor.fetchall()

        # 将 sqlite3.Row 对象转换为字典列表
        result = [dict(row) for row in rows]

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("机构类型统计查询错误：", e)
        return jsonify({"error": str(e)}), 500




@app.route('/api/org-type-stats2')
def org_type_stats2():
    # 必传参数：startDate, endDate
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')

    # 校验必传参数
    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate 和 endDate"}), 400

    # 可选参数：jgmc（机构名称）
    jgmc = request.args.get('jgmc')

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 基础 SQL：统计每个机构的各模型类型记录数，限定日期范围
        sql = """
            SELECT 
                md.jgmc AS 机构名称,
                md.jgbm AS 机构编码,
                COUNT(CASE WHEN mt.type_code = 1 THEN 1 END) AS 模型类型1记录数,
                COUNT(CASE WHEN mt.type_code = 2 THEN 1 END) AS 模型类型2记录数,
                COUNT(CASE WHEN mt.type_code = 3 THEN 1 END) AS 模型类型3记录数,
                COUNT(CASE WHEN mt.type_code = 4 THEN 1 END) AS 模型类型4记录数,
                COUNT(CASE WHEN mt.type_code = 5 THEN 1 END) AS 模型类型5记录数
            FROM model_data md
            LEFT JOIN model_type mt ON md.model_name = mt.model_name
            WHERE md.sjrq BETWEEN ? AND ?
        """
        params = [start_date, end_date]

        # 可选：按机构名称筛选
        if jgmc:
            sql += " AND md.jgmc = ?"
            params.append(jgmc)

        sql += " GROUP BY md.jgmc, md.jgbm ORDER BY md.jgmc"

        cursor.execute(sql, params)
        rows = cursor.fetchall()

        # 将 sqlite3.Row 对象转换为字典列表
        result = [dict(row) for row in rows]

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("机构类型统计查询错误：", e)
        return jsonify({"error": str(e)}), 500


from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta  # 需要安装 python-dateutil


# 如果没有安装，可以使用 pip install python-dateutil
# 或者用下面提供的纯 Python 替代函数
from datetime import datetime
import calendar


@app.route('/api/chart-data-monthly')
def chart_data_monthly():
    type_code = request.args.get('typeCode')   # 可选，若不传则统计所有模型
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')

    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate, endDate"}), 400

    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 1. 获取该类型下的所有模型名称（如果 type_code 为空则取所有模型）
        if type_code:
            sql_models = """
                SELECT DISTINCT md.model_name
                FROM model_data md
                INNER JOIN model_type mt ON md.model_name = mt.model_name
                WHERE mt.type_code = ?
                ORDER BY md.model_name
            """
            cursor.execute(sql_models, (type_code,))
        else:
            sql_models = """
                SELECT DISTINCT md.model_name
                FROM model_data md
                ORDER BY md.model_name
            """
            cursor.execute(sql_models)
        model_rows = cursor.fetchall()
        model_names = [row["model_name"] for row in model_rows]

        if not model_names:
            return Response(json.dumps([], ensure_ascii=False), mimetype='application/json')

        # 2. 查询实际数据：按月、按模型分组统计
        if type_code:
            sql_stats = """
                SELECT 
                    SUBSTR(md.sjrq, 1, 6) AS month,
                    md.model_name,
                    COUNT(*) AS count
                FROM model_data md
                INNER JOIN model_type mt ON md.model_name = mt.model_name
                WHERE mt.type_code = ?
                    AND md.sjrq BETWEEN ? AND ?
                GROUP BY SUBSTR(md.sjrq, 1, 6), md.model_name
                ORDER BY month, md.model_name
            """
            cursor.execute(sql_stats, (type_code, start_date, end_date))
        else:
            sql_stats = """
                SELECT 
                    SUBSTR(md.sjrq, 1, 6) AS month,
                    md.model_name,
                    COUNT(*) AS count
                FROM model_data md
                WHERE md.sjrq BETWEEN ? AND ?
                GROUP BY SUBSTR(md.sjrq, 1, 6), md.model_name
                ORDER BY month, md.model_name
            """
            cursor.execute(sql_stats, (start_date, end_date))

        rows = cursor.fetchall()
        # 构建字典 {(month, model_name): count}
        data_map = {}
        for row in rows:
            key = (row["month"], row["model_name"])
            data_map[key] = row["count"]

        # 3. 生成完整月份列表
        start_year = int(start_date[:4])
        start_month = int(start_date[4:6])
        end_year = int(end_date[:4])
        end_month = int(end_date[4:6])

        months_yyyymm = []
        year, month = start_year, start_month
        while (year < end_year) or (year == end_year and month <= end_month):
            months_yyyymm.append(f"{year}{month:02d}")
            if month == 12:
                month = 1
                year += 1
            else:
                month += 1

        # 4. 构建结果集：每个月份下，每个模型的数量
        result = []
        for yyyymm in months_yyyymm:
            month_str = f"{yyyymm[:4]}-{yyyymm[4:]}"
            model_counts = []
            for model_name in model_names:
                count = data_map.get((yyyymm, model_name), 0)
                model_counts.append({"model_name": model_name, "count": count})
            result.append({"month": month_str, "data": model_counts})

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("按月统计查询错误：", e)
        return jsonify({"error": str(e)}), 500



if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)