from flask import Flask, render_template, request, flash, redirect, jsonify, Response
import pandas as pd
import sqlite3
import os
import json
from dotenv import load_dotenv
from datetime import date

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "default-secret-key")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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


@app.route('/', methods=['GET', 'POST'])
def dashboard_ref():
    # ---------- 原有导入逻辑（占位符修改为 ?）----------
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename.endswith(('.xlsx', '.xls')):
            flash('请上传正确的 Excel 文件', 'danger')
            return redirect('/')

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        df = pd.read_excel(file_path, engine='openpyxl', header=None)
        data = df.values.tolist()

        if len(data) < 2:
            flash('文件至少需要两行数据', 'danger')
            return redirect('/')

        first_row = data[0]  # 表头行
        model_name = get_val(data[1], 0)  # 第二行第一列作为模型名称

        # 从表头行提取三个新字段的标题
        jgmc_title = get_val(first_row, 1)
        jgbm_title = get_val(first_row, 2)
        sjrq_title = get_val(first_row, 3)

        # 提取 field1~field20 的标题（从第4列开始，共20列）
        field_titles = [get_val(first_row, i) for i in range(4, 24)]

        conn = get_db()
        cursor = conn.cursor()

        if model_name:
            cursor.execute("SELECT id FROM model_config WHERE model_name = ?", (model_name,))
            if not cursor.fetchone():
                # 插入 model_config 时包含三个新字段
                cursor.execute('''
                    INSERT INTO model_config (
                        model_name, jgmc, jgbm, sjrq,
                        field1, field2, field3, field4, field5,
                        field6, field7, field8, field9, field10,
                        field11, field12, field13, field14, field15,
                        field16, field17, field18, field19, field20
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    model_name,
                    jgmc_title, jgbm_title, sjrq_title,
                    *field_titles  # Python 3.6+ 支持展开列表
                ))

        # 插入 model_data 数据行
        for row in data[1:]:  # 从第二行开始是数据行
            cursor.execute('''
                INSERT INTO model_data (
                    model_name, jgmc, jgbm, sjrq,
                    field1, field2, field3, field4, field5,
                    field6, field7, field8, field9, field10,
                    field11, field12, field13, field14, field15,
                    field16, field17, field18, field19, field20
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                get_val(row, 0),  # model_name
                get_val(row, 1),  # jgmc
                get_val(row, 2),  # jgbm
                get_val(row, 3),  # sjrq
                get_val(row, 4), get_val(row, 5), get_val(row, 6), get_val(row, 7), get_val(row, 8),
                get_val(row, 9), get_val(row, 10), get_val(row, 11), get_val(row, 12), get_val(row, 13),
                get_val(row, 14), get_val(row, 15), get_val(row, 16), get_val(row, 17), get_val(row, 18),
                get_val(row, 19), get_val(row, 20), get_val(row, 21), get_val(row, 22), get_val(row, 23)
            ))

        conn.commit()
        cursor.close()
        conn.close()

        flash(f"导入成功！模型名称：{model_name}", "success")
        return redirect('/')

    # ---------- 驾驶舱统计数据 ----------
    conn = get_db()
    cursor = conn.cursor()

    # 全局概览
    cursor.execute("SELECT COUNT(*) AS cnt FROM model_config")
    totalModel = cursor.fetchone()['cnt']

    cursor.execute("SELECT COUNT(*) AS cnt FROM model_data")
    totalData = cursor.fetchone()['cnt']

    cursor.execute("SELECT COUNT(DISTINCT model_name) AS cnt FROM model_data")
    totalOrg = cursor.fetchone()['cnt']

    today = date.today()
    cursor.execute("SELECT COUNT(*) AS cnt FROM model_data WHERE DATE(create_time)=?", (today,))
    todayImport = cursor.fetchone()['cnt']

    # 五大模型统计
    cursor.execute("""
        SELECT model_name, COUNT(*) AS cnt 
        FROM model_data 
        GROUP BY model_name 
        ORDER BY cnt DESC 
        LIMIT 5
    """)
    top5Models = cursor.fetchall()

    # 机构汇总
    cursor.execute("""
        SELECT model_name, COUNT(*) AS count 
        FROM model_data 
        GROUP BY model_name 
        ORDER BY count DESC
    """)
    orgData = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template("dashboard.html",
                           totalModel=totalModel,
                           totalData=totalData,
                           totalOrg=totalOrg,
                           todayImport=todayImport,
                           top5Models=top5Models,
                           orgData=orgData)


@app.route('/api/chart-data')
def api_chart_data1():
    type_code = request.args.get('typeCode')
    if not type_code:
        return jsonify({"error": "缺少必传参数：typeCode!"}), 400
    sql = """
           SELECT t1.model_name, COUNT(t1.model_name) AS count
           FROM model_data t1
           LEFT JOIN model_type t2 ON t1.model_name = t2.model_name
           WHERE t2.type_code = ?
           GROUP BY t1.model_name
       """
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute(sql, (type_code,))
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
    sjrq = request.args.get('sjrq')
    jgbm = request.args.get('jgbm')          # 新增：机构编码筛选（可选）

    # 参数校验：model_name 和 sjrq 必须存在
    if not model_name or not sjrq:
        return Response(json.dumps({"error": "缺少 modelName 或 sjrq 参数", "code": 400}, ensure_ascii=False),
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
            WHERE t1.model_name = ? AND t1.sjrq = ?
        """
        params = [model_name, sjrq]

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

@app.route('/api/export-excel')
def export_excel():
    # 接收参数（与原接口一致）
    model_name = request.args.get('modelName')
    sjrq = request.args.get('sjrq')
    jgbm = request.args.get('jgbm')

    if not model_name or not sjrq:
        return jsonify({"error": "缺少 modelName 或 sjrq 参数"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # ---------- 查询 model_config（获取字段顺序/表头） ----------
        sql_config = """
            SELECT t1.model_name, t1.jgmc, t1.jgbm, t1.sjrq,
                   t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
                   t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
                   t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
                   t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
            FROM model_config t1
            WHERE t1.model_name = ?
            LIMIT 1
        """
        cursor.execute(sql_config, (model_name,))
        config_row = cursor.fetchone()
        if not config_row:
            return jsonify({"error": "未找到模型配置"}), 404

        # 获取表头（字段名）
        headers = [desc[0] for desc in cursor.description]

        # ---------- 查询 model_data（明细数据） ----------
        sql_data = """
            SELECT t1.id, t1.model_name, t1.jgmc, t1.jgbm, t1.sjrq,
                   t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
                   t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
                   t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
                   t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
            FROM model_data t1
            WHERE t1.model_name = ? AND t1.sjrq = ?
        """
        params = [model_name, sjrq]
        if jgbm:
            sql_data += " AND t1.jgbm = ?"
            params.append(jgbm)

        cursor.execute(sql_data, params)
        data_rows = cursor.fetchall()
        conn.close()

    except Exception as e:
        print("导出失败：", e)
        return jsonify({"error": str(e)}), 500

    # ---------- 生成 Excel ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{model_name}_{sjrq}"

    # 写入表头（加粗、居中）
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # 写入数据行
    for row_idx, row_data in enumerate(data_rows, start=2):
        for col_idx, val in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)

    # 自动调整列宽（可选）
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        adjusted_width = min(max_len + 2, 30)
        ws.column_dimensions[col_letter].width = adjusted_width

    # 保存到内存字节流
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # 构造下载文件名
    filename = f"{model_name}_{sjrq}_{jgbm if jgbm else 'all'}.xlsx"
    # 使用 send_file 返回
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename   # Flask 2.0+ 使用 download_name；旧版可用 attachment_filename
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


@app.route('/api/org-type-stats')
def org_type_stats():
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
    type_code = request.args.get('typeCode')
    start_date = request.args.get('startDate')  # 格式：yyyymmdd，如 20260101
    end_date = request.args.get('endDate')

    if not type_code or not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：typeCode, startDate, endDate"}), 400

    # 校验日期格式（8位数字）
    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 按月分组统计，关联 model_type 表，过滤 type_code
        # 假设 md.sjrq 存储格式也是 yyyymmdd 字符串，直接字符串比较
        sql = """
            SELECT 
                SUBSTR(md.sjrq, 1, 6) AS month,
                COUNT(*) AS count
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE mt.type_code = ?
                AND md.sjrq BETWEEN ? AND ?
            GROUP BY SUBSTR(md.sjrq, 1, 6)
            ORDER BY month
        """
        cursor.execute(sql, (type_code, start_date, end_date))
        rows = cursor.fetchall()

        # 将查询结果转成字典 {yyyymm: count}
        data_map = {row["month"]: row["count"] for row in rows}

        # 生成从 start_date 到 end_date 的所有月份（yyyymm）
        start_year = int(start_date[:4])
        start_month = int(start_date[4:6])
        end_year = int(end_date[:4])
        end_month = int(end_date[4:6])

        months_yyyymm = []
        year = start_year
        month = start_month
        while (year < end_year) or (year == end_year and month <= end_month):
            months_yyyymm.append(f"{year}{month:02d}")
            # 增加一个月
            if month == 12:
                month = 1
                year += 1
            else:
                month += 1

        # 构建结果集，将 yyyymm 转为 yyyy-mm 格式，并补全0
        result = []
        for yyyymm in months_yyyymm:
            count = data_map.get(yyyymm, 0)
            # 转换为 yyyy-mm 格式便于前端显示
            formatted_month = f"{yyyymm[:4]}-{yyyymm[4:]}"
            result.append({"month": formatted_month, "count": count})

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