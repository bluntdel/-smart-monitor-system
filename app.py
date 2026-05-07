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

    # 创建 model_config 表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS model_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL UNIQUE,
            field1 TEXT, field2 TEXT, field3 TEXT, field4 TEXT, field5 TEXT,
            field6 TEXT, field7 TEXT, field8 TEXT, field9 TEXT, field10 TEXT,
            field11 TEXT, field12 TEXT, field13 TEXT, field14 TEXT, field15 TEXT,
            field16 TEXT, field17 TEXT, field18 TEXT, field19 TEXT, field20 TEXT
        )
    ''')

    # 创建 model_data 表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS model_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL,
            field1 TEXT, field2 TEXT, field3 TEXT, field4 TEXT, field5 TEXT,
            field6 TEXT, field7 TEXT, field8 TEXT, field9 TEXT, field10 TEXT,
            field11 TEXT, field12 TEXT, field13 TEXT, field14 TEXT, field15 TEXT,
            field16 TEXT, field17 TEXT, field18 TEXT, field19 TEXT, field20 TEXT,
            create_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 创建 model_type 表（如果使用）
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

        first_row = data[0]
        model_name = get_val(data[1], 0)

        conn = get_db()
        cursor = conn.cursor()

        if model_name:
            cursor.execute("SELECT id FROM model_config WHERE model_name = ?", (model_name,))
            if not cursor.fetchone():
                cursor.execute('''
                    INSERT INTO model_config (
                        model_name,
                        field1,field2,field3,field4,field5,
                        field6,field7,field8,field9,field10,
                        field11,field12,field13,field14,field15,
                        field16,field17,field18,field19,field20
                    ) VALUES (
                        ?,?,?,?,?,?,
                        ?,?,?,?,?,
                        ?,?,?,?,?,
                        ?,?,?,?,?
                    )
                ''', (
                    model_name,
                    get_val(first_row, 1), get_val(first_row, 2), get_val(first_row, 3), get_val(first_row, 4),
                    get_val(first_row, 5),
                    get_val(first_row, 6), get_val(first_row, 7), get_val(first_row, 8), get_val(first_row, 9),
                    get_val(first_row, 10),
                    get_val(first_row, 11), get_val(first_row, 12), get_val(first_row, 13), get_val(first_row, 14),
                    get_val(first_row, 15),
                    get_val(first_row, 16), get_val(first_row, 17), get_val(first_row, 18), get_val(first_row, 19),
                    get_val(first_row, 20)
                ))

        for row in data[1:]:
            cursor.execute('''
                INSERT INTO model_data (
                    model_name,
                    field1,field2,field3,field4,field5,
                    field6,field7,field8,field9,field10,
                    field11,field12,field13,field14,field15,
                    field16,field17,field18,field19,field20
                ) VALUES (
                    ?,?,?,?,?,?,
                    ?,?,?,?,?,
                    ?,?,?,?,?,
                    ?,?,?,?,?
                )
            ''', (
                get_val(row, 0),
                get_val(row, 1), get_val(row, 2), get_val(row, 3), get_val(row, 4), get_val(row, 5),
                get_val(row, 6), get_val(row, 7), get_val(row, 8), get_val(row, 9), get_val(row, 10),
                get_val(row, 11), get_val(row, 12), get_val(row, 13), get_val(row, 14), get_val(row, 15),
                get_val(row, 16), get_val(row, 17), get_val(row, 18), get_val(row, 19), get_val(row, 20)
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


def get_chart_data(sql):
    """通用方法：执行 SQL 并返回 JSON 化的 {model_name, count} 列表"""
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
        data = [{"model_name": row["model_name"], "count": row["count"]} for row in rows]
        cursor.close()
        conn.close()
        return jsonify(data)
    except Exception as e:
        print("查询错误：", e)
        return jsonify([])


@app.route('/api/chart-data1')
def api_chart_data1():
    sql = "SELECT t1.model_name, COUNT(t1.model_name) AS count FROM model_data t1 LEFT JOIN model_type t2 ON t1.model_name = t2.model_name WHERE t2.type_code = 1 GROUP BY t1.model_name"
    return get_chart_data(sql)


@app.route('/api/chart-data2')
def api_chart_data2():
    sql = "SELECT t1.model_name, COUNT(t1.model_name) AS count FROM model_data t1 LEFT JOIN model_type t2 ON t1.model_name = t2.model_name WHERE t2.type_code = 2 GROUP BY t1.model_name"
    return get_chart_data(sql)


@app.route('/api/chart-data3')
def api_chart_data3():
    sql = "SELECT t1.model_name, COUNT(t1.model_name) AS count FROM model_data t1 LEFT JOIN model_type t2 ON t1.model_name = t2.model_name WHERE t2.type_code = 3 GROUP BY t1.model_name"
    return get_chart_data(sql)


@app.route('/api/chart-data4')
def api_chart_data4():
    sql = "SELECT t1.model_name, COUNT(t1.model_name) AS count FROM model_data t1 LEFT JOIN model_type t2 ON t1.model_name = t2.model_name WHERE t2.type_code = 4 GROUP BY t1.model_name"
    return get_chart_data(sql)


@app.route('/api/chart-data5')
def api_chart_data5():
    sql = "SELECT t1.model_name, COUNT(t1.model_name) AS count FROM model_data t1 LEFT JOIN model_type t2 ON t1.model_name = t2.model_name WHERE t2.type_code = 5 GROUP BY t1.model_name"
    return get_chart_data(sql)


@app.route('/api/chart-data-detail1/<model_name>')
def chart_data_detail1(model_name):
    data2 = []  # config 字段值列表
    data = []  # data 多条记录的值列表
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor2 = conn.cursor()

        # 注意：SQL 占位符改为 ?
        sql = """
            SELECT t1.id, t1.model_name,
                   t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
                   t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
                   t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
                   t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
            FROM model_data t1
            WHERE t1.model_name = ?
        """
        sql2 = """
            SELECT t1.model_name,
                   t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
                   t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
                   t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
                   t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
            FROM model_config t1
            LEFT JOIN model_data t2 ON t1.model_name = t2.model_name
            WHERE t1.model_name = ?
            LIMIT 1
        """

        cursor.execute(sql, (model_name,))
        cursor2.execute(sql2, (model_name,))

        rows = cursor.fetchall()
        rows2 = cursor2.fetchall()

        # 处理 model_data 多条记录
        for row in rows:
            datatmp = []
            for i in range(1, 21):
                field = f"field{i}"
                datatmp.append(row[field] if row[field] is not None else "")
            data.append(datatmp)

        # 处理 model_config 单条记录
        if rows2:
            row2 = rows2[0]
            for i in range(1, 21):
                field = f"field{i}"
                data2.append(row2[field] if row2[field] is not None else "")

        cursor.close()
        cursor2.close()
        conn.close()

    except Exception as e:
        print("查询错误：", e)
        # 确保返回空列表
        data2 = []
        data = []

    res = [data2, data]
    return Response(json.dumps(res, ensure_ascii=False), mimetype="application/json")


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)