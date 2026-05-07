from flask import Flask, render_template, request, flash, redirect
import pandas as pd
import pymysql
import os
from dotenv import load_dotenv
from datetime import date
from flask import jsonify

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def get_db():
    return pymysql.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME"),
        charset=os.getenv("DB_CHARSET"),
        cursorclass=pymysql.cursors.DictCursor
    )

def get_val(arr, idx):
    try:
        return str(arr[idx]).strip() if idx < len(arr) else ""
    except:
        return ""

@app.route('/', methods=['GET', 'POST'])
def dashboard_ref():
    # ---------- 保留原有导入逻辑 完全不变 ----------
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

        db = get_db()
        cursor = db.cursor()

        if model_name:
            cursor.execute("SELECT id FROM model_config WHERE model_name = %s", (model_name,))
            if not cursor.fetchone():
                cursor.execute('''
                    INSERT INTO model_config (
                        model_name,
                        field1,field2,field3,field4,field5,
                        field6,field7,field8,field9,field10,
                        field11,field12,field13,field14,field15,
                        field16,field17,field18,field19,field20
                    ) VALUES (
                        %s,%s,%s,%s,%s,%s,
                        %s,%s,%s,%s,%s,
                        %s,%s,%s,%s,%s,
                        %s,%s,%s,%s,%s
                    )
                ''', (
                    model_name,
                    get_val(first_row, 1), get_val(first_row, 2), get_val(first_row, 3), get_val(first_row, 4), get_val(first_row, 5),
                    get_val(first_row, 6), get_val(first_row, 7), get_val(first_row, 8), get_val(first_row, 9), get_val(first_row, 10),
                    get_val(first_row, 11),get_val(first_row, 12),get_val(first_row, 13),get_val(first_row, 14),get_val(first_row, 15),
                    get_val(first_row, 16),get_val(first_row, 17),get_val(first_row, 18),get_val(first_row, 19),get_val(first_row, 20)
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
                    %s,%s,%s,%s,%s,%s,
                    %s,%s,%s,%s,%s,
                    %s,%s,%s,%s,%s,
                    %s,%s,%s,%s,%s
                )
            ''', (
                get_val(row, 0),
                get_val(row, 1), get_val(row, 2), get_val(row, 3), get_val(row, 4), get_val(row, 5),
                get_val(row, 6), get_val(row, 7), get_val(row, 8), get_val(row, 9), get_val(row, 10),
                get_val(row, 11),get_val(row, 12),get_val(row, 13),get_val(row, 14),get_val(row, 15),
                get_val(row, 16),get_val(row, 17),get_val(row, 18),get_val(row, 19),get_val(row, 20)
            ))

        db.commit()
        cursor.close()
        db.close()

        flash(f"导入成功！模型名称：{model_name}", "success")
        return redirect('/')

    # ---------- 新增：驾驶舱统计数据 ----------
    db = get_db()
    cursor = db.cursor()

    # 1. 全局概览
    cursor.execute("SELECT COUNT(*) AS cnt FROM model_config")
    totalModel = cursor.fetchone()['cnt']

    cursor.execute("SELECT COUNT(*) AS cnt FROM model_data")
    totalData = cursor.fetchone()['cnt']

    cursor.execute("SELECT COUNT(DISTINCT model_name) AS cnt FROM model_data")
    totalOrg = cursor.fetchone()['cnt']

    today = date.today()
    cursor.execute("SELECT COUNT(*) AS cnt FROM model_data WHERE DATE(create_time)=%s", (today,))
    todayImport = cursor.fetchone()['cnt']

    # 2. 五大模型统计（取数据量前5）
    cursor.execute("""
        SELECT model_name, COUNT(*) AS cnt 
        FROM model_data 
        GROUP BY model_name 
        ORDER BY cnt DESC 
        LIMIT 5
    """)
    top5Models = cursor.fetchall()

    # 3. 机构汇总（这里用model_name模拟机构）
    cursor.execute("""
        SELECT model_name, COUNT(*) AS count 
        FROM model_data 
        GROUP BY model_name 
        ORDER BY count DESC
    """)
    orgData = cursor.fetchall()

    cursor.close()
    db.close()

    return render_template("dashboard.html",
                           totalModel=totalModel,
                           totalData=totalData,
                           totalOrg=totalOrg,
                           todayImport=todayImport,
                           top5Models=top5Models,
                           orgData=orgData)

@app.route('/api/chart-data')
def api_chart_data():
    try:
        db = get_db()
        cursor = db.cursor()
        # 🔥 关键：SQL 别名必须和后面取的一致
        cursor.execute("SELECT model_name, COUNT(*) as count FROM model_data GROUP BY model_name")
        rows = cursor.fetchall()

        # 正确取值方式
        data = []
        for row in rows:
            data.append({
                "model_name": row['model_name'],   # 第一列：model_name
                "count": row['count']         # 第二列：count
            })

    except Exception as e:
        print("查询错误：", e)
        data = []

    return jsonify(data)





if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)