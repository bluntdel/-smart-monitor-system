from flask import Flask, render_template, request, flash, redirect, jsonify, Response,url_for,  send_file
import pandas as pd
import sqlite3
import os
import json
from dotenv import load_dotenv
from datetime import date
import io
import openpyxl
from openpyxl.styles import Font, Alignment
from datetime import datetime
from openpyxl.utils import get_column_letter

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
    printf(DB_PATH)
    conn.commit()
    conn.close()

# 应用启动时初始化数据库
# with app.app_context():
#     init_db()


def get_val(arr, idx):
    try:
        return str(arr[idx]).strip() if idx < len(arr) else ""
    except:
        return ""


# ---------- 独立导入接口 ----------
from datetime import datetime
def get_val(arr, idx):
    try:
        val = arr[idx]
        if val is None:
            return ""
        return str(val).strip()
    except (IndexError, TypeError):
        return ""

def is_valid(value):
    """检查值是否有效（非空、非 'None' 字符串）"""
    if value is None:
        return False
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == '' or stripped.lower() == 'none' or stripped.lower() == 'nan':
            return False
    return True

@app.route('/api/import', methods=['POST'])
def import_excel():
    files = request.files.getlist('file')
    if not files:
        flash('请至少选择一个 Excel 文件', 'danger')
        return redirect(url_for('dashboard'))

    total_inserted = 0
    success_count = 0
    errors = []

    # 2. 获取数据库连接（全局）
    conn = get_db()
    cursor = conn.cursor()

    for file in files:
        if not file.filename.endswith(('.xlsx', '.xls')):
            errors.append(f'{file.filename}: 不支持的文件类型，仅支持 .xlsx 或 .xls')
            continue

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        try:
            file.save(file_path)
            # 自动选择引擎（openpyxl 或 xlrd）
            df = pd.read_excel(file_path,engine='openpyxl', header=None)
            data = df.values.tolist()

            if len(data) < 2:
                raise ValueError('文件至少需要两行数据（表头+数据行）')

            first_row = data[0]
            model_name = get_val(data[1], 0)

            # 表头提取
            jgbm_title = get_val(first_row, 1)
            jgmc_title = get_val(first_row, 2)
            sjrq_title = get_val(first_row, 3)
            field_titles = [get_val(first_row, i) for i in range(4, 24)]

            # 插入或忽略 model_config
            if model_name:
                cursor.execute("SELECT id FROM model_config WHERE model_name = ?", (model_name,))
                if not cursor.fetchone():
                    base_fields = ['model_name', 'jgbm', 'jgmc', 'sjrq'] + [f'field{i}' for i in range(1, 21)]
                    base_values = [model_name, jgbm_title, jgmc_title, sjrq_title] + field_titles
                    all_fields = []
                    all_values = []
                    for idx, field in enumerate(base_fields):
                        orig_val = base_values[idx]
                        all_fields.extend([field, f'{field}_des', f'{field}_disable'])
                        all_values.extend([orig_val, orig_val, "1"])
                    placeholders = ','.join(['?'] * len(all_fields))
                    insert_sql = f"INSERT INTO model_config ({','.join(all_fields)}) VALUES ({placeholders})"
                    cursor.execute(insert_sql, all_values)

            # 生成批次号（不含扩展名）
            base_name = os.path.splitext(file.filename)[0]
            batch_no = f"{base_name}_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}"

            # 过滤合格数据行
            valid_rows = []
            for row in data[1:]:
                jgbm_val = get_val(row, 1)
                jgmc_val = get_val(row, 2)
                sjrq_val = get_val(row, 3)
                if is_valid(jgbm_val) and is_valid(jgmc_val) and is_valid(sjrq_val):
                    valid_rows.append(row)

            if not valid_rows:
                raise ValueError('没有符合条件的数据行（机构信息为空）')

            # 插入数据
            for row in valid_rows:
                cursor.execute('''
                    INSERT INTO model_data (
                        model_name, jgbm, jgmc, sjrq, batch_no,
                        field1, field2, field3, field4, field5,
                        field6, field7, field8, field9, field10,
                        field11, field12, field13, field14, field15,
                        field16, field17, field18, field19, field20
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    get_val(row, 0),
                    get_val(row, 1),
                    get_val(row, 2),
                    get_val(row, 3),
                    batch_no,
                    get_val(row, 4), get_val(row, 5), get_val(row, 6), get_val(row, 7), get_val(row, 8),
                    get_val(row, 9), get_val(row, 10), get_val(row, 11), get_val(row, 12), get_val(row, 13),
                    get_val(row, 14), get_val(row, 15), get_val(row, 16), get_val(row, 17), get_val(row, 18),
                    get_val(row, 19), get_val(row, 20), get_val(row, 21), get_val(row, 22), get_val(row, 23)
                ))
                total_inserted += 1

            conn.commit()
            success_count += 1

        except Exception as e:
            conn.rollback()   # 回滚当前文件的修改
            errors.append(f'{file.filename}: {str(e)}')
        finally:
            # 删除临时文件
            if os.path.exists(file_path):
                os.remove(file_path)

    cursor.close()
    conn.close()

    if success_count == 0:
        return jsonify({"code": 500, "msg": f"导入失败: {'; '.join(errors)}"}), 200
    else:
        msg = f'成功导入 {success_count} 个文件，共插入 {total_inserted} 条有效记录'
        if errors:
            msg += f'，部分文件失败: {"；".join(errors)}'
        return Response(json.dumps({"code": 200, "msg": msg}, ensure_ascii=False),
                        mimetype='application/json')
        return jsonify({"code": 200, "msg": msg}), 200


# @app.route('/import', methods=['POST'])
# def import_excel():
#     file = request.files.get('file')
#     if not file or not file.filename.endswith(('.xlsx', '.xls')):
#         flash('请上传正确的 Excel 文件', 'danger')
#         return redirect(url_for('dashboard'))
#
#     file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
#     file.save(file_path)
#
#     df = pd.read_excel(file_path,engine='xlrd', header=None)
#     data = df.values.tolist()
#
#     if len(data) < 2:
#         flash('文件至少需要两行数据', 'danger')
#         return redirect(url_for('dashboard'))
#
#     first_row = data[0]
#     model_name = get_val(data[1], 0)
#
#     # 表头提取（无需过滤）
#     jgbm_title = get_val(first_row, 1)
#     jgmc_title = get_val(first_row, 2)
#     sjrq_title = get_val(first_row, 3)
#     field_titles = [get_val(first_row, i) for i in range(4, 24)]
#
#     conn = get_db()
#     cursor = conn.cursor()
#
#     # 处理 model_config
#     if model_name:
#         cursor.execute("SELECT id FROM model_config WHERE model_name = ?", (model_name,))
#         if not cursor.fetchone():
#             base_fields = ['model_name', 'jgbm', 'jgmc', 'sjrq'] + [f'field{i}' for i in range(1, 21)]
#             base_values = [model_name, jgbm_title, jgmc_title, sjrq_title] + field_titles
#             all_fields = []
#             all_values = []
#             for idx, field in enumerate(base_fields):
#                 original_value = base_values[idx]
#                 all_fields.extend([field, f'{field}_des', f'{field}_disable'])
#                 all_values.extend([original_value, original_value, "1"])
#             placeholders = ','.join(['?'] * len(all_fields))
#             insert_sql = f"INSERT INTO model_config ({','.join(all_fields)}) VALUES ({placeholders})"
#             cursor.execute(insert_sql, all_values)
#
#     # --- 生成批次号 ---
#     base_name = os.path.splitext(file.filename)[0]
#     batch_no = f"{base_name}_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}"
#
#     # --- 过滤数据行，构建 dataFilter ---
#     dataFilter = []
#     for row in data[1:]:
#         jgbm_val = get_val(row, 1)
#         jgmc_val = get_val(row, 2)
#         sjrq_val = get_val(row, 3)
#         if is_valid(jgbm_val) and is_valid(jgmc_val) and is_valid(sjrq_val):
#             dataFilter.append(row)
#
#     # --- 插入 model_data（使用 dataFilter）---
#     for row in dataFilter:
#         cursor.execute('''
#             INSERT INTO model_data (
#                 model_name, jgbm, jgmc, sjrq, batch_no,
#                 field1, field2, field3, field4, field5,
#                 field6, field7, field8, field9, field10,
#                 field11, field12, field13, field14, field15,
#                 field16, field17, field18, field19, field20
#             ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
#         ''', (
#             get_val(row, 0),
#             get_val(row, 1),
#             get_val(row, 2),
#             get_val(row, 3),
#             batch_no,
#             get_val(row, 4), get_val(row, 5), get_val(row, 6), get_val(row, 7), get_val(row, 8),
#             get_val(row, 9), get_val(row, 10), get_val(row, 11), get_val(row, 12), get_val(row, 13),
#             get_val(row, 14), get_val(row, 15), get_val(row, 16), get_val(row, 17), get_val(row, 18),
#             get_val(row, 19), get_val(row, 20), get_val(row, 21), get_val(row, 22), get_val(row, 23)
#         ))
#
#     conn.commit()
#     cursor.close()
#     conn.close()
#
#     flash(f"导入成功！模型名称：{model_name}，共导入 {len(dataFilter)} 条有效数据（已过滤空机构信息）", "success")
#     return redirect(url_for('dashboard'))


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
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate, endDate"}), 400

    # 可选：校验日期格式（8位数字）
    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 1. 统计各类型在指定日期范围内的记录数量
        sql_stats = """
            SELECT mt.type_code, COUNT(*) AS count
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE md.sjrq BETWEEN ? AND ?
            GROUP BY mt.type_code
        """
        cursor.execute(sql_stats, (start_date, end_date))
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
        if not type_rows:
            # 默认描述映射
            type_desc_map = {1: "类型1", 2: "类型2", 3: "类型3", 4: "类型4", 5: "类型5"}
        else:
            type_desc_map = {row["type_code"]: row["type_des"] for row in type_rows}

        # 3. 补全所有类型（1~5）
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

@app.route('/api/chart-org-data')
def api_chart_org_data():
    type_code = request.args.get('typeCode')
    month = request.args.get('month')
    jgmc = request.args.get('jgmc')
    if not type_code:
        return jsonify({"error": "缺少必传参数：typeCode!"}), 400
    if not month:
        return jsonify({"error": "缺少必传参数：month!"}), 400
    if not jgmc:
        return jsonify({"error": "缺少必传参数：jgmc!"}), 400

    sql = """
        SELECT t1.model_name, COUNT(t1.model_name) AS count
        FROM model_data t1
        LEFT JOIN model_type t2 ON t1.model_name = t2.model_name
        WHERE t2.type_code = ? AND SUBSTR(t1.sjrq, 1, 6) = ? AND t1.jgmc = ?
        GROUP BY t1.model_name
    """
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute(sql, (type_code, month, jgmc))
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
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    jgbm = request.args.get('jgbm')          # 机构编码（可选）

    if not type_code:
        return jsonify({"error": "缺少必传参数：typeCode!"}), 400
    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate, endDate"}), 400
    # 可选：校验日期格式（8位数字）
    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

    # 基础 SQL
    sql = """
        SELECT t1.model_name, COUNT(t1.model_name) AS count
        FROM model_data t1
        INNER JOIN model_type t2 ON t1.model_name = t2.model_name
        WHERE t2.type_code = ? AND t1.sjrq BETWEEN ? AND ?
    """
    params = [type_code, start_date, end_date]

    # 可选机构筛选
    if jgbm:
        sql += " AND t1.jgbm = ?"
        params.append(jgbm)

    sql += " GROUP BY t1.model_name"

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute(sql, params)
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
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    jgbm = request.args.get('jgbm')
    # 分页参数
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    if page < 1:
        page = 1
    if per_page < 1:
        per_page = 20
    offset = (page - 1) * per_page

    if not model_name or not start_date or not end_date:
        return Response(json.dumps({"error": "缺少 modelName 或 startDate/endDate 参数", "code": 400}, ensure_ascii=False),
                        mimetype='application/json')

    # 可选：校验日期格式（8位数字）
    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return Response(json.dumps({"error": "日期格式必须为 yyyymmdd，如 20260101", "code": 400}, ensure_ascii=False),
                        mimetype='application/json')

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 基础字段列表（顺序决定了展示顺序）
        base_fields = ['model_name', 'jgmc', 'jgbm', 'sjrq'] + [f'field{i}' for i in range(1, 21)]

        # 查询 model_config，获取所有 _des 和 _disable
        select_config = []
        for f in base_fields:
            select_config.append(f + '_des')
            select_config.append(f + '_disable')
        sql_config = f"SELECT {','.join(select_config)} FROM model_config WHERE model_name = ? LIMIT 1"
        cursor.execute(sql_config, (model_name,))
        config_row = cursor.fetchone()
        if not config_row:
            return Response(json.dumps({"error": "未找到模型配置", "code": 404}, ensure_ascii=False),
                            mimetype='application/json')

        # 构建标题列表和需要查询的字段列表
        headers = []
        data_fields = []
        for f in base_fields:
            disable_val = config_row[f + '_disable']
            if disable_val == "1":
                title = config_row[f + '_des'] or f
                headers.append(title)
                data_fields.append(f)

        if not data_fields:
            return Response(json.dumps({"headers": [], "data": [], "total": 0, "page": page, "per_page": per_page}, ensure_ascii=False),
                            mimetype='application/json')

        # 构建 WHERE 条件（日期范围）
        where_clause = "model_name = ? AND sjrq BETWEEN ? AND ?"
        params = [model_name, start_date, end_date]
        if jgbm:
            where_clause += " AND jgbm = ?"
            params.append(jgbm)

        # 1. 查询总记录数
        count_sql = f"SELECT COUNT(*) AS total FROM model_data WHERE {where_clause}"
        cursor.execute(count_sql, params)
        total = cursor.fetchone()['total']

        # 2. 分页查询数据
        select_data = ','.join(data_fields)
        data_sql = f"""
            SELECT {select_data}
            FROM model_data
            WHERE {where_clause}
            LIMIT ? OFFSET ?
        """
        data_params = params + [per_page, offset]
        cursor.execute(data_sql, data_params)
        rows = cursor.fetchall()
        data = [list(row) for row in rows]

        cursor.close()
        conn.close()

        result = {
            "headers": headers,
            "data": data,
            "total": total,
            "page": page,
            "per_page": per_page
        }
        return Response(json.dumps(result, ensure_ascii=False), mimetype="application/json")
    except Exception as e:
        print("查询错误：", e)
        return Response(json.dumps({"error": str(e), "code": 500}, ensure_ascii=False),
                        mimetype='application/json')

# @app.route('/api/chart-data-detail1')
# def chart_data_detail1():
#     model_name = request.args.get('modelName')
#     month = request.args.get('month')
#     jgbm = request.args.get('jgbm')          # 新增：机构编码筛选（可选）
#
#     # 参数校验：model_name 和 sjrq 必须存在
#     if not model_name or not month:
#         return Response(json.dumps({"error": "缺少 modelName 或 month 参数", "code": 400}, ensure_ascii=False),
#                     mimetype='application/json')
#
#     data2 = []  # config 字段值列表（包含所有字段）
#     data = []   # data 多条记录的值列表（每行所有字段）
#
#     try:
#         conn = get_db()
#         cursor = conn.cursor()
#         cursor2 = conn.cursor()
#
#         # ---------- 动态构建 model_data 查询 ----------
#         sql = """
#             SELECT t1.id, t1.model_name,
#                    t1.jgmc, t1.jgbm, t1.sjrq,
#                    t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
#                    t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
#                    t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
#                    t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
#             FROM model_data t1
#             WHERE t1.model_name = ? AND SUBSTR(t1.sjrq, 1, 6)= ?
#         """
#         params = [model_name, month]
#
#         # 如果提供了机构编码，则添加筛选条件
#         if jgbm:
#             sql += " AND t1.jgbm = ?"
#             params.append(jgbm)
#
#         cursor.execute(sql, params)
#         rows = cursor.fetchall()
#
#         # model_config 查询（不需要机构筛选）
#         sql2 = """
#             SELECT t1.model_name, t1.jgmc, t1.jgbm, t1.sjrq,
#                    t1.field1, t1.field2, t1.field3, t1.field4, t1.field5,
#                    t1.field6, t1.field7, t1.field8, t1.field9, t1.field10,
#                    t1.field11, t1.field12, t1.field13, t1.field14, t1.field15,
#                    t1.field16, t1.field17, t1.field18, t1.field19, t1.field20
#             FROM model_config t1
#             WHERE t1.model_name = ?
#             LIMIT 1
#         """
#         cursor2.execute(sql2, (model_name,))
#         rows2 = cursor2.fetchall()
#
#         # 处理 model_data 数据（移除第一个id字段）
#         for row in rows:
#             row_list = list(row)
#             row_list.pop(0)  # 删除第一个元素（id字段）
#             data.append(row_list)
#
#         # 处理 model_config 数据
#         if rows2:
#             data2 = list(rows2[0])
#
#         cursor.close()
#         cursor2.close()
#         conn.close()
#
#     except Exception as e:
#         print("查询错误：", e)
#         data2 = []
#         data = []
#
#     # 返回格式：[config_fields, data_rows]
#     res = [data2, data]
#     return Response(json.dumps(res, ensure_ascii=False), mimetype="application/json")




# @app.route('/api/export-excel')
# def export_excel():
#     model_name = request.args.get('modelName')
#     month = request.args.get('month')
#     jgbm = request.args.get('jgbm')          # 可选
#
#     if not model_name or not month:
#         return jsonify({"error": "缺少 modelName 或 month 参数"}), 400
#
#     try:
#         conn = get_db()
#         cursor = conn.cursor()
#
#         # ---------- 1. 查询 model_config，获取固定字段的值（如 model_name, jgmc, jgbm, sjrq） ----------
#         #    这里只取配置用于动态标题，查询顺序不影响最终结果
#         sql_config = """
#             SELECT model_name, jgbm, jgmc, sjrq,
#                    field1, field2, field3, field4, field5,
#                    field6, field7, field8, field9, field10,
#                    field11, field12, field13, field14, field15,
#                    field16, field17, field18, field19, field20
#             FROM model_config
#             WHERE model_name = ?
#             LIMIT 1
#         """
#         cursor.execute(sql_config, (model_name,))
#         config_row = cursor.fetchone()
#         if not config_row:
#             return jsonify({"error": "未找到模型配置"}), 404
#
#         # 固定字段（Excel中的顺序：模型名称、机构编码、机构名称、数据日期）
#         fixed_headers = ['模型名称', '机构编码', '机构名称', '数据日期']
#
#         # 动态字段标题：从 config_row 中提取 field1~field20
#         dynamic_raw = config_row[4:]   # 从第5个元素（field1）开始
#         dynamic_raw = list(dynamic_raw) + [None] * (20 - len(dynamic_raw))
#         dynamic_headers = []
#         for idx, val in enumerate(dynamic_raw, start=1):
#             if val and str(val).strip():
#                 dynamic_headers.append(str(val).strip())
#             else:
#                 dynamic_headers.append(f"列{idx}")
#
#         all_headers = fixed_headers + dynamic_headers
#
#         # ---------- 2. 查询 model_data 明细数据 ----------
#         # 关键修改：SELECT 子句中固定字段的顺序必须与 all_headers 的固定部分一致
#         # 原顺序：model_name, jgmc, jgbm, sjrq
#         # 新顺序：model_name, jgbm, jgmc, sjrq
#         sql_data = """
#             SELECT model_name, jgbm, jgmc, sjrq,
#                    field1, field2, field3, field4, field5,
#                    field6, field7, field8, field9, field10,
#                    field11, field12, field13, field14, field15,
#                    field16, field17, field18, field19, field20
#             FROM model_data
#             WHERE model_name = ? AND SUBSTR(sjrq, 1, 6) = ?
#         """
#         params = [model_name, month]
#         if jgbm:
#             sql_data += " AND jgbm = ?"
#             params.append(jgbm)
#
#         cursor.execute(sql_data, params)
#         data_rows = cursor.fetchall()
#         conn.close()
#
#     except Exception as e:
#         print("导出失败：", e)
#         return jsonify({"error": str(e)}), 500
#
#     # ---------- 3. 生成 Excel ----------
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = f"{model_name}_{month}"
#
#     # 写入表头
#     for col_idx, header in enumerate(all_headers, start=1):
#         cell = ws.cell(row=1, column=col_idx, value=header)
#         cell.font = Font(bold=True)
#         cell.alignment = Alignment(horizontal='center')
#
#     # 写入数据行（此时 row_data 的列顺序已与 all_headers 完全一致）
#     for row_idx, row_data in enumerate(data_rows, start=2):
#         for col_idx, value in enumerate(row_data, start=1):
#             ws.cell(row=row_idx, column=col_idx, value=value)
#
#     # 自动调整列宽
#     for col in ws.columns:
#         max_len = 0
#         col_letter = col[0].column_letter
#         for cell in col:
#             if cell.value:
#                 try:
#                     max_len = max(max_len, len(str(cell.value)))
#                 except:
#                     pass
#         adjusted_width = min(max_len + 2, 40)
#         ws.column_dimensions[col_letter].width = adjusted_width
#
#     output = io.BytesIO()
#     wb.save(output)
#     output.seek(0)
#
#     filename = f"{model_name}_{month}_{jgbm if jgbm else 'all'}.xlsx"
#     return send_file(
#         output,
#         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
#         as_attachment=True,
#         download_name=filename
#     )

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

        # 1. 基础字段顺序（决定最终 Excel 列顺序）
        base_fields = ['model_name', 'jgmc', 'jgbm', 'sjrq'] + [f'field{i}' for i in range(1, 21)]

        # 2. 查询 model_config，获取所有字段的 _des 和 _disable
        select_config = []
        for f in base_fields:
            select_config.append(f + '_des')
            select_config.append(f + '_disable')
        # 注意：不查询原始值，只取 _des, _disable
        sql_config = f"SELECT {','.join(select_config)} FROM model_config WHERE model_name = ? LIMIT 1"
        cursor.execute(sql_config, (model_name,))
        config_row = cursor.fetchone()
        if not config_row:
            return jsonify({"error": "未找到模型配置"}), 404

        # 3. 构建表头和数据字段列表
        headers = []
        data_fields = []
        for f in base_fields:
            disable_val = config_row[f + '_disable']
            if disable_val == "1":   # 只有启用状态才导出
                title = config_row[f + '_des'] or f
                headers.append(title)
                data_fields.append(f)

        if not data_fields:
            # 没有可导出字段，返回空 Excel 或提示错误
            return jsonify({"error": "无可导出的字段，请检查模型配置"}), 400

        # 4. 查询 model_data（不分页，导出所有符合条件的记录）
        where_clause = "model_name = ? AND SUBSTR(sjrq, 1, 6) = ?"
        params = [model_name, month]
        if jgbm:
            where_clause += " AND jgbm = ?"
            params.append(jgbm)

        select_data = ','.join(data_fields)
        data_sql = f"""
            SELECT {select_data}
            FROM model_data
            WHERE {where_clause}
        """
        cursor.execute(data_sql, params)
        rows = cursor.fetchall()
        data = [list(row) for row in rows]

        cursor.close()
        conn.close()

        # 5. 生成 Excel 文件
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"{model_name}_{month}"

        # 写入表头
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        # 写入数据行
        for row_idx, row_data in enumerate(data, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        # 自动调整列宽
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
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

        filename = f"{model_name}_{month}{'_' + jgbm if jgbm else ''}.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        print("导出失败：", e)
        return jsonify({"error": str(e)}), 500


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
            lEFT JOIN model_type mt ON md.model_name = mt.model_name
            WHERE md.model_name IS NOT NULL AND md.model_name != ''
        """
        params = []

        if type_code:
            sql += " AND mt.type_code = ?"
            params.append(type_code)

        sql += " ORDER BY md.create_time"

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
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')

    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate, endDate"}), 400

    # 可选参数：jgmc（机构名称）
    jgmc = request.args.get('jgmc')

    # 日期格式校验（8位数字 yyyymmdd）
    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

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
    type_code = request.args.get('typeCode')
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    jgbm = request.args.get('jgbm')            # 机构编码（可选）

    # 必传参数校验
    if not type_code:
        return jsonify({"error": "缺少必传参数：typeCode"}), 400
    if not start_date or not end_date:
        return jsonify({"error": "缺少必传参数：startDate, endDate"}), 400
    if not (start_date.isdigit() and len(start_date) == 8 and end_date.isdigit() and len(end_date) == 8):
        return jsonify({"error": "日期格式必须为 yyyymmdd，如 20260101"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 1. 获取该类型下（且符合 jgbm 筛选）的所有模型名称
        sql_models = """
            SELECT DISTINCT md.model_name
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE mt.type_code = ?
        """
        params_models = [type_code]
        if jgbm:
            sql_models += " AND md.jgbm = ?"
            params_models.append(jgbm)
        sql_models += " ORDER BY md.model_name"

        cursor.execute(sql_models, params_models)
        model_rows = cursor.fetchall()
        model_names = [row["model_name"] for row in model_rows]

        if not model_names:
            return Response(json.dumps([], ensure_ascii=False), mimetype='application/json')

        # 2. 查询实际数据：按月、按模型分组统计（同样应用 jgbm 筛选）
        sql_stats = """
            SELECT 
                SUBSTR(md.sjrq, 1, 6) AS month,
                md.model_name,
                COUNT(*) AS count
            FROM model_data md
            INNER JOIN model_type mt ON md.model_name = mt.model_name
            WHERE mt.type_code = ?
                AND md.sjrq BETWEEN ? AND ?
        """
        params_stats = [type_code, start_date, end_date]
        if jgbm:
            sql_stats += " AND md.jgbm = ?"
            params_stats.append(jgbm)
        sql_stats += " GROUP BY SUBSTR(md.sjrq, 1, 6), md.model_name ORDER BY month, md.model_name"

        cursor.execute(sql_stats, params_stats)
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

        # 4. 构建结果集
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

@app.route('/api/model-config')
def get_model_config():
    """
    根据模型名称获取 model_config 完整配置信息，并关联 model_type 获取类型代码
    返回格式：{"id": 记录ID, "type_code": 类型代码, "data": [{"字段名": 原值, "字段名_des": 描述, "字段名_disable": 禁用标志}, ...]}
    """
    model_name = request.args.get('model_name')
    if not model_name:
        return jsonify({"error": "缺少 model_name 参数"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()
        # 获取 model_config 所有列名
        cursor.execute("PRAGMA table_info(model_config)")
        all_columns = [row['name'] for row in cursor.fetchall()]
        sql = f"SELECT {','.join(all_columns)} FROM model_config WHERE model_name = ?"
        cursor.execute(sql, (model_name,))
        row = cursor.fetchone()
        if not row:
            return jsonify({"error": f"未找到模型: {model_name}"}), 404

        row_dict = dict(row)
        record_id = row_dict.get('id')

        # 关联查询 model_type 中的 type_code
        cursor.execute("SELECT type_code FROM model_type WHERE model_name = ? LIMIT 1", (model_name,))
        type_row = cursor.fetchone()
        type_code = type_row['type_code'] if type_row else None

        # 基础字段列表（决定返回顺序）
        base_fields = ['model_name', 'jgmc', 'jgbm', 'sjrq'] + [f'field{i}' for i in range(1, 21)]

        data = []
        for field in base_fields:
            original = row_dict.get(field, "")
            des = row_dict.get(f"{field}_des", "")
            disable = row_dict.get(f"{field}_disable", "")
            item = {
                field: original,
                f"{field}_des": des,
                f"{field}_disable": disable
            }
            data.append(item)

        result = {
            "id": record_id,
            "type_code": type_code,   # 新增字段，与 id 平级
            "data": data
        }

        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("查询模型配置错误：", e)
        return jsonify({"error": str(e)}), 500


import time
import random

@app.route('/api/model-config-upd', methods=['PUT', 'POST', 'GET'])
def update_model_config():
    data = request.get_json()
    if not data:
        return jsonify({"error": "请求体不能为空"}), 400

    required_fields = ['id', 'type_code', 'type_des']
    for field in required_fields:
        if field not in data:
            return jsonify({"error": f"缺少必传字段: {field}"}), 400

    record_id = data['id']
    type_code = data['type_code']
    type_des = data['type_des']

    try:
        conn = get_db()
        cursor = conn.cursor()

        # 获取 model_name
        cursor.execute("SELECT model_name FROM model_config WHERE id = ?", (record_id,))
        config_row = cursor.fetchone()
        if not config_row:
            return jsonify({"error": f"未找到 id 为 {record_id} 的记录"}), 404
        model_name = config_row['model_name']

        # 更新 model_config（动态字段，与原逻辑一致）
        cursor.execute("PRAGMA table_info(model_config)")
        columns = [row['name'] for row in cursor.fetchall() if row['name'] != 'id']
        update_fields = []
        params = []
        for col in columns:
            if col in data and col not in ['type_code', 'type_des']:
                update_fields.append(f"{col} = ?")
                params.append(data[col])
        if update_fields:
            params.append(record_id)
            sql = f"UPDATE model_config SET {', '.join(update_fields)} WHERE id = ?"
            cursor.execute(sql, params)

        # 处理 model_type 表
        cursor.execute("SELECT id FROM model_type WHERE model_name = ?", (model_name,))
        existing = cursor.fetchone()
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if existing:
            # 更新
            cursor.execute(
                "UPDATE model_type SET type_code = ?, type_des = ? WHERE model_name = ?",
                (type_code, type_des, model_name)
            )
        else:
            # 新增：生成随机ID（时间戳毫秒 + 随机数）
            random_id = int(time.time() * 1000) + random.randint(0, 9999)
            cursor.execute(
                "INSERT INTO model_type (id, model_name, type_code, type_des, create_time) VALUES (?, ?, ?, ?, ?)",
                (random_id, model_name, type_code, type_des, current_time)
            )

        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({"message": "更新成功"}), 200
    except Exception as e:
        print("更新模型配置错误：", e)
        return jsonify({"error": str(e)}), 500


@app.route('/api/batch-stats')
def batch_stats():
    """
    按批次号分组统计 model_data 中的记录数量
    返回格式: [{"batch_no": "批次号", "count": 数量}, ...]
    可选参数: limit - 限制返回的批次数量（默认返回全部）
              order - 排序方式，可选 desc（倒序，默认）或 asc
    """
    try:
        conn = get_db()
        cursor = conn.cursor()

        limit = request.args.get('limit', default=None, type=int)
        order = request.args.get('order', default='desc', type=str)

        sql = """
            SELECT batch_no, COUNT(*) AS count
            FROM model_data
            WHERE batch_no IS NOT NULL AND batch_no != ''
            GROUP BY batch_no
        """
        if order.lower() == 'desc':
            sql += " ORDER BY create_time DESC"
        else:
            sql += " ORDER BY create_time ASC"
        if limit is not None:
            sql += " LIMIT ?"
            params = (limit,)
        else:
            params = ()

        cursor.execute(sql, params)
        rows = cursor.fetchall()
        result = [{"batch_no": row["batch_no"], "count": row["count"]} for row in rows]
        cursor.close()
        conn.close()
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        print("批次统计查询错误：", e)
        return jsonify({"error": str(e)}), 500


@app.route('/api/batch-delete', methods=['POST'])
def batch_delete():
    """
    根据批次号删除 model_data 中的记录
    参数（JSON 或 query string）: batch_no
    返回: {"message": "删除成功", "deleted_count": n}
    """
    data = request.get_json(silent=True)
    if data:
        batch_no = data.get('batch_no')
    else:
        batch_no = request.args.get('batch_no')

    if not batch_no:
        return jsonify({"error": "缺少 batch_no 参数"}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM model_data WHERE batch_no = ?", (batch_no,))
        deleted = cursor.rowcount
        conn.commit()
        cursor.close()
        conn.close()

        if deleted == 0:
            return jsonify({"message": f"未找到批次 {batch_no} 的数据", "deleted_count": 0}), 200
        else:
            return jsonify({"message": f"成功删除 {deleted} 条记录", "deleted_count": deleted}), 200
    except Exception as e:
        print("批次删除错误：", e)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=25125, debug=True)
