1.项目结构
smart-monitor-system/
├── app.py                  # 项目启动入口
├── .env                    # 环境配置
├── requirements.txt        # 固定依赖
├── uploads/                # 上传文件目录
├── static/                 # 静态资源
│   └── echarts.min.js      # 图表库
└── templates/              # 页面模板
    ├── base.html           # 全局布局
    ├── dashboard.html      # 【核心】驾驶舱
    ├── model_import.html   # 模型导入
    ├── db_config.html      # 数据库配置
    └── chart.html          # 图表分


2.安装依赖
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple




3.数据结构变更
记录：增加导入日志表，model_data表增加data_hash_code字段

CREATE TABLE "import_log" (
	"id"	INTEGER,
	"batch_no"	TEXT,
	"status"	TEXT NOT NULL,
	"error_msg"	TEXT,
	"create_time"	TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
	PRIMARY KEY("id" AUTOINCREMENT)
)

-- 为 model_data 表添加 data_hash_code 列，用于存储文件数据内容的哈希值（MD5）
ALTER TABLE model_data ADD COLUMN data_hash_code TEXT;
CREATE INDEX idx_model_data_hash ON model_data(data_hash_code);
CREATE INDEX idx_model_name ON model_data(model_name);
CREATE INDEX idx_jgbm ON model_data(jgbm);

--若是要保证表中结构，参看model_data.sql文件，只是显示顺序对数据操作无影响。

