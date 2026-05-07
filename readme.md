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

