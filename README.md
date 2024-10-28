# iTop 运维服务月报生成器

这是一个使用Python和Streamlit开发的iTop运维服务月报生成器。该应用程序连接到iTop数据库，获取相关数据，并生成一份详细的月度报告。

## 功能

- 显示指定日期范围内的工单统计
- 按服务类型(服务请求、事件、变更)分析工单
- 展示团队和个人的工单处理情况
- 列出未解决的工单
- 使用饼图可视化工单状态分布

## 安装

1. 克隆此仓库:
   ```bash
   git clone https://github.com/reven-tang/itop-report.git
   ```

2. 进入项目目录:
   ```bash
   cd itop-report
   ```

3. 创建并激活Python虚拟环境:
   ```bash
   python3 -m venv myenv
   source myenv/bin/activate  # 在Windows上使用 myenv\Scripts\activate
   ```

4. 安装所需的Python包:
   ```bash
   pip install -r requirements.txt
   ```

5. 创建配置文件:
   在项目根目录下创建一个名为`config.ini`的文件,并按以下格式填写数据库连接信息:
   ```ini
   [database]
   host = your_database_host
   port = your_database_port
   user = your_database_username
   password = your_database_password
   database = your_database_name
   ```
   请确保将上述占位符替换为实际的数据库连接信息。

## 使用方法

1. 确保你已经激活了虚拟环境。如果没有，请运行:
   ```bash
   source myenv/bin/activate  # 在Windows上使用 myenv\Scripts\activate.bat
   ```

2. 运行Streamlit应用:
   ```bash
   streamlit run itop_report.py
   ```

3. 在浏览器中打开显示的URL(通常是 http://localhost:8501)

4. 在左侧边栏选择要生成报告的日期范围（默认为上个月）

5. 查看生成的报告，包括工单统计、服务类型分析、团队和个人统计以及未解决工单列表

6. 当你完成使用后，可以通过以下命令退出虚拟环境:
   ```bash
   deactivate
   ```

## 添加为系统守护服务

如下配置假设我们将代码放置在/data/itop-report/下。

1. 创建服务启动脚本:
   在项目根目录下创建 `service.sh` 文件并添加以下内容:
   ```bash
   cat > /data/itop-report/service.sh  <<"EOF"
   #!/bin/sh

   source ./myenv/bin/activate

   # 终止所有 itop 报表服务进程
   pgrep -f 'streamlit run itop_report.py' | xargs kill -9

   # 启动
   streamlit run itop_report.py --server.port 8080 >/dev/nu11 2>&1 &
   sleep 2
   itop_pids=($(pgrep -f 'streamlit run itop_report.py'))
   if [ ${#itop_pids[@]} -ge 1 ]; then
       echo "服务启动成功"
   else
       echo "服务启动异常"
   fi
   EOF
   ```

2. 创建系统服务文件:
   创建 `/usr/lib/systemd/system/itop-report.service` 文件并添加以下内容:
   ```bash
   cat > /usr/lib/systemd/system/itop-report.service  <<"EOF"
   [Unit]
   Description=itop-report
   After=network.target

   [Service] 
   WorkingDirectory=/data/itop-report
   Type=forking
   ExecStart=/data/itop-report/service.sh
   Restart=always 
   PrivateTmp=true
   RestartSec=10

   [Install]
   WantedBy=multi-user.target
   EOF
   ```

9. 启用并启动服务:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable itop-report.service
   sudo systemctl start itop-report.service
   ```

现在，itop-report 服务将作为系统守护进程运行，并在系统启动时自动启动。

## 注意事项

- 确保您有权限访问iTop数据库
- 请妥善保管您的`config.ini`文件,不要将其上传到公共仓库
- 根据需要修改数据库连接信息（在 `config.ini` 文件中）

## 主要依赖

- Python 3.7+
- Streamlit
- Pandas
- SQLAlchemy
- PyMySQL
- Plotly
- ReportLab

### windows安装版本：

- Python 3.7+
- streamlit==1.39.0
- pandas==2.2.3
- SQLAlchemy==2.0.35
- PyMySQL==1.1.1
- plotly==5.24.1
- reportlab==4.2.5

### linux安装版本：

- Python 3.7+
- streamlit==1.23.1
- pandas==2.1.4
- SQLAlchemy==1.4.39
- PyMySQL==1.1.1
- plotly==5.18.1
- reportlab==4.2.5

国内加速源：
```
tee ~/.pip/pip.conf <<-'EOF'
[global]
timeout=600
index-url = http://mirrors.aliyun.com/pypi/simple/
[install]
trusted-host = mirrors.aliyun.com
EOF
```

详细的依赖列表请查看 `requirements.txt` 文件。

## 贡献

欢迎提交问题和拉取请求。对于重大更改，请先开issue讨论您想要更改的内容。

## 许可

[MIT](https://choosealicense.com/licenses/mit/)
