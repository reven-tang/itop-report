import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta, date
import plotly.express as px
import calendar
import configparser
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.colors import HexColor
from reportlab.graphics.shapes import String
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.charts.textlabels import Label
import os

# 连接到iTop数据库
def connect_to_itop_db():
    # 从配置文件读取数据库连接信息    
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    db_host = config['Database']['host']
    db_user = config['Database']['user']
    db_password = config['Database']['password']
    db_port = config['Database']['port']
    db_name = config['Database']['database']
    
    return create_engine(f'mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}')

# 执行SQL查询并返回DataFrame
def execute_query(engine, query, params):
    # 将日期转换为字符串格式
    for key, value in params.items():
        if isinstance(value, (date, datetime)):
            params[key] = value.strftime('%Y-%m-%d')
    
    print("Executing query:", query)
    print("With parameters:", params)
    
    with engine.connect() as connection:
        df = pd.read_sql(query, connection, params=params)
    return df

# 1. 工单统计
def get_ticket_summary(engine, start_date, end_date):
    query = """
    SELECT 
        count(1) as total,
        SUM(CASE WHEN t.finalclass = 'UserRequest' THEN 1 ELSE 0 END) as request_total,
        SUM(CASE WHEN t.finalclass LIKE '%%Change%%' THEN 1 ELSE 0 END) as change_total,
        SUM(CASE WHEN t.finalclass = 'Incident' THEN 1 ELSE 0 END) as Incident_total
    FROM ticket t 
    LEFT JOIN ticket_request tr ON tr.id = t.id 
    LEFT JOIN ticket_incident ti ON ti.id = t.id 
    LEFT JOIN `change` c ON c.id = t.id 
    WHERE t.finalclass <> 'Problem' 
    AND t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND (tr.status <> 'new' OR ti.status <> 'new' OR c.status <> 'new')
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 2. 服务请求状态统计
def get_user_request_stats(engine, start_date, end_date):
    query = """
    SELECT
        count(1) as total,
        SUM(CASE WHEN tr.status in ('closed','resolved') THEN 1 ELSE 0 END) as resolved_total,
        SUM(CASE WHEN tr.status = 'closed' THEN 1 ELSE 0 END) as closed_total,
        SUM(CASE WHEN tr.status not in ('closed','resolved') THEN 1 ELSE 0 END) as unresolved_total
    FROM ticket_request tr  
    LEFT JOIN ticket t ON tr.id = t.id 
    WHERE t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND tr.status <> 'new'
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 3. 事件状态统计
def get_incident_stats(engine, start_date, end_date):
    query = """
    SELECT
        count(1) as total,
        SUM(CASE WHEN ti.status in ('closed','resolved') THEN 1 ELSE 0 END) as resolved_total,
        SUM(CASE WHEN ti.status = 'closed' THEN 1 ELSE 0 END) as closed_total,
        SUM(CASE WHEN ti.status not in ('closed','resolved') THEN 1 ELSE 0 END) as unresolved_total
    FROM ticket_incident ti  
    LEFT JOIN ticket t ON ti.id = t.id 
    WHERE t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND ti.status <> 'new'
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 4. 变更状态统计
def get_change_stats(engine, start_date, end_date):
    query = """
    SELECT
        count(1) as total,
        SUM(CASE WHEN c.status in ('closed','resolved') THEN 1 ELSE 0 END) as resolved_total,
        SUM(CASE WHEN c.status = 'closed' THEN 1 ELSE 0 END) as closed_total,
        SUM(CASE WHEN c.status not in ('closed','resolved') THEN 1 ELSE 0 END) as unresolved_total
    FROM `change` c 
    LEFT JOIN ticket t ON c.id = t.id 
    WHERE t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND c.status <> 'new'
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 5. 按团队统计处理时长
def get_team_stats(engine, start_date, end_date):
    query = """
    SELECT 
        DATE_FORMAT(subquery.start_date, '%%Y-%%m') AS '月份',
        c.name AS '团队',
        subquery.ticket_type AS '工单类型',
        COUNT(*) AS '工单数量',
        SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END) AS '未解决',
        SUM(CASE WHEN (subquery.tto_75_passed = 1 OR subquery.ttr_75_passed = 1) THEN 1 ELSE 0 END) AS '超时工单',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单解决率',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN (subquery.tto_75_passed = 1 OR subquery.ttr_75_passed = 1) THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单及时率',
        CASE 
            WHEN subquery.ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(AVG(subquery.response_time) / 60, 2)
        END AS '平均响应时长(分钟)',
        ROUND(AVG(subquery.resolution_time) / 60, 2) AS '平均解决时长(分钟)',
        CASE 
            WHEN subquery.ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(MAX(subquery.response_time) / 60, 2)
        END AS '最大响应时长(分钟)',
        ROUND(MAX(subquery.resolution_time) / 60, 2) AS '最大解决时长(分钟)'
    FROM (
        SELECT 
            t.team_id,
            t.start_date,
            '服务请求' AS ticket_type,
            tr.status,
            tr.tto_75_passed,
            tr.ttr_75_passed,
            TIMESTAMPDIFF(SECOND, tr.tto_started, tr.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, tr.tto_stopped, tr.ttr_stopped) AS resolution_time
        FROM ticket t 
        JOIN ticket_request tr ON tr.id = t.id 
        WHERE tr.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.team_id,
            t.start_date,
            '事件' AS ticket_type,
            ti.status,
            ti.tto_75_passed,
            ti.ttr_75_passed,
            TIMESTAMPDIFF(SECOND, ti.tto_started, ti.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, ti.tto_stopped, ti.ttr_stopped) AS resolution_time
        FROM ticket t 
        JOIN ticket_incident ti ON ti.id = t.id
        WHERE ti.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.team_id,
            t.start_date,
            '变更' AS ticket_type,
            c2.status,
            0 AS tto_75_passed,  -- 变更工单没有响应时间要求
            0 AS ttr_75_passed,  -- 变更工单暂不考虑解决时间超时
            NULL AS response_time,
            TIMESTAMPDIFF(SECOND, t.start_date, t.end_date) AS resolution_time
        FROM ticket t 
        JOIN `change` c2 ON c2.id = t.id
        WHERE c2.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
    ) AS subquery
    JOIN contact c ON subquery.team_id = c.id 
    WHERE c.finalclass = 'Team'
    GROUP BY DATE_FORMAT(subquery.start_date, '%%Y-%%m'), c.name, subquery.ticket_type
    HAVING COUNT(*) > 0
    ORDER BY DATE_FORMAT(subquery.start_date, '%%Y-%%m') DESC, subquery.ticket_type DESC, c.name
    """
    
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 6. 按人员统计处理时长
def get_person_stats(engine, start_date, end_date):
    query = """
    WITH agent_info AS (
        SELECT p1.id, CONCAT(COALESCE(c1.name, ''), ' ', COALESCE(p1.first_name, '')) AS agent_name
        FROM person p1
        JOIN contact c1 ON p1.id = c1.id
    )

    SELECT 
        DATE_FORMAT(start_date, '%%Y-%%m') AS '月份',
        ai.agent_name AS '办理人',
        ticket_type AS '工单类型',
        COUNT(*) AS '工单数量',
        SUM(CASE WHEN status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END) AS '未解决',
        SUM(CASE WHEN (tto_75_passed = 1 OR ttr_75_passed = 1) THEN 1 ELSE 0 END) AS '超时工单',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单解决率',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN (tto_75_passed = 1 OR ttr_75_passed = 1) THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单及时率',
        CASE 
            WHEN ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(AVG(response_time) / 60, 2)
        END AS '平均响应时长(分钟)',
        ROUND(AVG(resolution_time) / 60, 2) AS '平均解决时长(分钟)',
        CASE 
            WHEN ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(MAX(response_time) / 60, 2)
        END AS '最大响应时长(分钟)',
        ROUND(MAX(resolution_time) / 60, 2) AS '最大解决时长(分钟)'
    FROM (
        SELECT 
            t.agent_id,
            '服务请求' AS ticket_type,
            tr.status,
            t.start_date,
            tr.tto_75_passed,
            tr.ttr_75_passed,
            TIMESTAMPDIFF(SECOND, tr.tto_started, tr.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, tr.tto_stopped, tr.ttr_stopped) AS resolution_time
        FROM ticket_request tr
        JOIN ticket t ON tr.id = t.id
        WHERE tr.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.agent_id,
            '事件' AS ticket_type,
            ti.status,
            t.start_date,
            ti.tto_75_passed,
            ti.ttr_75_passed,
            TIMESTAMPDIFF(SECOND, ti.tto_started, ti.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, ti.tto_stopped, ti.ttr_stopped) AS resolution_time
        FROM ticket_incident ti
        JOIN ticket t ON ti.id = t.id
        WHERE ti.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.agent_id,
            '变更' AS ticket_type,
            c2.status,
            t.start_date,
            0 AS tto_75_passed,  -- 变更工单没有响应时间要求
            0 AS ttr_75_passed,  -- 变更工单暂不考虑解决时间超时
            NULL AS response_time,
            TIMESTAMPDIFF(SECOND, t.start_date, t.end_date) AS resolution_time
        FROM `change` c2
        JOIN ticket t ON c2.id = t.id
        WHERE c2.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
    ) AS subquery
    JOIN agent_info ai ON subquery.agent_id = ai.id
    GROUP BY ai.agent_name, ticket_type, DATE_FORMAT(start_date, '%%Y-%%m')
    ORDER BY DATE_FORMAT(start_date, '%%Y-%%m') DESC, ticket_type DESC, ai.agent_name
    """

    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 7. 未解决的工单
def get_unresolved_tickets(engine, start_date, end_date):
    query = """
    SELECT 
        t.ref AS '工单号', 
        t.title AS '标题', 
        t.start_date AS '开始时间',
        CASE 
            WHEN tr.id IS NOT NULL THEN tr.status
            WHEN ti.id IS NOT NULL THEN ti.status
            WHEN cg.id IS NOT NULL THEN cg.status
        END AS '状态', 
        CONCAT(IFNULL(c.name, ''), ' ', IFNULL(p.first_name, '')) AS '发起人', 
        c2.name AS '团队名称', 
        CONCAT(IFNULL(c1.name, ''), ' ', IFNULL(p1.first_name, '')) AS '办理人'
    FROM ticket t
    LEFT JOIN ticket_request tr ON tr.id = t.id
    LEFT JOIN ticket_incident ti ON ti.id = t.id
    LEFT JOIN `change` cg ON cg.id = t.id
    LEFT JOIN (person p JOIN contact c ON p.id = c.id) ON t.caller_id = p.id 
    LEFT JOIN (person p1 JOIN contact c1 ON p1.id = c1.id) ON t.agent_id = p1.id 
    LEFT JOIN contact c2 ON t.team_id = c2.id 
    WHERE (tr.status NOT IN ('closed','new','resolved') 
        OR ti.status NOT IN ('closed','new','resolved')
        OR cg.status NOT IN ('closed','new','resolved'))
    AND t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 8. 超时工单
def get_overdue_tickets(engine, start_date, end_date):
    query = """
    SELECT 
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(t.ref, '')) AS CHAR) ELSE NULL END) AS '工单号', 
        t.title AS '标题',
        tr.status AS '状态', 
        t.start_date AS '开始日期',
        t.last_update AS '最后日期',
        ROUND(tr.tto_100_overrun / 60,2) AS '响应时间超过(分钟)',
        ROUND(tr.ttr_100_overrun / 60,2) AS '解决时间超过(分钟)',
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(c1.name, ''), COALESCE(' ', ''), COALESCE(p1.first_name, '')) AS CHAR) ELSE NULL END) AS '发起人', 
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(c2.name, '')) AS CHAR) ELSE NULL END) AS '团队名称', 
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(c1.name, ''), COALESCE(' ', ''), COALESCE(p1.first_name, '')) AS CHAR) ELSE NULL END) AS '办理人',
        tr.assignment_date AS '实际响应时间',
        tr.resolution_date AS '实际解决时间',
        tr.tto_100_deadline AS '响应最后期限',
        tr.ttr_100_deadline AS '解决最后期限',
        ROUND(AVG(TIMESTAMPDIFF(SECOND,tr.tto_started,tr.tto_stopped) / 60),2) AS '响应时长(分钟)' , 
        ROUND(AVG(TIMESTAMPDIFF(SECOND,tr.tto_stopped,tr.ttr_stopped) / 60),2) AS '解决时长(分钟)'
    FROM ticket t 
    LEFT JOIN ticket_request tr ON tr.id = t.id 
    LEFT JOIN ( person AS p INNER JOIN contact AS c ON p.id = c.id ) ON tr.approver_id = p.id 
    LEFT JOIN ( person AS p1 INNER JOIN contact AS c1 ON p1.id = c1.id ) ON t.agent_id = p1.id 
    LEFT JOIN contact AS c2 ON t.team_id = c2.id 
    WHERE t.finalclass <> 'Problem' 
    AND ( tr.tto_75_passed = 1 or tr.ttr_75_passed )
    AND t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    GROUP BY tr.status
    HAVING COUNT(*) > 0
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

def generate_pdf(start_date, end_date, ticket_summary, user_request_stats, incident_stats, change_stats, team_stats, person_stats, unresolved_tickets, overdue_tickets):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []

    # 注册中文字体
    font_path = "./simkai.ttf"
    pdfmetrics.registerFont(TTFont('SimKai', font_path))

    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    subtitle_style = styles['Heading2']
    normal_style = styles['Normal']

    # 设置中文字体
    title_style.fontName = 'SimKai'
    subtitle_style.fontName = 'SimKai'
    normal_style.fontName = 'SimKai'

    # 添加标题
    if start_date.month == end_date.month:
        title = f"<para alignment='center'>iTop 运维服务报表 ({start_date.year}年{start_date.month}月)</para>"
    else:
        title = f"<para alignment='center'>iTop 运维服务报表 ({start_date.year}年{start_date.month}月至{end_date.year}年{end_date.month}月)</para>"
    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 12))

    # 1. 工单统计
    total_tickets = ticket_summary['total'].iloc[0] if not ticket_summary.empty else 0
    elements.append(Paragraph(f"iTop共接收工单数 {total_tickets} 起，各类工单处理情况如下：", subtitle_style))
    elements.append(Spacer(1, 12))

    # 2. 按服务类型统计分析
    elements.append(Paragraph("1. 按服务类型统计分析如下：", subtitle_style))

    def create_pie_chart(data, labels, title):
        drawing = Drawing(400, 250)
        pie = Pie()
        pie.x = 100
        pie.y = 25
        pie.width = 200
        pie.height = 200
        pie.data = data
        
        # 直接使用标签作为饼图标签
        pie.labels = labels
        pie.slices.strokeWidth = 0.5
        
        # 设置标签样式
        pie.sideLabels = True  # 将标签放在饼图外侧
        pie.sideLabelsOffset = 0.1  # 调整标签距离
        pie.simpleLabels = False  # 允许自定义标签样式
        pie.slices.fontName = 'SimKai'  # 设置标签字体为楷体
        # 将标签替换为百分比
        total = sum(pie.data)
        pie.labels = ['%.1f%%' % (value/total*100) for value in pie.data]

        # 设置颜色
        colors = [HexColor('#00b8a9'), HexColor('#f6416c'), HexColor('#f8f3d4')]
        for i, color in enumerate(colors[:len(data)]):
            pie.slices[i].fillColor = color

        drawing.add(pie)

        # 添加标题
        title_label = String(200, 250, title)  # 从240上移到250
        title_label.fontName = 'SimKai'
        title_label.fontSize = 12
        title_label.textAnchor = 'middle'
        drawing.add(title_label)

        # 添加图例
        legend = Legend()
        legend.x = 320
        legend.y = 150
        legend.deltay = 15
        legend.fontSize = 10
        legend.fontName = 'SimKai'
        legend.alignment = 'right'
        legend.columnMaximum = 8
        legend.colorNamePairs = list(zip(colors[:len(data)], labels))
        drawing.add(legend)

        return drawing

    # 2.1 服务请求统计
    elements.append(Paragraph("1) 服务请求统计", subtitle_style))
    if not user_request_stats.empty:
        total = user_request_stats['total'].iloc[0]
        resolved = user_request_stats['resolved_total'].iloc[0]
        closed = user_request_stats['closed_total'].iloc[0]
        unresolved = user_request_stats['unresolved_total'].iloc[0]
        
        if total and total > 0:
            resolved_percentage = resolved / total * 100
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100
            
            elements.append(Paragraph(f"本周期内共接收服务请求 {total:g} 个，其中 {resolved:g} 个服务请求被解决，占比约 {resolved_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(f"已解决的服务请求中，{closed:g} 个服务请求被按时关闭，占比约 {closed_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(f"未解决的服务请求有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。", normal_style))
            
            # 添加两行空行
            elements.append(Spacer(1, 12))
            elements.append(Spacer(1, 12))

            # 添加饼图
            pie_data = [resolved, unresolved, closed]
            pie_labels = ['已解决', '未解决', '已关闭']
            elements.append(create_pie_chart(pie_data, pie_labels, "服务请求状态分布"))
        else:
            elements.append(Paragraph("本周期内没有接收到服务请求。", normal_style))
    else:
        elements.append(Paragraph("无法获取服务请求统计数据。", normal_style))
    elements.append(Spacer(1, 12))

    # 2.2 事件统计
    elements.append(Paragraph("2) 事件统计", subtitle_style))
    if not incident_stats.empty:
        total = incident_stats['total'].iloc[0]
        resolved = incident_stats['resolved_total'].iloc[0]
        closed = incident_stats['closed_total'].iloc[0]
        unresolved = incident_stats['unresolved_total'].iloc[0]

        if total and total > 0:
            resolved_percentage = resolved / total * 100
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100

            elements.append(Paragraph(f"本周期内共发生事件 {total:g} 个，其中 {resolved:g} 个事件被解决，占比约 {resolved_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(f"已解决的事件中，{closed:g} 个事件被按时关闭，占比约 {closed_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(f"未解决的事件有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。", normal_style))
            
            # 添加一行空行
            elements.append(Spacer(1, 12)) 
            
            # 添加饼图
            pie_data = [resolved, unresolved, closed]
            pie_labels = ['已解决', '未解决', '已关闭']
            elements.append(create_pie_chart(pie_data, pie_labels, "事件状态分布"))
        else:
            elements.append(Paragraph("本周期内没有发生事件。", normal_style))
    else:
        elements.append(Paragraph("无法获取事件统计数据。", normal_style))
    elements.append(Spacer(1, 12))

    # 2.3 变更统计
    elements.append(Paragraph("3) 变更统计", subtitle_style))
    if not change_stats.empty:
        total = change_stats['total'].iloc[0]
        resolved = change_stats['resolved_total'].iloc[0]
        closed = change_stats['closed_total'].iloc[0]

        if total and total > 0:
            closed_percentage = closed / total * 100
            resolved_percentage = resolved / closed * 100 if closed > 0 else 0

            elements.append(Paragraph(f"本周期内共发生变更 {total:g} 个，其中 {closed:g} 个变更已关闭，占比约 {closed_percentage:.2f}%。", normal_style))
            elements.append(Paragraph(f"已关闭的变更中，{resolved:g} 个变更被成功执行，占比约 {resolved_percentage:.2f}%。", normal_style))
            
            # 添加一行空行
            elements.append(Spacer(1, 12)) 
            
            # 添加饼图
            pie_data = [resolved, total-resolved]
            pie_labels = ['已解决', '未解决']
            elements.append(create_pie_chart(pie_data, pie_labels, "变更状态分布"))
        else:
            elements.append(Paragraph("本周期内没有发生变更。", normal_style))
    else:
        elements.append(Paragraph("无法获取变更统计数据。", normal_style))
    elements.append(Spacer(1, 12))

    # 3. 按照工单处理团队统计
    elements.append(Paragraph("2. 按照工单处理团队统计，具体如下", subtitle_style))
    if not team_stats.empty:
        team_data = [team_stats.columns.tolist()] + team_stats.values.tolist()
        # 计算表格宽度为页面宽度的85%
        table_width = letter[0] * 0.85
        col_widths = [table_width / len(team_data[0])] * len(team_data[0])
        
        team_table = Table(team_data, colWidths=col_widths)
        team_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'SimKai'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        # 设置单元格自动换行
        for i, row in enumerate(team_data):
            for j, cell in enumerate(row):
                team_table._cellvalues[i][j] = Paragraph(str(cell), normal_style)
        
        elements.append(team_table)     
        
        # 3.1 按照工单处理团队绘制服务请求的解决率
        # 将team_stats转换为pandas DataFrame
        df = pd.DataFrame(team_stats)
        
        # 按月份和团队分组计算平均解决率和及时率
        df['工单解决率'] = df['工单解决率'].apply(lambda x: float(str(x).rstrip('%')))
        
        # 检查是否跨月
        if len(df['月份'].unique()) > 1:
            # 仅保留服务请求数据
            service_request_df = df[df['工单类型'] == '服务请求']
            
            # 如果有服务请求数据才继续绘图
            if not service_request_df.empty:
                # 创建折线图,设置画布大小为500x300
                drawing = Drawing(600, 300)  # 使用页面宽度的85%作为图表宽度
                # 创建折线图对象
                lp = LinePlot()
                lp.data = []
                # 设置折线图在画布中的位置和大小
                lp.x = 10  # 从50改为10,向左偏移40个单位
                lp.y = 50
                lp.height = 200
                lp.width = 450
                
                # 获取所有月份和团队列表并排序
                months = sorted(service_request_df['月份'].unique())
                teams = sorted(service_request_df['团队'].unique())
                
                # 为每个团队创建一条折线
                for i, team in enumerate(teams):
                    # 获取当前团队的数据
                    team_data = service_request_df[service_request_df['团队'] == team]
                    data = []  # 存储解决率数据
                    x_data = [] # 存储月份数据
                    # 遍历每个月份获取数据点
                    for j, month in enumerate(months):
                        month_data = team_data[team_data['月份'] == month]
                        if not month_data.empty:
                            # 添加该月的解决率数据
                            data.append(month_data['工单解决率'].iloc[0])
                            # 将月份转换为202301格式
                            x_data.append(int(month.split('-')[1]))
                            print(x_data)
                    # 只有当有数据时才添加到图表中
                    if data and x_data:  
                        lp.data.append(list(zip(x_data, data)))
                
                # 配置x轴,使用实际的月份值作为刻度
                unique_months = sorted(set([x for team_data in lp.data for x,_ in team_data]))
                lp.xValueAxis.valueSteps = unique_months
                lp.xValueAxis.labels = [Label() for _ in unique_months]
                # 设置x轴标签样式
                for i, label in enumerate(lp.xValueAxis.labels):
                    # 找到对应的完整月份字符串
                    month_num = unique_months[i]
                    month_str = next(m for m in months if int(m.split('-')[1]) == month_num)
                    label._text = month_str
                    label.fontName = 'SimKai'
                    label.fontSize = 10
                    label.angle = 0  # 标签水平显示
                    label.dx = 0   # x方向偏移量向左20个单位
                    label.dy = -20   # y方向偏移量(向下)
                
                # 配置y轴范围和刻度
                lp.yValueAxis.valueMin = 0
                lp.yValueAxis.valueMax = 100
                lp.yValueAxis.valueStep = 10
                # 添加横向网格线
                lp.yValueAxis.visibleGrid = True  # 显示网格线
                lp.yValueAxis.gridStrokeColor = colors.Color(0.9, 0.9, 0.9)  # 设置网格线颜色为更浅的灰色
                lp.yValueAxis.gridStrokeWidth = 0.5  # 设置网格线宽度
                
                # 创建图例对象并设置样式
                legend = Legend()  # 实例化一个Legend对象用于显示图例
                
                # 设置图例位置和大小，使其居中显示在图表上方
                legend.x = lp.x + (lp.width / 2)  # 设置图例的x坐标为图表宽度的一半,使其水平居中
                legend.y = lp.y + lp.height + 30  # 设置图例的y坐标,使其位于图表上方30个单位处
                legend.dx = 8  # 设置图例的x方向偏移量为8
                legend.dy = 8  # 设置图例的y方向偏移量为8
                
                # 设置图例文字样式
                legend.fontName = 'SimKai'  # 设置图例字体为楷体
                legend.fontSize = 9  # 设置图例字体大小为10
                legend.boxAnchor = 'n'  # 设置图例框的锚点位置为north(北)
                legend.columnMaximum = 1  # 设置图例最大列数为1
                
                # 设置图例边框样式
                legend.strokeWidth = 0.5  # 设置图例边框线宽为0.5
                legend.strokeColor = colors.black  # 设置图例边框颜色为黑色
                
                # 设置图例内部布局
                legend.deltax = 75  # 设置图例项之间的水平间距为75
                legend.deltay = 10  # 设置图例项之间的垂直间距为10
                legend.autoXPadding = 5  # 设置图例自动水平内边距为5
                legend.yGap = 0  # 设置图例垂直间隙为0
                legend.dxTextSpace = 5  # 设置图例文本与标记之间的间距为5
                
                # 设置图例分隔线
                legend.dividerLines = 1|2|4  # 设置图例分隔线的显示方式(上|下|中间)
                legend.dividerOffsY = 4.5  # 设置分隔线的垂直偏移量为4.5
                legend.subCols.rpad = 30  # 设置子列的右侧内边距为30
                
                # 为每个有数据的团队添加图例项
                legend.colorNamePairs = []
                for i, team in enumerate(teams):
                    if not service_request_df[service_request_df['团队'] == team].empty:
                        # 第一个团队用红色,其他用蓝色
                        color = colors.red if i == 0 else colors.blue
                        legend.colorNamePairs.append((color, team))
                
                # 只有当有图例数据时才添加图表和图例
                if legend.colorNamePairs:  
                    drawing.add(lp)
                    drawing.add(legend)
                    
                    # 在每个数据点上添加数值标签
                    for i, team_data in enumerate(lp.data):
                        for x, y in team_data:
                            label = String(lp.x + (x - min(unique_months)) * (lp.width / (max(unique_months) - min(unique_months))) + 12,  # 向右偏移12个单位
                                         lp.y + y * (lp.height / 100) + 5,  # 向上偏移5个单位
                                         '%.1f%%' % y,  # 显示一位小数
                                         fontSize=9,
                                         fontName='SimKai',
                                         textAnchor='middle')
                            drawing.add(label)
                    
                    # 添加图表标题和图表到PDF
                    elements.append(Paragraph("各团队服务请求月度解决率趋势", subtitle_style))
                    elements.append(drawing)
                    elements.append(Spacer(1, 12))
    else:
        elements.append(Paragraph("本周期内没有要处理的工单", normal_style))
    elements.append(Spacer(1, 12))

    # 4. 按照工程师统计
    elements.append(Paragraph("3. 按照工程师统计，具体如下", subtitle_style))
    if not person_stats.empty:
        person_data = [person_stats.columns.tolist()] + person_stats.values.tolist()
        # 计算表格宽度为页面宽度的85%
        table_width = letter[0] * 0.85
        col_widths = [table_width / len(person_data[0])] * len(person_data[0])
        
        person_table = Table(person_data, colWidths=col_widths)
        person_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'SimKai'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        # 设置单元格自动换行
        for i, row in enumerate(person_data):
            for j, cell in enumerate(row):
                person_table._cellvalues[i][j] = Paragraph(str(cell), normal_style)
        
        elements.append(person_table)
    else:
        elements.append(Paragraph("本周期内没有要处理的工单", normal_style))
    elements.append(Spacer(1, 12))

    # 5. 未解决的工单
    elements.append(Paragraph("4. 未解决的工单如下", subtitle_style))
    if not unresolved_tickets.empty:
        unresolved_data = [unresolved_tickets.columns.tolist()] + unresolved_tickets.values.tolist()
        # 计算表格宽度为页面宽度的85%
        table_width = letter[0] * 0.85
        col_widths = [table_width / len(unresolved_data[0])] * len(unresolved_data[0])
        
        unresolved_table = Table(unresolved_data, colWidths=col_widths)
        unresolved_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'SimKai'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        # 设置单元格自动换行
        for i, row in enumerate(unresolved_data):
            for j, cell in enumerate(row):
                unresolved_table._cellvalues[i][j] = Paragraph(str(cell), normal_style)
        
        elements.append(unresolved_table)
    else:
        elements.append(Paragraph("本周期内没有未解决的工单。", normal_style))
    elements.append(Spacer(1, 12))

    # 6. 超时的工单
    elements.append(Paragraph("5. SLA超时的工单如下", subtitle_style))
    if not overdue_tickets.empty:
        overdue_data = [overdue_tickets.columns.tolist()] + overdue_tickets.values.tolist()
        # 计算表格宽度为页面宽度的85%
        table_width = letter[0] * 0.85
        col_widths = [table_width / len(overdue_data[0])] * len(overdue_data[0])
        
        overdue_table = Table(overdue_data, colWidths=col_widths)
        overdue_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'SimKai'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        # 设置单元格自动换行
        for i, row in enumerate(overdue_data):
            for j, cell in enumerate(row):
                overdue_table._cellvalues[i][j] = Paragraph(str(cell), normal_style)
        
        elements.append(overdue_table)
    else:
        elements.append(Paragraph("本周期内没有SLA超时的工单。", normal_style))
        
    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

def main():
    # 创建左边栏
    with st.sidebar:
        st.title("iTop 报表查询")
        st.markdown("<style>h1{text-align: center;}</style>", unsafe_allow_html=True)
        # 添加一条横线
        st.markdown("---")

        # 添加日期选择提示
        st.markdown("""
        <div>  </div>
        <div style='color: #808080; font-style: italic;'>
        请选择要查询的开始日期和结束日期\r\n
        (系统默认为上一个月的数据)
        </div>
        """, unsafe_allow_html=True)

        # 日期选择
        today = datetime.now()
        last_month = today.replace(day=1) - timedelta(days=1)
        
        st.markdown("开始日期", unsafe_allow_html=True)
        start_date = st.date_input("", last_month.replace(day=1), key="start_date", label_visibility="collapsed")
        
        st.markdown("结束日期", unsafe_allow_html=True)
        end_date = st.date_input("", last_month.replace(day=calendar.monthrange(last_month.year, last_month.month)[1]), key="end_date", label_visibility="collapsed")

        # 连接数据库
        engine = connect_to_itop_db()

        # 获取数据
        ticket_summary = get_ticket_summary(engine, start_date, end_date)
        user_request_stats = get_user_request_stats(engine, start_date, end_date)
        incident_stats = get_incident_stats(engine, start_date, end_date)
        change_stats = get_change_stats(engine, start_date, end_date)
        team_stats = get_team_stats(engine, start_date, end_date)
        person_stats = get_person_stats(engine, start_date, end_date)
        unresolved_tickets = get_unresolved_tickets(engine, start_date, end_date)
        overdue_tickets = get_overdue_tickets(engine, start_date, end_date)

        # 插入一行空行
        st.write("")

        # 添加导出PDF按钮
        col1, col2, col3 = st.columns([1, 1, 2])
        with col3:
            if st.button('导出PDF报表'):
                try:
                    pdf = generate_pdf(start_date, end_date, ticket_summary, user_request_stats, incident_stats, change_stats, team_stats, person_stats, unresolved_tickets, overdue_tickets)
                    with col3:
                        st.download_button(
                            label="下载PDF报表",
                            data=pdf,
                            file_name="itop_report.pdf",
                            mime="application/pdf"
                        )
                except Exception as e:
                    st.error(f"生成PDF时发生错误: {str(e)}")
                    st.error("请检查是否安装了所需的中文字体。")

    # 主要内容区域
    st.markdown("<h2 style='text-align: center;'>iTop 运维服务报表</h2>", unsafe_allow_html=True)

    # 显示报告
    if start_date.month == end_date.month:
        st.markdown(f"<div style='text-align: right; color: #808080; font-style: italic;'>服务周期：{start_date.year}年{start_date.month}月</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div style='text-align: right; color: #808080; font-style: italic;'>服务周期：{start_date.year}年{start_date.month}月至{end_date.year}年{end_date.month}月</div>", unsafe_allow_html=True)

    # 添加一条横线
    st.markdown("---")
    # 1. 工单统计
    total_tickets = ticket_summary['total'].iloc[0] if not ticket_summary.empty else 0
    if start_date.month == end_date.month:
        st.write(f"#### {start_date.year}年{start_date.month}月iTop共接收工单数 {total_tickets} 起，各类工单处理情况如下：")
    else:
        st.write(f"#### {start_date.year}年{start_date.month}月至{end_date.year}年{end_date.month}月iTop共接收工单数 {total_tickets} 个，各类工单处理情况如下：")

    # 2. 按服务类型统计分析
    st.write("#### 1. 按服务类型统计分析如下：")

    # 2.1 服务请求统计
    st.write("##### 1) 服务请求统计")
    if not user_request_stats.empty:
        total = user_request_stats['total'].iloc[0]
        resolved = user_request_stats['resolved_total'].iloc[0]
        closed = user_request_stats['closed_total'].iloc[0]
        unresolved = user_request_stats['unresolved_total'].iloc[0]

        if total and total > 0:
            resolved_percentage = resolved / total * 100 if total > 0 else 0
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100 if total > 0 else 0

            st.write(f"""
                     本周期内共接收服务请求 {total:g} 个，其中 {resolved:g} 个服务请求被解决，占比约 {resolved_percentage:.2f}%；
                     已解决的服务请求中，{closed:g} 个服务请求被按时关闭，占比约 {closed_percentage:.2f}%；
                     未解决的服务请求有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。
                     """)

            # 饼图：服务请求状态
            fig = px.pie(names=['已解决', '未解决', '已关闭'], 
                         values=[resolved, unresolved, closed], 
                         title="服务请求状态分布", 
                         color_discrete_sequence=['#00b8a9', '#f6416c', '#f8f3d4'])
            fig.update_traces(textposition='inside', 
                            textinfo='label+percent')
            fig.update_layout(title_x=0.35)  # 将标题居中显示
            st.plotly_chart(fig)
        else:
            st.write("本周期内没有接收到服务请求。")
    else:
        st.write("无法获取服务请求统计数据。")

    # 2.2 事件统计
    st.write("##### 2) 事件统计")
    if not incident_stats.empty:
        total = incident_stats['total'].iloc[0]
        resolved = incident_stats['resolved_total'].iloc[0]
        closed = incident_stats['closed_total'].iloc[0]
        unresolved = incident_stats['unresolved_total'].iloc[0]

        if total and total > 0:
            resolved_percentage = resolved / total * 100 if total > 0 else 0
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100 if total > 0 else 0

            st.write(f"""
                     本周期内共发生事件 {total:g} 个，其中 {resolved:g} 个事件被解决，占比约 {resolved_percentage:.2f}%；
                     已解决的事件中，{closed:g} 个事件被按时关闭，占比约 {closed_percentage:.2f}%；
                     未解决的事件有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。
                     """)

            # 饼图：事件状态
            fig = px.pie(names=['已解决', '未解决', '已关闭'], 
                         values=[resolved, unresolved, closed], 
                         title="事件状态分布", 
                         color_discrete_sequence=['#00b8a9', '#f6416c', '#f8f3d4'])
            fig.update_traces(textposition='inside', 
                            textinfo='label+percent')
            fig.update_layout(title_x=0.35)  # 将标题居中显示
            st.plotly_chart(fig)
        else:
            st.write("本周期内没有发生事件。")
    else:
        st.write("无法获取事件统计数据。")

    # 2.3 变更统计
    st.write("##### 3) 变更统计")
    if not change_stats.empty:
        total = change_stats['total'].iloc[0]
        resolved = change_stats['resolved_total'].iloc[0]
        closed = change_stats['closed_total'].iloc[0]

        if total and total > 0:
            closed_percentage = closed / total * 100 if total > 0 else 0
            resolved_percentage = resolved / closed * 100 if closed > 0 else 0
            
            st.write(f"""
                     本周期内共发生变更 {total:g} 个，其中 {closed:g} 个变更已关闭，占比约 {closed_percentage:.2f}%。
                     已关闭的变更中，{resolved:g} 个变更被成功执行，占比约 {resolved_percentage:.2f}%。
                     """)

            # 饼图：变更状态
            fig = px.pie(names=['已解决', '未解决'], 
                        values=[resolved, total-resolved], 
                        title="变更状态分布", 
                        color_discrete_sequence=['#00b8a9', '#f6416c'])
            fig.update_traces(textposition='inside', 
                            textinfo='label+percent')
            fig.update_layout(title_x=0.35)  # 将标题居中显示
            st.plotly_chart(fig)
        else:
            st.write("本周期内没有发生变更。")
    else:
        st.write("无法获取变更统计数据。")

    # 3. 按照工单处理团队统计
    st.write("#### 2. 按照工单处理团队统计，具体如下")
    st.dataframe(team_stats, use_container_width=True)

    # 3.1 按照工单处理团队绘制服务请求的解决率
    # 将team_stats转换为pandas DataFrame
    df = pd.DataFrame(team_stats)
    
    # 按月份和团队分组计算平均解决率和及时率
    df['工单解决率'] = df['工单解决率'].apply(lambda x: float(str(x).rstrip('%')))
    
    # 检查是否跨月
    if len(df['月份'].unique()) > 1:
        # 仅保留服务请求数据
        service_request_df = df[df['工单类型'] == '服务请求']
        # 创建解决率曲线图
        fig1 = px.line(service_request_df, 
                      x='月份', 
                      y='工单解决率',
                      color='团队',
                      markers=True,
                      text='工单解决率',  # 添加数值标签
                      title='各团队服务请求月度解决率趋势')
        
        # 配置数值标签的显示
        fig1.update_traces(
            textposition="top center",  # 将数值显示在点的上方居中
            texttemplate='%{text:.1f}%'  # 显示格式:保留1位小数并加上%号
        )
        
        fig1.update_layout(
            title_x=0.35,
            title_y=0.95, # 将标题向上移动
            xaxis_title='月份',
            yaxis_title='解决率(%)',
            yaxis=dict(range=[0, 110]),
            xaxis=dict(
                type='category',
                categoryorder='category ascending'
            ),
            margin=dict(t=100), # 增加顶部边距
            legend=dict(
                orientation="h",  # 水平方向
                yanchor="bottom",
                y=1.05,  # 调整图例位置,与标题保持10px间距
                xanchor="center",
                x=0.5,  # 图例水平居中
                itemwidth=30,  # 设置图例项的宽度,使团队名称显示在一行
                title=None  # 取消图例标题
            )
        )
        st.plotly_chart(fig1)

    # 4. 按照工程师统计
    st.write("#### 3. 按照工单处理工程师统计，具体如下")
    st.dataframe(person_stats, use_container_width=True)

    # 5. 未解决的工单
    st.write("#### 4. 未解决的工单如下")
    st.dataframe(unresolved_tickets, use_container_width=True)

    # 6. 超时的工单
    st.write("#### 5. SLA超时的工单如下")
    st.dataframe(overdue_tickets, use_container_width=True)

if __name__ == "__main__":
    main()
