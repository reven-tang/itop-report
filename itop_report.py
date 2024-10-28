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
        c.name AS '团队',
        ticket_type AS '工单类型',
        COUNT(*) AS '工单数量',
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
            t.team_id,
            '服务请求' AS ticket_type,
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
            '事件' AS ticket_type,
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
            '变更' AS ticket_type,
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
    GROUP BY c.name, ticket_type
    HAVING COUNT(*) > 0
    ORDER BY ticket_type desc, c.name
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
        ai.agent_name AS '办理人',
        ticket_type AS '工单类型',
        COUNT(*) AS '工单数量',
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
            NULL AS response_time,
            TIMESTAMPDIFF(SECOND, t.start_date, t.end_date) AS resolution_time
        FROM `change` c2
        JOIN ticket t ON c2.id = t.id
        WHERE c2.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
    ) AS subquery
    JOIN agent_info ai ON subquery.agent_id = ai.id
    GROUP BY ai.agent_name, ticket_type
    ORDER BY ticket_type desc, ai.agent_name
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
        title = f"<para alignment='center'>iTop 运维服务月报 ({start_date.year}年{start_date.month}月)</para>"
    else:
        title = f"<para alignment='center'>iTop 运维服务月报 ({start_date.year}年{start_date.month}月至{end_date.year}年{end_date.month}月)</para>"
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
        pie.labels = ['' for _ in labels]  # 清空饼图上的标签
        pie.slices.strokeWidth = 0.5

        # 设置颜色
        colors = [HexColor('#00b8a9'), HexColor('#f6416c'), HexColor('#f8f3d4')]
        for i, color in enumerate(colors[:len(data)]):
            pie.slices[i].fillColor = color

        drawing.add(pie)

        # 添加标题
        title_label = String(200, 240, title)
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
            
            # 添加一行空行
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
            fig = px.pie(names=['已解决', '未解决', '已关闭'], values=[resolved, unresolved, closed], title="服务请求状态分布", color_discrete_sequence=['#00b8a9', '#f6416c', '#f8f3d4'])
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
            fig = px.pie(names=['已解决', '未解决', '已关闭'], values=[resolved, unresolved, closed], title="事件状态分布", color_discrete_sequence=['#00b8a9', '#f6416c', '#f8f3d4'])
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
            fig = px.pie(names=['已解决', '未解决'], values=[resolved, total-resolved], title="变更状态分布", color_discrete_sequence=['#00b8a9', '#f6416c'])
            fig.update_layout(title_x=0.35)  # 将标题居中显示
            st.plotly_chart(fig)
        else:
            st.write("本周期内没有发生变更。")
    else:
        st.write("无法获取变更统计数据。")

    # 3. 按照工单处理团队统计
    st.write("#### 2. 按照工单处理团队统计，具体如下")
    st.dataframe(team_stats, use_container_width=True)

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