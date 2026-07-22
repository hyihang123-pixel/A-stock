import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import warnings

st.set_page_config(page_title="📈 多板块分时图分析器 Pro", page_icon="📊", layout="wide")

# ---------- 1. 初始化 Session State ----------
# 核心改动：用字典存储四个板块的数据
if "data_dict" not in st.session_state:
    st.session_state.data_dict = {
        "上证主板": None,
        "深圳主板": None,
        "科创综指": None,
        "创业板指": None
    }

if "events" not in st.session_state:
    st.session_state.events = pd.DataFrame({
        "时间": ["09:40", "10:30"],
        "事件描述": ["示例事件A", "示例事件B"]
    })

# ---------- 2. 智能解析 Excel（强化版，自动忽略样式错误） ----------
@st.cache_data
def parse_uploaded_file(uploaded_file):
    """读取并解析上传的文件，自动忽略损坏的样式格式"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='gbk')
        else:
            # 正常读取（先试一下）
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            except Exception as e:
                # 如果报错且包含 "NamedCellStyle"（样式错误），则启动“只读裸奔模式”
                if 'NamedCellStyle' in str(e):
                    from openpyxl import load_workbook
                    warnings.filterwarnings('ignore')
                    
                    # 关键：read_only=True 会跳过样式解析，data_only=True 拿数值
                    wb = load_workbook(uploaded_file, data_only=True, read_only=True)
                    ws = wb.active
                    
                    data_rows = list(ws.values)
                    if not data_rows:
                        st.error("❌ Excel 中无数据")
                        return None
                    
                    columns = data_rows[0]
                    values = data_rows[1:]
                    df = pd.DataFrame(values, columns=columns)
                    wb.close()
                else:
                    raise e
        
        if df.empty:
            return None

        # ---------- 智能列名匹配 ----------
        cols = df.columns.tolist()
        time_col, price_col, vol_col = None, None, None
        for col in cols:
            col_lower = str(col).lower()
            if '时间' in col_lower or 'time' in col_lower:
                time_col = col
            elif '收盘' in col_lower or 'close' in col_lower or '价格' in col_lower or '指数' in col_lower or '点位' in col_lower:
                price_col = col
            elif '成交' in col_lower or 'volume' in col_lower or '量' in col_lower:
                vol_col = col

        if time_col is None and len(cols) >= 1:
            time_col = cols[0]
        if price_col is None and len(cols) >= 2:
            price_col = cols[1]
        if vol_col is None and len(cols) >= 3:
            vol_col = cols[2]

        if time_col is None or price_col is None:
            st.error("❌ 无法识别列：请确保包含'时间'和'收盘价'列")
            return None

        result_df = pd.DataFrame()
        result_df['时间'] = df[time_col].astype(str).str.strip()
        result_df['价格'] = pd.to_numeric(df[price_col], errors='coerce')
        if vol_col:
            result_df['成交量'] = pd.to_numeric(df[vol_col], errors='coerce').fillna(0)
        else:
            result_df['成交量'] = 0

        result_df = result_df.dropna(subset=['价格'])
        result_df = result_df[result_df['时间'].str.contains(':', na=False)]

        if result_df.empty:
            st.error("❌ 解析后无有效数据，请确认时间格式为 HH:MM")
            return None

        return result_df

    except Exception as e:
        st.error(f"❌ 文件解析失败: {e}")
        return None

# ---------- 3. 核心绘图函数 ----------
def draw_chart(data_df, events_df, index_name):
    """使用 Plotly 绘制双轴分时图 + 事件标记"""
    if data_df is None or data_df.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"请上传 {index_name} 数据", x=0.5, y=0.5, showarrow=False, font=dict(size=20, color="gray"))
        fig.update_layout(height=500)
        return fig

    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        row_heights=[0.65, 0.35],
        subplot_titles=(f"{index_name} 指数走势", "成交量")
    )

    # 指数折线
    fig.add_trace(
        go.Scatter(
            x=data_df['时间'], y=data_df['价格'],
            mode='lines', name='指数点位',
            line=dict(color='#ff4d4f', width=2.5),
            fill='tozeroy', fillcolor='rgba(255,77,79,0.1)'
        ),
        row=1, col=1
    )

    # 成交量柱状
    fig.add_trace(
        go.Bar(
            x=data_df['时间'], y=data_df['成交量'],
            name='成交量', marker_color='#3b82f6', opacity=0.7
        ),
        row=2, col=1
    )

    # 事件标记
    if events_df is not None and not events_df.empty:
        for _, event in events_df.iterrows():
            event_time = str(event['时间']).strip()
            event_desc = str(event['事件描述']).strip()
            if event_time and event_desc and event_time in data_df['时间'].values:
                fig.add_vline(
                    x=event_time, line_dash="dash", line_color="orange", line_width=1.5,
                    annotation_text=event_desc, annotation_position="top left",
                    annotation_font_size=11, annotation_font_color="orange",
                    annotation_bgcolor="rgba(0,0,0,0.6)",
                    row=1, col=1
                )

    fig.update_layout(
        height=520, hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template='plotly_dark',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=40, b=20)
    )

    # 纵坐标范围：为了看清走势，根据数据自动调整，留出上下边距
    if not data_df.empty:
        min_price = data_df['价格'].min()
        max_price = data_df['价格'].max()
        padding = (max_price - min_price) * 0.1 if max_price > min_price else 10
        fig.update_yaxes(title_text="点位", row=1, col=1, range=[min_price - padding, max_price + padding])
    
    fig.update_yaxes(title_text="成交量", row=2, col=1)
    fig.update_xaxes(title_text="时间", row=2, col=1, tickangle=45, tickfont=dict(size=10))

    return fig

# ---------- 4. 页面UI布局 ----------
st.markdown("<h1 style='color:#e8edf5; border-left: 4px solid #ff4d4f; padding-left: 16px;'>📈 多板块分时图分析器 <span style='font-size:16px; color:#7a8ba3;'>上传数据 · 点击切换板块</span></h1>", unsafe_allow_html=True)

# ---------- 第一行：四个上传区域（并排） ----------
st.markdown("### 📤 第一步：分别上传四个板块的数据")
cols_upload = st.columns(4)
index_keys = list(st.session_state.data_dict.keys())

for i, (col, key) in enumerate(zip(cols_upload, index_keys)):
    with col:
        # 显示当前是否已上传的状态标记
        status = "✅" if st.session_state.data_dict[key] is not None else "⬜"
        st.markdown(f"**{status} {key}**")
        
        uploaded_file = st.file_uploader(
            f"上传 {key} 数据",
            type=['xlsx', 'xls', 'csv'],
            label_visibility="collapsed",
            key=f"upload_{key}"  # 每个上传器必须唯一 Key
        )
        
        if uploaded_file is not None:
            df = parse_uploaded_file(uploaded_file)
            if df is not None:
                st.session_state.data_dict[key] = df
                st.success(f"✅ {len(df)} 条数据")
            else:
                st.session_state.data_dict[key] = None
                st.error("解析失败")

# ---------- 第二行：图表 + 右侧面板 ----------
col_chart, col_right = st.columns([2.2, 1])

with col_chart:
    # 核心改动：使用 Tabs（标签页）作为“点击切换”的按钮
    st.markdown("### 📊 第二步：点击标签切换板块视图")
    tabs = st.tabs(index_keys)  # 生成四个标签页按钮
    
    # 循环填充每个标签页
    for tab, key in zip(tabs, index_keys):
        with tab:
            data = st.session_state.data_dict.get(key)
            # 获取当前板块的事件（事件是全局共享的，也可以做成独立的，但通常市场事件通用）
            fig = draw_chart(data, st.session_state.events, key)
            st.plotly_chart(fig, width='stretch', use_container_width=True)
            
            # 如果没数据，显示提示
            if data is None:
                st.info(f"👆 请先在顶部上传 {key} 的 Excel 文件")

with col_right:
    # ---------- 热点事件管理（全局通用） ----------
    st.subheader("⏱️ 热点事件 (全局)")
    edited_events = st.data_editor(
        st.session_state.events,
        num_rows="dynamic",
        width='stretch',
        column_config={
            "时间": st.column_config.TextColumn("时间", help="格式: 09:30"),
            "事件描述": st.column_config.TextColumn("事件描述", help="输入热点事件")
        },
        key="events_editor"
    )
    if not edited_events.equals(st.session_state.events):
        st.session_state.events = edited_events
        st.rerun()
    st.caption("💡 点击表格下方 '添加行' 增加事件，删除行会自动移除")

    # ---------- 底部统计信息 ----------
    st.markdown("---")
    st.subheader("📊 市场统计")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        turnover = st.text_input("成交额（亿元）", placeholder="27182.89", key="stat_turnover")
        up_cnt = st.text_input("上涨家数", placeholder="1740", key="stat_up")
    with col_s2:
        change = st.text_input("较前日增减", placeholder="+464.24", key="stat_change")
        down_cnt = st.text_input("下跌家数", placeholder="3710", key="stat_down")
    
    sectors = st.text_input("涨幅居前板块", placeholder="油气 · 煤炭 · 白酒", key="stat_sectors")

# ---------- 底部时间戳 ----------
st.markdown("---")
col_f1, col_f2 = st.columns([3, 1])
with col_f1:
    st.caption(f"🕒 数据来源：手动上传 ｜ 当前时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
with col_f2:
    st.caption("📌 所有数据仅在本地处理，不上传服务器")
