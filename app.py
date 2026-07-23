import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import warnings
import re

st.set_page_config(page_title="📈 多板块分时图分析器 Pro", page_icon="📊", layout="wide")

# ---------- 1. 初始化 Session State ----------
if "data_dict" not in st.session_state:
    st.session_state.data_dict = {
        "上证主板": None,
        "深圳主板": None,
        "科创综指": None,
        "创业板指": None
    }

if "prev_close_dict" not in st.session_state:
    st.session_state.prev_close_dict = {
        "上证主板": None,
        "深圳主板": None,
        "科创综指": None,
        "创业板指": None
    }

# 热点事件存储
if "pending_events" not in st.session_state:
    st.session_state.pending_events = [] 

if "submitted_flag" not in st.session_state:
    st.session_state.submitted_flag = False
if "submitted_events" not in st.session_state:
    st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件描述"])
if "submitted_stats" not in st.session_state:
    st.session_state.submitted_stats = {
        "turnover": "", "change": "", "up_cnt": "", "down_cnt": "", "sectors": ""
    }

# ---------- 2. 工具函数：智能提取时间格式 ----------
def normalize_time_to_hm(t_str):
    m = re.search(r'(\d{1,2})[:：](\d{2})', str(t_str))
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return str(t_str).strip()

def is_valid_trading_time(t_str):
    hm = normalize_time_to_hm(t_str)
    if "11:30" < hm < "13:00":
        return False
    return True

# ---------- 3. 智能解析 Excel ----------
@st.cache_data
def parse_uploaded_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='gbk')
        else:
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            except Exception as e:
                if 'NamedCellStyle' in str(e):
                    from openpyxl import load_workbook
                    warnings.filterwarnings('ignore')
                    wb = load_workbook(uploaded_file, data_only=True, read_only=True)
                    ws = wb.active
                    data_rows = list(ws.values)
                    if not data_rows:
                        return None
                    columns = data_rows[0]
                    values = data_rows[1:]
                    df = pd.DataFrame(values, columns=columns)
                    wb.close()
                else:
                    raise e
        
        if df.empty: return None

        cols = df.columns.tolist()
        time_col, price_col, vol_col = None, None, None
        for col in cols:
            col_lower = str(col).lower()
            if '时间' in col_lower or 'time' in col_lower: time_col = col
            elif '收盘' in col_lower or 'close' in col_lower or '价格' in col_lower or '指数' in col_lower or '点位' in col_lower: price_col = col
            elif '成交' in col_lower or 'volume' in col_lower or '量' in col_lower: vol_col = col

        if time_col is None and len(cols) >= 1: time_col = cols[0]
        if price_col is None and len(cols) >= 2: price_col = cols[1]
        if vol_col is None and len(cols) >= 3: vol_col = cols[2]

        if time_col is None or price_col is None: return None

        result_df = pd.DataFrame()
        result_df['时间'] = df[time_col].astype(str).str.strip()
        result_df['价格'] = pd.to_numeric(df[price_col], errors='coerce')
        result_df['成交量'] = pd.to_numeric(df[vol_col], errors='coerce').fillna(0) if vol_col else 0

        result_df = result_df.dropna(subset=['价格'])
        result_df = result_df[result_df['时间'].str.contains(':', na=False)]
        result_df = result_df[result_df['时间'].apply(is_valid_trading_time)]

        return result_df if not result_df.empty else None

    except Exception:
        return None


# ---------- 4. 核心绘图函数（红绿双色分时） ----------
def draw_chart(data_df, events_df, index_name, prev_close=None):
    if data_df is None or data_df.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"暂无 {index_name} 数据", x=0.5, y=0.5, showarrow=False, font=dict(size=20, color="gray"))
        fig.update_layout(height=550) 
        return fig

    # 确定那条“横线”在哪里 (昨收或开盘价)
    baseline = prev_close if prev_close and prev_close > 0 else data_df['价格'].iloc[0]
    prices = data_df['价格'].tolist()
    times = data_df['时间'].tolist()
    volumes = data_df['成交量'].tolist()

    # 高于横线标红，低于横线标绿
    y_red = []
    y_green = []
    for p in prices:
        if p >= baseline:
            y_red.append(p)
            y_green.append(None)
        else:
            y_red.append(None)
            y_green.append(p)

    # 简单处理交界处相连
    for i in range(1, len(prices)):
        p_prev = prices[i-1]
        p_curr = prices[i]
        if p_prev >= baseline and p_curr < baseline:
            y_green[i-1] = p_prev  
        elif p_prev < baseline and p_curr >= baseline:
            y_red[i-1] = p_prev    

    # 柱状图也做相同的红绿处理
    vol_colors = ['#ff4d4f' if p >= baseline else '#3ecf8e' for p in prices]

    fig = make_subplots(
        rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.08,
        row_heights=[0.65, 0.35], subplot_titles=(f"{index_name} 指数走势", "成交量")
    )

    # 画红色段
    fig.add_trace(go.Scatter(x=times, y=y_red, mode='lines', name='涨', line=dict(color='#ff4d4f', width=2.5), connectgaps=False), row=1, col=1)
    # 画绿色段
    fig.add_trace(go.Scatter(x=times, y=y_green, mode='lines', name='跌', line=dict(color='#3ecf8e', width=2.5), connectgaps=False), row=1, col=1)

    # ⭐️ 画出你说的那条明亮的基准“横线”
    baseline_label = "昨日收盘" if (prev_close and prev_close > 0) else "今日开盘"
    fig.add_hline(
        y=baseline, line_dash="dash", line_color="rgba(255,255,255,0.7)", line_width=1.5,
        annotation_text=f"{baseline_label}: {baseline:.2f}", 
        annotation_position="bottom right", annotation_font_color="white",
        row=1, col=1
    )

    # 成交量
    fig.add_trace(go.Bar(x=times, y=volumes, marker_color=vol_colors, opacity=0.8, showlegend=False), row=2, col=1)

    # 绘制热点事件
    if events_df is not None and not events_df.empty:
        time_map = {normalize_time_to_hm(t): t for t in times}
        for _, event in events_df.iterrows():
            event_time = str(event['时间']).strip()
            event_desc = str(event['事件描述']).strip()
            norm_ev_time = normalize_time_to_hm(event_time)
            if norm_ev_time and event_desc and norm_ev_time in time_map:
                actual_x = time_map[norm_ev_time]
                fig.add_vline(
                    x=actual_x, line_dash="dash", line_color="orange", line_width=1.5,
                    annotation_text=event_desc, annotation_position="top left",
                    annotation_font_size=12, annotation_font_color="orange",
                    row=1, col=1
                )

    # 布局更新
    fig.update_layout(
        height=550, hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template='plotly_dark', margin=dict(l=20, r=20, t=40, b=20)
    )

    if not data_df.empty:
        min_price, max_price = min(prices), max(prices)
        padding = (max_price - min_price) * 0.1 if max_price > min_price else 10
        y_min_bound, y_max_bound = min(min_price, baseline) - padding, max(max_price, baseline) + padding
        fig.update_yaxes(title_text="点位", row=1, col=1, range=[y_min_bound, y_max_bound])
    
    fig.update_yaxes(title_text="成交量", row=2, col=1)
    
    # 强制类目轴
    fig.update_xaxes(type='category', row=1, col=1)
    fig.update_xaxes(title_text="时间", row=2, col=1, type='category', tickangle=45, tickfont=dict(size=10), nticks=15)

    return fig

# ---------- 5. UI及其他逻辑 (保持不变) ----------
def render_index_overview(edit_mode=True):
    st.markdown("### 📊 今日指数概览")
    cols_overview = st.columns(4)

    for i, (col, key) in enumerate(zip(cols_overview, st.session_state.data_dict.keys())):
        with col:
            data = st.session_state.data_dict.get(key)
            current_prev = st.session_state.prev_close_dict.get(key, None)
            
            if data is not None and not data.empty:
                close_price = data['价格'].iloc[-1]
                open_price = data['价格'].iloc[0]
                
                if current_prev is not None and current_prev > 0:
                    change_pct = ((close_price - current_prev) / current_prev) * 100
                    calc_label = f"昨收 {current_prev:.2f}"
                else:
                    change_pct = ((close_price - open_price) / open_price) * 100 if open_price != 0 else None
                    calc_label = "⚠️ 请输昨收，暂基于开盘价" if edit_mode else "基于开盘价计算"
                
                color = "#ff4d4f" if (change_pct is not None and change_pct >= 0) else "#3ecf8e"
                arrow = "▲" if (change_pct is not None and change_pct >= 0) else "▼"
                change_display = f"{arrow} {abs(change_pct):.2f}%" if change_pct is not None else "N/A"
                
                st.markdown(f"""
                <div style="background: rgba(255,255,255,0.04); border-radius: 12px; padding: 12px 14px; border: 1px solid #1e2a3a;">
                    <div style="color: #7a8ba3; font-size: 13px; font-weight: 500;">{key}</div>
                    <div style="color: #e8edf5; font-size: 22px; font-weight: 700;">{close_price:.2f}</div>
                    <div style="color: {color}; font-size: 15px; font-weight: 600;">{change_display}</div>
                    <div style="color: #4a5a6e; font-size: 11px;">{calc_label}</div>
                </div>
                """, unsafe_allow_html=True)
                
                if edit_mode:
                    prev_input = st.number_input(
                        "昨日收盘", value=None if current_prev is None else float(current_prev),
                        step=1.0, format="%.2f", key=f"prev_input_{key}",
                        label_visibility="collapsed", placeholder="输入昨收",
                    )
                    if prev_input != current_prev:
                        st.session_state.prev_close_dict[key] = prev_input if prev_input is not None and prev_input > 0 else None
                        st.rerun()
            else:
                st.markdown(f"""
                <div style="background: rgba(255,255,255,0.02); border-radius: 12px; padding: 12px 14px; border: 1px dashed #2a3a4a;">
                    <div style="color: #7a8ba3; font-size: 13px;">{key}</div>
                    <div style="color: #4a5a6e; font-size: 16px; padding: 8px 0;">⏳ 暂无数据</div>
                </div>
                """, unsafe_allow_html=True)

if st.session_state.submitted_flag:
    col_t1, col_t2 = st.columns([5, 1])
    with col_t1:
        st.markdown(f"<h1 style='color:#e8edf5;'>📄 市场复盘分析报告 <span style='font-size:16px; color:#7a8ba3;'>{datetime.now().strftime('%Y-%m-%d')}</span></h1>", unsafe_allow_html=True)
    with col_t2:
        st.write("") 
        if st.button("⬅️ 返回修改", use_container_width=True):
            st.session_state.submitted_flag = False
            if not st.session_state.submitted_events.empty:
                st.session_state.pending_events = st.session_state.submitted_events.to_dict('records')
            st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件描述"])
            st.rerun()
    
    st.markdown("---")
    render_index_overview(edit_mode=False)

    st.markdown("<br>### 📈 核心市场统计", unsafe_allow_html=True)
    stats = st.session_state.submitted_stats
    st.markdown(f"""
    <div style="display: flex; justify-content: space-around; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 25px 10px; border: 1px solid #1e2a3a;">
        <div style="text-align: center;"><div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">成交额</div><div style="color: #e8edf5; font-size: 26px; font-weight: 700;">{stats.get('turnover', '-')}</div></div>
        <div style="text-align: center;"><div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">较前日增减</div><div style="color: #e8edf5; font-size: 26px; font-weight: 700;">{stats.get('change', '-')}</div></div>
        <div style="text-align: center;"><div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">上涨家数</div><div style="color: #ff4d4f; font-size: 26px; font-weight: 700;">{stats.get('up_cnt', '-')}</div></div>
        <div style="text-align: center;"><div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">下跌家数</div><div style="color: #3ecf8e; font-size: 26px; font-weight: 700;">{stats.get('down_cnt', '-')}</div></div>
        <div style="text-align: center;"><div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">涨幅居前板块</div><div style="color: #f5c542; font-size: 22px; font-weight: 700;">{stats.get('sectors', '-')}</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>### 📊 各板块分时走势（含热点事件）", unsafe_allow_html=True)
    index_keys = list(st.session_state.data_dict.keys())
    tabs = st.tabs(index_keys)
    for tab, key in zip(tabs, index_keys):
        with tab:
            data = st.session_state.data_dict.get(key)
            current_prev = st.session_state.prev_close_dict.get(key, None)
            fig = draw_chart(data, st.session_state.submitted_events, key, current_prev)
            st.plotly_chart(fig, width='stretch', use_container_width=True, key=f"final_chart_{key}")

else:
    st.markdown("<h1 style='color:#e8edf5; border-left: 4px solid #ff4d4f; padding-left: 16px;'>📈 多板块分时图分析器 <span style='font-size:16px; color:#7a8ba3;'>工作台模式</span></h1>", unsafe_allow_html=True)
    render_index_overview(edit_mode=True)
    st.markdown("### 📤 第一步：分别上传四个板块的分时数据（Excel/CSV）")
    cols_upload = st.columns(4)
    index_keys = list(st.session_state.data_dict.keys())

    for i, (col, key) in enumerate(zip(cols_upload, index_keys)):
        with col:
            status = "✅" if st.session_state.data_dict[key] is not None else "⬜"
            st.markdown(f"**{status} {key}**")
            uploaded_file = st.file_uploader(f"上传 {key} 数据", type=['xlsx', 'xls', 'csv'], label_visibility="collapsed", key=f"upload_{key}")
            
            if uploaded_file is not None:
                df = parse_uploaded_file(uploaded_file)
                if df is not None:
                    st.session_state.data_dict[key] = df
                    st.success(f"✅ 解析成功")
                else:
                    st.session_state.data_dict[key] = None

    col_chart, col_right = st.columns([2.2, 1])

    with col_chart:
        st.markdown("### 📊 第二步：图表实时预览")
        display_events = pd.DataFrame(st.session_state.pending_events) if st.session_state.pending_events else pd.DataFrame(columns=["时间", "事件描述"])
        st.info("✏️ 预览模式：在上方输入昨收即可自动生成横线变色，在右侧添加事件后可实时预览。")

        tabs = st.tabs(index_keys)
        for tab, key in zip(tabs, index_keys):
            with tab:
                data = st.session_state.data_dict.get(key)
                current_prev = st.session_state.prev_close_dict.get(key, None)
                fig = draw_chart(data, display_events, key, current_prev)
                st.plotly_chart(fig, width='stretch', use_container_width=True)

    with col_right:
        st.subheader("⏱️ 热点事件编辑")
        col_time, col_desc, col_btn = st.columns([1.2, 2.5, 0.8])
        with col_time: new_time = st.text_input("时间", placeholder="09:30", key="new_event_time", label_visibility="collapsed")
        with col_desc: new_desc = st.text_input("事件描述", placeholder="输入事件", key="new_event_desc", label_visibility="collapsed")
        with col_btn:
            st.write(""); st.write("") 
            if st.button("➕", use_container_width=True, help="添加事件"):
                if new_time.strip() and new_desc.strip():
                    st.session_state.pending_events.append({"时间": new_time.strip(), "事件描述": new_desc.strip()})
                    st.rerun()
                else: st.warning("填写不完整")
        
        if st.session_state.pending_events:
            st.markdown("**📋 待提交事件列表**")
            for idx, event in enumerate(st.session_state.pending_events):
                c1, c2, c3 = st.columns([1.2, 2.5, 0.8])
                with c1: st.text(event["时间"])
                with c2: st.text(event["事件描述"])
                with c3:
                    if st.button("✕", key=f"del_{idx}"):
                        del st.session_state.pending_events[idx]; st.rerun()

        st.markdown("---")
        st.subheader("📊 市场统计编辑")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            turnover = st.text_input("成交额", placeholder="27182.89亿", key="stat_turnover")
            up_cnt = st.text_input("上涨家数", placeholder="1740", key="stat_up")
        with col_s2:
            change = st.text_input("较前日增减", placeholder="+464.24", key="stat_change")
            down_cnt = st.text_input("下跌家数", placeholder="3710", key="stat_down")
        sectors = st.text_input("涨幅居前板块", placeholder="油气 · 煤炭 · 白酒", key="stat_sectors")

        st.markdown("---")
        if st.button("🚀 生成复盘报告", use_container_width=True, type="primary"):
            st.session_state.submitted_stats = {
                "turnover": st.session_state.get("stat_turnover", ""), "change": st.session_state.get("stat_change", ""),
                "up_cnt": st.session_state.get("stat_up", ""), "down_cnt": st.session_state.get("stat_down", ""),
                "sectors": st.session_state.get("stat_sectors", "")
            }
            st.session_state.submitted_events = pd.DataFrame(st.session_state.pending_events) if st.session_state.pending_events else pd.DataFrame(columns=["时间", "事件描述"])
            st.session_state.submitted_flag = True
            st.rerun()

st.markdown("---")
st.caption(f"🕒 数据仅在本地处理 ｜ 当前时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
