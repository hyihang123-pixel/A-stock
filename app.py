import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import warnings
import re # 新增：用于正则匹配时间格式

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

# ---------- 热点事件存储 ----------
if "pending_events" not in st.session_state:
    st.session_state.pending_events = []  # 元素格式: {"时间": "09:40", "事件描述": "xxx"}

if "submitted_flag" not in st.session_state:
    st.session_state.submitted_flag = False
if "submitted_events" not in st.session_state:
    st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件描述"])
if "submitted_stats" not in st.session_state:
    st.session_state.submitted_stats = {
        "turnover": "", "change": "", "up_cnt": "", "down_cnt": "", "sectors": ""
    }

# ---------- 2. 智能解析 Excel ----------
@st.cache_data
def parse_uploaded_file(uploaded_file):
    """读取并解析上传的文件，自动忽略损坏的样式格式"""
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

# ---------- 3. 智能提取时间格式 (新增工具函数) ----------
def normalize_time_to_hm(t_str):
    """将任何包含时间的字符串标准化为 HH:MM 格式，兼容中文冒号和单数小时"""
    m = re.search(r'(\d{1,2})[:：](\d{2})', str(t_str))
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return str(t_str).strip()

# ---------- 4. 核心绘图函数 ----------
def draw_chart(data_df, events_df, index_name):
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

    fig.add_trace(
        go.Scatter(
            x=data_df['时间'], y=data_df['价格'],
            mode='lines', name='指数点位',
            line=dict(color='#ff4d4f', width=2.5),
            fill='tozeroy', fillcolor='rgba(255,77,79,0.1)'
        ),
        row=1, col=1
    )

    fig.add_trace(
        go.Bar(
            x=data_df['时间'], y=data_df['成交量'],
            name='成交量', marker_color='#3b82f6', opacity=0.7
        ),
        row=2, col=1
    )

    # 绘制热点事件逻辑化修改：使用标准化时间进行匹配映射
    if events_df is not None and not events_df.empty:
        # 创建一个 规范时间 HH:MM -> 数据框实际时间字符串 的映射字典
        time_map = {normalize_time_to_hm(t): t for t in data_df['时间'].values}
        
        for _, event in events_df.iterrows():
            event_time = str(event['时间']).strip()
            event_desc = str(event['事件描述']).strip()
            
            # 将用户输入的事件时间标准化 (例如 9:30 或者 09：30 都转为 09:30)
            norm_ev_time = normalize_time_to_hm(event_time)
            
            # 如果标准化时间在时间映射中找得到对应的值
            if norm_ev_time and event_desc and norm_ev_time in time_map:
                actual_x_in_df = time_map[norm_ev_time] # 获取数据框内对应时间原格式
                fig.add_vline(
                    x=actual_x_in_df, line_dash="dash", line_color="orange", line_width=1.5,
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

    if not data_df.empty:
        min_price = data_df['价格'].min()
        max_price = data_df['价格'].max()
        padding = (max_price - min_price) * 0.1 if max_price > min_price else 10
        fig.update_yaxes(title_text="点位", row=1, col=1, range=[min_price - padding, max_price + padding])
    
    fig.update_yaxes(title_text="成交量", row=2, col=1)
    fig.update_xaxes(title_text="时间", row=2, col=1, tickangle=45, tickfont=dict(size=10))

    return fig

# ---------- 5. 页面UI布局 ----------
st.markdown("<h1 style='color:#e8edf5; border-left: 4px solid #ff4d4f; padding-left: 16px;'>📈 多板块分时图分析器 <span style='font-size:16px; color:#7a8ba3;'>上传数据 · 提交生成</span></h1>", unsafe_allow_html=True)

# ============================================================
# 📊 今日指数概览（含昨日收盘价输入框）
# ============================================================
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
                calc_label = "⚠️ 请输昨收"
            
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
            
            prev_input = st.number_input(
                "昨日收盘",
                value=None if current_prev is None else float(current_prev),
                step=1.0,
                format="%.2f",
                key=f"prev_input_{key}",
                label_visibility="collapsed",
                placeholder="输入昨收",
                help="输入昨日收盘价，用于精确计算涨跌幅"
            )
            
            if prev_input != current_prev:
                st.session_state.prev_close_dict[key] = prev_input if prev_input is not None and prev_input > 0 else None
                st.rerun()
        else:
            st.markdown(f"""
            <div style="background: rgba(255,255,255,0.02); border-radius: 12px; padding: 12px 14px; border: 1px dashed #2a3a4a;">
                <div style="color: #7a8ba3; font-size: 13px;">{key}</div>
                <div style="color: #4a5a6e; font-size: 16px; padding: 8px 0;">⏳ 等待上传</div>
            </div>
            """, unsafe_allow_html=True)

# ============================================================

st.markdown("### 📤 第一步：分别上传四个板块的分时数据（Excel/CSV）")
cols_upload = st.columns(4)
index_keys = list(st.session_state.data_dict.keys())

for i, (col, key) in enumerate(zip(cols_upload, index_keys)):
    with col:
        status = "✅" if st.session_state.data_dict[key] is not None else "⬜"
        st.markdown(f"**{status} {key}**")
        
        uploaded_file = st.file_uploader(
            f"上传 {key} 数据",
            type=['xlsx', 'xls', 'csv'],
            label_visibility="collapsed",
            key=f"upload_{key}"
        )
        
        if uploaded_file is not None:
            df = parse_uploaded_file(uploaded_file)
            if df is not None:
                st.session_state.data_dict[key] = df
                st.success(f"✅ {len(df)} 条数据")
            else:
                st.session_state.data_dict[key] = None
                st.error("解析失败")

# ---------- 第二步：图表 + 右侧面板 ----------
col_chart, col_right = st.columns([2.2, 1])

with col_chart:
    st.markdown("### 📊 第二步：点击标签切换板块视图")
    # 修改：动态渲染事件展示内容，实现添加即预览
    if st.session_state.submitted_flag:
        display_events = st.session_state.submitted_events
        st.info("📌 当前显示【已提交】版本，点击右侧「重置」可重新编辑")
    else:
        # 未提交时，将缓存区事件转为DataFrame用于图表实时预览
        if st.session_state.pending_events:
            display_events = pd.DataFrame(st.session_state.pending_events)
        else:
            display_events = pd.DataFrame(columns=["时间", "事件描述"])
        st.info("✏️ 预览模式：在右侧添加事件后可实时在下方图表预览，最终完成请点击「提交」")

    tabs = st.tabs(index_keys)
    for tab, key in zip(tabs, index_keys):
        with tab:
            data = st.session_state.data_dict.get(key)
            fig = draw_chart(data, display_events, key)
            st.plotly_chart(fig, width='stretch', use_container_width=True)
            if data is None:
                st.info(f"👆 请先在顶部上传 {key} 的 Excel 文件")

with col_right:
    # ---------- 热点事件管理 ----------
    st.subheader("⏱️ 热点事件")
    
    if st.session_state.submitted_flag:
        st.info("已提交，事件锁定。如需修改请点击「重置」")
        st.dataframe(st.session_state.submitted_events, use_container_width=True)
    else:
        col_time, col_desc, col_btn = st.columns([1.2, 2.5, 0.8])
        with col_time:
            new_time = st.text_input("时间", placeholder="09:30", key="new_event_time", label_visibility="collapsed")
        with col_desc:
            new_desc = st.text_input("事件描述", placeholder="输入事件内容", key="new_event_desc", label_visibility="collapsed")
        with col_btn:
            st.write("") 
            st.write("") 
            if st.button("➕ 添加", use_container_width=True):
                if new_time.strip() and new_desc.strip():
                    st.session_state.pending_events.append({
                        "时间": new_time.strip(),
                        "事件描述": new_desc.strip()
                    })
                    st.rerun()
                else:
                    st.warning("请完整填写时间和事件描述")
        
        if st.session_state.pending_events:
            st.markdown("**📋 待提交事件列表**")
            for idx, event in enumerate(st.session_state.pending_events):
                col1, col2, col3 = st.columns([1.2, 2.5, 0.8])
                with col1:
                    st.text(event["时间"])
                with col2:
                    st.text(event["事件描述"])
                with col3:
                    if st.button("✕", key=f"del_{idx}"):
                        del st.session_state.pending_events[idx]
                        st.rerun()
        else:
            st.caption("暂无待提交事件，请添加")

        st.caption("💡 添加事件后，可立即在左侧图表中预览。点击下方「提交」最终生效。")

    # ---------- 市场统计 ----------
    st.markdown("---")
    st.subheader("📊 市场统计")

    if st.session_state.submitted_flag:
        stats = st.session_state.submitted_stats
        st.markdown(f"""
        <div style="background: rgba(255,255,255,0.05); border-radius: 12px; padding: 16px;">
            <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                <span style="color: #7a8ba3; font-size: 14px;">成交额</span>
                <span style="color: #e8edf5; font-size: 22px; font-weight: 700;">{stats.get('turnover', '')}</span>
            </div>
            <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                <span style="color: #7a8ba3; font-size: 14px;">较前日增减</span>
                <span style="color: #e8edf5; font-size: 22px; font-weight: 700;">{stats.get('change', '')}</span>
            </div>
            <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                <span style="color: #7a8ba3; font-size: 14px;">上涨家数</span>
                <span style="color: #ff4d4f; font-size: 22px; font-weight: 700;">{stats.get('up_cnt', '')}</span>
            </div>
            <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                <span style="color: #7a8ba3; font-size: 14px;">下跌家数</span>
                <span style="color: #3ecf8e; font-size: 22px; font-weight: 700;">{stats.get('down_cnt', '')}</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
                <span style="color: #7a8ba3; font-size: 14px;">涨幅居前板块</span>
                <span style="color: #f5c542; font-size: 20px; font-weight: 700;">{stats.get('sectors', '')}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            turnover = st.text_input("成交额（亿元）", placeholder="27182.89", key="stat_turnover")
            up_cnt = st.text_input("上涨家数", placeholder="1740", key="stat_up")
        with col_s2:
            change = st.text_input("较前日增减", placeholder="+464.24", key="stat_change")
            down_cnt = st.text_input("下跌家数", placeholder="3710", key="stat_down")
        sectors = st.text_input("涨幅居前板块", placeholder="油气 · 煤炭 · 白酒", key="stat_sectors")

    # ---------- 提交 & 重置按钮 ----------
    st.markdown("---")
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("✅ 提交", use_container_width=True, disabled=st.session_state.submitted_flag):
            stats = {
                "turnover": st.session_state.get("stat_turnover", ""),
                "change": st.session_state.get("stat_change", ""),
                "up_cnt": st.session_state.get("stat_up", ""),
                "down_cnt": st.session_state.get("stat_down", ""),
                "sectors": st.session_state.get("stat_sectors", "")
            }
            if st.session_state.pending_events:
                submitted_events_df = pd.DataFrame(st.session_state.pending_events)
            else:
                submitted_events_df = pd.DataFrame(columns=["时间", "事件描述"])
            
            st.session_state.submitted_events = submitted_events_df
            st.session_state.submitted_stats = stats
            st.session_state.submitted_flag = True
            st.rerun()
    with col_btn2:
        if st.button("🔄 重置", use_container_width=True, disabled=not st.session_state.submitted_flag):
            st.session_state.submitted_flag = False
            # 优化交互：点击重置时不再清空数据，而是把已提交的事件退回到待编辑状态
            if not st.session_state.submitted_events.empty:
                st.session_state.pending_events = st.session_state.submitted_events.to_dict('records')
            st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件描述"])
            st.rerun()

# ---------- 底部时间戳 ----------
st.markdown("---")
col_f1, col_f2 = st.columns([3, 1])
with col_f1:
    st.caption(f"🕒 数据来源：手动上传 ｜ 当前时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
with col_f2:
    st.caption("📌 所有数据仅在本地处理，不上传服务器")
