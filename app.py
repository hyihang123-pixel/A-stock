import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import warnings

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

if "events" not in st.session_state:
    st.session_state.events = pd.DataFrame({
        "时间": ["09:40", "10:30"],
        "事件描述": ["示例事件A", "示例事件B"]
    })

# ---------- 新增：提交状态 ----------
if "submitted_flag" not in st.session_state:
    st.session_state.submitted_flag = False          # 是否已提交
if "submitted_events" not in st.session_state:
    st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件描述"])
if "submitted_stats" not in st.session_state:
    st.session_state.submitted_stats = {
        "turnover": "",
        "change": "",
        "up_cnt": "",
        "down_cnt": "",
        "sectors": ""
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

# ---------- 3. 计算收盘价和涨跌幅 ----------
def get_index_summary(data_df, prev_close):
    if data_df is None or data_df.empty:
        return None, None, None, None, None
    
    close_price = data_df['价格'].iloc[-1]
    open_price = data_df['价格'].iloc[0]
    
    if prev_close is not None and prev_close > 0:
        change_pct = ((close_price - prev_close) / prev_close) * 100
        used_prev = prev_close
        calc_method = "昨日收盘"
    else:
        change_pct = ((close_price - open_price) / open_price) * 100 if open_price != 0 else None
        used_prev = open_price
        calc_method = "开盘价（估算）"
    
    return close_price, change_pct, open_price, used_prev, calc_method

# ---------- 4. 辅助函数：将时间字符串转为分钟数 ----------
def time_to_min(t):
    try:
        h, m = map(int, t.split(':'))
        return h*60 + m
    except:
        return 0

# ---------- 5. 核心绘图函数（含分段着色 + 跳过休市） ----------
def draw_chart(data_df, events_df, index_name, prev_close):
    """
    绘制分时图，支持：
    - x轴跳过11:30-13:00
    - 分时线根据昨日收盘价分段着色（红 > 昨收，绿 < 昨收）
    """
    if data_df is None or data_df.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"请上传 {index_name} 数据", x=0.5, y=0.5, showarrow=False, font=dict(size=20, color="gray"))
        fig.update_layout(height=500)
        return fig

    # 复制数据并添加分钟列
    df = data_df.copy()
    df['分钟'] = df['时间'].apply(time_to_min)
    # 只保留交易时间内的数据（9:30-11:30 和 13:00-15:00）
    df = df[(df['分钟'] >= 570) & (df['分钟'] <= 900)]  # 9:30=570, 15:00=900
    # 剔除中午休市期间可能存在的异常数据
    df = df[~((df['分钟'] > 690) & (df['分钟'] < 780))]

    if df.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"{index_name} 无有效数据", x=0.5, y=0.5, showarrow=False, font=dict(size=20, color="gray"))
        fig.update_layout(height=500)
        return fig

    # 创建子图
    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        row_heights=[0.65, 0.35],
        subplot_titles=(f"{index_name} 指数走势", "成交量")
    )

    # ----- 1. 指数折线（分段着色） -----
    # 如果prev_close有效，则分段；否则全红
    if prev_close is not None and prev_close > 0:
        # 分段：遍历数据，当价格相对于prev_close的方向改变时切分
        segments = []
        current_seg = [df.iloc[0]]
        current_dir = 1 if df.iloc[0]['价格'] >= prev_close else -1  # 1: 高于/等于, -1: 低于

        for i in range(1, len(df)):
            row = df.iloc[i]
            dir_now = 1 if row['价格'] >= prev_close else -1
            if dir_now == current_dir:
                current_seg.append(row)
            else:
                segments.append((current_dir, pd.DataFrame(current_seg)))
                current_seg = [row]
                current_dir = dir_now
        segments.append((current_dir, pd.DataFrame(current_seg)))

        # 为每个段绘制折线
        for direction, seg_df in segments:
            color = '#ff4d4f' if direction == 1 else '#3ecf8e'  # 红涨绿跌（A股习惯）
            fig.add_trace(
                go.Scatter(
                    x=seg_df['分钟'],
                    y=seg_df['价格'],
                    mode='lines',
                    name='指数点位' if direction == 1 else None,  # 只显示一个图例
                    line=dict(color=color, width=2.5),
                    showlegend=(direction == 1)  # 只在第一个段显示图例
                ),
                row=1, col=1
            )
    else:
        # 无昨收，统一红色
        fig.add_trace(
            go.Scatter(
                x=df['分钟'],
                y=df['价格'],
                mode='lines',
                name='指数点位',
                line=dict(color='#ff4d4f', width=2.5)
            ),
            row=1, col=1
        )

    # ----- 2. 成交量柱状图 -----
    fig.add_trace(
        go.Bar(
            x=df['分钟'],
            y=df['成交量'],
            name='成交量',
            marker_color='#3b82f6',
            opacity=0.7,
            width=1.5  # 调整柱宽，避免重叠
        ),
        row=2, col=1
    )

    # ----- 3. 事件标记（使用分钟数值定位） -----
    if events_df is not None and not events_df.empty:
        for _, event in events_df.iterrows():
            event_time = str(event['时间']).strip()
            event_desc = str(event['事件描述']).strip()
            if event_time and event_desc:
                minute_val = time_to_min(event_time)
                # 检查该时间是否在数据范围内
                if minute_val in df['分钟'].values:
                    fig.add_vline(
                        x=minute_val,
                        line_dash="dash",
                        line_color="orange",
                        line_width=1.5,
                        annotation_text=event_desc,
                        annotation_position="top left",
                        annotation_font_size=11,
                        annotation_font_color="orange",
                        annotation_bgcolor="rgba(0,0,0,0.6)",
                        row=1, col=1
                    )

    # ----- 4. 布局与轴设置 -----
    fig.update_layout(
        height=520,
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template='plotly_dark',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=40, b=20)
    )

    # x轴：使用分钟数值，设置刻度标签为时间字符串，并跳过休市时段
    tick_vals = list(range(570, 901, 30))  # 9:30 到 15:00 每30分钟一个刻度
    tick_text = [f"{h:02d}:{m:02d}" for h, m in [(v//60, v%60) for v in tick_vals]]
    # 但我们的数据可能不是整30分钟，但刻度仍可显示
    fig.update_xaxes(
        title_text="时间",
        row=2, col=1,
        tickvals=tick_vals,
        ticktext=tick_text,
        tickangle=45,
        tickfont=dict(size=10),
        range=[570, 900],
        rangebreaks=[dict(bounds=[690, 780])]  # 跳过 11:30-13:00
    )

    # y轴
    if not df.empty:
        min_price = df['价格'].min()
        max_price = df['价格'].max()
        padding = (max_price - min_price) * 0.1 if max_price > min_price else 10
        fig.update_yaxes(title_text="点位", row=1, col=1, range=[min_price - padding, max_price + padding])
    
    fig.update_yaxes(title_text="成交量", row=2, col=1)

    return fig

# ---------- 6. 页面UI布局 ----------
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
                help="输入昨日收盘价，用于精确计算涨跌幅和分时线颜色"
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

# ---------- 第一步：四个上传区域 ----------
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
    st.markdown("### 📊 第二步：查看分时图")
    # 根据提交状态决定显示的事件
    if st.session_state.submitted_flag:
        display_events = st.session_state.submitted_events
        st.info("📌 当前显示的是【已提交】的版本，点击右侧「重置」可重新编辑")
    else:
        display_events = st.session_state.events
        st.info("✏️ 当前为编辑模式，完成事件和统计后点击「提交」生成最终图")

    # 使用tabs显示四个板块
    tabs = st.tabs(index_keys)
    for tab, key in zip(tabs, index_keys):
        with tab:
            data = st.session_state.data_dict.get(key)
            # 获取该板块的昨日收盘价（用于着色）
            prev_close = st.session_state.prev_close_dict.get(key, None)
            fig = draw_chart(data, display_events, key, prev_close)
            st.plotly_chart(fig, width='stretch', use_container_width=True)
            if data is None:
                st.info(f"👆 请先在顶部上传 {key} 的 Excel 文件")

with col_right:
    # ---------- 热点事件管理 ----------
    st.subheader("⏱️ 热点事件 (全局)")
    # 如果已提交，禁用编辑；否则允许编辑
    if st.session_state.submitted_flag:
        st.info("已提交，事件锁定。如需修改请点击「重置」")
        # 显示提交的事件（只读）
        st.dataframe(st.session_state.submitted_events, use_container_width=True)
    else:
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
        st.caption("💡 点击表格下方 '添加行' 增加事件，删除行自动移除")

    # ---------- 市场统计 ----------
    st.markdown("---")
    st.subheader("📊 市场统计")

    if st.session_state.submitted_flag:
        # 提交后显示大号加粗的统计数字
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
        # 编辑模式：输入框
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
            # 收集当前统计输入值
            stats = {
                "turnover": st.session_state.get("stat_turnover", ""),
                "change": st.session_state.get("stat_change", ""),
                "up_cnt": st.session_state.get("stat_up", ""),
                "down_cnt": st.session_state.get("stat_down", ""),
                "sectors": st.session_state.get("stat_sectors", "")
            }
            # 保存提交的事件和统计
            st.session_state.submitted_events = st.session_state.events.copy()
            st.session_state.submitted_stats = stats
            st.session_state.submitted_flag = True
            st.rerun()
    with col_btn2:
        if st.button("🔄 重置", use_container_width=True, disabled=not st.session_state.submitted_flag):
            # 重置提交状态，保留原始数据
            st.session_state.submitted_flag = False
            st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件描述"])
            st.session_state.submitted_stats = {
                "turnover": "",
                "change": "",
                "up_cnt": "",
                "down_cnt": "",
                "sectors": ""
            }
            st.rerun()

# ---------- 底部时间戳 ----------
st.markdown("---")
col_f1, col_f2 = st.columns([3, 1])
with col_f1:
    st.caption(f"🕒 数据来源：手动上传 ｜ 当前时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
with col_f2:
    st.caption("📌 所有数据仅在本地处理，不上传服务器")
