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

# 热点事件存储 (更新字段结构)
if "pending_events" not in st.session_state:
    st.session_state.pending_events = [] 

if "submitted_flag" not in st.session_state:
    st.session_state.submitted_flag = False
if "submitted_events" not in st.session_state:
    st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件标题", "事件内容"])
if "submitted_stats" not in st.session_state:
    st.session_state.submitted_stats = {
        "turnover": "", "change": "", "up_cnt": "", "down_cnt": "", "sectors": ""
    }

# ---------- 2. 工具函数：智能提取时间格式 ----------
def normalize_time_to_hm(t_str):
    """将任何包含时间的字符串标准化为 HH:MM 格式，兼容中文冒号和单数小时"""
    m = re.search(r'(\d{1,2})[:：](\d{2})', str(t_str))
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return str(t_str).strip()

def is_valid_trading_time(t_str):
    """判断是否为有效交易时间，剔除 11:30 - 13:00 的休盘时间"""
    hm = normalize_time_to_hm(t_str)
    if "11:30" < hm < "13:00":
        return False
    return True

# ---------- 3. 智能解析 Excel ----------
@st.cache_data
def parse_uploaded_file(uploaded_file):
    """读取并解析上传的文件，自动忽略损坏的样式格式，并过滤休盘时间"""
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
        
        # 过滤休盘数据
        result_df = result_df[result_df['时间'].apply(is_valid_trading_time)]

        if result_df.empty:
            st.error("❌ 解析后无有效数据")
            return None

        return result_df

    except Exception as e:
        st.error(f"❌ 文件解析失败: {e}")
        return None


# ---------- 4. 核心绘图函数（红绿双色分时算法 + 智能防遮挡） ----------
def draw_chart(data_df, events_df, index_name, prev_close=None):
    if data_df is None or data_df.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"暂无 {index_name} 数据", x=0.5, y=0.5, showarrow=False, font=dict(size=20, color="gray"))
        fig.update_layout(height=550) 
        return fig

    baseline = prev_close if prev_close and prev_close > 0 else data_df['价格'].iloc[0]
    prices = data_df['价格'].tolist()
    times = data_df['时间'].tolist()
    volumes = data_df['成交量'].tolist()

    # 👉 为智能错位计算Y轴边界 (大幅增加顶部留白 40%)
    min_price = min(prices)
    max_price = max(prices)
    price_range = max_price - min_price if max_price > min_price else (baseline * 0.02 if baseline else 10)
    
    top_padding = price_range * 0.40  # 顶部多留 40% 空间，用于叠放事件气泡
    bottom_padding = price_range * 0.10
    
    y_min_bound = min(min_price, baseline) - bottom_padding
    y_max_bound = max(max_price, baseline) + top_padding

    # 核心算法：拆分红绿折线
    y_red = []
    y_green = []
    
    for p in prices:
        if p >= baseline:
            y_red.append(p)
            y_green.append(None)
        else:
            y_red.append(None)
            y_green.append(p)

    for i in range(1, len(prices)):
        p_prev = prices[i-1]
        p_curr = prices[i]
        if p_prev >= baseline and p_curr < baseline:
            y_green[i-1] = p_prev
        elif p_prev < baseline and p_curr >= baseline:
            y_red[i-1] = p_prev

    vol_colors = ['#ff4d4f' if p >= baseline else '#3ecf8e' for p in prices]

    # 初始化画布
    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        row_heights=[0.65, 0.35],
        subplot_titles=(f"{index_name} 指数走势", "成交量")
    )

    fig.add_trace(go.Scatter(
        x=times, y=prices, mode='lines', line=dict(color='rgba(0,0,0,0)', width=0),
        fill='tozeroy', fillcolor='rgba(150,150,150,0.06)', showlegend=False, hoverinfo='skip'
    ), row=1, col=1)

    fig.add_trace(go.Scatter(
        x=times, y=y_red, mode='lines', name='高于基准 (涨)',
        line=dict(color='#ff4d4f', width=2.5), connectgaps=False
    ), row=1, col=1)

    fig.add_trace(go.Scatter(
        x=times, y=y_green, mode='lines', name='低于基准 (跌)',
        line=dict(color='#3ecf8e', width=2.5), connectgaps=False
    ), row=1, col=1)

    baseline_label = "昨收基准" if (prev_close and prev_close > 0) else "开盘基准"
    fig.add_hline(
        y=baseline, line_dash="dash", line_color="rgba(255,255,255,0.4)", line_width=1.5,
        annotation_text=f"{baseline_label}: {baseline:.2f}", annotation_position="bottom right",
        annotation_font_color="rgba(255,255,255,0.7)", row=1, col=1
    )

    fig.add_trace(go.Bar(
        x=times, y=volumes, name='成交量', marker_color=vol_colors, opacity=0.8, showlegend=False
    ), row=2, col=1)

    # 👉 升级版：带防遮挡算法的事件绘制
    if events_df is not None and not events_df.empty:
        # 记录每个时间的索引位置，用于计算横向距离
        time_map = {normalize_time_to_hm(t): (idx, t) for idx, t in enumerate(times)}
        
        event_xs = []
        event_ys = []
        event_texts = []
        event_hovers = []
        
        placed_events = [] # 存储已放置的事件 (x轴索引, 所在层级)
        safe_distance = 25 # 防遮挡安全距离（间隔小于25分钟视为会遮挡）
        
        for _, event in events_df.iterrows():
            event_time = str(event.get('时间', '')).strip()
            # 兼容旧版本字段
            event_title = str(event.get('事件标题', event.get('事件描述', ''))).strip()
            event_content = str(event.get('事件内容', '')).strip()
            
            norm_ev_time = normalize_time_to_hm(event_time)
            
            if norm_ev_time and event_title and norm_ev_time in time_map:
                x_idx, actual_x_in_df = time_map[norm_ev_time]
                
                # 画出竖直虚线
                fig.add_vline(
                    x=actual_x_in_df, line_dash="dash", line_color="orange", line_width=1.5,
                    row=1, col=1
                )
                
                # ---------------- 防遮挡碰撞检测算法 ----------------
                level = 0
                while True:
                    conflict = False
                    for prev_idx, prev_level in placed_events:
                        # 如果在同一层，且横向距离小于安全距离，则发生碰撞
                        if prev_level == level and abs(x_idx - prev_idx) < safe_distance:
                            conflict = True
                            break
                    if not conflict:
                        break # 找到没有冲突的层级
                    level += 1 # 碰撞了，往下一层挪
                
                placed_events.append((x_idx, level))
                
                # 根据所在层级计算 Y 坐标 (每降一层，Y轴往下挪)
                # 留出 5% 的绝对顶边距，每一层占据总高度的 ~6%
                y_val = y_max_bound - (top_padding * 0.05) - (level * top_padding * 0.18)
                
                event_xs.append(actual_x_in_df)
                event_ys.append(y_val)
                event_texts.append(f"{event_time} {event_title}")
                
                # 组合悬浮显示的富文本
                hover_html = f"<b>时间：</b>{event_time}<br><b>标题：</b>{event_title}"
                if event_content:
                    # 每30个字符自动换行一下，防止内容过长导致悬浮框溢出
                    wrapped_content = "<br>".join([event_content[i:i+30] for i in range(0, len(event_content), 30)])
                    hover_html += f"<br><b>内容：</b>{wrapped_content}"
                
                event_hovers.append(hover_html)
        
        # 批量绘制所有计算好坐标的气泡
        if event_xs:
            fig.add_trace(
                go.Scatter(
                    x=event_xs, y=event_ys,
                    mode='markers+text',
                    marker=dict(symbol='triangle-down', size=10, color='orange'),
                    text=event_texts,
                    textposition='middle right', # 放到居中偏右，方便多层堆叠时对齐
                    textfont=dict(color='orange', size=12),
                    hovertext=event_hovers,       # 悬浮显示标题+内容
                    hoverinfo='text',             
                    showlegend=False
                ),
                row=1, col=1
            )

    fig.update_layout(
        height=550, hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template='plotly_dark', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=40, b=20)
    )

    fig.update_yaxes(title_text="点位", row=1, col=1, range=[y_min_bound, y_max_bound])
    fig.update_yaxes(title_text="成交量", row=2, col=1)
    
    fig.update_xaxes(type='category', row=1, col=1)
    fig.update_xaxes(
        title_text="时间", row=2, col=1, type='category',
        tickangle=45, tickfont=dict(size=10), nticks=15
    )

    return fig


# ---------- 5. 模块化 UI 组件 ----------
def render_index_overview(edit_mode=True):
    """渲染指数概览"""
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


# ============================================================
# 主界面渲染逻辑分流
# ============================================================

if st.session_state.submitted_flag:
    # ----------------------------------------------------
    # 🎯 报告模式：沉浸式阅读界面（提交后）
    # ----------------------------------------------------
    col_t1, col_t2 = st.columns([5, 1])
    with col_t1:
        st.markdown(f"<h1 style='color:#e8edf5;'>📄 市场复盘分析报告 <span style='font-size:16px; color:#7a8ba3;'>{datetime.now().strftime('%Y-%m-%d')}</span></h1>", unsafe_allow_html=True)
    with col_t2:
        st.write("") 
        if st.button("⬅️ 返回修改", use_container_width=True):
            st.session_state.submitted_flag = False
            if not st.session_state.submitted_events.empty:
                st.session_state.pending_events = st.session_state.submitted_events.to_dict('records')
            st.session_state.submitted_events = pd.DataFrame(columns=["时间", "事件标题", "事件内容"])
            st.rerun()
    
    st.markdown("---")

    render_index_overview(edit_mode=False)

    st.markdown("<br>### 📈 核心市场统计", unsafe_allow_html=True)
    stats = st.session_state.submitted_stats
    st.markdown(f"""
    <div style="display: flex; justify-content: space-around; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 25px 10px; border: 1px solid #1e2a3a;">
        <div style="text-align: center;">
            <div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">成交额</div>
            <div style="color: #e8edf5; font-size: 26px; font-weight: 700;">{stats.get('turnover', '-')}</div>
        </div>
        <div style="text-align: center;">
            <div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">较前日增减</div>
            <div style="color: #e8edf5; font-size: 26px; font-weight: 700;">{stats.get('change', '-')}</div>
        </div>
        <div style="text-align: center;">
            <div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">上涨家数</div>
            <div style="color: #ff4d4f; font-size: 26px; font-weight: 700;">{stats.get('up_cnt', '-')}</div>
        </div>
        <div style="text-align: center;">
            <div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">下跌家数</div>
            <div style="color: #3ecf8e; font-size: 26px; font-weight: 700;">{stats.get('down_cnt', '-')}</div>
        </div>
        <div style="text-align: center;">
            <div style="color: #7a8ba3; font-size: 15px; margin-bottom: 8px;">涨幅居前板块</div>
            <div style="color: #f5c542; font-size: 22px; font-weight: 700;">{stats.get('sectors', '-')}</div>
        </div>
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
    # ----------------------------------------------------
    # 🛠️ 编辑模式：工作台界面
    # ----------------------------------------------------
    st.markdown("<h1 style='color:#e8edf5; border-left: 4px solid #ff4d4f; padding-left: 16px;'>📈 多板块分时图分析器 <span style='font-size:16px; color:#7a8ba3;'>工作台模式</span></h1>", unsafe_allow_html=True)
    
    render_index_overview(edit_mode=True)

    st.markdown("### 📤 第一步：分别上传四个板块的分时数据（Excel/CSV）")
    cols_upload = st.columns(4)
    index_keys = list(st.session_state.data_dict.keys())

    for i, (col, key) in enumerate(zip(cols_upload, index_keys)):
        with col:
            status = "✅" if st.session_state.data_dict[key] is not None else "⬜"
            st.markdown(f"**{status} {key}**")
            
            uploaded_file = st.file_uploader(
                f"上传 {key} 数据", type=['xlsx', 'xls', 'csv'],
                label_visibility="collapsed", key=f"upload_{key}"
            )
            
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
        
        if st.session_state.pending_events:
            display_events = pd.DataFrame(st.session_state.pending_events)
        else:
            display_events = pd.DataFrame(columns=["时间", "事件标题", "事件内容"])
        st.info("✏️ 预览模式：在上方输入昨收可实时触发红绿变色，添加事件后可实时预览（事件密集时会自动阶梯排列防遮挡）。")

        tabs = st.tabs(index_keys)
        for tab, key in zip(tabs, index_keys):
            with tab:
                data = st.session_state.data_dict.get(key)
                current_prev = st.session_state.prev_close_dict.get(key, None)
                fig = draw_chart(data, display_events, key, current_prev)
                st.plotly_chart(fig, width='stretch', use_container_width=True)
                if data is None:
                    st.info(f"👆 请先在顶部上传 {key} 的文件")

    with col_right:
        st.subheader("⏱️ 热点事件编辑")
        
        # 将输入框布局拆分为 时间、标题、内容
        col_time, col_title, col_content, col_btn = st.columns([1, 1.5, 2.5, 0.8])
        with col_time:
            new_time = st.text_input("时间", placeholder="09:30", key="new_event_time", label_visibility="collapsed")
        with col_title:
            new_title = st.text_input("事件标题", placeholder="短标题(图上直显)", key="new_event_title", label_visibility="collapsed")
        with col_content:
            new_content = st.text_input("事件内容", placeholder="详细内容(悬浮查看)", key="new_event_content", label_visibility="collapsed")
        with col_btn:
            if st.button("➕", use_container_width=True, help="添加事件"):
                if new_time.strip() and new_title.strip():
                    st.session_state.pending_events.append({
                        "时间": new_time.strip(),
                        "事件标题": new_title.strip(),
                        "事件内容": new_content.strip()
                    })
                    st.rerun()
                else:
                    st.warning("时间和标题必填")
        
        if st.session_state.pending_events:
            st.markdown("**📋 待提交事件列表**")
            for idx, event in enumerate(st.session_state.pending_events):
                c1, c2, c3, c4 = st.columns([1, 1.5, 2.5, 0.8])
                with c1: st.caption(event.get("时间", ""))
                with c2: st.caption(event.get("事件标题", event.get("事件描述","")))
                with c3: 
                    content_str = event.get("事件内容", "")
                    st.caption(content_str[:10] + "..." if len(content_str) > 10 else content_str)
                with c4:
                    if st.button("✕", key=f"del_{idx}"):
                        del st.session_state.pending_events[idx]
                        st.rerun()
        else:
            st.caption("暂无事件，添加后即可在左侧预览")

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
                submitted_events_df = pd.DataFrame(columns=["时间", "事件标题", "事件内容"])
            
            st.session_state.submitted_events = submitted_events_df
            st.session_state.submitted_stats = stats
            st.session_state.submitted_flag = True
            st.rerun()

# ---------- 底部时间戳 ----------
st.markdown("---")
col_f1, col_f2 = st.columns([3, 1])
with col_f1:
    st.caption(f"🕒 数据仅在本地处理 ｜ 当前时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
