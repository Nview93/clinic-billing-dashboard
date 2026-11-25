# -*- coding: utf-8 -*-
"""
综合内科门诊开单项目效益看板 —— Streamlit 版
数据源：综合内科门诊9月、10月项目统计.xlsx（字段：分类/月份/医生/开单科室/执行科室/类别/项目/现价格/数量/金额）
部署：本文件 + data.csv（或原始 xlsx）一起上传到 GitHub → share.streamlit.io 部署
"""
import io
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============ 配置区 ============
PASSWORD = "ks2026"          # 访问密码，可自行修改
REFRESH_INTERVAL = 600       # 页面自动刷新，单位秒（600=10分钟）
DATA_FILE = "data.csv"       # 默认数据源
COLOR_SEQ = ["#0E6E7B", "#1B9AAA", "#7FD1D9", "#B9E2E8", "#58A6B5", "#0B3B44", "#D98A2B"]
# ================================

st.set_page_config(page_title="综合内科门诊开单效益看板", page_icon="📊", layout="wide")

# ---------- 密码 ----------
if "auth" not in st.session_state:
    st.session_state["auth"] = False
if not st.session_state["auth"]:
    st.title("🔒 访问验证")
    pwd = st.text_input("请输入访问密码", type="password")
    if st.button("登录"):
        if pwd == PASSWORD:
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("密码错误")
    st.stop()

# ---------- 自动刷新 ----------
st.components.v1.html(
    f"<script>setTimeout(()=>{{window.location.reload();}},{REFRESH_INTERVAL*1000});</script>", height=0
)


# ---------- 读取数据 ----------
@st.cache_data(show_spinner=False)
def load_csv(path):
    return pd.read_csv(path)


def load_excel(file):
    """支持上传原始 xlsx：多个月份子表自动合并，自动剔除合计/空行"""
    xls = pd.ExcelFile(file)
    frames = []
    for s in xls.sheet_names:
        d = pd.read_excel(xls, sheet_name=s)
        if "金额" in d.columns and "医生" in d.columns:
            frames.append(d)
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    return df


def clean(df):
    need = {"月份", "医生", "项目", "金额"}
    if df.empty or not need.issubset(df.columns):
        return pd.DataFrame()
    df = df.copy()
    df = df[df["医生"].notna() & df["项目"].notna()]
    df["金额"] = pd.to_numeric(df["金额"], errors="coerce").fillna(0)
    if "数量" in df.columns:
        df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0)
    if "现价格" in df.columns:
        df["现价格"] = pd.to_numeric(df["现价格"], errors="coerce").fillna(0)
    df["月份"] = df["月份"].astype(str)
    df["月标签"] = df["月份"].str.replace("2026-0", "", regex=False).str.replace("2026-", "", regex=False).str.lstrip("0") + "月"
    return df


# 侧边栏
with st.sidebar:
    st.header("📤 数据更新")
    up = st.file_uploader("上传新的 Excel / CSV（支持多月子表合并）", type=["xlsx", "xls", "csv"])
    if up is not None:
        if up.name.lower().endswith(".csv"):
            raw = pd.read_csv(up)
        else:
            raw = load_excel(up)
        st.session_state["df_upload"] = clean(raw)
        st.success(f"已载入 {len(st.session_state['df_upload'])} 条明细")

if "df_upload" in st.session_state:
    df = st.session_state["df_upload"]
    source_tip = "本次上传文件（临时生效，重启回退到 data.csv）"
else:
    try:
        df = clean(load_csv(DATA_FILE))
        source_tip = DATA_FILE
    except Exception:
        df = pd.DataFrame()
        source_tip = "未找到 data.csv"

if df.empty:
    st.error("未读取到有效数据：请确认表头包含 月份 / 医生 / 项目 / 金额")
    st.stop()

# ---------- 标题 ----------
st.markdown(
    "<h2 style='margin-bottom:0'>综合内科门诊开单项目效益看板</h2>"
    f"<p style='color:#7a8b90;margin-top:4px'>数据源：{source_tip} ｜ 明细 {len(df)} 条 ｜ "
    f"更新 {pd.Timestamp.now():%Y-%m-%d %H:%M} ｜ {REFRESH_INTERVAL//60} 分钟自动刷新</p>",
    unsafe_allow_html=True,
)

# ---------- 筛选 ----------
months_all = sorted(df["月份"].unique())
c1, c2, c3, c4 = st.columns([1.1, 1.3, 1.6, 1.3])
sel_months = c1.multiselect("月份", months_all, default=months_all)
sel_doc = c2.multiselect("医生", sorted(df["医生"].unique()), default=[])
sel_proj = c3.multiselect("项目", sorted(df["项目"].unique()), default=[])
sel_dept = c4.multiselect("开单科室", sorted(df["开单科室"].dropna().unique()), default=[]) if "开单科室" in df else []

d = df[df["月份"].isin(sel_months)]
if sel_doc:
    d = d[d["医生"].isin(sel_doc)]
if sel_proj:
    d = d[d["项目"].isin(sel_proj)]
if sel_dept:
    d = d[d["开单科室"].isin(sel_dept)]

# ---------- 指标卡 ----------
tot = d["金额"].sum()
tot_qty = d["数量"].sum() if "数量" in d else 0
mom = None
MONTH_LABEL = dict(zip(df["月份"], df["月标签"]))
mon_sum = d.groupby("月份")["金额"].sum().sort_index()
if len(mon_sum) >= 2:
    mom = (mon_sum.iloc[-1] - mon_sum.iloc[-2]) / mon_sum.iloc[-2] * 100 if mon_sum.iloc[-2] else None

k1, k2, k3, k4 = st.columns(4)
k1.metric("筛选范围总金额", f"{tot:,.1f} 元")
if len(mon_sum) >= 1:
    k2.metric(f"{MONTH_LABEL.get(mon_sum.index[-1], mon_sum.index[-1])}金额", f"{mon_sum.iloc[-1]:,.1f} 元",
              f"{mom:+.1f}% 环比" if mom is not None else None)
k3.metric("开单医生数", f"{d['医生'].nunique()} 人")
k4.metric("开单项目数 / 总数量", f"{d['项目'].nunique()} 项 / {tot_qty:,.0f}")

st.divider()

# ---------- 图1：月度对比 ----------
r1c1, r1c2 = st.columns([1, 1.35])
with r1c1:
    st.subheader("① 两月开单金额对比")
    fig = go.Figure(go.Bar(
        x=[MONTH_LABEL.get(m, m) for m in mon_sum.index], y=mon_sum.values,
        marker_color=[COLOR_SEQ[1], COLOR_SEQ[0], COLOR_SEQ[2]][:len(mon_sum)],
        text=[f"{v:,.0f} 元" for v in mon_sum.values], textposition="outside", textfont_size=13,
    ))
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10), yaxis_title="金额（元）",
                      plot_bgcolor="white", yaxis=dict(gridcolor="#E6F0F2"), showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

with r1c2:
    st.subheader("② 各医生开单金额（按月分组）")
    piv = d.pivot_table(index="医生", columns="月份", values="金额", aggfunc="sum").fillna(0)
    piv = piv.sort_index(axis=1)
    piv.columns = [MONTH_LABEL.get(c, c) for c in piv.columns]
    piv["合计"] = piv.sum(axis=1)
    piv = piv.sort_values("合计", ascending=False)
    mcols = [c for c in piv.columns if c != "合计"]
    fig = px.bar(piv.reset_index().melt(id_vars="医生", value_vars=mcols, var_name="月份", value_name="金额"),
                 x="金额", y="医生", color="月份", orientation="h", color_discrete_sequence=COLOR_SEQ,
                 category_orders={"医生": piv.index.tolist()})
    fig.update_traces(texttemplate="%{x:,.0f}", textposition="outside", cliponaxis=False)
    fig.update_layout(height=380, margin=dict(l=10, r=30, t=20, b=10), plot_bgcolor="white",
                      xaxis=dict(gridcolor="#E6F0F2", title="金额（元）"),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=1, xanchor="right"))
    st.plotly_chart(fig, use_container_width=True)

# ---------- 图2：项目结构 / 医生×项目热力 ----------
r2c1, r2c2 = st.columns([1, 1.15])
with r2c1:
    st.subheader("③ 各项目金额与占比")
    proj = d.groupby("项目")["金额"].sum().sort_values(ascending=True)
    fig = go.Figure(go.Bar(
        x=proj.values, y=proj.index, orientation="h",
        marker_color=[COLOR_SEQ[2]] * len(proj),
        text=[f"{v:,.0f} 元（{v/tot*100:.1f}%）" if tot else "" for v in proj.values],
        textposition="outside", cliponaxis=False,
    ))
    fig.update_traces(marker=dict(color=[COLOR_SEQ[0] if i >= len(proj) - 2 else COLOR_SEQ[2]
                                         for i in range(len(proj))]))
    fig.update_layout(height=380, margin=dict(l=10, r=60, t=20, b=10), plot_bgcolor="white",
                      xaxis=dict(title="金额（元）", gridcolor="#E6F0F2"))
    st.plotly_chart(fig, use_container_width=True)

with r2c2:
    st.subheader("④ 医生 × 项目 金额热力图")
    hm = d.pivot_table(index="医生", columns="项目", values="金额", aggfunc="sum").fillna(0)
    hm = hm.loc[hm.sum(axis=1).sort_values(ascending=False).index]
    fig = px.imshow(hm, text_auto=".2s", aspect="auto", color_continuous_scale="BuGn",
                    labels=dict(x="项目", y="医生", color="金额（元）"))
    fig.update_traces(textfont_size=11, hovertemplate="医生 %{y}<br>项目 %{x}<br>金额 %{z:,.0f} 元<extra></extra>")
    fig.update_xaxes(tickangle=-25)
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=60), coloraxis_colorbar=dict(title="元"))
    st.plotly_chart(fig, use_container_width=True)

# ---------- 图3：开单科室 / 医生占比 ----------
r3c1, r3c2 = st.columns([1.3, 1])
if "开单科室" in d:
    with r3c1:
        st.subheader("⑤ 开单科室 × 月份 金额构成")
        dept = d.pivot_table(index="开单科室", columns="月份", values="金额", aggfunc="sum").fillna(0)
        dept = dept.sort_index(axis=1)
        dept.columns = [MONTH_LABEL.get(c, c) for c in dept.columns]
        dept["合计"] = dept.sum(axis=1)
        dept = dept.sort_values("合计", ascending=True)
        fig = go.Figure()
        for i, m in enumerate([c for c in dept.columns if c != "合计"]):
            fig.add_bar(y=dept.index, x=dept[m], name=m, orientation="h", marker_color=COLOR_SEQ[i % len(COLOR_SEQ)])
        fig.update_layout(barmode="stack", height=360, margin=dict(l=10, r=10, t=20, b=10),
                          plot_bgcolor="white", xaxis=dict(title="金额（元）", gridcolor="#E6F0F2"),
                          legend=dict(orientation="h", y=1.05, x=1, xanchor="right"))
        st.plotly_chart(fig, use_container_width=True)

with r3c2:
    st.subheader("⑥ 医生金额占比")
    doc = d.groupby("医生")["金额"].sum().sort_values(ascending=False)
    fig = px.pie(values=doc.values, names=doc.index, hole=0.5, color_discrete_sequence=COLOR_SEQ)
    fig.update_traces(textinfo="percent+label", textfont_size=12)
    fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# ---------- 明细 & 汇总表 ----------
st.divider()
st.subheader("📋 医生×月份 汇总")
piv2 = d.pivot_table(index="医生", columns="月份", values="金额", aggfunc="sum").fillna(0)
piv2 = piv2.sort_index(axis=1)
piv2.columns = [MONTH_LABEL.get(c, c) for c in piv2.columns]
piv2["合计"] = piv2.sum(axis=1)
piv2["占比%"] = (piv2["合计"] / piv2["合计"].sum() * 100).round(1) if piv2["合计"].sum() else 0
show = piv2.sort_values("合计", ascending=False).copy()
st.dataframe(show, use_container_width=True, column_config={
    c: st.column_config.NumberColumn(c, format="%.1f") for c in show.columns})

st.subheader("🗂 开单明细")
st.dataframe(d, use_container_width=True, height=300)

c1, c2 = st.columns(2)
c1.download_button("📥 下载明细 CSV", d.to_csv(index=False).encode("utf-8-sig"),
                   "开单明细_筛选.csv", "text/csv")
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    piv2.to_excel(w, sheet_name="医生月度汇总")
    d.to_excel(w, sheet_name="筛选明细", index=False)
c2.download_button("📥 下载汇总 Excel", buf.getvalue(), "门诊开单汇总.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
