# -*- coding: utf-8 -*-
"""
综合内科门诊开单项目效益看板 v2
数据源：综合内科门诊9月、10月项目统计.xlsx
字段：分类/月份/医生/开单科室/执行科室/类别/项目/现价格/数量/金额

本次修订（对照原始表格）：
1. 项目统一为「简称」展示，全称为准，并新增简称↔全称对照表；
2. 开单科室统一清理括号写法（儿内科(书林专家)→儿内科·书林专家 等）；
3. 月份标签统一为「2026年9月」；
4. 金额一律用「元」+ 千分位，彻底去掉 Plotly 的 k/M 缩略；
5. 六个图改为单列排布并加宽边距，避免中文标签被截断、重叠。
"""
import io
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============ 配置区 ============
PASSWORD = "ks2026"
REFRESH_INTERVAL = 600
DATA_FILE = "data.csv"
COLOR_SEQ = ["#0E6E7B", "#1B9AAA", "#7FD1D9", "#B9E2E8", "#58A6B5", "#0B3B44", "#D98A2B"]

PROJECT_SHORT = {
    "肺通气功能检查费": "肺通气功能",
    "肺通气功能检查费-儿童（加收）": "肺通气（儿童加收）",
    "一氧化氮呼气测定": "一氧化氮呼气测定",
    "变应原皮肤试验费": "变应原皮试",
    "呼吸阻力检查费": "呼吸阻力",
    "支气管激发试验检查费": "支气管激发试验",
    "结核菌素试验": "结核菌素试验",
}
DEPT_DISPLAY = {
    "儿内科(书林专家)": "儿内科·书林专家",
    "儿内科(书林)A": "儿内科·书林A区",
    "儿内科(书林)C": "儿内科·书林C区",
    "儿内科复诊(书林)": "儿内科复诊·书林",
}
# ================================

st.set_page_config(page_title="综合内科门诊开单效益看板", page_icon="📊", layout="wide")

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

st.components.v1.html(
    f"<script>setTimeout(()=>{{window.location.reload();}},{REFRESH_INTERVAL*1000});</script>", height=0
)


@st.cache_data(show_spinner=False)
def load_csv(path):
    return pd.read_csv(path)


def load_excel(file):
    xls = pd.ExcelFile(file)
    frames = []
    for s in xls.sheet_names:
        d = pd.read_excel(xls, sheet_name=s)
        if "金额" in d.columns and "医生" in d.columns:
            frames.append(d)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def clean(df):
    need = {"月份", "医生", "项目", "金额"}
    if df.empty or not need.issubset(df.columns):
        return pd.DataFrame()
    df = df.copy()
    df = df[df["医生"].notna() & df["项目"].notna()].copy()
    for c in ["医生", "项目", "开单科室", "执行科室", "类别"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.replace("\u3000", "").str.strip()
    for c in ["金额", "数量", "现价格"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["月份"] = df["月份"].astype(str).str.strip()

    def mon_label(m):
        try:
            y, mm = m.split("-")[:2]
            return f"{y}年{int(mm)}月"
        except Exception:
            return m

    df["月标签"] = df["月份"].map(mon_label)
    df["项目简称"] = df["项目"].map(lambda x: PROJECT_SHORT.get(x, x))
    if "开单科室" in df.columns:
        df["科室显示"] = df["开单科室"].map(lambda x: DEPT_DISPLAY.get(x, x))
    return df


with st.sidebar:
    st.header("📤 数据更新")
    up = st.file_uploader("上传新的 Excel / CSV（支持多月子表合并）", type=["xlsx", "xls", "csv"])
    if up is not None:
        raw = pd.read_csv(up) if up.name.lower().endswith(".csv") else load_excel(up)
        st.session_state["df_upload"] = clean(raw)
        st.success(f"已载入 {len(st.session_state['df_upload'])} 条明细")
    st.divider()
    st.caption("手机上如显示不全，请用电脑浏览器打开或点右上角全屏。")

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

st.markdown(
    "<h2 style='margin-bottom:0'>综合内科门诊开单项目效益看板</h2>"
    f"<p style='color:#7a8b90;margin-top:4px'>数据源：{source_tip} ｜ 明细 {len(df)} 条 ｜ "
    f"更新 {pd.Timestamp.now():%Y-%m-%d %H:%M} ｜ {REFRESH_INTERVAL//60} 分钟自动刷新</p>",
    unsafe_allow_html=True,
)

MONTH_LABEL = dict(zip(df["月份"], df["月标签"]))
MONTH_SHORT = {m: f"{int(m.split('-')[1])}月" if "-" in m else m for m in df["月份"].unique()}
months_all = sorted(df["月份"].unique())

c1, c2, c3, c4 = st.columns([1.1, 1.3, 1.6, 1.3])
sel_months = c1.multiselect("月份", months_all, default=months_all,
                            format_func=lambda x: MONTH_SHORT.get(x, x))
sel_doc = c2.multiselect("医生", sorted(df["医生"].unique()), default=[])
sel_proj = c3.multiselect("项目", sorted(df["项目"].unique()), default=[],
                          format_func=lambda x: PROJECT_SHORT.get(x, x))
sel_dept = (c4.multiselect("开单科室", sorted(df["开单科室"].dropna().unique()), default=[],
                           format_func=lambda x: DEPT_DISPLAY.get(x, x))
            if "开单科室" in df else [])

d = df[df["月份"].isin(sel_months)]
if sel_doc:
    d = d[d["医生"].isin(sel_doc)]
if sel_proj:
    d = d[d["项目"].isin(sel_proj)]
if sel_dept:
    d = d[d["开单科室"].isin(sel_dept)]

# ---------- 指标卡（万元口径）----------
tot = d["金额"].sum()
tot_qty = d["数量"].sum() if "数量" in d else 0
mon_sum = d.groupby("月份")["金额"].sum().sort_index()
mom = None
if len(mon_sum) >= 2 and mon_sum.iloc[-2]:
    mom = (mon_sum.iloc[-1] - mon_sum.iloc[-2]) / mon_sum.iloc[-2] * 100

k1, k2, k3, k4 = st.columns(4)
k1.metric("总金额（万元）", f"{tot/10000:,.2f}", help=f"合计 {tot:,.1f} 元")
if len(mon_sum) >= 1:
    k2.metric(f"{MONTH_LABEL.get(mon_sum.index[-1])}（万元）", f"{mon_sum.iloc[-1]/10000:,.2f}",
              f"{mom:+.1f}% 环比" if mom is not None else None,
              help=f"{mon_sum.iloc[-1]:,.1f} 元")
k3.metric("开单医生", f"{d['医生'].nunique()} 人")
k4.metric("开单项目", f"{d['项目'].nunique()} 项", help=f"项目次数合计 {tot_qty:,.0f} 次")

st.divider()
BASE = dict(height=400, margin=dict(l=110, r=130, t=40, b=50), plot_bgcolor="white", font=dict(size=13))

# ---------- ① 月度对比 ----------
st.subheader("① 两月开单金额对比")
fig = go.Figure(go.Bar(
    x=[MONTH_LABEL.get(m, m) for m in mon_sum.index], y=mon_sum.values,
    marker_color=[COLOR_SEQ[0], COLOR_SEQ[1], COLOR_SEQ[2]][:len(mon_sum)],
    text=[f"{v:,.0f} 元" for v in mon_sum.values], textposition="outside",
    textfont_size=15, cliponaxis=False, width=0.42,
))
fig.update_layout(**BASE, yaxis=dict(title="金额（元）", tickformat=",d", gridcolor="#E6F0F2",
                                     range=[0, mon_sum.max() * 1.2]), showlegend=False)
fig.update_traces(hovertemplate="%{x}<br>%{y:,.0f} 元<extra></extra>")
st.plotly_chart(fig, width="stretch")
st.caption("口径：开单收入（现价格×数量）合计，不等于科室结余或绩效。")

# ---------- ② 各医生 ----------
st.subheader("② 各医生开单金额（按月堆叠）")
piv = d.pivot_table(index="医生", columns="月份", values="金额", aggfunc="sum").fillna(0)
piv = piv.sort_index(axis=1)
piv.columns = [MONTH_LABEL.get(c, c) for c in piv.columns]
piv["合计"] = piv.sum(axis=1)
piv = piv.sort_values("合计", ascending=False)
mcols = [c for c in piv.columns if c != "合计"]
fig = go.Figure()
for i, m in enumerate(mcols):
    fig.add_bar(y=piv.index, x=piv[m], name=m, orientation="h",
                marker_color=COLOR_SEQ[i % len(COLOR_SEQ)],
                hovertemplate="%{y}｜" + m + "<br>%{x:,.0f} 元<extra></extra>")
for y, v in zip(piv.index, piv["合计"]):
    fig.add_annotation(x=v, y=y, text=f"{v:,.0f} 元", showarrow=False,
                       xanchor="left", xshift=4, font=dict(size=12, color="#0B3B44"))
fig.update_layout(**BASE, barmode="stack",
                  xaxis=dict(title="金额（元）", tickformat=",d", gridcolor="#E6F0F2",
                             range=[0, max(piv["合计"].max() * 1.3, 1)]),
                  yaxis=dict(title="医生", autorange="reversed"),
                  legend=dict(orientation="h", yanchor="bottom", y=1.0, x=1, xanchor="right"))
st.plotly_chart(fig, width="stretch")
st.caption("条形为两月金额相叠，右侧标注该医生两月合计。")

# ---------- ③ 项目结构 ----------
st.subheader("③ 各项目金额与占比")
proj = d.groupby("项目简称")["金额"].sum().sort_values(ascending=True)
pfull = d.groupby("项目简称")["项目"].agg(lambda s: s.iloc[0])
fig = go.Figure(go.Bar(
    x=proj.values, y=proj.index, orientation="h",
    marker_color=[COLOR_SEQ[0] if i >= len(proj) - 2 else COLOR_SEQ[2] for i in range(len(proj))],
    text=[f"{v:,.0f} 元 · {v/tot*100:.1f}%" if tot else "" for v in proj.values],
    textposition="outside", cliponaxis=False, textfont_size=12,
    customdata=[[pfull.get(k, k)] for k in proj.index],
    hovertemplate="%{customdata[0]}<br>%{x:,.0f} 元<extra></extra>",
))
fig.update_layout(**BASE, xaxis=dict(title="金额（元）", tickformat=",d", gridcolor="#E6F0F2",
                                     range=[0, max(proj.max() * 1.5, 1)]),
                  yaxis=dict(title="项目（简称）"))
st.plotly_chart(fig, width="stretch")
st.caption("文字为「金额 元 · 占比」；悬停显示项目全称。下方有简称↔全称对照表。")

# ---------- ④ 热力图 ----------
st.subheader("④ 医生 × 项目 金额热力图")
hm = d.pivot_table(index="医生", columns="项目简称", values="金额", aggfunc="sum").fillna(0)
hm = hm.loc[hm.sum(axis=1).sort_values(ascending=False).index]
hm = hm[hm.sum(axis=0).sort_values(ascending=False).index]
textm = [[f"{v:,.0f}" if v else "—" for v in row] for row in hm.values]
fig = go.Figure(go.Heatmap(
    z=hm.values, x=list(hm.columns), y=list(hm.index), text=textm, texttemplate="%{text}",
    textfont=dict(size=11), colorscale="BuGn", colorbar=dict(title="金额（元）", tickformat=",d"),
    hovertemplate="医生 %{y}<br>项目 %{x}<br>%{z:,.0f} 元<extra></extra>", xgap=2, ygap=2,
))
fig.update_layout(height=430, margin=dict(l=110, r=40, t=30, b=90), plot_bgcolor="white",
                  font=dict(size=13), xaxis=dict(tickangle=-28, title="项目（简称）", tickfont=dict(size=11)),
                  yaxis=dict(title="医生", autorange="reversed"))
st.plotly_chart(fig, width="stretch")
st.caption("单元格为金额（元），「—」表示该医生未开此项目。")

# ---------- ⑤ 科室构成 ----------
if "开单科室" in d:
    st.subheader("⑤ 开单科室 × 月份 金额构成")
    dept = d.pivot_table(index="科室显示", columns="月份", values="金额", aggfunc="sum").fillna(0)
    dept = dept.sort_index(axis=1)
    dept.columns = [MONTH_LABEL.get(c, c) for c in dept.columns]
    dept["合计"] = dept.sum(axis=1)
    dept = dept.sort_values("合计", ascending=True)
    mcols2 = [c for c in dept.columns if c != "合计"]
    fig = go.Figure()
    for i, m in enumerate(mcols2):
        fig.add_bar(y=dept.index, x=dept[m], name=m, orientation="h",
                    marker_color=COLOR_SEQ[i % len(COLOR_SEQ)],
                    hovertemplate="%{y}｜" + m + "<br>%{x:,.0f} 元<extra></extra>")
    for y, v in zip(dept.index, dept["合计"]):
        fig.add_annotation(x=v, y=y, text=f"{v:,.0f} 元", showarrow=False,
                           xanchor="left", xshift=4, font=dict(size=11, color="#0B3B44"))
    fig.update_layout(**BASE, barmode="stack",
                      xaxis=dict(title="金额（元）", tickformat=",d", gridcolor="#E6F0F2",
                                 range=[0, dept["合计"].max() * 1.35]),
                      yaxis=dict(title="开单科室"),
                      legend=dict(orientation="h", yanchor="bottom", y=1.0, x=1, xanchor="right"))
    st.plotly_chart(fig, width="stretch")
    st.caption("科室名已按原表规范整理：儿内科(书林专家)→儿内科·书林专家，儿内科(书林)A/C→A区/C区。")

# ---------- ⑥ 医生占比 ----------
st.subheader("⑥ 医生金额占比")
doc = d.groupby("医生")["金额"].sum().sort_values(ascending=False)
dtot = doc.sum()
keep = doc[doc / dtot >= 0.01]
rest = doc[doc / dtot < 0.01]
if len(rest) > 0:
    keep = pd.concat([keep, pd.Series({f"其他（{len(rest)}人）": rest.sum()})])
fig = px.pie(values=keep.values, names=keep.index, hole=0.45,
             color_discrete_sequence=COLOR_SEQ[:len(keep)])
fig.update_traces(textinfo="label+percent", texttemplate="%{label}<br>%{percent:.1%}",
                  textposition="outside", textfont_size=12,
                  hovertemplate="%{label}<br>%{value:,.0f} 元<extra></extra>")
fig.update_layout(height=470, margin=dict(l=10, r=10, t=40, b=30), showlegend=False,
                  font=dict(size=13))
st.plotly_chart(fig, width="stretch")
st.caption("悬停可看具体金额；占比按筛选范围总额计算。")

# ---------- 汇总表 ----------
st.divider()
st.subheader("📋 医生×月份 汇总（元）")
piv3 = d.pivot_table(index="医生", columns="月份", values="金额", aggfunc="sum").fillna(0)
piv3 = piv3.sort_index(axis=1)
piv3.columns = [MONTH_LABEL.get(c, c) for c in piv3.columns]
piv3["合计"] = piv3.sum(axis=1)
piv3["占比%"] = (piv3["合计"] / piv3["合计"].sum() * 100).round(1) if piv3["合计"].sum() else 0
show = piv3.sort_values("合计", ascending=False).copy()
st.dataframe(show, width="stretch", column_config={
    c: st.column_config.NumberColumn(c, format="%.1f") for c in show.columns})

st.subheader("🔤 项目名称对照（简称 ↔ 全称）")
rows = []
for full in sorted(set(df["项目"])):
    short = PROJECT_SHORT.get(full, full)
    rows.append({"项目简称": short, "项目全称": full,
                 "开单金额（元）": round(d.loc[d["项目简称"] == short, "金额"].sum(), 1)})
st.dataframe(pd.DataFrame(rows).sort_values("开单金额（元）", ascending=False),
             width="stretch", hide_index=True)

st.subheader("🗂 开单明细")
cols = [c for c in ["月份", "月标签", "医生", "科室显示", "类别", "项目", "现价格", "数量", "金额"] if c in d.columns]
det = d[cols].rename(columns={"月标签": "月份标签", "科室显示": "开单科室（规范化）"})
st.dataframe(det, width="stretch", height=300)

c1, c2 = st.columns(2)
c1.download_button("📥 下载明细 CSV", det.to_csv(index=False).encode("utf-8-sig"),
                   "开单明细_筛选.csv", "text/csv")
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    piv3.to_excel(w, sheet_name="医生月度汇总")
    det.to_excel(w, sheet_name="筛选明细", index=False)
c2.download_button("📥 下载汇总 Excel", buf.getvalue(), "门诊开单汇总.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
