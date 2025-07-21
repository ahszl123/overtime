import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO
from fpdf import FPDF
import seaborn as sns
import matplotlib.font_manager as fm

# ========== 한글 폰트 설정 (이 부분을 추가) ==========
plt.rcParams['font.family'] = 'Malgun Gothic'  # Windows
plt.rcParams['axes.unicode_minus'] = False

# 만약 Malgun Gothic이 없다면 다른 폰트 시도
try:
    plt.rcParams['font.family'] = 'Malgun Gothic'
except:
    try:
        plt.rcParams['font.family'] = 'AppleGothic'  # macOS
    except:
        plt.rcParams['font.family'] = 'NanumGothic'  # Linux

st.set_page_config(layout="wide")
st.title("근태 자동 분석 프로그램 (사내표준양식 + 팀별비교분석 + 확장분석)")

uploaded_file = st.file_uploader("근태 엑셀(.xlsx) 파일을 업로드하세요.", type=["xlsx"])

# 전체 그래프 expander 크게/작게 토글
st.sidebar.markdown("### 🔍 그래프 전체 크게/작게 보기")
is_big = st.sidebar.toggle("그래프 전체 크게 보기", value=False)
chart_size = (10, 6) if is_big else (5, 3.5)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 결측치 처리
    for col in df.columns:
        if df[col].dtype == 'O':
            df[col] = df[col].fillna("미입력")
        elif pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)

    # 1. 필터 설정
    st.sidebar.markdown("## 조직(담당) 선택")
    levels = ['전체'] + sorted(df['Level 05'].astype(str).unique())
    level = st.sidebar.selectbox("조직(담당)을 선택하세요", levels)
    if level == '전체':
        orgs = ['전체'] + sorted(df['조직명'].astype(str).unique())
    else:
        orgs = ['전체'] + sorted(df[df['Level 05'] == level]['조직명'].astype(str).unique())
    org = st.sidebar.selectbox("조직(팀)을 선택하세요", orgs)
    positions = ['전체'] + sorted(df['직위명'].astype(str).unique())
    pos = st.sidebar.selectbox("직위 선택", positions)
    role_list = ['전체'] + sorted(df['직책명'].astype(str).unique())
    selected_role = st.sidebar.selectbox("직책(전체/선택)", role_list)
    names = ['전체'] + sorted(df['성명'].astype(str).unique())
    person = st.sidebar.selectbox("성명 선택", names)

    filtered = df.copy()
    if level != '전체':
        filtered = filtered[filtered['Level 05'] == level]
    if org != '전체':
        filtered = filtered[filtered['조직명'] == org]
    if pos != '전체':
        filtered = filtered[filtered['직위명'] == pos]
    if selected_role != '전체':
        filtered = filtered[filtered['직책명'] == selected_role]
    if person != '전체':
        filtered = filtered[filtered['성명'] == person]

    st.write("#### 필터링된 데이터 미리보기", filtered.head())

    # ========== 1. 실무 주요 분석 (표+그래프 통일화) ========== #

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**조직(팀)별 당월 OT(A+B) 합계**")
        org_ot = filtered.groupby('조직명')["당월 OT\n(A+B)"].sum()
        with st.expander("그래프 크게 보기", expanded=is_big):
            fig1, ax1 = plt.subplots(figsize=chart_size)
            org_ot.plot(kind="bar", ax=ax1)
            ax1.set_ylabel("OT시간")
            ax1.set_xlabel("조직(팀)")
            ax1.set_title("조직(팀)별 당월 OT(A+B) 합계")
            st.pyplot(fig1)
        st.dataframe(org_ot.reset_index().rename(columns={"당월 OT\n(A+B)":"OT시간"}), use_container_width=True)

    with col2:
        st.markdown("**성명별 당월 OT(A+B) Top 5**")
        person_ot = filtered.groupby('성명')["당월 OT\n(A+B)"].sum().sort_values(ascending=False)
        with st.expander("그래프 크게 보기", expanded=is_big):
            fig2, ax2 = plt.subplots(figsize=chart_size)
            person_ot.head(5).plot(kind="bar", ax=ax2)
            ax2.set_ylabel("OT시간")
            ax2.set_title("성명별 당월 OT(A+B) Top 5")
            st.pyplot(fig2)
        st.dataframe(person_ot.head(5).reset_index().rename(columns={"당월 OT\n(A+B)": "OT시간"}), use_container_width=True)

    st.markdown("**조직(팀)별 당월 재택일수 합계**")
    if "당월\n재택일수" in filtered.columns:
        org_wfh = filtered.groupby('조직명')["당월\n재택일수"].sum()
        col3, col4 = st.columns(2)
        with col3:
            with st.expander("그래프 크게 보기", expanded=is_big):
                fig3, ax3 = plt.subplots(figsize=chart_size)
                org_wfh.plot(kind="bar", ax=ax3)
                ax3.set_ylabel("재택일수")
                ax3.set_xlabel("조직(팀)")
                ax3.set_title("조직(팀)별 당월 재택일수 합계")
                st.pyplot(fig3)
        with col4:
            st.dataframe(org_wfh.reset_index().rename(columns={"당월\n재택일수":"재택일수"}), use_container_width=True)

    st.markdown("**근무시간 부족 상위 5명 (당월 기준)**")
    if "당월 근무부족\n시간" in filtered.columns:
        lack = filtered.groupby('성명')["당월 근무부족\n시간"].sum().sort_values(ascending=False)
        col5, col6 = st.columns(2)
        with col5:
            with st.expander("그래프 크게 보기", expanded=is_big):
                fig4, ax4 = plt.subplots(figsize=chart_size)
                lack.head(5).plot(kind="bar", ax=ax4)
                ax4.set_ylabel("근무부족(시간)")
                ax4.set_title("근무시간 부족 상위 5명 (당월 기준)")
                st.pyplot(fig4)
        with col6:
            st.dataframe(lack.head(5).reset_index().rename(columns={"당월 근무부족\n시간":"근무부족(시간)"}), use_container_width=True)

    # 팀별 비교분석 (특근비율, OT, 연차, 재택 등)
    st.header("팀별 비교 분석 (특근비율, OT, 연차, 재택 등)")

    # 특근비율/OT
    st.subheader("팀별 특근비율/OT (전월/당월) 비교")
    df["전월특근비율"] = pd.to_numeric(df["전월특근비율\n(휴일시간\n/총OT시간)"], errors="coerce").fillna(0)
    df["당월특근비율"] = pd.to_numeric(df["당월특근비율\n(휴일시간\n/총OT시간)"], errors="coerce").fillna(0)
    df["전월OT"] = pd.to_numeric(df["전월 OT\n(A+B)"], errors="coerce").fillna(0)
    df["당월OT"] = pd.to_numeric(df["당월 OT\n(A+B)"], errors="coerce").fillna(0)

    group = df.groupby("조직명").agg({
        "전월특근비율": "mean",
        "당월특근비율": "mean",
        "전월OT": "sum",
        "당월OT": "sum"
    }).reset_index()

    with st.expander("그래프 크게 보기", expanded=is_big):
        fig, ax1 = plt.subplots(figsize=chart_size)
        width = 0.2
        x = np.arange(len(group["조직명"]))
        ax1.bar(x - width, group["전월특근비율"], width, label="전월 특근비율(%)")
        ax1.bar(x, group["당월특근비율"], width, label="당월 특근비율(%)")
        ax1.bar(x + width, group["전월OT"], width, label="전월 OT(시간)")
        ax1.bar(x + 2*width, group["당월OT"], width, label="당월 OT(시간)")
        ax1.set_xticks(x + width/2)
        ax1.set_xticklabels(group["조직명"])
        ax1.legend()
        plt.title("팀별 특근비율·OT(전월/당월) 비교")
        st.pyplot(fig)
    st.dataframe(group, use_container_width=True)

    # 팀별 누적 연차사용율 비교
    st.subheader("팀별 누적 연차사용율(%) 비교")
    if "누적\n 연차사용율" in df.columns:
        df["누적연차사용율"] = pd.to_numeric(df["누적\n 연차사용율"], errors="coerce").fillna(0)
        y = df.groupby("조직명")["누적연차사용율"].mean()
        with st.expander("그래프 크게 보기", expanded=is_big):
            fig3, ax3 = plt.subplots(figsize=chart_size)
            y.plot(kind="bar", ax=ax3)
            ax3.set_ylabel("평균 연차사용률(%)")
            st.pyplot(fig3)
        st.dataframe(y.reset_index().rename(columns={"누적연차사용율":"평균 연차사용률(%)"}), use_container_width=True)

    # 재택횟수별 인원 전월/당월 비교
    st.subheader("재택횟수별 인원 (전월/당월) 분포 비교")
    def group_bin(val):
        if val < 1: return "미경험"
        elif val < 3: return "3회 미만"
        elif val < 5: return "3회이상"
        elif val < 10: return "5회이상"
        elif val < 15: return "10회이상"
        else: return "15회이상"

    df["전월재택일수"] = pd.to_numeric(df["전월 재택일수"], errors="coerce").fillna(0)
    df["당월재택일수"] = pd.to_numeric(df["당월\n재택일수"], errors="coerce").fillna(0)
    df["전월구간"] = df["전월재택일수"].apply(group_bin)
    df["당월구간"] = df["당월재택일수"].apply(group_bin)

    bin_order = ["미경험","3회 미만","3회이상","5회이상","10회이상","15회이상"]
    prev_counts = df["전월구간"].value_counts().reindex(bin_order, fill_value=0)
    curr_counts = df["당월구간"].value_counts().reindex(bin_order, fill_value=0)
    idx = np.arange(len(bin_order))
    with st.expander("그래프 크게 보기", expanded=is_big):
        fig4, ax4 = plt.subplots(figsize=chart_size)
        ax4.bar(idx - 0.15, prev_counts, width=0.3, label="전월")
        ax4.bar(idx + 0.15, curr_counts, width=0.3, label="당월")
        ax4.set_xticks(idx)
        ax4.set_xticklabels(bin_order)
        ax4.set_ylabel("인원 수")
        ax4.set_title("재택횟수별 인원 전월/당월 비교")
        ax4.legend()
        st.pyplot(fig4)
    st.dataframe(pd.DataFrame({"전월":prev_counts,"당월":curr_counts}), use_container_width=True)

    st.divider()  # 시각적으로 분석 영역 구분 (Streamlit 1.27 이상)

    # ========== 2. 확장·추가 분석 및 시각화 ========== #
    st.header("🔎 추가 인사 분석 및 시각화")

    # (1) OT 상위/하위 TOP 10
    st.subheader("1. OT 상위/하위 TOP 10")
    df["당월OT"] = pd.to_numeric(df["당월 OT\n(A+B)"], errors="coerce").fillna(0)
    ot_top = df.groupby("성명")["당월OT"].sum().sort_values(ascending=False)
    ot_bottom = df.groupby("성명")["당월OT"].sum().sort_values()
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**OT 상위 TOP 10**")
        st.dataframe(ot_top.head(10).reset_index().rename(columns={"당월OT":"OT시간"}), use_container_width=True)
    with col2:
        st.markdown("**OT 하위 TOP 10**")
        st.dataframe(ot_bottom.head(10).reset_index().rename(columns={"당월OT":"OT시간"}), use_container_width=True)
    with st.expander("그래프 크게 보기", expanded=is_big):
        fig, ax = plt.subplots(figsize=chart_size)
        ot_top.head(10).plot(kind="bar", ax=ax)
        ax.set_title("OT 상위 TOP 10")
        st.pyplot(fig)

    # (2) 재택-OT 상관관계(Scatter Plot)
    st.subheader("2. 재택근무일수-OT 상관관계")
    if "당월\n재택일수" in df.columns:
        df["재택일수"] = pd.to_numeric(df["당월\n재택일수"], errors="coerce").fillna(0)
        with st.expander("그래프 크게 보기", expanded=is_big):
            fig, ax = plt.subplots(figsize=chart_size)
            sns.scatterplot(data=df, x='재택일수', y='당월OT', hue='조직명', ax=ax)
            ax.set_title("재택근무일수와 OT(당월)의 상관관계")
            st.pyplot(fig)

    # (3) 부서별 근무유형 분포(Pie Chart)
    st.subheader("3. 팀별 근무유형 분포 (Pie Chart)")
    if "근무유형" in df.columns:
        team_list = sorted(df['조직명'].unique())
        team = st.selectbox("분포를 볼 팀 선택", team_list)
        subset = df[df['조직명'] == team]
        g = subset['근무유형'].value_counts()
        with st.expander("그래프 크게 보기", expanded=is_big):
            fig, ax = plt.subplots(figsize=chart_size)
            ax.pie(g, labels=g.index, autopct='%1.1f%%')
            ax.set_title(f"{team} 근무유형 분포")
            st.pyplot(fig)
        st.dataframe(g.reset_index().rename(columns={"index":"근무유형", "근무유형":"건수"}), use_container_width=True)

    # (4) 연차 미사용자(30% 미만) 리스트
    st.subheader("4. 연차사용률 30% 미만 직원")
    if "누적\n 연차사용율" in df.columns:
        df["누적연차"] = pd.to_numeric(df["누적\n 연차사용율"], errors="coerce").fillna(0)
        low_annual = df[df["누적연차"] < 30][["성명","조직명","누적연차"]].sort_values("누적연차")
        st.dataframe(low_annual, use_container_width=True)

    # (5) 근무부족자/지각/결근자
    st.subheader("5. 근무부족/지각/결근 다발 인원")
    if "당월 근무부족\n시간" in df.columns:
        lack = df[df["당월 근무부족\n시간"] > 0][["성명","조직명","당월 근무부족\n시간"]]
        st.markdown("**근무부족(시간>0) 인원**")
        st.dataframe(lack, use_container_width=True)
    if "지각" in df.columns:
        late = df[df["지각"] != "N"][["성명","조직명","지각"]]
        st.markdown("**지각 인원**")
        st.dataframe(late, use_container_width=True)
    if "결근" in df.columns:
        absn = df[df["결근"] != "N"][["성명","조직명","결근"]]
        st.markdown("**결근 인원**")
        st.dataframe(absn, use_container_width=True)

    # (6) 재택 10회+OT 20시간 이상자
    st.subheader("6. 재택 10회+OT 20시간 이상자")
    if "당월재택일수" in df.columns:
        cond_wfh = df['당월재택일수'] >= 10
        wfh_df = df[cond_wfh]
        st.markdown(f"### 재택 10회 이상 직원 수: **{len(wfh_df)}명**")
        st.dataframe(wfh_df[["성명", "조직명", "당월재택일수", "당월OT"]], use_container_width=True)

    if "당월OT" in df.columns:
        cond_ot = df['당월OT'] >= 20
        ot_df = df[cond_ot]
        st.markdown(f"### OT 20시간 이상 직원 수: **{len(ot_df)}명**")
        st.dataframe(ot_df[["성명", "조직명", "당월재택일수", "당월OT"]], use_container_width=True)

    # (8) 자동 요약 리포트
    st.subheader("8. 자동 요약 리포트")
    top_team = df.groupby("조직명")["당월OT"].sum().sort_values(ascending=False).head(1)
    low_annual_team = df.groupby("조직명")["누적연차"].mean().sort_values().head(1)
    st.markdown(f"- 이번 달 OT 상위 1위 팀: **{top_team.index[0]}** ({top_team.values[0]:.1f}시간)")
    st.markdown(f"- 연차 사용률 최저 팀: **{low_annual_team.index[0]}** ({low_annual_team.values[0]:.1f}%)")
    st.markdown(f"- 재택 10회 이상 & OT 20시간 이상 직원 수: **{len(combo)}명**")
    st.markdown(f"- 연차 30% 미만 직원 수: **{len(low_annual)}명**")

    # === [추가 시각화] === #
    st.header("📊 추가 시각화")

    # Heatmap: 조직(팀)별 OT
    st.subheader("Heatmap: 팀별 OT/재택/근무부족 시간")
    pivot = pd.pivot_table(df, index='조직명', values=['당월OT','재택일수','당월 근무부족\n시간'], aggfunc='sum')
    with st.expander("그래프 크게 보기", expanded=is_big):
        fig, ax = plt.subplots(figsize=chart_size)
        sns.heatmap(pivot, annot=True, cmap="YlOrRd", ax=ax)
        st.pyplot(fig)
    st.dataframe(pivot, use_container_width=True)

    # Boxplot: OT 분포
    st.subheader("Boxplot: 조직별 OT 분포")
    with st.expander("그래프 크게 보기", expanded=is_big):
        fig, ax = plt.subplots(figsize=chart_size)
        sns.boxplot(data=df, x='조직명', y='당월OT', ax=ax)
        st.pyplot(fig)

    # Line Chart: (월별 OT/특근/연차) - 샘플컬럼 있으면 자동
    if "근무월" in df.columns:
        st.subheader("LineChart: 월별 OT/특근/연차 추이")
        df["근무월"] = pd.to_datetime(df["근무월"], errors="coerce")
        grp = df.groupby([df["근무월"].dt.to_period('M'),"조직명"]).agg({"당월OT":"sum"})
        grp = grp.reset_index()
        with st.expander("그래프 크게 보기", expanded=is_big):
            fig, ax = plt.subplots(figsize=chart_size)
            for team in grp['조직명'].unique():
                team_data = grp[grp['조직명']==team]
                ax.plot(team_data["근무월"].astype(str), team_data["당월OT"], label=team)
            ax.legend()
            ax.set_title("월별 OT 추이")
            st.pyplot(fig)

    # --- 엑셀 및 PDF 저장 버튼 추가 --- #
    st.markdown("### 📁 분석결과 다운로드")
    result_excel = BytesIO()
    with pd.ExcelWriter(result_excel, engine='xlsxwriter') as writer:
        filtered.to_excel(writer, index=False, sheet_name='Filtered')
        ot_top.head(10).reset_index().to_excel(writer, index=False, sheet_name='OT_TOP10')
        pivot.to_excel(writer, index=True, sheet_name='HeatmapData')
        # 추가 필요시 원하는 표도 저장
    result_excel.seek(0)
    st.download_button("엑셀로 결과 저장", data=result_excel, file_name="분석결과_확장.xlsx", mime="application/vnd.ms-excel")

    # PDF (한글 폰트 사용 예시 - 폰트파일 프로젝트 폴더에 위치)
    def make_pdf(df, title="근태 분석 리포트"):
        pdf = FPDF()
        pdf.add_page()
        pdf.add_font('Nanum', '', 'NanumGothic.ttf', uni=True)
        pdf.set_font('Nanum', size=12)
        pdf.cell(0, 10, txt=title, ln=True, align='C')
        pdf.ln(5)
        for i, row in df.iterrows():
            pdf.cell(0, 8, txt=', '.join(map(str, row)), ln=True)
        return pdf.output(dest='S').encode('latin1')
    try:
        pdf_bytes = make_pdf(ot_top.head(10).reset_index())
        st.download_button("PDF로 저장(OT TOP10)", data=pdf_bytes, file_name="OT_TOP10.pdf", mime="application/pdf")
    except Exception as e:
        st.warning("PDF저장 오류: 한글폰트파일(NanumGothic.ttf)이 같은 폴더에 없거나 한글 폰트 관련 오류입니다.")

    st.info("원하는 분석/표/그래프 더 추가, 스타일/자동 코멘트/다운로드 모두 언제든 요청!")

else:
    st.info("회사 표준 양식(1행 컬럼명) 그대로의 엑셀 파일을 업로드하세요!")
