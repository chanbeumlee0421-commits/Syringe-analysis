import streamlit as st
import pandas as pd

st.set_page_config(page_title="주사기 거래 재개 병원 분석", layout="wide")
st.title("주사기 거래 재개 병원 분석")
st.caption("주사기 첫 주문 전, 다른 제품 주문이 오랫동안 없었던 병원을 찾습니다.")

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

gap_days = st.slider("공백 기준 (일)", min_value=30, max_value=1000, value=365, step=30,
                     help="이 기간 이상 주문이 없다가 주사기를 주문한 병원만 표시")

if uploaded:
    df = pd.read_excel(uploaded, header=0)
    df['매출일_date'] = pd.to_datetime(df['매출일(배송완료일)'], errors='coerce')

    direct = df[df['유통'] == '직거래'].copy()
    syringe_mask = direct['제품명'].str.contains('Syringe', case=False, na=False)
    syringe_hospitals = direct[syringe_mask]['거래처명'].unique()

    results = []
    for hosp in syringe_hospitals:
        hosp_data = direct[direct['거래처명'] == hosp].copy()
        s_mask = syringe_mask[hosp_data.index]
        syringe_data = hosp_data[s_mask]

        first_syringe = syringe_data['매출일_date'].min()
        syringe_products = syringe_data['제품명'].unique().tolist()

        non_syringe_before = hosp_data[
            (~s_mask) & (hosp_data['매출일_date'] < first_syringe)
        ]

        if len(non_syringe_before) > 0:
            last_before = non_syringe_before['매출일_date'].max()
            last_product = non_syringe_before.loc[
                non_syringe_before['매출일_date'] == last_before, '제품명'
            ].iloc[0]
            gap = (first_syringe - last_before).days
        else:
            last_before = None
            last_product = None
            gap = None

        region = hosp_data['지역1'].dropna().iloc[0] if '지역1' in hosp_data.columns and len(hosp_data['지역1'].dropna()) > 0 else ''
        manager = hosp_data['담당자'].dropna().iloc[0] if '담당자' in hosp_data.columns and len(hosp_data['담당자'].dropna()) > 0 else ''

        results.append({
            '거래처명': hosp,
            '지역': region,
            '담당자': manager,
            '주사기 첫주문일': first_syringe.strftime('%Y-%m-%d') if pd.notna(first_syringe) else '',
            '주문 주사기': ', '.join(syringe_products),
            '직전 마지막주문일': last_before.strftime('%Y-%m-%d') if last_before else '(기록없음)',
            '직전 마지막제품': last_product if last_product else '(기록없음)',
            '공백일수': gap,
        })

    result_df = pd.DataFrame(results)

    # 탭 구성
    tab1, tab2, tab3 = st.tabs([
        f" 공백 {gap_days}일 이상 재개 병원",
        "🆕 이력없는 신규 병원",
        "📋 전체 주사기 병원"
    ])

    with tab1:
        comeback = result_df[result_df['공백일수'] >= gap_days].sort_values('공백일수', ascending=False).reset_index(drop=True)
        st.metric("해당 병원 수", f"{len(comeback)}개")
        st.dataframe(
            comeback.rename(columns={'공백일수': '공백일수(일)'}),
            use_container_width=True,
            hide_index=True
        )
        csv = comeback.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(" CSV 다운로드", csv, f"재개병원_{gap_days}일이상.csv", "text/csv")

    with tab2:
        new_hosp = result_df[result_df['직전 마지막주문일'] == '(기록없음)'].reset_index(drop=True)
        st.metric("신규 병원 수", f"{len(new_hosp)}개")
        st.dataframe(
            new_hosp[['거래처명', '지역', '담당자', '주사기 첫주문일', '주문 주사기']],
            use_container_width=True,
            hide_index=True
        )

    with tab3:
        st.metric("주사기 주문 병원 전체", f"{len(result_df)}개")
        st.dataframe(
            result_df.sort_values('공백일수', ascending=False).reset_index(drop=True),
            use_container_width=True,
            hide_index=True
        )

else:
    st.info("엑셀 파일을 업로드하면 분석이 시작됩니다.")
