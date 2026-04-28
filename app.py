import streamlit as st
import pandas as pd

st.set_page_config(page_title="주사기 거래 재개 병원 분석", layout="wide")
st.title("주사기 거래 재개 병원 분석")
st.caption("2026년 4월 이후 주사기를 첫 주문한 병원 중, 이전 주문과 공백이 긴 병원을 찾습니다.")

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

gap_days = st.slider("공백 기준 (일)", min_value=30, max_value=1000, value=365, step=1,
                     help="이 기간 이상 주문이 없다가 주사기를 주문한 병원만 표시")

METRO = {'서울', '경기', '인천'}

def format_region(r1, r2):
    if pd.isna(r1):
        return ''
    r1 = str(r1).strip()
    if r1 in METRO:
        if pd.notna(r2):
            r2 = str(r2).strip()
            if r2.endswith('시') or r2.endswith('구'):
                r2 = r2[:-1]
            return r2
        return r1
    return r1

def fmt_won(val):
    if pd.isna(val):
        return '-'
    return f"{int(val):,}원"

if uploaded:
    df = pd.read_excel(uploaded, header=0)
    df['매출일_date'] = pd.to_datetime(df['매출일(배송완료일)'], errors='coerce')

    direct = df[df['유통'] == '직거래'].copy()
    syringe_mask = direct['제품명'].str.contains('Syringe|주사기', case=False, na=False) & ~direct['제품명'].str.contains('INSULIN', case=False, na=False)
    syringe_hospitals = direct[syringe_mask]['거래처명'].unique()

    results = []
    for hosp in syringe_hospitals:
        hosp_data = direct[direct['거래처명'] == hosp].copy()
        s_mask = syringe_mask[hosp_data.index]
        syringe_data = hosp_data[s_mask]

        first_syringe = syringe_data['매출일_date'].min()

        if pd.isna(first_syringe) or first_syringe < pd.Timestamp('2026-04-01'):
            continue

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

        recent = hosp_data[hosp_data['매출일_date'] >= pd.Timestamp('2026-04-01')]
        non_syringe_recent = recent[~syringe_mask[recent.index]]
        if len(non_syringe_recent) > 0:
            other_products = (
                non_syringe_recent.groupby('제품명')['매출수량']
                .sum()
                .apply(lambda x: int(x))
                .reset_index()
                .apply(lambda r: f"{r['제품명']} ({r['매출수량']}개)", axis=1)
                .tolist()
            )
        else:
            other_products = []

        total_all = hosp_data['매출액(vat 제외)'].sum()
        total_recent = recent['매출액(vat 제외)'].sum()

        r1 = hosp_data['지역1'].dropna().iloc[0] if '지역1' in hosp_data.columns and len(hosp_data['지역1'].dropna()) > 0 else None
        r2 = hosp_data['지역2'].dropna().iloc[0] if '지역2' in hosp_data.columns and len(hosp_data['지역2'].dropna()) > 0 else None
        manager = hosp_data['담당자'].dropna().iloc[0] if '담당자' in hosp_data.columns and len(hosp_data['담당자'].dropna()) > 0 else ''

        results.append({
            '거래처명': hosp,
            '지역': format_region(r1, r2),
            '담당자': manager,
            '주사기 첫주문일': first_syringe.strftime('%Y-%m-%d') if pd.notna(first_syringe) else '',
            '주문 주사기': ', '.join(syringe_products),
            '동반 주문 제품': ', '.join(other_products) if other_products else '(주사기만 주문)',
            '이번달 매출(VAT제외)': total_recent,
            '누적 매출(VAT제외)': total_all,
            '직전 마지막주문일': last_before.strftime('%Y-%m-%d') if last_before else '(기록없음)',
            '직전 마지막제품': last_product if last_product else '(기록없음)',
            '공백일수': gap,
        })

    result_df = pd.DataFrame(results)

    tab1, tab2 = st.tabs([
        f"공백 {gap_days}일 이상 재개 병원",
        "이력없는 신규 병원",
    ])

    with tab1:
        comeback = result_df[result_df['공백일수'] >= gap_days].sort_values('공백일수', ascending=False).reset_index(drop=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("해당 병원 수", f"{len(comeback)}개")


        display = comeback.copy()
        display['이번달 매출(VAT제외)'] = display['이번달 매출(VAT제외)'].apply(fmt_won)
        display['누적 매출(VAT제외)'] = display['누적 매출(VAT제외)'].apply(fmt_won)
        display = display.rename(columns={'공백일수': '공백일수(일)'})

        st.dataframe(display, use_container_width=True, hide_index=True, height=36 * len(comeback) + 38)
        csv = comeback.to_csv(index=False, encoding='utf-8-sig')
        st.download_button("CSV 다운로드", csv, f"재개병원_{gap_days}일이상.csv", "text/csv")

    with tab2:
        new_hosp = result_df[result_df['직전 마지막주문일'] == '(기록없음)'].reset_index(drop=True)
        st.metric("신규 병원 수", f"{len(new_hosp)}개")
        st.dataframe(
            new_hosp[['거래처명', '지역', '담당자', '주사기 첫주문일', '주문 주사기', '동반 주문 제품', '이번달 매출(VAT제외)']].assign(
                **{'이번달 매출(VAT제외)': new_hosp['이번달 매출(VAT제외)'].apply(fmt_won)}
            ),
            use_container_width=True, hide_index=True, height=36 * len(new_hosp) + 38
        )

else:
    st.info("엑셀 파일을 업로드하면 분석이 시작됩니다.")
