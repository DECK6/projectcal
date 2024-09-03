import streamlit as st
import pandas as pd
import plotly.figure_factory as ff
import plotly.graph_objs as go
import requests
import io
from datetime import datetime, timedelta
import traceback
import re

# 스트림릿 페이지 설정을 와이드 모드로 변경
st.set_page_config(layout="wide")

# 특정 Google Sheets를 CSV로 공개한 URL
SHEET_URL = "https://docs.google.com/spreadsheets/d/1bm08ehPCzorBhhZFm6PocZZgGMN9LiOfEXwrOxUiX7U/export?format=csv&gid=37685594"

@st.cache_data
def load_data():
    try:
        response = requests.get(SHEET_URL)
        response.raise_for_status()
        csv_content = response.content.decode('utf-8')
        df = pd.read_csv(io.StringIO(csv_content), header=None)
        header_row = df[df.apply(lambda row: row.astype(str).str.contains('사업명').any(), axis=1)].index
        if not header_row.empty:
            header_row = header_row[0]
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
        else:
            st.error("'사업명' 열을 찾을 수 없습니다. 데이터 구조를 확인해주세요.")
            return pd.DataFrame()
        return df
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {str(e)}")
        st.error(traceback.format_exc())
        return pd.DataFrame()

def parse_date(date_str):
    if pd.isna(date_str):
        return pd.NaT
    
    original_date_str = date_str
    date_str = str(date_str).strip()
    
    if date_str in ['사전규격', '견적서 요청', '실행중']:
        return pd.NaT
    
    if '이전' in date_str:
        date_part = re.search(r'\d{4}-\d{2}-\d{2}', date_str)
        if date_part:
            return pd.to_datetime(date_part.group()) - timedelta(days=1)
    
    if '행사' in date_str:
        date_part = re.search(r'\d{4}\.\d{2}\.\d{2}', date_str)
        if date_part:
            return pd.to_datetime(date_part.group().replace('.', '-'))
    
    am_pm_match = re.search(r'(오전|오후)', date_str)
    if am_pm_match:
        am_pm = am_pm_match.group(1)
        date_str = date_str.replace(am_pm, '').strip()
        try:
            if am_pm == '오전':
                dt = pd.to_datetime(date_str, format='%Y.%m.%d %I:%M')
                if dt.hour == 12:
                    dt -= timedelta(hours=12)
            elif am_pm == '오후':
                dt = pd.to_datetime(date_str, format='%Y.%m.%d %I:%M')
                if dt.hour != 12:
                    dt += timedelta(hours=12)
            return dt
        except ValueError:
            pass
    
    if re.match(r'^\d{4}$', date_str):
        return pd.to_datetime(date_str, format='%Y')
    
    if re.match(r'^\d{4}\.\d{2}$', date_str):
        return pd.to_datetime(date_str, format='%Y.%m')
    
    date_formats = [
        '%Y-%m-%d',
        '%Y.%m.%d',
        '%Y-%m-%d %H:%M',
        '%Y.%m.%d %H:%M',
        '%Y-%m-%d %p %I:%M',
        '%Y.%m.%d %p %I:%M',
        '%Y-%m-%d %H:%M:%S',
        '%Y.%m.%d %H:%M:%S',
        '%Y. %m. %d. %H:%M',
        '%Y.%m.%d %H:%M',
        '%Y.%m.%d.%H',
    ]
    
    for fmt in date_formats:
        try:
            return pd.to_datetime(date_str, format=fmt)
        except ValueError:
            continue
    
    return pd.NaT

def find_column(df, keywords):
    for col in df.columns:
        if any(keyword.lower() in str(col).lower() for keyword in keywords):
            return col
    return None

def generate_colors(n):
    import matplotlib.pyplot as plt
    import numpy as np
    cmap = plt.get_cmap("tab20")
    colors = cmap(np.linspace(0, 1, n))
    return ['rgb({}, {}, {})'.format(int(color[0]*255), int(color[1]*255), int(color[2]*255)) for color in colors]

# Streamlit 앱 시작
st.title('프로젝트 일정 관리 대시보드')

# 데이터 로드
df = load_data()

if not df.empty:
    business_name_col = find_column(df, ['사업명'])
    end_date_col = find_column(df, ['제출일', '종료일'])
    manager_col = find_column(df, ['담당자'])

    if business_name_col and end_date_col and manager_col:
        try:
            df = df.dropna(subset=[business_name_col, end_date_col])
            df[end_date_col] = df[end_date_col].apply(parse_date)
            df['시작일'] = df[end_date_col] - timedelta(days=14)
            df = df.dropna(subset=['시작일', end_date_col])
            df[manager_col] = df[manager_col].fillna("Unknown").astype(str)

            unique_managers = df[manager_col].unique()
            colors = generate_colors(len(unique_managers))

            tasks = []
            annotations = []  # 각 바 끝에 추가할 디데이 텍스트를 저장할 리스트
            for index, row in df.iterrows():
                task = dict(
                    Task=row[business_name_col],
                    Start=row['시작일'].strftime('%Y-%m-%d'),
                    Finish=row[end_date_col].strftime('%Y-%m-%d'),
                    Resource=row[manager_col]
                )
                tasks.append(task)

                # 디데이 계산
                d_day = (row[end_date_col] - datetime.now()).days
                d_day_text = f"D-{d_day}" if d_day >= 0 else f"D+{-d_day}"

                # 디데이 텍스트를 바 끝에 추가
                annotations.append(dict(
                    x=row[end_date_col].strftime('%Y-%m-%d'),
                    y=index + 0.5,  # 각 바의 중간에 위치하도록 조정
                    text=d_day_text,
                    showarrow=False,
                    xanchor='left',
                    font=dict(color='black', size=10)
                ))

            fig = ff.create_gantt(
                tasks, 
                index_col='Resource', 
                show_colorbar=True, 
                group_tasks=True, 
                showgrid_x=True, 
                showgrid_y=True,
                colors=colors
            )

            # 높이 설정
            fig.update_layout(
                height=40 * len(tasks),  # 각 항목에 대한 높이 설정 (항목당 40 픽셀)
                annotations=annotations  # 디데이 텍스트 추가
            )

            # 가로축을 하루 단위로 설정
            fig.update_xaxes(
                tickformat="%b %d",  # 날짜 포맷
                dtick=24*60*60*1000  # 하루 단위로 설정 (밀리초 단위, 24*60*60*1000 밀리초 = 1일)
            )

            fig.add_vline(x=datetime.now().date(), line_width=2, line_dash="dash", line_color="red")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("프로젝트 상세 정보")
            display_columns = [col for col in [business_name_col, '수요기관(발주처)', '캠프명', manager_col, '사업 금액(VAT포함)', '시작일', end_date_col] if col in df.columns]
            st.dataframe(df[display_columns])
        except Exception as e:
            st.error(f"그래프 생성 중 오류 발생: {str(e)}")
            st.error(traceback.format_exc())
    else:
        st.write("필요한 열을 찾을 수 없습니다. 데이터 구조를 확인해주세요.")
else:
    st.write("데이터를 가져오는 데 실패했습니다. 파일의 내용을 확인해주세요.")
