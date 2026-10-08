import os
import re
import json
import base64
import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# 1. 페이지 설정
st.set_page_config(page_title="설비 유지 보수 종합 대시보드", layout="wide")

# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (확장 메뉴 호버/포커스 색상 완벽 분리)
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (확장 메뉴 및 버튼 텍스트 색상 완벽 수정)
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (버튼 디자인 완벽 통일)
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (확장 메뉴 및 버튼 디자인 완벽 통일)
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (클릭 시 팝업 백화현상 완벽 차단)
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (스트림릿 포커스 백화현상 완벽 차단)
# 🌟 상단 공백 제거 및 남색 배경 동기화 CSS (스트림릿 포커스 방어 & 가로 길이 축소)
st.markdown(
    """
    <style>
    header {visibility: hidden;}
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 0rem !important;
    }
    .stApp {
        background-color: #071321 !important;
        color: white !important;
    }
    h1, h2, h3, p {
        color: #ffffff !important;
    }

    /* 1. 확장 메뉴 (Expander) 철벽 방어 */
    div[data-testid="stExpander"] details summary,
    div[data-testid="stExpander"] details summary:hover,
    div[data-testid="stExpander"] details summary:focus,
    div[data-testid="stExpander"] details summary:active,
    div[data-testid="stExpander"] details:focus-within summary {
        background-color: #0f233c !important;
        border: 1px solid #1e3a5f !important;
        border-radius: 8px !important;
        color: #ffffff !important;
    }
    div[data-testid="stExpander"] details summary p {
        color: #ffffff !important; 
    }
    div[data-testid="stExpander"] details summary svg {
        fill: #ffffff !important; 
    }

    /* 🌟 2. 파일 업로드 구역 뚱뚱한 가로 길이 다이어트! */
    div[data-testid="stFileUploader"] {
        max-width: 500px !important; /* 👈 최대 너비를 500px로 묶어서 아담하고 세련되게 만듭니다 */
    }

    /* 3. 파일 업로드 박스 전체 영역 (Dropzone) 하얀색 철벽 방어 */
    div[data-testid="stFileUploadDropzone"],
    div[data-testid="stFileUploadDropzone"]:hover,
    div[data-testid="stFileUploadDropzone"]:focus,
    div[data-testid="stFileUploadDropzone"]:active,
    div[data-testid="stFileUploadDropzone"]:focus-within {
        background-color: #0f233c !important;
        border: 1px dashed #4fa0ff !important;
    }
    /* 업로드 박스 안의 회색 텍스트 유지 */
    div[data-testid="stFileUploadDropzone"] * {
        color: #a0aec0 !important; 
    }

    /* 4. 버튼 상태 고정 (팝업 뜰 때 하얘지는 것 방지) */
    div[data-testid="stFileUploader"] button,
    div[data-testid="stButton"] button,
    div[data-testid="stFileUploader"] button:focus,
    div[data-testid="stButton"] button:focus,
    div[data-testid="stFileUploader"] button:active,
    div[data-testid="stButton"] button:active {
        background-color: #0f233c !important;
        border: 1px solid #4fa0ff !important;
        border-radius: 6px !important;
    }
    div[data-testid="stFileUploader"] button p,
    div[data-testid="stButton"] button p {
        color: #ffffff !important;
        font-weight: bold !important;
    }

    /* 마우스를 올렸을 때만 살짝 밝아지는 효과 */
    div[data-testid="stFileUploader"] button:hover,
    div[data-testid="stButton"] button:hover {
        background-color: #1a365d !important;
        border-color: #63b3ed !important;
        box-shadow: 0 0 8px rgba(99, 179, 237, 0.4) !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("🏭 스마트 설비 유지보수 & 고장 이력 통합 대시보드")
st.write("상단 메뉴에서 새로운 월별 PDF 보고서를 업로드하면, 원본 엑셀이 자동 확장되며 대시보드와 실시간 연동됩니다.")

current_dir = os.path.dirname(os.path.abspath(__file__))
excel_file_path = os.path.join(current_dir, "고장이력_트렌드_무결점_최종.xlsx")
bems_file_path = os.path.join(current_dir, "벰스.xlsx")

images_dir = os.path.join(current_dir, "images")
if not os.path.exists(images_dir):
    images_dir = os.path.join(current_dir, "static")

# -------------------------------------------------------------------------
# [파트 1] 정밀 PDF 분석 및 엑셀 자동 확장 갱신 모듈
# -------------------------------------------------------------------------
with st.expander("📁 [여기서 클릭] 신규 월별 PDF 업로드 및 자동 확장 갱신", expanded=False):
    uploaded_pdfs = st.file_uploader("📂 새로운 월별 PDF 보고서들 업로드 (여러 개 선택 가능)", type=["pdf"], accept_multiple_files=True)

    if st.button("🚀 PDF 분석 및 엑셀 자동 확장 반영"):
        if uploaded_pdfs:
            if not os.path.exists(excel_file_path):
                st.error(f"⚠️ 폴더 내에서 원본 엑셀 파일을 찾을 수 없습니다!")
            else:
                temp_dir = os.path.join(current_dir, "pdf_temp")
                os.makedirs(temp_dir, exist_ok=True)

                for pdf_file in uploaded_pdfs:
                    pdf_path = os.path.join(temp_dir, pdf_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_file.getbuffer())

                with st.spinner("🔄 PDF 정밀 스캔 및 신규 날짜 자동 확장 분석 중..."):
                    def expand_shorthand(line):
                        for _ in range(3):
                            line = re.sub(r'([A-Za-z]-)(\d+)\s*,\s*(\d+)', r'\g<1>\2, \g<1>\3', line)
                        return line


                    daily_faults = {}
                    date_pattern = re.compile(r'(202\d)년\s*(\d+)월\s*(\d+)일')

                    for file_name in os.listdir(temp_dir):
                        if not file_name.lower().endswith('.pdf'):
                            continue
                        pdf_path = os.path.join(temp_dir, file_name)
                        try:
                            with pdfplumber.open(pdf_path) as pdf:
                                current_date = None
                                current_stage = "공통"
                                skip_mode = False

                                for page in pdf.pages:
                                    text = page.extract_text()
                                    if not text: continue

                                    for line in text.split('\n'):
                                        line_stripped = line.strip()
                                        line_no_space = line.replace(" ", "")
                                        line_lower = line.lower()

                                        date_match = date_pattern.search(line)
                                        if date_match:
                                            y, m, d = date_match.groups()
                                            current_date = f"{y[-2:]}.{int(m):02d}.{int(d):02d}"
                                            if current_date not in daily_faults:
                                                daily_faults[current_date] = []
                                            skip_mode = False
                                            current_stage = "공통"
                                            continue

                                        if "[1단계]" in line_no_space:
                                            current_stage = "1단계"; skip_mode = False
                                        elif "[2단계]" in line_no_space:
                                            current_stage = "2단계"; skip_mode = False
                                        elif "[3단계]" in line_no_space:
                                            current_stage = "3단계"; skip_mode = False
                                        elif any(sec in line_no_space for sec in ["주간", "야간", "설비현황"]):
                                            skip_mode = False

                                        if line_no_space == "비고" or line_no_space == "비" or line_stripped.startswith(
                                                "비고"):
                                            skip_mode = True
                                            continue

                                        if skip_mode: continue

                                        if current_date and any(kw in line_lower for kw in
                                                                ['fault', '불량', '파손', '이탈', '누기', '탈락', '보수']):
                                            expanded_line = expand_shorthand(line)
                                            daily_faults[current_date].append((current_stage, expanded_line))
                        except Exception:
                            pass

                    df = pd.read_excel(excel_file_path, sheet_name=0, engine='openpyxl')
                    equip_cols = df.columns[:4]
                    df[equip_cols] = df[equip_cols].ffill()

                    date_col_map = {}
                    for col in df.columns:
                        col_str = str(col).split(' ')[0]
                        if re.search(r'\d{2}\.\d{2}', col_str) or isinstance(col, pd.Timestamp) or re.search(r'202\d',
                                                                                                             col_str):
                            if '-' in col_str and len(col_str) >= 10:
                                try:
                                    col_str = pd.to_datetime(col_str).strftime('%y.%m.%d')
                                except:
                                    pass
                            date_col_map[col_str] = col

                    for pdf_date_str in daily_faults.keys():
                        if pdf_date_str not in date_col_map:
                            df[pdf_date_str] = None
                            date_col_map[pdf_date_str] = pdf_date_str

                    stage_3_kws = ["3차", "여과", "역세척"]
                    categories = ['펌프', '송풍기', '스크린', '제진기', '컨베이어', '인양기', '수집기', '탈수기', '계면', '교반기', '냉각수', '냉각탑',
                                  '소포수', '잡용수', '농축기']

                    row_props = []
                    for idx, row in df.iterrows():
                        cols_vals = [str(x).replace(' ', '') for x in row[equip_cols] if
                                     pd.notna(x) and str(x) != 'nan']
                        if not cols_vals:
                            row_props.append(None)
                            continue

                        full_equip_str = "".join(cols_vals)
                        if "컨베이어" in full_equip_str and ("1단계" in full_equip_str or "2단계" in full_equip_str):
                            df.at[idx, equip_cols[1]] = "1단계 E"
                            full_equip_str = full_equip_str.replace("1단계", "1단계E")

                        last_val = str(row[equip_cols[-1]]).strip()
                        req_stage = "1단계" if "1단계" in full_equip_str else ("2단계" if "2단계" in full_equip_str else None)
                        req_process = []
                        if any(x in full_equip_str for x in ["1차", "초침"]):
                            req_process = ["1차", "초침"]
                        elif any(x in full_equip_str for x in ["2차", "종침", "이차"]):
                            req_process = ["2차", "종침", "이차"]

                        row_props.append({
                            'idx': idx,
                            'full_str': full_equip_str,
                            'is_3rd': any(kw in full_equip_str for kw in stage_3_kws),
                            'req_stage': req_stage,
                            'req_process': req_process,
                            'my_cats': [cat for cat in categories if cat in full_equip_str],
                            'dash_tags': re.findall(r'[A-Za-z]-\d+', last_val),
                            'alpha_tags': re.findall(r'(?<![A-Za-z])[A-Za-z](?![A-Za-z\-])', last_val),
                            'num_tags': re.findall(r'\d+', last_val)
                        })

                    cleaned_daily_faults = {}
                    for d_col, items in daily_faults.items():
                        cleaned_items = []
                        for log_stage, log_line in items:
                            line_clean = re.sub(r'^\s*\d+\.\s*', '', log_line)
                            line_super_clean = re.sub(r'1단계|2단계|3단계|1차|2차|3차|A계열|B계열|C계열', '', line_clean)
                            is_3rd_log = any(kw in line_clean for kw in stage_3_kws)
                            cleaned_items.append((log_stage, line_clean, line_super_clean, is_3rd_log))
                        cleaned_daily_faults[d_col] = cleaned_items

                    for date_col, pre_cleaned_faults in cleaned_daily_faults.items():
                        if date_col not in date_col_map: continue
                        actual_excel_col = date_col_map[date_col]
                        df[actual_excel_col] = df[actual_excel_col].astype(object)

                        for props in row_props:
                            if props is None: continue
                            is_fault = False
                            for log_stage, line_clean, line_super_clean, is_3rd_log in pre_cleaned_faults:
                                if props['is_3rd'] != is_3rd_log: continue
                                if props['req_stage'] and (props['req_stage'] != log_stage and props[
                                    'req_stage'] not in line_clean): continue
                                if props['req_process'] and not any(
                                    p in line_clean for p in props['req_process']): continue

                                if props['my_cats']:
                                    cat_matched = False
                                    for cat in props['my_cats']:
                                        if cat in line_clean: cat_matched = True; break
                                        if cat == '수집기' and (
                                                '찌꺼기' in line_clean or '슬러지' in line_clean): cat_matched = True; break
                                    if not cat_matched: continue

                                tag_matched = False
                                if props['dash_tags']:
                                    if any(dt in line_super_clean for dt in props['dash_tags']): tag_matched = True
                                elif props['alpha_tags']:
                                    for at in props['alpha_tags']:
                                        if re.search(r'(?<![A-Za-z])' + at + r'(?![A-Za-z\-])',
                                                     line_super_clean): tag_matched = True; break
                                elif props['num_tags']:
                                    line_nums = re.findall(r'\d+', line_super_clean)
                                    if any(nt in line_nums for nt in props['num_tags']): tag_matched = True
                                else:
                                    if props['my_cats']: tag_matched = True

                                if tag_matched:
                                    is_fault = True
                                    break

                            if is_fault:
                                existing_val = df.at[props['idx'], actual_excel_col]
                                if pd.isna(existing_val):
                                    df.at[props['idx'], actual_excel_col] = 'Fault'
                                elif 'Fault' not in str(existing_val):
                                    df.at[props['idx'], actual_excel_col] = str(existing_val) + ', Fault'

                    df.to_excel(excel_file_path, index=False, engine='openpyxl')

                st.success("🎉 PDF 분석 완료! 원본 엑셀이 자동으로 확장·갱신되었습니다.")
                st.rerun()
        else:
            st.warning("⚠️ 분석할 PDF 파일들을 업로드해 주세요!")

st.markdown("<br>", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# [파트 2] 벰스(BEMS) 파일 로드 및 메타데이터 매핑
# -------------------------------------------------------------------------
bems_meta_map = {}
if os.path.exists(bems_file_path):
    try:
        df_bems = pd.read_excel(bems_file_path, sheet_name=0, engine='openpyxl')

        # 엑셀의 열 이름(헤더)을 직접 지정하여 데이터를 가져옵니다.
        for _, r in df_bems.iterrows():
            # 💡 아래 3줄과 if 문의 시작 위치(세로줄)가 정확히 맞아야 에러가 안 납니다!
            o_val = r.iloc[14]  # 기존 O열 (설비 경로/이름)
            q_val = r['설비규격']  # Q열
            ab_val = r['취득일자']  # AB열

            if pd.notna(o_val):
                key_norm = re.sub(r'\s+', '', str(o_val))
                spec_str = str(q_val).strip() if pd.notna(q_val) and str(q_val).strip() != 'nan' else "-"
                date_str = str(ab_val).split(' ')[0] if pd.notna(ab_val) and str(ab_val).strip() != 'nan' else "공단자체취득"

                if hasattr(ab_val, 'strftime'):
                    date_str = ab_val.strftime('%Y-%m-%d')

                bems_meta_map[key_norm] = {
                    "capacity": spec_str,
                    "install_date": date_str
                }
    except Exception as e:
        pass
# -------------------------------------------------------------------------
# [파트 3] 스마트 이미지 매칭 및 Base64 변환 로직
# -------------------------------------------------------------------------
available_images = []
if os.path.exists(images_dir):
    available_images = [f for f in os.listdir(images_dir)]


def get_image_base64(image_filename):
    img_path = os.path.join(images_dir, image_filename)
    if os.path.exists(img_path):
        try:
            with open(img_path, "rb") as img_file:
                encoded = base64.b64encode(img_file.read()).decode('utf-8')
                ext = image_filename.split('.')[-1].lower()
                mime = "image/jpeg" if ext in ['jpg', 'jpeg'] else ("image/png" if ext == 'png' else "image/jpeg")
                return f"data:{mime};base64,{encoded}"
        except:
            pass
    return ""


def find_best_image_filename(vals):
    if not available_images:
        return "default.jpg"

    clean_vals = [str(v).strip() for v in vals if pd.notna(v) and str(v).strip() != '']

    possible_names = []
    if len(clean_vals) >= 3:
        possible_names.append(f"{clean_vals[2]}_{clean_vals[1]}_{clean_vals[0]}.JPG")
        possible_names.append(f"{clean_vals[2]}_{clean_vals[1]}_{clean_vals[0]}.jpg")
    if len(clean_vals) >= 2:
        possible_names.append(f"{clean_vals[1]}_{clean_vals[0]}.JPG")
        possible_names.append(f"{clean_vals[1]}_{clean_vals[0]}.jpg")
        possible_names.append(f"{clean_vals[0]}_{clean_vals[1]}.JPG")
        possible_names.append(f"{clean_vals[0]}_{clean_vals[1]}.jpg")

    for name in possible_names:
        if name in available_images:
            return name

    best_img = "default.jpg"
    max_match = 0
    for img in available_images:
        img_lower = img.lower()
        match_count = sum(1 for v in clean_vals if len(v) > 1 and v.lower() in img_lower)
        if match_count > max_match:
            max_match = match_count
            best_img = img

    return best_img if max_match > 0 else "default.jpg"


default_b64 = ""
if available_images:
    default_b64 = get_image_base64(available_images[0])

# -------------------------------------------------------------------------
# [파트 4] 고장 이력 엑셀 읽기
# -------------------------------------------------------------------------
asset_db_dict = {}

if os.path.exists(excel_file_path):
    try:
        df_dash = pd.read_excel(excel_file_path, sheet_name=0, engine='openpyxl')
        equip_cols = df_dash.columns[:4]
        df_dash[equip_cols] = df_dash[equip_cols].ffill()

        date_cols = []
        for col in df_dash.columns:
            col_str = str(col).split(' ')[0]
            if re.search(r'\d{2}\.\d{2}', col_str) or isinstance(col, pd.Timestamp) or re.search(r'202\d', col_str):
                date_cols.append(col)

        parsed_date_cols = []
        for d_col in date_cols:
            col_str = str(d_col).split(' ')[0]
            try:
                if re.match(r'\d{2}\.\d{2}\.\d{2}', col_str):
                    dt = pd.to_datetime(col_str, format='%y.%m.%d')
                else:
                    dt = pd.to_datetime(col_str)
                parsed_date_cols.append((dt, d_col))
            except:
                pass
        parsed_date_cols.sort(key=lambda x: x[0])

        for idx, row in df_dash.iterrows():
            vals = [str(row[col]).strip() for col in equip_cols if
                    pd.notna(row[col]) and str(row[col]).strip() != '' and str(row[col]).strip() != 'nan']
            if not vals:
                continue
            equip_name = " > ".join(vals)

            raw_joined_key = "".join(vals)
            norm_key = re.sub(r'\s+', '', raw_joined_key)

            capacity_val = equip_name
            install_date_val = "공단자체취득"

            if norm_key in bems_meta_map:
                capacity_val = bems_meta_map[norm_key]["capacity"]
                install_date_val = bems_meta_map[norm_key]["install_date"]
            else:
                for b_key, b_data in bems_meta_map.items():
                    if norm_key in b_key or b_key in norm_key:
                        capacity_val = b_data["capacity"]
                        install_date_val = b_data["install_date"]
                        break

            fault_count = 0
            fault_dates = []

            for dt, d_col in parsed_date_cols:
                val = row[d_col]
                if pd.notna(val) and 'fault' in str(val).lower():
                    fault_count += 1
                    fault_dates.append(dt)

            avg_interval = 0.0
            if fault_count > 1:
                intervals = [(fault_dates[i] - fault_dates[i - 1]).days for i in range(1, len(fault_dates))]
                if intervals:
                    avg_interval = round(sum(intervals) / len(intervals), 1)

            real_today = pd.Timestamp.now()

            if fault_count == 0:
                calculated_reliability = 100.0
            else:
                if avg_interval <= 0:
                    avg_interval = 7.0

                last_fault_date = fault_dates[-1]
                days_since_last = max(0, (real_today - last_fault_date).days)
                current_cycle_progress = days_since_last % avg_interval

                drop_ratio = current_cycle_progress / avg_interval
                calculated_reliability = max(0, round(100 - (drop_ratio * 100), 1))

            if calculated_reliability >= 70:
                status_text = "양호 (Normal)"
                status_color = "#28a745"
                lamp_idx = 0
            elif calculated_reliability >= 30:
                status_text = "주의 (Caution)"
                status_color = "#ffc107"
                lamp_idx = 1
            else:
                status_text = "경고 (Warning)"
                status_color = "#fd7e14"
                lamp_idx = 2

            fail_prob = round(100.0 - calculated_reliability, 1)

            img_filename = find_best_image_filename(vals)
            img_b64 = get_image_base64(img_filename)
            if not img_b64: img_b64 = default_b64

            asset_db_dict[equip_name] = {
                "cnt": fault_count,
                "avg": avg_interval,
                "image_data": img_b64,
                "install_date": install_date_val,
                "hi": 70,
                "voltage": "380V",
                "capacity": capacity_val,
                "reliability": calculated_reliability,
                "fail_prob": fail_prob,
                "status_text": status_text,
                "status_color": status_color,
                "lamp_idx": lamp_idx
            }
    except Exception as e:
        pass

asset_db_json = json.dumps(asset_db_dict, ensure_ascii=False)
# -------------------------------------------------------------------------
# [파트 5] HTML 대시보드 렌더링 (🌟 점수판 전체를 왼쪽으로 이동)
# -------------------------------------------------------------------------
html_dashboard_code = '''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>설비 유지보수 대시보드</title>
    <style>
        #spec-capacity { font-size: 14px; font-weight: bold; color: #fff; text-align: right; max-width: 220px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; transition: all 0.3s ease; }
        #spec-capacity:hover { white-space: normal; word-break: break-all; background: rgba(0, 123, 255, 0.2); padding: 4px 8px; border-radius: 4px; }
        body { background: #071321; color: white; font-family: sans-serif; padding: 0px; margin: 0; }
        .main-layout { display: grid; grid-template-columns: 340px 1fr; gap: 20px; }
        .left-column { display: flex; flex-direction: column; gap: 20px; }
        .right-column { display: flex; flex-direction: column; gap: 20px; }
        .card { background: #0f233c; border-radius: 12px; padding: 20px; border: 1px solid #1e3a5f; box-shadow: 0 4px 12px rgba(0,0,0,0.5); }
        select { width: 100%; padding: 12px; background: #071321; color: white; border: 1px solid #4fa0ff; border-radius: 6px; font-size: 15px; cursor: pointer; margin-top: 10px; }
        .asset-spec-card { display: flex; flex-direction: column; gap: 15px; }
        .asset-img-box { width: 100%; height: 160px; border-radius: 8px; border: 1px solid #2a4365; background: #000; overflow: hidden; display: flex; align-items: center; justify-content: center; }
        .asset-img { width: 100%; height: 100%; object-fit: cover; }
        .spec-row { display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px; border-bottom: 1px solid rgba(255,255,255,0.05); padding-bottom: 6px; }
        .spec-label { color: #a0aec0; }
        .spec-val { font-weight: bold; color: #fff; }
        .stats-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; text-align: center; align-items: center; height: 100%; }
        .val { font-size: 2.2em; color: #63b3ed; font-weight: bold; margin-top: 5px; }
        .label { color: #a0aec0; font-size: 1em; }
        .top-right-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .reliability-content { display: grid; grid-template-columns: 1fr 160px; gap: 10px; align-items: center; }

        .gauge-container { display: flex; flex-direction: column; align-items: center; justify-content: center; position: relative; height: 135px; margin-top: 5px; }
        .gauge-visual-box { position: relative; width: 180px; height: 95px; }

        .gauge-num-label { position: absolute; top: -6px; font-size: 14px; font-weight: bold; z-index: 6; text-shadow: 0 0 3px rgba(0,0,0,0.8); }
        #num-30 { left: 32px; color: #ff9500; }
        #num-70 { right: 32px; color: #a8e063; }

        .semi-gauge { 
            width: 180px; height: 90px; 
            border-top-left-radius: 90px; border-top-right-radius: 90px; 
            background: linear-gradient(90deg, #ff3b30 0%, #ff9500 30%, #a8e063 70%, #28a745 100%);
            position: absolute; bottom: 0; left: 0; overflow: hidden; 
        }
        .gauge-cover { 
            width: 130px; height: 65px; background: #0f233c; 
            border-top-left-radius: 65px; border-top-right-radius: 65px; 
            position: absolute; bottom: 0; left: 25px; 
        }

        .gauge-needle {
            position: absolute; bottom: 0px; left: 50%; 
            width: 6px; height: 70px; background-color: #ffffff; 
            transform-origin: bottom center; 
            transform: translateX(-50%) rotate(-90deg);
            transition: transform 0.6s cubic-bezier(0.4, 2.3, 0.6, 1);
            z-index: 5;
            clip-path: polygon(50% 0%, 100% 100%, 0% 100%); 
        }
        .gauge-pivot {
            position: absolute; bottom: -7px; left: 50%; 
            transform: translateX(-50%);
            width: 14px; height: 14px; background-color: #ffffff; border-radius: 50%; 
            border: 2px solid #0f233c; z-index: 6; box-shadow: 0 2px 4px rgba(0,0,0,0.5);
        }

        .tick-line { position: absolute; bottom: 0px; left: 50%; width: 2px; height: 25px; transform-origin: bottom center; z-index: 4; }
        #tick-30 { background: #ff9500; transform: translateX(-50%) rotate(-36deg) translateY(-65px); }
        #tick-70 { background: #a8e063; transform: translateX(-50%) rotate(36deg) translateY(-65px); }

        .gauge-status-label { position: absolute; font-size: 10px; font-weight: bold; white-space: nowrap; z-index: 12; }
        #lbl-left { bottom: 20px; left: -40px; color: #ff3b30; }
        #lbl-top { top: -16px; left: 50%; transform: translateX(-50%); color: #ff9500; }
        #lbl-right { bottom: 20px; right: -40px; color: #34c759; }

        /* 🌟 점수판 전체 그룹의 위치를 강제로 왼쪽으로 당겼습니다! */
        .score-board-wrapper {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-left: -70px; /* 👈 이 숫자를 조절하면 점수판이 통째로 좌우로 움직입니다! (-40px, -50px 등) */
        }

        .gauge-side-score-box {
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            background: rgba(0, 123, 255, 0.12); border: 2px solid rgba(99, 179, 237, 0.5);
            border-radius: 12px; padding: 14px 18px; flex-grow: 1; text-align: center; box-shadow: 0 4px 10px rgba(0,0,0,0.4);
        }
        .gauge-side-val { font-size: 3.2em; font-weight: bold; color: #63b3ed; line-height: 1.0; text-shadow: 0 0 10px rgba(99,179,237,0.4); }

        .lamp-panel { display: flex; flex-direction: column; gap: 8px; align-items: center; border-left: 1px dashed #2a4365; padding-left: 10px; }
        .lamp { width: 14px; height: 14px; border-radius: 50%; background: #2a4365; box-shadow: inset 0 2px 4px rgba(0,0,0,0.5); }
        .gauge-status-bar { margin-top: 15px; text-align: center; font-size: 13px; font-weight: bold; padding: 10px; border-radius: 8px; background: rgba(0,0,0,0.3); border: 1px solid #1e3a5f; }
        .prob-card { background: linear-gradient(135deg, #2b1b42, #1a102f); border: 1px solid #6b46c1; border-radius: 8px; padding: 20px; text-align: center; display: flex; flex-direction: column; justify-content: center; height: 100%; box-sizing: border-box; }
        .prob-val { font-size: 2.5em; color: #ff6b6b; font-weight: bold; margin-top: 10px; text-shadow: 0 0 10px rgba(255,107,107,0.5); }
    </style>
</head>
<body>
    <div class="main-layout">
        <div class="left-column">
            <div class="card">
                <h3 style="margin-top:0; color:#63b3ed; font-size:1.1em; margin-bottom: 15px;">🔍 설비 선택</h3>
                <select id="level1-select" onchange="onLevel1Change()"><option value="">-- 대분류 선택 --</option></select>
                <select id="level2-select" onchange="onLevel2Change()" disabled style="margin-top: 12px;"><option value="">-- 중분류 선택 --</option></select>
                <select id="asset-select" onchange="onAssetSelectChange()" disabled style="margin-top: 12px;"><option value="">-- 세부 설비 선택 --</option></select>
            </div>
            <div class="card asset-spec-card">
                <h3 id="asset-name" style="margin-top:0; margin-bottom:5px; color:#63b3ed; font-size:1.1em; word-break: keep-all;">설비를 선택하세요</h3>
                <div class="asset-img-box"><img id="asset-img" src="" class="asset-img" alt="설비 사진"></div>
                <div>
                    <div class="spec-row"><span class="spec-label">취득일자</span><span id="spec-date" class="spec-val">-</span></div>
                    <div class="spec-row"><span class="spec-label">설비규격</span><span id="spec-capacity" class="spec-val">-</span></div>
                </div>
            </div>
        </div>
        <div class="right-column">
            <div class="card" style="min-height: 85px; display: flex; flex-direction: column; justify-content: center;">
                <h3 style="margin-top:0; margin-bottom:10px; color:#63b3ed; font-size:1.1em;">📊 고장 이력 통계 분석</h3>
                <div class="stats-grid">
                    <div><div class="label">총 고장 횟수</div><div class="val" id="cnt">--</div></div>
                    <div><div class="label">평균 고장 간격</div><div class="val" id="avg">--</div></div>
                </div>
            </div>
            <div class="top-right-grid">
                <div class="card" style="display: flex; flex-direction: column; justify-content: space-between;">
                    <div class="reliability-content">
                        <div>
                            <div class="gauge-container">
                                <div class="gauge-visual-box">
                                    <div class="gauge-status-label" id="lbl-left">즉시정비필요</div>
                                    <div class="gauge-status-label" id="lbl-top">정비요망</div>
                                    <div class="gauge-status-label" id="lbl-right">상태양호</div>
                                    <div class="gauge-num-label" id="num-30">30</div>
                                    <div class="gauge-num-label" id="num-70">70</div>
                                    <div class="semi-gauge">
                                        <div class="tick-line" id="tick-30"></div>
                                        <div class="tick-line" id="tick-70"></div>
                                        <div class="gauge-cover"></div>
                                    </div>
                                    <div class="gauge-needle" id="gauge-needle"></div>
                                    <div class="gauge-pivot"></div>
                                </div>
                            </div>
                            <div style="text-align: center; font-size: 12px; color: #a0aec0; margin-top: 6px;">종합신뢰도지수</div>
                        </div>

                        <!-- 🌟 새로 묶어준 점수판 + 램프 그룹 -->
                        <div class="score-board-wrapper">
                            <div class="gauge-side-score-box">
                                <div class="gauge-side-val" id="reliability-val">100</div>
                                <div style="font-size: 10px; color: #a0aec0; margin-top: 4px; font-weight: bold; letter-spacing: 1px;">SCORE</div>
                            </div>
                            <div class="lamp-panel">
                                <div class="lamp" id="lamp-0"></div><div class="lamp" id="lamp-1"></div>
                                <div class="lamp" id="lamp-2"></div><div class="lamp" id="lamp-3"></div>
                            </div>
                        </div>

                    </div>
                    <div id="status-badge" class="gauge-status-bar">상태 분석 중...</div>
                </div>
                <div class="prob-card">
                    <div style="color: #cbd5e0; font-size: 0.95em;">실시간 설비 고장 발생 확률</div>
                    <div class="prob-val" id="fail-prob">0.0 %</div>
                    <div style="color: #a0aec0; font-size: 0.85em; margin-top: 5px;">신뢰도 지수 역산 AI 추정치</div>
                </div>
            </div>
        </div>
    </div>
    <script>
        const ASSET_DB = __ASSET_DB_JSON__;
        let rawAssetList = []; let hierarchyData = {};
        window.addEventListener('DOMContentLoaded', (event) => {
            if (typeof ASSET_DB === 'undefined') return;
            hierarchyData = {}; rawAssetList = Object.keys(ASSET_DB);
            rawAssetList.forEach(dbKey => {
                let parts = dbKey.split('>').map(p => p.trim());
                let l1 = parts[0] || '기타'; let l2 = parts[1] || '기타';
                if (!hierarchyData[l1]) hierarchyData[l1] = {};
                if (!hierarchyData[l1][l2]) hierarchyData[l1][l2] = [];
                hierarchyData[l1][l2].push(dbKey);
            });
            let l1Select = document.getElementById('level1-select');
            l1Select.innerHTML = '<option value="">-- 대분류 선택 --</option>';
            for (let l1 in hierarchyData) { l1Select.innerHTML += `<option value="${l1}">${l1}</option>`; }
            showInitialDashboardSummary();
        });

        function showInitialDashboardSummary() {
            document.getElementById('asset-name').innerText = "💡 공장 전체 통합 모니터링";
            let imgBox = document.querySelector('.asset-img-box');
            imgBox.style.background = "linear-gradient(135deg, #0f233c, #1a365d)";
            imgBox.innerHTML = `<div style="text-align: center; color: #fff; padding: 15px;">
                <div style="font-size: 15px; font-weight: bold; color: #63b3ed; margin-bottom: 8px;">🏭 스마트 하수처리장 APM</div>
                <div style="font-size: 13px; color: #cbd5e0; line-height: 1.6;">
                    총 관리 설비: <strong style="color: #48bb78;">${rawAssetList.length}개</strong><br>
                    시스템 상태: <strong style="color: #4299e1;">정상 가동 중</strong>
                </div></div>`;
            document.getElementById('spec-date').innerText = "실시간 연동 완료";
            document.getElementById('spec-capacity').innerText = "자동 로드 완료";
            document.getElementById('cnt').innerText = "통합 관리중";
            document.getElementById('avg').innerText = "실시간 연동";
            updateNeedle(100);
        }

        function updateNeedle(score) {
            let needle = document.getElementById('gauge-needle');
            if (!needle) return;
            let degree = (score / 100) * 180 - 90;
            if (degree < -90) degree = -90; if (degree > 90) degree = 90;
            needle.style.transform = `translateX(-50%) rotate(${degree}deg)`;
        }

        function onLevel1Change() {
            let l1 = document.getElementById('level1-select').value;
            let l2Select = document.getElementById('level2-select');
            let assetSelect = document.getElementById('asset-select');
            l2Select.innerHTML = '<option value="">-- 중분류 선택 --</option>';
            assetSelect.innerHTML = '<option value="">-- 세부 설비 선택 --</option>';
            assetSelect.disabled = true;
            if (!l1) { l2Select.disabled = true; showInitialDashboardSummary(); return; }
            l2Select.disabled = false;
            for (let l2 in hierarchyData[l1]) { l2Select.innerHTML += `<option value="${l2}">${l2}</option>`; }
        }

        function onLevel2Change() {
            let l1 = document.getElementById('level1-select').value;
            let l2 = document.getElementById('level2-select').value;
            let assetSelect = document.getElementById('asset-select');
            assetSelect.innerHTML = '<option value="">-- 세부 설비 선택 --</option>';
            if (!l2) { assetSelect.disabled = true; return; }
            assetSelect.disabled = false;
            let items = hierarchyData[l1][l2];
            items.forEach(dbKey => { assetSelect.innerHTML += `<option value="${dbKey}">${dbKey.split('>').pop().trim()}</option>`; });
        }

        function onAssetSelectChange() {
            let selectedKey = document.getElementById('asset-select').value;
            if (!selectedKey) return;
            let data = ASSET_DB[selectedKey]; if (!data) return;
            document.querySelector('.asset-img-box').innerHTML = `<img id="asset-img" src="${data.image_data}" class="asset-img">`;
            document.getElementById('asset-name').innerText = selectedKey;
            document.getElementById('spec-date').innerText = data.install_date || '-';
            document.getElementById('spec-capacity').innerText = data.capacity || '-';
            document.getElementById('cnt').innerText = (data.cnt !== undefined ? data.cnt : '0') + '회';
            document.getElementById('avg').innerText = (data.avg !== undefined ? data.avg : '0') + '일';
            document.getElementById('reliability-val').innerText = data.reliability;
            updateNeedle(data.reliability);
            const statusBadge = document.getElementById('status-badge');
            statusBadge.textContent = "신뢰도 지수 " + data.reliability + " 점 상태 : " + data.status_text;
            statusBadge.style.color = data.status_color;
            document.getElementById('fail-prob').textContent = ((typeof data.fail_prob === 'number') ? data.fail_prob.toFixed(1) : data.fail_prob) + " %";
            for(let i=0; i<4; i++) {
                let lamp = document.getElementById('lamp-' + i);
                lamp.style.background = "#2a4365"; lamp.style.boxShadow = "inset 0 2px 4px rgba(0,0,0,0.5)";
            }
            let activeLamp = document.getElementById('lamp-' + data.lamp_idx);
            if(activeLamp) {
                activeLamp.style.background = data.status_color;
                activeLamp.style.boxShadow = "0 0 10px " + data.status_color + ", inset 0 -2px 4px rgba(0,0,0,0.5)";
            }
        }
    </script>
</body>
</html>
'''

# JSON 데이터 주입
html_dashboard_code = html_dashboard_code.replace("__ASSET_DB_JSON__", asset_db_json)

# 최종 대시보드 렌더링
components.html(html_dashboard_code, height=650, scrolling=True)