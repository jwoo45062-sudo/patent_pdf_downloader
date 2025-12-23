import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
import fitz  # PyMuPDF
import os
import io
import zipfile

# 페이지 설정
st.set_page_config(page_title="KIPRIS 특허 PDF 다운로더", layout="wide")

st.title("📂 특허 공고전문 PDF 일괄 다운로더")
st.info("엑셀 파일을 업로드하면 1,2 페이지를 추출하여 ZIP 파일로 제공합니다.")

# --- 사이드바: 설정 ---
with st.sidebar:
    st.header("⚙️ 설정")
    service_key = st.text_input("KIPRIS API 서비스키 입력", type="password")
    st.markdown("[KIPRIS Plus](https://plus.kipris.or.kr/)에서 발급받은 키를 입력하세요.")

# --- PDF 처리 함수 ---
def get_pdf_pages(pdf_url, num_pages=2):
    try:
        response = requests.get(pdf_url, timeout=30)
        if response.status_code == 200:
            # 메모리 내에서 PDF 열기
            pdf_stream = io.BytesIO(response.content)
            doc = fitz.open(stream=pdf_stream, filetype="pdf")
            
            # 실제 문서의 페이지 수와 요청한 페이지 수 중 작은 값을 선택
            # (1페이지만 있는 문서에서 2페이지를 추출하려 할 때 에러 방지)
            end_page = min(len(doc), num_pages) - 1
            
            # 새 PDF 생성 및 페이지 복사 (0번부터 end_page까지)
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=0, to_page=end_page)
            
            # 결과물을 바이너리로 변환
            output_buffer = io.BytesIO()
            new_doc.save(output_buffer)
            
            doc.close()
            new_doc.close()
            return output_buffer.getvalue()
    except Exception as e:
        return None
    return None

# --- 메인 로직 ---
uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and service_key:
    df = pd.read_excel(uploaded_file)
    st.write("📋 데이터 미리보기 (총", len(df), "건)")
    st.dataframe(df.head())

    if st.button("🚀 다운로드 시작"):
        zip_buffer = io.BytesIO()
        success_count = 0
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for index, row in df.iterrows():
                # 1. 데이터 추출 및 하이픈 제거
                idx_num = str(row.iloc[0])
                app_num = str(row.iloc[1]).replace("-", "") # 하이픈 제거
                
                file_name = f"{idx_num}_{app_num}.pdf"
                status_text.text(f"처리 중 ({index+1}/{len(df)}): {file_name}")
                
                # 2. API 요청
                api_url = "http://plus.kipris.or.kr/kipo-api/kipi/patUtiModInfoSearchSevice/getAnnFullTextInfoSearch"
                params = {
                    'applicationNumber': app_num,
                    'ServiceKey': service_key
                }

                try:
                    res = requests.get(api_url, params=params)
                    root = ET.fromstring(res.text)
                    pdf_url_node = root.find('.//path')
                    
                    if pdf_url_node is not None:
                        pdf_url = pdf_url_node.text
                        pdf_content = get_pdf_pages(pdf_url, num_pages=2)
                        
                        if pdf_content:
                            # ZIP 파일에 PDF 데이터 추가
                            zip_file.writestr(file_name, pdf_content)
                            success_count += 1
                    
                except Exception as e:
                    st.error(f"에러 발생 ({app_num}): {e}")

                # 진행률 업데이트
                progress_bar.progress((index + 1) / len(df))

            status_text.text("✅ 모든 처리가 완료되었습니다!")

        # 3. ZIP 파일 다운로드 버튼 생성
        if success_count > 0:
            st.success(f"총 {success_count}건의 파일을 압축했습니다.")
            st.download_button(
                label="📦 압축 파일(ZIP) 다운로드",
                data=zip_buffer.getvalue(),
                file_name="patent_pdfs.zip",
                mime="application/zip"
            )
        else:
            st.warning("다운로드된 파일이 없습니다. API 키나 데이터를 확인하세요.")

elif not service_key:
    st.warning("👈 왼쪽 사이드바에 API 키를 입력해 주세요.")
