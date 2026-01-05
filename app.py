import streamlit as st
import pandas as pd
import base64
import time
from pypdf import PdfReader
from pptx import Presentation
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage

# 페이지 설정 (넓은 화면 사용)
st.set_page_config(page_title="Ultra AI Analyst", layout="wide", page_icon="⚡")

# --- CSS 스타일 (말풍선, 헤더 등 이쁘게 꾸미기) ---
st.markdown("""
<style>
    .stChatMessage {border-radius: 20px; padding: 10px;}
    .stHeader {background-color: transparent;}
</style>
""", unsafe_allow_html=True)

def get_image_base64(file):
    """이미지 파일을 base64로 변환하는 함수"""
    img_bytes = file.getvalue()
    return base64.b64encode(img_bytes).decode('utf-8')

def main():
    # 사이드바: 설정 및 파일 업로드
    with st.sidebar:
        st.header("⚙️ Control Panel")
        # 시연용 API 키 고정 (여기에 키를 넣으면 입력창 사라짐)
        api_key = st.text_input("OpenAI API Key", type="password")
        
        st.divider()
        st.subheader("📂 자료 업로드")
        uploaded_files = st.file_uploader(
            "엑셀, PDF, PPT, 사진을 모두 올려주세요.", 
            accept_multiple_files=True,
            type=['xlsx', 'csv', 'pdf', 'pptx', 'png', 'jpg', 'jpeg']
        )
        
        if st.button("🔄 대화 내용 초기화"):
            st.session_state.messages = []
            st.session_state.context_data = ""
            st.rerun()

    # 메인 타이틀
    st.title("⚡ Ultra Multi-Modal AI Agent")
    st.caption("🚀 엑셀 + PDF + PPT + 이미지 통합 분석 시스템")
    st.divider()

    # --- Session State 초기화 (대화 기억장치) ---
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "context_data" not in st.session_state:
        st.session_state.context_data = ""
    if "processed_files" not in st.session_state:
        st.session_state.processed_files = []

    # 1. 파일 처리 로직 (파일이 올라오면 딱 한 번만 실행)
    if uploaded_files and sorted([f.name for f in uploaded_files]) != sorted(st.session_state.processed_files):
        # (3번 기능) 있어 보이는 로딩 애니메이션
        with st.status("🔍 문서를 스캔하고 데이터를 추출하는 중...", expanded=True) as status:
            
            raw_text = ""
            image_contents = []
            
            for file in uploaded_files:
                ext = file.name.split('.')[-1].lower()
                time.sleep(0.5) # 시연용 딜레이 (너무 빠르면 재미없음)
                
                # 엑셀 처리 + (2번 기능) 데이터 시각화 자동 생성
                if ext in ['xlsx', 'csv']:
                    st.write(f"📊 엑셀 데이터 분석 중: {file.name}")
                    df = pd.read_excel(file) if ext == 'xlsx' else pd.read_csv(file)
                    raw_text += f"\n[Excel Data: {file.name}]\n{df.to_string()}\n"
                    
                    # 엑셀이 있으면 사이드바나 상단에 차트 바로 그려버리기
                    with st.expander(f"📈 {file.name} - 데이터 자동 시각화 (Click to Open)", expanded=False):
                        st.dataframe(df.head())
                        # 숫자 데이터만 뽑아서 차트 그리기
                        numeric_df = df.select_dtypes(include=['float64', 'int64'])
                        if not numeric_df.empty:
                            st.line_chart(numeric_df)
                            st.info("AI가 숫자 데이터를 감지하여 자동으로 트렌드 차트를 생성했습니다.")

                # PDF 처리
                elif ext == 'pdf':
                    st.write(f"📄 PDF 텍스트 추출 중: {file.name}")
                    reader = PdfReader(file)
                    text = "".join([page.extract_text() for page in reader.pages])
                    raw_text += f"\n[PDF Document: {file.name}]\n{text}\n"

                # PPT 처리
                elif ext == 'pptx':
                    st.write(f"📢 프레젠테이션 분석 중: {file.name}")
                    prs = Presentation(file)
                    ppt_text = ""
                    for i, slide in enumerate(prs.slides):
                        txts = [shape.text for shape in slide.shapes if hasattr(shape, "text")]
                        ppt_text += f"Slide {i+1}: {' '.join(txts)}\n"
                    raw_text += f"\n[PPT Slides: {file.name}]\n{ppt_text}\n"

                # 이미지 처리
                elif ext in ['png', 'jpg', 'jpeg']:
                    st.write(f"🖼️ 이미지 비전 인식 중: {file.name}")
                    b64_img = get_image_base64(file)
                    image_contents.append({
                        "type": "image_url",
                        "image_url": {"url": f"data:image/{ext};base64,{b64_img}"}
                    })

            # 텍스트와 이미지 정보를 세션에 저장
            st.session_state.context_data = {"text": raw_text, "images": image_contents}
            st.session_state.processed_files = [f.name for f in uploaded_files]
            
            status.update(label="✅ 모든 문서 분석 완료! AI가 준비되었습니다.", state="complete", expanded=False)

    # 2. (1번 기능) 챗봇 인터페이스 구현
    # 이전 대화 기록 표시
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 사용자 입력 대기
    if prompt := st.chat_input("문서에 대해 무엇이든 물어보세요 (예: 이 엑셀의 매출 추이와 보고서 내용을 비교해줘)"):
        if not api_key:
            st.error("API 키를 먼저 입력해주세요!")
            st.stop()
            
        # 사용자 메시지 표시
        st.chat_message("user").markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})

        # AI 응답 생성
        with st.chat_message("assistant"):
            # (3번 기능) 답변 생성 중 로딩 효과
            message_placeholder = st.empty()
            with st.spinner("AI가 문서 내용을 기반으로 생각 중입니다..."):
                
                try:
                    llm = ChatOpenAI(model="gpt-4o", api_key=api_key, temperature=0.1)
                    
                    # LLM에 보낼 메시지 구성
                    content_payload = []
                    
                    # 1. 텍스트 컨텍스트 추가
                    if st.session_state.context_data.get("text"):
                        content_payload.append({
                            "type": "text", 
                            "text": f"다음은 사용자가 업로드한 문서들의 내용입니다. 이 내용을 바탕으로 질문에 답하세요:\n{st.session_state.context_data['text']}"
                        })
                    
                    # 2. 이미지 컨텍스트 추가
                    if st.session_state.context_data.get("images"):
                        content_payload.extend(st.session_state.context_data['images'])
                        
                    # 3. 사용자 질문 추가
                    content_payload.append({
                        "type": "text",
                        "text": prompt
                    })
                    
                    # 시스템 프롬프트 (페르소나 설정)
                    system_msg = SystemMessage(content="""
                        당신은 탁월한 데이터 분석가이자 비즈니스 컨설턴트입니다. 
                        제공된 엑셀, PDF, PPT, 이미지 자료를 종합적으로 분석하여 통찰력 있는 답변을 주세요.
                        답변할 때는 중요한 숫자에 볼드체를 사용하고, 필요하다면 마크다운 표를 그려서 가독성을 높이세요.
                        한국어로 답변하세요.
                    """)
                    
                    human_msg = HumanMessage(content=content_payload)
                    
                    # 스트리밍 효과 (타자 치듯 나오는 효과)
                    full_response = ""
                    response = llm.stream([system_msg, human_msg])
                    
                    for chunk in response:
                        if chunk.content:
                            full_response += chunk.content
                            message_placeholder.markdown(full_response + "▌")
                    
                    message_placeholder.markdown(full_response)
                    
                    # 대화 기록에 저장
                    st.session_state.messages.append({"role": "assistant", "content": full_response})

                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()