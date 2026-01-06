import streamlit as st
import pandas as pd
import base64
import time
from io import BytesIO

# 문서 처리용 라이브러리
from pypdf import PdfReader
from pptx import Presentation
from docx import Document 

# AI & 검색용 라이브러리
from langchain_openai import ChatOpenAI
from langchain_community.tools import DuckDuckGoSearchRun
from langchain_core.messages import HumanMessage, SystemMessage

# 1. 페이지 설정
st.set_page_config(page_title="Ultra AI Analyst Pro", layout="wide", page_icon="🕵️‍♂️")

# --- [핵심 수정] 스타일 설정 (다크모드 완벽 대응) ---
st.markdown("""
<style>
    /* 1. 말풍선 및 버튼 디자인 */
    .stChatMessage {border-radius: 15px; padding: 10px;}
    .stButton>button {width: 100%; border-radius: 5px;}
    
    /* 2. 라디오 버튼(모드 선택)이 들어있는 박스 전체를 흰색으로 */
    div[role="radiogroup"] {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #ddd;
    }
    
    /* 3. 라디오 버튼 '안쪽'의 모든 글자(p태그)를 검정색으로 강제 고정 (!important) */
    div[role="radiogroup"] p {
        color: #000000 !important;
        font-weight: bold;
    }
    
    /* 4. 라디오 버튼 선택 시 체크되는 동그라미 색상 조정 (선택 사항) */
    div[role="radiogroup"] div[data-testid="stMarkdownContainer"] {
        color: #000000 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 헬퍼 함수들 ---
def get_image_base64(file):
    """이미지 파일을 base64 문자열로 변환"""
    img_bytes = file.getvalue()
    return base64.b64encode(img_bytes).decode('utf-8')

def create_word_report(messages):
    """대화 내용을 워드 파일(.docx)로 변환"""
    doc = Document()
    doc.add_heading('AI 분석 결과 보고서', 0)
    
    for msg in messages:
        role = "사용자" if msg['role'] == "user" else "AI"
        doc.add_heading(role, level=2)
        doc.add_paragraph(msg['content'])
        doc.add_paragraph("-" * 50)
    
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def main():
    # --- 사이드바: 설정 및 파일 관리 ---
    with st.sidebar:
        st.header("⚙️ Pro Control Panel")
        
        # API 키 관리
        api_key = None
        try:
            if "OPENAI_API_KEY" in st.secrets:
                api_key = st.secrets["OPENAI_API_KEY"]
        except:
            pass
            
        if not api_key:
            api_key = st.text_input("OpenAI API Key", type="password")

        st.divider()

        # --- 페르소나(모드) 선택 ---
        st.subheader("🎭 AI 모드 선택 (Persona)")
        persona_mode = st.radio(
            "분석 관점을 선택하세요:",
            ["1. 친절한 비서 (요약 & 설명)", 
             "2. 깐깐한 감사관 (불일치 & 오류 적발)", 
             "3. 창의적 기획자 (아이디어 제안)"],
            index=0
        )
        
        # 안내 메시지
        if "감사관" in persona_mode:
            st.warning("🚨 [감사 모드] AI가 매우 비판적으로 변합니다.")
        elif "기획자" in persona_mode:
            st.success("💡 [기획 모드] 창의적인 아이디어를 제안합니다.")
        else:
            st.info("😊 [비서 모드] 친절하고 명확하게 설명합니다.")

        st.divider()

        # 파일 업로드
        st.subheader("📂 문서 보관함")
        uploaded_files = st.file_uploader(
            "파일을 추가하면 목록에 쌓입니다.", 
            accept_multiple_files=True,
            type=['xlsx', 'csv', 'pdf', 'pptx', 'png', 'jpg', 'jpeg']
        )

        if "file_cache" not in st.session_state:
            st.session_state.file_cache = {} 
        if "processed_file_names" not in st.session_state:
            st.session_state.processed_file_names = []

        if uploaded_files:
            for file in uploaded_files:
                if file.name not in st.session_state.processed_file_names:
                    with st.spinner(f"📥 새 파일 분석 중... {file.name}"):
                        content = ""
                        images = []
                        ext = file.name.split('.')[-1].lower()
                        
                        try:
                            if ext in ['xlsx', 'csv']:
                                df = pd.read_excel(file) if ext == 'xlsx' else pd.read_csv(file)
                                content = f"[Data: {file.name}]\n{df.to_string()}\n"
                            elif ext == 'pdf':
                                reader = PdfReader(file)
                                content = f"[Doc: {file.name}]\n" + "".join([p.extract_text() for p in reader.pages])
                            elif ext == 'pptx':
                                prs = Presentation(file)
                                txts = []
                                for slide in prs.slides:
                                    txts.extend([s.text for s in slide.shapes if hasattr(s, "text")])
                                content = f"[Slide: {file.name}]\n" + "\n".join(txts)
                            elif ext in ['png', 'jpg', 'jpeg']:
                                b64_img = get_image_base64(file)
                                images.append({
                                    "type": "image_url",
                                    "image_url": {"url": f"data:image/{ext};base64,{b64_img}"}
                                })
                                content = f"[Image File: {file.name}] (이미지 데이터 포함됨)"

                            st.session_state.file_cache[file.name] = {"text": content, "images": images}
                            st.session_state.processed_file_names.append(file.name)
                            
                        except Exception as e:
                            st.error(f"파일 처리 실패 ({file.name}): {e}")

        # 파일 선택 (Context Control)
        st.markdown("👇 **이번 질문에 참고할 파일 선택**")
        if st.session_state.file_cache:
            selected_files = st.multiselect(
                "체크된 파일만 AI가 읽습니다.",
                options=list(st.session_state.file_cache.keys()),
                default=list(st.session_state.file_cache.keys())
            )
        else:
            selected_files = []
            st.caption("업로드된 파일이 없습니다.")

        st.divider()
        
        # 보고서 다운로드
        st.subheader("💾 결과 저장")
        if st.session_state.get("messages"):
            report_file = create_word_report(st.session_state.messages)
            st.download_button(
                label="📝 워드 보고서 다운로드",
                data=report_file,
                file_name="AI_분석_보고서.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        if st.button("🗑️ 대화 내용 지우기"):
            st.session_state.messages = []
            st.rerun()

    # --- 메인 화면 ---
    st.title("🕵️‍♂️ Ultra Analyst Pro")
    st.caption(f"현재 모드: {persona_mode}") 
    
    if "messages" not in st.session_state:
        st.session_state.messages = []

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("질문을 입력하세요 (예: 엑셀과 PDF 내용을 비교해서 틀린 부분 찾아줘)"):
        if not api_key:
            st.warning("왼쪽 사이드바에 OpenAI API 키를 입력해주세요!")
            st.stop()

        st.chat_message("user").markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})

        with st.chat_message("assistant"):
            message_placeholder = st.empty()
            full_response = ""
            
            # 검색 로직
            search_result = ""
            search_keywords = ["검색", "찾아", "조사", "최신", "search", "구글링"]
            
            if any(keyword in prompt for keyword in search_keywords):
                with st.status("🌍 인터넷에서 정보를 찾는 중...", expanded=False) as status:
                    try:
                        search_tool = DuckDuckGoSearchRun()
                        search_result = search_tool.run(prompt)
                        status.update(label="✅ 최신 정보 검색 완료!", state="complete")
                    except Exception as e:
                        status.update(label="⚠️ 검색 실패 (일시적 오류)", state="error")
            
            # 컨텍스트 조립
            context_text = ""
            context_images = []
            
            for fname in selected_files:
                data = st.session_state.file_cache[fname]
                context_text += f"\n--- 문서: {fname} ---\n{data['text']}\n"
                if data['images']:
                    context_images.extend(data['images'])

            if search_result:
                context_text += f"\n\n--- [인터넷 검색 결과] ---\n{search_result}\n"

            try:
                llm = ChatOpenAI(model="gpt-4o", api_key=api_key, temperature=0.1)
                
                content_payload = []
                
                # 페르소나 프롬프트 설정
                if "친절한 비서" in persona_mode:
                    system_instruction = """
                    당신은 친절하고 유능한 비서입니다. 
                    문서의 내용을 이해하기 쉽게 요약하고, 사용자의 질문에 부드러운 톤으로 답변하세요.
                    복잡한 데이터는 표로 정리해주고, 초보자도 알기 쉽게 설명하세요.
                    """
                elif "깐깐한 감사관" in persona_mode:
                    system_instruction = """
                    당신은 냉철한 '품질 관리 감사관(Auditor)'입니다.
                    당신의 목표는 칭찬이 아니라 **'오류와 불일치 발견'**입니다.
                    제공된 문서들(엑셀, PDF 등) 간에 데이터가 다르거나, 계산이 틀린 부분이 있다면
                    반드시 빨간색 글씨나 볼드체(**Bold**)를 사용하여 강력하게 경고하세요.
                    말투는 직설적이고 단호하게 하세요. "확인 부탁드립니다" 대신 "오류입니다"라고 하세요.
                    """
                elif "창의적 기획자" in persona_mode:
                    system_instruction = """
                    당신은 아이디어가 넘치는 '마케팅 기획자'입니다.
                    문서의 데이터를 바탕으로 숨겨진 인사이트를 찾고, 새로운 사업 기회나 개선 아이디어를 제안하세요.
                    이모지(💡, 🚀)를 적극적으로 사용하고, 열정적이고 긍정적인 톤으로 답변하세요.
                    """
                else:
                    system_instruction = "당신은 유능한 AI 분석가입니다."

                final_system_prompt = f"""
                {system_instruction}
                
                [참고 문서 및 데이터]
                {context_text if context_text else "(참고할 문서 없음)"}
                """
                
                content_payload.append(SystemMessage(content=final_system_prompt))
                
                if context_images:
                    content_payload.extend([HumanMessage(content=[img]) for img in context_images])
                
                # 기억력
                for msg in st.session_state.messages[-5:]: 
                    if msg['role'] == 'user':
                        pass 
                    else:
                        content_payload.append(HumanMessage(content=msg['content']))

                content_payload.append(HumanMessage(content=prompt))

                response = llm.invoke(content_payload)
                full_response = response.content
                
                message_placeholder.markdown(full_response)
                st.session_state.messages.append({"role": "assistant", "content": full_response})

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
