import streamlit as st
from utils import (
    generate_outline_and_ppt,
    analyze_ppt_with_llm,
    run_simulation_round,
    generate_eval_report_docx,
    parse_outline_from_text, 
    build_ppt_from_outline_mixed,
    ready_openai, ready_search,
    aoai_chat, search_topk,
    render_rag_section,
    get_judge_roles, default_specialties, make_default_profile,
    judge_greet_and_first_impression, judge_next_turn, judge_score_answer,
    build_brief_from_slides
)

from dotenv import load_dotenv
load_dotenv()

st.set_page_config(page_title="AI 제안서 심사위원 & 시뮬레이터", layout="wide")

def get_outline():
    # 탭2에서 만든 발표용 개요
    return st.session_state.get("sim_outline_brief", "")

def get_criteria():
    # 탭2에서 입력한 평가기준
    return st.session_state.get("criteria_list", [])


tab1, tab2, tab3 = st.tabs([
    "📝 개요 생성 + PPT 만들기",
    "📤 PPT 업로드 평가",
    "🎤 발표 시뮬레이터"
])

# ----------------------------
# 탭1: 개요 생성 + PPT 만들기
# ----------------------------
with tab1:
    st.header("📝 개요 생성 + PPT 만들기")

    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("주제", "AI 기반 펫케어 헬스체크 서비스", placeholder="예: AI 기반 펫케어 헬스체크 서비스")
        summary = st.text_area("핵심내용", "반려동물 사진으로 건강상태를 진단하는 AI 솔루션", placeholder="예: 반려동물 사진으로 건강상태를 진단하는 AI 솔루션")
        audience = st.selectbox("대상", ["투자유치용", "정부사업용", "사내용"])
        tone = st.selectbox("톤", ["간결", "설득", "기술"])
        slide_count = st.slider("슬라이드 수", 10, 25, 15)
        use_rag = st.checkbox("RAG 참고문서 사용(Azure AI Search)")

    with col2:
        deck_title = st.text_input("PPT 제목", "PetCare AI 제안서", placeholder="예: PetCare AI 제안서")
        template_file = st.file_uploader("PPT 템플릿 업로드", type=["pptx"], help="업로드한 템플릿의 레이아웃과 디자인을 적용합니다")
        font_name = st.text_input("폰트 이름", "맑은 고딕", placeholder="예: 맑은 고딕 / Arial")

        # 템플릿 미리보기 정보
        if template_file:
            st.success(f"✅ 템플릿 업로드됨: {template_file.name}")
            st.caption(f"파일 크기: {len(template_file.getvalue()) / 1024:.1f} KB")


    if st.button("🚀 개요 만들고 PPT 생성하기"):
        # 입력 검증
        if not topic.strip():
            st.error("주제를 입력해주세요.")
            st.stop()
            
        if not summary.strip():
            st.error("핵심내용을 입력해주세요.")
            st.stop()

        with st.spinner("개요 생성 중..."):
            # 1) RAG 참고 문서 가져오기(선택)
            ground = ""
            if use_rag and ready_search():
                try:
                    docs = search_topk(topic, 3)
                    ground = "\n\n".join([f"- {d['title']}: {d['content']}" for d in docs])[:3000]
                    st.info(f"참고문서 {len(docs)}개를 찾았습니다.")
                except Exception as e:
                    st.warning(f"RAG 참조 실패(무시 후 진행): {e}")

            # 2) LLM에 마크다운 개요 요청
            sys = "You are a concise Korean pitch-deck planner."
            usr = f"""
        다음 정보를 반영해 **마크다운 개요**를 만들어줘. 
        각 슬라이드는 '## 슬라이드 N: 제목' 형식으로, 아래에 불릿을 '-'로 3~5개 작성.
        추가 설명도 붙여줘.

        주제: {topic}
        핵심내용: {summary}
        타겟: {audience}
        톤: {tone}
        슬라이드 수: {slide_count}

        참고 자료(있으면 반영):
        {ground}
        """
            try:
                outline_md = aoai_chat(
                    [{"role": "system", "content": sys}, {"role": "user", "content": usr}],
                    max_tokens=1200, temperature=0.6
                )
                slides = parse_outline_from_text(outline_md)

                if not slides:
                    st.error("개요 생성에 실패했습니다. 다시 시도해주세요.")
                    st.stop()

                st.session_state["outline"] = {"slides": slides, "target": audience}
                st.session_state["last_outline_md"] = outline_md

                st.success(f"✅ 개요 생성 완료! ({len(slides)}개 슬라이드)")
            except Exception as e:
                st.error(f"개요 생성 실패: {e}")
                st.stop()
        
        with st.spinner("PPT 생성 중..."):
            try:
                # 3) PPT 생성 (템플릿 적용)
                template_data = None
                if template_file:
                    # 파일 포인터가 끝에 있을 수 있으므로 처음으로 리셋
                    template_file.seek(0)
                    template_data = template_file.getvalue()
                    st.info("템플릿을 적용하여 PPT를 생성합니다...")
                else:
                    st.info("기본 템플릿으로 PPT를 생성합니다...")

                ppt_io = build_ppt_from_outline_mixed(
                    outline_slides=slides,
                    project_title=deck_title,
                    template_bytes=template_data,
                    font_name=font_name
                )
                
                if ppt_io:
                    st.session_state["ppt_bytes"] = ppt_io.getvalue()
                    st.success("✅ PPT 생성 완료!")
                else:
                    st.error("PPT 생성에 실패했습니다.")
                    
            except Exception as e:
                st.error(f"PPT 생성 실패: {e}")
                # 디버깅용 상세 정보
                st.expander("오류 세부사항").write(str(e))

        # ---------------- Preview & Download ----------------
        if st.session_state.get("outline"):
            st.markdown("### 📑 생성된 개요 미리보기")
            with st.expander("개요 상세보기", expanded=False):
                for i, s in enumerate(st.session_state["outline"]["slides"], 1):
                    st.markdown(f"**{i}. {s['title']}**")
                    for b in s["bullets"]:
                        st.markdown(f"- {b}")

        if st.session_state.get("ppt_bytes"):
            st.markdown("### 📥 PPT 다운로드")
            file_name = f"{deck_title.replace(' ', '_')}_outline.pptx"
            st.download_button(
                "📄 다운로드: PPT 파일",
                data=st.session_state["ppt_bytes"],
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            st.success("위 버튼을 클릭하여 생성된 PPT를 다운로드하세요!")
        
            # 통계 정보
            ppt_size = len(st.session_state["ppt_bytes"])
            slide_count_actual = len(st.session_state["outline"]["slides"])
            st.caption(f"생성된 PPT: {slide_count_actual}개 슬라이드, {ppt_size/1024:.1f} KB")

# ----------------------------
# 탭2: PPT 업로드 평가
# ----------------------------
with tab2:
    st.header("📤 PPT 업로드 평가 (LLM)")

    uploaded_ppt = st.file_uploader("PPT 파일 업로드", type=["pptx"], key="eval_upload")

    st.markdown("### 1) 평가 기준 입력")
    st.write("기준명 + 가중치(%) + 루브릭(LLM에게 설명할 평가 기준)을 직접 입력하세요.")

    # 기본 4개 기준 예시 보여주고 편집 가능하게
    default_rows = [
        {"crit": "시장분석", "weight": 30, "rubric": "TAM·SAM·SOM 수치와 근거"},
        {"crit": "비즈니스모델", "weight": 25, "rubric": "수익모델·CAC→LTV 수치 제시"},
        {"crit": "기술구체성", "weight": 25, "rubric": "시스템 아키텍처·성능지표"},
        {"crit": "재무현실성", "weight": 20, "rubric": "3년 재무추정·민감도 분석"},
    ]

    criteria = []
    for i, row in enumerate(default_rows):
        c1, c2, c3 = st.columns([1, 1, 3])
        with c1:
            crit = st.text_input(f"기준 {i+1} 이름", row["crit"], placeholder=row["crit"])
        with c2:
            weight = st.number_input(f"가중치 {i+1}(%)", 0, 100, row["weight"])
        with c3:
            rubric = st.text_input(f"루브릭 {i+1}", row["rubric"], placeholder=row["rubric"])
        criteria.append({"name": crit, "weight": weight, "rubric": rubric})
    
    # 🔹 RAG 옵션
    st.markdown("### 2) 참고 문헌(RAG) 옵션")
    use_rag = st.checkbox("참고 문헌 기반으로 검증하기 (Azure Search 필요)")
    rag_top_k = st.slider("문헌 Top-K", 1, 5, 2, disabled=not use_rag)

    if st.button("🔎 PPT 분석 & 평가"):
        if uploaded_ppt:
            with st.spinner("PPT 분석 중..."):
                # 분석 호출
                scores, feedback, rich = analyze_ppt_with_llm(uploaded_ppt, criteria, use_rag, rag_top_k)

                st.success("분석 완료!")

                col_left, col_right = st.columns([3, 1])  # 오른쪽 칼럼은 남겨둬도 되고 안 써도 됨
                with col_left:
                    st.markdown("### 📊 내용 평가")
                    ce = rich.get("content_evaluation", {}) if rich else {}
                    st.write("- **가중 총점**:", ce.get("weighted_total"))
                    st.write("**기준별 점수**")
                    st.json(scores)

                    st.markdown("### 🧾 총평")
                    st.write(rich.get("summary", "총평 없음"))

                    st.markdown("### 💡 핵심 개선 포인트")
                    for fb in (feedback or []):
                        st.write(f"- {fb}")

                # 구조 분석
                with st.expander("🔎 구조 분석 (슬라이드 정리 / 논리 흐름 / 빠진 내용)"):
                    struct = rich.get("structure", {}) if rich else {}
                    st.markdown("**슬라이드별 핵심 정리**")
                    for s in struct.get("slides_outline", []):
                        st.markdown(f"- **{s.get('idx','?')}. {s.get('title','')}:** " +
                                    " / ".join((s.get('key_points') or [])[:5]))
                    st.markdown("**논리 흐름 이슈**")
                    for iss in struct.get("logic_flow", {}).get("issues", []):
                        st.write(f"- {iss}")
                    st.markdown("**빠진 내용**")
                    for miss in struct.get("missing", []):
                        st.write(f"- {miss}")

                # 맞춤법 & 문장 체크
                with st.expander("✨ 맞춤법 · 문장 체크 (오타/띄어쓰기/어색한 문장/용어 일관성/간결화)"):
                    wc = rich.get("writing_check", {}) if rich else {}
                    typos = wc.get("typos", []) or []
                    awkward = wc.get("awkward", []) or []
                    terms = wc.get("terminology", []) or []
                    concise = wc.get("concise_rewrites", []) or []

                    st.markdown("**오타/띄어쓰기**")
                    printed = False
                    for t in typos:
                        if isinstance(t, dict):
                            st.write(f"- `{t.get('before','')}` → **{t.get('after','')}** ({t.get('why','')})")
                            printed = True
                        elif isinstance(t, str):
                            st.write(f"- `{t}`")
                            printed = True
                    if not printed:
                        st.caption("· 발견 없음")

                    st.markdown("**어색한 문장 → 제안**")
                    printed = False
                    for a in awkward:
                        if isinstance(a, dict):
                            st.write(f"- `{a.get('before','')}` → **{a.get('suggest','')}** ({a.get('reason','')})")
                            printed = True
                        elif isinstance(a, str):
                            st.write(f"- `{a}`")
                            printed = True
                    if not printed:
                        st.caption("· 발견 없음")

                    st.markdown("**용어 일관성**")
                    printed = False
                    for tm in terms:
                        if isinstance(tm, dict):
                            st.write(f"- **{tm.get('term','')}**: {tm.get('note','')}")
                            printed = True
                        elif isinstance(tm, str):
                            st.write(f"- **{tm}**")
                            printed = True
                    if not printed:
                        st.caption("· 발견 없음")

                    st.markdown("**PPT답게 간결화**")
                    printed = False
                    for rw in concise:
                        if isinstance(rw, dict):
                            st.write(f"- (Slide {rw.get('slide_idx','?')}) `{rw.get('line_before','')}` → **{rw.get('line_after','')}**")
                            printed = True
                        elif isinstance(rw, str):
                            st.write(f"- {rw}")
                            printed = True
                    if not printed:
                        st.caption("· 제안 없음")

                # RAG 근거 보여주기
                if use_rag:
                    with st.expander("📚 RAG 참고 문헌"):
                        slides_for_llm = (rich or {}).get("slides_for_llm", [])
                        rag_map = (rich or {}).get("rag_references", {})
                        if slides_for_llm and rag_map:
                            render_rag_section(st, slides_for_llm, rag_map)
                        else:
                            st.caption("참고 문헌 없음")

                # DOCX 리포트 다운로드 (업데이트된 섹션 포함)
                # DOCX 리포트 다운로드 (RAG/구조/맞춤법 포함)
                docx_io = generate_eval_report_docx(uploaded_ppt, criteria, rich)
                st.download_button(
                    "📥 평가 리포트 다운로드 (DOCX)",
                    data=docx_io.getvalue(),
                    file_name="ppt_evaluation_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                # 분석 호출 후 rich 결과가 있을 때
                slides_for_llm = (rich or {}).get("slides_for_llm", [])
                brief = build_brief_from_slides(slides_for_llm)

                # 요약이 너무 빈약하면 LLM 총평으로 보강
                if (not brief) and rich.get("summary"):
                    brief = rich["summary"]

                # ✅ 탭3에서 쓸 세션 저장
                st.session_state["sim_outline_brief"] = brief
                st.session_state["criteria_list"] = criteria

                # 구조 요약을 탭3에서 쓸 개요 텍스트로 저장
                slides_outline = (rich or {}).get("structure", {}).get("slides_outline", [])
                sim_outline_brief = "\n".join(
                    [f"{s.get('idx','?')}. {s.get('title','')}" for s in slides_outline]
                ) or "제안 개요 없음"
                st.session_state["sim_outline_brief"] = sim_outline_brief

                st.markdown("### 🎤 발표 시뮬레이션 안내")
                st.info("👉 발표 시뮬레이션을 진행하려면 **3번 탭**으로 이동하세요!")
        else:
            st.warning("먼저 PPT 파일을 업로드하세요.")

# ----------------------------
# 탭3: 발표 시뮬레이션
# ----------------------------

def render_tab3_judges(get_outline=None, get_criteria=None):
    # 세션 키 보장
    if "judges" not in st.session_state:
        st.session_state["judges"] = []   # [{profile, chats:[{role,content}], scores:[...], progress:int}]
    if "active_judge_idx" not in st.session_state:
        st.session_state["active_judge_idx"] = 0

    st.title("🧑‍⚖️ AI 심사위원 커스터마이징 & 시뮬레이션")

    # ===== 발표 개요/기준 준비 =====
    # outline_text: getter → 세션 → 기본값
    outline_text = ""
    if get_outline:
        try:
            outline_obj = get_outline()
            if isinstance(outline_obj, str):
                outline_text = outline_obj.strip()
            elif isinstance(outline_obj, dict) and "slides" in outline_obj:
                outline_text = "\n".join(
                    [f"{i+1}. {s.get('title','')}" for i, s in enumerate(outline_obj["slides"])]
                )
        except Exception:
            pass
    if not outline_text:
        outline_text = (
            st.session_state.get("sim_outline_brief")
            or st.session_state.get("last_outline_md")
            or ""
        )

    # criteria_list: getter → 세션 → 빈 리스트
    criteria_list = []
    if get_criteria:
        try:
            criteria_list = get_criteria() or []
        except Exception:
            criteria_list = []
    if not criteria_list:
        criteria_list = st.session_state.get("criteria_list", [])

    # ===== 우측 대화 영역 헤더 =====
    st.markdown("#### 💬 실시간 시뮬레이션")
    if not outline_text:
        st.warning("⚠️ 발표 시뮬레이션을 시작하려면 먼저 **2번 탭에서 PPT 분석**을 완료하세요.")
        return
    else:
        st.success("✅ PPT 분석이 완료되었습니다. 발표 시뮬레이션을 시작할 수 있습니다.")

    # 심사위원 존재 확인
    if not st.session_state["judges"]:
        st.info("우측(사이드바)에서 심사위원을 최소 1명 추가하세요.")
        return

    # 활성 심사위원 선택
    idx = st.session_state["active_judge_idx"]
    idx = max(0, min(idx, len(st.session_state["judges"]) - 1))
    st.session_state["active_judge_idx"] = idx

    judge = st.session_state["judges"][idx]
    prof  = judge["profile"]
    judge.setdefault("chats", [])
    judge.setdefault("scores", [])
    judge.setdefault("progress", 0)

    # 상단 카드: 프로필 요약 + 진행률
    colA, colB, colC = st.columns([2,1,1])
    with colA:
        st.markdown(f"**{prof.get('name') or prof.get('role','심사위원')}** · {prof.get('role','심사위원')}")
        st.caption(
            f"스타일: {prof.get('style_carefulness','보통')} / "
            f"{prof.get('style_question','직설적')} / "
            f"{prof.get('style_focus','큰그림')} / "
            f"{prof.get('style_tone','공손')}"
        )
        if prof.get("specialties"):
            st.caption("전문분야: " + ", ".join(prof["specialties"]))
    with colB:
        st.metric("진행 질문 수", judge.get("progress", 0))
    with colC:
        last_score = judge["scores"][-1]["weighted_total"] if judge["scores"] else "-"
        st.metric("최근 가중 총점", last_score)

    st.divider()

    # === 시작 버튼 ===
    if not judge["chats"]:
        if st.button("👋 인사 및 첫인상 받기"):
            greet = judge_greet_and_first_impression(prof, outline_text or "제안 개요 없음")
            judge["chats"].append({"role":"assistant", "content": greet})
            st.session_state["judges"][idx] = judge
            try:
                st.rerun()
            except Exception:
                st.experimental_rerun()

    # === 대화 로그 표시 ===
    for msg in judge["chats"]:
        st.chat_message("assistant" if msg["role"]=="assistant" else "user").markdown(msg["content"])

    # === 채팅 입력 → 즉시 평가 → 다음 질문 ===
    if judge["chats"]:  # 시작 이후에만 입력 가능
        user_text = st.chat_input("심사위원에게 답변을 입력하세요...")
        if user_text:
            # 1) 사용자 답 저장
            judge["chats"].append({"role":"user","content":user_text})

            # 2) 즉시 평가(JSON)
            try:
                eval_res = judge_score_answer(prof, outline_text or "제안 개요 없음", criteria_list, judge["chats"], user_text)
            except Exception as e:
                eval_res = {"scores": {}, "weighted_total": 0, "comments": [f"평가 실패: {e}"]}
            judge["scores"].append(eval_res)

            # 3) 다음 질문 생성(자연어)
            nxt = judge_next_turn(prof, outline_text or "제안 개요 없음", judge["chats"])
            judge["chats"].append({"role":"assistant","content":nxt})

            # 진행 질문 수 +1
            judge["progress"] = judge.get("progress", 0) + 1

            # 저장 & 리렌더
            st.session_state["judges"][idx] = judge
            try:
                st.rerun()
            except Exception:
                st.experimental_rerun()

    # === 이번 턴 미니 평가 요약 ===
    st.divider()
    st.markdown("#### 📝 이번 턴 평가(요약)")
    last_eval = (judge.get("scores") or [])[-1] if (judge.get("scores")) else None
    if last_eval:
        c1, c2, c3 = st.columns([1,2,2])
        with c1:
            st.metric("가중 총점", last_eval.get("weighted_total", 0))
        with c2:
            sdict = last_eval.get("scores", {})
            top2 = sorted(sdict.items(), key=lambda x: x[1], reverse=True)[:2]
            if top2:
                st.caption("상위 기준")
                for k, v in top2:
                    st.write(f"- {k}: {v}")
        with c3:
            cmts = last_eval.get("comments", [])[:2]
            if cmts:
                st.caption("코멘트")
                for c in cmts:
                    st.write(f"- {c}")
    else:
        st.caption("아직 평가 데이터가 없습니다. 답변을 입력하면 자동 평가됩니다.")

def render_judges_panel(criteria_getter):
    roles = get_judge_roles()
    specs = default_specialties()

    st.markdown("### ⚙️ 심사위원 구성")
    with st.expander("➕ 심사위원 추가", expanded=True):
        new_prof = make_default_profile()
        new_prof["name"] = st.text_input("이름(표시용)", key="nj_name", placeholder="예: 박VC")
        new_prof["role"] = st.selectbox("역할", roles, index=0, key="nj_role")

        col1, col2 = st.columns(2)
        with col1:
            new_prof["style_carefulness"] = st.select_slider("까다로움 정도", ["온화함","보통","매우 까다로움"], value="보통", key="nj_care")
            new_prof["style_question"] = st.select_slider("질문 스타일", ["논리적","직관적","감성적"], value="논리적", key="nj_q")
        with col2:
            new_prof["style_focus"] = st.select_slider("관심 영역", ["디테일 중시","큰 그림 중시"], value="큰 그림 중시", key="nj_focus")
            new_prof["style_tone"] = st.select_slider("의사소통", ["직설적","우회적","격려형"], value="직설적", key="nj_tone")

        new_prof["specialties"] = st.multiselect(
            "전문 분야(다중선택)", sum(specs.values(), []), default=[]
        )

        new_prof["career_years"] = st.select_slider("경력 연차", options=["3년 미만", "3~7년", "7~15년", "15년 이상"], key="nj_years")
        new_prof["company_size"] = st.radio("회사 규모", ["스타트업", "중견기업", "대기업", "공공기관"], horizontal=True, key="nj_company")

        st.markdown("**페르소나 설명 (500자 이내)**")
        new_prof["persona_text"] = st.text_area("자유 텍스트", height=90, key="nj_persona")

        crits = criteria_getter() if criteria_getter else []
        if crits:
            st.markdown("#### 평가 우선순위")
            for c in crits:
                opt = st.radio(f"{c['name']}", ["1순위", "2순위", "3순위", "4순위"], horizontal=True, key=f"prio_{c['name']}")
                if opt == "1순위":
                    new_prof.setdefault("priority_1", []).append(c['name'])
                elif opt == "2순위":
                    new_prof.setdefault("priority_2", []).append(c['name'])
                elif opt == "3순위":
                    new_prof.setdefault("priority_3", []).append(c['name'])
                elif opt == "4순위":
                    new_prof.setdefault("priority_4", []).append(c['name'])

        if st.button("추가", use_container_width=True):
            if not new_prof["name"]:
                st.warning("이름을 입력해 주세요.")
            elif len(st.session_state.get("judges", [])) >= 5:
                st.warning("최대 5명까지 추가할 수 있어요.")
            else:
                st.session_state.setdefault("judges", []).append({
                    "profile": new_prof, "chats": [], "scores": [], "progress": 0
                })
                st.success(f"심사위원 '{new_prof['name']}' 추가됨")

    if st.session_state.get("judges"):
        st.markdown("---")
        st.markdown("### 현재 심사위원")
        for i, j in enumerate(st.session_state["judges"]):
            prof = j["profile"]
            sel = st.radio(
                "활성 심사위원 선택",
                options=list(range(len(st.session_state["judges"]))),
                format_func=lambda idx: f"{st.session_state['judges'][idx]['profile'].get('name') or '이름없음'} · {st.session_state['judges'][idx]['profile'].get('role')}",
                index=st.session_state.get("active_judge_idx", 0),
                key=f"judge_pick_{i}"
            )
            st.session_state["active_judge_idx"] = sel
            if st.button(f"삭제: {prof.get('name') or prof.get('role')}", key=f"del_{i}"):
                st.session_state["judges"].pop(i)
                st.success("삭제 완료")
                st.rerun()

with tab3:
     # 우측 슬림 패널 컬럼
    col_main, col_side = st.columns([3, 1])

    with col_side:
        render_judges_panel(get_criteria)  # ← 아래 함수로 교체

    with col_main:
        render_tab3_judges(get_outline, get_criteria)  # 기존 대화/평가 본문
