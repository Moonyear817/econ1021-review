from pathlib import Path
import random
import re

import streamlit as st

from econ_quiz.data_pipeline import load_or_build_dataset, sample_final_questions
from econ_quiz.progress_store import load_progress, record_attempt, save_progress

st.set_page_config(page_title="经济学 Lecture 刷题系统", page_icon="📘", layout="wide")

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / ".data"
DATASET_PATH = DATA_DIR / "quiz_dataset.json"
PROGRESS_PATH = DATA_DIR / "progress.json"

DEFAULT_PPT_DIR = Path("/Users/yihao/Desktop/econ/课件")
DEFAULT_QUESTION_DIR = Path("/Users/yihao/Desktop/econ/题库")
DEFAULT_SOLUTION_DIR = Path("/Users/yihao/Desktop/econ/(Solution Manual)Economics, 6th Global Edition by Hubbard")


def to_question_map(dataset: dict) -> dict:
    return {q["qid"]: q for q in dataset["questions"]}


def clean_stem_text(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r"\bDiff:\s*[^\n\r]*", "", text, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bPage\s*Ref:\s*[^\n\r]*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bAACSB:\s*[^\n\r]*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\*:\s*Recurring", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return cleaned.strip()


def render_bilingual_explanation(question: dict) -> None:
    zh = question.get("explanation_zh", "")
    en = question.get("explanation_en", "")
    if zh:
        st.markdown("#### 中文解析")
        st.write(zh)
    if en:
        st.markdown("#### English Explanation")
        st.write(en)


def render_question_block(question: dict, mode: str, progress: dict) -> tuple[dict, bool]:
    st.markdown(f"### 题目 {question['qid']}")
    st.write(clean_stem_text(question["stem"]))

    option_keys = [k for k in ["A", "B", "C", "D"] if k in question["options"]]
    labels = [f"{k}. {question['options'][k]}" for k in option_keys]
    choice = st.radio(
        "请选择答案",
        labels,
        index=None,
        key=f"choice_{mode}_{question['qid']}",
    )

    submitted = st.button("提交本题", key=f"submit_{mode}_{question['qid']}")
    if not submitted:
        return progress, False

    if not choice:
        st.warning("请先选择一个选项再提交。")
        return progress, False

    selected = choice.split(".", 1)[0].strip()
    correct = selected == question["answer"]

    progress = record_attempt(
        progress,
        qid=question["qid"],
        selected=selected,
        correct_answer=question["answer"],
        stem=question["stem"],
        options=question["options"],
        explanation_zh=question.get("explanation_zh", ""),
        explanation_en=question.get("explanation_en", ""),
        lecture_title=question["lecture_title"],
        mode=mode,
    )
    save_progress(PROGRESS_PATH, progress)

    if correct:
        st.success(f"回答正确，答案是 {question['answer']}。")
    else:
        st.error(f"回答错误。你的答案是 {selected}，正确答案是 {question['answer']}。")

    st.info(f"Topic: {question.get('topic', 'N/A')}")
    st.markdown("#### 解析")
    render_bilingual_explanation(question)
    return progress, True


def render_practice_mode(dataset: dict, progress: dict, lecture_id: str) -> dict:
    lectures = {lec["lecture_id"]: lec for lec in dataset["lectures"]}
    qmap = to_question_map(dataset)

    lecture = lectures[lecture_id]
    qids = lecture["question_ids"]
    if not qids:
        st.warning("这个分类下暂时没有题目。")
        return progress

    st.markdown(f"## {lecture['lecture_title']}")
    st.caption(f"题量：{len(qids)}")

    order_key = f"order_{lecture_id}"
    if order_key not in st.session_state:
        shuffled = qids.copy()
        random.shuffle(shuffled)
        st.session_state[order_key] = shuffled

    if st.button("重新随机本 Lecture 题序", key=f"shuffle_{lecture_id}"):
        shuffled = qids.copy()
        random.shuffle(shuffled)
        st.session_state[order_key] = shuffled
        st.session_state[f"idx_{lecture_id}"] = 0
        st.rerun()

    display_qids = st.session_state[order_key]

    idx_key = f"idx_{lecture_id}"
    if idx_key not in st.session_state:
        st.session_state[idx_key] = 0

    current_idx = st.session_state[idx_key]
    current_idx = max(0, min(current_idx, len(display_qids) - 1))
    st.session_state[idx_key] = current_idx

    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        if st.button("上一题", disabled=current_idx == 0, key=f"prev_{lecture_id}"):
            st.session_state[idx_key] = max(0, current_idx - 1)
            st.rerun()
    with c2:
        st.write(f"第 {current_idx + 1} / {len(display_qids)} 题")
    with c3:
        if st.button("下一题", disabled=current_idx >= len(display_qids) - 1, key=f"next_{lecture_id}"):
            st.session_state[idx_key] = min(len(display_qids) - 1, current_idx + 1)
            st.rerun()

    q = qmap[display_qids[st.session_state[idx_key]]]
    progress, _ = render_question_block(q, mode=f"practice_{lecture_id}", progress=progress)
    return progress


def render_final_mode(dataset: dict, progress: dict) -> dict:
    qmap = to_question_map(dataset)

    st.markdown("## Final 抽测模式（微观随机 30 题）")

    if "final_state" not in st.session_state:
        st.session_state.final_state = {
            "started": False,
            "qids": [],
            "idx": 0,
            "answers": {},
            "submitted": False,
        }

    state = st.session_state.final_state

    if not state["started"]:
        if st.button("开始 Final 抽测"):
            state["started"] = True
            state["submitted"] = False
            state["qids"] = sample_final_questions(dataset, n=30)
            state["idx"] = 0
            state["answers"] = {}
            st.rerun()
        return progress

    if state["submitted"]:
        qids = state["qids"]
        answers = state["answers"]
        total = len(qids)
        correct_count = 0
        wrong_items = []
        for qid in qids:
            q = qmap[qid]
            user_ans = answers.get(qid, "")
            if user_ans == q["answer"]:
                correct_count += 1
            else:
                wrong_items.append((q, user_ans))

        score = 100 * correct_count / total if total else 0
        st.success(f"考试完成。得分：{correct_count}/{total} ({score:.1f}%)")

        if wrong_items:
            st.markdown("### 错题解析")
            for q, ua in wrong_items:
                st.markdown(f"#### {q['qid']}")
                st.write(clean_stem_text(q["stem"]))
                st.write(f"你的答案：{ua if ua else '未作答'}")
                st.write(f"正确答案：{q['answer']}")
                render_bilingual_explanation(q)
        else:
            st.info("本次抽测全对，暂无错题。")

        if st.button("重新开始 Final 抽测"):
            st.session_state.final_state = {
                "started": False,
                "qids": [],
                "idx": 0,
                "answers": {},
                "submitted": False,
            }
            st.rerun()

        return progress

    qids = state["qids"]
    idx = state["idx"]

    idx = max(0, min(idx, len(qids) - 1))
    state["idx"] = idx
    answered_count = len(state["answers"])
    st.caption(f"已完成 {answered_count}/{len(qids)} 题")

    left, right = st.columns([3, 2])

    q = qmap[qids[idx]]
    option_keys = [k for k in ["A", "B", "C", "D"] if k in q["options"]]
    labels = [f"{k}. {q['options'][k]}" for k in option_keys]

    existing_choice = state["answers"].get(q["qid"])
    selected_index = None
    if existing_choice in option_keys:
        selected_index = option_keys.index(existing_choice)

    with left:
        st.write(f"第 {idx + 1} / {len(qids)} 题")
        st.write(clean_stem_text(q["stem"]))
        sel = st.radio(
            "请选择答案",
            labels,
            index=selected_index,
            key=f"final_choice_{q['qid']}_{idx}",
        )

        def save_current_selection() -> None:
            chosen = sel.split(".", 1)[0].strip() if sel else ""
            if chosen:
                state["answers"][q["qid"]] = chosen

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("上一题", disabled=idx == 0, key=f"final_prev_{idx}"):
                save_current_selection()
                state["idx"] = max(0, idx - 1)
                st.rerun()
        with c2:
            if st.button("保存本题答案", key=f"final_save_{idx}"):
                save_current_selection()
                st.success("已保存。")
        with c3:
            if st.button("下一题", disabled=idx >= len(qids) - 1, key=f"final_nav_next_{idx}"):
                save_current_selection()
                state["idx"] = min(len(qids) - 1, idx + 1)
                st.rerun()

        if st.button("提交 Final 并统一批改", key="final_submit_all"):
            save_current_selection()
            for qid in qids:
                q_item = qmap[qid]
                chosen = state["answers"].get(qid, "")
                progress = record_attempt(
                    progress,
                    qid=qid,
                    selected=chosen,
                    correct_answer=q_item["answer"],
                    stem=q_item["stem"],
                    options=q_item["options"],
                    explanation_zh=q_item.get("explanation_zh", ""),
                    explanation_en=q_item.get("explanation_en", ""),
                    lecture_title="Final 抽测",
                    mode="final",
                )
            save_progress(PROGRESS_PATH, progress)
            state["submitted"] = True
            st.rerun()

    with right:
        st.markdown("### 30 题完成情况")
        for row_start in range(0, len(qids), 5):
            cols = st.columns(5)
            for offset in range(5):
                i = row_start + offset
                if i >= len(qids):
                    continue
                qid = qids[i]
                done = qid in state["answers"]
                current = i == idx
                label = f"{i + 1}{'✓' if done else ''}{'•' if current else ''}"
                if cols[offset].button(label, key=f"jump_final_{i}"):
                    save_current_selection()
                    state["idx"] = i
                    st.rerun()

    return progress


def render_wrong_book(progress: dict, dataset: dict) -> dict:
    wrong_book = progress.get("wrong_book", {})
    qmap = to_question_map(dataset)
    st.markdown("## 错题本")
    st.caption("答错自动加入；后续做对会自动移除。")

    if not wrong_book:
        st.info("错题本为空。")
        return progress

    st.write(f"当前错题数量：{len(wrong_book)}")
    for qid, item in sorted(list(wrong_book.items())):
        st.markdown(f"### {qid}")
        st.write(clean_stem_text(item.get("stem", "")))
        options = item.get("options", {})
        for key in ["A", "B", "C", "D"]:
            if key in options:
                st.write(f"{key}. {options[key]}")
        st.write(f"你的答案：{item.get('your_answer', '')}")
        st.write(f"正确答案：{item.get('correct_answer', '')}")
        if st.button("从错题本删除", key=f"delete_wrong_{qid}"):
            wrong_book.pop(qid, None)
            progress["wrong_book"] = wrong_book
            save_progress(PROGRESS_PATH, progress)
            st.rerun()
        q = qmap.get(qid, {})
        zh_exp = item.get("explanation_zh") or q.get("explanation_zh", "")
        en_exp = item.get("explanation_en") or q.get("explanation_en", "")
        legacy_exp = item.get("explanation", "")

        if zh_exp:
            st.markdown("#### 中文解析")
            st.write(zh_exp)
        if en_exp:
            st.markdown("#### English Explanation")
            st.write(en_exp)
        if (not zh_exp and not en_exp) and legacy_exp:
            st.markdown("#### 解析")
            st.write(legacy_exp)

    return progress


def main() -> None:
    st.title("经济学 Lecture 刷题系统")
    st.caption("按课件自动分类题目，支持 Final 抽测与错题本")

    with st.sidebar:
        st.markdown("## 数据源设置")
        ppt_dir = Path(st.text_input("课件目录", value=str(DEFAULT_PPT_DIR)))
        question_dir = Path(st.text_input("题库目录", value=str(DEFAULT_QUESTION_DIR)))
        solution_dir = Path(st.text_input("解析目录", value=str(DEFAULT_SOLUTION_DIR)))
        force_rebuild = st.checkbox("强制重建数据", value=False)
        trigger_rebuild = st.button("重建数据")

    dataset = load_or_build_dataset(
        ppt_dir=ppt_dir,
        question_dir=question_dir,
        solution_dir=solution_dir,
        out_path=DATASET_PATH,
        force_rebuild=force_rebuild or trigger_rebuild,
    )
    progress = load_progress(PROGRESS_PATH)

    with st.sidebar:
        st.markdown("## 数据概览")
        meta = dataset["meta"]
        st.write(f"总题数：{meta['total_questions']}")
        st.write(f"微观题：{meta['micro_questions']}")
        st.write(f"宏观题：{meta['macro_questions']}")
        st.write(f"Lecture 数：{meta['total_lectures']}")
        st.write(f"错题本：{len(progress.get('wrong_book', {}))}")

    lecture_options = [
        (lec["lecture_id"], f"Lecture {lec['lecture_no']:02d} | {lec['lecture_title']}")
        if lec["lecture_id"] != "other"
        else (lec["lecture_id"], "Other 其他题目")
        for lec in dataset["lectures"]
    ]

    page_items = [f"Lecture 练习 | {name}" for _, name in lecture_options]
    page_items += ["Final 抽测模式", "错题本"]
    page = st.selectbox("选择模式", page_items)

    if page == "Final 抽测模式":
        progress = render_final_mode(dataset, progress)
    elif page == "错题本":
        progress = render_wrong_book(progress, dataset)
    else:
        selected_name = page.split("|", 1)[1].strip()
        selected_id = None
        for lid, lname in lecture_options:
            if lname == selected_name:
                selected_id = lid
                break
        if selected_id is None:
            st.error("未找到对应 Lecture。")
            return
        progress = render_practice_mode(dataset, progress, selected_id)

    save_progress(PROGRESS_PATH, progress)


if __name__ == "__main__":
    main()
