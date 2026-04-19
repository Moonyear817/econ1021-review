import html
import json
import random
import re
import subprocess
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZipFile

STOPWORDS = {
    "the", "and", "for", "are", "with", "that", "from", "this", "your", "you",
    "into", "their", "have", "has", "had", "its", "not", "can", "will", "would",
    "should", "what", "which", "when", "where", "why", "how", "all", "more", "less",
    "than", "over", "under", "using", "used", "use", "about", "into", "only", "also",
    "chapter", "lecture", "spring", "economics", "market", "system", "micro", "macro",
    "pages", "page", "objective", "learning", "outcome", "identify", "discuss", "explain",
    "analyze", "analysis", "question", "questions", "answer", "answers", "topic", "diff",
    "recurring", "figure", "refer", "one", "two", "three", "four", "five", "six",
    "seven", "eight", "nine", "ten", "a", "b", "c", "d"
}


@dataclass
class Question:
    qid: str
    chapter: int
    qnum: int
    stem: str
    options: dict[str, str]
    answer: str
    topic: str
    learning_outcome: str
    explanation_en: str
    explanation_zh: str
    lecture_id: str
    lecture_title: str
    source_file: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "qid": self.qid,
            "chapter": self.chapter,
            "qnum": self.qnum,
            "stem": self.stem,
            "options": self.options,
            "answer": self.answer,
            "topic": self.topic,
            "learning_outcome": self.learning_outcome,
            "explanation_en": self.explanation_en,
            "explanation_zh": self.explanation_zh,
            "lecture_id": self.lecture_id,
            "lecture_title": self.lecture_title,
            "source_file": self.source_file,
            "is_micro": self.chapter <= 18,
        }


def read_doc_text(doc_path: Path) -> str:
    cmd = ["textutil", "-convert", "txt", "-stdout", str(doc_path)]
    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        return ""
    return result.stdout


def tokenize(text: str) -> list[str]:
    words = re.findall(r"[a-zA-Z][a-zA-Z\-]{1,}", text.lower())
    return [w for w in words if w not in STOPWORDS and len(w) >= 3]


def extract_ppt_text(pptx_path: Path) -> str:
    texts: list[str] = []
    try:
        with ZipFile(pptx_path, "r") as zf:
            slide_names = [
                name for name in zf.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ]
            slide_names.sort(key=lambda s: int(re.search(r"slide(\d+)\.xml", s).group(1)))
            for name in slide_names:
                xml_text = zf.read(name).decode("utf-8", errors="ignore")
                for frag in re.findall(r"<a:t>(.*?)</a:t>", xml_text):
                    clean = html.unescape(frag).strip()
                    if clean:
                        texts.append(clean)
    except Exception:
        return ""
    return "\n".join(texts)


def parse_lecture_meta(ppt_dir: Path) -> list[dict[str, Any]]:
    lectures: list[dict[str, Any]] = []
    for ppt in sorted(ppt_dir.glob("*.pptx")):
        m = re.search(r"Lecture\s*(\d+)", ppt.name, flags=re.IGNORECASE)
        if not m:
            continue
        lecture_no = int(m.group(1))
        lecture_id = f"lecture_{lecture_no:02d}"
        text = extract_ppt_text(ppt)
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        title = lines[0] if lines else f"Lecture {lecture_no:02d}"

        token_counts = Counter(tokenize(text))
        keywords = [w for w, _ in token_counts.most_common(120)]

        chapter_refs = sorted(
            {
                int(x)
                for x in re.findall(r"\\bchapter\\s*(\\d+)\\b", text, flags=re.IGNORECASE)
            }
        )

        title = f"Lecture {lecture_no:02d}"
        if chapter_refs:
            chapter_str = ", ".join(str(c) for c in chapter_refs[:3])
            title = f"Lecture {lecture_no:02d} (Chapter {chapter_str})"

        lectures.append(
            {
                "lecture_id": lecture_id,
                "lecture_no": lecture_no,
                "lecture_title": title,
                "keywords": keywords,
                "chapter_refs": chapter_refs,
                "raw_text": text,
                "ppt_file": str(ppt),
            }
        )

    lectures.sort(key=lambda x: x["lecture_no"])
    return lectures


def chapter_from_filename(name: str) -> int | None:
    patterns = [r"_C(\d+)\.doc$", r"Ch(\d+)\.doc$"]
    for p in patterns:
        m = re.search(p, name, flags=re.IGNORECASE)
        if m:
            return int(m.group(1))
    return None


def parse_question_doc(doc_text: str, chapter: int, source_file: str) -> list[Question]:
    lines = [ln.rstrip() for ln in doc_text.splitlines()]
    questions: list[Question] = []

    current: dict[str, Any] | None = None
    current_option: str | None = None

    q_start = re.compile(r"^\s*(\d+)\)\s+(.+)\s*$")
    opt_start = re.compile(r"^\s*([A-D])\)\s+(.+)\s*$")
    answer_line = re.compile(r"^\s*Answer:\s*([A-D])\s*$", flags=re.IGNORECASE)
    topic_line = re.compile(r"^\s*Topic:\s*(.+)$", flags=re.IGNORECASE)
    lo_line = re.compile(r"^\s*Learning Outcome:\s*(.+)$", flags=re.IGNORECASE)
    diff_line = re.compile(r"^\s*Diff:\s*.*$", flags=re.IGNORECASE)
    page_ref_line = re.compile(r"^\s*Page\s*Ref:\s*.*$", flags=re.IGNORECASE)
    recurring_line = re.compile(r"^\s*\*:\s*Recurring\s*$", flags=re.IGNORECASE)
    aacsb_line = re.compile(r"^\s*AACSB:\s*.*$", flags=re.IGNORECASE)
    special_feature_line = re.compile(r"^\s*Special\s*Feature:\s*.*$", flags=re.IGNORECASE)

    def strip_stem_artifacts(text: str) -> str:
        cleaned = re.sub(r"\bDiff:\s*[^\n\r]*", "", text, flags=re.IGNORECASE)
        cleaned = re.sub(r"\bPage\s*Ref:\s*[^\n\r]*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\bAACSB:\s*[^\n\r]*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\*:\s*Recurring", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s{2,}", " ", cleaned)
        return cleaned.strip()

    def flush_current() -> None:
        nonlocal current
        if not current:
            return
        if current.get("stem") and current.get("options") and current.get("answer"):
            qid = f"C{chapter:02d}-Q{current['qnum']:04d}"
            q = Question(
                qid=qid,
                chapter=chapter,
                qnum=current["qnum"],
                stem=strip_stem_artifacts(current["stem"].strip()),
                options=current["options"],
                answer=current["answer"].upper(),
                topic=current.get("topic", "").strip(),
                learning_outcome=current.get("learning_outcome", "").strip(),
                explanation_en="",
                explanation_zh="",
                lecture_id="unclassified",
                lecture_title="Other 其他题目",
                source_file=source_file,
            )
            questions.append(q)
        current = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        m_q = q_start.match(line)
        if m_q:
            flush_current()
            current = {
                "qnum": int(m_q.group(1)),
                "stem": m_q.group(2).strip(),
                "options": {},
                "answer": "",
                "topic": "",
                "learning_outcome": "",
            }
            current_option = None
            continue

        if not current:
            continue

        m_opt = opt_start.match(line)
        if m_opt:
            opt_key = m_opt.group(1).upper()
            current["options"][opt_key] = m_opt.group(2).strip()
            current_option = opt_key
            continue

        m_ans = answer_line.match(line)
        if m_ans:
            current["answer"] = m_ans.group(1).upper()
            current_option = None
            continue

        m_topic = topic_line.match(line)
        if m_topic:
            current["topic"] = m_topic.group(1).strip()
            current_option = None
            continue

        m_lo = lo_line.match(line)
        if m_lo:
            current["learning_outcome"] = m_lo.group(1).strip()
            current_option = None
            continue

        if (
            diff_line.match(line)
            or page_ref_line.match(line)
            or recurring_line.match(line)
            or aacsb_line.match(line)
            or special_feature_line.match(line)
        ):
            current_option = None
            continue

        if current_option and current_option in current["options"]:
            current["options"][current_option] += " " + line
        else:
            current["stem"] += " " + line

    flush_current()
    return questions


def parse_solution_notes(solution_dir: Path) -> dict[int, str]:
    notes: dict[int, str] = {}
    for doc in sorted(solution_dir.glob("*.doc")):
        chapter = chapter_from_filename(doc.name)
        if chapter is None:
            continue
        txt = read_doc_text(doc)
        if not txt:
            continue
        # Keep a compact note body for explanation lookup.
        compact = "\n".join([ln.strip() for ln in txt.splitlines() if ln.strip()])
        notes[chapter] = compact[:60000]
    return notes


EN2ZH_GLOSSARY = {
    "opportunity cost": "机会成本",
    "trade-off": "权衡取舍",
    "tradeoffs": "权衡取舍",
    "scarcity": "稀缺性",
    "comparative advantage": "比较优势",
    "absolute advantage": "绝对优势",
    "production possibilities frontier": "生产可能性边界",
    "ppf": "生产可能性边界",
    "market": "市场",
    "demand": "需求",
    "supply": "供给",
    "equilibrium": "均衡",
    "marginal": "边际",
    "efficiency": "效率",
    "inflation": "通货膨胀",
    "gdp": "国内生产总值",
    "unemployment": "失业",
    "fiscal policy": "财政政策",
    "monetary policy": "货币政策",
}


def summarize_en_excerpt(topic: str, chapter_notes: str, answer: str) -> str:
    if not chapter_notes:
        return (
            f"Correct answer: {answer}. "
            "Focus on the core concept in the stem and compare all options by elimination."
        )

    lower_notes = chapter_notes.lower()
    topic_tokens = tokenize(topic)
    for tk in topic_tokens[:5]:
        pos = lower_notes.find(tk)
        if pos != -1:
            start = max(0, pos - 280)
            end = min(len(chapter_notes), pos + 420)
            excerpt = chapter_notes[start:end].strip()
            return f"Correct answer: {answer}. Relevant chapter excerpt:\n\n{excerpt}"

    fallback = chapter_notes[:700].strip()
    return f"Correct answer: {answer}. Chapter reference:\n\n{fallback}"


def build_chinese_explanation(
    *,
    answer: str,
    topic: str,
    stem: str,
    options: dict[str, str],
    explanation_en: str,
) -> str:
    lower_all = f"{topic} {stem} {explanation_en}".lower()
    hit_terms: list[str] = []
    for en_term, zh_term in EN2ZH_GLOSSARY.items():
        if en_term in lower_all and zh_term not in hit_terms:
            hit_terms.append(zh_term)
        if len(hit_terms) >= 5:
            break

    option_text = options.get(answer, "")
    key_line = ""
    if option_text:
        key_line = f"正确选项 {answer} 的关键点是：{option_text}。"

    concept_line = ""
    if hit_terms:
        concept_line = f"本题核心概念：{'、'.join(hit_terms)}。"

    en_ref = "\n\n[English Reference]\n" + explanation_en[:1200]
    return (
        f"正确答案：{answer}。"
        f"{key_line}"
        f"{concept_line}"
        "解题建议：先抓题干中的经济学概念，再逐项比对选项，优先排除与概念冲突的选项。"
        f"{en_ref}"
    )


def classify_question(question: Question, lectures: list[dict[str, Any]]) -> tuple[str, str]:
    # Deterministic grouping: Lecture 01-06 first, anything else goes to Other.
    top_lectures = [lec for lec in lectures if 1 <= lec.get("lecture_no", 0) <= 6]
    top_lectures.sort(key=lambda x: x["lecture_no"])

    for lec in top_lectures:
        chapter_refs = lec.get("chapter_refs", [])
        if question.chapter in chapter_refs:
            return lec["lecture_id"], lec["lecture_title"]

    # Fallback scoring only inside lecture 01-06 to avoid noisy mapping.
    q_tokens = set(tokenize(f"{question.stem} {question.topic} {question.learning_outcome}"))
    if q_tokens:
        best_score = 0
        best_lecture: dict[str, Any] | None = None
        for lec in top_lectures:
            lec_kw = set(lec.get("keywords", []))
            overlap = q_tokens.intersection(lec_kw)
            score = len(overlap) + 2 * len(set(tokenize(question.topic)).intersection(lec_kw))
            if score > best_score:
                best_score = score
                best_lecture = lec
        if best_lecture and best_score >= 3:
            return best_lecture["lecture_id"], best_lecture["lecture_title"]

    return "other", "Other 其他题目"


def build_dataset(ppt_dir: Path, question_dir: Path, solution_dir: Path, out_path: Path) -> dict[str, Any]:
    lectures = parse_lecture_meta(ppt_dir)
    solution_notes = parse_solution_notes(solution_dir)

    top_lectures = [lec for lec in lectures if 1 <= lec.get("lecture_no", 0) <= 6]
    top_lectures.sort(key=lambda x: x["lecture_no"])

    all_questions: list[Question] = []
    for doc in sorted(question_dir.glob("*.doc")):
        chapter = chapter_from_filename(doc.name)
        if chapter is None:
            continue
        txt = read_doc_text(doc)
        if not txt:
            continue
        all_questions.extend(parse_question_doc(txt, chapter, doc.name))

    for q in all_questions:
        q.explanation_en = summarize_en_excerpt(q.topic, solution_notes.get(q.chapter, ""), q.answer)
        q.explanation_zh = build_chinese_explanation(
            answer=q.answer,
            topic=q.topic,
            stem=q.stem,
            options=q.options,
            explanation_en=q.explanation_en,
        )
        lecture_id, lecture_title = classify_question(q, top_lectures)
        q.lecture_id = lecture_id
        q.lecture_title = lecture_title

    lecture_map: dict[str, dict[str, Any]] = {
        lec["lecture_id"]: {
            "lecture_id": lec["lecture_id"],
            "lecture_no": lec["lecture_no"],
            "lecture_title": lec["lecture_title"],
            "keywords": lec["keywords"][:30],
            "question_ids": [],
        }
        for lec in top_lectures
    }
    lecture_map["other"] = {
        "lecture_id": "other",
        "lecture_no": 999,
        "lecture_title": "Other 其他题目",
        "keywords": [],
        "question_ids": [],
    }

    question_dicts = [q.to_dict() for q in all_questions]
    question_dicts.sort(key=lambda x: (x["chapter"], x["qnum"]))

    for q in question_dicts:
        lid = q["lecture_id"] if q["lecture_id"] in lecture_map else "other"
        lecture_map[lid]["question_ids"].append(q["qid"])

    # Keep question order stable inside each lecture by chapter then question number.
    qmeta = {q["qid"]: (q["chapter"], q["qnum"]) for q in question_dicts}
    for lec in lecture_map.values():
        lec["question_ids"].sort(key=lambda qid: qmeta.get(qid, (999, 999999)))

    dataset = {
        "meta": {
            "total_questions": len(question_dicts),
            "total_lectures": len(top_lectures),
            "micro_questions": sum(1 for q in question_dicts if q["is_micro"]),
            "macro_questions": sum(1 for q in question_dicts if not q["is_micro"]),
        },
        "lectures": sorted(lecture_map.values(), key=lambda x: x["lecture_no"]),
        "questions": question_dicts,
    }

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    return dataset


def load_or_build_dataset(
    ppt_dir: Path,
    question_dir: Path,
    solution_dir: Path,
    out_path: Path,
    force_rebuild: bool = False,
) -> dict[str, Any]:
    if out_path.exists() and not force_rebuild:
        return json.loads(out_path.read_text(encoding="utf-8"))
    return build_dataset(ppt_dir, question_dir, solution_dir, out_path)


def sample_final_questions(dataset: dict[str, Any], n: int = 30, seed: int | None = None) -> list[str]:
    micro_ids = [q["qid"] for q in dataset["questions"] if q.get("is_micro")]
    if seed is not None:
        random.seed(seed)
    if len(micro_ids) <= n:
        return micro_ids
    return random.sample(micro_ids, n)
