# -*- coding: utf-8 -*-
"""
qa-queries.json -> qa-query-sheet.md / .csv

필드명은 planning-analyst.md(슬림 버전)의 출력 스키마와 1:1로 일치한다.
id, element_no, category, case, quote, risk, expected_action, priority, owner, status
"""
import json, csv, sys
from pathlib import Path

PRIORITY = {"필수": 0, "높음": 1, "중간": 2}
HEADER = ["ID", "요소", "구분", "상황", "원문인용", "예상리스크",
          "기대동작(QA권고)", "우선순위", "담당", "상태"]


def sort_key(q):
    # 요소 번호 순 -> 같은 요소 안에서 우선순위 순. 번호 없는 항목은 맨 뒤
    elem = q.get("element_no") or "￿"
    return (elem, PRIORITY.get(q.get("priority"), 9))


def row(q):
    return [q.get("id", ""), q.get("element_no", "-"), q.get("category", ""),
            q.get("case", ""), q.get("quote", ""), q.get("risk", ""),
            f"[Q] {q.get('expected_action', '')}", q.get("priority", ""),
            q.get("owner", ""), q.get("status", "대기")]


def render(json_path, out_dir):
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    queries = sorted(data.get("queries", []), key=sort_key)
    conflicts = data.get("conflicts", [])
    a11y = data.get("accessibility", [])
    unreadable = data.get("unreadable", [])

    L = [f"# {data.get('document', '')} - QA 기획 질의서\n",
         f"- 분석일: {data.get('generated_at', '')}",
         f"- 질의 {len(queries)}건 / 문서 간 충돌 {len(conflicts)}건 / "
         f"접근성 {len(a11y)}건\n",
         "\n## 1. 확인 필요 사항\n",
         "| " + " | ".join(HEADER) + " |",
         "|" + "---|" * len(HEADER)]

    for q in queries:
        L.append("| " + " | ".join(str(c).replace("\n", " ") for c in row(q)) + " |")

    L.append("\n## 2. 문서 간 충돌\n")
    L.append("\n".join(f"- {c}" for c in conflicts) if conflicts else "0건 확인.")

    L.append("\n## 3. 접근성 (KWCAG 2.2)\n")
    L.append("\n".join(f"- {a}" for a in a11y) if a11y else "미시행 (요청 시 수행).")

    L.append("\n## 4. 분석 불가 구간\n")
    L.append("\n".join(f"- {u}" for u in unreadable) if unreadable else "없음.")

    Path(out_dir, "qa-query-sheet.md").write_text("\n".join(L), encoding="utf-8")

    with open(Path(out_dir, "qa-query-sheet.csv"), "w",
              encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, quoting=csv.QUOTE_ALL)
        w.writerow(HEADER)
        for q in queries:
            w.writerow(row(q))

    print(f"질의 {len(queries)}건 -> qa-query-sheet.md / .csv 생성 완료")


if __name__ == "__main__":
    render(sys.argv[1] if len(sys.argv) > 1 else "output/qa-queries.json",
           sys.argv[2] if len(sys.argv) > 2 else "output")