# -*- coding: utf-8 -*-
"""
qa-queries.json 자체 검증기.

LLM이 파일을 다시 읽으며 눈으로 확인하던 자체검증을 대체한다.
같은 검사를 0.1초에 끝내므로 분석 시간이 크게 줄어든다.
"""
import json, sys
from pathlib import Path

REQUIRED = ["id", "element_no", "category", "case", "quote",
            "risk", "expected_action", "priority", "owner", "status"]

PLACEHOLDER = ["확정 필요", "확인 필요", "검토 필요", "기획자 확인",
               "미정", "TBD", "추후 결정", "협의 필요"]

VALID_PRIORITY = {"필수", "높음", "중간"}


def validate(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    queries = data.get("queries", [])
    errors, warnings = [], []

    if not queries:
        errors.append("queries 배열이 비어있음")

    seen_ids = set()
    for i, q in enumerate(queries):
        tag = q.get("id", f"[{i}번째 항목]")

        for f in REQUIRED:
            if f not in q:
                errors.append(f"{tag}: 필드 '{f}' 누락")
            elif f != "element_no" and not str(q[f]).strip():
                errors.append(f"{tag}: 필드 '{f}' 가 비어있음")

        if q.get("id") in seen_ids:
            errors.append(f"{tag}: ID 중복")
        seen_ids.add(q.get("id"))

        # 근거 없는 질의 = 창작
        quote = str(q.get("quote", ""))
        if quote and "[" not in quote:
            warnings.append(f"{tag}: quote에 출처 표기([슬라이드 N])가 없음")

        # 권고안 플레이스홀더 검출
        action = str(q.get("expected_action", ""))
        for p in PLACEHOLDER:
            if p in action:
                errors.append(f"{tag}: 권고안이 플레이스홀더('{p}'). 구체적 대안으로 재작성 필요")
                break

        # 내부 코드 노출 검출
        cat = str(q.get("category", ""))
        if cat.startswith(("A-", "B-", "KWCAG")):
            warnings.append(f"{tag}: category에 내부 코드('{cat}') 사용. 자연어로 변경 권장")

        pr = q.get("priority")
        if pr and pr not in VALID_PRIORITY:
            errors.append(f"{tag}: priority '{pr}' 는 허용값(필수/높음/중간) 아님")

    # ---- 결과 ----
    print(f"검증 대상: {path}")
    print(f"질의 {len(queries)}건 | 충돌 {len(data.get('conflicts', []))}건 | "
          f"접근성 {len(data.get('accessibility', []))}건")

    for w in warnings:
        print(f"  [경고] {w}")

    if errors:
        print(f"\n실패 {len(errors)}건:")
        for e in errors:
            print(f"  [오류] {e}")
        print("\n위 항목을 수정한 뒤 다시 실행하세요.")
        return 1

    by_pr = {}
    for q in queries:
        by_pr[q.get("priority")] = by_pr.get(q.get("priority"), 0) + 1
    print(f"\n통과. 우선순위별: " +
          ", ".join(f"{k} {v}건" for k, v in sorted(by_pr.items(),
                    key=lambda x: ["필수", "높음", "중간"].index(x[0])
                    if x[0] in VALID_PRIORITY else 9)))
    return 0


if __name__ == "__main__":
    sys.exit(validate(sys.argv[1] if len(sys.argv) > 1 else "output/qa-queries.json"))