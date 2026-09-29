"""공사 유형별 대표 공종 (작업 순서대로). 항목은 1일작업량(토목) 시트의 선택명을 찾는 키워드 또는 직접입력용 이름."""
import workrates_civil

KEYS = [k for k, *_ in workrates_civil.build()]


def find(*subs):
    hits = [k for k in KEYS if all(s in k for s in subs)]
    assert hits, subs
    return hits[0]


def custom(name):
    return f"{name} [1일 작업량 직접 입력]"


PRESETS = {
    "하수도 관로공사": [
        find("가설울타리"), find("터파기-토사"),
        find("배수관-흄관"), custom("맨홀 설치"), find("흙쌓기-되메우기"),
        custom("굴착부 아스팔트 복구포장 (폭 3m 미만)"),
    ],
    "도로 토공+아스팔트포장": [
        find("가설울타리"), find("벌개제근"), find("표토제거"), find("흙깎기-토사"),
        find("흙쌓기-노체"), find("흙쌓기-노상"), find("배수관-흄관"), find("측구공-U형측구"),
        find("아스팔트포장 1단계-보조기층"), find("아스팔트포장 2단계-기층", "5-7"),
        find("아스팔트포장 2단계-중간층/표층"), find("아스팔트포장-부대공사"),
    ],
    "도로 콘크리트포장": [
        find("가설울타리"), find("표토제거"), find("흙쌓기-노체"), find("흙쌓기-노상"),
        find("콘크리트포장 1단계-린콘크리트"), find("콘크리트포장 2단계-콘크리트표층", "1차로"),
        find("콘크리트포장 2단계-포장절단"), find("콘크리트포장-부대공사"),
    ],
    "교량공사(PSC빔)": [
        find("직접기초"), find("말뚝기초"), find("교대 벽체"), find("교각 기둥"), find("교각 코핑"),
        find("거더 제작"), find("거더 운반 및 거치"), find("상부슬래브"),
        find("교면포장(LMC) 포설"), find("교면포장(LMC) 마무리"), find("교량부대공사"),
    ],
    "터널공사": [
        find("갱구부 보강-시점부"), find("갱구부 보강-종점부"), find("굴착 및 보강-P-1"),
        find("강관보강 그라우팅"), find("후속공종-방수"), find("후속공종-배수공동구"),
        find("후속공종-라이닝 콘크리트"), find("후속공종-갱문조성"),
    ],
    "옹벽·배수구조물": [
        find("가설울타리"), find("터파기-토사"), find("암거공-철근콘크리트"), find("암거공-날개벽"),
        find("옹벽-콘크리트옹벽"), find("옹벽-보강토옹벽"), find("측구공-U형측구"), find("흙쌓기-되메우기"),
    ],
}
assert all(len(v) <= 15 for v in PRESETS.values())
