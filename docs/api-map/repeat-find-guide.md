# RepeatFind — HWP 유일의 프로그래밍용 찾기 API

## 왜 RepeatFind만 되는가?

HWP에는 찾기 관련 Action이 10개 존재하지만, **프로그래밍으로 작동하는 건 2개뿐**:

| API | 작동 | 용도 |
|-----|:----:|------|
| **RepeatFind** | ✅ | 하나씩 찾기 + 커서 이동 |
| **AllReplace** | ✅ | 전체 일괄 치환 (커서 이동 없음) |
| Find, FindReplace, FindDlg, ReplaceDlg, FindNext, FindPrev, GoToFind, FindContinue | ❌ | 전부 안 됨 (HWP 매크로에서도 동일) |

`FindReplace`는 `CreateAction("FindReplace")`가 null을 반환. HWP 자체적으로 프로그래밍 인터페이스를 제공하지 않는 Action.

---

## 기본 사용법

```js
hwp.Run("MoveDocBegin");
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "찾을텍스트";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;  // 필수: "찾지 못했습니다" 팝업 억제
var found = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
// found=true → 텍스트가 선택(드래그)된 상태
// found=false → 못 찾음
```

찾으면 해당 텍스트가 **선택된 상태**가 되므로 바로 작업 가능:
```js
hwp.Run("Delete");                          // 삭제
hwp.InsertPicture(path, 1, 0,0,0,0,0,0);   // 이미지로 교체
// 또는 InsertText로 다른 텍스트 삽입
```

---

## 찾은 위치 정보 읽기

```js
var ps = hwp.GetPosBySet();
var list = ps.Item("List");   // 0=본문, >0=표 셀
var para = ps.Item("Para");   // 문단 번호
var pos  = ps.Item("Pos");    // 오프셋

var sel = hwp.GetTextFile("UNICODE", "saveblock");  // 선택된 텍스트
```

---

## 여러 개 찾기 (순환 주의)

RepeatFind는 문서 끝에 도달하면 **처음으로 돌아가면서 true를 반환**. 무한루프 방지를 위해 위치 추적 필수:

```js
hwp.Run("MoveDocBegin");
var visited = {};

while (true) {
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.FindString = "{{MARKER}}";
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;

    if (!hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)) break;

    // 순환 감지
    var ps = hwp.GetPosBySet();
    var key = ps.Item("List") + "," + ps.Item("Para") + "," + ps.Item("Pos");
    if (visited[key]) break;  // 이미 방문한 위치 → 순환 끝
    visited[key] = true;

    // 찾은 위치에서 작업
    hwp.Run("Delete");
    // ...
}
```

---

## 주요 파라미터

| 파라미터 | 값 | 설명 |
|----------|-----|------|
| `FindString` | 문자열 | 찾을 텍스트 |
| `IgnoreMessage` | 1 | 팝업 억제 (필수) |
| `Direction` | `hwp.FindDir("Forward")` | 앞으로 검색 |
| `Direction` | `hwp.FindDir("AllDoc")` | 전체 문서 검색 (값: 2) |
| `FindRegExp` | 1 | 정규식 사용 (`^n`=줄바꿈 등) |
| `MatchCase` | 1 | 대소문자 구분 |
| `WholeWordOnly` | 1 | 전체 단어만 |
| `FindType` | 1 | 검색 유형 |

---

## 서식으로 찾기

텍스트가 아닌 **서식(볼드, 글꼴, 색상 등)**으로 찾기도 가능:

```js
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "";  // 빈 문자열 = 서식만 검색
hwp.HParameterSet.HFindReplace.FindCharShape.Bold = 1;         // 볼드 텍스트 찾기
hwp.HParameterSet.HFindReplace.FindCharShape.TextColor = hwp.RGBColor(255, 0, 0);  // 빨간 글자 찾기
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
```

---

## 검색 범위

| 범위 | 방법 |
|------|------|
| 전체 문서 | `MoveDocBegin` 후 RepeatFind |
| 표 안만 | 표 안에서 커서 위치 후 RepeatFind (표 밖도 찾아감 주의) |
| 특정 범위 | 현재로선 범위 제한 불가. 찾은 후 위치 체크로 필터링 |

---

## 실전 패턴

### 패턴 1: 플레이스홀더 → 이미지 삽입

```js
// 1. xml.set으로 플레이스홀더 삽입
xml.load();
xml.set("t0.r1.c0", "{{PHOTO}}");
xml.commit();

// 2. RepeatFind로 찾아서 이미지 교체
hwp.Run("MoveDocBegin");
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "{{PHOTO}}";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.Run("Delete");

var ctrl = hwp.InsertPicture("C:/path/photo.jpg", 1, 0, 0, 0, 0, 0, 0);
// 크기 조절
var props = ctrl.Properties;
props.SetItem("Width", 8041);
props.SetItem("Height", 10440);
ctrl.Properties = props;
```

### 패턴 2: 특정 셀로 커서 이동

```js
// 1. 유니크 마커 삽입
xml.load();
xml.set("t0.r5.c3", "{{CURSOR_HERE}}");
xml.commit();

// 2. RepeatFind로 이동
hwp.Run("MoveDocBegin");
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "{{CURSOR_HERE}}";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.Run("Delete");

// 이제 커서가 t0.r5.c3 셀 안에 위치
// InsertText, InsertPicture 등 자유롭게 작업
```

### 패턴 3: 여러 플레이스홀더 일괄 처리

```js
xml.load();
xml.set("t0.r1.c0", "{{PHOTO}}");
xml.set("t0.r6.c1", "{{PERIOD_1}}");
xml.set("t0.r7.c1", "{{PERIOD_2}}");
xml.commit();

// AllReplace로 텍스트 치환 (커서 이동 불필요)
AllReplace("{{PERIOD_1}}", "2014.03 ~ 2018.02");
AllReplace("{{PERIOD_2}}", "2011.03 ~ 2014.02");

// RepeatFind로 이미지만 별도 처리
hwp.Run("MoveDocBegin");
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "{{PHOTO}}";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.Run("Delete");
hwp.InsertPicture("C:/path/photo.jpg", 1, 0, 0, 0, 0, 0, 0);
```
