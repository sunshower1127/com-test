# 단계 11: 컨트롤 순회

## 테스트 결과

모든 테스트 ✅ 통과.

### 컨트롤 순회

```js
var ctrl = hwp.HeadCtrl;   // 첫 번째 컨트롤
while (ctrl) {
  var id = ctrl.CtrlID;       // "secd", "tbl", "gso" 등
  var desc = ctrl.UserDesc;   // "구역 정의", "표", "그림" 등
  var props = ctrl.Properties; // ParameterSet
  ctrl = ctrl.Next;            // 다음 컨트롤 (null이면 끝)
}

var last = hwp.LastCtrl;    // 마지막 컨트롤
last.Prev;                   // 역순회
```

### 컨트롤 ID

| CtrlID | UserDesc | Properties SetID | 설명 |
|--------|----------|-----------------|------|
| `secd` | 구역 정의 | SecDef | 섹션 속성 |
| `cold` | 단 정의 | ColDef | 칼럼 속성 |
| `tbl` | 표 | Table | 표 컨트롤 |
| `gso` | 그림 | ShapeObject | 그림/도형 |

### GetAnchorPos — 컨트롤 위치

```js
var anchorPos = ctrl.GetAnchorPos(0);
anchorPos.Item("List");  // 리스트 ID
anchorPos.Item("Para");  // 문단 번호
anchorPos.Item("Pos");   // 위치 오프셋
```

표: L=0, P=1, pos=0 / 그림: L=0, P=4, pos=0 (문서에 따라 다름)

### Properties 읽기

```js
var ctrl = hwp.HeadCtrl;
// 표 찾기
while (ctrl && ctrl.CtrlID !== "tbl") ctrl = ctrl.Next;

var props = ctrl.Properties;
props.Item("Width");        // 12244
props.Item("Height");       // 2564
props.Item("CellSpacing");  // 0
```

### DeleteCtrl — 컨트롤 삭제

```js
hwp.DeleteCtrl(ctrl);  // ✅ 작동 — 그림 삭제 확인
```

### SetPosBySet으로 컨트롤 위치로 이동

```js
var anchorPos = ctrl.GetAnchorPos(0);
hwp.SetPosBySet(anchorPos);  // 컨트롤 위치로 커서 이동
```
