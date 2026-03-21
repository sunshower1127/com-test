# HWP COM Enum 상수값 레퍼런스

`hwp.EnumMethod("StringValue")` → 정수 반환. 케이스 센시티브.

---

## HwpLineType (선 종류)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Solid"` | 실선 |
| 1 | `"Dash"` | 긴 점선 |
| 2 | `"Dot"` | 점선 |
| 3 | `"DashDot"` | -.-.-. |
| 4 | `"DashDotDot"` | -..-.. |
| 5 | `"LongDash"` | 긴 대시 |
| 6 | `"Circle"` | 큰 점 |
| 7 | `"DoubleSlim"` | 2중선 |
| 8 | `"SlimThick"` | 가는선+굵은선 |
| 9 | `"ThickSlim"` | 굵은선+가는선 |
| 10 | `"SlimThickSlim"` | 가는선+굵은선+가는선 |
| 11 | `"Wave"` | 물결 |
| 12 | `"DoubleWave"` | 물결 2중선 |
| 13 | `"Thick3D"` | 두꺼운 3D |
| 14 | `"Thick3DInset"` | 두꺼운 3D (광원 반대) |
| 15 | `"Slim3D"` | 3D 단선 |
| 16 | `"Slim3DInset"` | 3D 단선 (광원 반대) |

`"None"` 도 유효 (선 없음).

---

## HAlign (가로 정렬)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Justify"` | 양쪽 정렬 |
| 1 | `"Left"` | 왼쪽 |
| 2 | `"Right"` | 오른쪽 |
| 3 | `"Center"` | 가운데 |
| 4 | `"Distributive"` | 배분 |
| 5 | `"DistributiveSpace"` | 나눔 |

---

## VAlign (세로 정렬)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Baseline"` | 글꼴 기준 |
| 1 | `"Top"` | 위쪽 |
| 2 | `"Center"` | 가운데 |
| 3 | `"Bottom"` | 아래 |

---

## TextWrapType (텍스트 배치)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Square"` | 어울림 |
| 1 | `"Tight"` | outline 따라 |
| 2 | `"Through"` | 빈 공간까지 |
| 3 | `"TopAndBottom"` | 좌우 배치 안 함 |
| 4 | `"BehindText"` | 글 뒤로 |
| 5 | `"InFrontOfText"` | 글 앞으로 |

---

## HorzRel (가로 위치 기준)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Paper"` | 종이 |
| 1 | `"Page"` | 페이지 |
| 2 | `"Column"` | 단 |
| 3 | `"Para"` | 문단 |

---

## VertRel (세로 위치 기준)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Paper"` | 종이 |
| 1 | `"Page"` | 페이지 |
| 2 | `"Para"` | 문단 |

---

## TextFlowType (글 흐름 방향)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"BothSides"` | 양쪽 |
| 1 | `"LeftOnly"` | 왼쪽만 |
| 2 | `"RightOnly"` | 오른쪽만 |
| 3 | `"LargestOnly"` | 큰 쪽만 |

---

## BrushType (채우기 종류)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"None"` | 없음 |
| 1 | `"Color"` | 단색 |
| 2 | `"Image"` | 이미지 |
| 4 | `"Gradation"` | 그라데이션 |

---

## HatchStyle (해치 패턴)

| 값 | COM String | 설명 |
|----|-----------|------|
| -1 | `"None"` | 단색 |
| 0 | `"HorizontalLine"` | 가로줄 |
| 1 | `"VerticalLine"` | 세로줄 |
| 2 | `"BackDiagonalLine"` | \\\\ 대각선 |
| 3 | `"FrontDiagonalLine"` | //// 대각선 |
| 4 | `"CrossLine"` | 격자 |
| 5 | `"CrossDiagonalLine"` | X자 격자 |

---

## Gradation (그라데이션 종류)

| 값 | COM String | 설명 |
|----|-----------|------|
| 1 | `"Stripe"` | 줄무늬형 |
| 2 | `"Circle"` | 원형 |
| 3 | `"Cone"` | 원뿔형 |
| 4 | `"Square"` | 사각형 |

---

## HwpUnderlineType (밑줄 위치)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"None"` | 없음 |
| 1 | `"Bottom"` | 밑줄 |
| 2 | `"Top"` | 윗줄 |

밑줄 모양은 HwpLineType과 동일한 상수 사용.

---

## CharShadowType (글자 그림자)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"None"` | 없음 |
| 1 | `"Drop"` | 비연속 |
| 2 | `"Continuous"` | 연속 |

---

## DSMark (강조점)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"None"` | 없음 |
| 1 | `"DotAbove"` | ● |
| 2 | `"RingAbove"` | ○ |
| 3 | `"Caron"` | ˇ |
| 4 | `"Tilde"` | ~ |
| 5 | `"DotMiddle"` | ･ |
| 6 | `"Colon"` | : |

---

## LineSpacingMethod (줄간격 종류)

| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Percent"` | 글자에 따라 (%) |
| 1 | `"Fixed"` | 고정값 |
| 2 | `"BetweenLine"` | 여백만 지정 |
| 3 | `"AtLeast"` | 최소 |

---

## WidthRel / HeightRel (상대 크기 기준)

### WidthRel
| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Paper"` | 종이 |
| 1 | `"Page"` | 페이지 |
| 2 | `"Column"` | 단 |
| 3 | `"Paragraph"` | 문단 |
| 4 | `"Absolute"` | 절대값 |

### HeightRel
| 값 | COM String | 설명 |
|----|-----------|------|
| 0 | `"Paper"` | 종이 |
| 1 | `"Page"` | 페이지 |
| 2 | `"Absolute"` | 절대값 |

---

## MovePos moveID 값

| 값 | 이름 | 설명 |
|----|------|------|
| 0 | Main | 메인 |
| 1 | CurList | 현재 리스트 |
| 2 | TopOfFile | 문서 시작 |
| 3 | BottomOfFile | 문서 끝 |
| 4 | TopOfList | 리스트 시작 |
| 5 | BottomOfList | 리스트 끝 |
| 6 | StartOfPara | 문단 시작 |
| 7 | EndOfPara | 문단 끝 |
| 8 | StartOfWord | 단어 시작 |
| 9 | EndOfWord | 단어 끝 |
| 10 | NextPara | 다음 문단 |
| 11 | PrevPara | 이전 문단 |
| 12 | NextPos | 다음 위치 |
| 13 | PrevPos | 이전 위치 |
| 14 | NextPosEx | 다음 위치 (확장) |
| 15 | PrevPosEx | 이전 위치 (확장) |
| 16 | NextChar | 다음 글자 |
| 17 | PrevChar | 이전 글자 |
| 18 | NextWord | 다음 단어 |
| 19 | PrevWord | 이전 단어 |
| 20 | NextLine | 다음 줄 |
| 21 | PrevLine | 이전 줄 |
| 22 | StartOfLine | 줄 시작 |
| 23 | EndOfLine | 줄 끝 |
| 24 | ParentList | 부모 리스트 |
| 25 | TopLevelList | 최상위 리스트 |
| 26 | RootList | 루트 리스트 |
| 27 | CurrentCaret | 현재 캐럿 |
| 100 | LeftOfCell | 셀 왼쪽 |
| 101 | RightOfCell | 셀 오른쪽 |
| 102 | UpOfCell | 셀 위 |
| 103 | DownOfCell | 셀 아래 |
| 104 | StartOfCell | 셀 시작 |
| 105 | EndOfCell | 셀 끝 |
| 106 | TopOfCell | 셀 맨 위 |
| 107 | BottomOfCell | 셀 맨 아래 |

---

## InitScan range 플래그

시작 위치 (OR 결합):
| 플래그 | 값 | 설명 |
|--------|------|------|
| scanSposCurrent | 0x0000 | 캐럿부터 |
| scanSposSpecified | 0x0010 | spara/spos 지정 |
| scanSposParagraph | 0x0030 | 문단 시작 |
| scanSposSection | 0x0040 | 구역 시작 |
| scanSposList | 0x0050 | 리스트 시작 |
| scanSposDocument | 0x0070 | 문서 시작 |

끝 위치:
| 플래그 | 값 | 설명 |
|--------|------|------|
| scanEposCurrent | 0x0000 | 캐럿까지 |
| scanEposSpecified | 0x0001 | epara/epos 지정 |
| scanEposParagraph | 0x0003 | 문단 끝 |
| scanEposSection | 0x0004 | 구역 끝 |
| scanEposList | 0x0005 | 리스트 끝 |
| scanEposDocument | 0x0007 | 문서 끝 |

전체 문서 스캔: `0x0070 + 0x0007 = 0x0077`

---

## 출처

- https://docs.rs/hwp/latest/hwp/ (hwp-rs)
- https://github.com/neolord0/hwplib (hwplib Java)
- https://github.com/JunDamin/hwpapi (hwpapi Python)
- https://forum.developer.hancom.com/
- https://developer.hancom.com/hwpautomation
