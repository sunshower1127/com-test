# 단계 10: 페이지 설정

## 테스트 결과

모든 테스트 ✅ 통과.

### 사용법

```js
hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet);
var pd = hwp.HParameterSet.HSecDef.PageDef;

// 용지 크기
pd.PaperWidth = 59528;   // A4 210mm
pd.PaperHeight = 84186;  // A4 297mm

// 여백
pd.LeftMargin = 8504;    // 30mm
pd.RightMargin = 8504;
pd.TopMargin = 5668;     // 20mm
pd.BottomMargin = 4252;  // 15mm

// 용지 방향
pd.Landscape = 0;  // 0=세로, 1=가로

hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet);
```

### 주의사항

- **GetDefault가 전부 0으로 리셋** — 변경할 프로퍼티뿐 아니라 **모든 프로퍼티를 명시적으로 설정**해야 함 (PaperWidth, PaperHeight, 여백 전부)
- HWP 단위: 1mm ≈ 283.46 HWP unit (MiliToHwpUnit으로 변환 가능)

### 주요 용지 크기

| 용지 | Width | Height |
|------|-------|--------|
| A4 | 59528 | 84186 |
| B5 | 51590 | 72852 |
| Letter | 61200 | 79200 |
