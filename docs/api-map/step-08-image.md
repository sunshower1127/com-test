# 단계 8: 그림 삽입

## 1. 공식 API 문서 조사

### InsertPicture(Path, Embedded, sizeoption, Reverse, watermark, Effect, Width, Height)

- **Path**: 이미지 파일 경로 (String)
- **Embedded**: 문서에 포함 (1=포함, 0=링크)
- **sizeoption**: 크기 옵션 (0=원본)
- **Reverse**: 반전 (0=없음)
- **watermark**: 워터마크 (0=없음)
- **Effect**: 효과 (0=없음)
- **Width**: 너비 (0=원본)
- **Height**: 높이 (0=원본)
- **반환**: **Dispatch** — 그림 컨트롤 객체 (CtrlID="gso", UserDesc="그림")

### CreatePageImage(Path, pgno, resolution, depth, Format)

- **Path**: 저장 경로
- **pgno**: 페이지 번호 (0-based)
- **resolution**: DPI
- **depth**: 색심도 (24=트루컬러)
- **Format**: 이미지 형식

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 | 실제 | 비고 |
|------|-----------|------|------|
| InsertPicture 반환 | Bool | **Dispatch** | 그림 컨트롤 객체 반환. 문자열 변환 시 에러 주의 |
| InsertPicture 파라미터 | 일부 선택 | **8개 전부 필수** | 파라미터 부족하면 0x8002000E |
| CreatePageImage PNG | 지원 | **파일 생성 안 됨** | ret=true이지만 실제 파일 없음 |
| CreatePageImage JPG | 지원 | **파일 생성 안 됨** | 동일 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)
- 보안모듈 미설치

### InsertPicture

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 8-1 | InsertPicture (8개 파라미터 전부) | ✅ | Dispatch 객체 반환. CtrlID="gso", UserDesc="그림" |
| 8-2 | InsertPicture (파라미터 부족) | ❌ | 0x8002000E — 8개 전부 필수 |
| 8-3 | IsActionEnable("InsertPicture") | false | CreateAction도 empty — HAction 패턴은 안 됨 |

### CreatePageImage

| # | 포맷 | 결과 | 비고 |
|---|------|:----:|------|
| 8-4 | BMP | ✅ | 6.5MB, 정상 |
| 8-5 | GIF | ✅ | 3.7KB, 정상 |
| 8-6 | PNG | ❌ | ret=true이지만 파일 없음 |
| 8-7 | JPG/JPEG | ❌ | 동일 |
| 8-8 | EMF/WMF | ❌ | 동일 |
| 8-9 | TIFF | ❌ | 동일 |

### 반환 컨트롤 객체

InsertPicture가 반환하는 Dispatch 객체의 멤버:

| 멤버 | 종류 | 설명 |
|------|------|------|
| CtrlID | get | "gso" (Graphic Shape Object) |
| UserDesc | get | "그림" |
| CtrlCh | get | 컨트롤 문자 |
| HasList | get | 리스트 포함 여부 |
| Next / Prev | get | 다음/이전 컨트롤 |
| Properties | get/put | 속성 ParameterSet |
| GetAnchorPos | method | 앵커 위치 |

---

## 4. 분석

### 핵심 발견 1: InsertPicture는 Dispatch를 반환한다

```js
// ✅ 올바른 사용법 — 반환값을 Dispatch로 받기
var ctrl = hwp.InsertPicture(path, 1, 0, 0, 0, 0, 0, 0);
// ctrl.CtrlID = "gso"
// ctrl.Properties 로 속성 접근 가능

// ❌ 에러 — 반환값을 문자열로 쓰면 안 됨
console.log("result=" + hwp.InsertPicture(...));  // Cannot convert object to primitive value
```

파라미터 8개가 전부 필수. 부족하면 에러.

### 핵심 발견 2: CreatePageImage는 BMP/GIF만 작동

PNG, JPG, EMF 등은 `ret=true`를 반환하지만 실제 파일이 생성되지 않음.
BMP (무압축)와 GIF만 정상 작동. 보안모듈 또는 한글 2018 제한일 수 있음.

---

## 5. 최종 결론

### 정상 작동 (✅)

- **InsertPicture** — 8개 파라미터 전부 넘기면 정상. Dispatch 반환
- **CreatePageImage("BMP")** — 페이지를 BMP로 저장
- **CreatePageImage("GIF")** — 페이지를 GIF로 저장

### 주의 필요 (⚠️)

- **InsertPicture 반환값** — Dispatch 객체. 문자열 변환 시 에러
- **파라미터 8개 필수** — 생략 불가

### 이미지 소스별 결과

| 소스 | 결과 | 비고 |
|------|:----:|------|
| **로컬 파일 경로** | ✅ | JPG, BMP 등 |
| **HTTP/HTTPS URL** | ❌ | 객체 반환되지만 이미지 로드 안 됨 (빈 그림) |
| **존재하지 않는 파일** | ❌ | 객체 반환되지만 빈 그림 |
| **Buffer → 임시파일** | ✅ | 프로그래밍으로 생성한 이미지 삽입 가능 |
| **SetTextFile HTML img** | ❌ | 인코딩 팝업 뜸, 이미지 안 보임 |
| **SetTextFile HTML base64** | ❌ | 이미지 안 보임 |
| **Paste (클립보드)** | ? | 클립보드에 이미지 있으면 가능할 수 있음 (미확인) |

**외부 이미지 삽입 패턴:**
```js
// URL/blob → 임시파일로 저장 → InsertPicture
var tmpPath = "C:/tmp/image.bmp";
fs.writeFileSync(tmpPath, imageBuffer);
hwp.InsertPicture(tmpPath, 1, 0, 0, 0, 0, 0, 0);
```

### 작동하지 않음 (❌)

- **CreatePageImage PNG/JPG/EMF/WMF/TIFF** — ret=true이지만 파일 생성 안 됨
- **IsActionEnable("InsertPicture")** — false. HAction 패턴 사용 불가
- **InsertPicture URL** — 객체 반환되지만 실제 이미지 로드 안 됨
- **SetTextFile HTML img/base64** — 이미지 삽입 안 됨
