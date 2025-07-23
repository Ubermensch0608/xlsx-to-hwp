# xlsx-to-hwp

XLSX 파일의 차트를 이미지로 변환하여 HWP 파일에 삽입하는 서비스입니다.

## 명세

1. 인덱싱된 xlsx 내부 차트를 이미지화하여 hwp파일의 동일한 인덱스를 가진 영역에 삽입한다
2. xlsx와 hwp파일 업로더 UI
3. 차트 이미지가 삽입된 hwp 파일 다운로드 제공

## 기술 스택

- **Backend**: Node.js + Express.js
- **File Processing**:
  - XLSX: `xlsx` 라이브러리
  - Chart to Image: `puppeteer` + `chart.js`
  - HWP: `hwp.js` (기본 구조)
- **File Upload**: `multer`
- **Frontend**: 순수 HTML + CSS + JavaScript

## 설치 및 실행

### 1. 의존성 설치

```bash
npm install
```

### 2. 개발 서버 실행

```bash
npm run dev
```

### 3. 프로덕션 서버 실행

```bash
npm start
```

서버는 `http://localhost:3000`에서 실행됩니다.

## 프로젝트 구조

```
xlsx-to-hwp/
├── server.js              # Express 서버 메인 파일
├── package.json           # 프로젝트 설정 및 의존성
├── public/
│   └── index.html         # 파일 업로드 UI
├── utils/
│   ├── xlsxProcessor.js   # XLSX 파일 처리 유틸리티
│   └── hwpProcessor.js    # HWP 파일 처리 유틸리티
├── uploads/               # 업로드된 파일 저장소
├── output/                # 처리된 파일 출력 디렉토리
└── README.md
```

## 기능

### ✅ 구현 완료

- [x] 기본 Express 서버 설정
- [x] 파일 업로드 UI (XLSX, HWP)
- [x] 파일 업로드 처리 (Multer)
- [x] XLSX 차트 데이터 추출 구조
- [x] 차트 이미지 변환 (Puppeteer + Chart.js)
- [x] HWP 파일 처리 기본 구조

### 🔄 진행 중

- [ ] 실제 HWP 파일 수정 로직 구현
- [ ] 차트 인덱스 매핑
- [ ] 파일 다운로드 기능

### 📋 TODO

- [ ] 에러 처리 개선
- [ ] 파일 크기 제한
- [ ] 보안 검증 강화
- [ ] 로깅 시스템
- [ ] 테스트 코드 작성

## API 엔드포인트

- `GET /` - 메인 페이지 (파일 업로드 UI)
- `POST /upload` - 파일 업로드 처리

## 개발 노트

> - 간단한 HTML
> - xlsx, hwp 내부 수정은 서버 작업이 유리

위와 같은 이유로 초기 단계에서는 server로만 개발

### HWP 파일 처리 참고사항

- HWP 파일은 바이너리 형식이므로 특별한 라이브러리가 필요
- 현재는 기본 구조만 구현되어 있으며, 실제 HWP 수정 로직은 추가 개발 필요
- `pyhwp` 또는 `hwp.js` 라이브러리 활용 고려

## 라이선스

MIT License
