# 스마트 월간식단 뷰어 · 사이니지 통합본

## 포함 기능
- 나이스 API 학교 검색 및 월간 식단 자동조회
- 월간식단 Excel 업로드
- 월간 전체 식단 카드
- 오늘의 식단 및 오늘 날짜 이동
- 알레르기 1~19 필터
- 선택급식 A/B, 자율배식 지정
- 월 전체 편집 및 찾아 바꾸기
- 학교별·월별 수정본 Supabase 저장
- 당류·나트륨·식이섬유 영양교육 자동 표시
- TV 16:9, 스탠바이미, 키오스크 9:16 사이니지 URL 생성
- 사이니지에는 오늘/월간 식단, 알레르기, 영양정보, 영양교육만 표시

## 1. Supabase 설정
Supabase SQL Editor에서 `supabase.sql` 전체를 실행합니다.

## 2. Vercel 환경변수
Project Settings → Environment Variables에 아래 3개를 등록합니다.
- `NEIS_API_KEY`: 나이스 Open API 인증키
- `SUPABASE_URL`: Supabase Project URL
- `SUPABASE_SERVICE_ROLE_KEY`: Supabase service_role key

주의: service_role key는 브라우저 코드에 넣지 말고 Vercel 환경변수에만 저장합니다.

## 3. 배포
1. 이 폴더 전체를 GitHub 새 저장소에 업로드
2. Vercel → Add New Project → 해당 저장소 Import
3. Framework Preset: Vite
4. Build Command: `npm run build`
5. Output Directory: `dist`
6. 환경변수 등록 후 Deploy

## 수정본 저장
학교·월별 첫 저장 시 편집 비밀번호가 생성됩니다. 이후 같은 비밀번호를 입력해야 수정할 수 있습니다.

## 사이니지 사용
관리 화면에서 사이니지 버튼 → TV/스탠바이미/키오스크 선택 → 전체화면 미리보기 또는 URL 복사.
월을 URL에 포함하므로 해당 월 수정본을 표시합니다.
