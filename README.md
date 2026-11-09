# SlideCraft Vercel 배포 수정판

이 압축파일은 **폴더 없이 9개의 파일이 최상위에 있는 배포용**입니다.

## GitHub 업로드
1. ZIP 압축을 풉니다.
2. GitHub 저장소 최상위에서 Add file > Upload files.
3. 압축 해제된 파일들을 **전부 선택해** 업로드합니다. 기존 index.html, package.json, vercel.json과 같은 이름이 있으면 덮어씁니다.
4. 기존 src, tools 폴더는 남아 있어도 새 앱은 사용하지 않습니다.
5. GitHub 최상위에 index.html, main.js, converter.js, tableEditor.js, style.css, package.json, vercel.json이 모두 보여야 합니다.

## Vercel 설정
- Framework Preset: Vite
- Root Directory: ./ (저장소 최상위)
- Build Command: npm run build
- Output Directory: dist
- 기존 연결 프로젝트라면 GitHub 커밋 이후 자동 배포됩니다.

## 로컬 실행
```
npm install
npm run build
npm run dev
```

브라우저 OCR의 정확도와 대용량 파일 성능은 PDF마다 달라집니다.

### 변경사항
- `index.html`에서 존재하지 않던 `/src/main.js` 대신 최상위 `./main.js` 참조
- JS/CSS 파일도 모두 같은 폴더에 배치해 GitHub 웹 업로드 시 폴더 계층 누락 방지
- 코드의 상대 import 경로를 그대로 유지
