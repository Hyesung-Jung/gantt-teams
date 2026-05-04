# 프로젝트 간트 - Teams 배포 가이드

## 📁 폴더 구조
```
gantt-teams/
├── src/
│   ├── main.jsx        ← React 진입점
│   └── App.jsx         ← 간트 앱 본체
├── index.html
├── package.json
├── vite.config.js
└── teams-package/
    ├── manifest.json   ← Teams 앱 설정 (URL 수정 필요)
    ├── color.png       ← 192×192 아이콘
    └── outline.png     ← 32×32 아이콘
```

---

## 🚀 Step 1: GitHub에 올리기

```bash
# 터미널에서 gantt-teams 폴더로 이동 후
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/내아이디/gantt-teams.git
git push -u origin main
```

---

## ▲ Step 2: Vercel에 배포하기

1. [vercel.com](https://vercel.com) 접속 → GitHub으로 로그인
2. **"Add New Project"** 클릭
3. `gantt-teams` 저장소 선택 → **Deploy** 클릭
4. 배포 완료되면 URL 발급: `https://gantt-teams-xxx.vercel.app`

---

## 📝 Step 3: manifest.json URL 수정

`teams-package/manifest.json` 파일에서
`YOUR-APP.vercel.app` 을 실제 Vercel URL로 교체:

```json
"contentUrl": "https://gantt-teams-xxx.vercel.app",
"websiteUrl": "https://gantt-teams-xxx.vercel.app",
"validDomains": ["gantt-teams-xxx.vercel.app"]
```

---

## 📦 Step 4: Teams 패키지 zip 만들기

`teams-package` 폴더 안의 파일 3개를 zip으로 압축:
- manifest.json
- color.png  
- outline.png

> ⚠️ 폴더째로 압축하면 안 되고, **파일 3개만** 선택해서 zip으로!

---

## 🔧 Step 5: Teams에 설치

1. Teams 왼쪽 하단 **앱** 클릭
2. **Manage your apps** 클릭
3. **Upload an app** 클릭
4. zip 파일 선택 → **추가**
5. 왼쪽 사이드바에 앱 아이콘 등장! 🎉
