# 📦 Changelog

## [0.47.22] - 2025-03-27
### 📝 Misc
- 검색모듈을 docsearch(Algolia)로 변경함

---

## [0.47.21] - 2025-03-27
### 🐛 Fixed
- 다시 __init__에 from .pyhwpx import * 삽입

---

## [0.47.20] - 2025-03-27
### 🐛 Fixed
- 순환임포트는 pyhwpx.py 안에 있었다. from pyhwpx import Hwp 라인 제거

---

## [0.47.19] - 2025-03-27
### 🐛 Fixed
- 순환임포트 방지를 위한 __init__ 내 지연임포트 코드 추가

---

## [0.47.18] - 2025-03-27
### 📝 Misc
- docstring 구조 개선작업 진행중

---

## [0.47.17] - 2025-03-27
### 📝 Misc
- docstring 구조 개선작업 진행중

---

## [0.47.16] - 2025-03-27
### 📝 Misc
- docstring 구조 개선작업 진행중

---

## [0.47.15] - 2025-03-27
### 📝 Misc
- .gitignore, docstring 경미한 수정

---

## [0.47.14] - 2025-03-27
### 🚀 Added
- MKDocs 문서페이지 추가
- 이번 주 내에 algolia docsearch 기능 추가 예정

---

## [0.47.13] - 2025-03-27
### 🚀 Added
- MKDocs 문서페이지 추가
- 이번 주 내에 algolia docsearch 기능 추가 예정

---

## [0.47.12] - 2025-03-27
### 🚀 Added
- MKDocs 문서페이지 추가
- 이번 주 내에 algolia docsearch 기능 추가 예정

---

## [0.47.11] - 2025-03-27
### 🐛 Fixed
- 기존에 누락되었던 TableSubtractRow 메서드 추가 == hwp.TableSubtractRow()라고 실행할 수 있음. (기존방식 : `hwp.HAction.Run("TableSubtractRow")`)

---

## [0.47.10] - 2025-03-27
### 🐛 Fixed
- 이모티콘(=감성) 추가✨ 챗지피티가 이렇게도 도와주는구나!!!

---

## [0.47.9] - 2025-03-27
### 🐛 Fixed
- 버전 표기가 아직도 안 맞았다. 이번엔 잘 맞겠지!?

---

## [0.47.7] - 2025-03-27
### Fixed
- GitHub Releases 탭의 헤더 버전과 콘텐트 버전에 0.0.1 차이나는 오류 해결

---

## [0.47.6] - 2025-03-27
### Fixed
- GitHub Releases 탭에 릴리즈 자동 생성 기능 추가

---

## [0.47.5] - 2025-03-27
### Fixed
- GitHub Releases 탭에 릴리즈 자동 생성 기능 추가

---

## [0.47.4] - 2025-03-27
### Fixed
- GitHub Releases 탭에 릴리즈 자동 생성 기능 추가

---

## [0.47.3] - 2025-03-27
### Fixed
- 로컬 파이참 터미널에서 chcp 65001 추가. 배포 메시지에 유니코드 아이콘도 추가하고 싶다. 감성이 빠지는 건 싫어!

---

## [0.47.2] - 2025-03-27
### Fixed
- 버전 자동으로 올리는 로직 추가
- CHANGELOG.md 파일 자동작성
- 릴리즈와 커밋푸쉬 분리하기(커밋푸쉬는 해놔도 릴리즈는 좀 더 두고봐야 할 때가 있을 것)

---

## [0.47.1] - 2025-03-27
### Fixed
- 버전 자동으로 올리는 로직 추가
- CHANGELOG.md 파일 자동작성
- 릴리즈와 커밋푸쉬 분리하기(커밋푸쉬는 해놔도 릴리즈는 좀 더 두고봐야 할 때가 있을 것)

---

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
