# 📦 Changelog

## [1.6.5] - 2025-09-17
### 🐛 Fixed
- hwp.get_selected_pos() 메서드 실행시 낮은 버전에서도 안정적으로 실행되게 하기 위해, 내부적으로 BySet 메서드로 대체함

---
## [1.6.4] - 2025-09-16
### 🐛 Fixed
- hwp.is_empty_page 내부, get_text_file 실패시 Undo 추가

---
## [1.6.3] - 2025-09-16
### 📝 Misc
- hwp.is_empty_page 메서드 내부에서 GetTextFile 실패시 True 바로 리턴

---
## [1.6.2] - 2025-09-16
### 📝 Misc
- hwp.is_empty_page 메서드의 ignore_fwspace(고정폭빈칸 무시)도 True로 변경

---
## [1.6.1] - 2025-09-16
### 🐛 Fixed
- hwp.is_empty_page 내부 Undo 로직 변경

---
## [1.6.0] - 2025-09-16
### 🚀 Added
- hwp.is_empty_page 메서드 추가(보완 필요. 사용 중 오류 발생시 제보 바람)

---
## [1.5.6] - 2025-09-01
### 📝 Misc
- hwp.find_replace 메서드 내부 get_text_file에 `option=""` 추가

---
## [1.5.5] - 2025-08-22
### 📝 Misc
- hwp.ShapeObjUngroup 으로 메서드 이름 수정(G -> g)

---
## [1.5.4] - 2025-08-22
### 📝 Misc
- hwp.ShapeObjUnGroup 메서드 추가(개체 묶기 해제)

---
## [1.5.3] - 2025-08-21
### 🐛 Fixed
- hwp.SelectCtrl 및 hwp.select_ctrl 메서드 실행시 한글2024 이후 버전의 경우 다중선택 또는 선택추가 가능
- 기존 컨트롤인스턴스 아이디(digit문자열)를 입력하는 방식 대신 Ctrl 인스턴스나 List[Ctrl]를 넣어도 작동하도록 변경

---
## [1.5.2] - 2025-08-21
### 📝 Misc
- hwp.ShapeObjGroup 메서드 실행시 뜨는 팝업 무시(항상 "예" 선택)


---
## [1.5.1] - 2025-08-20
### 📝 Misc
- hwp.find_replace_all 안의 get_text_file의 option에서 saveblock:true 제거


---
## [1.5.0] - 2025-08-12
### 🚀 Added
- hwp.insert_hyperlink 메서드 추가(현재 선택한 문자열 구간에 하이퍼링크를 삽입할 수 있음)

### 🐛 Fixed
- hwp.ctrl_list의 " 끝" 컨트롤 삭제시 com_error 발생 대신 예외처리함 

---
## [1.4.11] - 2025-08-03
### 📝 Misc
- hwp.clipboard_to_pyfunc 메서드 경미한 정규식 오류 정정

---
## [1.4.10] - 2025-08-03
### 📝 Misc
- hwp.clipboard_to_pyfunc 메서드 경미한 수정(FindCtrl() 메서드 앞에 `hwp.` 추가)

---
## [1.4.9] - 2025-08-03
### 📝 Misc
- hwp.clipboard_to_pyfunc 메서드 경미한 수정(FindCtrl() 메서드 앞에 `hwp.` 추가)

---
## [1.4.8] - 2025-08-03
### 📝 Misc
- 헬퍼함수 hwp.PicEffect의 인수 설명 추가

---
## [1.4.7] - 2025-07-31
### 🐛 Fixed
- 하찮은 메서드, get_ctrl_by_ctrl_id를 추가했다. 표나 그림 컨트롤 목록을 더 간편하게 얻어오기 위함. 추후에 여러 종류를 가져오거나, UserDesc로 가져오는 메서드로 보완 예정

---
## [1.4.6] - 2025-07-31
### 🐛 Fixed
- 하찮은 메서드, get_ctrl_by_ctrl_id를 추가했다. 표나 그림 컨트롤 목록을 더 간편하게 얻어오기 위함. 추후에 여러 종류를 가져오거나, UserDesc로 가져오는 메서드로 보완 예정

---
## [1.4.5] - 2025-07-31
### 🐛 Fixed
- 하찮은 메서드, get_ctrl_by_ctrl_id를 추가했다. 표나 그림 컨트롤 목록을 더 간편하게 얻어오기 위함. 추후에 여러 종류를 가져오거나, UserDesc로 가져오는 메서드로 보완 예정

---
## [1.4.4] - 2025-07-31
### 🐛 Fixed
- 하찮은 메서드, get_ctrl_by_ctrl_id를 추가했다. 표나 그림 컨트롤 목록을 더 간편하게 얻어오기 위함. 추후에 여러 종류를 가져오거나, UserDesc로 가져오는 메서드로 보완 예정

---
## [1.4.3] - 2025-07-31
### 📝 Misc
- 하찮은 메서드, get_ctrl_by_ctrl_id를 추가했다. 표나 그림 컨트롤 목록을 더 간편하게 얻어오기 위함. 추후에 여러 종류를 가져오거나, UserDesc로 가져오는 메서드로 보완 예정

---
## [1.4.2] - 2025-07-31
### 📝 Misc
- 하찮은 메서드, get_ctrl_by_ctrl_id를 추가했다. 표나 그림 컨트롤 목록을 더 간편하게 얻어오기 위함. 추후에 여러 종류를 가져오거나, UserDesc로 가져오는 메서드로 보완 예정

---
## [1.4.1] - 2025-07-31
### 🐛 Fixed
- `hwp.get_pos()`의 리턴타입 힌트를 `Tuple[int, int, int]`로 정정

---
## [1.4.0] - 2025-07-30
### 🚀 Added
- hwp.get_viewstate, hwp.set_viewstate 메서드 추가 : 조판부호 보기 설정이 필요한 경우가 많아 추가하게 됨.

---
## [1.3.5] - 2025-07-30
### 🐛 Fixed
- delete_ctrl 메서드 오류 정정(삭제시 Ctrl 객체 대신 Ctrl의 _com_obj로 삭제 진행)
- HeadCtrl, LastCtrl에서 Next나 Prev시에 Optional 타입체크로 인한 오류 제거

---
## [1.3.4] - 2025-06-24
### 🐛 Fixed
- switch_to 메서드의 오타 정정(hwp -> self.hwp). 무식하면 버전이 빨리 올라간다!

---
## [1.3.3] - 2025-06-19
### 🐛 Fixed
- switch_to 메서드의 입력정수 검증방법을 기존방법인 Count로 변경함

---
## [1.3.2] - 2025-06-19
### 🐛 Fixed
- Ctrl 클래스의 __repr__ 및 Next / Prev 속성 수정(ctrl = hwp.HeadCtrl.Prev 시에 오류 없이 None 리턴)

- Run 단축메서드 주석 경미한 수정

---
## [1.3.1] - 2025-06-17
### 📝 Misc
- deveworld님의 기여내용 반영

---
## [1.3.0] - 2025-06-17
### 🚀 Added
- 변경내용 추적 등 202504 매뉴얼에 추가된 Run 메서드 추가(일부는 2022, 2024에서부터 지원)

---
## [1.2.8] - 2025-06-12
### 🐛 Fixed
- find 관련 모든 메서드의 초기화 코드 추가(FindDlg)

- 오늘 알게 된 건데, find의 정수 파라미터를 무시하는 방법은 65535를 대입하는 것.

---
## [1.2.7] - 2025-06-12
### 🐛 Fixed
- Hwp 클래스 내 get_selected_text의 keep_select 파라미터를 모두 True로 변경

---
## [1.2.6] - 2025-06-12
### 🐛 Fixed
- set_field_by_bracket 수정(get_selected_text 변경으로 인한 버그 제거)

---
## [1.2.5] - 2025-06-12
### 🐛 Fixed
- is_empty_para 메서드 보완(로직 변경)

---
## [1.2.4] - 2025-06-12
### 📝 Misc
- find 메서드에 글자크기(Height) 속성 파라미터 소심하게 추가해봄

---
## [1.2.3] - 2025-06-10
### 📝 Misc
- hwp.auto_spacing 글자선택 취소되는 버그 개선

- hwp.get_selected_text 메서드에 keep_select 파라미터 추가(기본값은 False로 선택해제임)

---
## [1.2.2] - 2025-06-10
### 📝 Misc
- hwp.get_selected_pos 메서드, hwp.select_text 메서드 독스트링 보완(특정 구간의 선택상태를 저장하고, 해당 저장블록을 재선택하는 데 쓸 수 있는 유용한 메서드니까!)

---
## [1.2.1] - 2025-06-10
### 🐛 Fixed
- hwp.find 메서드에 글자색(TextColor) 파라미터만 추가해봄(기본값은 "Black"이 아니라 None임). 적용방법은 hwp.find(TextColor=hwp.RGBColor("Blue"))

- hwp.find 메서드의 SeveralWords(콤마로 or연산 검색) 파라미터 기본값을 1에서 0으로 변경.

- hwp.find 메서드의 FindString(파라미터명은 src)를 필수파라미터에서, 빈 문자열("")로 기본값 추가. 특정 글자색 구간을 찾을 때 찾을 문자열을 비울 수 있게 함. 

- hwp.find 반복실행시, 초기화를 위해 메서드 상단에 FindDlg 액션 (파라미터셋은 비우고) 실행

---
## [1.2.0] - 2025-06-09
### 🚀 Added
- apply_parashape 메서드 추가. hwp.get_parashape_as_dict 메서드로 저장한 문단속성 딕셔너리를 다른 문단에 간편하게 적용할 수 있는 메서드. para_dict의 일부 키를 삭제(pop)하여 문단모양 적용범위를 조절할 수 있다는 점에서 get_parashape / set_parashape과 차이가 있다. 

---
## [1.1.7] - 2025-06-09
### 🐛 Fixed
- set_para 메서드의 AlignType 파라미터에 문자열값이 아닌 정수 입력시, 헬퍼함수를 거치지 않고 바로 파라미터셋에 넣을 수 있도록 보완

---
## [1.1.6] - 2025-06-09
### 🐛 Fixed
- set_col_width에 self.Cancel() 추가(리스트로 칼럼별 비율 넣어 실행하는 경우 너비가 동일하게 조절되는 버그 조치)

- 가끔 이런 버그가 생기면 현타가 올 때가 있다. 내가 꼼꼼하게 함수를 짜놔도, 한글 개발자들이 API실행방식을 바꾸면 pyhwpx는 버그덩어리가 된다. 예전엔 Execute 실행시 Cancel이 자동으로 실행되는 방식이어서 오류가 안 났던 것 같기도 하다. 사실 이런 버그가 자주 있는 일은 아니다. 한글 개발자들이 열일해서겠지.

---
## [1.1.5] - 2025-06-01
### 🐛 Fixed
- switch_to 오류 수정(내부에서 사용하는 FindItem 메서드는 0-index가 아니고 1-index임!ㅜㅜ)


---
## [1.1.4] - 2025-05-28
### 📝 Misc
- 문단모양을 저장하거나 적용할 수 있는 get_parashape, get_parashape_as_dict, set_parashape 3종세트!

- 그리고 만들어놓고 존재조차 까맣게 잊고 있었던 get_charshape, get_charshape_as_dict, set_charshape도 이번 용역을 통해.. 기억남ㅜ

- **기존 함수방식의 문단모양 변경 메서드인 set_paragraph는 set_para로 이름을 변경함ㅜㅜㅜ**  

---
## [1.1.3] - 2025-05-28
### 🐛 Fixed
- quit 메서드 내부에서 clear 동작 추가 

---
## [1.1.2] - 2025-05-28
### 🐛 Fixed
- set_parashape 메서드에 self가 빠져있었던 부분 정정

- HeadType 헬퍼메서드에 독스트링 추가 

---
## [1.1.1] - 2025-05-28
### 🐛 Fixed
- get_style_dict 메서드가 리턴하는 dict의 키 문자열을 HWPML2X와 일치시킴. 

---
## [1.1.0] - 2025-05-28
### 🚀 Added
- pyhwpx 설치 가능한 파이썬 최소버전을 3.10에서 3.9로 낮춤(버전 오기로 배포 재시도!ㅋ) 

---
## [1.0.17] - 2025-05-28
### minor
- pyhwpx 설치 가능한 파이썬 최소버전을 3.10에서 3.9로 낮춤 

---
## [1.0.16] - 2025-05-28
### minor
- set_parashape 메서드 추가. set_font와 유사하게 문단의 모양을 단순한 메서드 방식으로 변경할 수 있다.

- 위에 수반되는 get_parashape, get_parashape_as_dict 메서드도 추가했지만, 고급사용자 외에는 쓸 일이 없을 듯. 

---
## [1.0.15] - 2025-05-27
### minor
- set_parashape 메서드 추가. set_font와 유사하게 문단의 모양을 단순한 메서드 방식으로 변경할 수 있다.

- 위에 수반되는 get_parashape, get_parashape_as_dict 메서드도 추가했지만, 고급사용자 외에는 쓸 일이 없을 듯. 

---
## [1.0.14] - 2025-05-23
### patch
- hwp.get_selected_text()를 본문에서 사용 후 선택모드 해제(Cancel) 명령 추가

---
## [1.0.13] - 2025-05-23
### patch
- hwp.get_selected_text()를 본문에서 사용 후 선택모드 해제(Cancel) 명령 추가

---
## [1.0.12] - 2025-05-23
### minor
- delete_style_by_name 메서드에 삭제하고 싶은 스타일리스트를 넣어 일괄삭제할 수 있게 수정
- get_used_style_dict 메서드 추가(전체 스타일 중 문서에 실제 사용된 스타일 목록만 추출)
- remove_unused_styles 메서드 추가(실제 사용되지 않은 스타일들은 문서에서 제거하는 기능)

---
## [1.0.11] - 2025-05-23
### minor
- delete_style_by_name 메서드에 삭제하고 싶은 스타일리스트를 넣어 일괄삭제할 수 있게 수정
- get_used_style_dict 메서드 추가(전체 스타일 중 문서에 실제 사용된 스타일 목록만 추출)
- remove_unused_styles 메서드 추가(실제 사용되지 않은 스타일들은 문서에서 제거하는 기능)

---
## [1.0.10] - 2025-05-23
### minor
- delete_style_by_name 메서드에 삭제하고 싶은 스타일리스트를 넣어 일괄삭제할 수 있게 수정
- get_used_style_dict 메서드 추가(전체 스타일 중 문서에 실제 사용된 스타일 목록만 추출)
- remove_unused_styles 메서드 추가(실제 사용되지 않은 스타일들은 문서에서 제거하는 기능)

---
## [1.0.9] - 2025-05-23
### minor
- delete_style_by_name 메서드에 삭제하고 싶은 스타일리스트를 넣어 일괄삭제할 수 있게 수정
- get_used_style_dict 메서드 추가(전체 스타일 중 문서에 실제 사용된 스타일 목록만 추출)
- remove_unused_styles 메서드 추가(실제 사용되지 않은 스타일들은 문서에서 제거하는 기능)

---
## [1.0.8] - 2025-05-19
### patch
- SetMessageBoxMode 수정 : 급한 업데이트 중에 엉뚱한 함수가 들어가 있었음ㅜ 

---
## [1.0.7] - 2025-05-09
### default
- self.on_quit = on_quit 라인 추가ㅜ 

---
## [1.0.6] - 2025-05-09
### default
- __init__ 생성자의 quit 파라미터 -> on_quit 으로 이름 변경 

---
## [1.0.5] - 2025-05-08
### default
- fonts -> htf_fonts 

---
## [1.0.4] - 2025-05-08
### default
- del_on_quit -> self.del_on_quit 

---
## [1.0.3] - 2025-05-08
### default
- 폰트 사전을 별도 파일(fonts.py)로 분리

---
## [1.0.2] - 2025-04-28
### patch
- hwp.goto_addr() 실행시 표안의표 무시하기 패치
- 내부 헬퍼함수 tuple_to_addr의 열,행 입력순서를 행,열로 정상화

---
## [1.0.1] - 2025-04-28
### patch
- hwp.goto_addr() 실행시 표 안에 각주/미주 등의 컨트롤이 있는 경우 무한루프 처리
- todo: 표 안의 표를 무시하는 방법...ㅜ 

---
## [1.0.0] - 2025-04-27
### 💥 Breaking
- 1.0.0으로 업데이트(core, param_helpers, run_methods로 분리)
- 공식문서 카테고리 구분 시작 (앞길이 벌써 막막함ㅜ) 

---
## [0.51.2] - 2025-04-13
### 🐛 Fixed
- quit, Quit 메서드 실행시 빈 문서일 경우 clear 추가함(hwp.hwp 삭제 -> 삭제안함) 

---
## [0.51.1] - 2025-04-13
### 🐛 Fixed
- quit, Quit 메서드 실행시 저장여부 결정하는 save 파라미터 추가(단톡방 엄지척 제이지님 제보) 

---
## [0.51.0] - 2025-04-13
### 🚀 Added
- XHwpDocuments, XHwpDocument 오브젝트를 클래스로 추가 

---
## [0.50.40] - 2025-04-13
### Added
- XHwpDocuments, XHwpDocument 클래스 추가 

---

## [0.50.39] - 2025-04-13
### 🐛 Fixed
- 보안모듈 관련 RegisterModule, check_registry_key 등의 모듈이름 "FilePathCheckerModule"을 하드코딩 대신 파라미터로 옮김. (Ruzzy77님께서 기여해주심) 

---

### 📝 Misc
- Hwp 클래스의 __repr__ 출력 개선

---

## [0.50.38] - 2025-04-11
### 🐛 Fixed
- 믹스인? 상속? 뭘 펴서 공부해야 하나... 

---

## [0.50.37] - 2025-04-11
### 🐛 Fixed
- 타이핑이 문제가 되네.. str 또는 None을 리턴한다고 -> str|None: 을 쓰면 안 되는구나ㅜ 

---

## [0.50.36] - 2025-04-11
### 🐛 Fixed
- README.md 움짤 주소 변경 

---

## [0.50.35] - 2025-04-11
### 🐛 Fixed
- hwp.find_dir -> hwp.FindDir 오류 해결 

---

## [0.50.34] - 2025-04-11
### 📝 Misc
- 독스트링 업데이트(`:param` 요소 전부 제거하고 구글 독스트링 스타일의 `Args:`로 교체함) 

---

## [0.50.33] - 2025-04-11
### 📝 Misc
- 독스트링 업데이트(너무 피곤하다..ㅜ 오늘은 여기까지만) 

---

## [0.50.32] - 2025-04-11
### 📝 Misc
- Examples 섹션의 코드블록을 >>> 대신 ```로 바꿔봄 

---

## [0.50.31] - 2025-04-11
### 📝 Misc
- Examples 섹션의 코드블록을 >>> 대신 ```로 바꿔봄 

---

## [0.50.30] - 2025-04-11
### 📝 Misc
- Examples 섹션의 코드블록을 >>> 대신 ```로 바꿔봄 

---

## [0.50.29] - 2025-04-11
### 📝 Misc
- Examples 섹션의 코드블록을 >>> 대신 ```로 바꿔봄 

---

## [0.50.28] - 2025-04-11
### 📝 Misc
- gradation_on_cell 독스트링(움짤) 업데이트 

---

## [0.50.27] - 2025-04-11
### 📝 Misc
- SelectCtrl 등 독스트링 업데이트 

---

## [0.50.26] - 2025-04-11
### 📝 Misc
- get_field_info로 독스트링 테스트중.. 이런 와중에도 버전은 오른다ㅜ 

---

## [0.50.25] - 2025-04-11
### 📝 Misc
- 전체 독스트링의 \r\n과 \t 등 탈출문자열 정리 

---

## [0.50.24] - 2025-04-11
### 📝 Misc
- rgb_color 및 독스트링에 \r\n 문자열 정비 

---

## [0.50.23] - 2025-04-10
### 📝 Misc
- get_linespacing, NewNumber 독스트링 수정 

---

## [0.50.22] - 2025-04-10
### 📝 Misc
- shape_copy_paste, goto_style 등 독스트링 수정 

---

## [0.50.21] - 2025-04-10
### 📝 Misc
- set_font 독스트링 수정 

---

## [0.50.20] - 2025-04-10
### 📝 Misc
- set_font 독스트링 수정 

---

## [0.50.19] - 2025-04-09
### 📝 Misc
- 독스트링의 헤딩 크기를 와장창 키워봄
- 경미한 독스트링, 어노테이션 업데이트 

---

## [0.50.18] - 2025-04-09
### 📝 Misc
- 독스트링에 gif 삽입해봄ㅎ 

---

## [0.50.17] - 2025-04-09
### 📝 Misc
- API문서에 assets 이미지 삽입 가능한가? 로컬에서는 되는데ㅜ

---

## [0.50.16] - 2025-04-09
### 🐛 Fixed
- set_pos 파라미터 list를 List로 변경

---

## [0.50.15] - 2025-04-09
### 📝 Misc
- 독스트링 및 타입어노테이션 추가

---

## [0.50.14] - 2025-04-08
### 🐛 Fixed
- set_table_outside_margin 메서드 계산방식 변경
- HwpUnitToMili의 리턴값을 소숫점 네 자리에서, round(,2)로 변경

---

## [0.50.13] - 2025-04-08
### 🐛 Fixed
- get_table_outside_margin 및 _left, _right, _top, _bottom 메서드의 계산방식 변경. 표 안에 꽉 찬 그림이 있을 때 발생하는 오류 제거

---

## [0.50.12] - 2025-04-08
### 🐛 Fixed
- Ctrl 클래스의 Properties 속성의 세터 데코레이터 추가. (오류를 직접 겪어야 업데이트를 하냐ㅜ)

---

## [0.50.11] - 2025-04-08
### 🐛 Fixed
- 표 안에서 insert_picture 실행시, 전역셀 안여백과 특정셀 안여백 적용에 따라 이미지 너비/높이 자동조절

---

## [0.50.10] - 2025-04-06
### 🐛 Fixed
- Ctrl 클래스의 Next, Prev 리턴의 래핑 추가
- get_text_file에 블록저장 옵션을 기본값으로 변경

---

## [0.50.9] - 2025-04-06
### 🚀 Added
- Ctrl 래퍼클래스 추가(docstring 등을 위함). 단계적으로 클래스 하나씩 쪼개서 래핑할 예정

---

### 🐛 Fixed
- 스네이크 케이스의 헬퍼함수 일부 제거함

---

### 📝 Misc
- docstring 1차 완성

---

## [0.50.8] - 2025-04-06
### 🚀 Added
- Ctrl 래퍼클래스 추가(docstring 등을 위함). 단계적으로 클래스 하나씩 쪼개서 래핑할 예정

---

### 🐛 Fixed
- 스네이크 케이스의 헬퍼함수 일부 제거함

---

### 📝 Misc
- docstring 1차 완성

---

## [0.50.7] - 2025-04-06
### 🐛 Fixed
- self.switch_to 실행시, 인덱스를 넘었을 때 com_error 대신 IndexError를 리턴 

---

### 📝 Misc
- 일부 static 방식 메서드에 @staticmethod 데코레이터 추가

---

## [0.50.6] - 2025-04-06
### 🐛 Fixed
- self.Version을 문자열에서 리스트로 변경함
- self.Version을 사용하는 모든 메서드에서 .split을 제거함
- 그 외 경미한 독스트링 추가

---

## [0.50.5] - 2025-04-06
### 🚀 Added
- Title 프로퍼티 추가(제목표시줄의 타이틀임)

---

### 🐛 Fixed
- 기존 __repr__ 을 Title 문자열로 교체함(앞으로는 굳이 hwp.Path를 실행해보지 말고, hwp만 실행해보면 됨)
- 외부에서 불러오던 fonts.json 파일 때문에 pyinstaller 커맨드가 길어짐.. 사용자 편의를 위해 fonts.json을 다시 pyhwpx.Hwp의 __init__에 삽입!

---

## [0.50.4] - 2025-04-06
### 🚀 Added
- Title 프로퍼티 추가(제목표시줄의 타이틀임)

---

### 🐛 Fixed
- 기존 __repr__ 을 Title 문자열로 교체함(앞으로는 굳이 hwp.Path를 실행해보지 말고, hwp만 실행해보면 됨)
- 외부에서 불러오던 fonts.json 파일 때문에 pyinstaller 커맨드가 길어짐.. 사용자 편의를 위해 fonts.json을 다시 pyhwpx.Hwp의 __init__에 삽입!

---

## [0.50.3] - 2025-04-06
### 🐛 Fixed
- Title 프로퍼티 추가
- __repr__ 을 Title 리턴으로 변경
- fonts.json을 도로 pyhwpx.py 안에 딕셔너리로 넣음(컴파일하는 사용자가 불편함) 

---

## [0.50.2] - 2025-04-03
### 🐛 Fixed
- set_table_width 실행시 용지 방향에 따른 계산방식 변경(기존 오류 해결)
- get/set_table_width 실행시 캐럿이 표 안에 있지 않으면 오류 발생(AttributeError 같은 애매한 거 말고ㅜ 이런 작업 많이 해야겠다) 

---

## [0.50.1] - 2025-04-03
### 🐛 Fixed
- save_as 메서드에서 ".hwpx"로 저장할 때 포맷을 명시하지 않아도 내부적으로 "HWPX"로 처리하는 elif문 추가
- 버전 폭주 버그! 이러다 아무 것도 안 하고 조만간 1.0.0이 나와버릴 듯...ㅜ 자동버저닝 버그는 일요일에 꼭 고치자. 

---

## [0.50.0] - 2025-04-01
### 🐛 Fixed
- 항상 헷갈린다. 예시코드는 Example: 이 아니고 Examples:인데ㅜ Args, Returns, Examples 전부 s로 끝난다고 외워야겠다. 

---

## [0.49.0] - 2025-04-01
### 🚀 Added
- set_cur_field_name 독스트링 업데이트. 버전업이 뭔가 이상하다. Added 로그 다음 버전에 마이너가 올라가는 듯? 

---

## [0.48.0] - 2025-04-01
### 🚀 Added
- 독스트링 소폭 수정. 마이너 버전업 

---

## [0.47.47] - 2025-04-01
### 🐛 Fixed
- hwp.get_title 메서드 Added : 빈 문서 상태라도 "빈 문서 1 - 한글"이라는 타이틀 문자열을 리턴함. (아무 짝에도 쓸모 없어 보이지만 누군가는 필요로 할 수도 있을 듯ㅎ) 

---

## [0.47.46] - 2025-03-30
### 🐛 Fixed
- hwp.set_cur_field_name 셀블록 상태시에도, 아닌 경우에도 필드명 수정 되게~~

---

## [0.47.45] - 2025-03-29
### 🐛 Fixed
- docstring 개선(Example: -> Examples:)

---

## [0.47.44] - 2025-03-29
### 🐛 Fixed
- docstring 개선, release.py 업데이트, api.md 우측 TOC 숨김해제 등

---

## [0.47.43] - 2025-03-29
### 📝 Misc
- 독스트링 업데이트, api.md 화면 우측 목차 숨김해제(버전 마이너 번호를 올리고 싶은데 어떻게 하지???)

---

## [0.47.42] - 2025-03-28
### 🐛 Fixed📝 Misc
- 리턴 노테이션 오류가 있었어ㄷㄷㄷ

---

## [0.47.41] - 2025-03-27
### 📝 Misc
- 검색엔진을 algolia로 교체했지만.. (결과는 같을 것 같을 것 같다ㅋ)

---

## [0.47.40] - 2025-03-27
### 📝 Misc
- 컨트리뷰터스 수동나열을 빼봤는데.. 왜 용준님은 안 나오지?

---

## [0.47.39] - 2025-03-27
### 📝 Misc
- Hwp 클래스의 독스트링은 살리는 대신, 소스코드 제공을 뺐더니 검색이 좋아졌다!

---

## [0.47.38] - 2025-03-27
### 📝 Misc
- Hwp 클래스의 독스트링을 제거하면 어떨까?

---

## [0.47.37] - 2025-03-27
### 📝 Misc
- 로고도 추가했고, 커미터스도 추가된 듯 하다!

---

## [0.47.36] - 2025-03-27
### 📝 Misc
- 로고도 추가했고, 커미터스도 추가된 듯 하다!

---

## [0.47.35] - 2025-03-27
### 📝 Misc
- 로고도 추가했고, 커미터스도 추가된 듯 하다!

---

## [0.47.34] - 2025-03-27
### 📝 Misc
- git-committers 활성화를 위해 환경변수에 APIKEY 저장! 제발돼라!1

---

## [0.47.33] - 2025-03-27
### 📝 Misc
- git-committers 활성화를 위해 환경변수에 APIKEY 저장! 제발돼라!1

---

## [0.47.32] - 2025-03-27
### 📝 Misc
- git-committers 활성화를 위해 환경변수에 APIKEY 저장! 제발돼라!1

---

## [0.47.31] - 2025-03-27
### 📝 Misc
- git-committers, git-revision-date-localized-plugin 추가해봄. (뭔지는 열어봐야 알겠지)

---

## [0.47.30] - 2025-03-27
### 📝 Misc
- 웹문서에 피드백 기능 (조그맣게) 추가

---

## [0.47.29] - 2025-03-27
### 🚀 Added
- 웹문서에 설문조사 추가

---

## [0.47.28] - 2025-03-27
### 📝 Misc
- 웹문서에 기여 및 댓글 추가

---

## [0.47.27] - 2025-03-27
### 📝 Misc
- 웹문서에 기여 및 댓글 추가

---

## [0.47.26] - 2025-03-27
### 📝 Misc
- 웹문서에 기여 및 댓글 추가

---

## [0.47.25] - 2025-03-27
### 📝 Misc
- 웹문서에 기여 및 댓글 추가

---

## [0.47.24] - 2025-03-27
### 📝 Misc
- 웹문서에 기여 및 댓글 추가

---

## [0.47.23] - 2025-03-27
### 📝 Misc
- 웹문서에 구글태그 붙이기

---

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
