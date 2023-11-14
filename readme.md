# 파이썬-아래아한글 자동화 모듈

pywin32 패키지를 활용하여 

아래아한글을 다루는 

hwpx 모듈을 만들어보았습니다.

hwpx 모듈 안에는 아직 Hwp라는 클래스만 

하나 정의해놓았습니다.

해당 클래스 안에 HwpAutomation에서 제공하는 

모든 저수준 API메서드의 사용법과 

파라미터, (필요한 경우) 예시코드 등을 

docstring으로 추가하는 중입니다.

틈틈이 유용한 커스텀 메서드를 

추가할 예정입니다.

아직은 오픈소스라고 부르기엔 

다소 초라한 수준이지만,

hwp 문서업무 자동화에 많이 쓰이는 패턴들을 

추가하다보면?

나름 유용한 모듈이 되지 않을까 

생각해봅니다.

두서없이 디스크립션을 적어보았습니다.

행복한 하루 되세요!

일코, 2023. 11. 14.

# 사용법

```python
from hwpx import Hwp

# 아래아한글 시작
# 기존에 hwp문서가 열려 있는 경우 해당 문서를 제어
hwp = Hwp()  

# 텍스트 삽입
hwp.insert_text("Hello world!")

# 다른이름으로 저장
hwp.save_as("c:\\users\\user\\desktop\\hello.hwpx", format="hwpx")

# 한/글 종료
hwp.Quit()
```