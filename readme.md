# 파이썬-아래아한글 자동화 모듈

pywin32 패키지를 활용하여 

아래아한글을 다루는 

pyhwpx 모듈을 만들어보았습니다.

pyhwpx 모듈 안에는 아직 Hwp라는 클래스만 

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

일코, 2023. 11. 30.

# 사용법

```python
from pyhwpx import *

# 아래아한글이 바로 실행되며, hwp, hwpx 등 두 개의 인스턴스 자동생성됨.
# 편의를 위한 커스텀 인스턴스는 hwpx로, 
# 기존 win32com (로우레벨) 인스턴스는 hwp로 생성됨.
# (직관적이지 않은 부분이 있음을 인정함..ㅜ)
# (한/글 스크립트매크로와의 호환을 위해 어쩔 수 없었음. 향후 개선 예정)
# 기존에 한/글 프로그램이 열려 있는 경우 해당 문서를 제어

# 텍스트 삽입
hwpx.insert_text("Hello world!")

# 위의 코드는 아래처럼 hwp 인스턴스로도 실행 가능
pset = hwp.HParameterSet.HInsertText
hwp.HAction.GetDefault("InsertText", pset.HSet)
pset.Text = "Hello world!"
hwp.HAction.Execute("InsertText", pset.HSet)

# (ver.0.4.0 이상에서는 스크립트매크로 방식으로 사용 가능)
HAction.GetDefault("InsertText", HParameterSet.HInsertText.HSet)
HParameterSet.HInsertText = "Hello world!"
HAction.Execute("InsertText", HParameterSet.HInsertText.HSet)

# 다른이름으로 저장
hwpx.save_as("./helloworld.hwp")

# 한/글 종료
hwpx.quit()
```

이밖에 구체적인 hwpx 모듈 사용법과 주요업데이트는 

https://martinii.fun/ 블로그에

틈틈이 포스팅으로 남기겠습니다.

방문해 주셔서 감사합니다.

행복한 하루 되세요!!