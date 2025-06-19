from time import sleep
from typing import Any, TYPE_CHECKING, Protocol, Tuple

if TYPE_CHECKING:
    from .core import Hwp


class _InnerMethods(Protocol):
    hwp: Any

    def get_cell_addr(self, str) -> Tuple[int]:
        pass

    def get_pos(self) -> Tuple[int]:
        pass

    def get_message_box_mode(self) -> int:
        pass

    def set_message_box_mode(self, mode: int) -> int:
        pass

    @property
    def HParameterSet(self):
        pass


class RunMethods(_InnerMethods):
    """
    Run 메서드 모음
    """

    def ASendBrowserText(self) -> bool:
        """
        웹브라우저로 보내기
        """
        return self.hwp.HAction.Run("ASendBrowserText")

    def AutoChangeHangul(self) -> bool:
        """
        구버전의 "낱자모 우선입력" 활성화 토글기능. 현재는 사용하지 않으며, 최신버전에서 <도구-글자판-글자판 자동 변경(A)> 기능에 통합되었다.낱자모 우선입력 기능은 제거된 것으로 보임

        """
        return self.hwp.HAction.Run("AutoChangeHangul")

    def AutoChangeRun(self) -> bool:
        """
        위 커맨드를 실행할 때마다 "글자판 자동 변경 기능"이 활성화/비활성화로 토글된다. 다만 API 등으로 텍스트를 입력하는 경우 원래 한/영 자동변환이 되지 않으므로, 자동화에는 쓰일 일이 없는 액션.
        """
        return self.hwp.HAction.Run("AutoChangeRun")

    def AutoSpellRun(self) -> bool:
        """
        맞춤법 도우미(맞춤법이 틀린 단어 밑에 빨간 점선) 활성화/비활성화를 토글한다. 실행 후(비활성화시) 몇 초 뒤에 붉은 줄이 사라지는 것을 확인할 수 있다. 중간 스페이스에 유의.
        """
        return self.hwp.HAction.Run("AutoSpell Run")

    def AutoSpellSelect0(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect0")

    def AutoSpellSelect1(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect1")

    def AutoSpellSelect2(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect2")

    def AutoSpellSelect3(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect3")

    def AutoSpellSelect4(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect4")

    def AutoSpellSelect5(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect5")

    def AutoSpellSelect6(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect6")

    def AutoSpellSelect7(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect7")

    def AutoSpellSelect8(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect8")

    def AutoSpellSelect9(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect9")

    def AutoSpellSelect10(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect10")

    def AutoSpellSelect11(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect11")

    def AutoSpellSelect12(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect12")

    def AutoSpellSelect13(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect13")

    def AutoSpellSelect14(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect14")

    def AutoSpellSelect15(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect15")

    def AutoSpellSelect16(self) -> bool:
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.

        """
        return self.hwp.HAction.Run("AutoSpellSelect16")

    def BookmarkEditDialog(self) -> bool:
        """북마크 편집 대화상자 호출 액션 - 책갈피 작업창 에서 편집 대화상자를 호출하기 위한 액션"""
        return self.hwp.HAction.Run("BookmarkEditDialog")

    def BottomTabFrameClose(self) -> bool:
        """아래쪽 작업창 감추기"""
        return self.hwp.HAction.Run("BottomTabFrameClose")

    def BreakColDef(self) -> bool:
        """
        다단 레이아웃을 사용하는 경우의 "단 정의 삽입 액션(Ctrl-Alt-Enter)".

        단 정의 삽입 위치를 기점으로 구분된 다단을 하나 추가한다.
        다단이 아닌 경우에는 일반 "문단나누기(Enter)"와 동일하다.
        """
        return self.hwp.HAction.Run("BreakColDef")

    def BreakColumn(self) -> bool:
        """
        다단 레이아웃을 사용하는 경우 "단 나누기[배분다단] 액션(Ctrl-Shift-Enter)".

        단 정의 삽입 위치를 기점으로 구분된 다단을 하나 추가한다.
        다단이 아닌 경우에는 일반 "문단나누기(Enter)"와 동일하다.

        """
        return self.hwp.HAction.Run("BreakColumn")

    def BreakLine(self) -> bool:
        """
        라인나누기 액션(Shift-Enter).

        들여쓰기나 내어쓰기 등 문단속성이 적용되어 있는 경우에
        속성을 유지한 채로 줄넘김만 삽입한다. 이 단축키를 모르고 보고서를 작성하면,
        들여쓰기를 맞추기 위해 스페이스를 여러 개 삽입했다가,
        앞의 문구를 수정하는 과정에서 스페이스 뭉치가 문단 중간에 들어가버리는 대참사가 자주 발생할 수 있다.
        """
        return self.hwp.HAction.Run("BreakLine")

    def BreakPage(self) -> bool:
        """
        쪽 나누기 액션(Ctrl-Enter).

        캐럿 위치를 기준으로 하단의 글을 다음 페이지로 넘긴다.
        BreakLine과 마찬가지로 보고서 작성시 자주 사용해야 하는 액션으로,
        이 기능을 사용하지 않고 보고서 작성시 엔터를 십여개 치고 다음 챕터 제목을 입력했다가,
        일부 수정하면서 챕터 제목이 중간에 와 있는 경우 등의 불상사가 발생할 수 있다.
        """
        return self.hwp.HAction.Run("BreakPage")

    def BreakPara(self) -> bool:
        """
        줄바꿈(문단 나누기). 일반적인 엔터와 동일하다.

        Returns:
            True

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 한/글 창이 열림
            >>> hwp.BreakPara()  # 줄바꿈 삽입!
            True
            >>> hwp.insert_text("\\r\\n")  # BreakPara와 동일
            True
        """
        return self.hwp.HAction.Run("BreakPara")

    def BreakSection(self) -> bool:
        """
        구역[섹션] 나누기 액션(Shift-Alt-Enter). 새로 생성된 섹션에서는 편집용지를 다르게 설정하거나, 혹은 새 개요번호/모양을 만든다든지 할 수 있다. 단, 초깃값으로 새 섹션이 생성되는 게 아니라, 기존 섹션의 편집용지, 바탕쪽 상태, 각주/미주 모양, 프레젠테이션 상태, 쪽 테두리/배경 속성 및 단 모양 등 대부분을 그대로 이어받으며, 새 섹션에서 수정시 기존 섹션에는 대부분 영향을 미치지 않는다.

        """
        return self.hwp.HAction.Run("BreakSection")

    def Cancel(self) -> bool:
        """
        취소 액션. Esc 키를 눌렀을 때와 동일하다. 대표적인 예로는 텍스트 선택상태나 셀선택모드  해제 또는 이미지, 표 등의 개체 선택을 취소할 때 사용한다. Cancel 과 유사하게 쓰이는 액션으로 Close(또는 CloseEx, Shift-Esc)가 있다.

        """
        return self.hwp.HAction.Run("Cancel")

    def CaptureHandler(self) -> bool:
        """
        갈무리 시작 액션. 현재 버전에서는 Run커맨드로 사용할 수 없는 것 같다. 어떤 이미지포맷으로 저장하든 오류발생.

        """
        return self.hwp.HAction.Run("CaptureHandler")

    def CaptureDialog(self) -> bool:
        """
        갈무리 끝 액션. 현재 버전에서는 Run커맨드로 사용할 수 없는 것 같다. 어떤 이미지포맷으로 저장하든 오류발생.

        """
        return self.hwp.HAction.Run("CaptureDialog")

    def ChangeSkin(self) -> bool:
        """스킨 바꾸기"""
        return self.hwp.HAction.Run("ChangeSkin")

    def CharShapeBold(self) -> bool:
        """
        글자모양 중 "진하게Bold" 속성을 토글하는 액션. 이 액션을 실행하기 전에 특정 셀이나 텍스트가 선택된 상태여야 하며, 이 커맨드만으로는 확실히 "진하게" 속성이 적용되었는지 확인할 수 없다. 그 이유는 토글 커맨드라서, 기존에 진하게 적용되어 있었다면, 해제되어버리기 때문이다. 확실히 진하게를 적용하는 방법으로는, 초기에 모든 텍스트의 진하게를 해제(진하게 두 번??)한다든지, 파라미터셋을 활용하여 진하게 속성이 적용되어 있는지를 확인하는 방법 등이 있다.

        """
        return self.hwp.HAction.Run("CharShapeBold")

    def CharShapeCenterline(self) -> bool:
        """
        글자에 취소선 적용을 토글하는 액션. Bold와 마찬가지로 토글이므로, 기존에 취소선이 적용되어 있다면 해제되어버리므로 사용에 유의해야 한다.

        """
        return self.hwp.HAction.Run("CharShapeCenterline")

    def CharShapeEmboss(self) -> bool:
        """
        글자모양에 양각 속성(글자가 튀어나온 느낌) 적용을 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeEmboss")

    def CharShapeEngrave(self) -> bool:
        """
        글자모양에 음각 속성(글자가 움푹 들어간 느낌) 적용을 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeEngrave")

    def CharShapeHeight(self) -> bool:
        """
        글자모양(Alt-L) 대화상자를 열고, 포커스를 "기준 크기"로 이동한다. (수작업이 필요하므로 자동화에 사용하지는 않는다. 유사한 액션으로 글꼴 언어를 선택하는 CharShapeLang, CharShapeSpacing, CharShapeTypeFace, CharShapeWidth 등이 있다.)

        """
        return self.hwp.HAction.Run("CharShapeHeight")

    def CharShapeHeightDecrease(self) -> bool:
        """
        글자 크기를 1포인트씩 작게 한다. 단, 속도가 다소 느리므로.. 큰 폭으로 조정할 때에는 다른 방법을 쓰는 것을 추천.

        """
        return self.hwp.HAction.Run("CharShapeHeightDecrease")

    def CharShapeHeightIncrease(self) -> bool:
        """
        글자 크기를 1포인트씩 크게 한다. 단, 속도가 다소 느리므로.. 큰 폭으로 조정할 때에는 다른 방법을 쓰는 것을 추천.

        """
        return self.hwp.HAction.Run("CharShapeHeightIncrease")

    def CharShapeItalic(self) -> bool:
        """
        글자 모양에 이탤릭 속성을 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeItalic")

    def CharShapeLang(self) -> bool:
        """글자 언어"""
        return self.hwp.HAction.Run("CharShapeLang")

    def CharShapeNextFaceName(self) -> bool:
        """
        다음 글꼴로 이동(Shift-Alt-F)한다. 단, 이 액션으로 어떤 폰트가 선택되었는지를 파이썬에서 확인하려면 파라미터셋에 접근해야 한다. 유사한 커맨드로, CharShapePrevFaceName 이 있다.

        """
        return self.hwp.HAction.Run("CharShapeNextFaceName")

    def CharShapeNormal(self) -> bool:
        """
        글자모양에 적용된 속성 및 글자색 등 전부를 해제(Shift-Alt-C)한다. 단 글꼴, 크기 등은 바뀌지 않는다.

        """
        return self.hwp.HAction.Run("CharShapeNormal")

    def CharShapeOutline(self) -> bool:
        """
        글자모양의 외곽선 속성을 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeOutline")

    def CharShapePrevFaceName(self) -> bool:
        """
        이전 글꼴 ALT+SHIFT+G
        """
        return self.hwp.HAction.Run("CharShapePrevFaceName")

    def CharShapeShadow(self) -> bool:
        """
        선택한 텍스트 글자모양 중 그림자 속성을 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeShadow")

    def CharShapeSpacing(self) -> bool:
        """
        글자모양(alt-L) 창을 열고, 자간 값에 포커스를 옮긴다.
        """
        return self.hwp.HAction.Run("CharShapeSpacing")

    def CharShapeSpacingDecrease(self) -> bool:
        """
        자간을 1%씩 좁힌다. 최대 -50%까지 좁힐 수 있다. 다만 자동화 작업시 줄넘김을 체크하는 것이 상당히 번거로운 작업이므로, 크게 보고서의 틀이 바뀌지 않는 선에서는 자간을 좁히는 것보다 "한 줄로 입력"을 활용하는 편이 간단하고 자연스러울 수 있다. 한 줄로 입력 옵션 : 문단모양(Alt-T)의 확장 탭에 있음. 한 줄로 입력을 활성화해놓은 문단이나 셀에서는 자간이 아래와 같이 자동으로 좁혀진다.

        """
        return self.hwp.HAction.Run("CharShapeSpacingDecrease")

    def CharShapeSpacingIncrease(self) -> bool:
        """
        자간을 1%씩 넓힌다. 최대 50%까지 넓힐 수 있다.

        """
        return self.hwp.HAction.Run("CharShapeSpacingIncrease")

    def CharShapeSubscript(self) -> bool:
        """
        선택한 텍스트에 아래첨자 속성을 토글(Shift-Alt-S)한다.

        """
        return self.hwp.HAction.Run("CharShapeSubscript")

    def CharShapeSuperscript(self) -> bool:
        """
        선택한 텍스트에 위첨자 속성을 토글(Shift-Alt-P)한다.

        """
        return self.hwp.HAction.Run("CharShapeSuperscript")

    def CharShapeSuperSubscript(self) -> bool:
        """
        선택한 텍스트의 첨자속성을 위→아래→보통의 순서를 반복해서 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeSuperSubscript")

    def CharShapeTextColorBlack(self) -> bool:
        """
        선택한 텍스트의 글자색을 검정색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorBlack")

    def CharShapeTextColorBlue(self) -> bool:
        """
        선택한 텍스트의 글자색을 파란색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorBlue")

    def CharShapeTextColorBluish(self) -> bool:
        """
        선택한 텍스트의 글자색을 청록색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorBluish")

    def CharShapeTextColorGreen(self) -> bool:
        """
        선택한 텍스트의 글자색을 초록색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorGreen")

    def CharShapeTextColorRed(self) -> bool:
        """
        선택한 텍스트의 글자색을 빨간색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorRed")

    def CharShapeTextColorViolet(self) -> bool:
        """
        선택한 텍스트의 글자색을 보라색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorViolet")

    def CharShapeTextColorWhite(self) -> bool:
        """
        선택한 텍스트의 글자색을 흰색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorWhite")

    def CharShapeTextColorYellow(self) -> bool:
        """
        선택한 텍스트의 글자색을 노란색으로 변경한다.

        """
        return self.hwp.HAction.Run("CharShapeTextColorYellow")

    def CharShapeTypeface(self) -> bool:
        """글꼴 이름(글자 모양 대화상자에서 Focus이동용 으로 사용)"""
        return self.hwp.HAction.Run("CharShapeTypeface")

    def CharShapeUnderline(self) -> bool:
        """
        선택한 텍스트에 밑줄 속성을 토글한다. 대소문자에 유의해야 한다. (UnderLine이 아니다.)

        """
        return self.hwp.HAction.Run("CharShapeUnderline")

    def CharShapeWidth(self) -> bool:
        """글자 모양(Alt-L) 창에서 글자 장평에 포커스를 둔다."""
        return self.hwp.HAction.Run("CharShapeWidth")

    def CharShapeWidthDecrease(self) -> bool:
        """
        장평을 1%씩 줄인다. 장평 범위는 50~200%이며, 장평을 늘일 때는 Decrease 대신 Increase를 사용하면 된다.

        """
        return self.hwp.HAction.Run("CharShapeWidthDecrease")

    def CharShapeWidthIncrease(self) -> bool:
        """
        장평을 1%씩 줄인다. 장평 범위는 50~200%이며, 장평을 줄일 때는 Increase 대신 Decrease를 사용하면 된다.

        """
        return self.hwp.HAction.Run("CharShapeWidthIncrease")

    def CloseEx(self) -> bool:
        """
        현재 리스트를 닫고 상위 리스트로 이동하는 액션.

        Close와 CloseEx는 유사하나 두 가지 차이점이 있다.
        첫 번째로는 여러 계층의 표 안에서 CloseEx 실행시
        본문이 아니라 상위의 표(셀)로 캐럿이 단계적으로 이동한다는 점.
        반면 Close는 무조건 본문으로 나간다.
        두 번째로, CloseEx에는 전체화면(최대화 말고)을 해제하는 기능이 있다.
        Close로는 전체화면 해제가 되지 않는다.
        사용빈도가 가장 높은 액션 중의 하나라고 생각한다.
        """
        return self.hwp.HAction.Run("CloseEx")

    def Comment(self) -> bool:
        """
        새로운 숨은설명 컨트롤을 만들고, 해당 숨은설명으로 이동.

        아래아한글에 "숨은 설명"이 있다는 걸 아는 사람도 없다시피 한데,
        그 "숨은 설명" 관련한 Run 액션이 세 개나 있다.
        Comment 액션은 표현 그대로 숨은 설명을 붙일 수 있다.
        텍스트만 넣을 수 있을 것 같은 액션이름인데,
        사실 표나 그림도 자유롭게 삽입할 수 있기 때문에,
        문서 안에 몰래 숨겨놓은 또다른 문서 느낌이다.
        파일별로 자동화에 활용할 수 있는 특정 문자열을
        파이썬이 아니라 숨은설명 안에 붙여놓고 활용할 수도 있지 않을까
        이런저런 고민을 해봤는데, 개인적으로 자동화에 제대로 활용한 적은 한 번도 없었다.
        숨은 설명이라고 민감한 정보를 넣으면 안 되는데,
        완전히 숨겨져 있는 게 아니기 때문이다.
        현재 캐럿위치에 [숨은설명] 조판부호가 삽입되며,
        이를 통해 숨은 설명 내용이 확인 가능하므로 유념해야 한다.
        재미있는 점은, 숨은설명 안에 또 숨은설명을 삽입할 수 있다.
        숨은설명 안에다 숨은설명을 넣고 그 안에 또 숨은설명을 넣는...
        이런 테스트를 해봤는데 2,400단계 정도에서 한글이 종료돼버렸다.

        Returns:
            성공시 True, 실패시 False를 리턴.
        """
        return self.hwp.HAction.Run("Comment")

    def CommentDelete(self) -> bool:
        """
        숨은설명 지우기

        단어 그대로 숨은 설명을 지우는 액션이다.
        단, 사용방법이 까다로운데 숨은 설명 안에 들어가서 CommentDelete를 실행하면,
        지울지 말지(Yes/No) 팝업이 나타난다.
        나중에 자세히 설명하겠지만 이런 팝업을 자동처리하는 방법은 hwp.SetMessageBoxMode() 메서드를 미리 실행해놓는 것이다.
        Yes/No 방식의 팝업에서 Yes를 선택하는 파라미터는 0x10000 (또는 65536)이므로,
        hwp.SetMessageBoxMode(0x10000) 를 사용하면 된다.
        """
        return self.hwp.HAction.Run("CommentDelete")

    def CommentModify(self) -> bool:
        """
        숨은설명 수정하기

        단어 그대로 숨은 설명을 수정하는 액션이다.
        캐럿은 해당 [숨은설명] 조판부호 바로 앞에 위치하고 있어야 한다.

        """
        return self.hwp.HAction.Run("CommentModify")

    def Copy(self) -> bool:
        """
        복사하기. (비추천)

        선택되어 있는 문자열 혹은 개체(표, 이미지 등)를 클립보드에 저장한다.
        파이썬에서 클립보드를 다루는 모듈은 pyperclip이나,
        pywin32의 win32clipboard 두 가지가 가장 많이 쓰이는데,
        단순한 문자열의 경우 hwp.Copy() 실행 후 pyperclip.paste()로 파이썬으로 가져올 수 있지만,
        서식 등의 메타 정보를 모두 잃어버린다.

        또한 빠른 반복작업시 윈도우OS의 클립보드 잠금 때문에 오류가 발생할 수 있다.

        hwp.Copy()나 hwp.Paste() 대신 hwp.GetTextFile()이나 hwp.save_block_as 등의 메서드를 사용하는 것을 추천.

        """
        return self.hwp.HAction.Run("Copy")

    def CopyPage(self) -> bool:
        """
        쪽 복사

        한글2014 이하의 버전에서는 사용할 수 없다.
        """
        return self.hwp.HAction.Run("CopyPage")

    def CustCopyBtn(self) -> bool:
        """툴바 버튼 복사하기"""
        return self.hwp.HAction.Run("CustCopyBtn")

    def CustCutBtn(self) -> bool:
        """툴바 버튼 오려두기"""
        return self.hwp.HAction.Run("CustCutBtn")

    def CustEraseBtn(self) -> bool:
        """툴바 버튼 지우기"""
        return self.hwp.HAction.Run("CustEraseBtn")

    def CustInsSepBtn(self) -> bool:
        """툴바 버튼에 구분선 넣기"""
        return self.hwp.HAction.Run("CustInsSepBtn")

    def CustomizeToolbar(self) -> bool:
        """도구상자 사용자 설정"""
        return self.hwp.HAction.Run("CustomizeToolbar")

    def CustPasteBtn(self) -> bool:
        """툴바 버튼 붙여기"""
        return self.hwp.HAction.Run("CustPasteBtn")

    def CustRenameBtn(self) -> bool:
        """툴바 버튼 이름 바꾸기"""
        return self.hwp.HAction.Run("CustRenameBtn")

    def CustRestBtn(self) -> bool:
        """툴바 버튼 처음 상태로 되돌리기"""
        return self.hwp.HAction.Run("CustRestBtn")

    def CustViewIconBtn(self) -> bool:
        """툴바 버튼 아이콘만 보이기"""
        return self.hwp.HAction.Run("CustViewIconBtn")

    def CustViewIconNameBtn(self) -> bool:
        """툴바 버튼 이름과 아이콘 보이기"""
        return self.hwp.HAction.Run("CustViewIconNameBtn")

    def CustViewNameBtn(self) -> bool:
        """툴바 버튼 이름만 보이기"""
        return self.hwp.HAction.Run("CustViewNameBtn")

    def DeleteField(self) -> bool:
        """
        누름틀지우기.

        누름틀 안의 내용은 지우지 않고, 단순히 누름틀만 지운다.
        지울 때 캐럿의 위치는 누름틀 안이든, 앞이나 뒤든 붙어있기만 하면 된다.
        만약 최종문서에는 누름틀을 넣지 않고
        모두 일반 텍스트로 변환하려고 할 때
        이 기능을 활용하면 된다.

        Returns:
            성공시 True, 실패시 False를 리턴.

        """
        return self.hwp.HAction.Run("DeleteField")

    def DeleteFieldMemo(self) -> bool:
        """
        메모 지우기. 누름틀 지우기와 유사하다. 메모 누름틀에 붙어있거나, 메모 안에 들어가 있는 경우 위 액션 실행시 해당 메모가 삭제된다.
        """
        return self.hwp.HAction.Run("DeleteFieldMemo")

    def DeletePage(self) -> bool:
        """
        쪽 지우기

        한글2018 미만의 버전에서는 사용할 수 없다.
        """
        return self.hwp.HAction.Run("DeletePage")

    def DeletePrivateInfoMark(self) -> bool:
        """
        개인 정보 감추기한 정보 다시보기.
        (개인 정보 보호 암호를 설정해야 함)
        """
        return self.hwp.HAction.Run("DeletePrivateInfoMark")

    def DeletePrivateInfoMarkAtCurrentPos(self) -> bool:
        """
        현재 캐럿 위치의 감추기한 개인 정보 다시 보기
        (개인 정보 보호 암호를 설정해야 함)
        """
        return self.hwp.HAction.Run("DeletePrivateInfoMarkAtCurrentPos")

    def DrawObjCancelOneStep(self) -> bool:
        """
        다각형(곡선) 그리는 중 이전 선 지우기.

        현재 사용 안함(?)
        """
        return self.hwp.HAction.Run("DrawObjCancelOneStep")

    def DrawObjEditDetail(self) -> bool:
        """
        그리기 개체 중 다각형 점편집 액션.

        다각형이 선택된 상태에서만 실행가능.
        """
        return self.hwp.HAction.Run("DrawObjEditDetail")

    def DrawObjOpenClosePolygon(self) -> bool:
        """
        닫힌 다각형 열기 또는 열린 다각형 닫기 토글.

        ①다각형 개체 선택상태가 아니라 편집상태에서만 위 명령어가 실행된다.

        ②닫힌 다각형을 열 때는 마지막으로 봉합된 점에서 아주 조금만 열린다.

        ③아주 조금만 열린 상태에서 닫으면 노드(꼭지점)가 추가되지 않지만, 적절한 거리를 벌리고 닫기를 하면 추가됨.

        """
        return self.hwp.HAction.Run("DrawObjOpenClosePolygon")

    def DrawObjTemplateSave(self) -> bool:
        """
        그리기개체를 그리기마당에 템플릿으로 등록하는 액션

        (어떻게 써먹고 싶어도 방법을 모르겠다...)
        그리기개체가 선택된 상태에서만 실행 가능하다.
        여담으로, 그리기 마당에 임의로 등록한 개체 삭제 아이콘을 못 찾고 있는데;
        한글2020 기준으로, 개체 이름을 "얼굴"이라고 "기본도형"에 저장했을 경우,
        찾아가서 아래의 파일을 삭제해도 된다.
        `"C:/Users/이름/AppData/Roaming/HNC/User/Shared110/HwpTemplate/Draw/FG_Basic_Shapes/얼굴.drt"`

        """
        return self.hwp.HAction.Run("DrawObjTemplateSave")

    def EasyFind(self) -> bool:
        """
        쉬운 찾기

        사용하지 않음.

        """
        return self.hwp.HAction.Run("EasyFind")

    def EditFieldMemo(self) -> bool:
        """
        메모 내용 편집 액션.

        "메모 내용 보기" 창이 하단에 열린다.
        SplitMemoOpen과 동일한 기능으로 보이며,
        메모내용보기창에서 두 번째 이후의 메모 클릭시 메모내용 보기 창이 닫히는 버그가 있다.(한/글 2020 기준)참고로 메모내용 보기 창을 닫을 때는 SplitMemoClose 커맨드를 쓰면 된다.

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        return self.hwp.HAction.Run("EditFieldMemo")

    def EditParaDown(self) -> bool:
        """
        현재 캐럿이 위치한 문단 또는 선택한 문단 전체를, 하단에 위치한 문단 아래로 옮긴다.
        """
        return self.hwp.HAction.Run("EditParaDown")

    def EditParaUp(self) -> bool:
        """
        현재 캐럿이 위치한 문단 또는 선택한 문단 전체를, 상단에 위치한 문단의 위로 옮긴다.
        """
        return self.hwp.HAction.Run("EditParaUp")

    def Erase(self) -> bool:
        """
        선택한 문자나 개체 삭제.

        문자열이나 컨트롤 등을 삭제한다는 점에서는 Delete나 DeleteBack과 유사하지만,
        가장 큰 차이점은, 아무 것도 선택되어 있지 않은 상태일 때 Erase는 아무 것도 지우지 않는다는 점이다.
        (Delete나 DeleteBack은 어찌됐든 앞뒤의 뭔가를 지운다.)

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        return self.hwp.HAction.Run("Erase")

    def FileClose(self) -> bool:
        """
        문서 닫기.

        저장 이후 변경사항이 있으면 팝업이 뜨므로 주의
        """
        return self.hwp.HAction.Run("FileClose")

    def FileFind(self) -> bool:
        """문서 찾기"""
        return self.hwp.HAction.Run("FileFind")

    def FileNew(self) -> bool:
        """
        새 문서 창을 여는 명령어.

        참고로 현재 창에서 새 탭을 여는 명령어는 ``hwp.FileNewTab()``

        여담이지만 한/글2020 기준으로 새 창은 30개까지 열 수 있다.
        그리고 한 창에는 탭을 30개까지 열 수 있다.
        즉, (리소스만 충분하다면) 동시에 열어서 자동화를 돌릴 수 있는
        문서 갯수는 900개!
        """
        return self.hwp.HAction.Run("FileNew")

    def FileNewTab(self) -> bool:
        """
        새 탭을 여는 명령어.
        """
        return self.hwp.HAction.Run("FileNewTab")

    def FileNextVersionDiff(self) -> bool:
        """버전 비교 :　앞으로 이동"""
        return self.hwp.HAction.Run("FileNextVersionDiff")

    def FilePrevVersionDiff(self) -> bool:
        """버전 비교 : 뒤로 이동"""
        return self.hwp.HAction.Run("FilePrevVersionDiff")

    def FileOpen(self) -> bool:
        """
        문서를 여는 명령어.

        단 파일선택 팝업이 뜨므로,
        자동화작업시에는 이 명령어를 사용하지 않는다.
        대신 hwp.open(파일명)을 사용해야 한다.
        """
        return self.hwp.HAction.Run("FileOpen")

    def FileOpenMRU(self) -> bool:
        """
        최근 작업문서 열기

        현재는 FileOpen과 동일한 동작을 하는 것으로 보임.
        사용자입력을 요구하는 팝업이 뜨므로
        자동화에 사용하지 않으며, hwp.open(Path)을 써야 한다.

        """
        return self.hwp.HAction.Run("FileOpenMRU")

    def FilePreview(self) -> bool:
        """
        미리보기 창을 열어준다.

        자동화와 큰 연관이 없어 자주 쓰이지도 않고,
        더군다나 닫는 명령어가 없다.
        또한 이 명령어는 hwp.XHwpDocuments.Item(0).XHwpPrint.RunFilePreview()와 동일한 동작을 하는데,
        재미있는 점은,

            - ①스크립트 매크로 녹화 진행중에 ``hwp.FilePreview()``는 실행해도 반응이 없고, 녹화 로그에도 잡히지 않는다.
            - ②그리고 스크립트매크로 녹화 진행중에 [파일] - [미리보기(V)] 메뉴도 비활성화되어 있어 코드를 알 수 없다.
            - ③그런데 hwp.XHwpDocuments.Item(0).XHwpPrint.RunFilePreview()는 녹화중에도 실행이 된다. 녹화된 코드와 관련하여 남기고 싶은 코멘트가 많은데, 별도의 포스팅으로 남길 예정.
        """
        return self.hwp.HAction.Run("FilePreview")

    def FileQuit(self) -> bool:
        """
        한/글 프로그램을 종료한다.

        단, 저장 이후 문서수정이 있는 경우에는 팝업이 뜨므로,
        ①저장하거나 ②수정내용을 버리는 메서드를 활용해야 한다.

        """
        return self.hwp.HAction.Run("FileQuit")

    def FileSave(self) -> bool:
        """
        문서 저장(Alt-S).

        가급적 ``hwp.save()``를 사용하자.
        ``hwp.save()``와 ``hwp.FileSave()``에 차이가 있는데
        ``hwp.save()``는 실제 변경이 없으면 저장을 수행하지 않지만
        ``hwp.FileSave()``는 변경이 없어도 저장을 수행하므로 수정일자가 바뀐다.
        """
        return self.hwp.HAction.Run("FileSave")

    def FileSaveAs(self) -> bool:
        """
        다른 이름으로 저장(Alt-V).

        사용자입력을 필요로 하므로 이 액션은 사용하지 않는다.
        대신 hwp.save_as(Path)를 사용하면 된다.

        """
        return self.hwp.HAction.Run("FileSaveAs")

    def FileSaveAsDRM(self) -> bool:
        """배포용 문서로 저장하기"""
        return self.hwp.HAction.Run("FileSaveAsDRM")

    def FileSaveOptionDlg(self) -> bool:
        """저장 옵션창 열기"""
        return self.hwp.HAction.Run("FileSaveOptionDlg")

    def FileVersionDiffChangeAlign(self) -> bool:
        """버전 비교 : 비교화면 배열 변경 (좌우↔상하)"""
        return self.hwp.HAction.Run("FileVersionDiffChangeAlign")

    def FileVersionDiffSameAlign(self) -> bool:
        """버전 비교 : 비교화면 다시 정렬"""
        return self.hwp.HAction.Run("FileVersionDiffSameAlign")

    def FileVersionDiffSyncScroll(self) -> bool:
        """버전 비교 : 비교화면 동시에 이동"""
        return self.hwp.HAction.Run("FileVersionDiffSyncScroll")

    def FillColorShadeDec(self) -> bool:
        """면색 음영 비율 감소"""
        return self.hwp.HAction.Run("FillColorShadeDec")

    def FillColorShadeInc(self) -> bool:
        """면색 음영 비율 증가"""
        return self.hwp.HAction.Run("FillColorShadeInc")

    def FindForeBackBookmark(self) -> bool:
        """
        책갈피 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackBookmark")

    def FindForeBackCtrl(self) -> bool:
        """
        조판부호 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        참고로 ``hwp.FindForeBackSelectCtrl``은 선택.
        """
        return self.hwp.HAction.Run("FindForeBackCtrl")

    def FindForeBackFind(self) -> bool:
        """
        찾기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackFind")

    def FindForeBackLine(self) -> bool:
        """
        줄 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackLine")

    def FindForeBackPage(self) -> bool:
        """
        쪽 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackPage")

    def FindForeBackSection(self) -> bool:
        """
        구역 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackSection")

    def FindForeBackSelectCtrl(self) -> bool:
        """앞뒤로 찾아가기 : 조판 부호 찾기 (선택)"""
        return self.hwp.HAction.Run("FindForeBackSelectCtrl")

    def FindForeBackStyle(self) -> bool:
        """
        스타일 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackStyle")

    def FormDesignMode(self) -> bool:
        """양식 개체 디자인 모드 변경"""
        return self.hwp.HAction.Run("FormDesignMode")

    def FormObjCreatorCheckButton(self) -> bool:
        """양식 개체 체크 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorCheckButton")

    def FormObjCreatorComboBox(self) -> bool:
        """양식 개체 콤보 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorComboBox")

    def FormObjCreatorEdit(self) -> bool:
        """양식 개체 에디트 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorEdit")

    def FormObjCreatorListBox(self) -> bool:
        """양식 개체 리스트 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorListBox")

    def FormObjCreatorPushButton(self) -> bool:
        """양식 개체 푸쉬 버튼 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorPushButton")

    def FormObjCreatorRadioButton(self) -> bool:
        """양식 개체 라디오 버튼 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorRadioButton")

    def FormObjCreatorScrollBar(self) -> bool:
        """양식 개체 스크롤바 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorScrollBar")

    def FormObjRadioGroup(self) -> bool:
        """양식 개체 라디오 버튼 그룹 묶기"""
        return self.hwp.HAction.Run("FormObjRadioGroup")

    def FrameFullScreen(self) -> bool:
        """
        한/글 프로그램창 전체화면(창 최대화 아님).

        전체화면 해제는 hwp.FrameFullScreenEnd() 또는 hwp.CloseEx()
        """
        return self.hwp.HAction.Run("FrameFullScreen")

    def FrameFullScreenEnd(self) -> bool:
        """전체 화면 닫기"""
        return self.hwp.HAction.Run("FrameFullScreenEnd")

    def FrameHRuler(self) -> bool:
        """가로축 눈금자 보이기/감추기"""
        return self.hwp.HAction.Run("FrameHRuler")

    def FrameStatusBar(self) -> bool:
        """
        한/글 프로그램 하단의 상태바 보이기/숨기기 토글

        """
        return self.hwp.HAction.Run("FrameStatusBar")

    def FrameViewZoomRibon(self) -> bool:
        """화면 확대/축소"""
        return self.hwp.HAction.Run("FrameViewZoomRibon")

    def FrameVRuler(self) -> bool:
        """세로축 눈금자 보이기/감추기"""
        return self.hwp.HAction.Run("FrameVRuler")

    def HancomRoom(self) -> bool:
        """한컴 계약방"""
        return self.hwp.HAction.Run("HancomRoom")

    def HanThDIC(self) -> bool:
        """
        한/글에 내장되어 있는 "유의어/반의어 사전"을 여는 액션.

        """
        return self.hwp.HAction.Run("HanThDIC")

    def HeaderFooterDelete(self) -> bool:
        """
        머리말/꼬리말 지우기.

        본문이 아니라 머리말/꼬리말 편집상태에서 실행해야
        삭제 팝업이 뜬다.
        삭제팝업 없이 머리말/꼬리말을 삭제하려면
        hwp.SetMessageBoxMode(0x10000)을 미리 실행해놓아야 한다.

        자동화작업시에는 ``hwp.HeaderFooterModify()``을 통해
        편집상태로 들어가야 한다.

        """
        return self.hwp.HAction.Run("HeaderFooterDelete")

    def HeaderFooterModify(self) -> bool:
        """
        머리말/꼬리말 고치기.

        마우스를 쓰지 않고 머리말/꼬리말 편집상태로 들어갈 수 있다.
        단, 커서가 머리말/꼬리말 컨트롤에 닿아 있는 상태에서 실행해야 한다.
        """
        return self.hwp.HAction.Run("HeaderFooterModify")

    def HeaderFooterToNext(self) -> bool:
        """
        다음 머리말/꼬리말.

        당장은 사용방법을 모르겠다..
        """
        return self.hwp.HAction.Run("HeaderFooterToNext")

    def HeaderFooterToPrev(self) -> bool:
        """
        이전 머리말.

        당장은 사용방법을 모르겠다..
        """
        return self.hwp.HAction.Run("HeaderFooterToPrev")

    def HelpContents(self) -> bool:
        """내용"""
        return self.hwp.HAction.Run("HelpContents")

    def HelpIndex(self) -> bool:
        """찾아보기"""
        return self.hwp.HAction.Run("HelpIndex")

    def HelpWeb(self) -> bool:
        """온라인 고객 지원"""
        return self.hwp.HAction.Run("HelpWeb")

    def HiddenCredits(self) -> bool:
        """
        인터넷 정보.

        사용방법을 모르겠다.

        """
        return self.hwp.HAction.Run("HiddenCredits")

    def HideTitle(self) -> bool:
        """
        차례 숨기기(Ctrl-K-S)

        ([도구 - 차례/색인 - 차례 숨기기] 메뉴에 대응.
        실행한 개요라인을 자동생성되는 제목차례에서 숨긴다.
        즉시 변경되지 않으며,
        ``hwp.UpdateAllContents()``(모든 차례 새로고침, Ctrl-KA) 실행시
        제목차례가 업데이트된다.
        """
        return self.hwp.HAction.Run("HideTitle")

    def HimConfig(self) -> bool:
        """
        입력기 언어별 환경설정.

        현재는 실행되지 않는 듯 하다.
        대신 ``hwp.HimKbdChange()``로 환경설정창을 띄울 수 있다.
        자동화에는 쓰이지 않는다.

        """
        return self.hwp.HAction.Run("Him Config")

    def HimKbdChange(self) -> bool:
        """
        입력기 언어별 환경설정.

        """
        return self.hwp.HAction.Run("HimKbdChange")

    def HorzScrollbar(self) -> bool:
        """가로축 스크롤바 보이기/감추기"""
        return self.hwp.HAction.Run("HorzScrollbar")

    def HwpCtrlEquationCreate97(self) -> bool:
        """
        한/글97버전 수식 만들기

        실행되지 않는 듯 하다.

        """
        return self.hwp.HAction.Run("HwpCtrlEquationCreate97")

    def HwpCtrlFileNew(self) -> bool:
        """
        한글컨트롤 전용 새문서.

        실행되지 않는 듯 하다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileNew")

    def HwpCtrlFileOpen(self) -> bool:
        """
        한글컨트롤 전용 파일 열기.

        실행되지 않는 듯 하다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileOpen")

    def HwpCtrlFileSave(self) -> bool:
        """
        한글컨트롤 전용 파일 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSave")

    def HwpCtrlFileSaveAs(self) -> bool:
        """
        한글컨트롤 전용 다른 이름으로 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAs")

    def HwpCtrlFileSaveAsAutoBlock(self) -> bool:
        """
        한글컨트롤 전용 다른이름으로 블록 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAsAutoBlock")

    def HwpCtrlFileSaveAutoBlock(self) -> bool:
        """
        한/글 컨트롤 전용 블록 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAutoBlock")

    def HwpCtrlFindDlg(self) -> bool:
        """
        한/글 컨트롤 전용 찾기 대화상자.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFindDlg")

    def HwpCtrlReplaceDlg(self) -> bool:
        """
        한/글 컨트롤 전용 바꾸기 대화상자

        """
        return self.hwp.HAction.Run("HwpCtrlReplaceDlg")

    def HwpDic(self) -> bool:
        """
        한컴 사전(F12).

        현재 캐럿이 닿아 있거나, 블록선택한 구간을 검색어에 자동으로 넣는다.

        """
        return self.hwp.HAction.Run("HwpDic")

    def HwpTabViewAction(self) -> bool:
        """빠른 실행 작업창"""
        return self.hwp.HAction.Run("HwpTabViewAction")

    def HwpTabViewAttribute(self) -> bool:
        """양식 개체 속성 작업창"""
        return self.hwp.HAction.Run("HwpTabViewAttribute")

    def HwpTabViewClipboard(self) -> bool:
        """클립보드 작업창"""
        return self.hwp.HAction.Run("HwpTabViewClipboard")

    def HwpTabViewDistant(self) -> bool:
        """쪽모양 보기 작업창"""
        return self.hwp.HAction.Run("HwpTabViewDistant")

    def HwpTabViewHwpDic(self) -> bool:
        """사전 검색 작업창"""
        return self.hwp.HAction.Run("HwpTabViewHwpDic")

    def HwpTabViewMasterPage(self) -> bool:
        """바탕쪽 보기 작업창"""
        return self.hwp.HAction.Run("HwpTabViewMasterPage")

    def HwpTabViewOutline(self) -> bool:
        """개요 보기 작업창"""
        return self.hwp.HAction.Run("HwpTabViewOutline")

    def HwpTabViewScript(self) -> bool:
        """스크립트 작업창"""
        return self.hwp.HAction.Run("HwpTabViewScript")

    def HwpViewType(self) -> bool:
        """문서창 모양 설정"""
        return self.hwp.HAction.Run("HwpViewType")

    def HwpWSDic(self) -> bool:
        """사전 검색 작업창 (Shift + F12)"""
        return self.hwp.HAction.Run("HwpWSDic")

    def HyperlinkBackward(self) -> bool:
        """
        하이퍼링크 뒤로.

        하이퍼링크를 통해서 문서를 탐색하여 페이지나 캐럿을 이동한 경우, (브라우저의 "뒤로가기"처럼) 이동 전의 위치로 돌아간다.

        """
        return self.hwp.HAction.Run("HyperlinkBackward")

    def HyperlinkForward(self) -> bool:
        """
        하이퍼링크 앞으로.

        ``hwp.HyperlinkBackward()`` 에 상반되는 명령어로, 브라우저의 "앞으로 가기"나 한/글의 재실행과 유사하다. 하이퍼링크 등으로 이동한 후에 뒤로가기를 눌렀다면, 캐럿이 뒤로가기 전 위치로 다시 이동한다.

        """
        return self.hwp.HAction.Run("HyperlinkForward")

    def ImageFindPath(self) -> bool:
        """
        그림 경로 찾기.

        현재는 실행되지 않는 듯.

        """
        return self.hwp.HAction.Run("ImageFindPath")

    def ImportCharactersFromPiuctre(self) -> bool:
        """
        그림에서 글자 가져오기(OCR). (성능은 차차 좋아질 것...)
        문서에 포함된 그림이 아니라, 로컬의 파일을 선택해야 하고,
        추출된 텍스트는 클립보드에 저장되므로 사용시 주의.
        """
        return self.hwp.HAction.Run("ImportCharactersFromPicture")

    def InputCodeChange(self) -> bool:
        """
        문자/코드 변환

        현재 캐럿의 바로 앞 문자를 찾아서 문자이면 코드로, 코드이면 문자로 변환해준다.(변환 가능한 코드영역 0x0020 ~ 0x10FFFF 까지)

        """
        return self.hwp.HAction.Run("InputCodeChange")

    def InputHanja(self) -> bool:
        """
        한자로 바꾸기 창을 띄워준다.

        추가입력이 필요하여 자동화에는 쓰이지 않음.

        """
        return self.hwp.HAction.Run("InputHanja")

    def InputHanjaBusu(self) -> bool:
        """
        부수로 입력.

        자동화에는 쓰이지 않음.

        """
        return self.hwp.HAction.Run("InputHanjaBusu")

    def InputHanjaMean(self) -> bool:
        """
        한자 새김 입력창 띄우기.

        뜻과 음을 입력하면 적절한 한자를 삽입해준다.입력시 뜻과 음은 붙여서 입력. (예)하늘천

        """
        return self.hwp.HAction.Run("InputHanjaMean")

    def InsertAutoNum(self) -> bool:
        """
        번호 다시 넣기(?)

        실행이 안되는 듯.

        """
        return self.hwp.HAction.Run("InsertAutoNum")

    def InsertCpNo(self) -> bool:
        """
        현재 쪽번호(상용구) 삽입.

        쪽번호와 마찬가지로, 문자열이 실시간으로 변경된다.

        ※유의사항 : 이 쪽번호는 찾기, 찾아바꾸기, GetText 및 누름틀 안에 넣고 GetFieldText나 복붙 등 그 어떤 방법으로도 추출되지 않는다.
        한 마디로 눈에는 보이는 것 같지만 실재하지 않는 숫자임. 참고로 표번호도 그렇다. 값이 아니라 속성이라서 그렇다.

        """
        return self.hwp.HAction.Run("InsertCpNo")

    def InsertCpTpNo(self) -> bool:
        """
        상용구 코드 넣기(현재 쪽/전체 쪽).

        실시간으로 변경된다.

        """
        return self.hwp.HAction.Run("InsertCpTpNo")

    def InsertDateCode(self) -> bool:
        """
        상용구 코드 넣기(만든 날짜).

        현재날짜가 아님에 유의.

        """
        return self.hwp.HAction.Run("InsertDateCode")

    def InsertDocInfo(self) -> bool:
        """
        상용구 코드 넣기

        (만든 사람, 현재 쪽, 만든 날짜)

        """
        return self.hwp.HAction.Run("InsertDocInfo")

    def InsertEndnote(self) -> bool:
        """
        미주 입력

        """
        return self.hwp.HAction.Run("InsertEndnote")

    def InsertFieldCitation(self) -> bool:
        """
        인용(citation) 삽입
        """
        return self.hwp.HAction.Run("InsertFieldCitation")

    def InsertFieldDateTime(self) -> bool:
        """
        날짜/시간 코드로 넣기([입력-날짜/시간-날짜/시간 코드]메뉴와 동일)

        """
        return self.hwp.HAction.Run("InsertFieldDateTime")

    def InsertFieldMemo(self) -> bool:
        """
        메모 넣기([입력-메모-메모 넣기]메뉴와 동일)

        """
        return self.hwp.HAction.Run("InsertFieldMemo")

    def InsertFieldRevisionChagne(self) -> bool:
        """
        메모고침표 넣기

        (현재 한/글메뉴에 없음, 메모와 동일한 기능)

        """
        return self.hwp.HAction.Run("InsertFieldRevisionChagne")

    def InsertFixedWidthSpace(self) -> bool:
        """
        고정폭 빈칸 삽입

        """
        return self.hwp.HAction.Run("InsertFixedWidthSpace")

    def InsertFootnote(self) -> bool:
        """
        각주 입력

        """
        return self.hwp.HAction.Run("InsertFootnote")

    def InsertLastPrintDate(self) -> bool:
        """
        상용구 코드 넣기(마지막 인쇄한 날짜)

        """
        return self.hwp.HAction.Run("InsertLastPrintDate")

    def InsertLastSaveBy(self) -> bool:
        """
        상용구 코드 넣기(마지막 저장한 사람)

        """
        return self.hwp.HAction.Run("InsertLastSaveBy")

    def InsertLastSaveDate(self) -> bool:
        """
        상용구 코드 넣기(마지막 저장한 날짜)

        """
        return self.hwp.HAction.Run("InsertLastSaveDate")

    def InsertLine(self) -> bool:
        """
        선 넣기

        """
        return self.hwp.HAction.Run("InsertLine")

    def InsertNonBreakingSpace(self) -> bool:
        """
        묶음 빈칸 삽입

        """
        return self.hwp.HAction.Run("InsertNonBreakingSpace")

    def InsertPageNum(self) -> bool:
        """
        쪽 번호 넣기

        """
        return self.hwp.HAction.Run("InsertPageNum")

    def InsertSoftHyphen(self) -> bool:
        """
        하이픈 삽입

        """
        return self.hwp.HAction.Run("InsertSoftHyphen")

    def InsertSpace(self) -> bool:
        """
        공백 삽입

        """
        return self.hwp.HAction.Run("InsertSpace")

    def InsertStringDateTime(self) -> bool:
        """
        날짜/시간 넣기 - 문자열로 넣기([입력-날짜/시간-날짜/시간 문자열]메뉴와 동일)

        """
        return self.hwp.HAction.Run("InsertStringDateTime")

    def InsertTab(self) -> bool:
        """
        탭 삽입

        """
        return self.hwp.HAction.Run("InsertTab")

    def InsertTpNo(self) -> bool:
        """
        상용구 코드 넣기(전체 쪽수)

        """
        return self.hwp.HAction.Run("InsertTpNo")

    def Jajun(self) -> bool:
        """
        한자 자전

        """
        return self.hwp.HAction.Run("Jajun")

    def LabelAdd(self) -> bool:
        """
        라벨 새 쪽 추가하기

        """
        return self.hwp.HAction.Run("LabelAdd")

    def LabelTemplate(self) -> bool:
        """
        라벨 문서 만들기

        """
        return self.hwp.HAction.Run("LabelTemplate")

    def LeftShiftBlock(self) -> bool:
        """텍스트 블록 상태에서 블록 왼쪽에 있는 탭 또는 공백을 지운다."""
        return self.hwp.HAction.Run("LeftShiftBlock")

    def LeftTabFrameClose(self) -> bool:
        """왼쪽 작업창 감추기"""
        return self.hwp.HAction.Run("LeftTabFrameClose")

    def LinkTextBox(self) -> bool:
        """
        글상자 연결.

        글상자가 선택되지 않았거나, 캐럿이 글상자 내부에 있지 않으면 동작하지 않는다.

        """
        return self.hwp.HAction.Run("LinkTextBox")

    def MacroPause(self) -> bool:
        """
        매크로 실행 일시 중지 (정의/실행)

        """
        return self.hwp.HAction.Run("MacroPause")

    def MacroPlay1(self) -> bool:
        """
        매크로 1

        """
        return self.hwp.HAction.Run("MacroPlay1")

    def MacroPlay2(self) -> bool:
        """
        매크로 2

        """
        return self.hwp.HAction.Run("MacroPlay2")

    def MacroPlay3(self) -> bool:
        """
        매크로 3

        """
        return self.hwp.HAction.Run("MacroPlay3")

    def MacroPlay4(self) -> bool:
        """
        매크로 4

        """
        return self.hwp.HAction.Run("MacroPlay4")

    def MacroPlay5(self) -> bool:
        """
        매크로 5

        """
        return self.hwp.HAction.Run("MacroPlay5")

    def MacroPlay6(self) -> bool:
        """
        매크로 6

        """
        return self.hwp.HAction.Run("MacroPlay6")

    def MacroPlay7(self) -> bool:
        """
        매크로 7

        """
        return self.hwp.HAction.Run("MacroPlay7")

    def MacroPlay8(self) -> bool:
        """
        매크로 8

        """
        return self.hwp.HAction.Run("MacroPlay8")

    def MacroPlay9(self) -> bool:
        """
        매크로 9

        """
        return self.hwp.HAction.Run("MacroPlay9")

    def MacroPlay10(self) -> bool:
        """
        매크로 10

        """
        return self.hwp.HAction.Run("MacroPlay10")

    def MacroPlay11(self) -> bool:
        """
        매크로 11

        """
        return self.hwp.HAction.Run("MacroPlay11")

    def MacroRepeat(self) -> bool:
        """
        매크로 실행

        """
        return self.hwp.HAction.Run("MacroRepeat")

    def MacroStop(self) -> bool:
        """
        매크로 실행 중지 (정의/실행)

        """
        return self.hwp.HAction.Run("MacroStop")

    def MailMergeField(self) -> bool:
        """
        메일 머지 필드(표시달기 or 고치기)

        """
        return self.hwp.HAction.Run("MailMergeField")

    def MakeIndex(self) -> bool:
        """
        찾아보기 만들기

        """
        return self.hwp.HAction.Run("MakeIndex")

    def ManualChangeHangul(self) -> bool:
        """
        한영 수동 전환

        현재 커서위치 또는 문단나누기 이전에 입력된 내용에 대해서 강제적으로 한/영 전환을 한다.

        """
        return self.hwp.HAction.Run("ManualChangeHangul")

    def MarkPenColor(self) -> bool:
        """형광펜 색"""
        return self.hwp.HAction.Run("MarkPenColor")

    def MarkPenDelete(self) -> bool:
        """형광펜 삭제"""
        return self.hwp.HAction.Run("MarkPenDelete")

    def MarkPenNext(self) -> bool:
        """
        다음 형광펜 삽입 위치로 이동한다.
        """
        return self.hwp.HAction.Run("MarkPenNext")

    def MarkPenPrev(self) -> bool:
        """
        이전 형광펜 삽입 위치로 이동한다.
        """
        return self.hwp.HAction.Run("MarkPenPrev")

    def MemoToNext(self) -> bool:
        """메모 편집 상태에서 다음 메모로 이동"""
        return self.hwp.HAction.Run("MemoToNext")

    def MemoToPrev(self) -> bool:
        """메모 편집 상태에서 이전 메모로 이동"""
        return self.hwp.HAction.Run("MemoToNext")

    def MetatagExist(self, tag):
        """특정 이름의 메타태그가 존재하는지?"""
        return self.hwp.MetatagExist(tag=tag)

    def MarkTitle(self) -> bool:
        """
        제목 차례 표시([도구-차례/찾아보기-제목 차례 표시]메뉴에 대응).

        차례 코드가 삽입되어 나중에 차례 만들기에서 사용할 수 있다.
        적용여부는 Ctrl+G,C를 이용해 조판부호를 확인하면 알 수 있다.

        """
        return self.hwp.HAction.Run("MarkTitle")

    def MasterPage(self) -> bool:
        """
        바탕쪽 진입

        """
        return self.hwp.HAction.Run("MasterPage")

    def MasterPageDuplicate(self) -> bool:
        """
        기존 바탕쪽과 겹침.

        바탕쪽 편집상태가 활성화되어 있으며 [구역 마지막쪽], [구역임의 쪽]일 경우에만 사용 가능하다.

        """
        return self.hwp.HAction.Run("MasterPageDuplicate")

    def MasterPageExcept(self) -> bool:
        """
        첫 쪽 제외

        """
        return self.hwp.HAction.Run("MasterPageExcept")

    def MasterPageFront(self) -> bool:
        """
        바탕쪽 앞으로 보내기.

        바탕쪽 편집모드일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("MasterPageFront")

    def MasterPagePrevSection(self) -> bool:
        """
        앞 구역 바탕쪽 사용

        """
        return self.hwp.HAction.Run("MasterPagePrevSection")

    def MasterPageToNext(self) -> bool:
        """
        이후 바탕쪽

        """
        return self.hwp.HAction.Run("MasterPageToNext")

    def MasterPageToPrevious(self) -> bool:
        """
        이전 바탕쪽

        """
        return self.hwp.HAction.Run("MasterPageToPrevious")

    def MasterPageType(self) -> bool:
        """바탕쪽 종류"""
        return self.hwp.HAction.Run("MasterPageType")

    def MasterWsItemOnOff(self) -> bool:
        """바탕쪽 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("MasterWsItemOnOff")

    def ModifyComposeChars(self) -> bool:
        """
        고치기 - 글자 겹침

        """
        return self.hwp.HAction.Run("ModifyComposeChars")

    def ModifyCtrl(self) -> bool:
        """
        고치기 : 컨트롤

        """
        return self.hwp.HAction.Run("ModifyCtrl")

    def ModifyDutmal(self) -> bool:
        """
        고치기 - 덧말

        """
        return self.hwp.HAction.Run("ModifyDutmal")

    def ModifyFillProperty(self) -> bool:
        """
        고치기(채우기 속성 탭으로).

        만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다.
        채우기 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다.

        """
        return self.hwp.HAction.Run("ModifyFillProperty")

    def ModifyLineProperty(self) -> bool:
        """
        고치기(선/테두리 속성 탭으로).

        만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다.
        선/테두리 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다.

        """
        return self.hwp.HAction.Run("ModifyLineProperty")

    def ModifyShapeObject(self) -> bool:
        """
        고치기 - 개체 속성

        """
        return self.hwp.HAction.Run("ModifyShapeObject")

    def MoveColumnBegin(self) -> bool:
        """
        단의 시작점으로 이동

        단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다.

        """
        return self.hwp.HAction.Run("MoveColumnBegin")

    def MoveColumnEnd(self) -> bool:
        """
        단의 끝점으로 이동한다.

        단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다.

        """
        return self.hwp.HAction.Run("MoveColumnEnd")

    def MoveDocBegin(self) -> bool:
        """
        문서의 시작으로 이동

        만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다.
        현재 서브 리스트 내에 있으면 빠져나간다. 자동화에 아주 많이 사용된다.

        """
        return self.hwp.HAction.Run("MoveDocBegin")

    def MoveDocEnd(self) -> bool:
        """
        문서의 끝으로 이동

        만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다.
        현재 서브 리스트 내에 있으면 빠져나간다.

        """
        return self.hwp.HAction.Run("MoveDocEnd")

    def PutParaNumber(self) -> bool:
        """
        문단번호 삽입/제거 토글

        """
        return self.hwp.HAction.Run("PutParaNumber")

    def PutOutlinleNumber(self) -> bool:
        """
        개요번호 삽입/제거 토글

        """
        return self.hwp.HAction.Run("PutOutlineNumber")

    def Close(self) -> bool:
        """
        현재 리스트를 닫고 (최)상위 리스트로 이동하는 액션.

        대표적인 예로, 메모나 각주 등을 작성한 후 본문으로 빠져나올 때, 혹은 여러 겹의 표 안에 있을 때 한 번에 표 밖으로 캐럿을 옮길 때 사용한다. 굉장히 자주 쓰이는 액션이며, 경우에 따라 Close가 아니라 CloseEx를 써야 하는 경우도 있다.
        (레퍼런스 포인트가 등록되어 있으면 그 포인트로, 없으면 루트 리스트로 이동한다. 나머지 특성은 MoveRootList와 동일)
        """
        cur_pos = self.get_pos()
        try:
            return self.hwp.HAction.Run("Close")
        finally:
            for _ in range(5):
                if self.get_pos() != cur_pos:
                    break
                else:
                    sleep(0.05)

    def DeleteDocumentMasterPage(self):
        """
        문서의 마스터페이지 삭제하기
        """
        self.hwp.HAction.Run("MasterPage")
        cur_messagebox_mode = self.get_message_box_mode()
        self.set_message_box_mode(0x10001)
        try:
            return self.hwp.HAction.Run("DeleteDocumentMasterPage")
        finally:
            self.set_message_box_mode(cur_messagebox_mode)

    def DeleteSectionMasterPage(self):
        """
        섹션의 마스터페이지 삭제하기
        """
        self.hwp.HAction.Run("MasterPage")
        cur_messagebox_mode = self.get_message_box_mode()
        self.set_message_box_mode(0x10001)
        try:
            return self.hwp.HAction.Run("DeleteSectionMasterPage")
        finally:
            self.set_message_box_mode(cur_messagebox_mode)

    def UnSelectCtrl(self):
        """선택중인 컨트롤 선택해제"""
        return self.hwp.UnSelectCtrl()

    def WindowAlignCascade(self):
        """창 겹치게 배열"""
        return self.hwp.HAction.Run("WindowAlignCascade")

    def WindowAlignTileHorz(self):
        """창 가로로 배열"""
        return self.hwp.HAction.Run("WindowAlignTileHorz")

    def WindowAlignTileVert(self):
        """창 세로로 배열"""
        return self.hwp.HAction.Run("WindowAlignTileVert")

    def WindowList(self):
        """창 목록"""
        return self.hwp.HAction.Run("WindowList")

    def WindowMinimizeAll(self):
        """창 모두 아이콘으로 배열"""
        return self.hwp.HAction.Run("WindowMinimizeAll")

    def WindowNextPane(self):
        """다음 분할창 활성화"""
        return self.hwp.HAction.Run("WindowNextPane")

    def WindowNextTab(self):
        """다음 창 활성화"""
        return self.hwp.HAction.Run("WindowNextTab")

    def WindowPrevTab(self):
        """이전 창 활성화"""
        return self.hwp.HAction.Run("WindowPrevTab")

    def Cut(self, remove_cell=True) -> bool:
        """
        잘라내기.

        Copy 액션과 유사하지만,
        복사 대신 잘라내기 기능을 수행한다.
        유용하지만, 빠른 반복작업에서는 윈도우OS의 클립보드 잠금 때문에 오류가 발생할 수 있다.

        """
        if remove_cell:
            self.set_message_box_mode(0x2000)
            try:
                return self.hwp.HAction.Run("Cut")
            finally:
                self.set_message_box_mode(0xF000)
        else:
            self.set_message_box_mode(0x1000)
            try:
                return self.hwp.HAction.Run("Cut")
            finally:
                self.set_message_box_mode(0xF000)

    def Delete(self, delete_ctrl: bool = True) -> bool:
        """
        삭제(Del키).

        키보드의 Del 키를 눌렀을 때와 거의(?) 유사하다.
        아주 사용빈도가 높은 액션이다.
        수작업과 달리 표나 이미지 등의 컨트롤을 삭제할 때에도 팝업으로 묻지 않는다.
        (``hwp.Delete(False)``로 실행하면 경고팝업이 뜬다.)

        Args:
            delete_ctrl: 컨트롤(표, 이미지, 겹침문자 등)을 삭제할 때 처리방법(True: 삭제(기본값), False: 삭제안함)

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        cur_mode = self.hwp.GetMessageBoxMode()
        if delete_ctrl:
            self.hwp.SetMessageBoxMode(0x10)
        else:
            self.hwp.SetMessageBoxMode(0x00)
        try:
            return self.hwp.HAction.Run("Delete")
        finally:
            self.hwp.SetMessageBoxMode(cur_mode)

    def DeleteBack(self, delete_ctrl: bool = True) -> bool:
        """
        백스페이스 삭제

        Delete와 유사하지만, 이건 Backspace처럼 우측에서 좌측으로 삭제해준다. 많이 쓰인다.

        Args:
            delete_ctrl: 컨트롤 삭제시 경고팝업을 띄울지(기본값은 True)

        """
        cur_mode = self.hwp.GetMessageBoxMode()
        if delete_ctrl:
            self.hwp.SetMessageBoxMode(0x10)
        else:
            self.hwp.SetMessageBoxMode(0x00)
        try:
            return self.hwp.HAction.Run("DeleteBack")
        finally:
            self.hwp.SetMessageBoxMode(cur_mode)

    def DeleteLine(self, delete_ctrl=True):
        """
        한 줄 지우기(Ctrl-Y) 액션.

        문단나눔과 전혀 상관없이 딱 한 줄의 텍스트가 삭제된다.
        원래 액션과 달리 DeleteLine으로 표 등의 객체를 삭제하는 경우에
        경고팝업이 뜨지 않으므로 유의해야 한다.
        만약 컨트롤을 지우고 싶지 않다면 인수에 False를 넣으면 된다.

        """
        cur_mode = self.hwp.GetMessageBoxMode()
        if delete_ctrl:
            self.hwp.SetMessageBoxMode(0x10)
        else:
            self.hwp.SetMessageBoxMode(0x00)
        try:
            return self.hwp.HAction.Run("DeleteLine")
        finally:
            self.hwp.SetMessageBoxMode(cur_mode)

    def DeleteLineEnd(self, delete_ctrl=True):
        """
        현재 커서에서 줄 끝까지 지우기(Alt-Y).

        수작업시에 굉장히 유용한 기능일 수 있지만,
        자동화 작업시에는 DeleteLine이나 DeleteLineEnd 모두,
        한 줄 안에 어떤 내용까지 있는지 파악하기 어려운 관계로,
        자동화에 잘 쓰이지는 않는다.
        원래 액션과 달리 DeleteLineEnd로 표 등의 객체를 삭제하는 경우에
        경고팝업이 뜨지 않으므로 유의해야 한다.
        만약 컨트롤을 지우고 싶지 않다면 인수에 False를 넣으면 된다.
        """
        cur_mode = self.hwp.GetMessageBoxMode()
        if delete_ctrl:
            self.hwp.SetMessageBoxMode(0x10)
        else:
            self.hwp.SetMessageBoxMode(0x00)
        try:
            return self.hwp.HAction.Run("DeleteLineEnd")
        finally:
            self.hwp.SetMessageBoxMode(cur_mode)

    def DeleteWord(self, delete_ctrl=True):
        """
        단어 지우기(Ctrl-T) 액션.

        단, 커서 우측에 위치한 단어 한 개씩 삭제하며, 커서가 단어 중간에 있는 경우 우측 글자만 삭제한다.

        """
        cur_mode = self.hwp.GetMessageBoxMode()
        if delete_ctrl:
            self.hwp.SetMessageBoxMode(0x10)
        else:
            self.hwp.SetMessageBoxMode(0x00)
        try:
            return self.hwp.HAction.Run("DeleteWord")
        finally:
            self.hwp.SetMessageBoxMode(cur_mode)

    def DeleteWordBack(self, delete_ctrl=True):
        """
        한 단어씩 좌측으로 삭제하는 액션(Ctrl-백스페이스).

        DeleteWord와 마찬가지로 커서가 단어 중간에 있는 경우
        좌측 글자만 삭제한다.

        """
        cur_mode = self.hwp.GetMessageBoxMode()
        if delete_ctrl:
            self.hwp.SetMessageBoxMode(0x10)
        else:
            self.hwp.SetMessageBoxMode(0x00)
        try:
            return self.hwp.HAction.Run("DeleteWordBack")
        finally:
            self.hwp.SetMessageBoxMode(cur_mode)

    def MoveDown(self):
        """
        캐럿을 (논리적 개념의) 아래로 이동시킨다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveLeft(self):
        """
        캐럿을 (논리적 개념의) 왼쪽으로 이동시킨다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveLeft")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveLineBegin(self):
        """
        현재 위치한 줄의 시작/끝으로 이동

        """
        return self.hwp.HAction.Run("MoveLineBegin")

    def MoveLineDown(self):
        """
        한 줄 아래로 이동한다.

        """
        return self.hwp.HAction.Run("MoveLineDown")

    def MoveLineEnd(self):
        """
        현재 위치한 줄의 시작/끝으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveLineEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveLineUp(self):
        """
        한 줄 위로 이동한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveLineUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveListBegin(self):
        """
        현재 리스트의 시작으로 이동

        """
        return self.hwp.HAction.Run("MoveListBegin")

    def MoveListEnd(self):
        """
        현재 리스트의 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveListEnd")

    def MoveNextChar(self):
        """
        한 글자 뒤로 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextChar")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveNextColumn(self):
        """
        뒤 단으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextColumn")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveNextParaBegin(self):
        """
        앞 문단의 끝/다음 문단의 시작으로 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextParaBegin")
        if self.get_pos()[1] != cwd[1]:
            return True
        else:
            return False

    def MoveNextPos(self):
        """
        한 글자 뒤로 이동.

        서브 리스트를 옮겨 다닐 수 있다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveNextPosEx(self):
        """
        한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다.

        (머리말, 꼬리말, 각주, 미주, 글상자 포함)
        예를 들어, 문단 중간에 글상자가 (글자처럼취급 꺼진상태로) 떠있다면
        MoveNextPos는 글상자를 패스하고 본문만 통과해서 지나가는데,
        MoveNextPosEx는 캐럿이 컨트롤을 만나는 시점에 글상자 안으로 들어갔다 나온다.
        문서 전체를 훑어야 하는 경우에는 굉장히 유용하게 쓰일 듯.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextPosEx")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveNextWord(self):
        """
        한 단어 뒤로 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePageBegin(self):
        """
        현재 페이지의 시작점으로 이동한다.

        만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 상단으로 페이지뷰잉이 맞춰진다.

        """
        return self.hwp.HAction.Run("MovePageBegin")

    def MovePageDown(self):
        """
        앞 페이지의 시작으로 이동.

        현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePageDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePageEnd(self):
        """
        현재 페이지의 끝점으로 이동한다.

        만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 하단으로 페이지뷰잉이 맞춰진다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePageEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePageUp(self):
        """
        뒤 페이지의 시작으로 이동.

        현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePageUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveParaBegin(self):
        """
        현재 위치한 문단의 시작/끝으로 이동

        """
        return self.hwp.HAction.Run("MoveParaBegin")

    def MoveParaEnd(self):
        """
        현재 위치한 문단의 시작/끝으로 이동

        """
        return self.hwp.HAction.Run("MoveParaEnd")

    def MoveParentList(self):
        """
        한 레벨 상위/탑레벨/루트 리스트로 이동한다

        현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다.
        이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다.
        위치 이동시 셀렉션은 무조건 풀린다.

        """
        return self.hwp.HAction.Run("MoveParentList")

    def MovePrevChar(self):
        """
        한 글자 앞 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevChar")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevColumn(self):
        """
        앞 단으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevColumn")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevParaBegin(self):
        """
        앞 문단의 시작으로 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevParaBegin")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevParaEnd(self):
        """
        앞 문단의 끝으로 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevParaEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevPos(self):
        """
        한 글자 앞으로 이동.

        서브 리스트를 옮겨 다닐 수 있다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevPosEx(self):
        """
        한 글자 앞으로 이동.

        서브 리스트를 옮겨 다닐 수 있다.
        (머리말, 꼬리말, 각주, 미주, 글상자 포함)
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevPosEx")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevWord(self):
        """
        한 단어 앞으로 이동.

        현재 리스트만을 대상으로 동작한다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveRight(self):
        """
        캐럿을 (논리적 개념의) 오른쪽으로 이동시킨다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveRight")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveRootList(self):
        """
        한 레벨 상위/탑레벨/루트 리스트로 이동한다

        현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다.
        이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다.
        위치 이동시 셀렉션은 무조건 풀린다.

        """
        return self.hwp.HAction.Run("MoveRootList")

    def MoveScrollDown(self):
        """
        아래 방향으로 스크롤하면서 이동

        """
        return self.hwp.HAction.Run("MoveScrollDown")

    def MoveScrollNext(self):
        """
        다음 방향으로 스크롤하면서 이동

        """
        return self.hwp.HAction.Run("MoveScrollNext")

    def MoveScrollPrev(self):
        """
        이전 방향으로 스크롤하면서 이동

        """
        return self.hwp.HAction.Run("MoveScrollPrev")

    def MoveScrollUp(self):
        """
        위 방향으로 스크롤하면서 이동

        """
        return self.hwp.HAction.Run("MoveScrollUp")

    def MoveSectionDown(self):
        """
        뒤 섹션으로 이동.

        현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSectionDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSectionUp(self):
        """
        앞 섹션으로 이동.

        현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSectionUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelDocBegin(self):
        """
        선택 상태로 문서 처음으로 이동

        """
        return self.hwp.HAction.Run("MoveSelDocBegin")

    def MoveSelDocEnd(self):
        """
        선택 상태로 문서 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveSelDocEnd")

    def MoveSelDown(self):
        """
        선택 상태로 캐럿을 (논리적 방향) 아래로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelLeft(self):
        """
        선택 상태로 캐럿을 (논리적 방향) 왼쪽으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelLeft")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelLineBegin(self):
        """
        선택 상태로 줄 처음으로 이동

        """
        return self.hwp.HAction.Run("MoveSelLineBegin")

    def MoveSelLineDown(self):
        """
        선택 상태로 한줄 아래로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelLineDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelLineEnd(self):
        """
        선택 상태로 줄 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveSelLineEnd")

    def MoveSelLineUp(self):
        """
        선택 상태로 한줄 위로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelLineUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelListBegin(self):
        """
        선택 상태로 리스트 처음으로 이동

        """
        return self.hwp.HAction.Run("MoveSelListBegin")

    def MoveSelListEnd(self):
        """
        선택 상태로 리스트 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveSelListEnd")

    def MoveSelNextChar(self):
        """
        선택 상태로 다음 글자로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextChar")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelNextParaBegin(self):
        """
        선택 상태로 다음 문단 처음으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextParaBegin")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelNextPos(self):
        """
        선택 상태로 다음 커서위치(글자)로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelNextWord(self):
        """
        선택 상태로 다음 단어로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPageDown(self):
        """
        선택 상태로 PageDown 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPageDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPageUp(self):
        """
        선택 상태로 PageUp 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPageUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelParaBegin(self):
        """
        선택 상태로 문단 처음으로 이동

        """
        return self.hwp.HAction.Run("MoveSelParaBegin")

    def MoveSelParaEnd(self):
        """
        선택 상태로 문단 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveSelParaEnd")

    def MoveSelPrevChar(self):
        """
        선택 상태로 이전 글자로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevChar")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevParaBegin(self):
        """
        선택 상태로 이전 문단 시작으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevParaBegin")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevParaEnd(self):
        """
        선택 상태로 이전 문단 끝으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevParaEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevPos(self):
        """
        선택 상태로 이전 위치(글자)로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevWord(self):
        """
        선택 상태로 이전 단어로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelRight(self):
        """
        선택 상태로 캐럿을 (논리적 방향) 오른쪽으로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelRight")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelTopLevelBegin(self):
        """
        선택 상태로 문서의 탑레벨 처음으로 이동(==``hwp.MoveDocBegin()``)

        """
        return self.hwp.HAction.Run("MoveSelTopLevelBegin")

    def MoveSelTopLevelEnd(self):
        """
        선택 상태로 탑레벨 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveSelTopLevelEnd")

    def MoveSelUp(self):
        """
        선택 상태로 캐럿을 (논리적 방향) 위로 이동

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelViewDown(self):
        """
        선택 상태로 화면 시점의 아래로 이동

        Shift-PgDn 기능과 동일

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelViewDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelViewUp(self):
        """
        선택 상태로 화면 시점의 위로 이동

        Shift-PgUp 기능과 동일

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelViewUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelWordBegin(self):
        """
        선택 상태로 단어 처음으로 이동

        """
        return self.hwp.HAction.Run("MoveSelWordBegin")

    def MoveSelWordEnd(self):
        """
        선택 상태로 단어 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveSelWordEnd")

    def MoveTopLevelBegin(self):
        """
        탑레벨 리스트의 시작으로 이동

        """
        return self.hwp.HAction.Run("MoveTopLevelBegin")

    def MoveTopLevelEnd(self):
        """
        탑레벨 리스트의 끝으로 이동

        """
        return self.hwp.HAction.Run("MoveTopLevelEnd")

    def MoveTopLevelList(self):
        """
        한 레벨 상위/탑레벨/루트 리스트로 이동

        현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다.
        이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다.
        위치 이동시 셀렉션은 무조건 풀린다.

        """
        return self.hwp.HAction.Run("MoveTopLevelList")

    def MoveUp(self):
        """
        캐럿을 (논리적 개념의) 위로 이동시킨다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveViewBegin(self):
        """
        현재 뷰의 시작에 위치한 곳으로 이동

        """
        return self.hwp.HAction.Run("MoveViewBegin")

    def MoveViewDown(self):
        """
        현재 뷰의 크기만큼 아래로 이동한다. PgDn 키의 기능이다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveViewDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveViewEnd(self):
        """
        현재 뷰의 끝에 위치한 곳으로 이동

        """
        return self.hwp.HAction.Run("MoveViewEnd")

    def MoveViewUp(self):
        """
        현재 뷰의 크기만큼 위로 이동한다. PgUp 키의 기능이다.

        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveViewUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveWordBegin(self):
        """
        현재 위치한 단어의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.

        """
        return self.hwp.HAction.Run("MoveWordBegin")

    def MoveWordEnd(self):
        """
        현재 위치한 단어의 끝으로 이동. 현재 리스트만을 대상으로 동작한다.

        """
        return self.hwp.HAction.Run("MoveWordEnd")

    def MPSectionToNext(self):
        """
        이후 구역으로

        """
        return self.hwp.HAction.Run("MPSectionToNext")

    def MPSectionToPrevious(self):
        """
        이전 구역으로

        """
        return self.hwp.HAction.Run("MPSectionToPrevious")

    def MPShowMarginBorder(self) -> bool:
        """바탕쪽 편집 상태에서 여백 보기 토글"""
        return self.hwp.HAction.Run("MPShowMarginBorder")

    def NextTextBoxLinked(self):
        """
        연결된 글상자의 다음 글상자로 이동

        """
        return self.hwp.HAction.Run("NextTextBoxLinked")

    def NoteDelete(self):
        """
        주석 지우기

        """
        return self.hwp.HAction.Run("NoteDelete")

    def NoSplit(self):
        """창 나누지 않음"""
        return self.hwp.HAction.Run("NoSplit")

    def NoteLineColor(self):
        """주석 구분선 색"""
        return self.hwp.HAction.Run("NoteLineColor")

    def NoteLineLength(self):
        """주석 구분선 길이"""
        return self.hwp.HAction.Run("NoteLineLength")

    def NoteLineShape(self):
        """주석 구분선 모양"""
        return self.hwp.HAction.Run("NoteLineShape")

    def NoteLineWeight(self):
        """주석 구분선 굵기"""
        return self.hwp.HAction.Run("NoteLineWeight")

    def NoteModify(self):
        """
        주석 고치기

        """
        return self.hwp.HAction.Run("NoteModify")

    def NoteNumProperty(self):
        """
        주석 번호 속성

        """
        return self.hwp.HAction.Run("NoteNumProperty")

    def NoteNumShape(self):
        """주석 번호 모양"""
        return self.hwp.HAction.Run("NoteNumShape")

    def NotePosition(self):
        """각주 위치"""
        return self.hwp.HAction.Run("NotePosition")

    def NoteToNext(self):
        """
        주석 다음으로 이동

        """
        return self.hwp.HAction.Run("NoteToNext")

    def NoteToPrev(self):
        """
        주석 앞으로 이동

        """
        return self.hwp.HAction.Run("NoteToPrev")

    def ParagraphShapeAlignCenter(self):
        """
        가운데 정렬

        """
        return self.hwp.HAction.Run("ParagraphShapeAlignCenter")

    def ParagraphShapeAlignDistribute(self):
        """
        배분 정렬

        """
        return self.hwp.HAction.Run("ParagraphShapeAlignDistribute")

    def ParagraphShapeAlignDivision(self):
        """
        나눔 정렬

        """
        return self.hwp.HAction.Run("ParagraphShapeAlignDivision")

    def ParagraphShapeAlignJustify(self):
        """
        양쪽 정렬

        """
        return self.hwp.HAction.Run("ParagraphShapeAlignJustify")

    def ParagraphShapeAlignLeft(self):
        """
        왼쪽 정렬

        """
        return self.hwp.HAction.Run("ParagraphShapeAlignLeft")

    def ParagraphShapeAlignRight(self):
        """
        오른쪽 정렬

        """
        return self.hwp.HAction.Run("ParagraphShapeAlignRight")

    def ParagraphShapeDecreaseLeftMargin(self):
        """
        왼쪽 여백 줄이기

        """
        return self.hwp.HAction.Run("ParagraphShapeDecreaseLeftMargin")

    def ParagraphShapeDecreaseLineSpacing(self):
        """
        줄 간격을 점점 좁힘

        """
        return self.hwp.HAction.Run("ParagraphShapeDecreaseLineSpacing")

    def ParagraphShapeDecreaseMargin(self):
        """
        왼쪽-오른쪽 여백 줄이기

        """
        return self.hwp.HAction.Run("ParagraphShapeDecreaseMargin")

    def ParagraphShapeDecreaseRightMargin(self):
        """
        오른쪽 여백 키우기

        """
        return self.hwp.HAction.Run("ParagraphShapeDecreaseRightMargin")

    def ParagraphShapeIncreaseLeftMargin(self):
        """
        왼쪽 여백 키우기

        """
        return self.hwp.HAction.Run("ParagraphShapeIncreaseLeftMargin")

    def ParagraphShapeIncreaseLineSpacing(self):
        """
        줄 간격을 점점 넓힘

        """
        return self.hwp.HAction.Run("ParagraphShapeIncreaseLineSpacing")

    def ParagraphShapeIncreaseMargin(self):
        """
        왼쪽-오른쪽 여백 키우기

        """
        return self.hwp.HAction.Run("ParagraphShapeIncreaseMargin")

    def ParagraphShapeIncreaseRightMargin(self):
        """
        오른쪽 여백 줄이기

        """
        return self.hwp.HAction.Run("ParagraphShapeIncreaseRightMargin")

    def ParagraphShapeIndentAtCaret(self):
        """
        첫 줄 내어 쓰기

        """
        return self.hwp.HAction.Run("ParagraphShapeIndentAtCaret")

    def ParagraphShapeIndentNegative(self):
        """
        첫 줄을 한 글자 내어 씀

        """
        return self.hwp.HAction.Run("ParagraphShapeIndentNegative")

    def ParagraphShapeIndentPositive(self):
        """
        첫 줄을 한 글자 들여 씀

        """
        return self.hwp.HAction.Run("ParagraphShapeIndentPositive")

    def ParagraphShapeProtect(self):
        """
        문단 보호
        """
        return self.hwp.HAction.Run("ParagraphShapeProtect")

    def ParagraphShapeSingleRow(self) -> bool:
        """
        문단 한 줄로 입력 토글
        """
        return self.hwp.HAction.Run("ParagraphShapeSingleRow")

    def ParagraphShapeWithNext(self):
        """
        다음 문단과 함께

        """
        return self.hwp.HAction.Run("ParagraphShapeWithNext")

    def ParaShapeLineSpace(self):
        """문단 모양"""
        return self.hwp.HAction.Run("ParaShapeLineSpace")

    def Paste(self):
        """
        일반 붙이기(Ctrl-V)

        확장 붙여넣기는 hwp.paste() 사용.
        """
        return self.hwp.HAction.Run("Paste")

    def PastePage(self):
        """
        쪽 붙여넣기

        한글2014 이하의 버전에서는 작동하지 않는다.

        """
        return self.hwp.HAction.Run("PastePage")

    def PasteSpecial(self):
        """
        골라 붙이기

        """
        return self.hwp.HAction.Run("PasteSpecial")

    def PictureEffect1(self):
        """
        그림 그레이 스케일

        """
        return self.hwp.HAction.Run("PictureEffect1")

    def PictureEffect2(self):
        """
        그림 흑백으로

        """
        return self.hwp.HAction.Run("PictureEffect2")

    def PictureEffect3(self):
        """
        그림 워터마크

        """
        return self.hwp.HAction.Run("PictureEffect3")

    def PictureEffect4(self):
        """
        그림 효과 없음

        """
        return self.hwp.HAction.Run("PictureEffect4")

    def PictureEffect5(self):
        """
        그림 밝기 증가

        """
        return self.hwp.HAction.Run("PictureEffect5")

    def PictureEffect6(self):
        """
        그림 밝기 감소

        """
        return self.hwp.HAction.Run("PictureEffect6")

    def PictureEffect7(self):
        """
        그림 명암 증가

        """
        return self.hwp.HAction.Run("PictureEffect7")

    def PictureEffect8(self):
        """
        그림 명암 감소

        """
        return self.hwp.HAction.Run("PictureEffect8")

    def PictureInsertDialog(self):
        """
        그림 넣기 대화상자

        (대화상자를 띄워 선택한 이미지 파일을 문서에 삽입하는 액션 : API용)

        """
        return self.hwp.HAction.Run("PictureInsertDialog")

    def PictureLinkedToEmbedded(self):
        """
        연결된 그림을 모두 삽입그림으로

        """
        return self.hwp.HAction.Run("PictureLinkedToEmbedded")

    def PictureSave(self):
        """
        그림 빼내기

        """
        return self.hwp.HAction.Run("PictureSave")

    def PictureScissor(self):
        """
        그림 자르기

        """
        return self.hwp.HAction.Run("PictureScissor")

    def PictureToOriginal(self):
        """
        그림을 원래 그림으로

        """
        return self.hwp.HAction.Run("PictureToOriginal")

    def PrevTextBoxLinked(self):
        """
        연결된 글상자의 이전 글상자로 이동.

        현재 글상자가 선택되거나, 글상자 내부에 캐럿이 존재하지 않으면 동작하지 않는다.

        """
        return self.hwp.HAction.Run("PrevTextBoxLinked")

    def PstAutoPlay(self):
        """프리젠테이션 자동 시연"""
        return self.hwp.HAction.Run("PstAutoPlay")

    def PstBlackToWhite(self):
        """프리젠테이션 검은색 글자를 흰색으로 변경"""
        return self.hwp.HAction.Run("PstBlackToWhite")

    def PstGradientType(self):
        """프리젠테이션 그라데이션 형태"""
        return self.hwp.HAction.Run("PstGradientType")

    def PstScrChangeType(self):
        """프리젠테이션 화면 전환 형태"""
        return self.hwp.HAction.Run("PstScrChangeType")

    def PstSetupNextSec(self):
        """프리젠테이션 뒤 구역 설정"""
        return self.hwp.HAction.Run("PstSetupNextSec")

    def PstSetupPrevSec(self):
        """프리젠테이션 앞 구역 설정"""
        return self.hwp.HAction.Run("PstSetupPrevSec")

    def QuickCommandRun(self):
        """
        입력 자동 명령 동작

        """
        return self.hwp.HAction.Run("QuickCommand Run")

    def QuickCorrect(self):
        """
        빠른 교정 (실질적인 동작 Action)
        """
        return self.hwp.HAction.Run("QuickCorrect")

    def QuickCorrectRun(self):
        """
        빠른 교정 ― 내용 편집
        """
        return self.hwp.HAction.Run("QuickCorrect Run")

    def QuickCorrectSound(self):
        """
        빠른 교정 ― 메뉴에서 효과음 On/Off
        """
        return self.hwp.HAction.Run("QuickCorrect Sound")

    def QuickMarkInsert0(self):
        """
        쉬운 책갈피0 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert0")

    def QuickMarkInsert1(self):
        """
        쉬운 책갈피1 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert1")

    def QuickMarkInsert2(self):
        """
        쉬운 책갈피2 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert2")

    def QuickMarkInsert3(self):
        """
        쉬운 책갈피3 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert3")

    def QuickMarkInsert4(self):
        """
        쉬운 책갈피4 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert4")

    def QuickMarkInsert5(self):
        """
        쉬운 책갈피5 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert5")

    def QuickMarkInsert6(self):
        """
        쉬운 책갈피6 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert6")

    def QuickMarkInsert7(self):
        """
        쉬운 책갈피7 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert7")

    def QuickMarkInsert8(self):
        """
        쉬운 책갈피8 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert8")

    def QuickMarkInsert9(self):
        """
        쉬운 책갈피9 - 삽입

        """
        return self.hwp.HAction.Run("QuickMarkInsert9")

    def QuickMarkMove0(self):
        """
        쉬운 책갈피0 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove0")

    def QuickMarkMove1(self):
        """
        쉬운 책갈피1 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove1")

    def QuickMarkMove2(self):
        """
        쉬운 책갈피2 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove2")

    def QuickMarkMove3(self):
        """
        쉬운 책갈피3 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove3")

    def QuickMarkMove4(self):
        """
        쉬운 책갈피4 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove4")

    def QuickMarkMove5(self):
        """
        쉬운 책갈피5 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove5")

    def QuickMarkMove6(self):
        """
        쉬운 책갈피6 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove6")

    def QuickMarkMove7(self):
        """
        쉬운 책갈피7 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove7")

    def QuickMarkMove8(self):
        """
        쉬운 책갈피8 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove8")

    def QuickMarkMove9(self):
        """
        쉬운 책갈피9 - 이동

        """
        return self.hwp.HAction.Run("QuickMarkMove9")

    def RecalcPageCount(self):
        """
        현재 페이지의 쪽 번호 재계산

        """
        return self.hwp.HAction.Run("RecalcPageCount")

    def RecentCode(self):
        """
        최근에 사용한 문자표 입력.

        최근에 사용한 문자표가 없을 경우에는 문자표 대화상자를 띄운다.

        """
        return self.hwp.HAction.Run("RecentCode")

    def RecentEmpty(self):
        """최근 목록 지우기"""
        return self.hwp.HAction.Run("RecentEmpty")

    def RecentNoExistDel(self):
        """최근 목록에서 존재하지 않는 아이템 지우기"""
        return self.hwp.HAction.Run("RecentNoExistDel")

    def Redo(self):
        """
        다시 실행

        """
        return self.hwp.HAction.Run("Redo")

    def ReplacePrivateInfoDlg(self) -> bool:
        """
        개인정보 바꾸기 창 열고, 기타 바꾸기(문자열 치환)에 포커스
        """
        return self.hwp.HAction.Run("ReplacePrivateInfoDlg")

    def ReplyMemo(self) -> bool:
        """
        메모 회신(한/글2022부터 지원)
        """
        try:
            return self.hwp.HAction.Run("ReplyMemo")
        except:
            print("이 기능은 한/글2022부터 지원합니다.")
            return False

    def returnKeyInField(self):
        """
        캐럿이 필드 안에 위치한 상태에서 return Key에 대한 액션 분기
        첫 글자 r은 오타인 것으로 추정.
        """
        return self.hwp.HAction.Run("returnKeyInField")

    def ReturnKeyInField(self):
        """
        캐럿이 필드 안에 위치한 상태에서 return Key에 대한 액션 분기
        """
        return self.hwp.HAction.Run("returnKeyInField")

    def returnPrevPos(self):
        """
        직전위치로 돌아가기
        첫 글자 r은 오타인 것으로 추정.
        """
        return self.hwp.HAction.Run("returnPrevPos")

    def ReturnPrevPos(self):
        """
        직전위치로 돌아가기
        """
        return self.hwp.HAction.Run("returnPrevPos")

    def RightShiftBlock(self):
        """텍스트 블록 상태에서 블록이 문단의 시작위치에서 시작할 경우 블록 왼쪽에 탭을 삽입한다."""
        return self.hwp.HAction.Run("RightShiftBlock")

    def RightTabFrameClose(self):
        """오른쪽 작업창 감추기"""
        return self.hwp.HAction.Run("RightTabFrameClose")

    def RunUserKeyLayout(self):
        """사용자 글자판 제작 툴"""
        return self.hwp.HAction.Run("RunUserKeyLayout")

    def ScanHFTFonts(self) -> bool:
        """
        한/글 글꼴 검색
        """
        return self.hwp.HAction.Run("ScanHFTFonts")

    def ScrMacroPause(self):
        """
        매크로 기록 일시정지/재시작

        """
        return self.hwp.HAction.Run("ScrMacroPause")

    def ScrMacroPlay1(self):
        """
        1번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay1")

    def ScrMacroPlay2(self):
        """
        2번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay2")

    def ScrMacroPlay3(self):
        """
        3#번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay3")

    def ScrMacroPlay4(self):
        """
        4번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay4")

    def ScrMacroPlay5(self):
        """
        5번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay5")

    def ScrMacroPlay6(self):
        """
        6번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay6")

    def ScrMacroPlay7(self):
        """
        7번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay7")

    def ScrMacroPlay8(self):
        """
        8번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay8")

    def ScrMacroPlay9(self):
        """
        9번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay9")

    def ScrMacroPlay10(self):
        """
        10번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay10")

    def ScrMacroPlay11(self):
        """
        11번 매크로 실행(Alt+Shift+#)

        """
        return self.hwp.HAction.Run("ScrMacroPlay11")

    def ScrMacroRepeat(self):
        """스크립트 매크로 실행"""
        return self.hwp.HAction.Run("ScrMacroRepeat")

    def ScrMacroStop(self):
        """
        매크로 기록 중지

        """
        return self.hwp.HAction.Run("ScrMacroStop")

    def SearchPrivateInfo(self) -> bool:
        """
        개인정보 보호하기 창 열기(개인 정보 찾아 감추기-암호화?)
        """
        return self.hwp.HAction.Run("SearchPrivateInfo")

    def Select(self):
        """
        선택 (F3 Key를 누른 효과)

        """
        return self.hwp.HAction.Run("Select")

    def SelectAll(self):
        """
        모두 선택(Ctrl-A)

        """
        return self.hwp.HAction.Run("SelectAll")

    def SelectColumn(self):
        """
        칸 블록 선택 (F4 Key를 누른 효과)

        """
        return self.hwp.HAction.Run("SelectColumn")

    def SelectCtrlFront(self):
        """
        정방향으로 가장 가까운 컨트롤을 찾아 선택

        """
        return self.hwp.HAction.Run("SelectCtrlFront")

    def SelectCtrlReverse(self):
        """
        역방향으로 가장 가까운 컨트롤을 찾아 선택

        """
        return self.hwp.HAction.Run("SelectCtrlReverse")

    def SendBrowserText(self):
        """
        브라우저로 보내기

        브라우저가 실행되고, 현재 문서가 브라우저에 나타난다.
        """
        return self.hwp.HAction.Run("SendBrowserText")

    def SetWorkSpaceView(self):
        """작업창 보기 설정"""
        return self.hwp.HAction.Run("SetWorkSpaceView")

    def ShapeObjAlignBottom(self):
        """
        아래로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignBottom")

    def ShapeObjAlignCenter(self):
        """
        가운데로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignCenter")

    def ShapeObjAlignHeight(self):
        """
        높이 맞춤

        """
        return self.hwp.HAction.Run("ShapeObjAlignHeight")

    def ShapeObjAlignHorzSpacing(self):
        """
        왼쪽/오른쪽 일정한 비율로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignHorzSpacing")

    def ShapeObjAlignLeft(self):
        """
        왼쪽으로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignLeft")

    def ShapeObjAlignMiddle(self):
        """
        중간 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignMiddle")

    def ShapeObjAlignRight(self):
        """
        오른쪽으로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignRight")

    def ShapeObjAlignSize(self):
        """
        폭/높이 맞춤

        """
        return self.hwp.HAction.Run("ShapeObjAlignSize")

    def ShapeObjAlignTop(self):
        """
        위로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignTop")

    def ShapeObjAlignVertSpacing(self):
        """
        위/아래 일정한 비율로 정렬

        """
        return self.hwp.HAction.Run("ShapeObjAlignVertSpacing")

    def ShapeObjAlignWidth(self):
        """
        폭 맞춤

        """
        return self.hwp.HAction.Run("ShapeObjAlignWidth")

    def ShapeObjAttachCaption(self):
        """
        역방향으로 가장 가까운 컨트롤에 캡션 추가

        캐럿 위쪽의 표, 이미지 등에 캡션을 생성한다.
        실행 후에는 캡션 안에 캐럿이 있으므로
        직접 insert_text 등으로 입력하고 CloseEx로 빠져나오면 된다.

        위쪽 컨트롤에 캡션이 있을 때에는 캡션 시작부분으로 진입한다.
        """
        return self.hwp.HAction.Run("ShapeObjAttachCaption")

    def ShapeObjAttachTextBox(self):
        """
        사각형을 글상자로 변경

        사각형 선택 상태에서 실행하면 글을 삽입할 수 있게 된다.
        """
        return self.hwp.HAction.Run("ShapeObjAttachTextBox")

    def ShapeObjBringForward(self):
        """
        선택개체를 앞으로

        """
        return self.hwp.HAction.Run("ShapeObjBringForward")

    def ShapeObjBringInFrontOfText(self):
        """
        선택개체를 글 앞으로

        """
        return self.hwp.HAction.Run("ShapeObjBringInFrontOfText")

    def ShapeObjBringToFront(self):
        """
        선택개체를 맨 앞으로

        """
        return self.hwp.HAction.Run("ShapeObjBringToFront")

    def ShapeObjCtrlSendBehindText(self):
        """
        선택개체를 글 뒤로

        """
        return self.hwp.HAction.Run("ShapeObjCtrlSendBehindText")

    def ShapeObjDetachCaption(self):
        """
        캡션 없애기

        """
        return self.hwp.HAction.Run("ShapeObjDetachCaption")

    def ShapeObjDetachTextBox(self):
        """
        글상자를 사각형으로 변경

        (주의)글상자 속의 텍스트는 사라짐
        """
        return self.hwp.HAction.Run("ShapeObjDetachTextBox")

    def ShapeObjFillProperty(self):
        """
        고치기 대화상자중 fill tab

        """
        return self.hwp.HAction.Run("ShapeObjFillProperty")

    def ShapeObjGroup(self):
        """
        틀 묶기

        """
        return self.hwp.HAction.Run("ShapeObjGroup")

    def ShapeObjGuideLine(self) -> bool:
        """
        개체이동 안내선 설정창 열기
        """
        return self.hwp.HAction.Run("ShapeObjGuideLine")

    def ShapeObjHorzFlip(self):
        """
        그리기 개체 좌우 뒤집기

        """
        return self.hwp.HAction.Run("ShapeObjHorzFlip")

    def ShapeObjHorzFlipOrgState(self):
        """
        그리기 개체 좌우 뒤집기를 원상태로 되돌리기

        """
        return self.hwp.HAction.Run("ShapeObjHorzFlipOrgState")

    def ShapeObjInsertCaptionNum(self):
        """
        캡션 번호 넣기

        """
        return self.hwp.HAction.Run("ShapeObjInsertCaptionNum")

    def ShapeObjLineProperty(self):
        """
        고치기 대화상자중 line tab

        """
        return self.hwp.HAction.Run("ShapeObjLineProperty")

    def ShapeObjLineStyleOther(self):
        """
        개체속성 - 선 탭 열기(다른 선 종류)
        """
        return self.hwp.HAction.Run("ShapeObjLineStyleOhter")

    def ShapeObjLineWidthOther(self):
        """
        개체속성 - 선 탭 열기(다른 선 굵기)
        """
        return self.hwp.HAction.Run("ShapeObjLineWidthOhter")

    def ShapeObjLock(self):
        """
        개체 Lock

        """
        return self.hwp.HAction.Run("ShapeObjLock")

    def ShapeObjMoveDown(self):
        """
        키로 움직이기(아래)

        """
        return self.hwp.HAction.Run("ShapeObjMoveDown")

    def ShapeObjMoveLeft(self):
        """
        키로 움직이기(왼쪽)

        """
        return self.hwp.HAction.Run("ShapeObjMoveLeft")

    def ShapeObjMoveRight(self):
        """
        키로 움직이기(오른쪽)

        """
        return self.hwp.HAction.Run("ShapeObjMoveRight")

    def ShapeObjMoveUp(self):
        """
        키로 움직이기(위)

        """
        return self.hwp.HAction.Run("ShapeObjMoveUp")

    def ShapeObjNextObject(self):
        """
        이후 개체로 이동(tab키)

        """
        return self.hwp.HAction.Run("ShapeObjNextObject")

    def ShapeObjNorm(self):
        """
        기본 도형 설정

        """
        return self.hwp.HAction.Run("ShapeObjNorm")

    def ShapeObjPrevObject(self):
        """
        이전 개체로 이동(shift + tab키)

        """
        return self.hwp.HAction.Run("ShapeObjPrevObject")

    def ShapeObjResizeDown(self):
        """
        키로 크기 조절(shift + 아래)

        """
        return self.hwp.HAction.Run("ShapeObjResizeDown")

    def ShapeObjResizeLeft(self):
        """
        키로 크기 조절(shift + 왼쪽)

        """
        return self.hwp.HAction.Run("ShapeObjResizeLeft")

    def ShapeObjResizeRight(self):
        """
        키로 크기 조절(shift + 오른쪽)

        """
        return self.hwp.HAction.Run("ShapeObjResizeRight")

    def ShapeObjResizeUp(self):
        """
        키로 크기 조절(shift + 위)

        """
        return self.hwp.HAction.Run("ShapeObjResizeUp")

    def ShapeObjRightAngleRotater(self):
        """
        90도 시계방향 회전

        """
        return self.hwp.HAction.Run("ShapeObjRightAngleRotater")

    def ShapeObjRightAngleRotaterAnticlockwise(self):
        """
        90도 반시계 방향 회전
        """
        return self.hwp.HAction.Run("ShapeObjRightAngleRotaterAnticlockwise")

    def ShapeObjRotater(self):
        """
        자유각 회전(회전중심 고정)

        """
        return self.hwp.HAction.Run("ShapeObjRotater")

    def ShapeObjSaveAsPicture(self):
        """
        개체를 그림으로 저장하기

        """
        return self.hwp.HAction.Run("ShapeObjSaveAsPicture")

    def ShapeObjSelect(self):
        """
        틀 선택 도구

        """
        return self.hwp.HAction.Run("ShapeObjSelect")

    def ShapeObjSendBack(self):
        """
        개체를 뒤로 보내기
        """
        return self.hwp.HAction.Run("ShapeObjSendBack")

    def ShapeObjSendToBack(self):
        """
        맨 뒤로
        """
        return self.hwp.HAction.Run("ShapeObjSendToBack")

    def ShapeObjShowGuideLine(self) -> bool:
        """
        그리기 개체 안내선 보기 토글
        """
        return self.hwp.HAction.Run("ShapeObjShowGuideLine")


    def ShapeObjShowGuideLineBase(self) -> bool:
        """
        그리기 안내선(한/글2024부터 지원)
        """
        try:
            return self.hwp.HAction.Run("ShapeObjShowGuideLineBase")
        except:
            print("이 기능은 한/글2024부터 지원합니다.")
            return False

    def ShapeObjTableSelCell(self):
        """
        표의 첫 번째 셀 선택

        테이블 선택상태에서 실행해야 한다.
        비슷한 명령어로 ``hwp.ShapeObjTextBoxEdit()``가 있다.
        ``hwp.ShapeObjTableSelCell()``이 셀블록상태인데 반해
        ``hwp.ShapeObjTextBoxEdit()``는 편집상태이다.
        """
        return self.hwp.HAction.Run("ShapeObjTableSelCell")

    def ShapeObjTextBoxEdit(self):
        """
        표나 글상자 선택상태에서 편집모드로 들어가기

        표를 선택하고 있는 경우 A1 셀 안으로 이동한다.

        """
        return self.hwp.HAction.Run("ShapeObjTextBoxEdit")

    def ShapeObjToggleTextBox(self) -> bool:
        """
        도형을 글 상자로 만들기 토글
        """
        return self.hwp.HAction.Run("ShapeObjToggleTextBox")

    def ShapeObjUngroup(self):
        """
        틀 풀기(그룹해제)

        """
        return self.hwp.HAction.Run("ShapeObjUngroup")

    def ShapeObjUnlockAll(self):
        """
        개체 잠금해제(Unlock All)

        """
        return self.hwp.HAction.Run("ShapeObjUnlockAll")

    def ShapeObjVertFlip(self):
        """
        그리기 개체 상하 뒤집기

        """
        return self.hwp.HAction.Run("ShapeObjVertFlip")

    def ShapeObjVertFlipOrgState(self):
        """
        그리기 개체 상하 뒤집기 원상태로 되돌리기

        """
        return self.hwp.HAction.Run("ShapeObjVertFlipOrgState")

    def ShapeObjWrapSquare(self):
        """
        직사각형

        """
        return self.hwp.HAction.Run("ShapeObjWrapSquare")

    def ShapeObjWrapTopAndBottom(self):
        """
        자리 차지

        """
        return self.hwp.HAction.Run("ShapeObjWrapTopAndBottom")

    def ShowAttributeTab(self):
        """속성 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowAttributeTab")

    def ShowBottomWorkspace(self):
        """아래쪽 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowBottomWorkspace")

    def ShowFloatTabFrame(self):
        """플로팅 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowFloatTabFrame")

    def ShowLeftWorkspace(self):
        """왼쪽 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowLeftWorkspace")

    def ShowRightWorkspace(self):
        """오른쪽 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowRightWorkspace")

    def ShowScriptTab(self):
        """스크립트 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowScriptTab")

    def ShowTopWorkspace(self):
        """위쪽 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("ShowTopWorkspace")

    def SoftKeyboard(self):
        """
        소프트키보드 보기 토글

        """
        return self.hwp.HAction.Run("Soft Keyboard")

    def SpellingCheck(self):
        """
        맞춤법

        """
        return self.hwp.HAction.Run("SpellingCheck")

    def SplitAll(self):
        """창 가로 세로 나누기"""
        return self.hwp.HAction.Run("SplitAll")

    def SplitHorz(self):
        """창 가로로 나누기"""
        return self.hwp.HAction.Run("SplitHorz")

    def SplitMainActive(self):
        """메모창 활성화"""
        return self.hwp.HAction.Run("SplitMainActive")

    def SplitMemo(self):
        """메모창 보이기/감추기"""
        return self.hwp.HAction.Run("SplitMemo")

    def SplitMemoClose(self):
        """
        메모창 닫기

        """
        return self.hwp.HAction.Run("SplitMemoClose")

    def SplitMemoOpen(self):
        """
        메모창 열기

        """
        return self.hwp.HAction.Run("SplitMemoOpen")

    def SplitVert(self):
        """창 세로로 나누기"""
        return self.hwp.HAction.Run("SplitVert")

    def StyleClearCharStyle(self):
        """
        글자 스타일 해제

        """
        return self.hwp.HAction.Run("StyleClearCharStyle")

    def StyleCombo(self):
        """글자 스타일"""
        return self.hwp.HAction.Run("StyleCombo")

    def StyleShortcut1(self):
        """
        스타일 단축키1

        """
        return self.hwp.HAction.Run("StyleShortcut1")

    def StyleShortcut2(self):
        """
        스타일 단축키2

        """
        return self.hwp.HAction.Run("StyleShortcut2")

    def StyleShortcut3(self):
        """
        스타일 단축키3

        """
        return self.hwp.HAction.Run("StyleShortcut3")

    def StyleShortcut4(self):
        """
        스타일 단축키4

        """
        return self.hwp.HAction.Run("StyleShortcut4")

    def StyleShortcut5(self):
        """
        스타일 단축키5

        """
        return self.hwp.HAction.Run("StyleShortcut5")

    def StyleShortcut6(self):
        """
        스타일 단축키6

        """
        return self.hwp.HAction.Run("StyleShortcut6")

    def StyleShortcut7(self):
        """
        스타일 단축키7

        """
        return self.hwp.HAction.Run("StyleShortcut7")

    def StyleShortcut8(self):
        """
        스타일 단축키8

        """
        return self.hwp.HAction.Run("StyleShortcut8")

    def StyleShortcut9(self):
        """
        스타일 단축키9

        """
        return self.hwp.HAction.Run("StyleShortcut9")

    def StyleShortcut10(self):
        """
        스타일 단축키10

        """
        return self.hwp.HAction.Run("StyleShortcut10")

    def TabClose(self):
        """
        현재 탭 닫기

        """
        return self.hwp.HAction.Run("TabClose")

    def TableAppendRow(self):
        """
        표 안에서, 현재 행 아래에 새 행 추가

        """
        return self.hwp.HAction.Run("TableAppendRow")

    def TableSubtractRow(self):
        """
        표 안에서, 현재 행 삭제

        """
        return self.hwp.HAction.Run("TableSubtractRow")

    def TableAutoFill(self):
        """표 자동 채우기"""
        return self.hwp.HAction.Run("TableAutoFill")

    def TableAutoFillDlg(self):
        """자동 채우기 창 열기"""
        return self.hwp.HAction.Run("TableAutoFillDlg")

    def TableCellAlignCenterBottom(self):
        """셀 가운데 아래 정렬"""
        return self.hwp.HAction.Run("TableCellAlignCenterBottom")

    def TableCellAlignCenterCenter(self):
        """셀 가운데 정렬"""
        return self.hwp.HAction.Run("TableCellAlignCenterCenter")

    def TableCellAlignCenterTop(self):
        """셀 가운데 위 정렬"""
        return self.hwp.HAction.Run("TableCellAlignCenterTop")

    def TableCellAlignLeftBottom(self):
        """셀 왼쪽 아래 정렬"""
        return self.hwp.HAction.Run("TableCellAlignLeftBottom")

    def TableCellAlignLeftCenter(self):
        """셀 왼쪽 가운데 정렬"""
        return self.hwp.HAction.Run("TableCellAlignLeftCenter")

    def TableCellAlignLeftTop(self):
        """셀 왼쪽 위 정렬"""
        return self.hwp.HAction.Run("TableCellAlignLeftTop")

    def TableCellAlignRightBottom(self):
        """셀 오른쪽 아래 정렬"""
        return self.hwp.HAction.Run("TableCellAlignRightBottom")

    def TableCellAlignRightCenter(self):
        """셀 오른쪽 가운데 정렬"""
        return self.hwp.HAction.Run("TableCellAlignRightCenter")

    def TableCellAlignRightTop(self):
        """셀 가운데 오른쪽 정렬"""
        return self.hwp.HAction.Run("TableCellAlignRightTop")

    def TableCellBlock(self) -> bool:
        """
        셀 블록 상태로 전환

        Returns:
            성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_into_nth_table(0)  # 문서 첫 번째 표로 이동
            >>> hwp.TableCellBlock()  # 셀블록
            >>> hwp.TableCellBlockExtend()  # 셀블록 확장
            >>> hwp.TableCellBlockExtend()  # 전체 셀 선택
            >>> rgb = (50, 150, 250)
            >>> hwp.cell_fill(rgb)

        """
        return self.hwp.HAction.Run("TableCellBlock")

    def TableCellBlockCol(self):
        """
        표 안에서 현재 열(Column) 전체 선택

        """
        return self.hwp.HAction.Run("TableCellBlockCol")

    def TableCellBlockExtend(self):
        """
        셀 블록 연장(F5 + F5)

        실행 전 hwp.TableCellBlock()을 실행해야 한다.

        """
        return self.hwp.HAction.Run("TableCellBlockExtend")

    def TableCellBlockExtendAbs(self):
        """
        셀 블록 연장(SHIFT + F5)

        """
        return self.hwp.HAction.Run("TableCellBlockExtendAbs")

    def TableCellBlockRow(self):
        """
        현재 셀이 포함된 행 전체 선택

        """
        return self.hwp.HAction.Run("TableCellBlockRow")

    def TableCellBorderAll(self):
        """
        모든 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderAll")

    def TableCellBorderBottom(self):
        """
        가장 아래 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderBottom")

    def TableCellBorderDiagonalDown(self):
        """
        대각선(⍂) 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderDiagonalDown")

    def TableCellBorderDiagonalUp(self):
        """
        대각선(⍁) 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderDiagonalUp")

    def TableCellBorderInside(self):
        """
        모든 안쪽 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderInside")

    def TableCellBorderInsideHorz(self):
        """
        모든 안쪽 가로 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderInsideHorz")

    def TableCellBorderInsideVert(self):
        """
        모든 안쪽 세로 셀 테두리 toggle(있음/없음).

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderInsideVert")

    def TableCellBorderLeft(self):
        """
        가장 왼쪽의 셀 테두리 toggle(있음/없음)

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderLeft")

    def TableCellBorderNo(self):
        """
        모든 셀 테두리 지움.

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderNo")

    def TableCellBorderOutside(self):
        """
        바깥 셀 테두리 toggle(있음/없음)

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderOutside")

    def TableCellBorderRight(self):
        """
        가장 오른쪽의 셀 테두리 toggle(있음/없음)

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderRight")

    def TableCellBorderTop(self):
        """
        가장 위의 셀 테두리 toggle(있음/없음)

        셀 블록 상태일 경우에만 동작한다.

        """
        return self.hwp.HAction.Run("TableCellBorderTop")

    def TableColBegin(self):
        """
        셀 이동: 현재 행의 가장 왼쪽 셀로 이동

        """
        return self.hwp.HAction.Run("TableColBegin")

    def TableColEnd(self):
        """
        셀 이동: 현재 행의 가장 오른쪽 셀로 이동

        """
        return self.hwp.HAction.Run("TableColEnd")

    def TableColPageDown(self):
        """
        셀 이동: 현재 열의 맨 아래로 이동

        """
        return self.hwp.HAction.Run("TableColPageDown")

    def TableColPageUp(self):
        """
        셀 이동: 현재 열의 맨 위로 이동

        """
        return self.hwp.HAction.Run("TableColPageUp")

    def TableDeleteCell(self, remain_cell: bool = False) -> bool:
        """
        셀 삭제

        셀블록 상태에서 실행해야 한다.
        행 또는 열 전체 선택상태에서는 해당 행/열을 제거한다.

        Args:
             remain_cell: 지울 때 빈 셀은 남겨둘 것인가? (기본값은 남겨두지 않음=False)

        Returns:
            성공시 True, 실패시 False를 리턴.

        """
        if remain_cell:
            self.set_message_box_mode(0x1000)
        else:
            self.set_message_box_mode(0x2000)
        try:
            return self.hwp.HAction.Run("TableDeleteCell")
        finally:
            self.set_message_box_mode(0xF000)

    def TableDeleteComma(self) -> bool:
        """
        표 안에 천단위구분콤마 빼기
        """
        return self.hwp.HAction.Run("TableDeleteComma")

    def TableInsertComma(self) -> bool:
        """
        표 안에 천단위구분콤마 넣기
        """
        return self.hwp.HAction.Run("TableInsertComma")

    def TableDistributeCellHeight(self):
        """
        표 안에서 셀 높이를 같게
        """
        return self.hwp.HAction.Run("TableDistributeCellHeight")

    def TableDistributeCellWidth(self):
        """
        표 안에서 셀 너비를 같게
        """
        return self.hwp.HAction.Run("TableDistributeCellWidth")

    def TableDrawPen(self):
        """
        표 그리기
        """
        return self.hwp.HAction.Run("TableDrawPen")

    def TableDrawPenStyle(self):
        """표 그리기 선 모양"""
        return self.hwp.HAction.Run("TableDrawPenStyle")

    def TableDrawPenWidth(self):
        """표 그리기 선 굵기"""
        return self.hwp.HAction.Run("TableDrawPenWidth")

    def TableEraser(self):
        """
        표 지우개
        """
        return self.hwp.HAction.Run("TableEraser")

    def TableFormulaAvgAuto(self):
        """
        표 안에 블록 평균 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaAvgAuto")

    def TableFormulaAvgHor(self):
        """
        표 안에 가로 평균 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaAvgHor")

    def TableFormulaAvgVer(self):
        """
        표 안에 세로 평균 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaAvgVer")

    def TableFormulaProAuto(self):
        """
        표 안에 블록 자동곱 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaProAuto")

    def TableFormulaProHor(self):
        """
        표 안에 가로 곱 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaProHor")

    def TableFormulaProVer(self):
        """
        표 안에 세로 곱 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaProVer")

    def TableFormulaSumAuto(self):
        """
        표 안에 블록 합계 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaSumAuto")

    def TableFormulaSumHor(self):
        """
        표 안에 가로 합계 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaSumHor")

    def TableFormulaSumVer(self):
        """
        표 안에 세로 합계 수식 삽입
        """
        return self.hwp.HAction.Run("TableFormulaSumVer")

    def TableLeftCell(self):
        """
        셀 이동: 셀 왼쪽
        """
        return self.hwp.HAction.Run("TableLeftCell")

    def TableLowerCell(self):
        """
        셀 이동: 셀 아래
        """
        return self.hwp.HAction.Run("TableLowerCell")

    def TableMergeCell(self):
        """
        셀 합치기

        셀 병합(m)
        """
        return self.hwp.HAction.Run("TableMergeCell")

    def TableMergeTable(self):
        """
        표와 표 붙이기
        """
        self.Cancel()
        result = self.hwp.HAction.Run("TableMergeTable")
        if result:
            return result
        else:
            self.set_message_box_mode(0x1)
            sleep(0.1)
            self.set_message_box_mode(0xF)
            return result

    def TableResizeCellDown(self):
        """
        셀 크기 변경: 셀 아래
        """
        return self.hwp.HAction.Run("TableResizeCellDown")

    def TableResizeCellLeft(self):
        """
        셀 크기 변경: 셀 왼쪽
        """
        return self.hwp.HAction.Run("TableResizeCellLeft")

    def TableResizeCellRight(self):
        """
        셀 크기 변경: 셀 오른쪽
        """
        return self.hwp.HAction.Run("TableResizeCellRight")

    def TableResizeCellUp(self):
        """
        셀 크기 변경: 셀 위
        """
        return self.hwp.HAction.Run("TableResizeCellUp")

    def TableResizeDown(self):
        """
        셀 크기 변경: 셀 아래
        """
        return self.hwp.HAction.Run("TableResizeDown")

    def TableResizeExDown(self):
        """
        셀 크기 변경: 셀 아래.

        TebleResizeDown과 다른 점은
        셀 블록 상태가 아니어도 동작한다는 점이다.

        """
        return self.hwp.HAction.Run("TableResizeExDown")

    def TableResizeExLeft(self):
        """
        셀 크기 변경: 셀 왼쪽.

        TableResizeLeft와 다른 점은 셀 블록 상태가 아니어도 동작한다는 점이다.

        """
        return self.hwp.HAction.Run("TableResizeExLeft")

    def TableResizeExRight(self):
        """
        셀 크기 변경: 셀 오른쪽.

        TableResizeRight와 다른 점은 셀 블록 상태가 아니어도 동작한다는 점이다.
        """
        return self.hwp.HAction.Run("TableResizeExRight")

    def TableResizeExUp(self):
        """
        셀 크기 변경: 셀 위쪽.

        TableResizeUp과 다른 점은 셀 블록 상태가 아니어도 동작한다는 점이다.
        """
        return self.hwp.HAction.Run("TableResizeExUp")

    def TableResizeLeft(self):
        """
        셀 크기 변경: 왼쪽
        """
        return self.hwp.HAction.Run("TableResizeLeft")

    def TableResizeLineDown(self):
        """
        셀 크기 변경: 선아래
        """
        return self.hwp.HAction.Run("TableResizeLineDown")

    def TableResizeLineLeft(self):
        """
        셀 크기 변경: 선 왼쪽
        """
        return self.hwp.HAction.Run("TableResizeLineLeft")

    def TableResizeLineRight(self):
        """
        셀 크기 변경: 선 오른쪽
        """
        return self.hwp.HAction.Run("TableResizeLineRight")

    def TableResizeLineUp(self):
        """
        셀 크기 변경: 선 위
        """
        return self.hwp.HAction.Run("TableResizeLineUp")

    def TableResizeRight(self):
        """
        셀 크기 변경: 우측으로
        """
        return self.hwp.HAction.Run("TableResizeRight")

    def TableResizeUp(self):
        """
        셀 크기 변경: 위로
        """
        return self.hwp.HAction.Run("TableResizeUp")

    def TableRightCell(self):
        """
        셀 이동: 셀 오른쪽
        """
        return self.hwp.HAction.Run("TableRightCell")

    def TableRightCellAppend(self):
        """
        셀 이동: 셀 오른쪽에 이어서

        우측 셀로 이동하다 끝에 도달하면 다음 행의 첫 번째 셀로 이동.
        그리고 다음 행이 없는 경우에는 새 행을 아래 추가하고 첫 번째 셀로 이동.

        """
        return self.hwp.HAction.Run("TableRightCellAppend")

    def TableSplitCell(
        self, Rows: int = 2, Cols: int = 0, DistributeHeight: int = 0, Merge: int = 0
    ) -> bool:
        """
        셀 나누기.

        Args:
            Rows: 나눌 행 수(기본값:2)
            Cols: 나눌 열 수(기본값:0)
            DistributeHeight: 줄 높이를 같게 나누기(0 or 1)
            Merge: 셀을 합친 후 나누기(0 or 1)
        """
        pset = self.hwp.HParameterSet.HTableSplitCell
        pset.Rows = Rows
        pset.Cols = Cols
        pset.DistributeHeight = DistributeHeight
        pset.Merge = Merge
        return self.hwp.HAction.Execute("TableSplitCell", pset.HSet)

    def TableSplitTable(self):
        """
        표 나누기

        현재 캐럿이 있는 행을 포함한 아랫쪽 부분을 잘라서
        별도의 표로 만든다.

        셀 블록 상태에서는 작동하지 않는다. 일반 편집상태여야 한다.
        """
        if self.get_cell_addr("tuple")[0] == 0:
            return False
        return self.hwp.HAction.Run("TableSplitTable")

    def TableUpperCell(self):
        """
        셀 이동: 셀 위로
        """
        return self.hwp.HAction.Run("TableUpperCell")

    def TableVAlignBottom(self):
        """
        셀의 텍스트를 아래로 세로정렬

        """
        return self.hwp.HAction.Run("TableVAlignBottom")

    def TableVAlignCenter(self):
        """
        셀의 텍스트를 가운데로 세로정렬

        """
        return self.hwp.HAction.Run("TableVAlignCenter")

    def TableVAlignTop(self):
        """
        셀의 텍스트를 위로 세로정렬

        """
        return self.hwp.HAction.Run("TableVAlignTop")

    def ToggleOverwrite(self):
        """
        Toggle Overwrite

        """
        return self.hwp.HAction.Run("ToggleOverwrite")

    def TopTabFrameClose(self):
        """위쪽 작업창 감추기"""
        return self.hwp.HAction.Run("TopTabFrameClose")

    def TrackChangeApply(self):
        """변경추적: 변경내용 적용"""
        return self.hwp.HAction.Run("TrackChangeApply")

    def TrackChangeApplyAll(self):
        """변경추적: 문서에서 변경내용 모두 적용"""
        return self.hwp.HAction.Run("TrackChangeApplyAll")

    def TrackChangeApplyNext(self):
        """변경추적: 적용 후 다음으로 이동"""
        return self.hwp.HAction.Run("TrackChangeApplyNext")

    def TrackChangeApplyPrev(self):
        """변경추적: 적용 후 이전으로 이동"""
        return self.hwp.HAction.Run("TrackChangeApplyPrev")

    def TrackChangeApplyViewAll(self):
        """변경추적: 표시된 변경내용 모두 적용"""
        return self.hwp.HAction.Run("TrackChangeApplyViewAll")

    def TrackChangeAuthor(self):
        """변경추적: 사용자 이름 변경"""
        return self.hwp.HAction.Run("TrackChangeAuthor")

    def TrackChangeCancel(self):
        """변경추적: 변경내용 취소"""
        return self.hwp.HAction.Run("TrackChangeCancel")

    def TrackChangeCancelAll(self):
        """변경추적: 문서에서 변경내용 모두 취소"""
        return self.hwp.HAction.Run("TrackChangeCancelAll")

    def TrackChangeCancelNext(self):
        """변경추적: 취소 후 다음으로 이동"""
        return self.hwp.HAction.Run("TrackChangeCancelNext")

    def TrackChangeCancelPrev(self):
        """변경추적: 취소 후 이전으로 이동"""
        return self.hwp.HAction.Run("TrackChangeCancelPrev")

    def TrackChangeCancelViewAll(self):
        """변경추적: 표시된 변경내용 모두 취소"""
        return self.hwp.HAction.Run("TrackZChangeCancelViewAll")

    def TrackChangeNext(self):
        """변경추적: 다음 변경내용"""
        return self.hwp.HAction.Run("TrackChangeNext")

    def TrackChangePrev(self):
        """변경추적: 이전 변경내용"""
        return self.hwp.HAction.Run("TrackChangePrev")

    def ViewOptionTrackChange(self):
        """변경추적 보기"""
        return self.hwp.HAction.Run("ViewOptionTrackChange")

    def ViewOptionTrackChangeFinal(self):
        """변경추적 보기: 최종본 보기"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeFinal")

    def ViewOptionTrackChangeFinalMemo(self):
        """변경추적 보기: 메모 및 변경 내용 최종본"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeFinalMemo")

    def ViewOptionTrackChangeInline(self):
        """변경추적 보기: 안내문에 표시"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeInline")

    def ViewOptionTrackChangeInsertDelete(self):
        """변경추적 보기: 삽입 및 삭제"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeInsertDelete")

    def ViewOptionTrackChangeOriginal(self):
        """변경추적 보기: 원본 보기"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeOriginal")

    def ViewOptionTrackChangeOriginalMemo(self):
        """변경추적 보기: 메모 및 변경 내용 원본"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeOriginalMemo")

    def ViewOptionTrackChangeShape(self):
        """변경추적 보기: 서식"""
        return self.hwp.HAction.Run("ViewOptionTrackChangeShape")

    def ViewOptionTrackChnageInfo(self):
        """변경추적 보기: 변경 내용 보기"""
        return self.hwp.HAction.Run("ViewOptionTrackChnageInfo")

    def Undo(self):
        """
        실행 취소(Ctrl-Z)
        """
        return self.hwp.HAction.Run("Undo")

    def UnlinkTextBox(self):
        """
        글상자 연결 끊기
        """
        return self.hwp.HAction.Run("UnlinkTextBox")

    def VersionDeleteAll(self):
        """
        모든 버전정보 지우기
        """
        return self.hwp.HAction.Run("VersionDeleteAll")

    def VertScrollbar(self):
        """세로축 스크롤바 보이기/감추기"""
        return self.hwp.HAction.Run("VertScrollbar")

    def ViewIdiom(self):
        """
        상용구 보기

        """
        return self.hwp.HAction.Run("ViewIdiom")

    def ViewOptionCtrlMark(self):
        """
        조판 부호 보기 토글

        """
        return self.hwp.HAction.Run("ViewOptionCtrlMark")

    def ViewOptionGuideLine(self):
        """
        안내선 보기 토글
        """
        return self.hwp.HAction.Run("ViewOptionGuideLine")

    def ViewOptionMemo(self):
        """
        메모 보이기/숨기기 토글

        ([보기-메모-메모 보이기/숨기기]메뉴와 동일)

        """
        return self.hwp.HAction.Run("ViewOptionMemo")

    def ViewOptionMemoGuideline(self) -> bool:
        """
        메모 안내선 표시 토글

        ([보기-메모-메모 안내선 표시]메뉴와 동일)

        """
        return self.hwp.HAction.Run("ViewOptionMemoGuideline")

    def ViewOptionPaper(self) -> bool:
        """
        쪽 윤곽 보기/숨기기 토글

        """
        return self.hwp.HAction.Run("ViewOptionPaper")

    def ViewOptionParaMark(self) -> bool:
        """
        문단 부호 보기/숨기기 토글

        """
        return self.hwp.HAction.Run("ViewOptionParaMark")

    def ViewOptionPicture(self) -> bool:
        """
        그림 보이기/숨기기 토글

        ([보기-그림]메뉴와 동일)

        """
        return self.hwp.HAction.Run("ViewOptionPicture")

    def ViewOptionRevision(self) -> bool:
        """
        교정부호 보이기/숨기기 토글

        ([보기-교정부호]메뉴와 동일)

        """
        return self.hwp.HAction.Run("ViewOptionRevision")

    def ViewTabButton(self) -> bool:
        """
        문서탭 보이기/감추기 토글
        """
        return self.hwp.HAction.Run("ViewTabButton")

    def ViewZoomFitPage(self) -> bool:
        """화면 확대: 페이지에 맞춤"""
        return self.hwp.HAction.Run("ViewZoomFitPage")

    def ViewZoomNormal(self) -> bool:
        """화면 확대: 폭에 맞춤"""
        return self.hwp.HAction.Run("ViewZoomFitNormal")

    def ViewZoomFitWidth(self) -> bool:
        """화면 확대: 폭에 맞춤"""
        return self.hwp.HAction.Run("ViewZoomFitWidth")

    def ViewZoomRibon(self) -> bool:
        """화면 확대"""
        return self.hwp.HAction.Run("ViewZoomRibon")

    def VoiceCommandConfig(self) -> bool:
        """
        음성 명령 설정
        """
        return self.hwp.HAction.Run("VoiceCommand Config")

    def VoiceCommandResume(self) -> bool:
        """
        음성 명령 레코딩 시작
        """
        return self.hwp.HAction.Run("VoiceCommand Resume")

    def VoiceCommandStop(self) -> bool:
        """
        음성 명령 레코딩 중지
        """
        return self.hwp.HAction.Run("VoiceCommand Stop")
