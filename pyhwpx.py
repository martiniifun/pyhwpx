import os
import re
from typing import Literal, Union
from time import sleep

import numpy as np
import pandas as pd
import pythoncom
import win32com.client as win32
import pyperclip as cb
import shutil

# temp 폴더 삭제
try:
    shutil.rmtree(os.path.join(os.environ["USERPROFILE"], "AppData/Local/Temp/gen_py"))
except FileNotFoundError as e:
    pass

# Type Library 파일 재생성
win32.gencache.EnsureModule('{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}', 0, 1, 0)


# 아래아한글 오토메이션 클래스 정의
class Hwp:
    """
    아래아한글 인스턴스를 실행한다.

    :param new:
        new=True인 경우, 기존에 열려 있는 한/글 인스턴스와 무관한 새 인스턴스를 생성한다.
        new=False(기본값)인 경우, 기존에 열려 있는 한/글 인스턴스를 조작하게 된다.
    :param visible:
        한/글 인스턴스를 백그라운드에서 실행할지, 화면에 나타낼지 선택한다.
        기본값은 True로, 화면에 나타나게 된다.
        visible=False일 경우 백그라운드에서 작업할 수 있다.
    :param register_module:
        기존의 hwp__.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") 메서드를 실행한다.
        레지스트리 키를 직접 추가(수정)한다.
    """

    def __repr__(self):
        return "<파이썬+아래아한글 자동화를 돕기 위한 함수모음 및 추상화 인스턴스>"

    def __init__(self, new=False, visible=True, register_module=True):
        self.hwp = 0
        context = pythoncom.CreateBindCtx(0)

        # 현재 실행중인 프로세스를 가져옵니다.
        running_coms = pythoncom.GetRunningObjectTable()
        monikers = running_coms.EnumRunning()

        if not new:
            for moniker in monikers:
                name = moniker.GetDisplayName(context, moniker)
                # moniker의 DisplayName을 통해 한글을 가져옵니다
                # 한글의 경우 HwpObject.버전으로 각 버전별 실행 이름을 설정합니다.
                if name.startswith('!HwpObject.'):
                    # 120은 한글 2022의 경우입니다.
                    # 현재 moniker를 통해 ROT에서 한글의 object를 가져옵니다.
                    obj = running_coms.GetObject(moniker)
                    # 가져온 object를 Dispatch를 통해 사용할수 있는 객체로 변환시킵니다.
                    self.hwp = win32.gencache.EnsureDispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                    # 그이후는 오토메이션 api를 사용할수 있습니다
        if not self.hwp:
            self.hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
        try:
            self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible
        except:
            sleep(0.01)
            self.hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
            self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible

        if register_module:
            self.register_module()

    @property
    def HeadCtrl(self):
        return self.hwp.HeadCtrl

    @property
    def LastCtrl(self):
        return self.hwp.LastCtrl

    @property
    def PageCount(self):
        return self.hwp.PageCount

    def message_box(self, string, flag: int = 0):
        msgbox = self.hwp.XHwpMessageBox  # 메시지박스 생성
        msgbox.string = string
        msgbox.Flag = flag  # [확인] 버튼만 나타나게 설정
        msgbox.DoModal()  # 메시지박스 보이기
        return msgbox.Result

    def insert_file(self, filename, keep_section=0, keep_charshape=0, keep_parashape=0, keep_style=0):
        if not filename.lower().startswith("c:"):
            filename = os.path.join(os.getcwd(), filename)
        pset = self.hwp.HParameterSet.HInsertFile
        self.hwp.HAction.GetDefault("InsertFile", pset.HSet)
        pset.filename = filename
        pset.KeepSection = keep_section
        pset.KeepCharshape = keep_charshape
        pset.KeepParashape = keep_parashape
        pset.KeepStyle = keep_style

        return self.hwp.HAction.Execute("InsertFile", pset.HSet)

    def insert_memo(self, text):
        """
        선택한 단어 범위에 메모고침표를 삽입하는 코드.
        한/글에서 일반 문자열을 삽입하는 코드와 크게 다르지 않다.
        선택모드가 아닌 경우 캐럿이 위치한 단어에 메모고침표를 삽입한다.
        :param text: str
        :return: None
        """
        self.InsertFieldRevisionChagne()  # 이 라인이 메모고침표 삽입하는 코드
        self.insert_text(text)
        self.CloseEx()

    def is_cell(self):
        """
        캐럿이 현재 표 안에 있는지 알려주는 메서드
        :return:
            표 안에 있으면 True, 그렇지 않으면 False를 리턴
        """
        if self.key_indicator()[-1].startswith("("):
            return True
        else:
            return False

    def find(self, src, direction: Literal["Forward", "Backward", "AllDoc"] = "AllDoc"):
        """
        캐럿 뒤의 특정 단어를 찾아가는 메서드.
        해당 단어를 선택한 상태가 되며,
        문서 끝에 도달시 문서 처음부터 탐색을 재개하고,
        원래 캐럿위치까지 왔을 때 False를 리턴하므로
        while문의 조건으로 사용할 수 있다.

        :param src:
            찾을 단어
        :return:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴,
            단어가 더이상 없으면 False를 리턴
        """
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.MatchCase = 1
        pset.SeveralWords = 1
        pset.UseWildCards = 1
        pset.AutoSpell = 1
        pset.Direction = self.find_dir(direction)
        pset.FindString = src
        pset.IgnoreMessage = 1
        pset.HanjaFromHangul = 1
        return self.hwp.HAction.Execute("RepeatFind", pset.HSet)

    def set_field_by_bracket(self):
        """
        필드를 지정하는 일련의 반복작업을 간소화하기 위한 메서드.

        중괄호 두 겹({{}})으로 둘러싸인 구문을 누름틀로 변환해준다.
        만약 본문에 "{{name}}"이라는 문구가 있었다면 해당 단어를 삭제하고
        그 위치에 name이라는 누름틀을 생성한다.
        지시문(direction)과 메모(memo)도 추가가 가능한데,
        "{{name:direction}}" 또는 "{{name:direction:memo}}" 방식으로
        콜론으로 구분하여 지정할 수 있다.
        (가급적 direction을 지정해주도록 하자. 그렇지 않으면 누름틀이 보이지 않는다.)
        셀 안에서 누름틀을 삽입할 수도 있지만,
        편의상 셀필드를 삽입하고 싶은 경우 "[[name]]"으로 지정하면 된다.
        :return:
        """
        while self.find("{{"):
            while True:
                self.hwp.HAction.Run("MoveSelRight")
                if self.get_selected_text().endswith("}}"):
                    field_name = self.get_selected_text()[2:-2]
                    if ":" in field_name:
                        field_name, direction = field_name.split(":", maxsplit=1)
                        if ":" in direction:
                            direction, memo = direction.split(":", maxsplit=1)
                        else:
                            memo = ""
                    else:
                        direction = memo = ""
                    break
                if self.get_selected_text().endswith("\r\n"):
                    raise Exception("필드를 닫는 중괄호가 없습니다.")
            self.hwp.HAction.Run("Delete")
            self.create_field(field_name, direction, memo)

        while self.find("[["):
            while True:
                self.hwp.HAction.Run("MoveSelRight")
                if self.get_selected_text().endswith("]]"):
                    field_name = self.get_selected_text()[2:-2]
                    if ":" in field_name:
                        field_name, direction = field_name.split(":", maxsplit=1)
                        if ":" in direction:
                            direction, memo = direction.split(":", maxsplit=1)
                        else:
                            memo = ""
                    else:
                        direction = memo = ""
                    break
                if self.get_selected_text().endswith("\r\n"):
                    raise Exception("필드를 닫는 중괄호가 없습니다.")
            self.hwp.HAction.Run("Delete")
            if self.is_cell():
                self.set_cur_field_name(field_name, option=1, direction=direction, memo=memo)
            else:
                pass

    def find_replace_all(self, src, dst):
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("AllReplace", pset.HSet)
        pset.Direction = self.hwp.FindDir("AllDoc")
        pset.FindString = src  # "\\r\\n"
        pset.ReplaceString = dst  # "^n"
        pset.ReplaceMode = 1
        pset.IgnoreMessage = 1
        pset.FindType = 1
        self.hwp.HAction.Execute("AllReplace", pset.HSet)

    def clipboard_to_pyfunc(self):
        """
        한/글 프로그램에서 스크립트매크로 녹화 코드를 클립보드에 복사하고
        clipboard_to_pyfunc()을 실행하면, 클립보드의 매크로가 파이썬 함수로 변경된다.
        곧 정규식으로 업데이트 예정(2023. 11. 30)
        """
        text = cb.paste()
        text = text.replace("\t", "    ").replace(";", "")
        text = text.split("{\r\n", maxsplit=1)[1][:-5]
        if "with" in text:
            pset_name = text.split("with (")[1].split(")")[0]
            inner_param = text.split("{\r\n")[1].split("\r\n}")[0].replace("        ", f"    {pset_name}")
            result = f"def script_macro():\r\n    pset = {pset_name}\r\n    " + text.replace("    ", "").split("with")[
                0].replace(pset_name, "pset").replace("\r\n", "\r\n    ") + inner_param.replace(pset_name,
                                                                                                "pset.").replace("    ",
                                                                                                                 "").replace(
                "}\r\n", "").replace("..", ".").replace("\r\n", "\r\n    ")
        else:
            pset_name = text.split(", ")[1].split(".HSet")[0]
            result = f"def script_macro():\r\n    pset = {pset_name}\r\n    " + text.replace("    ", "").replace(
                pset_name, "pset").replace("\r\n", "\r\n    ")
        print(result)
        cb.copy(result)

    def clear_field_text(self):
        for i in self.hwp.GetFieldList(1).split("\x02"):
            self.hwp.PutFieldText(i, "")

    def switch_to(self, num):
        """
        여러 개의 hwp인스턴스가 열려 있는 경우 해당 인스턴스를 활성화한다.
        :param num:
            인스턴스 번호
        :return:
            None
        """
        self.hwp.XHwpDocuments.Item(num).SetActive_XHwpDocument()

    def add_tab(self):
        """
        새 문서를 현재 창의 새 탭에 추가한다.
        백그라운드 상태에서 새 창을 만들 때 윈도우에 나타나는 경우가 있는데,
        add_tab() 함수를 사용하면 백그라운드 작업이 보장된다.
        탭 전환은 switch_to() 메서드로 가능하다.

        새 창을 추가하고 싶은 경우는 add_tab 대신 hwp__.FileNew()나 hwp__.add_doc()을 실행하면 된다.
        :return:
        """
        self.hwp.XHwpDocuments.Add(1)  # 0은 새 창, 1은 새 탭

    def add_doc(self):
        """
        새 문서를 추가한다. 원래 창이 백그라운드로 숨겨져 있어도
        추가된 문서는 보이는 상태가 기본값이다. 숨기려면 set_visible(False)를 실행해야 한다.
        새 탭을 추가하고 싶은 경우는 add_doc 대신 add_tab()을 실행하면 된다.
        :return:
        """
        self.hwp.XHwpDocuments.Add(0)  # 0은 새 창, 1은 새 탭

    def hwp_unit_to_mili(self, hwp_unit):
        """
        HwpUnit 값을 밀리미터로 변환한 값을 리턴한다.
        HwpUnit으로 리턴되었거나, 녹화된 코드의 HwpUnit값을 확인할 때 유용하게 사용할 수 있다.

        :return:
            HwpUnit을 7200으로 나눈 후 25.4를 곱하고 반올림한 값
        """
        return round(hwp_unit / 7200 * 25.4)

    def create_table(self, rows, cols, treat_as_char=1, width_type=0, height_type=0):
        """
        표를 생성하는 메서드.
        기본적으로 rows와 cols만 지정하면 되며,
        용지여백을 제외한 구간에 맞춰 표 너비가 결정된다.
        이는 일반적인 표 생성과 동일한 수치이다.

        아래의 148mm는 종이여백 210mm에서 60mm(좌우 각 30mm)를 뺀 150mm에다가,
        표 바깥여백 각 1mm를 뺀 148mm이다. (TableProperties.Width = 41954)
        각 열의 너비는 5개 기준으로 26mm인데 이는 셀마다 안쪽여백 좌우 각각 1.8mm를 뺀 값으로,
        148 - (1.8 x 10 =) 18mm = 130mm
        그래서 셀 너비의 총 합은 130이 되어야 한다.
        아래의 라인28~32까지 셀너비의 합은 16+36+46+16+16=130
        표를 생성하는 시점에는 표 안팎의 여백을 없애거나 수정할 수 없으므로
        이는 고정된 값으로 간주해야 한다.

        :return:
            표 생성 성공시 True, 실패시 False를 리턴한다.
        """
        pset = self.hwp.HParameterSet.HTableCreation
        self.hwp.HAction.GetDefault("TableCreate", pset.HSet)  # 표 생성 시작
        pset.Rows = rows  # 행 갯수
        pset.Cols = cols  # 열 갯수
        pset.WidthType = width_type  # 너비 지정(0:단에맞춤, 1:문단에맞춤, 2:임의값)
        pset.HeightType = height_type  # 높이 지정(0:자동, 1:임의값)

        sec_def = self.hwp.HParameterSet.HSecDef
        self.hwp.HAction.GetDefault("PageSetup", sec_def.HSet)
        total_width = (sec_def.PageDef.PaperWidth - sec_def.PageDef.LeftMargin
                       - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen
                       - self.mili_to_hwp_unit(2))

        pset.WidthValue = self.hwp.MiliToHwpUnit(total_width)  # 표 너비
        # pset.HeightValue = self.hwp.MiliToHwpUnit(150)  # 표 높이
        pset.CreateItemArray("ColWidth", cols)  # 열 5개 생성
        each_col_width = total_width - self.mili_to_hwp_unit(3.6 * cols)
        for i in range(cols):
            pset.ColWidth.SetItem(i, self.hwp.MiliToHwpUnit(each_col_width))  # 1열
        pset.TableProperties.TreatAsChar = treat_as_char  # 글자처럼 취급
        pset.TableProperties.Width = total_width  # self.hwp.MiliToHwpUnit(148)  # 표 너비
        return self.hwp.HAction.Execute("TableCreate", pset.HSet)  # 위 코드 실행

    def get_selected_text(self):
        """
        한/글 문서 선택 구간의 텍스트를 리턴하는 메서드.
        :return:
            선택한 문자열
        """
        self.hwp.InitScan(Range=0xff)
        total_text = ""
        state = 2
        while state not in [0, 1]:
            state, text = self.hwp.GetText()
            total_text += text
        self.hwp.ReleaseScan()
        return total_text

    def table_to_csv(self, idx=1, filename="result.csv"):
        """
        한/글 문서의 idx번째 표를 현재 폴더에 filename으로 csv포맷으로 저장한다.
        filename을 지정하지 않는 경우 "./result.csv"가 기본값이다.
        :return:
            None
        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> hwp.table_to_csv(1, "table.csv")
        """
        start_pos = self.hwp.GetPos()
        table_num = 0
        ctrl = self.HeadCtrl
        while ctrl.Next:
            if ctrl.UserDesc == "표":
                table_num += 1
            if table_num == idx:
                break
            ctrl = ctrl.Next

        self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
        self.hwp.FindCtrl()
        self.hwp.HAction.Run("ShapeObjTableSelCell")
        data = [self.get_selected_text()]
        col_count = 1
        while self.hwp.HAction.Run("TableRightCell"):
            # a.append(get_text().replace("\r\n", "\n"))
            if re.match("\([A-Z]1\)", self.hwp.KeyIndicator()[-1]):
                col_count += 1
            data.append(self.get_selected_text())

        array = np.array(data).reshape(-1, col_count)
        df = pd.DataFrame(array[1:], columns=array[0])
        df.to_csv(filename, index=False)
        self.hwp.SetPos(*start_pos)
        print(os.path.join(os.getcwd(), filename))
        return None

    def table_to_df(self, idx=1):
        """
        한/글 문서의 idx번째 표를 판다스 데이터프레임으로 리턴하는 메서드.
        :return:
            pd.DataFrame
        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> df = hwp.table_to_df(1)
        """
        start_pos = self.hwp.GetPos()
        table_num = 0
        ctrl = self.HeadCtrl
        while ctrl.Next:
            if ctrl.UserDesc == "표":
                table_num += 1
            if table_num == idx:
                break
            ctrl = ctrl.Next

        self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
        self.hwp.FindCtrl()
        self.hwp.HAction.Run("ShapeObjTableSelCell")
        data = [self.get_selected_text()]
        col_count = 1
        while self.hwp.HAction.Run("TableRightCell"):
            # a.append(get_text().replace("\r\n", "\n"))
            if re.match("\([A-Z]1\)", self.hwp.KeyIndicator()[-1]):
                col_count += 1
            data.append(self.get_selected_text())

        array = np.array(data).reshape(-1, col_count)
        df = pd.DataFrame(array[1:], columns=array[0])
        self.hwp.SetPos(*start_pos)
        return df

    def table_to_bottom(self, offset=0.):
        """
        표 앞에 캐럿을 둔 상태 또는 캐럿이 표 안에 있는 상태에서 위 함수 실행시
        표를 (페이지 기준) 하단으로 위치시킨다.
        :param offset:
            페이지 하단 기준 오프셋(mm)
        :return:
        """
        self.hwp.FindCtrl()
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.VertAlign = self.hwp.VAlign("Bottom")
        pset.VertRelTo = self.hwp.VertRel("Page")
        pset.TreatAsChar = 0
        pset.VertOffset = self.hwp.MiliToHwpUnit(offset)
        pset.HSet.SetItem("ShapeType", 3)
        self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
        self.hwp.Run("Cancel")

    def insert_text(self, text):
        """
        한/글 문서 내 캐럿 위치에 문자열을 삽입하는 메서드.
        :return:
            삽입 성공시 True, 실패시 False를 리턴함.
        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_text("Hello world!\r\n")
        """
        param = self.hwp.HParameterSet.HInsertText
        self.hwp.HAction.GetDefault("InsertText", param.HSet)
        param.Text = text
        return self.hwp.HAction.Execute("InsertText", param.HSet)

    def move_caption(self, location: Literal["Top", "Bottom", "Left", "Right"] = "Bottom"):
        """
        한/글 문서 내 모든 표의 주석 위치를 이동하는 메서드.
        """
        start_pos = self.hwp.GetPos()
        ctrl = self.HeadCtrl
        while ctrl:
            if ctrl.UserDesc == "번호 넣기":
                self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                self.hwp.HAction.Run("ParagraphShapeAlignCenter")
                param = self.hwp.HParameterSet.HShapeObject
                self.hwp.HAction.GetDefault("TablePropertyDialog", param.HSet)
                param.ShapeCaption.Side = self.hwp.SideType(location)
                self.hwp.HAction.Execute("TablePropertyDialog", param.HSet)
            ctrl = ctrl.Next
        self.hwp.SetPos(*start_pos)
        return None

    def is_empty(self) -> bool:
        """
        아무 내용도 들어있지 않은 빈 문서인지 여부를 나타낸다. 읽기전용
        """
        return self.hwp.IsEmpty

    def is_modified(self) -> bool:
        """
        최근 저장 또는 생성 이후 수정이 있는지 여부를 나타낸다. 읽기전용
        """
        return self.hwp.IsModified

    def arc_type(self, arc_type):
        return self.hwp.ArcType(ArcType=arc_type)

    def auto_num_type(self, autonum):
        return self.hwp.AutoNumType(autonum=autonum)

    def border_shape(self, border_type):
        return self.hwp.BorderShape(BorderType=border_type)

    def break_word_latin(self, break_latin_word):
        return self.hwp.BreakWordLatin(BreakLatinWord=break_latin_word)

    def brush_type(self, brush_type):
        return self.hwp.BrushType(BrushType=brush_type)

    def canonical(self, canonical):
        return self.hwp.Canonical(Canonical=canonical)

    def cell_apply(self, cell_apply):
        return self.hwp.CellApply(CellApply=cell_apply)

    def char_shadow_type(self, shadow_type):
        return self.hwp.CharShadowType(ShadowType=shadow_type)

    def check_xobject(self, bstring):
        return self.hwp.CheckXObject(bstring=bstring)

    def clear(self, option: int = 1):
        """
        현재 편집중인 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다.

        :param option:
            편집중인 문서의 내용에 대한 처리 방법, 생략하면 1(hwpDiscard)가 선택된다.
            0: 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다. (hwpAskSave)
            1: 문서의 내용을 버린다. (hwpDiscard, 기본값)
            2: 문서가 변경된 경우 저장한다. (hwpSaveIfDirty)
            3: 무조건 저장한다. (hwpSave)

        :return:
            None

        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> hwp.clear()
        """
        return self.hwp.XHwpDocuments.Active_XHwpDocument.Clear(option=option)

    def close(self, is_dirty: bool = False):
        return self.hwp.XHwpDocuments.Active_XHwpDocument.Close(isDirty=is_dirty)

    def col_def_type(self, col_def_type):
        return self.hwp.ColDefType(ColDefType=col_def_type)

    def col_layout_type(self, col_layout_type):
        return self.hwp.ColLayoutType(ColLayoutType=col_layout_type)

    def convert_pua_hangul_to_unicode(self, reverse):
        return self.hwp.ConvertPUAHangulToUnicode(Reverse=reverse)

    def create_action(self, actidstr: str):
        """
        Action 객체를 생성한다.
        액션에 대한 세부적인 제어가 필요할 때 사용한다.
        예를 들어 기능을 수행하지 않고 대화상자만을 띄운다든지,
        대화상자 없이 지정한 옵션에 따라 기능을 수행하는 등에 사용할 수 있다.

        :param actidstr:
            액션 ID (ActionIDTable.hwp 참조)

        :return:
            Action object

        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> # 현재 커서의 폰트 크기(Height)를 구하는 코드
            >>> act = hwp.hwp.CreateAction("CharShape")
            >>> cs = act.CreateSet()  # == cs = self.hwp.CreateSet(act)
            >>> act.GetDefault(cs)
            >>> print(cs.Item("Height"))
            2800

            >>> # 현재 선택범위의 폰트 크기를 20pt로 변경하는 코드
            >>> act = hwp.hwp.CreateAction("CharShape")
            >>> cs = act.CreateSet()  # == cs = self.hwp.CreateSet(act)
            >>> act.GetDefault(cs)
            >>> cs.SetItem("Height", self.hwp.PointToHwpUnit(20))
            >>> act.Execute(cs)
            True

        """
        return self.hwp.CreateAction(actidstr=actidstr)

    def create_field(self, name: str, direction: str = "", memo: str = "") -> bool:
        """
        캐럿의 현재 위치에 누름틀을 생성한다.

        :param name:
            누름틀 필드에 대한 필드 이름(중요)

        :param direction:
            누름틀에 입력이 안 된 상태에서 보이는 안내문/지시문.

        :param memo:
            누름틀에 대한 설명/도움말

        :return:
            성공이면 True, 실패면 False

        :example:
            >>> hwp.create_field(direction="이름", memo="이름을 입력하는 필드", name="name")
            True
            >>> hwp.PutFieldText("name", "일코")
        """
        return self.hwp.CreateField(Direction=direction, memo=memo, name=name)

    def create_id(self, creation_id):
        return self.hwp.CreateID(CreationID=creation_id)

    def create_mode(self, creation_mode):
        return self.hwp.CreateMode(CreationMode=creation_mode)

    def create_page_image(self, path: str, pgno: int = 0, resolution: int = 300, depth: int = 24,
                          format: str = "bmp") -> bool:
        """
        지정된 페이지를 이미지파일로 저장한다.
        저장되는 이미지파일의 포맷은 비트맵 또는 GIF 이미지이다.
        만약 이 외의 포맷이 입력되었다면 비트맵으로 저장한다.

        :param path:
            생성할 이미지 파일의 경로(전체경로로 입력해야 함)

        :param pgno:
            페이지 번호. 0부터 PageCount-1 까지. 생략하면 0이 사용된다.
            페이지 복수선택은 불가하므로,
            for나 while 등 반복문을 사용해야 한다.

        :param resolution:
            이미지 해상도. DPI단위(96, 300, 1200 등)로 지정한다.
            생략하면 300이 사용된다.

        :param depth:
            이미지파일의 Color Depth(1, 4, 8, 24)를 지정한다.
            생략하면 24

        :param format:
            이미지파일의 포맷. "bmp", "gif"중의 하나. 생략하면 "bmp"가 사용된다.

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> self.hwp.create_page_image("c:/Users/User/Desktop/a.bmp")
            True
        """
        if not path.lower().startswith("c:"):
            path = os.path.join(os.getcwd(), path)
        return self.hwp.CreatePageImage(Path=path, pgno=pgno, resolution=resolution, depth=depth, Format=format)

    def create_set(self, setidstr):
        """
        ParameterSet을 생성한다.
        단독으로 쓰이는 경우는 거의 없으며,
        대부분 create_action과 같이 사용한다.

        ParameterSet은 일종의 정보를 지니는 객체이다.
        어떤 Action들은 그 Action이 수행되기 위해서 정보가 필요한데
        이 때 사용되는 정보를 ParameterSet으로 넘겨준다.
        또한 한/글 컨트롤은 특정정보(ViewProperties, CellShape, CharShape 등)를
        ParameterSet으로 변환하여 넘겨주기도 한다.
        사용 가능한 ParameterSet의 ID는 ParameterSet Table.hwp문서를 참조한다.

        :param setidstr:
            생성할 ParameterSet의 ID (ParameterSet Table.hwp 참고)

        :return:
            생성된 ParameterSet Object
        """
        return self.hwp.CreateSet(setidstr=setidstr)

    def crooked_slash(self, crooked_slash):
        return self.hwp.CrookedSlash(CrookedSlash=crooked_slash)

    def ds_mark(self, diac_sym_mark):
        return self.hwp.DSMark(DiacSymMark=diac_sym_mark)

    def dbf_code_type(self, dbf_code):
        return self.hwp.DbfCodeType(DbfCode=dbf_code)

    def delete_ctrl(self, ctrl) -> bool:
        """
        문서 내 컨트롤을 삭제한다.

        :param ctrl:
            삭제할 문서 내 컨트롤

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> ctrl = self.hwp.HeadCtrl.Next.Next
            >>> if ctrl.UserDesc == "표":
            ...     self.hwp.delete_ctrl(ctrl)
            ...
            True
        """
        return self.hwp.DeleteCtrl(ctrl=ctrl)

    def delimiter(self, delimiter):
        return self.hwp.Delimiter(Delimiter=delimiter)

    def draw_aspect(self, draw_aspect):
        return self.hwp.DrawAspect(DrawAspect=draw_aspect)

    def draw_fill_image(self, fillimage):
        return self.hwp.DrawFillImage(fillimage=fillimage)

    def draw_shadow_type(self, shadow_type):
        return self.hwp.DrawShadowType(ShadowType=shadow_type)

    def encrypt(self, encrypt):
        return self.hwp.Encrypt(Encrypt=encrypt)

    def end_size(self, end_size):
        return self.hwp.EndSize(EndSize=end_size)

    def end_style(self, end_style):
        return self.hwp.EndStyle(EndStyle=end_style)

    def export_style(self, sty_filepath: str) -> bool:
        """
        현재 문서의 Style을 sty 파일로 Export한다.

        :param sty_filepath:
            Export할 sty 파일의 전체경로 문자열

        :return:
            성공시 True, 실패시 False

        :example:
            >>> self.hwp.export_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if not sty_filepath.lower().startswith("c:"):
            sty_filepath = os.path.join(os.getcwd(), sty_filepath)

        style_set = self.hwp.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.hwp.ExportStyle(param=style_set.HSet)

    def field_exist(self, field):
        """
        문서에 지정된 데이터 필드가 존재하는지 검사한다.

        :param field:
            필드이름

        :return:
            필드가 존재하면 True, 존재하지 않으면 False
        """
        return self.hwp.FieldExist(Field=field)

    def file_translate(self, cur_lang, trans_lang):
        return self.hwp.FileTranslate(curLang=cur_lang, transLang=trans_lang)

    def fill_area_type(self, fill_area):
        return self.hwp.FillAreaType(FillArea=fill_area)

    def find_ctrl(self):
        return self.hwp.FindCtrl()

    def find_dir(self, find_dir: Literal["Forward", "Backward", "AllDoc"] = "AllDoc"):
        return self.hwp.FindDir(FindDir=find_dir)

    def find_private_info(self, private_type, private_string):
        """
        개인정보를 찾는다.
        (비밀번호 설정 등의 이유, 현재 비활성화된 것으로 추정)

        :param private_type:
            보호할 개인정보 유형. 다음의 값을 하나이상 조합한다.
            0x0001: 전화번호
            0x0002: 주민등록번호
            0x0004: 외국인등록번호
            0x0008: 전자우편
            0x0010: 계좌번호
            0x0020: 신용카드번호
            0x0040: IP 주소
            0x0080: 생년월일
            0x0100: 주소
            0x0200: 사용자 정의
            0x0400: 기타

        :param private_string:
            기타 문자열. 예: "신한카드"
            0x0400 유형이 존재할 경우에만 유효하므로, 생략가능하다

        :return:
            찾은 개인정보의 유형 값. 다음과 같다.
            0x0001 : 전화번호
            0x0002 : 주민등록번호
            0x0004 : 외국인등록번호
            0x0008 : 전자우편
            0x0010 : 계좌번호
            0x0020 : 신용카드번호
            0x0040 : IP 주소
            0x0080 : 생년월일
            0x0100 : 주소
            0x0200 : 사용자 정의
            0x0400 : 기타
            개인정보가 없는 경우에는 0을 반환한다.
            또한, 검색 중 문서의 끝(end of document)을 만나면 –1을 반환한다. 이는 함수가 무한히 반복하는 것을 막아준다.
        """
        return self.hwp.FindPrivateInfo(PrivateType=private_type, PrivateString=private_string)

    def font_type(self, font_type):
        return self.hwp.FontType(FontType=font_type)

    def get_bin_data_path(self, binid):
        """
        Binary Data(Temp Image 등)의 경로를 가져온다.

        :param binid:
            바이너리 데이터의 ID 값 (1부터 시작)

        :return:
            바이너리 데이터의 경로

        :example:
            >>> path = self.hwp.GetBinDataPath(2)
            >>> print(path)
            C:/Users/User/AppData/Local/Temp/Hnc/BinData/EMB00004dd86171.jpg
        """
        return self.hwp.GetBinDataPath(binid=binid)

    def get_cur_field_name(self, option=0):
        """
        현재 캐럿이 위치하는 곳의 필드이름을 구한다.
        이 함수를 통해 현재 필드가 셀필드인지 누름틀필드인지 구할 수 있다.
        참고로, 필드 좌측에 커서가 붙어있을 때는 이름을 구할 수 있지만,
        우측에 붙어 있을 때는 작동하지 않는다.
        GetFieldList()의 옵션 중에 hwpFieldSelection(=4)옵션은 사용하지 않는다.


        :param option:
            다음과 같은 옵션을 지정할 수 있다.
            0: 모두 off. 생략하면 0이 지정된다.
            1: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다.(hwpFieldCell)
            2: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)

        :return:
            필드이름이 돌아온다.
            필드이름이 없는 경우 빈 문자열이 돌아온다.
        """
        return self.hwp.GetCurFieldName(option=option)

    def get_cur_metatag_name(self):
        return self.hwp.GetCurMetatagName()

    def get_field_list(self, number=0, option=0):
        """
        문서에 존재하는 필드의 목록을 구한다.
        문서 중에 동일한 이름의 필드가 여러 개 존재할 때는
        number에 지정한 타입에 따라 3 가지의 서로 다른 방식 중에서 선택할 수 있다.
        예를 들어 문서 중 title, body, title, body, footer 순으로
        5개의 필드가 존재할 때, hwpFieldPlain, hwpFieldNumber, HwpFieldCount
        세 가지 형식에 따라 다음과 같은 내용이 돌아온다.
        hwpFieldPlain: "title\x02body\x02title\x02body\x02footer"
        hwpFieldNumber: "title{{0}}\x02body{{0}}\x02title{{1}}\x02body{{1}}\x02footer{{0}}"
        hwpFieldCount: "title{{2}}\x02body{{2}}\x02footer{{1}}"

        :param number:
            문서 내에서 동일한 이름의 필드가 여러 개 존재할 경우
            이를 구별하기 위한 식별방법을 지정한다.
            생략하면 0(hwpFieldPlain)이 지정된다.
            0: 아무 기호 없이 순서대로 필드의 이름을 나열한다.(hwpFieldPlain)
            1: 필드이름 뒤에 일련번호가 {{#}}과 같은 형식으로 붙는다.(hwpFieldNumber)
            2: 필드이름 뒤에 그 필드의 개수가 {{#}}과 같은 형식으로 붙는다.(hwpFieldCount)

        :param option:
            다음과 같은 옵션을 조합할 수 있다. 0을 지정하면 모두 off이다.
            생략하면 0이 지정된다.
            0x01: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere과는 함께 지정할 수 없다.(hwpFieldCell)
            0x02: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)
            0x04: 선택된 내용 안에 존재하는 필드 리스트를 구한다.(HwpFieldSelection)

        :return:
            각 필드 사이를 문자코드 0x02로 구분하여 다음과 같은 형식으로 리턴 한다.
            (가장 마지막 필드에는 0x02가 붙지 않는다.)
            "필드이름#1\x02필드이름#2\x02...필드이름#n"
        """
        return self.hwp.GetFieldList(Number=number, option=option)

    def get_field_text(self, field):
        """
        지정한 필드에서 문자열을 구한다.


        :param field:
            텍스트를 구할 필드 이름의 리스트.
            다음과 같이 필드 사이를 문자 코드 0x02로 구분하여
            한 번에 여러 개의 필드를 지정할 수 있다.
            "필드이름#1\x02필드이름#2\x02...필드이름#n"
            지정한 필드 이름이 문서 중에 두 개 이상 존재할 때의 표현 방식은 다음과 같다.
            "필드이름": 이름의 필드 중 첫 번째
            "필드이름{{n}}": 지정한 이름의 필드 중 n 번째
            예를 들어 "제목{{1}}\x02본문\x02이름{{0}}" 과 같이 지정하면
            '제목'이라는 이름의 필드 중 두 번째,
            '본문'이라는 이름의 필드 중 첫 번째,
            '이름'이라는 이름의 필드 중 첫 번째를 각각 지정한다.
            즉, '필드이름'과 '필드이름{{0}}'은 동일한 의미로 해석된다.

        :return:
            텍스트 데이터가 돌아온다.
            텍스트에서 탭은 '\t'(0x9),
            문단 바뀜은 CR/LF(0x0D/0x0A == \r\n)로 표현되며,
            이외의 특수 코드는 포함되지 않는다.
            필드 텍스트의 끝은 0x02(\x02)로 표현되며,
            그 이후 다음 필드의 텍스트가 연속해서
            지정한 필드 리스트의 개수만큼 위치한다.
            지정한 이름의 필드가 없거나,
            사용자가 해당 필드에 아무 텍스트도 입력하지 않았으면
            해당 텍스트에는 빈 문자열이 돌아온다.
        """
        return self.hwp.GetFieldText(Field=field)

    def get_file_info(self, filename):
        """
        파일 정보를 알아낸다.
        한글 문서를 열기 전에 암호가 걸린 문서인지 확인할 목적으로 만들어졌다.
        (현재 한/글2022 기준으로 hwpx포맷에 대해서는 파일정보를 파악할 수 없다.)

        :param filename:
            정보를 구하고자 하는 hwp 파일의 전체 경로

        :return:
            "FileInfo" ParameterSet이 반환된다.
            파라미터셋의 ItemID는 아래와 같다.
            Format(string) : 파일의 형식.(HWP : 한/글 파일, UNKNOWN : 알 수 없음.)
            VersionStr(string) : 파일의 버전 문자열. ex)5.0.0.3
            VersionNum(unsigned int) : 파일의 버전. ex) 0x05000003
            Encrypted(int) : 암호 여부 (현재는 파일 버전 3.0.0.0 이후 문서-한/글97, 한/글 워디안 및 한/글 2002 이상의 버전-에 대해서만 판단한다.)
            (-1: 판단할 수 없음, 0: 암호가 걸려 있지 않음, 양수: 암호가 걸려 있음.)

        :example:
            >>> pset = self.hwp.GetFileInfo("C:/Users/Administrator/Desktop/이력서.hwp")
            >>> print(pset.Item("Format"))
            >>> print(pset.Item("VersionStr"))
            >>> print(hex(pset.Item("VersionNum")))
            >>> print(pset.Item("Encrypted"))
            HWP
            5.1.1.0
            0x5010100
            0
        """
        if not filename.lower().startswith("c:"):
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.GetFileInfo(filename=filename)

    def get_font_list(self, langid):
        self.scan_font()
        return self.hwp.GetFontList(langid=langid)

    def get_heading_string(self):
        """
        현재 커서가 위치한 문단의 글머리표/문단번호/개요번호를 추출한다.
        글머리표/문단번호/개요번호가 있는 경우, 해당 문자열을 얻어올 수 있다.
        문단에 글머리표/문단번호/개요번호가 없는 경우, 빈 문자열이 추출된다.

        :return:
            (글머리표/문단번호/개요번호가 있다면) 해당 문자열이 반환된다.
        """
        return self.hwp.GetHeadingString()

    def get_message_box_mode(self):
        """
        현재 메시지 박스의 Mode를 int로 얻어온다.
        set_message_box_mode와 함께 쓰인다.
        6개의 대화상자에서 각각 확인/취소/종료/재시도/무시/예/아니오 버튼을
        자동으로 선택할 수 있게 설정할 수 있으며 조합 가능하다.

        :return:
            // 메시지 박스의 종류
            MB_MASK: 0x00FFFFFF
            // 1. 확인(MB_OK) : IDOK(1)
            MB_OK_IDOK: 0x00000001
            MB_OK_MASK: 0x0000000F
            // 2. 확인/취소(MB_OKCANCEL) : IDOK(1), IDCANCEL(2)
            MB_OKCANCEL_IDOK: 0x00000010
            MB_OKCANCEL_IDCANCEL: 0x00000020
            MB_OKCANCEL_MASK: 0x000000F0
            // 3. 종료/재시도/무시(MB_ABORTRETRYIGNORE) : IDABORT(3), IDRETRY(4), IDIGNORE(5)
            MB_ABORTRETRYIGNORE_IDABORT: 0x00000100
            MB_ABORTRETRYIGNORE_IDRETRY: 0x00000200
            MB_ABORTRETRYIGNORE_IDIGNORE: 0x00000400
            MB_ABORTRETRYIGNORE_MASK: 0x00000F00
            // 4. 예/아니오/취소(MB_YESNOCANCEL) : IDYES(6), IDNO(7), IDCANCEL(2)
            MB_YESNOCANCEL_IDYES: 0x00001000
            MB_YESNOCANCEL_IDNO: 0x00002000
            MB_YESNOCANCEL_IDCANCEL: 0x00004000
            MB_YESNOCANCEL_MASK: 0x0000F000
            // 5. 예/아니오(MB_YESNO) : IDYES(6), IDNO(7)
            MB_YESNO_IDYES: 0x00010000
            MB_YESNO_IDNO: 0x00020000
            MB_YESNO_MASK: 0x000F0000
            // 6. 재시도/취소(MB_RETRYCANCEL) : IDRETRY(4), IDCANCEL(2)
            MB_RETRYCANCEL_IDRETRY: 0x00100000
            MB_RETRYCANCEL_IDCANCEL: 0x00200000
            MB_RETRYCANCEL_MASK: 0x00F00000
        """
        return self.hwp.GetMessageBoxMode()

    def get_metatag_list(self, number, option):
        return self.hwp.GetMetatagList(Number=number, option=option)

    def get_metatag_name_text(self, tag):
        return self.hwp.GetMetatagNameText(tag=tag)

    def get_mouse_pos(self, x_rel_to=1, y_rel_to=1):
        """
        마우스의 현재 위치를 얻어온다.
        단위가 HWPUNIT임을 주의해야 한다.
        (1 inch = 7200 HWPUNIT, 1mm = 283.465 HWPUNIT)

        :param x_rel_to:
            X좌표계의 기준 위치(기본값은 1:쪽기준)
            0: 종이 기준으로 좌표를 가져온다.
            1: 쪽 기준으로 좌표를 가져온다.

        :param y_rel_to:
            Y좌표계의 기준 위치(기본값은 1:쪽기준)
            0: 종이 기준으로 좌표를 가져온다.
            1: 쪽 기준으로 좌표를 가져온다.

        :return:
            "MousePos" ParameterSet이 반환된다.
            아이템ID는 아래와 같다.
            XRelTo(unsigned long): 가로 상대적 기준(0: 종이, 1: 쪽)
            YRelTo(unsigned long): 세로 상대적 기준(0: 종이, 1: 쪽)
            Page(unsigned long): 페이지 번호(0-based)
            X(long): 가로 클릭한 위치(HWPUNIT)
            Y(long): 세로 클릭한 위치(HWPUNIT)

        :example:
            >>> pset = self.hwp.GetMousePos(1, 1)
            >>> print("X축 기준:", "쪽" if pset.Item("XRelTo") else "종이")
            >>> print("Y축 기준:", "쪽" if pset.Item("YRelTo") else "종이")
            >>> print("현재", pset.Item("Page")+1, "페이지에 커서 위치")
            >>> print("좌상단 기준 우측으로", int(pset.Item("X") / 283.465), "mm에 위치")
            >>> print("좌상단 기준 아래로", int(pset.Item("Y") / 283.465), "mm에 위치")
            X축 기준: 쪽
            Y축 기준: 쪽
            현재 2 페이지에 커서 위치
            좌상단 기준 우측으로 79 mm에 위치
            좌상단 기준 아래로 217 mm에 위치
        """
        return self.hwp.GetMousePos(XRelTo=x_rel_to, YRelTo=y_rel_to)

    def get_page_text(self, pgno: int = 0, option: hex = 0xffffffff) -> str:
        """
        페이지 단위의 텍스트 추출
        일반 텍스트(글자처럼 취급 도형 포함)를 우선적으로 추출하고,
        도형(표, 글상자) 내의 텍스트를 추출한다.
        팁: get_text로는 글머리를 추출하지 않지만, get_page_text는 추출한다.
        팁2: 아무리 get_page_text라도 유일하게 표번호는 추출하지 못한다.
        표번호는 XML태그 속성 안에 저장되기 때문이다.

        :param pgno:
            텍스트를 추출 할 페이지의 번호(0부터 시작)

        :param option:
            추출 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다.
            생략(또는 0xffffffff)하면 모든 텍스트를 추출한다.
            0x00: 본문 텍스트만 추출한다.(maskNormal)
            0x01: 표에대한 텍스트를 추출한다.(maskTable)
            0x02: 글상자 텍스트를 추출한다.(maskTextbox)
            0x04: 캡션 텍스트를 추출한다. (표, ShapeObject)(maskCaption)

        :return:
            해당 페이지의 텍스트가 추출된다.
            글머리는 추출하지만, 표번호는 추출하지 못한다.
        """
        return self.hwp.GetPageText(pgno=pgno, option=option)

    def get_pos(self) -> tuple[int]:
        """
        캐럿의 위치를 얻어온다.
        파라미터 중 리스트는, 문단과 컨트롤들이 연결된 한/글 문서 내 구조를 뜻한다.
        리스트 아이디는 문서 내 위치 정보 중 하나로서 SelectText에 넘겨줄 때 사용한다.
        (파이썬 자료형인 list가 아님)

        :return:
            (List, para, pos) 튜플.
            list: 캐럿이 위치한 문서 내 list ID(본문이 0)
            para: 캐럿이 위치한 문단 ID(0부터 시작)
            pos: 캐럿이 위치한 문단 내 글자 위치(0부터 시작)

        """
        return self.hwp.GetPos()

    def get_pos_by_set(self):
        """
        현재 캐럿의 위치 정보를 ParameterSet으로 얻어온다.
        해당 파라미터셋은 set_pos_by_set에 직접 집어넣을 수 있어 간편히 사용할 수 있다.

        :return:
            캐럿 위치에 대한 ParameterSet
            해당 파라미터셋의 아이템은 아래와 같다.
            "List": 캐럿이 위치한 문서 내 list ID(본문이 0)
            "Para": 캐럿이 위치한 문단 ID(0부터 시작)
            "Pos": 캐럿이 위치한 문단 내 글자 위치(0부터 시작)

        :example:
            >>> pset = self.hwp.get_pos_by_set()  # 캐럿위치 저장
            >>> print(pset.Item("List"))
            6
            >>> print(pset.Item("Para"))
            3
            >>> print(pset.Item("Pos"))
            2
            >>> self.hwp.set_pos_by_set(pset)  # 캐럿위치 복원
            True
        """
        return self.hwp.GetPosBySet()

    def get_script_source(self, filename: str) -> str:
        """
        문서에 포함된 매크로(스크립트매크로 제외) 소스코드를 가져온다.
        문서포함 매크로는 기본적으로
        ```
        function OnDocument_New() {
        }
        function OnDocument_Open() {
        }
        ```
        형태로 비어있는 상태이며,
        OnDocument_New와 OnDocument_Open 두 개의 함수에 한해서만
        코드를 추가하고 실행할 수 있다.

        :param filename:
            매크로 소스를 가져올 한/글 문서의 전체경로

        :return:
            (문서에 포함된) 스크립트의 소스코드

        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> print(hwp.get_script_source("C:/Users/User/Desktop/script.hwp"))
            function OnDocument_New()
            {
                HAction.GetDefault("InsertText", HParameterSet.HInsertText.HSet);
                HParameterSet.HInsertText.Text = "ㅁㄴㅇㄹㅁㄴㅇㄹ";
                HAction.Execute("InsertText", HParameterSet.HInsertText.HSet);
            }
            function OnDocument_Open()
            {
                HAction.GetDefault("InsertText", HParameterSet.HInsertText.HSet);
                HParameterSet.HInsertText.Text = "ㅋㅌㅊㅍㅋㅌㅊㅍ";
                HAction.Execute("InsertText", HParameterSet.HInsertText.HSet);
            }
        """
        if not filename.lower().startswith("c:"):
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.GetScriptSource(filename=filename)

    def get_selected_pos(self):
        """
        현재 설정된 블록의 위치정보를 얻어온다.

        :return:
            블록상태여부, 시작과 끝위치 인덱스인 6개 정수 등 7개 요소의 튜플을 리턴
            (is_block, slist, spara, spos, elist, epara, epos)
            is_block: 현재 블록선택상태 여부(블록상태이면 True)
            slist: 설정된 블록의 시작 리스트 아이디.
            spara: 설정된 블록의 시작 문단 아이디.
            spos: 설정된 블록의 문단 내 시작 글자 단위 위치.
            elist: 설정된 블록의 끝 리스트 아이디.
            epara: 설정된 블록의 끝 문단 아이디.
            epos: 설정된 블록의 문단 내 끝 글자 단위 위치.

        :example:
            >>> self.hwp.get_selected_pos()
            (True, 0, 0, 16, 0, 7, 16)
        """
        return self.hwp.GetSelectedPos()

    def get_selected_pos_by_set(self, sset, eset):
        """
        현재 설정된 블록의 위치정보를 얻어온다.
        (GetSelectedPos의 ParameterSet버전)
        실행 전 GetPos 형태의 파라미터셋 두 개를 미리 만들어서
        인자로 넣어줘야 한다.

        :param sset:
            설정된 블록의 시작 파라메터셋 (ListParaPos)

        :param eset:
            설정된 블록의 끝 파라메터셋 (ListParaPos)

        :return:
            성공하면 True, 실패하면 False.
            실행시 sset과 eset의 아이템 값이 업데이트된다.

        :example:
            >>> sset = self.hwp.get_pos_by_set()
            >>> eset = self.hwp.get_pos_by_set()
            >>> self.hwp.GetSelectedPosBySet(sset, eset)
            >>> self.hwp.SetPosBySet(eset)
            True
        """
        return self.hwp.GetSelectedPosBySet(sset=sset, eset=eset)

    def get_text(self):
        """
        문서 내에서 텍스트를 얻어온다.
        줄바꿈 기준으로 텍스트를 얻어오므로 반복실행해야 한다.
        get_text()의 사용이 끝나면 release_scan()을 반드시 호출하여
        관련 정보를 초기화 해주어야 한다.
        get_text()로 추출한 텍스트가 있는 문단으로 캐럿을 이동 시키려면
        move_pos(201)을 실행하면 된다.

        :return:
            (state: int, text: str) 형태의 튜플을 리턴한다.
            state의 의미는 아래와 같다.
            0: 텍스트 정보 없음
            1: 리스트의 끝
            2: 일반 텍스트
            3: 다음 문단
            4: 제어문자 내부로 들어감
            5: 제어문자를 빠져나옴
            101: 초기화 안 됨(init_scan() 실패 또는 init_scan()을 실행하지 않은 경우)
            102: 텍스트 변환 실패
            text는 추출한 텍스트 데이터이다.
            텍스트에서 탭은 '\t'(0x9), 문단 바뀜은 '\r\n'(0x0D/0x0A)로 표현되며,
            이외의 특수 코드는 포함되지 않는다.

        :example:
            >>> self.hwp.init_scan()
            >>> while True:
            ...     state, text = self.hwp.get_text()
            ...     print(state, text)
            ...     if state <= 1:
            ...         break
            ... self.hwp.release_scan()
            2
            2
            2 ㅁㄴㅇㄹ
            3
            4 ㅂㅈㄷㄱ
            2 ㅂㅈㄷㄱ
            5
            1

        """
        return self.hwp.GetText()

    def get_text_file(self, format="UNICODE", option=""):
        """
        현재 열린 문서를 문자열로 넘겨준다.
        이 함수는 JScript나 VBScript와 같이
        직접적으로 local disk를 접근하기 힘든 언어를 위해 만들어졌으므로
        disk를 접근할 수 있는 언어에서는 사용하지 않기를 권장.
        disk를 접근할 수 있다면, Save나 SaveBlockAction을 사용할 것.
        이 함수 역시 내부적으로는 save나 SaveBlockAction을 호출하도록 되어있고
        텍스트로 저장된 파일이 메모리에서 3~4번 복사되기 때문에 느리고,
        메모리를 낭비함.
        팁: HTML로 추출시 표번호가 유지된다.

        :param format:
            파일의 형식. 기본값은 "UNICODE"
            "HWP": HWP native format, BASE64로 인코딩되어 있다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다.ver:0x0505010B
            "HWPML2X": HWP 형식과 호환. 문서의 모든 정보를 유지
            "HTML": 인터넷 문서 HTML 형식. 한/글 고유의 서식은 손실된다.
            "UNICODE": 유니코드 텍스트, 서식정보가 없는 텍스트만 저장.
            "TEXT": 일반 텍스트. 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다.
            소문자로 입력해도 된다.

        :param option:
            "saveblock": 선택된 블록만 저장. 개체 선택 상태에서는 동작하지 않는다.
            기본값은 빈 문자열("")

        :return:
            지정된 포맷에 맞춰 파일을 문자열로 변환한 값을 반환한다.

        :example:
            >>> self.hwp.get_text_file()
            'ㅁㄴㅇㄹ\r\nㅁㄴㅇㄹ\r\nㅁㄴㅇㄹ\r\n\r\nㅂㅈㄷㄱ\r\nㅂㅈㄷㄱ\r\nㅂㅈㄷㄱ\r\n'
        """
        return self.hwp.GetTextFile(Format=format, option=option)

    def get_translate_lang_list(self, cur_lang):
        return self.hwp.GetTranslateLangList(curLang=cur_lang)

    def get_user_info(self, user_info_id):
        return self.hwp.GetUserInfo(userInfoId=user_info_id)

    def gradation(self, gradation):
        return self.hwp.Gradation(Gradation=gradation)

    def grid_method(self, grid_method):
        return self.hwp.GridMethod(GridMethod=grid_method)

    def grid_view_line(self, grid_view_line):
        return self.hwp.GridViewLine(GridViewLine=grid_view_line)

    def gutter_method(self, gutter_type):
        return self.hwp.GutterMethod(GutterType=gutter_type)

    def h_align(self, h_align):
        return self.hwp.HAlign(HAlign=h_align)

    def handler(self, handler):
        return self.hwp.Handler(Handler=handler)

    def hash(self, hash):
        return self.hwp.Hash(Hash=hash)

    def hatch_style(self, hatch_style):
        return self.hwp.HatchStyle(HatchStyle=hatch_style)

    def head_type(self, heading_type):
        return self.hwp.HeadType(HeadingType=heading_type)

    def height_rel(self, height_rel):
        return self.hwp.HeightRel(HeightRel=height_rel)

    def hiding(self, hiding):
        return self.hwp.Hiding(Hiding=hiding)

    def horz_rel(self, horz_rel):
        return self.hwp.HorzRel(HorzRel=horz_rel)

    def hwp_line_type(self, line_type):
        return self.hwp.HwpLineType(LineType=line_type)

    def hwp_line_width(self, line_width):
        return self.hwp.HwpLineWidth(LineWidth=line_width)

    def hwp_outline_style(self, hwp_outline_style):
        return self.hwp.HwpOutlineStyle(HwpOutlineStyle=hwp_outline_style)

    def hwp_outline_type(self, hwp_outline_type):
        return self.hwp.HwpOutlineType(HwpOutlineType=hwp_outline_type)

    def hwp_underline_shape(self, hwp_underline_shape):
        return self.hwp.HwpUnderlineShape(HwpUnderlineShape=hwp_underline_shape)

    def hwp_underline_type(self, hwp_underline_type):
        return self.hwp.HwpUnderlineType(HwpUnderlineType=hwp_underline_type)

    def hwp_zoom_type(self, zoom_type):
        return self.hwp.HwpZoomType(ZoomType=zoom_type)

    def image_format(self, image_format):
        return self.hwp.ImageFormat(ImageFormat=image_format)

    def import_style(self, sty_filepath):
        """
        미리 저장된 특정 sty파일의 스타일을 임포트한다.

        :param sty_filepath:
            sty파일의 경로

        :return:
            성공시 True, 실패시 False

        :example:
            >>> self.hwp.import_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if not sty_filepath.lower().startswith("c:"):
            sty_filepath = os.path.join(os.getcwd(), sty_filepath)

        style_set = self.hwp.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.hwp.ImportStyle(style_set.HSet)

    def init_hparameter_set(self):
        return self.hwp.InitHParameterSet()

    def init_scan(self, option=0x07, range=0x77, spara=0, spos=0, epara=-1, epos=-1):
        """
        문서의 내용을 검색하기 위해 초기설정을 한다.
        문서의 검색 과정은 InitScan()으로 검색위한 준비 작업을 하고
        GetText()를 호출하여 본문의 텍스트를 얻어온다.
        GetText()를 반복호출하면 연속하여 본문의 텍스트를 얻어올 수 있다.
        검색이 끝나면 ReleaseScan()을 호출하여 관련 정보를 Release해야 한다.

        :param option:
            찾을 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다.
            생략하면 모든 컨트롤을 찾을 대상으로 한다.
            0x00: 본문을 대상으로 검색한다.(서브리스트를 검색하지 않는다.) - maskNormal
            0x01: char 타입 컨트롤 마스크를 대상으로 한다.(강제줄나눔, 문단 끝, 하이픈, 묶움빈칸, 고정폭빈칸, 등...) - maskChar
            0x02: inline 타입 컨트롤 마스크를 대상으로 한다.(누름틀 필드 끝, 등...) - maskInline
            0x04: extende 타입 컨트롤 마스크를 대상으로 한다.(바탕쪽, 프레젠테이션, 다단, 누름틀 필드 시작, Shape Object, 머리말, 꼬리말, 각주, 미주, 번호관련 컨트롤, 새 번호 관련 컨트롤, 감추기, 찾아보기, 글자 겹침, 등...) - maskCtrl

        :param range:
            검색의 범위를 다음과 같은 옵션을 조합(sum)하여 지정할 수 있다.
            생략하면 "문서 시작부터 - 문서의 끝까지" 검색 범위가 지정된다.
            0x0000: 캐럿 위치부터. (시작 위치) - scanSposCurrent
            0x0010: 특정 위치부터. (시작 위치) - scanSposSpecified
            0x0020: 줄의 시작부터. (시작 위치) - scanSposLine
            0x0030: 문단의 시작부터. (시작 위치) - scanSposParagraph
            0x0040: 구역의 시작부터. (시작 위치) - scanSposSection
            0x0050: 리스트의 시작부터. (시작 위치) - scanSposList
            0x0060: 컨트롤의 시작부터. (시작 위치) - scanSposControl
            0x0070: 문서의 시작부터. (시작 위치) - scanSposDocument
            0x0000: 캐럿 위치까지. (끝 위치) - scanEposCurrent
            0x0001: 특정 위치까지. (끝 위치) - scanEposSpecified
            0x0002: 줄의 끝까지. (끝 위치) - scanEposLine
            0x0003: 문단의 끝까지. (끝 위치) - scanEposParagraph
            0x0004: 구역의 끝까지. (끝 위치) - scanEposSection
            0x0005: 리스트의 끝까지. (끝 위치) - scanEposList
            0x0006: 컨트롤의 끝까지. (끝 위치) - scanEposControl
            0x0007: 문서의 끝까지. (끝 위치) - scanEposDocument
            0x00ff: 검색의 범위를 블록으로 제한. - scanWithinSelection
            0x0000: 정뱡향. (검색 방향) - scanForward
            0x0100: 역방향. (검색 방향) - scanBackward

        :param spara:
            검색 시작 위치의 문단 번호.
            scanSposSpecified 옵션이 지정되었을 때만 유효하다.
            예) range=0x0011

        :param spos:
            검색 시작 위치의 문단 중에서 문자의 위치.
            scanSposSpecified 옵션이 지정되었을 때만 유효하다.
            예) range=0x0011

        :param epara:
            검색 끝 위치의 문단 번호.
            scanEposSpecified 옵션이 지정되었을 때만 유효하다.

        :param epos:
            검색 끝 위치의 문단 중에서 문자의 위치.
            scanEposSpecified 옵션이 지정되었을 때만 유효하다.

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> self.hwp.init_scan(range=0xff)
            >>> _, text = self.hwp.get_text()
            >>> self.hwp.release_scan()
            >>> print(text)
            Hello, world!
        """
        return self.hwp.InitScan(option=option, Range=range, spara=spara,
                                 spos=spos, epara=epara, epos=epos)

    def insert(self, path, format="", arg=""):
        """
        현재 캐럿 위치에 문서파일을 삽입한다.
        format, arg에 대해서는 self.hwp.open 참조

        :param path:
            문서파일의 경로

        :param format:
            문서형식. **빈 문자열을 지정하면 자동으로 선택한다.**
            생략하면 빈 문자열이 지정된다.
            아래에 쓰여 있는 대로 대문자로만 써야 한다.
            "HWPX": 한/글 hwpx format
            "HWP": 한/글 native format
            "HWP30": 한/글 3.X/96/97
            "HTML": 인터넷 문서
            "TEXT": 아스키 텍스트 문서
            "UNICODE": 유니코드 텍스트 문서
            "HWP20": 한글 2.0
            "HWP21": 한글 2.1/2.5
            "HWP15": 한글 1.X
            "HWPML1X": HWPML 1.X 문서 (Open만 가능)
            "HWPML2X": HWPML 2.X 문서 (Open / SaveAs 가능)
            "RTF": 서식 있는 텍스트 문서
            "DBF": DBASE II/III 문서
            "HUNMIN": 훈민정음 3.0/2000
            "MSWORD": 마이크로소프트 워드 문서
            "DOCRTF": MS 워드 문서 (doc)
            "OOXML": MS 워드 문서 (docx)
            "HANA": 하나워드 문서
            "ARIRANG": 아리랑 문서
            "ICHITARO": 一太郞 문서 (일본 워드프로세서)
            "WPS": WPS 문서
            "DOCIMG": 인터넷 프레젠테이션 문서(SaveAs만 가능)
            "SWF": Macromedia Flash 문서(SaveAs만 가능)

        :param arg:
            세부옵션. 의미는 format에 지정한 파일형식에 따라 다르다.
            조합 가능하며, 생략하면 빈 문자열이 지정된다.
            <공통>
            "setcurdir:FALSE;" :로드한 후 해당 파일이 존재하는 폴더로 현재 위치를 변경한다. hyperlink 정보가 상대적인 위치로 되어 있을 때 유용하다.
            <HWP/HWPX>
            "lock:TRUE;": 로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
            "notext:FALSE;": 텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)
            "template:FALSE;": 새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다. 이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.
            "suspendpassword:FALSE;": TRUE로 지정하면 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.
            "forceopen:FALSE;": TRUE로 지정하면 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다.
            "versionwarning:FALSE;": TRUE로 지정하면 문서가 상위버전일 경우 메시지 박스를 띄우게 된다.
            <HTML>
            "code"(string, codepage): 문서변환 시 사용되는 코드 페이지를 지정할 수 있으며 code키가 존재할 경우 필터사용 시 사용자 다이얼로그를  띄우지 않는다.
            (코드페이지 종류는 아래와 같다.)
            ("utf8" : UTF8)
            ("unicode": 유니코드)
            ("ks":  한글 KS 완성형)
            ("acp" : Active Codepage 현재 시스템의 코드 페이지)
            ("kssm": 한글 조합형)
            ("sjis" : 일본)
            ("gb" : 중국 간체)
            ("big5" : 중국 번체)
            "textunit:(string, pixel);": Export될 Text의 크기의 단위 결정.pixel, point, mili 지정 가능.
            "formatunit:(string, pixel);": Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능
            <DOCIMG>
            "asimg:FALSE;": 저장할 때 페이지를 image로 저장
            "ashtml:FALSE;": 저장할 때 페이지를 html로 저장
            <TEXT>
            "code:(string, codepage);": 문서 변환 시 사용되는 코드 페이지를 지정할 수 있으며
            code키가 존재할 경우 필터 사용 시 사용자 다이얼로그를  띄우지 않는다.

        :return:
            성공하면 True, 실패하면 False
        """
        if not path.lower().startswith("c:"):
            path = os.path.join(os.getcwd(), path)
        return self.hwp.Insert(Path=path, Format=format, arg=arg)

    def insert_background_picture(self, path, border_type="SelectedCell",
                                  embedded=True, filloption=5, effect=1,
                                  watermark=False, brightness=0, contrast=0) -> bool:
        """
        **셀**에 배경이미지를 삽입한다.
        CellBorderFill의 SetItem 중 FillAttr 의 SetItem FileName 에
        이미지의 binary data를 지정해 줄 수가 없어서 만든 함수다.
        기타 배경에 대한 다른 조정은 Action과 ParameterSet의 조합으로 가능하다.

        :param path:
            삽입할 이미지 파일

        :param border_type:
            배경 유형을 문자열로 지정(파라미터 이름과는 다르게 삽입/제거 기능이다.)
            "SelectedCell": 현재 선택된 표의 셀의 배경을 변경한다.
            "SelectedCellDelete": 현재 선택된 표의 셀의 배경을 지운다.
            단, 배경 제거시 반드시 셀이 선택되어 있어야함.
            커서가 위치하는 것만으로는 동작하지 않음.

        :param embedded:
            이미지 파일을 문서 내에 포함할지 여부 (True/False). 생략하면 True

        :param filloption:
            삽입할 그림의 크기를 지정하는 옵션
            0: 바둑판식으로 - 모두
            1: 바둑판식으로 - 가로/위
            2: 바둑판식으로 - 가로/아로
            3: 바둑판식으로 - 세로/왼쪽
            4: 바둑판식으로 - 세로/오른쪽
            5: 크기에 맞추어(기본값)
            6: 가운데로
            7: 가운데 위로
            8: 가운데 아래로
            9: 왼쪽 가운데로
            10: 왼쪽 위로
            11: 왼쪽 아래로
            12: 오른쪽 가운데로
            13: 오른쪽 위로
            14: 오른쪽 아래로

        :param effect:
            이미지효과
            0: 원래 그림(기본값)
            1: 그레이 스케일
            2: 흑백으로

        :param watermark:
            watermark효과 유무 (True/False)
            기본값은 False
            이 옵션이 True이면 brightness 와 contrast 옵션이 무시된다.

        :param brightness:
            밝기 지정(-100 ~ 100), 기본 값은 0

        :param contrast:
            선명도 지정(-100 ~ 100), 기본 값은 0

        :return:
            성공했을 경우 True, 실패했을 경우 False

        :example:
            >>> self.hwp.insert_background_picture(path="C:/Users/User/Desktop/KakaoTalk_20230709_023118549.jpg")
            True
        """
        if not path.lower().startswith("c:"):
            path = os.path.join(os.getcwd(), path)

        return self.hwp.InsertBackgroundPicture(Path=path, BorderType=border_type,
                                                Embedded=embedded, filloption=filloption,
                                                Effect=effect, watermark=watermark,
                                                Brightness=brightness, Contrast=contrast)

    def insert_ctrl(self, ctrl_id, initparam):
        """
        현재 캐럿 위치에 컨트롤을 삽입한다.
        ctrlid에 지정할 수 있는 컨트롤 ID는 HwpCtrl.CtrlID가 반환하는 ID와 동일하다.
        자세한 것은  Ctrl 오브젝트 Properties인 CtrlID를 참조.
        initparam에는 컨트롤의 초기 속성을 지정한다.
        대부분의 컨트롤은 Ctrl.Properties와 동일한 포맷의 parameter set을 사용하지만,
        컨트롤 생성 시에는 다른 포맷을 사용하는 경우도 있다.
        예를 들어 표의 경우 Ctrl.Properties에는 "Table" 셋을 사용하지만,
        생성 시 initparam에 지정하는 값은 "TableCreation" 셋이다.

        :param ctrl_id:
            삽입할 컨트롤 ID

        :param initparam:
            컨트롤 초기속성. 생략하면 default 속성으로 생성한다.

        :return:
            생성된 컨트롤 object

        :example:
            >>> # 3행5열의 표를 삽입한다.
            >>> from time import sleep
            >>> tbset = self.hwp.CreateSet("TableCreation")
            >>> tbset.SetItem("Rows", 3)
            >>> tbset.SetItem("Cols", 5)
            >>> row_set = tbset.CreateItemArray("RowHeight", 3)
            >>> col_set = tbset.CreateItemArray("ColWidth", 5)
            >>> row_set.SetItem(0, self.hwp.PointToHwpUnit(10))
            >>> row_set.SetItem(1, self.hwp.PointToHwpUnit(10))
            >>> row_set.SetItem(2, self.hwp.PointToHwpUnit(10))
            >>> col_set.SetItem(0, self.hwp.MiliToHwpUnit(26))
            >>> col_set.SetItem(1, self.hwp.MiliToHwpUnit(26))
            >>> col_set.SetItem(2, self.hwp.MiliToHwpUnit(26))
            >>> col_set.SetItem(3, self.hwp.MiliToHwpUnit(26))
            >>> col_set.SetItem(4, self.hwp.MiliToHwpUnit(26))
            >>> table = self.hwp.InsertCtrl("tbl", tbset)
            >>> sleep(3)  # 표 생성 3초 후 다시 표 삭제
            >>> self.hwp.delete_ctrl(table)


        """
        return self.hwp.InsertCtrl(CtrlID=ctrl_id, initparam=initparam)

    def insert_picture(self, path, embedded=True, sizeoption=2, reverse=False, watermark=False, effect=0, width=0,
                       height=0):
        """
        현재 캐럿의 위치에 그림을 삽입한다.
        다만, 그림의 종횡비를 유지한 채로 셀의 높이만 키워주는 옵션이 없다.
        이런 작업을 원하는 경우에는 그림을 클립보드로 복사하고,
        Ctrl-V로 붙여넣기를 하는 수 밖에 없다.
        또한, 셀의 크기를 조절할 때 이미지의 크기도 따라 변경되게 하고 싶다면
        insert_background_picture 함수를 사용하는 것도 좋다.

        :param path:
            삽입할 이미지 파일의 전체경로

        :param embedded:
            이미지 파일을 문서 내에 포함할지 여부 (True/False). 생략하면 True

        :param sizeoption:
            삽입할 그림의 크기를 지정하는 옵션. 기본값은 2
            0: 이미지 원래의 크기로 삽입한다. width와 height를 지정할 필요 없다.(realSize)
            1: width와 height에 지정한 크기로 그림을 삽입한다.(specificSize)
            2: 현재 캐럿이 표의 셀 안에 있을 경우, 셀의 크기에 맞게 자동 조절하여 삽입한다. (종횡비 유지안함)(cellSize)
               캐럿이 셀 안에 있지 않으면 이미지의 원래 크기대로 삽입된다.
            3: 현재 캐럿이 표의 셀 안에 있을 경우, 셀의 크기에 맞추어 원본 이미지의 가로 세로의 비율이 동일하게 확대/축소하여 삽입한다.(cellSizeWithSameRatio)

        :param reverse: 이미지의 반전 유무 (True/False). 기본값은 False

        :param watermark: watermark효과 유무 (True/False). 기본값은 False

        :param effect:
            그림 효과
            0: 실제 이미지 그대로
            1: 그레이 스케일
            2: 흑백효과

        :param width:
            그림의 가로 크기 지정. 단위는 mm(HWPUNIT 아님!)

        :param height:
            그림의 높이 크기 지정. 단위는 mm

        :return:
            생성된 컨트롤 object.

        :example:
            >>> ctrl = self.hwp.insert_picture("C:/Users/Administrator/Desktop/KakaoTalk_20230709_023118549.jpg")
            >>> pset = ctrl.Properties  # == self.hwp.create_set("ShapeObject")
            >>> pset.SetItem("TreatAsChar", False)  # 글자처럼취급 해제
            >>> pset.SetItem("TextWrap", 2)  # 그림을 글 뒤로
            >>> ctrl.Properties = pset  # 설정한 값 적용(간단!)
        """
        if not path.lower().startswith("c:"):
            path = os.path.join(os.getcwd(), path)

        return self.hwp.InsertPicture(Path=path, Embedded=embedded, sizeoption=sizeoption,
                                      Reverse=reverse, watermark=watermark, Effect=effect,
                                      Width=width, Height=height)

    def is_action_enable(self, action_id):
        return self.hwp.IsActionEnable(actionID=action_id)

    def is_command_lock(self, action_id):
        """
        해당 액션이 잠겨있는지 확인한다.

        :param action_id: 액션 ID. (ActionIDTable.Hwp 참조)

        :return:
            잠겨있으면 True, 잠겨있지 않으면 False를 반환한다.
        """
        return self.hwp.IsCommandLock(actionID=action_id)

    def key_indicator(self) -> tuple:
        """
        상태 바의 정보를 얻어온다.
        (캐럿이 표 안에 있을 때 셀의 주소를 얻어오는 거의 유일한 방법이다.)

        :return:
            튜플(succ, seccnt, secno, prnpageno, colno, line, pos, over, ctrlname)
            succ: 성공하면 True, 실패하면 False (항상 True임..)
            seccnt: 총 구역
            secno: 현재 구역
            prnpageno: 쪽
            colno: 단
            line: 줄
            pos: 칸
            over: 삽입모드 (True: 수정, False: 삽입)
            ctrlname: 캐럿이 위치한 곳의 컨트롤이름

        :example:
            >>> # 현재 셀 주소(표 안에 있을 때)
            >>> self.hwp.KeyIndicator()[-1][1:].split(")")[0]
            "A1"
        """
        return self.hwp.KeyIndicator()

    def line_spacing_method(self, line_spacing):
        return self.hwp.LineSpacingMethod(LineSpacing=line_spacing)

    def line_wrap_type(self, line_wrap):
        return self.hwp.LineWrapType(LineWrap=line_wrap)

    def lock_command(self, act_id, is_lock):
        """
        특정 액션이 실행되지 않도록 잠근다.

        :param act_id: 액션 ID. (ActionIDTable.Hwp 참조)

        :param is_lock:
            True이면 액션의 실행을 잠그고, False이면 액션이 실행되도록 한다.

        :return: None

        :example:
            >>> # Undo와 Redo 잠그기
            >>> self.hwp.LockCommand("Undo", True)
            >>> self.hwp.LockCommand("Redo", True)
        """
        return self.hwp.LockCommand(ActID=act_id, isLock=is_lock)

    def lunar_to_solar(self, l_year, l_month, l_day, l_leap, s_year, s_month, s_day):
        return self.hwp.LunarToSolar(lYear=l_year, lMonth=l_month, lDay=l_day, lLeap=l_leap,
                                     sYear=s_year, sMonth=s_month, sDay=s_day)

    def lunar_to_solar_by_set(self, l_year, l_month, l_day, l_leap):
        return self.hwp.LunarToSolarBySet(lYear=l_year, lMonth=l_month, lLeap=l_leap)

    def macro_state(self, macro_state):
        return self.hwp.MacroState(MacroState=macro_state)

    def mail_type(self, mail_type):
        return self.hwp.MailType(MailType=mail_type)

    def metatag_exist(self, tag):
        return self.hwp.MetatagExist(tag=tag)

    def mili_to_hwp_unit(self, mili):
        return self.hwp.MiliToHwpUnit(mili=mili)

    def modify_field_properties(self, field, remove, add):
        """
        지정한 필드의 속성을 바꾼다.
        양식모드에서 편집가능/불가 여부를 변경하는 메서드지만,
        현재 양식모드에서 어떤 속성이라도 편집가능하다..
        혹시 필드명이나 메모, 지시문을 수정하고 싶다면
        set_cur_field_name 메서드를 사용하자.

        :param field:
        :param remove:
        :param add:
        :return:
        """
        return self.hwp.ModifyFieldProperties(Field=field, remove=remove, Add=add)

    def modify_metatag_properties(self, tag, remove, add):
        return self.hwp.ModifyMetatagProperties(tag=tag, remove=remove, Add=add)

    def move_pos(self, move_id=1, para=0, pos=0):
        """
        캐럿의 위치를 옮긴다.
        move_id를 200(moveScrPos)으로 지정한 경우에는
        스크린 좌표로 마우스 커서의 (x,y)좌표를 그대로 넘겨주면 된다.
        201(moveScanPos)는 문서를 검색하는 중 캐럿을 이동시키려 할 경우에만 사용이 가능하다.
        (솔직히 200 사용법은 잘 모르겠다;)

        :param move_id:
            아래와 같은 값을 지정할 수 있다. 생략하면 1(moveCurList)이 지정된다.
            0: 루트 리스트의 특정 위치.(para pos로 위치 지정) moveMain
            1: 현재 리스트의 특정 위치.(para pos로 위치 지정) moveCurList
            2: 문서의 시작으로 이동. moveTopOfFile
            3: 문서의 끝으로 이동. moveBottomOfFile
            4: 현재 리스트의 시작으로 이동 moveTopOfList
            5: 현재 리스트의 끝으로 이동 moveBottomOfList
            6: 현재 위치한 문단의 시작으로 이동 moveStartOfPara
            7: 현재 위치한 문단의 끝으로 이동 moveEndOfPara
            8: 현재 위치한 단어의 시작으로 이동.(현재 리스트만을 대상으로 동작한다.) moveStartOfWord
            9: 현재 위치한 단어의 끝으로 이동.(현재 리스트만을 대상으로 동작한다.) moveEndOfWord
            10: 다음 문단의 시작으로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextPara
            11: 앞 문단의 끝으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevPara
            12: 한 글자 뒤로 이동.(서브 리스트를 옮겨 다닐 수 있다.) moveNextPos
            13: 한 글자 앞으로 이동.(서브 리스트를 옮겨 다닐 수 있다.) movePrevPos
            14: 한 글자 뒤로 이동.(서브 리스트를 옮겨 다닐 수 있다. 머리말/꼬리말, 각주/미주, 글상자 포함.) moveNextPosEx
            15: 한 글자 앞으로 이동.(서브 리스트를 옮겨 다닐 수 있다. 머리말/꼬리말, 각주/미주, 글상자 포함.) movePrevPosEx
            16: 한 글자 뒤로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextChar
            17: 한 글자 앞으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevChar
            18: 한 단어 뒤로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextWord
            19: 한 단어 앞으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevWord
            20: 한 줄 아래로 이동. moveNextLine
            21: 한 줄 위로 이동. movePrevLine
            22: 현재 위치한 줄의 시작으로 이동. moveStartOfLine
            23: 현재 위치한 줄의 끝으로 이동. moveEndOfLine
            24: 한 레벨 상위로 이동한다. moveParentList
            25: 탑레벨 리스트로 이동한다. moveTopLevelList
            26: 루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 반환한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다. moveRootList
            27: 현재 캐럿이 위치한 곳으로 이동한다. (캐럿 위치가 뷰의 맨 위쪽으로 올라간다.) moveCurrentCaret
            100: 현재 캐럿이 위치한 셀의 왼쪽 moveLeftOfCell
            101: 현재 캐럿이 위치한 셀의 오른쪽 moveRightOfCell
            102: 현재 캐럿이 위치한 셀의 위쪽 moveUpOfCell
            103: 현재 캐럿이 위치한 셀의 아래쪽 moveDownOfCell
            104: 현재 캐럿이 위치한 셀에서 행(row)의 시작 moveStartOfCell
            105: 현재 캐럿이 위치한 셀에서 행(row)의 끝 moveEndOfCell
            106: 현재 캐럿이 위치한 셀에서 열(column)의 시작 moveTopOfCell
            107: 현재 캐럿이 위치한 셀에서 열(column)의 끝 moveBottomOfCell
            200: 한/글 문서창에서의 screen 좌표로서 위치를 설정 한다. moveScrPos
            201: GetText() 실행 후 위치로 이동한다. moveScanPos

        :param para:
            이동할 문단의 번호.
            0(moveMain) 또는 1(moveCurList)가 지정되었을 때만 사용된다.
            200(moveScrPos)가 지정되었을 때는 문단번호가 아닌 스크린 좌표로 해석된다.
            (스크린 좌표 : LOWORD = x좌표, HIWORD = y좌표)

        :param pos:
            이동할 문단 중에서 문자의 위치.
            0(moveMain) 또는 1(moveCurList)가 지정되었을 때만 사용된다.

        :return:
            성공하면 True, 실패하면 False
        """
        return self.hwp.MovePos(moveID=move_id, Para=para, pos=pos)

    def move_to_field(self, field, text=True, start=True, select=False):
        """
        지정한 필드로 캐럿을 이동한다.

        :param field:
            필드이름. GetFieldText()/PutFieldText()와 같은 형식으로
            이름 뒤에 ‘{{#}}’로 번호를 지정할 수 있다.

        :param text:
            필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(True)
            누름틀 코드로 이동할지(False)를 지정한다.
            누름틀이 아닌 필드일 경우 무시된다. 생략하면 True가 지정된다.

        :param start:
            필드의 처음(True)으로 이동할지 끝(False)으로 이동할지 지정한다.
            select를 True로 지정하면 무시된다. 생략하면 True가 지정된다.

        :param select:
            필드 내용을 블록으로 선택할지(True), 캐럿만 이동할지(False) 지정한다.
            생략하면 False가 지정된다.
        :return:
        """
        return self.hwp.MoveToField(Field=field, Text=text, start=start, select=select)

    def move_to_metatag(self, tag, text, start, select):
        return self.hwp.MoveToMetatag(tag=tag, Text=text, start=start, select=select)

    def number_format(self, num_format):
        return self.hwp.NumberFormat(NumFormat=num_format)

    def numbering(self, numbering):
        return self.hwp.Numbering(Numbering=numbering)

    def open(self, filename, format="", arg=""):
        """
        문서를 연다.

        :param filename:
            문서 파일의 전체경로

        :param format:
            문서 형식. 빈 문자열을 지정하면 자동으로 인식한다. 생략하면 빈 문자열이 지정된다.
            "HWP": 한/글 native format
            "HWP30": 한/글 3.X/96/97
            "HTML": 인터넷 문서
            "TEXT": 아스키 텍스트 문서
            "UNICODE": 유니코드 텍스트 문서
            "HWP20": 한글 2.0
            "HWP21": 한글 2.1/2.5
            "HWP15": 한글 1.X
            "HWPML1X": HWPML 1.X 문서 (Open만 가능)
            "HWPML2X": HWPML 2.X 문서 (Open / SaveAs 가능)
            "RTF": 서식 있는 텍스트 문서
            "DBF": DBASE II/III 문서
            "HUNMIN": 훈민정음 3.0/2000
            "MSWORD": 마이크로소프트 워드 문서
            "DOCRTF": MS 워드 문서 (doc)
            "OOXML": MS 워드 문서 (docx)
            "HANA": 하나워드 문서
            "ARIRANG": 아리랑 문서
            "ICHITARO": 一太郞 문서 (일본 워드프로세서)
            "WPS": WPS 문서
            "DOCIMG": 인터넷 프레젠테이션 문서(SaveAs만 가능)
            "SWF": Macromedia Flash 문서(SaveAs만 가능)

        :param arg:
            세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.
            arg에 지정할 수 있는 옵션의 의미는 필터가 정의하기에 따라 다르지만,
            syntax는 다음과 같이 공통된 형식을 사용한다.
            "key:value;key:value;..."
            * key는 A-Z, a-z, 0-9, _ 로 구성된다.
            * value는 타입에 따라 다음과 같은 3 종류가 있다.
	        boolean: ex) fullsave:true (== fullsave)
	        integer: ex) type:20
	        string:  ex) prefix:_This_
            * value는 생략 가능하며, 이때는 콜론도 생략한다.
            * arg에 지정할 수 있는 옵션
            <모든 파일포맷>
                - setcurdir(boolean, true/false)
                    로드한 후 해당 파일이 존재하는 폴더로 현재 위치를 변경한다.
                    hyperlink 정보가 상대적인 위치로 되어 있을 때 유용하다.
            <HWP(HWPX)>
                - lock (boolean, TRUE)
                    로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
                - notext (boolean, FALSE)
                    텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)
                - template (boolean, FALSE)
                    새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다.
                    이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.
                - suspendpassword (boolean, FALSE)
                    TRUE로 지정하면 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.
                - forceopen (boolean, FALSE)
                    TRUE로 지정하면 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다.
                - versionwarning (boolean, FALSE)
                    TRUE로 지정하면 문서가 상위버전일 경우 메시지 박스를 띄우게 된다.
            <HTML>
                - code(string, codepage)
                    문서변환 시 사용되는 코드 페이지를 지정할 수 있으며 code키가 존재할 경우 필터사용 시 사용자 다이얼로그를  띄우지 않는다.
                - textunit(boolean, pixel)
                    Export될 Text의 크기의 단위 결정.pixel, point, mili 지정 가능.
                - formatunit(boolean, pixel)
                    Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능
                ※ [codepage 종류]
                    - ks :  한글 KS 완성형
                    - kssm: 한글 조합형
                    - sjis : 일본
                    - utf8 : UTF8
                    - unicode: 유니코드
                    - gb : 중국 간체
                    - big5 : 중국 번체
                    - acp : Active Codepage 현재 시스템의 코드 페이지
            <DOCIMG>
                - asimg(boolean, FALSE)
                    저장할 때 페이지를 image로 저장
                - ashtml(boolean, FALSE)
                    저장할 때 페이지를 html로 저장

        :return:
            성공하면 True, 실패하면 False
        """
        if not filename.lower().startswith("c:"):
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.Open(filename=filename, Format=format, arg=arg)

    def page_num_position(self, pagenumpos):
        return self.hwp.PageNumPosition(pagenumpos=pagenumpos)

    def page_type(self, page_type):
        return self.hwp.PageType(PageType=page_type)

    def para_head_align(self, para_head_align):
        return self.hwp.ParaHeadAlign(ParaHeadAlign=para_head_align)

    def pic_effect(self, pic_effect):
        return self.hwp.PicEffect(PicEffect=pic_effect)

    def placement_type(self, restart):
        return self.hwp.PlacementType(Restart=restart)

    def point_to_hwp_unit(self, point):
        return self.hwp.PointToHwpUnit(Point=point)

    def present_effect(self, prsnteffect):
        return self.hwp.PresentEffect(prsnteffect=prsnteffect)

    def print_device(self, print_device):
        return self.hwp.PrintDevice(PrintDevice=print_device)

    def print_paper(self, print_paper):
        return self.hwp.PrintPaper(PrintPaper=print_paper)

    def print_range(self, print_range):
        return self.hwp.PrintRange(PrintRange=print_range)

    def print_type(self, print_method):
        return self.hwp.PrintType(PrintMethod=print_method)

    def protect_private_info(self, protecting_char, private_pattern_type):
        """
        개인정보를 보호한다.
        한/글의 경우 “찾아서 보호”와 “선택 글자 보호”를 다른 기능으로 구현하였지만,
        API에서는 하나의 함수로 구현한다.

        :param protecting_char:
            보호문자. 개인정보는 해당문자로 가려진다.

        :param private_pattern_type:
            보호유형. 개인정보 유형마다 설정할 수 있는 값이 다르다.
            0값은 기본 보호유형으로 모든 개인정보를 보호문자로 보호한다.

        :return:
            개인정보를 보호문자로 치환한 경우에 true를 반환한다.
	        개인정보를 보호하지 못할 경우 false를 반환한다.
	        문자열이 선택되지 않은 상태이거나, 개체가 선택된 상태에서는 실패한다.
	        또한, 보호유형이 잘못된 설정된 경우에도 실패한다.
	        마지막으로 보호암호가 설정되지 않은 경우에도 실패하게 된다.
        """
        return self.hwp.ProtectPrivateInfo(PotectingChar=protecting_char, PrivatePatternType=private_pattern_type)

    def put_field_text(self, field, text: Union[str, list, tuple, pd.Series]):
        """
        지정한 필드의 내용을 채운다.
        현재 필드에 입력되어 있는 내용은 지워진다.
        채워진 내용의 글자모양은 필드에 지정해 놓은 글자모양을 따라간다.
        fieldlist의 필드 개수와, textlist의 텍스트 개수는 동일해야 한다.
        존재하지 않는 필드에 대해서는 무시한다.

        :param field:
            내용을 채울 필드 이름의 리스트.
            한 번에 여러 개의 필드를 지정할 수 있으며,
            형식은 GetFieldText와 동일하다.
            다만 필드 이름 뒤에 "{{#}}"로 번호를 지정하지 않으면
            해당 이름을 가진 모든 필드에 동일한 텍스트를 채워 넣는다.
            즉, PutFieldText에서는 ‘필드이름’과 ‘필드이름{{0}}’의 의미가 다르다.

        :param text:
            필드에 채워 넣을 문자열의 리스트.
            형식은 필드 리스트와 동일하게 필드의 개수만큼
            텍스트를 0x02로 구분하여 지정한다.

        :return: None

        :example:
            >>> # 현재 캐럿 위치에 zxcv 필드 생성
            >>> self.hwp.create_field("zxcv")
            >>> # zxcv 필드에 "Hello world!" 텍스트 삽입
            >>> self.hwp.put_field_text("zxcv", "Hello world!")
        """
        if type(field) in [list, tuple]:
            field = "\x02".join(field)
        if type(text) in [list, tuple, pd.Series]:
            text = "\x02".join(text)
        return self.hwp.PutFieldText(Field=field, Text=text)

    def put_metatag_name_text(self, tag, text):
        return self.hwp.PutMetatagNameText(tag=tag, Text=text)

    def quit(self):
        """
        한/글을 종료한다.
        단, 저장되지 않은 변경사항이 있는 경우 팝업이 뜨므로
        clear나 save 등의 메서드를 실행한 후에 quit을 실행해야 한다.
        :return:
        """
        self.hwp.Quit()
        del self.hwp

    def rgb_color(self, red, green, blue):
        return self.hwp.RGBColor(red=red, green=green, blue=blue)

    def register_module(self, module_type="FilePathCheckDLL", module_data="FilePathCheckerModule"):
        """
        (인스턴스 생성시 자동으로 실행된다.)
        한/글 컨트롤에 부가적인 모듈을 등록한다.
        사용자가 모르는 사이에 파일이 수정되거나 서버로 전송되는 것을 막기 위해
        한/글 오토메이션은 파일을 불러오거나 저장할 때 사용자로부터 승인을 받도록 되어있다.
        그러나 이미 검증받은 웹페이지이거나,
        이미 사용자의 파일 시스템에 대해 강력한 접근 권한을 갖는 응용프로그램의 경우에는
        이러한 승인절차가 아무런 의미가 없으며 오히려 불편하기만 하다.
        이런 경우 register_module을 통해 보안승인모듈을 등록하여 승인절차를 생략할 수 있다.

        :param module_type:
            모듈의 유형. 기본값은 "FilePathCheckDLL"이다.
            파일경로 승인모듈을 DLL 형태로 추가한다.

        :param module_data:
            Registry에 등록된 DLL 모듈 ID

        :return:
            추가모듈등록에 성공하면 True를, 실패하면 False를 반환한다.

        :example:
            >>> # 사전에 레지스트리에 보안모듈이 등록되어 있어야 한다.
            >>> # 보다 자세한 설명은 공식문서 참조
            >>> self.hwp.register_module("FilePathChekDLL", "FilePathCheckerModule")
            True
        """
        self.register_regedit()
        return self.hwp.RegisterModule(ModuleType=module_type, ModuleData=module_data)

    def register_regedit(self):
        import os
        import subprocess
        from winreg import ConnectRegistry, HKEY_CURRENT_USER, OpenKey, KEY_WRITE, SetValueEx, REG_SZ, CloseKey

        try:
            location = [i.split(": ")[1] for i in
                        subprocess.check_output(['pip', 'show', 'pyhwpx']).decode(encoding="cp949").split("\r\n") if
                        i.startswith("Location: ")][0]
        except:
            location = [i.split(": ")[1] for i in
                        subprocess.check_output(['pip', 'show', 'pyhwpx']).decode().split("\r\n") if
                        i.startswith("Location: ")][0]
        winup_path = r"Software\HNC\HwpAutomation\Modules"

        # HKEY_LOCAL_MACHINE와 연결 생성 후 핸들 얻음
        reg_handle = ConnectRegistry(None, HKEY_CURRENT_USER)

        # 얻은 행동을 사용해 WRITE 권한으로 레지스트리 키를 엶
        file_path_checker_module = winup_path + r"\FilePathCheckerModule"
        try:
            key = OpenKey(reg_handle, winup_path, 0, KEY_WRITE)
        except FileNotFoundError as e:
            winup_path = r"Software\Hnc\HwpUserAction\Modules"
            key = OpenKey(reg_handle, winup_path, 0, KEY_WRITE)
        SetValueEx(key, "FilePathCheckerModule", 0, REG_SZ, os.path.join(location, "FilePathCheckerModule.dll"))
        CloseKey(key)

    def register_private_info_pattern(self, private_type, private_pattern):
        """
        개인정보의 패턴을 등록한다.
        (현재 작동하지 않는다.)

        :param private_type:
            등록할 개인정보 유형. 다음의 값 중 하나다.
			0x0001: 전화번호
			0x0002: 주민등록번호
			0x0004: 외국인등록번호
			0x0008: 전자우편
			0x0010: 계좌번호
			0x0020: 신용카드번호
			0x0040: IP 주소
			0x0080: 생년월일
			0x0100: 주소
			0x0200: 사용자 정의

        :param private_pattern:
            등록할 개인정보 패턴. 예를 들면 이런 형태로 입력한다.
			(예) 주민등록번호 - "NNNNNN-NNNNNNN"
			한/글이 이미 정의한 패턴은 정의하면 안 된다.
			함수를 여러 번 호출하는 것을 피하기 위해 패턴을 “;”기호로 구분
			반속해서 입력할 수 있도록 한다.

        :return:
            등록이 성공하였으면 True, 실패하였으면 False

        :example:
            >>> self.hwp.RegisterPrivateInfoPattern(0x01, "NNNN-NNNN;NN-NN-NNNN-NNNN")  # 전화번호패턴
        """
        return self.hwp.RegisterPrivateInfoPattern(PrivateType=private_type, PrivatePattern=private_pattern)

    def release_action(self, action):
        return self.hwp.ReleaseAction(action=action)

    def release_scan(self):
        """
        InitScan()으로 설정된 초기화 정보를 해제한다.
        텍스트 검색작업이 끝나면 반드시 호출하여 설정된 정보를 해제해야 한다.

        :return: None
        """
        return self.hwp.ReleaseScan()

    def rename_field(self, oldname, newname):
        """
        지정한 필드의 이름을 바꾼다.
        예를 들어 oldname에 "title{{0}}\x02title{{1}}",
        newname에 "tt1\x02tt2로 지정하면 첫 번째 title은 tt1로, 두 번째 title은 tt2로 변경된다.
        oldname의 필드 개수와, newname의 필드 개수는 동일해야 한다.
        존재하지 않는 필드에 대해서는 무시한다.

        :param oldname:
            이름을 바꿀 필드 이름의 리스트. 형식은 PutFieldText와 동일하게 "\x02"로 구분한다.

        :param newname:
            새로운 필드 이름의 리스트. oldname과 동일한 개수의 필드 이름을 "\x02"로 구분하여 지정한다.

        :return: None

        :example:
            >>> self.hwp.create_field("asdf")  # "asdf" 필드 생성
            >>> self.hwp.rename_field("asdf", "zxcv")  # asdf 필드명을 "zxcv"로 변경
            >>> self.hwp.put_field_text("zxcv", "Hello world!")  # zxcv 필드에 텍스트 삽입
        """
        return self.hwp.RenameField(oldname=oldname, newname=newname)

    def rename_metatag(self, oldtag, newtag):
        return self.hwp.RenameMetatag(oldtag=oldtag, newtag=newtag)

    def replace_action(self, old_action_id, new_action_id):
        """
        특정 Action을 다른 Action으로 대체한다.
        이는 메뉴나 단축키로 호출되는 Action을 대체할 뿐,
        CreateAction()이나, Run() 등의 함수를 이용할 때에는 아무런 영향을 주지 않는다.
        즉, ReplaceAction(“Cut", "Copy")을 호출하여
        ”오려내기“ Action을 ”복사하기“ Action으로 교체하면
        Ctrl+X 단축키나 오려내기 메뉴/툴바 기능을 수행하더라도 복사하기 기능이 수행되지만,
        코드 상에서 Run("Cut")을 실행하면 오려내기 Action이 실행된다.
        또한, 대체된 Action을 원래의 Action으로 되돌리기 위해서는
        NewActionID의 값을 원래의 Action으로 설정한 뒤 호출한다. 이를테면 이런 식이다.
        >>> self.hwp.replace_action("Cut", "Cut")

        :param old_action_id:
            변경될 원본 Action ID.
            한/글 컨트롤에서 사용할 수 있는 Action ID는
            ActionTable.hwp(별도문서)를 참고한다.

        :param new_action_id:
            변경할 대체 Action ID.
            기존의 Action ID와 UserAction ID(ver:0x07050206) 모두 사용가능하다.

        :return:
            Action을 바꾸면 True를 바꾸지 못했다면 False를 반환한다.
        """

        return self.hwp.ReplaceAction(OldActionID=old_action_id, NewActionID=new_action_id)

    def replace_font(self, langid, des_font_name, des_font_type, new_font_name, new_font_type):
        return self.hwp.ReplaceFont(langid=langid, desFontName=des_font_name, desFontType=des_font_type,
                                    newFontName=new_font_name, newFontType=new_font_type)

    def revision(self, revision):
        return self.hwp.Revision(Revision=revision)

    def run(self, act_id):
        """
        액션을 실행한다. ActionTable.hwp 액션 리스트 중에서
        "별도의 파라미터가 필요하지 않은" 단순 액션을 run으로 호출할 수 있다.

        :param act_id:
            액션 ID (ActionIDTable.hwp 참조)

        :return:
            성공시 True, 실패시 False를 반환한다.
        """
        return self.hwp.HAction.Run(act_id)

    def AutoChangeHangul(self):
        """
        구버전의 "낱자모 우선입력" 활성화 토글기능. 현재는 사용하지 않으며, 최신버전에서 <도구-글자판-글자판 자동 변경(A)> 기능에 통합되었다.낱자모 우선입력 기능은 제거된 것으로 보임
        """
        return self.hwp.HAction.Run("AutoChangeHangul")

    def AutoChangeRun(self):
        """
        위 커맨드를 실행할 때마다 "글자판 자동 변경 기능"이 활성화/비활성화로 토글된다. 다만 API 등으로 텍스트를 입력하는 경우 원래 한/영 자동변환이 되지 않으므로, 자동화에는 쓰일 일이 없는 액션.
        """
        return self.hwp.HAction.Run("AutoChangeRun")

    def AutoSpellRun(self):
        """
        맞춤법 도우미(맞춤법이 틀린 단어 밑에 빨간 점선) 활성화/비활성화를 토글한다. 실행 후(비활성화시) 몇 초 뒤에 붉은 줄이 사라지는 것을 확인할 수 있다. 중간 스페이스에 유의.
        """
        return self.hwp.HAction.Run("AutoSpell Run")

    def AutoSpellSelect0(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect0")

    def AutoSpellSelect1(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect1")

    def AutoSpellSelect2(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect2")

    def AutoSpellSelect3(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect3")

    def AutoSpellSelect4(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect4")

    def AutoSpellSelect5(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect5")

    def AutoSpellSelect6(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect6")

    def AutoSpellSelect7(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect7")

    def AutoSpellSelect8(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect8")

    def AutoSpellSelect9(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect9")

    def AutoSpellSelect10(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect10")

    def AutoSpellSelect11(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect11")

    def AutoSpellSelect12(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect12")

    def AutoSpellSelect13(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect13")

    def AutoSpellSelect14(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect14")

    def AutoSpellSelect15(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect15")

    def AutoSpellSelect16(self):
        """
        맞춤법 도우미를 통해, 미리 입력되어 있는 어휘로 변경하는 액션. 어휘는 0에서부터 최대 16번 인덱스까지 존재. 예를 들어 아래 "aaple"이라는 오타의 경우 0~9까지 10개의 슬롯에 네 개의 어휘(ample, maple, apple, leap)만 랜덤하게 나타난다.
        """
        return self.hwp.HAction.Run("AutoSpellSelect16")

    def BreakColDef(self):
        """
        다단 레이아웃을 사용하는 경우의 "단 정의 삽입 액션(Ctrl-Alt-Enter)"이다. 아래 이미지의 중간페이지 참조. 단 정의 삽입 위치를 기점으로 구분된 다단을 하나 추가한다. 다단이 아닌 경우에는 일반 "문단나누기(Enter)"와 동일하다.
        """
        return self.hwp.HAction.Run("BreakColDef")

    def BreakColumn(self):
        """
        다단 레이아웃을 사용하는 경우 "단 나누기[배분다단] 액션(Ctrl-Shift-Enter)"이다. 아래 이미지의 중간페이지 참조. 단 정의 삽입 위치를 기점으로 구분된 다단을 하나 추가한다. 다단이 아닌 경우에는 일반 "문단나누기(Enter)"와 동일하다.
        """
        return self.hwp.HAction.Run("BreakColumn")

    def BreakLine(self):
        """
        라인나누기 액션(Shift-Enter). 들여쓰기나 내어쓰기 등 문단속성이 적용되어 있는 경우에 속성을 유지한 채로 줄넘김만 삽입한다. 이 단축키를 모르고 보고서를 작성하면, 들여쓰기를 맞추기 위해 스페이스를 여러 개 삽입했다가, 앞의 문구를 수정하는 과정에서 스페이스 뭉치가 문단 중간에 들어가버리는 대참사가 자주 발생할 수 있다.
        """
        return self.hwp.HAction.Run("BreakLine")

    def BreakPage(self):
        """
        쪽 나누기 액션(Ctrl-Enter). 캐럿 위치를 기준으로 하단의 글을 다음 페이지로 넘긴다. BreakLine과 마찬가지로 보고서 작성시 자주 사용해야 하는 액션으로, 이 기능을 사용하지 않고 보고서 작성시 엔터를 십여개 치고 다음 챕터 제목을 입력했다가, 일부 수정하면서 챕터 제목이 중간에 와 있는 경우 등의 불상사가 발생할 수 있다.
        """
        return self.hwp.HAction.Run("BreakPage")

    def BreakPara(self):
        """
        문단 나누기. 일반적인 엔터와 동일하다.
        """
        return self.hwp.HAction.Run("BreakPara")

    def BreakSection(self):
        """
        구역[섹션] 나누기 액션(Shift-Alt-Enter). 새로 생성된 섹션에서는 편집용지를 다르게 설정하거나, 혹은 새 개요번호/모양을 만든다든지 할 수 있다. 단, 초깃값으로 새 섹션이 생성되는 게 아니라, 기존 섹션의 편집용지, 바탕쪽 상태, 각주/미주 모양, 프레젠테이션 상태, 쪽 테두리/배경 속성 및 단 모양 등 대부분을 그대로 이어받으며, 새 섹션에서 수정시 기존 섹션에는 대부분 영향을 미치지 않는다.
        """
        return self.hwp.HAction.Run("BreakSection")

    def Cancel(self):
        """
        취소 액션. Esc 키를 눌렀을 때와 동일하다. 대표적인 예로는 텍스트 선택상태나 셀선택모드  해제 또는 이미지, 표 등의 개체 선택을 취소할 때 사용한다. Cancel 과 유사하게 쓰이는 액션으로 Close(또는 CloseEx, Shift-Esc)가 있다.
        """
        return self.hwp.HAction.Run("Cancel")

    def CaptureHandler(self):
        """
        갈무리 시작 액션. 현재 버전에서는 Run커맨드로 사용할 수 없는 것 같다. 어떤 이미지포맷으로 저장하든 오류발생.
        """
        return self.hwp.HAction.Run("CaptureHandler")

    def CaptureDialog(self):
        """
        갈무리 끝 액션. 현재 버전에서는 Run커맨드로 사용할 수 없는 것 같다. 어떤 이미지포맷으로 저장하든 오류발생.
        """
        return self.hwp.HAction.Run("CaptureDialog")

    def CharShapeBold(self):
        """
        글자모양 중 "진하게Bold" 속성을 토글하는 액션. 이 액션을 실행하기 전에 특정 셀이나 텍스트가 선택된 상태여야 하며, 이 커맨드만으로는 확실히 "진하게" 속성이 적용되었는지 확인할 수 없다. 그 이유는 토글 커맨드라서, 기존에 진하게 적용되어 있었다면, 해제되어버리기 때문이다. 확실히 진하게를 적용하는 방법으로는, 초기에 모든 텍스트의 진하게를 해제(진하게 두 번??)한다든지, 파라미터셋을 활용하여 진하게 속성이 적용되어 있는지를 확인하는 방법 등이 있다.
        """
        return self.hwp.HAction.Run("CharShapeBold")

    def CharShapeCenterline(self):
        """
        글자에 취소선 적용을 토글하는 액션. Bold와 마찬가지로 토글이므로, 기존에 취소선이 적용되어 있다면 해제되어버리므로 사용에 유의해야 한다.
        """
        return self.hwp.HAction.Run("CharShapeCenterline")

    def CharShapeEmboss(self):
        """
        글자모양에 양각 속성(글자가 튀어나온 느낌) 적용을 토글한다.
        """
        return self.hwp.HAction.Run("CharShapeEmboss")

    def CharShapeEngrave(self):
        """
        글자모양에 음각 속성(글자가 움푹 들어간 느낌) 적용을 토글한다.
        """
        return self.hwp.HAction.Run("CharShapeEngrave")

    def CharShapeHeight(self):
        """
        글자모양(Alt-L) 대화상자를 열고, 포커스를 "기준 크기"로 이동한다. (수작업이 필요하므로 자동화에 사용하지는 않는다. 유사한 액션으로 글꼴 언어를 선택하는 CharShapeLang, CharShapeSpacing, CharShapeTypeFace, CharShapeWidth 등이 있다.)
        """
        return self.hwp.HAction.Run("CharShapeHeight")

    def CharShapeHeightDecrease(self):
        """
        글자 크기를 1포인트씩 작게 한다. 단, 속도가 다소 느리므로.. 큰 폭으로 조정할 때에는 다른 방법을 쓰는 것을 추천.
        """
        return self.hwp.HAction.Run("CharShapeHeightDecrease")

    def CharShapeHeightIncrease(self):
        """
        글자 크기를 1포인트씩 크게 한다. 단, 속도가 다소 느리므로.. 큰 폭으로 조정할 때에는 다른 방법을 쓰는 것을 추천.
        """
        return self.hwp.HAction.Run("CharShapeHeightIncrease")

    def CharShapeItalic(self):
        """
        글자 모양에 이탤릭 속성을 토글한다.
        """
        return self.hwp.HAction.Run("CharShapeItalic")

    def CharShapeNextFaceName(self):
        """
        다음 글꼴로 이동(Shift-Alt-F)한다. 단, 이 액션으로 어떤 폰트가 선택되었는지를 파이썬에서 확인하려면 파라미터셋에 접근해야 한다. 유사한 커맨드로, CharShapePrevFaceName 이 있다.
        """
        return self.hwp.HAction.Run("CharShapeNextFaceName")

    def CharShapeNormal(self):
        """
        글자모양에 적용된 속성 및 글자색 등 전부를 해제(Shift-Alt-C)한다. 단 글꼴, 크기 등은 바뀌지 않는다.
        """
        return self.hwp.HAction.Run("CharShapeNormal")

    def CharShapeOutline(self):
        """
        글자모양의 외곽선 속성을 토글한다.
        """
        return self.hwp.HAction.Run("CharShapeOutline")

    def CharShapeShadow(self):
        """
        선택한 텍스트 글자모양 중 그림자 속성을 토글한다.
        """
        return self.hwp.HAction.Run("CharShapeShadow")

    def CharShapeSpacingDecrease(self):
        """
        자간을 1%씩 좁힌다. 최대 -50%까지 좁힐 수 있다. 다만 자동화 작업시 줄넘김을 체크하는 것이 상당히 번거로운 작업이므로, 크게 보고서의 틀이 바뀌지 않는 선에서는 자간을 좁히는 것보다 "한 줄로 입력"을 활용하는 편이 간단하고 자연스러울 수 있다. 한 줄로 입력 옵션 : 문단모양(Alt-T)의 확장 탭에 있음. 한 줄로 입력을 활성화해놓은 문단이나 셀에서는 자간이 아래와 같이 자동으로 좁혀진다.
        """
        return self.hwp.HAction.Run("CharShapeSpacingDecrease")

    def CharShapeSpacingIncrease(self):
        """
        자간을 1%씩 넓힌다. 최대 50%까지 넓힐 수 있다.
        """
        return self.hwp.HAction.Run("CharShapeSpacingIncrease")

    def CharShapeSubscript(self):
        """
        선택한 텍스트에 아래첨자 속성을 토글(Shift-Alt-S)한다.
        """
        return self.hwp.HAction.Run("CharShapeSubscript")

    def CharShapeSuperscript(self):
        """
        선택한 텍스트에 위첨자 속성을 토글(Shift-Alt-P)한다.
        """
        return self.hwp.HAction.Run("CharShapeSuperscript")

    def CharShapeSuperSubscript(self):
        """
        선택한 텍스트의 첨자속성을 위→아래→보통의 순서를 반복해서 토글한다.
        """
        return self.hwp.HAction.Run("CharShapeSuperSubscript")

    def CharShapeTextColorBlack(self):
        """
        선택한 텍스트의 글자색을 검정색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorBlack")

    def CharShapeTextColorBlue(self):
        """
        선택한 텍스트의 글자색을 파란색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorBlue")

    def CharShapeTextColorBluish(self):
        """
        선택한 텍스트의 글자색을 청록색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorBluish")

    def CharShapeTextColorGreen(self):
        """
        선택한 텍스트의 글자색을 초록색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorGreen")

    def CharShapeTextColorRed(self):
        """
        선택한 텍스트의 글자색을 빨간색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorRed")

    def CharShapeTextColorViolet(self):
        """
        선택한 텍스트의 글자색을 보라색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorViolet")

    def CharShapeTextColorWhite(self):
        """
        선택한 텍스트의 글자색을 흰색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorWhite")

    def CharShapeTextColorYellow(self):
        """
        선택한 텍스트의 글자색을 노란색으로 변경한다.
        """
        return self.hwp.HAction.Run("CharShapeTextColorYellow")

    def CharShapeUnderline(self):
        """
        선택한 텍스트에 밑줄 속성을 토글한다. 대소문자에 유의해야 한다. (UnderLine이 아니다.)
        """
        return self.hwp.HAction.Run("CharShapeUnderline")

    def CharShapeWidthDecrease(self):
        """
        장평을 1%씩 줄인다. 장평 범위는 50~200%이며, 장평을 늘일 때는 Decrease 대신 Increase를 사용하면 된다.
        """
        return self.hwp.HAction.Run("CharShapeWidthDecrease")

    def CharShapeWidthIncrease(self):
        """
        장평을 1%씩 줄인다. 장평 범위는 50~200%이며, 장평을 줄일 때는 Increase 대신 Decrease를 사용하면 된다.
        """
        return self.hwp.HAction.Run("CharShapeWidthIncrease")

    def Close(self):
        """
        현재 리스트를 닫고 (최)상위 리스트로 이동하는 액션. 대표적인 예로, 메모나 각주 등을 작성한 후 본문으로 빠져나올 때, 혹은 여러 겹의 표 안에 있을 때 한 번에 표 밖으로 캐럿을 옮길 때 사용한다. 굉장히 자주 쓰이는 액션이며, 경우에 따라 Close가 아니라 CloseEx를 써야 하는 경우도 있다. 아래 영상의 캐럿 위치에 주목.
        """
        return self.hwp.HAction.Run("Close")

    def CloseEx(self):
        """
        현재 리스트를 닫고 상위 리스트로 이동하는 액션. Close와 유사하나 두 가지 차이점이 있다. 첫 번째로는 여러 계층의 표 안에서 CloseEx 실행시 본문이 아니라 상위의 표(셀)로 캐럿이 이동한다는 점. Close는 무조건 본문으로 나간다. 두 번째로, CloseEx에는 전체화면(최대화 말고)을 해제하는 기능이 있다. Close로는 전체화면 해제가 되지 않는다. 사용빈도가 가장 높은 액션 중의 하나라고 생각한다.
        """
        return self.hwp.HAction.Run("CloseEx")

    def Comment(self):
        """
        아래아한글에 "숨은 설명"이 있다는 걸 아는 사람도 없다시피 한데, 그 "숨은 설명" 관련한 Run 액션이 세 개나 있다. Comment 액션은 표현 그대로 숨은 설명을 붙일 수 있다. 텍스트만 넣을 수 있을 것 같은 액션이름인데, 사실 표나 그림도 자유롭게 삽입할 수 있기 때문에, 문서 안에 몰래 숨겨놓은 또다른 문서 느낌이다. 파일별로 자동화에 활용할 수 있는 특정 문자열을 파이썬이 아니라 숨은설명 안에 붙여놓고 활용할 수도 있지 않을까 이런저런 고민을 해봤는데, 개인적으로 자동화에 제대로 활용한 적은 한 번도 없었다. 숨은 설명이라고 민감한 정보를 넣으면 안 되는데, 완전히 숨겨져 있는 게 아니기 때문이다. 현재 캐럿위치에 [숨은설명] 조판부호가 삽입되며, 이를 통해 숨은 설명 내용이 확인 가능하므로 유념해야 한다. 재미있는 점은, 숨은설명 안에 또 숨은설명을 삽입할 수 있다. 숨은설명 안에다 숨은설명을 넣고 그 안에 또 숨은설명을 넣는... 이런 테스트를 해봤는데 2,400단계 정도에서 한글이 종료돼버렸다.
        """
        return self.hwp.HAction.Run("Comment")

    def CommentDelete(self):
        """
        단어 그대로 숨은 설명을 지우는 액션이다. 단, 사용방법이 까다로운데 숨은 설명 안에 들어가서 CommentDelete를 실행하면, 지울지 말지(Yes/No) 팝업이 나타난다. 나중에 자세히 설명하겠지만 이런 팝업을 자동처리하는 방법은 hwp__.SetMessageBoxMode() 메서드를 미리 실행해놓는 것이다. Yes/No 방식의 팝업에서 Yes를 선택하는 파라미터는 0x10000 (또는 65536)이므로, hwp__.SetMessageBoxMode(0x10000) 를 사용하면 된다.
        """
        return self.hwp.HAction.Run("CommentDelete")

    def CommentModify(self):
        """
        단어 그대로 숨은 설명을 수정하는 액션이다. 캐럿은 [숨은설명] 조판부호 바로 앞에 위치하고 있어야 한다.
        """
        return self.hwp.HAction.Run("CommentModify")

    def Copy(self):
        """
        복사하기. 선택되어 있는 문자열 혹은 개체(표, 이미지 등)를 클립보드에 저장한다. 파이썬에서 클립보드를 다루는 모듈은 pyperclip이나, pywin32의 win32clipboard 두 가지가 가장 많이 쓰이는데, 단순한 문자열의 경우 아래처럼
        """
        return self.hwp.HAction.Run("Copy")

    def CopyPage(self):
        """
        쪽 복사
        """
        return self.hwp.HAction.Run("CopyPage")

    def Cut(self):
        """
        잘라내기. Copy 액션과 유사하지만, 복사 대신 잘라내기 기능을 수행한다. 자주 쓰이는 메서드이다.
        """
        return self.hwp.HAction.Run("Cut")

    def Delete(self):
        """
        삭제액션. 키보드의 Del 키를 눌렀을 때와 대부분 유사하다. 아주 사용빈도가 높은 액션이다.
        """
        return self.hwp.HAction.Run("Delete")

    def DeleteBack(self):
        """
        Delete와 유사하지만, 이건 Backspace처럼 우측에서 좌측으로 삭제해준다. 많이 쓰인다.
        """
        return self.hwp.HAction.Run("DeleteBack")

    def DeleteField(self):
        """
        누름틀지우기. 누름틀 안의 내용은 지우지 않고, 단순히 누름틀만 지운다. 지울 때 캐럿의 위치는 누름틀 안이든, 앞이나 뒤든 붙어있기만 하면 된다. 만약 최종문서에는 누름틀을 넣지 않고 모두 일반 텍스트로 변환하려고 하면 이 기능을 활용할 수 있다.
        """
        return self.hwp.HAction.Run("DeleteField")

    def DeleteFieldMemo(self):
        """
        메모 지우기. 누름틀 지우기와 유사하다. 메모 누름틀에 붙어있거나, 메모 안에 들어가 있는 경우 위 액션 실행시 해당 메모가 삭제된다.
        """
        return self.hwp.HAction.Run("DeleteFieldMemo")

    def DeleteLine(self):
        """
        한 줄 지우기(Ctrl-Y) 액션. 문단나눔과 전혀 상관없이 딱 한 줄의 텍스트가 삭제된다. DeleteLine으로 표 등의 객체를 삭제하는 경우에는 팝업이 뜨므로 유의해야 한다. (hwp.SetMessageBoxMode 메서드를 추가로 사용하면 해결된다.)
        """
        return self.hwp.HAction.Run("DeleteLine")

    def DeleteLineEnd(self):
        """
        현재 커서에서 줄 끝까지 지우기(Alt-Y). 수작업시에 굉장히 유용한 기능일 수 있지만, 자동화 작업시에는 DeleteLine이나 DeleteLineEnd 모두, 한 줄 안에 어떤 내용까지 있는지 파악하기 어려운 관계로, 자동화에 잘 쓰이지는 않는다.
        """
        return self.hwp.HAction.Run("DeleteLineEnd")

    def DeletePage(self):
        """
        쪽 지우기
        """
        return self.hwp.HAction.Run("DeletePage")

    def DeleteWord(self):
        """
        단어 지우기(Ctrl-T) 액션. 단, 커서 우측에 위치한 단어 한 개씩 삭제하며, 커서가 단어 중간에 있는 경우 우측 글자만 삭제한다.
        """
        return self.hwp.HAction.Run("DeleteWord")

    def DeleteWordBack(self):
        """
        한 단어씩 좌측으로 삭제하는 액션(Ctrl-백스페이스). DeleteWord와 마찬가지로 커서가 단어 중간에 있는 경우 좌측 글자만 삭제한다.
        """
        return self.hwp.HAction.Run("DeleteWordBack")

    def DrawObjCancelOneStep(self):
        """
        다각형(곡선) 그리는 중 이전 선 지우기. 현재 사용 안함(?)
        """
        return self.hwp.HAction.Run("DrawObjCancelOneStep")

    def DrawObjEditDetail(self):
        """
        그리기 개체 중 다각형 점편집 액션. 다각형이 선택된 상태에서만 실행가능.
        """
        return self.hwp.HAction.Run("DrawObjEditDetail")

    def DrawObjOpenClosePolygon(self):
        """
        닫힌 다각형 열기 또는 열린 다각형 닫기 토글.①다각형 개체 선택상태가 아니라 편집상태에서만 위 명령어가 실행된다.②닫힌다각형을 열 때는 마지막으로 봉합된 점에서 아주 조금만 열린다.③아주 조금만 열린 상태에서 닫으면 노드(꼭지점)가 추가되지 않지만, 적절한 거리를 벌리고 닫기를 하면 추가됨.
        """
        return self.hwp.HAction.Run("DrawObjOpenClosePolygon")

    def DrawObjTemplateSave(self):
        """
        그리기개체를 그리기마당에 템플릿으로 등록하는 액션(어떻게 써먹고 싶어도 방법을 모르겠다...)그리기개체가 선택된 상태에서만 실행 가능하다.여담으로, 그리기 마당에 임의로 등록한 개체 삭제 아이콘을 못 찾고 있는데; 한글2020 기준으로, 개체 이름을 "얼굴"이라고 "기본도형"에 저장했을 경우, 찾아가서 아래의 파일을 삭제해도 된다."C:/Users/이름/AppData/Roaming/HNC/User/Shared110/HwpTemplate/Draw/FG_Basic_Shapes/얼굴.drt"
        """
        return self.hwp.HAction.Run("DrawObjTemplateSave")

    def EditFieldMemo(self):
        """
        메모 내용 편집 액션. "메모 내용 보기" 창이 하단에 열린다. SplitMemoOpen과 동일한 기능으로 보이며, 메모내용보기창에서 두 번째 이후의 메모 클릭시 메모내용보기창이 닫히는 버그가 있다.(한/글 2020 기준)참고로 메모내용 보기 창을 닫을 때는 SplitMemoClose 커맨드를 쓰면 된다.
        """
        return self.hwp.HAction.Run("EditFieldMemo")

    def Erase(self):
        """
        선택한 문자나 개체 삭제. 문자열이나 컨트롤 등을 삭제한다는 점에서는 Delete나 DeleteBack과 유사하지만, 가장 큰 차이점은, 아무 것도 선택되어 있지 않은 상태일 때 Erase는 아무 것도 지우지 않는다는 점이다. (Delete나 DeleteBack은 어찌됐든 앞뒤의 뭔가를 지운다.)
        """
        return self.hwp.HAction.Run("Erase")

    def FileClose(self):
        """
        문서 닫기. 한/글을 종료하는 명령어는 아니다. 다만 문서저장 이후 수정을 한 상태이거나, 빈 문서를 열어서 편집한 경우에는, 팝업이 나타나고 사용자 입력을 요구하므로 자동화작업에 걸림돌이 된다.이를 해결하는 세 가지(?) 옵션이 있는데,①문서를 저장한 후 FileClose 실행저장하는 방법은, hwp__.SaveAs(Path)②변경된 내용을 버린 후 FileClose 실행(탬플릿문서를 쓰고 있거나, 이미 PDF로 저장했다든지, 캡쳐를 완료한 경우 등)버리는 방법은 hwp__.Clear(option=1)※ Clear 메서드는 경우에 따라 심각한 오류를 뱉기도 한다. 그것도 상당히 빈도가 잦아서 필자는 Clear를 사용하지 않는 편이다. 대신 아래의 XHwpDocument.Close(False)를 사용하는 편.③변경된 내용을 버리고 문서를 닫는 명령 실행hwp.XHwpDocuments.Item(0).Close(isDirty=False)위 명령어는 다소 길어 보이지만 hwp__.Clear(option=1), hwp__.HAction.Run("FileClose")와 동일하게 작동한다.
        """
        return self.hwp.HAction.Run("FileClose")

    def FileNew(self):
        """
        새 문서 창을 여는 명령어. 참고로 현재 창에서 새 탭을 여는 명령어는 hwp__.HAction.Run("FileNewTab"). 여담이지만 한/글2020 기준으로 새 창은 30개까지 열 수 있다. 그리고 한 창에는 탭을 30개까지 열 수 있다. 즉, (리소스만 충분하다면) 동시에 열어서 자동화를 돌릴 수 있는 문서 갯수는 900개.
        """
        return self.hwp.HAction.Run("FileNew")

    def FileOpen(self):
        """
        문서를 여는 명령어. 단 파일선택 팝업이 뜨므로, 자동화작업시에는 이 명령어를 사용하지 않는다.  대신 hwp__.Open(파일명)을 사용해야 한다. 레지스트리에디터에 보안모듈 등록(링크)을 해놓으면 hwp__.Open 명령 실행시에 보안팝업도 뜨지 않는다.
        """
        return self.hwp.HAction.Run("FileOpen")

    def FileOpenMRU(self):
        """
        API매뉴얼엔 "최근 작업 문서"를 여는 명령어라고 나와 있지만, 현재는 FileOpen과 동일한 동작을 하는 것으로 보인다. 이 액션 역시 사용자입력을 요구하는 팝업이 뜨므로 자동화에 사용하지 않으며, hwp__.Open(Path)을 써야 한다.
        """
        return self.hwp.HAction.Run("FileOpenMRU")

    def FilePreview(self):
        """
        미리보기 창을 열어준다. 자동화와 큰 연관이 없어 자주 쓰이지도 않고, 더군다나 닫는 명령어가 없다.또한 이 명령어는 hwp__.XHwpDocuments.Item(0).XHwpPrint.RunFilePreview()와 동일한 동작을 하는데,재미있는 점은,①스크립트 매크로 녹화 진행중에 hwp__.HAction.Run("FilePreview")는 실행해도 반응이 없고, 녹화 로그에도 잡히지 않는다.②그리고 스크립트매크로 녹화 진행중에 [파일] - [미리보기(V)] 메뉴도 비활성화되어 있어 코드를 알 수 없다.③그런데 hwp__.XHwpDocuments.Item(0).XHwpPrint.RunFilePreview()는 녹화중에도 실행이 된다.녹화된 코드와 관련하여 남기고 싶은 코멘트가 많은데, 별도의 포스팅으로 남길 예정.
        """
        return self.hwp.HAction.Run("FilePreview")

    def FileQuit(self):
        """
        한/글 프로그램을 종료한다. 단, 저장 이후 문서수정이 있는 경우에는 팝업이 뜨므로, ①저장하거나 ②수정내용을 버리는 메서드를 활용해야 한다.
        """
        return self.hwp.HAction.Run("FileQuit")

    def FileSave(self):
        """
        파일을 저장하는 액션(Alt-S). 자동화프로세스 중 빈 문서를 열어 작성하는 경우에는, 저장액션 실행시 아래와 같이 경로선택 팝업이 뜨므로, hwp__.SaveAs(Path) 메서드를 사용하여 저장한 후 Run("FileSave")를 써야 한다.Run("FileSave")는 hwp__.Save() 메서드와 거의 동일하지만 한 가지 차이점이 있는데,- hwp__.Save()는 수정사항이 있는 경우에만 저장 프로세스를 실행하여 부하를 줄이는데 반해 hwp__.HAction.Run("FileSave")는 매번 실행할 때마다 변동사항이 없더라도 저장 프로세스를 실행한다.단, hwp__.Save(save_if_dirty=False) 방식으로 파라미터를 주고 실행하면 Run("FileSave")와 동일하게, 수정이 없더라도 매번 저장을 수행하게 된다.
        """
        return self.hwp.HAction.Run("FileSave")

    def FileSaveAs(self):
        """
        다른 이름으로 저장(Alt-V). 사용자입력을 필요로 하므로 이 액션은 사용하지 않는다.대신 hwp__.SaveAs(Path)를 사용하면 된다.
        """
        return self.hwp.HAction.Run("FileSaveAs")

    def FindForeBackBookmark(self):
        """
        책갈피 찾아가기. 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackBookmark")

    def FindForeBackCtrl(self):
        """
        조판부호 찾아가기. FindForeBackBookmark와 마찬가지로 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackCtrl")

    def FindForeBackFind(self):
        """
        찾기. FindForeBackBookmark와 마찬가지로 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackFind")

    def FindForeBackLine(self):
        """
        줄 찾아가기. FindForeBackBookmark와 마찬가지로 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackLine")

    def FindForeBackPage(self):
        """
        쪽 찾아가기. FindForeBackBookmark와 마찬가지로 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackPage")

    def FindForeBackSection(self):
        """
        구역 찾아가기. FindForeBackBookmark와 마찬가지로 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackSection")

    def FindForeBackStyle(self):
        """
        스타일 찾아가기. FindForeBackBookmark와 마찬가지로 사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        """
        return self.hwp.HAction.Run("FindForeBackStyle")

    def FrameStatusBar(self):
        """
        한/글 프로그램 하단의 상태바 보이기/숨기기 토글
        """
        return self.hwp.HAction.Run("FrameStatusBar")

    def HanThDIC(self):
        """
        한/글에 내장되어 있는 "유의어/반의어 사전"을 여는 액션.
        """
        return self.hwp.HAction.Run("HanThDIC")

    def HeaderFooterDelete(self):
        """
        머리말/꼬리말 지우기. 본문이 아니라 머리말/꼬리말 편집상태에서 실행해야 삭제 팝업이 뜬다.삭제팝업 없이 머리말/꼬리말을 삭제하려면 hwp__.SetMessageBoxMode(0x10000)을 미리 실행해놓아야 한다.참고로 아래 영상에서는 마우스 더블클릭을 했지만, 자동화작업시에는 아래의 Run("HeaderFooterModify")을 통해 편집상태로 들어가야 한다.
        """
        return self.hwp.HAction.Run("HeaderFooterDelete")

    def HeaderFooterModify(self):
        """
        머리말/꼬리말 고치기. 마우스를 쓰지 않고 머리말/꼬리말 편집상태로 들어갈 수 있다. 단, 커서가 조판부호에 닿아 있는 상태에서 실행해야 한다.
        """
        return self.hwp.HAction.Run("HeaderFooterModify")

    def HeaderFooterToNext(self):
        """
        다음 머리말/꼬리말. 당장은 사용방법을 모르겠다..
        """
        return self.hwp.HAction.Run("HeaderFooterToNext")

    def HeaderFooterToPrev(self):
        """
        이전 머리말. 당장은 사용방법을 모르겠다..
        """
        return self.hwp.HAction.Run("HeaderFooterToPrev")

    def HiddenCredits(self):
        """
        인터넷 정보. 사용방법을 모르겠다.
        """
        return self.hwp.HAction.Run("HiddenCredits")

    def HideTitle(self):
        """
        차례 숨기기([도구 - 차례/색인 - 차례 숨기기] 메뉴에 대응(Ctrl-K-S). 실행한 개요라인을 자동생성되는 제목차례에서 숨긴다. 즉시 변경되지 않으며, "모든 차례 새로고침(Ctrl-K-A)" 실행시 제목차례가 업데이트된다.모든차례 새로고침 명령어는 hwp__.HAction.Run("UpdateAllContents") 이다.적용여부는 Ctrl+G,C를 이용해 조판부호를 확인하면 알 수 있다.
        """
        return self.hwp.HAction.Run("HideTitle")

    def HimConfig(self):
        """
        입력기 언어별 환경설정. 현재는 실행되지 않는 듯 하다. 대신 Run("HimKbdChange")로 환경설정창을 띄울 수 있다.자동화에는 쓰이지 않는다.
        """
        return self.hwp.HAction.Run("Him Config")

    def HimKbdChange(self):
        """
        입력기 언어별 환경설정.
        """
        return self.hwp.HAction.Run("HimKbdChange")

    def HwpCtrlEquationCreate97(self):
        """
        "한/글97버전 수식 만들기"라고 하는데, 실행되지 않는 듯 하다.
        """
        return self.hwp.HAction.Run("HwpCtrlEquationCreate97")

    def HwpCtrlFileNew(self):
        """
        한글컨트롤 전용 새문서. 실행되지 않는 듯 하다.
        """
        return self.hwp.HAction.Run("HwpCtrlFileNew")

    def HwpCtrlFileOpen(self):
        """
        한글컨트롤 전용 파일 열기. 실행되지 않는 듯 하다.
        """
        return self.hwp.HAction.Run("HwpCtrlFileOpen")

    def HwpCtrlFileSave(self):
        """
        한글컨트롤 전용 파일 저장. 실행되지 않는다.
        """
        return self.hwp.HAction.Run("HwpCtrlFileSave")

    def HwpCtrlFileSaveAs(self):
        """
        한글컨트롤 전용 다른 이름으로 저장. 실행되지 않는다.
        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAs")

    def HwpCtrlFileSaveAsAutoBlock(self):
        """
        한글컨트롤 전용 다른이름으로 블록 저장. 실행되지 않는다.
        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAsAutoBlock")

    def HwpCtrlFileSaveAutoBlock(self):
        """
        한/글 컨트롤 전용 블록 저장. 실행되지 않는다.
        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAutoBlock")

    def HwpCtrlFindDlg(self):
        """
        한/글 컨트롤 전용 찾기 대화상자. 실행되지 않는다.
        """
        return self.hwp.HAction.Run("HwpCtrlFindDlg")

    def HwpCtrlReplaceDlg(self):
        """
        한/글 컨트롤 전용 바꾸기 대화상자
        """
        return self.hwp.HAction.Run("HwpCtrlReplaceDlg")

    def HwpDic(self):
        """
        한컴 사전(F12). 현재 캐럿이 닿아 있거나, 블록선택한 구간을 검색어에 자동으로 넣는다.
        """
        return self.hwp.HAction.Run("HwpDic")

    def HyperlinkBackward(self):
        """
        하이퍼링크 뒤로. 하이퍼링크를 통해서 문서를 탐색하여 페이지나 캐럿을 이동한 경우, (브라우저의 "뒤로가기"처럼) 이동 전의 위치로 돌아간다.
        """
        return self.hwp.HAction.Run("HyperlinkBackward")

    def HyperlinkForward(self):
        """
        하이퍼링크 앞으로. Run("HyperlinkBackward") 에 상반되는 명령어로, 브라우저의 "앞으로 가기"나 한/글의 재실행과 유사하다. 하이퍼링크 등으로 이동한 후에 뒤로가기를 눌렀다면, 캐럿이 뒤로가기 전 위치로 다시 이동한다.
        """
        return self.hwp.HAction.Run("HyperlinkForward")

    def ImageFindPath(self):
        """
        그림 경로 찾기. 현재는 실행되지 않는 듯.
        """
        return self.hwp.HAction.Run("ImageFindPath")

    def InputCodeChange(self):
        """
        문자/코드 변환.. 현재 캐럿의 바로 앞 문자를 찾아서 문자이면 코드로, 코드이면 문자로 변환해준다.(변환 가능한 코드영역 0x0020 ~ 0x10FFFF 까지)
        """
        return self.hwp.HAction.Run("InputCodeChange")

    def InputHanja(self):
        """
        한자로 바꾸기 창을 띄워준다. 추가입력이 필요하여 자동화에는 쓰이지 않음.
        """
        return self.hwp.HAction.Run("InputHanja")

    def InputHanjaBusu(self):
        """
        부수로 입력. 자동화에는 쓰이지 않음.
        """
        return self.hwp.HAction.Run("InputHanjaBusu")

    def InputHanjaMean(self):
        """
        한자 새김 입력창 띄우기. 뜻과 음을 입력하면 적절한 한자를 삽입해준다.입력시 뜻과 음은 붙여서 입력. (예)하늘천
        """
        return self.hwp.HAction.Run("InputHanjaMean")

    def InsertAutoNum(self):
        """
        번호 다시 넣기(?) 실행이 안되는 듯.
        """
        return self.hwp.HAction.Run("InsertAutoNum")

    def InsertCpNo(self):
        """
        현재 쪽번호(상용구) 삽입. 쪽번호와 마찬가지로, 문자열이 실시간으로 변경된다.※유의사항 : 이 쪽번호는 찾기, 찾아바꾸기, GetText 및 누름틀 안에 넣고 GetFieldText나 복붙 등 그 어떤 방법으로도 추출되지 않는다.한 마디로 눈에는 보이는 것 같지만 실재하지 않는 숫자임. 참고로 표번호도 그렇다. 값이 아니라 속성이라서 그렇다.
        """
        return self.hwp.HAction.Run("InsertCpNo")

    def InsertCpTpNo(self):
        """
        상용구 코드 넣기(현재 쪽/전체 쪽). 실시간으로 변경된다.
        """
        return self.hwp.HAction.Run("InsertCpTpNo")

    def InsertDateCode(self):
        """
        상용구 코드 넣기(만든 날짜). 현재날짜가 아님에 유의.
        """
        return self.hwp.HAction.Run("InsertDateCode")

    def InsertDocInfo(self):
        """
        상용구 코드 넣기(만든 사람, 현재 쪽, 만든 날짜)
        """
        return self.hwp.HAction.Run("InsertDocInfo")

    def InsertEndnote(self):
        """
        미주 입력
        """
        return self.hwp.HAction.Run("InsertEndnote")

    def InsertFieldDateTime(self):
        """
        날짜/시간 코드로 넣기([입력-날짜/시간-날짜/시간 코드]메뉴와 동일)
        """
        return self.hwp.HAction.Run("InsertFieldDateTime")

    def InsertFieldMemo(self):
        """
        메모 넣기([입력-메모-메모 넣기]메뉴와 동일)
        """
        return self.hwp.HAction.Run("InsertFieldMemo")

    def InsertFieldRevisionChagne(self):
        """
        메모고침표 넣기(현재 한/글메뉴에 없음, 메모와 동일한 기능)
        """
        return self.hwp.HAction.Run("InsertFieldRevisionChagne")

    def InsertFixedWidthSpace(self):
        """
        고정폭 빈칸 삽입
        """
        return self.hwp.HAction.Run("InsertFixedWidthSpace")

    def InsertFootnote(self):
        """
        각주 입력
        """
        return self.hwp.HAction.Run("InsertFootnote")

    def InsertLastPrintDate(self):
        """
        상용구 코드 넣기(마지막 인쇄한 날짜)
        """
        return self.hwp.HAction.Run("InsertLastPrintDate")

    def InsertLastSaveBy(self):
        """
        상용구 코드 넣기(마지막 저장한 사람)
        """
        return self.hwp.HAction.Run("InsertLastSaveBy")

    def InsertLastSaveDate(self):
        """
        상용구 코드 넣기(마지막 저장한 날짜)
        """
        return self.hwp.HAction.Run("InsertLastSaveDate")

    def InsertLine(self):
        """
        선 넣기
        """
        return self.hwp.HAction.Run("InsertLine")

    def InsertNonBreakingSpace(self):
        """
        묶음 빈칸 삽입
        """
        return self.hwp.HAction.Run("InsertNonBreakingSpace")

    def InsertPageNum(self):
        """
        쪽 번호 넣기
        """
        return self.hwp.HAction.Run("InsertPageNum")

    def InsertSoftHyphen(self):
        """
        하이픈 삽입
        """
        return self.hwp.HAction.Run("InsertSoftHyphen")

    def InsertSpace(self):
        """
        공백 삽입
        """
        return self.hwp.HAction.Run("InsertSpace")

    def InsertStringDateTime(self):
        """
        날짜/시간 넣기 - 문자열로 넣기([입력-날짜/시간-날짜/시간 문자열]메뉴와 동일)
        """
        return self.hwp.HAction.Run("InsertStringDateTime")

    def InsertTab(self):
        """
        탭 삽입
        """
        return self.hwp.HAction.Run("InsertTab")

    def InsertTpNo(self):
        """
        상용구 코드 넣기(전체 쪽수)
        """
        return self.hwp.HAction.Run("InsertTpNo")

    def Jajun(self):
        """
        한자 자전
        """
        return self.hwp.HAction.Run("Jajun")

    def LabelAdd(self):
        """
        라벨 새 쪽 추가하기
        """
        return self.hwp.HAction.Run("LabelAdd")

    def LabelTemplate(self):
        """
        라벨 문서 만들기
        """
        return self.hwp.HAction.Run("LabelTemplate")

    def LinkTextBox(self):
        """
        글상자 연결. 글상자가 선택되지 않았거나, 캐럿이 글상자 내부에 있지 않으면 동작하지 않는다.
        """
        return self.hwp.HAction.Run("LinkTextBox")

    def MacroPause(self):
        """
        매크로 실행 일시 중지 (정의/실행)
        """
        return self.hwp.HAction.Run("MacroPause")

    def MacroPlay1(self):
        """
        매크로 1
        """
        return self.hwp.HAction.Run("MacroPlay1")

    def MacroPlay10(self):
        """
        매크로 10
        """
        return self.hwp.HAction.Run("MacroPlay10")

    def MacroPlay11(self):
        """
        매크로 11
        """
        return self.hwp.HAction.Run("MacroPlay11")

    def MacroPlay2(self):
        """
        매크로 2
        """
        return self.hwp.HAction.Run("MacroPlay2")

    def MacroPlay3(self):
        """
        매크로 3
        """
        return self.hwp.HAction.Run("MacroPlay3")

    def MacroPlay4(self):
        """
        매크로 4
        """
        return self.hwp.HAction.Run("MacroPlay4")

    def MacroPlay5(self):
        """
        매크로 5
        """
        return self.hwp.HAction.Run("MacroPlay5")

    def MacroPlay6(self):
        """
        매크로 6
        """
        return self.hwp.HAction.Run("MacroPlay6")

    def MacroPlay7(self):
        """
        매크로 7
        """
        return self.hwp.HAction.Run("MacroPlay7")

    def MacroPlay8(self):
        """
        매크로 8
        """
        return self.hwp.HAction.Run("MacroPlay8")

    def MacroPlay9(self):
        """
        매크로 9
        """
        return self.hwp.HAction.Run("MacroPlay9")

    def MacroRepeat(self):
        """
        매크로 실행
        """
        return self.hwp.HAction.Run("MacroRepeat")

    def MacroStop(self):
        """
        매크로 실행 중지 (정의/실행)
        """
        return self.hwp.HAction.Run("MacroStop")

    def MailMergeField(self):
        """
        메일 머지 필드(표시달기 or 고치기)
        """
        return self.hwp.HAction.Run("MailMergeField")

    def MakeIndex(self):
        """
        찾아보기 만들기
        """
        return self.hwp.HAction.Run("MakeIndex")

    def ManualChangeHangul(self):
        """
        한영 수동 전환.. 현재 커서위치 또는 문단나누기 이전에 입력된 내용에 대해서 강제적으로 한/영 전환을 한다.
        """
        return self.hwp.HAction.Run("ManualChangeHangul")

    def MarkTitle(self):
        """
        제목 차례 표시([도구-차례/찾아보기-제목 차례 표시]메뉴에 대응). 차례 코드가 삽입되어 나중에 차례 만들기에서 사용할 수 있다.적용여부는 Ctrl+G,C를 이용해 조판부호를 확인하면 알 수 있다.
        """
        return self.hwp.HAction.Run("MarkTitle")

    def MasterPageDuplicate(self):
        """
        기존 바탕쪽과 겹침. 바탕쪽 편집상태가 활성화되어 있으며 [구역 마지막쪽], [구역임의 쪽]일 경우에만 사용 가능하다.
        """
        return self.hwp.HAction.Run("MasterPageDuplicate")

    def MasterPageExcept(self):
        """
        첫 쪽 제외
        """
        return self.hwp.HAction.Run("MasterPageExcept")

    def MasterPageFront(self):
        """
        바탕쪽 앞으로 보내기. 바탕쪽 편집모드일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("MasterPageFront")

    def MasterPagePrevSection(self):
        """
        앞 구역 바탕쪽 사용
        """
        return self.hwp.HAction.Run("MasterPagePrevSection")

    def MasterPageToNext(self):
        """
        이후 바탕쪽
        """
        return self.hwp.HAction.Run("MasterPageToNext")

    def MasterPageToPrevious(self):
        """
        이전 바탕쪽
        """
        return self.hwp.HAction.Run("MasterPageToPrevious")

    def ModifyComposeChars(self):
        """
        고치기 - 글자 겹침
        """
        return self.hwp.HAction.Run("ModifyComposeChars")

    def ModifyCtrl(self):
        """
        고치기 : 컨트롤
        """
        return self.hwp.HAction.Run("ModifyCtrl")

    def ModifyDutmal(self):
        """
        고치기 - 덧말
        """
        return self.hwp.HAction.Run("ModifyDutmal")

    def ModifyFillProperty(self):
        """
        고치기(채우기 속성 탭으로). 만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다. 채우기 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다.
        """
        return self.hwp.HAction.Run("ModifyFillProperty")

    def ModifyLineProperty(self):
        """
        고치기(선/테두리 속성 탭으로). 만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다. 선/테두리 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다.
        """
        return self.hwp.HAction.Run("ModifyLineProperty")

    def ModifyShapeObject(self):
        """
        고치기 - 개체 속성
        """
        return self.hwp.HAction.Run("ModifyShapeObject")

    def MoveColumnBegin(self):
        """
        단의 시작점으로 이동한다. 단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다.
        """
        return self.hwp.HAction.Run("MoveColumnBegin")

    def MoveColumnEnd(self):
        """
        단의 끝점으로 이동한다. 단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다.
        """
        return self.hwp.HAction.Run("MoveColumnEnd")

    def MoveDocBegin(self):
        """
        문서의 시작으로 이동.. 만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다. 현재 서브 리스트 내에 있으면 빠져나간다.
        """
        return self.hwp.HAction.Run("MoveDocBegin")

    def MoveDocEnd(self):
        """
        문서의 끝으로 이동.. 만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다. 현재 서브 리스트 내에 있으면 빠져나간다.
        """
        return self.hwp.HAction.Run("MoveDocEnd")

    def MoveDown(self):
        """
        캐럿을 (논리적 개념의) 아래로 이동시킨다.
        """
        return self.hwp.HAction.Run("MoveDown")

    def MoveLeft(self):
        """
        캐럿을 (논리적 개념의) 왼쪽으로 이동시킨다.
        """
        return self.hwp.HAction.Run("MoveLeft")

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
        return self.hwp.HAction.Run("MoveLineEnd")

    def MoveLineUp(self):
        """
        한 줄 위로 이동한다.
        """
        return self.hwp.HAction.Run("MoveLineUp")

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
        한 글자 뒤로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MoveNextChar")

    def MoveNextColumn(self):
        """
        뒤 단으로 이동
        """
        return self.hwp.HAction.Run("MoveNextColumn")

    def MoveNextParaBegin(self):
        """
        다음 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MoveNextParaBegin")

    def MoveNextPos(self):
        """
        한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다.
        """
        return self.hwp.HAction.Run("MoveNextPos")

    def MoveNextPosEx(self):
        """
        한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함)
        """
        return self.hwp.HAction.Run("MoveNextPosEx")

    def MoveNextWord(self):
        """
        한 단어 뒤로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MoveNextWord")

    def MovePageBegin(self):
        """
        현재 페이지의 시작점으로 이동한다.. 만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 상단으로 페이지뷰잉이 맞춰진다.
        """
        return self.hwp.HAction.Run("MovePageBegin")

    def MovePageDown(self):
        """
        앞 페이지의 시작으로 이동. 현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다.
        """
        return self.hwp.HAction.Run("MovePageDown")

    def MovePageEnd(self):
        """
        현재 페이지의 끝점으로 이동한다.. 만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 하단으로 페이지뷰잉이 맞춰진다.
        """
        return self.hwp.HAction.Run("MovePageEnd")

    def MovePageUp(self):
        """
        뒤 페이지의 시작으로 이동. 현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다.
        """
        return self.hwp.HAction.Run("MovePageUp")

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
        한 레벨 상위/탑레벨/루트 리스트로 이동한다.. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
        """
        return self.hwp.HAction.Run("MoveParentList")

    def MovePrevChar(self):
        """
        한 글자 앞 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MovePrevChar")

    def MovePrevColumn(self):
        """
        앞 단으로 이동
        """
        return self.hwp.HAction.Run("MovePrevColumn")

    def MovePrevParaBegin(self):
        """
        앞 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MovePrevParaBegin")

    def MovePrevParaEnd(self):
        """
        앞 문단의 끝으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MovePrevParaEnd")

    def MovePrevPos(self):
        """
        한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다.
        """
        return self.hwp.HAction.Run("MovePrevPos")

    def MovePrevPosEx(self):
        """
        한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함)
        """
        return self.hwp.HAction.Run("MovePrevPosEx")

    def MovePrevWord(self):
        """
        한 단어 앞으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        return self.hwp.HAction.Run("MovePrevWord")

    def MoveRight(self):
        """
        캐럿을 (논리적 개념의) 오른쪽으로 이동시킨다.
        """
        return self.hwp.HAction.Run("MoveRight")

    def MoveRootList(self):
        """
        한 레벨 상위/탑레벨/루트 리스트로 이동한다.. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
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
        뒤 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.
        """
        return self.hwp.HAction.Run("MoveSectionDown")

    def MoveSectionUp(self):
        """
        앞 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.
        """
        return self.hwp.HAction.Run("MoveSectionUp")

    def MoveSelDocBegin(self):
        """
        셀렉션: 문서 처음
        """
        return self.hwp.HAction.Run("MoveSelDocBegin")

    def MoveSelDocEnd(self):
        """
        셀렉션: 문서 끝
        """
        return self.hwp.HAction.Run("MoveSelDocEnd")

    def MoveSelDown(self):
        """
        셀렉션: 캐럿을 (논리적 방향) 아래로 이동
        """
        return self.hwp.HAction.Run("MoveSelDown")

    def MoveSelLeft(self):
        """
        셀렉션: 캐럿을 (논리적 방향) 왼쪽으로 이동
        """
        return self.hwp.HAction.Run("MoveSelLeft")

    def MoveSelLineBegin(self):
        """
        셀렉션: 줄 처음
        """
        return self.hwp.HAction.Run("MoveSelLineBegin")

    def MoveSelLineDown(self):
        """
        셀렉션: 한줄 아래
        """
        return self.hwp.HAction.Run("MoveSelLineDown")

    def MoveSelLineEnd(self):
        """
        셀렉션: 줄 끝
        """
        return self.hwp.HAction.Run("MoveSelLineEnd")

    def MoveSelLineUp(self):
        """
        셀렉션: 한줄 위
        """
        return self.hwp.HAction.Run("MoveSelLineUp")

    def MoveSelListBegin(self):
        """
        셀렉션: 리스트 처음
        """
        return self.hwp.HAction.Run("MoveSelListBegin")

    def MoveSelListEnd(self):
        """
        셀렉션: 리스트 끝
        """
        return self.hwp.HAction.Run("MoveSelListEnd")

    def MoveSelNextChar(self):
        """
        셀렉션: 다음 글자
        """
        return self.hwp.HAction.Run("MoveSelNextChar")

    def MoveSelNextParaBegin(self):
        """
        셀렉션: 다음 문단 처음
        """
        return self.hwp.HAction.Run("MoveSelNextParaBegin")

    def MoveSelNextPos(self):
        """
        셀렉션: 다음 위치
        """
        return self.hwp.HAction.Run("MoveSelNextPos")

    def MoveSelNextWord(self):
        """
        셀렉션: 다음 단어
        """
        return self.hwp.HAction.Run("MoveSelNextWord")

    def MoveSelPageDown(self):
        """
        셀렉션: 페이지다운
        """
        return self.hwp.HAction.Run("MoveSelPageDown")

    def MoveSelPageUp(self):
        """
        셀렉션: 페이지 업
        """
        return self.hwp.HAction.Run("MoveSelPageUp")

    def MoveSelParaBegin(self):
        """
        셀렉션: 문단 처음
        """
        return self.hwp.HAction.Run("MoveSelParaBegin")

    def MoveSelParaEnd(self):
        """
        셀렉션: 문단 끝
        """
        return self.hwp.HAction.Run("MoveSelParaEnd")

    def MoveSelPrevChar(self):
        """
        셀렉션: 이전 글자
        """
        return self.hwp.HAction.Run("MoveSelPrevChar")

    def MoveSelPrevParaBegin(self):
        """
        셀렉션: 이전 문단 시작
        """
        return self.hwp.HAction.Run("MoveSelPrevParaBegin")

    def MoveSelPrevParaEnd(self):
        """
        셀렉션: 이전 문단 끝
        """
        return self.hwp.HAction.Run("MoveSelPrevParaEnd")

    def MoveSelPrevPos(self):
        """
        셀렉션: 이전 위치
        """
        return self.hwp.HAction.Run("MoveSelPrevPos")

    def MoveSelPrevWord(self):
        """
        셀렉션: 이전 단어
        """
        return self.hwp.HAction.Run("MoveSelPrevWord")

    def MoveSelRight(self):
        """
        셀렉션: 캐럿을 (논리적 방향) 오른쪽으로 이동
        """
        return self.hwp.HAction.Run("MoveSelRight")

    def MoveSelTopLevelBegin(self):
        """
        셀렉션: 처음
        """
        return self.hwp.HAction.Run("MoveSelTopLevelBegin")

    def MoveSelTopLevelEnd(self):
        """
        셀렉션: 끝
        """
        return self.hwp.HAction.Run("MoveSelTopLevelEnd")

    def MoveSelUp(self):
        """
        셀렉션: 캐럿을 (논리적 방향) 위로 이동
        """
        return self.hwp.HAction.Run("MoveSelUp")

    def MoveSelViewDown(self):
        """
        셀렉션: 아래
        """
        return self.hwp.HAction.Run("MoveSelViewDown")

    def MoveSelViewUp(self):
        """
        셀렉션: 위
        """
        return self.hwp.HAction.Run("MoveSelViewUp")

    def MoveSelWordBegin(self):
        """
        셀렉션: 단어 처음
        """
        return self.hwp.HAction.Run("MoveSelWordBegin")

    def MoveSelWordEnd(self):
        """
        셀렉션: 단어 끝
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
        한 레벨 상위/탑레벨/루트 리스트로 이동한다.. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
        """
        return self.hwp.HAction.Run("MoveTopLevelList")

    def MoveUp(self):
        """
        캐럿을 (논리적 개념의) 위로 이동시킨다.
        """
        return self.hwp.HAction.Run("MoveUp")

    def MoveViewBegin(self):
        """
        현재 뷰의 시작에 위치한 곳으로 이동
        """
        return self.hwp.HAction.Run("MoveViewBegin")

    def MoveViewDown(self):
        """
        현재 뷰의 크기만큼 아래로 이동한다. PgDn 키의 기능이다.
        """
        return self.hwp.HAction.Run("MoveViewDown")

    def MoveViewEnd(self):
        """
        현재 뷰의 끝에 위치한 곳으로 이동
        """
        return self.hwp.HAction.Run("MoveViewEnd")

    def MoveViewUp(self):
        """
        현재 뷰의 크기만큼 위로 이동한다. PgUp 키의 기능이다.
        """
        return self.hwp.HAction.Run("MoveViewUp")

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

    def ParagraphShapeWithNext(self):
        """
        다음 문단과 함께
        """
        return self.hwp.HAction.Run("ParagraphShapeWithNext")

    def Paste(self):
        """
        붙이기
        """
        return self.hwp.HAction.Run("Paste")

    def PastePage(self):
        """
        쪽 붙여넣기
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
        그림 넣기 (대화상자를 띄워 선택한 이미지 파일을 문서에 삽입하는 액션 : API용)
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
        그림 원래 그림으로
        """
        return self.hwp.HAction.Run("PictureToOriginal")

    def PrevTextBoxLinked(self):
        """
        연결된 글상자의 이전 글상자로 이동. 현재 글상자가 선택되거나, 글상자 내부에 캐럿이 존재하지 않으면 동작하지 않는다.
        """
        return self.hwp.HAction.Run("PrevTextBoxLinked")

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
        빠른 교정 ―내용 편집
        """
        return self.hwp.HAction.Run("QuickCorrect Run")

    def QuickCorrectSound(self):
        """
        빠른 교정 ― 메뉴에서 효과음 On/Off
        """
        return self.hwp.HAction.Run("QuickCorrect Sound")

    def QuickMarkInsert0(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert0")

    def QuickMarkInsert1(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert1")

    def QuickMarkInsert2(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert2")

    def QuickMarkInsert3(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert3")

    def QuickMarkInsert4(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert4")

    def QuickMarkInsert5(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert5")

    def QuickMarkInsert6(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert6")

    def QuickMarkInsert7(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert7")

    def QuickMarkInsert8(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert8")

    def QuickMarkInsert9(self):
        """
        쉬운 책갈피 - 삽입
        """
        return self.hwp.HAction.Run("QuickMarkInsert9")

    def QuickMarkMove0(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove0")

    def QuickMarkMove1(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove1")

    def QuickMarkMove2(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove2")

    def QuickMarkMove3(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove3")

    def QuickMarkMove4(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove4")

    def QuickMarkMove5(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove5")

    def QuickMarkMove6(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove6")

    def QuickMarkMove7(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove7")

    def QuickMarkMove8(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove8")

    def QuickMarkMove9(self):
        """
        쉬운 책갈피 - 이동
        """
        return self.hwp.HAction.Run("QuickMarkMove9")

    def RecalcPageCount(self):
        """
        현재 페이지의 쪽 번호 재계산
        """
        return self.hwp.HAction.Run("RecalcPageCount")

    def RecentCode(self):
        """
        최근에 사용한 문자표 입력. 최근에 사용한 문자표가 없을 경우에는 문자표 대화상자를 띄운다.
        """
        return self.hwp.HAction.Run("RecentCode")

    def Redo(self):
        """
        다시 실행
        """
        return self.hwp.HAction.Run("Redo")

    def returnKeyInField(self):
        """
        캐럿이 필드 안에 위치한 상태에서 return Key에 대한 액션 분기
        """
        return self.hwp.HAction.Run("returnKeyInField")

    def returnPrevPos(self):
        """
        직전위치로 돌아가기
        """
        return self.hwp.HAction.Run("returnPrevPos")

    def ScrMacroPause(self):
        """
        매크로 기록 일시정지/재시작
        """
        return self.hwp.HAction.Run("ScrMacroPause")

    def ScrMacroPlay1(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay1")

    def ScrMacroPlay2(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay2")

    def ScrMacroPlay3(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay3")

    def ScrMacroPlay4(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay4")

    def ScrMacroPlay5(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay5")

    def ScrMacroPlay6(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay6")

    def ScrMacroPlay7(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay7")

    def ScrMacroPlay8(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay8")

    def ScrMacroPlay9(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay9")

    def ScrMacroPlay10(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay10")

    def ScrMacroPlay11(self):
        """
        #번 매크로 실행(Alt+Shift+#)
        """
        return self.hwp.HAction.Run("ScrMacroPlay11")

    def ScrMacroStop(self):
        """
        매크로 기록 중지
        """
        return self.hwp.HAction.Run("ScrMacroStop")

    def Select(self):
        """
        선택 (F3 Key를 누른 효과)
        """
        return self.hwp.HAction.Run("Select")

    def SelectAll(self):
        """
        모두 선택
        """
        return self.hwp.HAction.Run("SelectAll")

    def SelectColumn(self):
        """
        칸 블록 선택 (F4 Key를 누른 효과)
        """
        return self.hwp.HAction.Run("SelectColumn")

    def SelectCtrlFront(self):
        """
        개체선택 정방향
        """
        return self.hwp.HAction.Run("SelectCtrlFront")

    def SelectCtrlReverse(self):
        """
        개체선택 역방향
        """
        return self.hwp.HAction.Run("SelectCtrlReverse")

    def SendBrowserText(self):
        """
        브라우저로 보내기
        """
        return self.hwp.HAction.Run("SendBrowserText")

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
        return self.hwp.HAction.Run("ShapeObjAttachCaption")

    def ShapeObjAttachTextBox(self):
        """
        글 상자로 만들기
        """
        return self.hwp.HAction.Run("ShapeObjAttachTextBox")

    def ShapeObjBringForward(self):
        """
        앞으로
        """
        return self.hwp.HAction.Run("ShapeObjBringForward")

    def ShapeObjBringInFrontOfText(self):
        """
        글 앞으로
        """
        return self.hwp.HAction.Run("ShapeObjBringInFrontOfText")

    def ShapeObjBringToFront(self):
        """
        맨 앞으로
        """
        return self.hwp.HAction.Run("ShapeObjBringToFront")

    def ShapeObjCtrlSendBehindText(self):
        """
        글 뒤로
        """
        return self.hwp.HAction.Run("ShapeObjCtrlSendBehindText")

    def ShapeObjDetachCaption(self):
        """
        캡션 없애기
        """
        return self.hwp.HAction.Run("ShapeObjDetachCaption")

    def ShapeObjDetachTextBox(self):
        """
        글상자 속성 없애기
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

    def ShapeObjHorzFlip(self):
        """
        그리기 개체 좌우 뒤집기
        """
        return self.hwp.HAction.Run("ShapeObjHorzFlip")

    def ShapeObjHorzFlipOrgState(self):
        """
        그리기 개체 좌우 뒤집기 원상태로 되돌리기
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
        90도 회전
        """
        return self.hwp.HAction.Run("ShapeObjRightAngleRotater")

    def ShapeObjRightAngleRotaterAnticlockwise(self):
        return self.hwp.HAction.Run("ShapeObjRightAngleRotaterAnticlockwise")

    def ShapeObjRotater(self):
        """
        자유각 회전(회전중심 고정)
        """
        return self.hwp.HAction.Run("ShapeObjRotater")

    def ShapeObjSaveAsPicture(self):
        """
        그리기개체를 그림으로 저장하기
        """
        return self.hwp.HAction.Run("ShapeObjSaveAsPicture")

    def ShapeObjSelect(self):
        """
        틀 선택 도구
        """
        return self.hwp.HAction.Run("ShapeObjSelect")

    def ShapeObjSendBack(self):
        """
        뒤로
        """
        return self.hwp.HAction.Run("ShapeObjSendBack")

    def ShapeObjSendToBack(self):
        """
        맨 뒤로
        """
        return self.hwp.HAction.Run("ShapeObjSendToBack")

    def ShapeObjTableSelCell(self):
        """
        테이블 선택상태에서 첫 번째 셀 선택하기
        """
        return self.hwp.HAction.Run("ShapeObjTableSelCell")

    def ShapeObjTextBoxEdit(self):
        """
        글상자 선택상태에서 편집모드로 들어가기
        """
        return self.hwp.HAction.Run("ShapeObjTextBoxEdit")

    def ShapeObjUngroup(self):
        """
        틀 풀기
        """
        return self.hwp.HAction.Run("ShapeObjUngroup")

    def ShapeObjUnlockAll(self):
        """
        개체 Unlock All
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

    def SoftKeyboard(self):
        """
        보기
        """
        return self.hwp.HAction.Run("Soft Keyboard")

    def SpellingCheck(self):
        """
        맞춤법
        """
        return self.hwp.HAction.Run("SpellingCheck")

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

    def StyleClearCharStyle(self):
        """
        글자 스타일 해제
        """
        return self.hwp.HAction.Run("StyleClearCharStyle")

    def StyleShortcut1(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut1")

    def StyleShortcut10(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut10")

    def StyleShortcut2(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut2")

    def StyleShortcut3(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut3")

    def StyleShortcut4(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut4")

    def StyleShortcut5(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut5")

    def StyleShortcut6(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut6")

    def StyleShortcut7(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut7")

    def StyleShortcut8(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut8")

    def StyleShortcut9(self):
        """
        스타일 단축키
        """
        return self.hwp.HAction.Run("StyleShortcut9")

    def TableAppendRow(self):
        """
        줄 추가
        """
        return self.hwp.HAction.Run("TableAppendRow")

    def TableCellBlock(self):
        """
        셀 블록
        """
        return self.hwp.HAction.Run("TableCellBlock")

    def TableCellBlockCol(self):
        """
        셀 블록 (칸)
        """
        return self.hwp.HAction.Run("TableCellBlockCol")

    def TableCellBlockExtend(self):
        """
        셀 블록 연장(F5 + F5)
        """
        return self.hwp.HAction.Run("TableCellBlockExtend")

    def TableCellBlockExtendAbs(self):
        """
        셀 블록 연장(SHIFT + F5)
        """
        return self.hwp.HAction.Run("TableCellBlockExtendAbs")

    def TableCellBlockRow(self):
        """
        셀 블록(줄)
        """
        return self.hwp.HAction.Run("TableCellBlockRow")

    def TableCellBorderAll(self):
        """
        모든 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderAll")

    def TableCellBorderBottom(self):
        """
        가장 아래 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderBottom")

    def TableCellBorderDiagonalDown(self):
        """
        대각선(⍂) 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderDiagonalDown")

    def TableCellBorderDiagonalUp(self):
        """
        대각선(⍁) 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderDiagonalUp")

    def TableCellBorderInside(self):
        """
        모든 안쪽 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderInside")

    def TableCellBorderInsideHorz(self):
        """
        모든 안쪽 가로 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderInsideHorz")

    def TableCellBorderInsideVert(self):
        """
        모든 안쪽 세로 셀 테두리 toggle(있음/없음). 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderInsideVert")

    def TableCellBorderLeft(self):
        """
        가장 왼쪽의 셀 테두리 toggle(있음/없음) 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderLeft")

    def TableCellBorderNo(self):
        """
        모든 셀 테두리 지움. 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderNo")

    def TableCellBorderOutside(self):
        """
        바깥 셀 테두리 toggle(있음/없음) 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderOutside")

    def TableCellBorderRight(self):
        """
        가장 오른쪽의 셀 테두리 toggle(있음/없음) 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderRight")

    def TableCellBorderTop(self):
        """
        가장 위의 셀 테두리 toggle(있음/없음) 셀 블록 상태일 경우에만 동작한다.
        """
        return self.hwp.HAction.Run("TableCellBorderTop")

    def TableColBegin(self):
        """
        셀 이동: 열 시작
        """
        return self.hwp.HAction.Run("TableColBegin")

    def TableColEnd(self):
        """
        셀 이동: 열 끝
        """
        return self.hwp.HAction.Run("TableColEnd")

    def TableColPageDown(self):
        """
        셀 이동: 페이지다운
        """
        return self.hwp.HAction.Run("TableColPageDown")

    def TableColPageUp(self):
        """
        셀 이동: 페이지 업
        """
        return self.hwp.HAction.Run("TableColPageUp")

    def TableDeleteCell(self):
        """
        셀 삭제
        """
        return self.hwp.HAction.Run("TableDeleteCell")

    def TableDistributeCellHeight(self):
        """
        셀 높이를 같게
        """
        return self.hwp.HAction.Run("TableDistributeCellHeight")

    def TableDistributeCellWidth(self):
        """
        셀 너비를 같게
        """
        return self.hwp.HAction.Run("TableDistributeCellWidth")

    def TableDrawPen(self):
        """
        표 그리기
        """
        return self.hwp.HAction.Run("TableDrawPen")

    def TableEraser(self):
        """
        표 지우개
        """
        return self.hwp.HAction.Run("TableEraser")

    def TableFormulaAvgAuto(self):
        """
        블록 평균
        """
        return self.hwp.HAction.Run("TableFormulaAvgAuto")

    def TableFormulaAvgHor(self):
        """
        가로 평균
        """
        return self.hwp.HAction.Run("TableFormulaAvgHor")

    def TableFormulaAvgVer(self):
        """
        세로 평균
        """
        return self.hwp.HAction.Run("TableFormulaAvgVer")

    def TableFormulaProAuto(self):
        """
        블록 곱
        """
        return self.hwp.HAction.Run("TableFormulaProAuto")

    def TableFormulaProHor(self):
        """
        가로 곱
        """
        return self.hwp.HAction.Run("TableFormulaProHor")

    def TableFormulaProVer(self):
        """
        세로 곱
        """
        return self.hwp.HAction.Run("TableFormulaProVer")

    def TableFormulaSumAuto(self):
        """
        블록 합계
        """
        return self.hwp.HAction.Run("TableFormulaSumAuto")

    def TableFormulaSumHor(self):
        """
        가로 합계
        """
        return self.hwp.HAction.Run("TableFormulaSumHor")

    def TableFormulaSumVer(self):
        """
        세로 합계
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
        """
        return self.hwp.HAction.Run("TableMergeCell")

    def TableMergeTable(self):
        """
        표 붙이기
        """
        return self.hwp.HAction.Run("TableMergeTable")

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
        셀 크기 변경
        """
        return self.hwp.HAction.Run("TableResizeDown")

    def TableResizeExDown(self):
        """
        셀 크기 변경: 셀 아래. TebleResizeDown과 다른 점은 셀 블록 상태가 아니어도 동작한다는 점이다.
        """
        return self.hwp.HAction.Run("TableResizeExDown")

    def TableResizeExLeft(self):
        """
        셀 크기 변경: 셀 왼쪽. TebleResizeLeft와 다른 점은 셀 블록 상태가 아니어도 동작한다는 점이다.
        """
        return self.hwp.HAction.Run("TableResizeExLeft")

    def TableResizeExRight(self):
        """
        셀 크기 변경: 셀 오른쪽. TebleResizeRight와 다른 점은 셀 블록 상태가 아니어도 동작한다는 점이다.
        """
        return self.hwp.HAction.Run("TableResizeExRight")

    def TableResizeExUp(self):
        return self.hwp.HAction.Run("TableResizeExUp")

    def TableResizeLeft(self):
        """
        셀 크기 변경
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
        셀 크기 변경
        """
        return self.hwp.HAction.Run("TableResizeRight")

    def TableResizeUp(self):
        """
        셀 크기 변경
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
        """
        return self.hwp.HAction.Run("TableRightCellAppend")

    def TableSplitTable(self):
        """
        표 나누기
        """
        return self.hwp.HAction.Run("TableSplitTable")

    def TableUpperCell(self):
        """
        셀 이동: 셀 위
        """
        return self.hwp.HAction.Run("TableUpperCell")

    def TableVAlignBottom(self):
        """
        셀 세로정렬 아래
        """
        return self.hwp.HAction.Run("TableVAlignBottom")

    def TableVAlignCenter(self):
        """
        셀 세로정렬 가운데
        """
        return self.hwp.HAction.Run("TableVAlignCenter")

    def TableVAlignTop(self):
        """
        셀 세로정렬 위
        """
        return self.hwp.HAction.Run("TableVAlignTop")

    def ToggleOverwrite(self):
        """
        Toggle Overwrite
        """
        return self.hwp.HAction.Run("ToggleOverwrite")

    def Undo(self):
        """
        되살리기
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

    def ViewIdiom(self):
        """
        상용구 보기
        """
        return self.hwp.HAction.Run("ViewIdiom")

    def ViewOptionCtrlMark(self):
        """
        조판 부호
        """
        return self.hwp.HAction.Run("ViewOptionCtrlMark")

    def ViewOptionGuideLine(self):
        """
        안내선
        """
        return self.hwp.HAction.Run("ViewOptionGuideLine")

    def ViewOptionMemo(self):
        """
        메모 보이기/숨기기([보기-메모-메모 보이기/숨기기]메뉴와 동일)
        """
        return self.hwp.HAction.Run("ViewOptionMemo")

    def ViewOptionMemoGuideline(self):
        """
        메모 안내선 표시([보기-메모-메모 안내선 표시]메뉴와 동일)
        """
        return self.hwp.HAction.Run("ViewOptionMemoGuideline")

    def ViewOptionPaper(self):
        """
        쪽 윤곽 보기
        """
        return self.hwp.HAction.Run("ViewOptionPaper")

    def ViewOptionParaMark(self):
        """
        문단 부호
        """
        return self.hwp.HAction.Run("ViewOptionParaMark")

    def ViewOptionPicture(self):
        """
        그림 보이기/숨기기([보기-그림]메뉴와 동일)
        """
        return self.hwp.HAction.Run("ViewOptionPicture")

    def ViewOptionRevision(self):
        """
        교정부호 보이기/숨기기([보기-교정부호]메뉴와 동일)
        """
        return self.hwp.HAction.Run("ViewOptionRevision")

    def VoiceCommandConfig(self):
        """
        음성 명령 설정
        """
        return self.hwp.HAction.Run("VoiceCommand Config")

    def VoiceCommandResume(self):
        """
        음성 명령 레코딩 시작
        """
        return self.hwp.HAction.Run("VoiceCommand Resume")

    def VoiceCommandStop(self):
        """
        음성 명령 레코딩 중지
        """
        return self.hwp.HAction.Run("VoiceCommand Stop")

    def run_script_macro(self, function_name, u_macro_type=0, u_script_type=0):
        """
        한/글 문서 내에 존재하는 매크로를 실행한다.
        문서매크로, 스크립트매크로 모두 실행 가능하다.
        재미있는 점은 한/글 내에서 문서매크로 실행시
        New, Open 두 개의 함수 밖에 선택할 수 없으므로
        별도의 함수를 정의하더라도 이 두 함수 중 하나에서 호출해야 하지만,
        (진입점이 되어야 함)
        self.hwp.run_script_macro 명령어를 통해서는 제한없이 실행할 수 있다.

        :param function_name:
            실행할 매크로 함수이름(전체이름)

        :param u_macro_type:
            매크로의 유형. 밑의 값 중 하나이다.
            0: 스크립트 매크로(전역 매크로-HWP_GLOBAL_MACRO_TYPE, 기본값)
            1: 문서 매크로(해당문서에만 저장/적용되는 매크로-HWP_DOCUMENT_MACRO_TYPE)

        :param u_script_type:
            스크립트의 유형. 현재는 javascript만을 유일하게 지원한다.
            아무 정수나 입력하면 된다. (기본값: 0)

        :return:
            무조건 True를 반환(매크로의 실행여부와 상관없음)

        :example:
            >>> self.hwp.run_script_macro("OnDocument_New", u_macro_type=1)
            True
            >>> self.hwp.run_script_macro("OnScriptMacro_중국어1성")
            True
        """
        return self.hwp.RunScriptMacro(FunctionName=function_name, uMacroType=u_macro_type, uScriptType=u_script_type)

    def save(self, save_if_dirty=True):
        """
        현재 편집중인 문서를 저장한다.
        문서의 경로가 지정되어있지 않으면 “새 이름으로 저장” 대화상자가 뜬다.

        :param save_if_dirty:
            True를 지정하면 문서가 변경된 경우에만 저장한다.
            False를 지정하면 변경여부와 상관없이 무조건 저장한다.
            생략하면 True가 지정된다.

        :return:
            성공하면 True, 실패하면 False
        """
        return self.hwp.Save(save_if_dirty=save_if_dirty)

    def save_as(self, path, format="HWP", arg=""):
        """
        현재 편집중인 문서를 지정한 이름으로 저장한다.
        format, arg의 일반적인 개념에 대해서는 Open()참조.
        "Hwp" 포맷으로 파일 저장 시 arg에 지정할 수 있는 옵션은 다음과 같다.
        "lock:true" - 저장한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
        "backup:false" - 백업 파일 생성 여부
        "compress:true" - 압축 여부
        "fullsave:false" - 스토리지 파일을 완전히 새로 생성하여 저장
        "prvimage:2" - 미리보기 이미지 (0=off, 1=BMP, 2=GIF)
        "prvtext:1" - 미리보기 텍스트 (0=off, 1=on)
        "autosave:false" - 자동저장 파일로 저장할 지 여부 (TRUE: 자동저장, FALSE: 지정 파일로 저장)
        "export" - 다른 이름으로 저장하지만 열린 문서는 바꾸지 않는다.(lock:false와 함께 설정되어 있을 시 동작)
        여러 개를 한꺼번에 할 경우에는 세미콜론으로 구분하여 연속적으로 사용할 수 있다.
        "lock:TRUE;backup:FALSE;prvtext:1"

        :param path:
            문서 파일의 전체경로

        :param format:
            문서 형식. 생략하면 "HWP"가 지정된다.

        :param arg:
            세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.

        :return:
            성공하면 True, 실패하면 False
        """
        if not path.lower().startswith("c:"):
            path = os.path.join(os.getcwd(), path)
        return self.hwp.SaveAs(Path=path, Format=format, arg=arg)

    def scan_font(self):
        return self.hwp.ScanFont()

    def select_text(self, spara: Union[int, list, tuple] = 0, spos=0, epara=0, epos=0, slist=0):
        """
        특정 범위의 텍스트를 블록선택한다.
        epos가 가리키는 문자는 포함되지 않는다.

        :param spara:
            블록 시작 위치의 문단 번호.

        :param spos:
            블록 시작 위치의 문단 중에서 문자의 위치.

        :param epara:
            블록 끝 위치의 문단 번호.

        :param epos:
            블록 끝 위치의 문단 중에서 문자의 위치.

        :return:
            성공하면 True, 실패하면 False
        """
        if type(spara) in [list, tuple]:
            _, slist, spara, spos, _, epara, epos = spara
        self.set_pos(slist, 0, 0)
        return self.hwp.SelectText(spara=spara, spos=spos, epara=epara, epos=epos)

    def set_bar_code_image(self, lp_image_path, pgno, index, x, y, width, height):
        """
        작동하지 않는다.

        :param lp_image_path:
        :param pgno:
        :param index:
        :param x:
        :param y:
        :param width:
        :param height:
        :return:
        """
        if not lp_image_path.lower().startswith("c:"):
            lp_image_path = os.path.join(os.getcwd(), lp_image_path)
        return self.hwp.SetBarCodeImage(lpImagePath=lp_image_path, pgno=pgno, index=index,
                                        X=x, Y=y, Width=width, Height=height)

    def set_cur_field_name(self, field, option, direction, memo):
        """
        현재 캐럿이 위치하는 곳의 필드이름을 설정한다.
        GetFieldList()의 옵션 중에 4(hwpFieldSelection) 옵션은 사용하지 않는다.

        :param field:
            데이터 필드 이름

        :param option:
            다음과 같은 옵션을 지정할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다.
            1: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다.(hwpFieldCell)
            2: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)

        :param direction:
            누름틀 필드의 안내문. 누름틀 필드일 때만 유효하다.

        :param memo:
            누름틀 필드의 메모. 누름틀 필드일 때만 유효하다.

        :return:
            성공하면 True, 실패하면 False
        """
        return self.hwp.SetCurFieldName(Field=field, option=option, Direction=direction, memo=memo)

    def set_cur_metatag_name(self, tag):
        return self.hwp.SetCurMetatagName(tag=tag)

    def set_drm_authority(self, authority):
        return self.hwp.SetDRMAuthority(authority=authority)

    def set_field_view_option(self, option):
        """
        양식모드와 읽기전용모드일 때 현재 열린 문서의 필드의 겉보기 속성(『』표시)을 바꾼다.
        EditMode와 비슷하게 현재 열려있는 문서에 대한 속성이다. 따라서 저장되지 않는다.
        (작동하지 않음)

        :param option:
            겉보기 속성 bit
            1: 누름틀의 『』을 표시하지 않음, 기타필드의 『』을 표시하지 않음
            2: 누름틀의 『』을 빨간색으로 표시, 기타필드의 『』을 흰색으로 표시(기본값)
            3: 누름틀의 『』을 흰색으로 표시, 기타필드의 『』을 흰색으로 표시

        :return:
            설정된 속성이 반환된다.
            에러일 경우 0이 반환된다.
        """
        return self.hwp.SetFieldViewOption(option=option)

    def set_message_box_mode(self, mode):
        """
        한/글에서 쓰는 다양한 메시지박스가 뜨지 않고,
        자동으로 특정 버튼을 클릭한 효과를 주기 위해 사용한다.
        한/글에서 한/글이 로드된 후 SetMessageBoxMode()를 호출해서 사용한다.
        SetMessageBoxMode는 하나의 파라메터를 받으며,
        해당 파라메터는 자동으로 스킵할 버튼의 값으로 설정된다.
        예를 들어, MB_OK_IDOK (0x00000001)값을 주면,
        MB_OK형태의 메시지박스에서 OK버튼이 눌린 효과를 낸다.

        :param mode:
            // 메시지 박스의 종류
            #define MB_MASK						0x00FFFFFF
            // 1. 확인(MB_OK) : IDOK(1)
            #define MB_OK_IDOK						0x00000001
            #define MB_OK_MASK						0x0000000F
            // 2. 확인/취소(MB_OKCANCEL) : IDOK(1), IDCANCEL(2)
            #define MB_OKCANCEL_IDOK					0x00000010
            #define MB_OKCANCEL_IDCANCEL				0x00000020
            #define MB_OKCANCEL_MASK					0x000000F0
            // 3. 종료/재시도/무시(MB_ABORTRETRYIGNORE) : IDABORT(3), IDRETRY(4), IDIGNORE(5)
            #define MB_ABORTRETRYIGNORE_IDABORT			0x00000100
            #define MB_ABORTRETRYIGNORE_IDRETRY			0x00000200
            #define MB_ABORTRETRYIGNORE_IDIGNORE			0x00000400
            #define MB_ABORTRETRYIGNORE_MASK				0x00000F00
            // 4. 예/아니오/취소(MB_YESNOCANCEL) : IDYES(6), IDNO(7), IDCANCEL(2)
            #define MB_YESNOCANCEL_IDYES				0x00001000
            #define MB_YESNOCANCEL_IDNO				0x00002000
            #define MB_YESNOCANCEL_IDCANCEL				0x00004000
            #define MB_YESNOCANCEL_MASK				0x0000F000
            // 5. 예/아니오(MB_YESNO) : IDYES(6), IDNO(7)
            #define MB_YESNO_IDYES					0x00010000
            #define MB_YESNO_IDNO					0x00020000
            #define MB_YESNO_MASK					0x000F0000
            // 6. 재시도/취소(MB_RETRYCANCEL) : IDRETRY(4), IDCANCEL(2)
            #define MB_RETRYCANCEL_IDRETRY				0x00100000
            #define MB_RETRYCANCEL_IDCANCEL				0x00200000
            #define MB_RETRYCANCEL_MASK				0x00F00000

        :return:
            실행 전의 MessageBoxMode
        """
        return self.hwp.SetMessageBoxMode(Mode=mode)

    def set_pos(self, list, para, pos):
        """
        캐럿을 문서 내 특정 위치로 옮긴다.
        지정된 위치로 캐럿을 옮겨준다.

        :param list:
            캐럿이 위치한 문서 내 list ID

        :param para:
            캐럿이 위치한 문단 ID. 음수거나, 범위를 넘어가면 문서의 시작으로 이동하며, pos는 무시한다.

        :param pos:
            캐럿이 위치한 문단 내 글자 위치. -1을 주면 해당문단의 끝으로 이동한다.
            단 para가 범위 밖일 경우 pos는 무시되고 문서의 시작으로 캐럿을 옮긴다.

        :return:
            성공하면 True, 실패하면 False
        """
        return self.hwp.SetPos(List=list, Para=para, pos=pos)

    def set_pos_by_set(self, disp_val):
        """
        캐럿을 ParameterSet으로 얻어지는 위치로 옮긴다.

        :param disp_val:
            캐럿을 옮길 위치에 대한 ParameterSet 정보

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> start_pos = self.hwp.GetPosBySet()  # 현재 위치를 저장하고,
            >>> self.hwp.set_pos_by_set(start_pos)  # 특정 작업 후에 저장위치로 재이동
        """
        return self.hwp.SetPosBySet(dispVal=disp_val)

    def set_private_info_password(self, password):
        """
        개인정보보호를 위한 암호를 등록한다.
        개인정보 보호를 설정하기 위해서는
        우선 개인정보 보호 암호를 먼저 설정해야 한다.
        그러므로 개인정보 보호 함수를 실행하기 이전에
        반드시 이 함수를 호출해야 한다.
        (현재 작동하지 않는다.)

        :param password:
            새 암호

        :return:
            정상적으로 암호가 설정되면 true를 반환한다.
            암호설정에 실패하면 false를 반환한다. false를 반환하는 경우는 다음과 같다
            1. 암호의 길이가 너무 짧거나 너무 길 때 (영문: 5~44자, 한글: 3~22자)
            2. 암호가 이미 설정되었음. 또는 암호가 이미 설정된 문서임
        """
        return self.hwp.SetPrivateInfoPassword(Password=password)

    def set_text_file(self, data: str, format="HWPML2X", option=""):
        """
        문서를 문자열로 지정한다.

        :param data:
            문자열로 변경된 text 파일
        :param format:
            파일의 형식
            "HWP": HWP native format. BASE64 로 인코딩되어 있어야 한다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다.ver:0x0505010B
            "HWPML2X": HWP 형식과 호환. 문서의 모든 정보를 유지
            "HTML": 인터넷 문서 HTML 형식. 한/글 고유의 서식은 손실된다.
            "UNICODE": 유니코드 텍스트, 서식정보가 없는 텍스트만 저장
            "TEXT": 일반 텍스트, 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다.

        :param option:
            "insertfile": 현재커서 이후에 지정된 파일 삽입

        :return:
            성공이면 1을, 실패하면 0을 반환한다.
        """
        return self.hwp.SetTextFile(data=data, Format=format, option=option)

    def set_title_name(self, title):
        """
        한/글 프로그램의 타이틀을 변경한다.
        파일명과 무관하게 설정할 수 있으며,
        모든 특수문자를 허용한다.

        :param title:
            변경할 타이틀 문자열

        :return:
            성공시 True
        """
        return self.hwp.SetTitleName(Title=title)

    def set_user_info(self, user_info_id, value):
        return self.hwp.SetUserInfo(userInfoId=user_info_id, Value=value)

    def set_visible(self, visible):
        """
        현재 조작중인 한/글 인스턴스의 백그라운드 숨김여부를 변경할 수 있다.

        :param visible:
            visible=False로 설정하면 현재 조작중인 한/글 인스턴스가 백그라운드로 숨겨진다.

        :return:
        """
        self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible

    def side_type(self, side_type):
        return self.hwp.SideType(SideType=side_type)

    def signature(self, signature):
        return self.hwp.Signature(Signature=signature)

    def slash(self, slash):
        return self.hwp.Slash(Slash=slash)

    def solar_to_lunar(self, s_year, s_month, s_day, l_year, l_month, l_day, l_leap):
        return self.hwp.SolarToLunar(sYear=s_year, sMonth=s_month, sDay=s_day,
                                     lYear=l_year, lMonth=l_month, lDay=l_day, lLeap=l_leap)

    def solar_to_lunar_by_set(self, s_year, s_month, s_day):
        return self.hwp.SolarToLunarBySet(sYear=s_year, sMonth=s_month, sDay=s_day)

    def sort_delimiter(self, sort_delimiter):
        return self.hwp.SortDelimiter(SortDelimiter=sort_delimiter)

    def strike_out(self, strike_out_type):
        return self.hwp.StrikeOut(StrikeOutType=strike_out_type)

    def style_type(self, style_type):
        return self.hwp.StyleType(StyleType=style_type)

    def subt_pos(self, subt_pos):
        return self.hwp.SubtPos(SubtPos=subt_pos)

    def table_break(self, page_break):
        return self.hwp.TableBreak(PageBreak=page_break)

    def table_format(self, table_format):
        return self.hwp.TableFormat(TableFormat=table_format)

    def table_swap_type(self, tableswap):
        return self.hwp.TableSwapType(tableswap=tableswap)

    def table_target(self, table_target):
        return self.hwp.TableTarget(TableTarget=table_target)

    def text_align(self, text_align):
        return self.hwp.TextAlign(TextAlign=text_align)

    def text_art_align(self, text_art_align):
        return self.hwp.TextArtAlign(TextArtAlign=text_art_align)

    def text_dir(self, text_direction):
        return self.hwp.TextDir(TextDirection=text_direction)

    def text_flow_type(self, text_flow):
        return self.hwp.TextFlowType(TextFlow=text_flow)

    def text_wrap_type(self, text_wrap):
        return self.hwp.TextWrapType(TextWrap=text_wrap)

    def un_select_ctrl(self):
        return self.hwp.UnSelectCtrl()

    def v_align(self, v_align):
        return self.hwp.VAlign(VAlign=v_align)

    def vert_rel(self, vert_rel):
        return self.hwp.VertRel(VertRel=vert_rel)

    def view_flag(self, view_flag):
        return self.hwp.ViewFlag(ViewFlag=view_flag)

    def watermark_brush(self, watermark_brush):
        return self.hwp.WatermarkBrush(WatermarkBrush=watermark_brush)

    def width_rel(self, width_rel):
        return self.hwp.WidthRel(WidthRel=width_rel)


hwpx__ = Hwp(visible=False)

hwp__ = hwpx__.hwp

try:
    Application = hwp__.Application
except:
    pass

try:
    ArcType = hwp__.ArcType
except:
    pass

try:
    AutoNumType = hwp__.AutoNumType
except:
    pass

try:
    BorderShape = hwp__.BorderShape
except:
    pass

try:
    BreakWordLatin = hwp__.BreakWordLatin
except:
    pass

try:
    BrushType = hwp__.BrushType
except:
    pass

try:
    CLSID = hwp__.CLSID
except:
    pass

try:
    Canonical = hwp__.Canonical
except:
    pass

try:
    CellApply = hwp__.CellApply
except:
    pass

try:
    CellShape = hwp__.CellShape
except:
    pass

try:
    CharShadowType = hwp__.CharShadowType
except:
    pass

try:
    CharShape = hwp__.CharShape
except:
    pass

try:
    CheckXObject = hwp__.CheckXObject
except:
    pass

try:
    Clear = hwp__.Clear
except:
    pass

try:
    ColDefType = hwp__.ColDefType
except:
    pass

try:
    ColLayoutType = hwp__.ColLayoutType
except:
    pass

try:
    ConvertPUAHangulToUnicode = hwp__.ConvertPUAHangulToUnicode
except:
    pass

try:
    CreateAction = hwp__.CreateAction
except:
    pass

try:
    CreateField = hwp__.CreateField
except:
    pass

try:
    CreateID = hwp__.CreateID
except:
    pass

try:
    CreateMode = hwp__.CreateMode
except:
    pass

try:
    CreatePageImage = hwp__.CreatePageImage
except:
    pass

try:
    CreateSet = hwp__.CreateSet
except:
    pass

try:
    CrookedSlash = hwp__.CrookedSlash
except:
    pass

try:
    CurFieldState = hwp__.CurFieldState
except:
    pass

try:
    CurMetatagState = hwp__.CurMetatagState
except:
    pass

try:
    CurSelectedCtrl = hwp__.CurSelectedCtrl
except:
    pass

try:
    DSMark = hwp__.DSMark
except:
    pass

try:
    DbfCodeType = hwp__.DbfCodeType
except:
    pass

try:
    DeleteCtrl = hwp__.DeleteCtrl
except:
    pass

try:
    Delimiter = hwp__.Delimiter
except:
    pass

try:
    DrawAspect = hwp__.DrawAspect
except:
    pass

try:
    DrawFillImage = hwp__.DrawFillImage
except:
    pass

try:
    DrawShadowType = hwp__.DrawShadowType
except:
    pass

try:
    EditMode = hwp__.EditMode
except:
    pass

try:
    Encrypt = hwp__.Encrypt
except:
    pass

try:
    EndSize = hwp__.EndSize
except:
    pass

try:
    EndStyle = hwp__.EndStyle
except:
    pass

try:
    EngineProperties = hwp__.EngineProperties
except:
    pass

try:
    ExportStyle = hwp__.ExportStyle
except:
    pass

try:
    FieldExist = hwp__.FieldExist
except:
    pass

try:
    FileTranslate = hwp__.FileTranslate
except:
    pass

try:
    FillAreaType = hwp__.FillAreaType
except:
    pass

try:
    FindCtrl = hwp__.FindCtrl
except:
    pass

try:
    FindDir = hwp__.FindDir
except:
    pass

try:
    FindPrivateInfo = hwp__.FindPrivateInfo
except:
    pass

try:
    FontType = hwp__.FontType
except:
    pass

try:
    GetBinDataPath = hwp__.GetBinDataPath
except:
    pass

try:
    GetCurFieldName = hwp__.GetCurFieldName
except:
    pass

try:
    GetCurMetatagName = hwp__.GetCurMetatagName
except:
    pass

try:
    GetFieldList = hwp__.GetFieldList
except:
    pass

try:
    GetFieldText = hwp__.GetFieldText
except:
    pass

try:
    GetFileInfo = hwp__.GetFileInfo
except:
    pass

try:
    GetFontList = hwp__.GetFontList
except:
    pass

try:
    GetHeadingString = hwp__.GetHeadingString
except:
    pass

try:
    GetMessageBoxMode = hwp__.GetMessageBoxMode
except:
    pass

try:
    GetMetatagList = hwp__.GetMetatagList
except:
    pass

try:
    GetMetatagNameText = hwp__.GetMetatagNameText
except:
    pass

try:
    GetMousePos = hwp__.GetMousePos
except:
    pass

try:
    GetPageText = hwp__.GetPageText
except:
    pass

try:
    GetPos = hwp__.GetPos
except:
    pass

try:
    GetPosBySet = hwp__.GetPosBySet
except:
    pass

try:
    GetScriptSource = hwp__.GetScriptSource
except:
    pass

try:
    GetSelectedPos = hwp__.GetSelectedPos
except:
    pass

try:
    GetSelectedPosBySet = hwp__.GetSelectedPosBySet
except:
    pass

try:
    GetText = hwp__.GetText
except:
    pass

try:
    GetTextFile = hwp__.GetTextFile
except:
    pass

try:
    GetTranslateLangList = hwp__.GetTranslateLangList
except:
    pass

try:
    GetUserInfo = hwp__.GetUserInfo
except:
    pass

try:
    Gradation = hwp__.Gradation
except:
    pass

try:
    GridMethod = hwp__.GridMethod
except:
    pass

try:
    GridViewLine = hwp__.GridViewLine
except:
    pass

try:
    GutterMethod = hwp__.GutterMethod
except:
    pass

try:
    HAction = hwp__.HAction
except:
    pass

try:
    HAlign = hwp__.HAlign
except:
    pass

try:
    HParameterSet = hwp__.HParameterSet
except:
    pass

try:
    Handler = hwp__.Handler
except:
    pass

try:
    Hash = hwp__.Hash
except:
    pass

try:
    HatchStyle = hwp__.HatchStyle
except:
    pass

try:
    HeadCtrl = hwp__.HeadCtrl
except:
    pass

try:
    HeadType = hwp__.HeadType
except:
    pass

try:
    HeightRel = hwp__.HeightRel
except:
    pass

try:
    Hiding = hwp__.Hiding
except:
    pass

try:
    HorzRel = hwp__.HorzRel
except:
    pass

try:
    HwpLineType = hwp__.HwpLineType
except:
    pass

try:
    HwpLineWidth = hwp__.HwpLineWidth
except:
    pass

try:
    HwpOutlineStyle = hwp__.HwpOutlineStyle
except:
    pass

try:
    HwpOutlineType = hwp__.HwpOutlineType
except:
    pass

try:
    HwpUnderlineShape = hwp__.HwpUnderlineShape
except:
    pass

try:
    HwpUnderlineType = hwp__.HwpUnderlineType
except:
    pass

try:
    HwpZoomType = hwp__.HwpZoomType
except:
    pass

try:
    ImageFormat = hwp__.ImageFormat
except:
    pass

try:
    ImportStyle = hwp__.ImportStyle
except:
    pass

try:
    InitHParameterSet = hwp__.InitHParameterSet
except:
    pass

try:
    InitScan = hwp__.InitScan
except:
    pass

try:
    Insert = hwp__.Insert
except:
    pass

try:
    InsertBackgroundPicture = hwp__.InsertBackgroundPicture
except:
    pass

try:
    InsertCtrl = hwp__.InsertCtrl
except:
    pass

try:
    InsertPicture = hwp__.InsertPicture
except:
    pass

try:
    IsActionEnable = hwp__.IsActionEnable
except:
    pass

try:
    IsCommandLock = hwp__.IsCommandLock
except:
    pass

try:
    IsEmpty = hwp__.IsEmpty
except:
    pass

try:
    IsModified = hwp__.IsModified
except:
    pass

try:
    IsPrivateInfoProtected = hwp__.IsPrivateInfoProtected
except:
    pass

try:
    IsTrackChange = hwp__.IsTrackChange
except:
    pass

try:
    IsTrackChangePassword = hwp__.IsTrackChangePassword
except:
    pass

try:
    KeyIndicator = hwp__.KeyIndicator
except:
    pass

try:
    LastCtrl = hwp__.LastCtrl
except:
    pass

try:
    LineSpacingMethod = hwp__.LineSpacingMethod
except:
    pass

try:
    LineWrapType = hwp__.LineWrapType
except:
    pass

try:
    LockCommand = hwp__.LockCommand
except:
    pass

try:
    LunarToSolar = hwp__.LunarToSolar
except:
    pass

try:
    LunarToSolarBySet = hwp__.LunarToSolarBySet
except:
    pass

try:
    MacroState = hwp__.MacroState
except:
    pass

try:
    MailType = hwp__.MailType
except:
    pass

try:
    MetatagExist = hwp__.MetatagExist
except:
    pass

try:
    MiliToHwpUnit = hwp__.MiliToHwpUnit
except:
    pass

try:
    ModifyFieldProperties = hwp__.ModifyFieldProperties
except:
    pass

try:
    ModifyMetatagProperties = hwp__.ModifyMetatagProperties
except:
    pass

try:
    MovePos = hwp__.MovePos
except:
    pass

try:
    MoveToField = hwp__.MoveToField
except:
    pass

try:
    MoveToMetatag = hwp__.MoveToMetatag
except:
    pass

try:
    NumberFormat = hwp__.NumberFormat
except:
    pass

try:
    Numbering = hwp__.Numbering
except:
    pass

try:
    Open = hwp__.Open
except:
    pass

try:
    PageCount = hwp__.PageCount
except:
    pass

try:
    PageNumPosition = hwp__.PageNumPosition
except:
    pass

try:
    PageType = hwp__.PageType
except:
    pass

try:
    ParaHeadAlign = hwp__.ParaHeadAlign
except:
    pass

try:
    ParaShape = hwp__.ParaShape
except:
    pass

try:
    ParentCtrl = hwp__.ParentCtrl
except:
    pass

try:
    Path = hwp__.Path
except:
    pass

try:
    PicEffect = hwp__.PicEffect
except:
    pass

try:
    PlacementType = hwp__.PlacementType
except:
    pass

try:
    PointToHwpUnit = hwp__.PointToHwpUnit
except:
    pass

try:
    PresentEffect = hwp__.PresentEffect
except:
    pass

try:
    PrintDevice = hwp__.PrintDevice
except:
    pass

try:
    PrintPaper = hwp__.PrintPaper
except:
    pass

try:
    PrintRange = hwp__.PrintRange
except:
    pass

try:
    PrintType = hwp__.PrintType
except:
    pass

try:
    ProtectPrivateInfo = hwp__.ProtectPrivateInfo
except:
    pass

try:
    PutFieldText = hwp__.PutFieldText
except:
    pass

try:
    PutMetatagNameText = hwp__.PutMetatagNameText
except:
    pass

try:
    Quit = hwp__.Quit
except:
    pass

try:
    RGBColor = hwp__.RGBColor
except:
    pass

try:
    RegisterModule = hwp__.RegisterModule
except:
    pass

try:
    RegisterPrivateInfoPattern = hwp__.RegisterPrivateInfoPattern
except:
    pass

try:
    ReleaseAction = hwp__.ReleaseAction
except:
    pass

try:
    ReleaseScan = hwp__.ReleaseScan
except:
    pass

try:
    RenameField = hwp__.RenameField
except:
    pass

try:
    RenameMetatag = hwp__.RenameMetatag
except:
    pass

try:
    ReplaceAction = hwp__.ReplaceAction
except:
    pass

try:
    ReplaceFont = hwp__.ReplaceFont
except:
    pass

try:
    Revision = hwp__.Revision
except:
    pass

try:
    Run = hwp__.Run
except:
    pass

try:
    RunScriptMacro = hwp__.RunScriptMacro
except:
    pass

try:
    Save = hwp__.Save
except:
    pass

try:
    SaveAs = hwp__.SaveAs
except:
    pass

try:
    ScanFont = hwp__.ScanFont
except:
    pass

try:
    SelectText = hwp__.SelectText
except:
    pass

try:
    SelectionMode = hwp__.SelectionMode
except:
    pass

try:
    SetBarCodeImage = hwp__.SetBarCodeImage
except:
    pass

try:
    SetCurFieldName = hwp__.SetCurFieldName
except:
    pass

try:
    SetCurMetatagName = hwp__.SetCurMetatagName
except:
    pass

try:
    SetDRMAuthority = hwp__.SetDRMAuthority
except:
    pass

try:
    SetFieldViewOption = hwp__.SetFieldViewOption
except:
    pass

try:
    SetMessageBoxMode = hwp__.SetMessageBoxMode
except:
    pass

try:
    SetPos = hwp__.SetPos
except:
    pass

try:
    SetPosBySet = hwp__.SetPosBySet
except:
    pass

try:
    SetPrivateInfoPassword = hwp__.SetPrivateInfoPassword
except:
    pass

try:
    SetTextFile = hwp__.SetTextFile
except:
    pass

try:
    SetTitleName = hwp__.SetTitleName
except:
    pass

try:
    SetUserInfo = hwp__.SetUserInfo
except:
    pass

try:
    SideType = hwp__.SideType
except:
    pass

try:
    Signature = hwp__.Signature
except:
    pass

try:
    Slash = hwp__.Slash
except:
    pass

try:
    SolarToLunar = hwp__.SolarToLunar
except:
    pass

try:
    SolarToLunarBySet = hwp__.SolarToLunarBySet
except:
    pass

try:
    SortDelimiter = hwp__.SortDelimiter
except:
    pass

try:
    StrikeOut = hwp__.StrikeOut
except:
    pass

try:
    StyleType = hwp__.StyleType
except:
    pass

try:
    SubtPos = hwp__.SubtPos
except:
    pass

try:
    TableBreak = hwp__.TableBreak
except:
    pass

try:
    TableFormat = hwp__.TableFormat
except:
    pass

try:
    TableSwapType = hwp__.TableSwapType
except:
    pass

try:
    TableTarget = hwp__.TableTarget
except:
    pass

try:
    TextAlign = hwp__.TextAlign
except:
    pass

try:
    TextArtAlign = hwp__.TextArtAlign
except:
    pass

try:
    TextDir = hwp__.TextDir
except:
    pass

try:
    TextFlowType = hwp__.TextFlowType
except:
    pass

try:
    TextWrapType = hwp__.TextWrapType
except:
    pass

try:
    UnSelectCtrl = hwp__.UnSelectCtrl
except:
    pass

try:
    VAlign = hwp__.VAlign
except:
    pass

try:
    Version = hwp__.Version
except:
    pass

try:
    VertRel = hwp__.VertRel
except:
    pass

try:
    ViewFlag = hwp__.ViewFlag
except:
    pass

try:
    ViewProperties = hwp__.ViewProperties
except:
    pass

try:
    WatermarkBrush = hwp__.WatermarkBrush
except:
    pass

try:
    WidthRel = hwp__.WidthRel
except:
    pass

try:
    XHwpDocuments = hwp__.XHwpDocuments
except:
    pass

try:
    XHwpMessageBox = hwp__.XHwpMessageBox
except:
    pass

try:
    XHwpODBC = hwp__.XHwpODBC
except:
    pass

try:
    XHwpWindows = hwp__.XHwpWindows
except:
    pass