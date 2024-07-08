import json
import os
import re
import shutil
import sys
import tempfile
import urllib.error
import zipfile
from collections import defaultdict
from io import StringIO
from time import sleep
from typing import Literal, Union
from urllib import request, parse

import numpy as np
import pandas as pd
import pyperclip as cb
import pythoncom
import win32com.client as win32
from PIL import Image

__version__ = "0.18.0"

# for pyinstaller
if getattr(sys, 'frozen', False):
    pyinstaller_path = sys._MEIPASS
else:
    pyinstaller_path = os.path.dirname(os.path.abspath(__file__))

# temp 폴더 삭제
try:
    shutil.rmtree(os.path.join(os.environ["USERPROFILE"], "AppData/Local/Temp/gen_py"))
except FileNotFoundError as e:
    pass

# Type Library 파일 재생성
win32.gencache.EnsureModule('{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}', 0, 1, 0)


# 헬퍼함수
def rename_duplicates_in_list(file_list):
    """
    문서 내 이미지를 파일로 저장할 때,
    동일한 이름의 파일 뒤에 (2), (3).. 붙여주는 함수
    """
    # 딕셔너리를 사용하여 중복 횟수를 추적합니다.
    counts = {}

    # 리스트를 순회하며 중복 횟수를 계산합니다.
    for i, item in enumerate(file_list):
        # 중복된 아이템을 찾았을 경우
        if item in counts:
            counts[item] += 1
            new_item = f"{os.path.splitext(item)[0]}({counts[item]}){os.path.splitext(item)[1]}"
        else:
            counts[item] = 0
            new_item = item

        # 리스트를 업데이트합니다.
        file_list[i] = new_item

    return file_list


def check_tuple_of_ints(var):
    if isinstance(var, tuple):  # 먼저 변수가 튜플인지 확인
        return all(isinstance(item, int) for item in var)  # 모든 요소가 int인지 확인
    return False  # 변수가 튜플이 아니면 False 반환


def excel_address_to_tuple_zero_based(address):
    column = 0
    row = 0
    for char in address:
        if char.isalpha():
            column = column * 26 + (ord(char.upper()) - ord('A'))
        elif char.isdigit():
            row = row * 10 + int(char)
        else:
            raise ValueError("Invalid address format")
    return row - 1, column


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
        # return "<파이썬+아래아한글 자동화를 돕기 위한 함수모음 및 추상화 인스턴스 by 일상의코딩>"
        return """
        파이썬으로 한/글 오토메이션API를 간편하게 사용하기 위한 메서드와 속성을 제공하는 클래스입니다.
        파이썬 콘솔 또는 쥬피터 노트북에서 아래와 같이 실행하여 한/글을 실행할 수 있습니다.

        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()

        코드 실행 전에 한/글이 실행되어 있었다면 가장 최근에 조작(또는 포커스)했던 한/글 창에 연결됩니다. 
        """

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
                    self.hwp = win32.gencache.EnsureDispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch))  # 그이후는 오토메이션 api를 사용할수 있습니다
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
    def Application(self):
        """
        저수준의 아래아한글 오토메이션API에 직접 접근하기 위한 속성
        사용예시는 아래와 같다.

        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()
        >>> hwp.Application.XHwpWindows.Item(0).Visible = True
        :return: HwpApplication 객체
        """
        return self.hwp.Application

    @property
    def CellShape(self):
        """
        셀(또는 표) 모양을 관리하는 파라미터셋 속성
        :return:
        """
        return self.hwp.CellShape

    @CellShape.setter
    def CellShape(self, prop):
        """
        셀(또는 표) 모양 파라미터셋을 변경할 수 있는 setter 속성
        :param prop:
        :return:
        """
        self.hwp.CellShape = prop

    @property
    def CharShape(self):
        """
        글자모양 파라미터셋을 조회할 수 있는 파라미터셋 속성
        :return:
        """
        return self.hwp.CharShape

    @CharShape.setter
    def CharShape(self, prop):
        """
        글자모양 파라미터셋을 변경할 수 있는 setter 속성
        :param prop:
        :return:
        """
        self.hwp.CharShape = prop

    @property
    def CLSID(self):
        """
        클래스아이디를 리턴하는 속성
        :return:
        """
        return self.hwp.CLSID

    @property
    def coclass_clsid(self):
        """
        coclass의 clsid를 리턴하는 속성
        :return:
        """
        return self.hwp.coclass_clsid

    @property
    def CurFieldState(self):
        """
        현재 캐럿이 들어가 있는 필드의 상태를 조회할 수 있는 속성
        :return:
        """
        return self.hwp.CurFieldState

    @property
    def CurMetatagState(self):
        """
        현재 캐럿이 들어가 있는 메타태그 상태를 조회할 수 있는 속성
        :return:
        """
        return self.hwp.CurMetatagState

    @property
    def CurSelectedCtrl(self):
        """
        현재 선택된 컨트롤을 리턴하는 속성
        :return:
        """
        return self.hwp.CurSelectedCtrl

    @property
    def EditMode(self):
        """
        현재 편집모드를 리턴하는 속성
        :return:
        """
        return self.hwp.EditMode

    @EditMode.setter
    def EditMode(self, prop):
        """
        현재 편집모드를 수정하기 위한 setter 속성
        :param prop:
        :return:
        """
        self.hwp.EditMode = prop

    @property
    def EngineProperties(self):
        return self.hwp.EngineProperties

    @property
    def HAction(self):
        """
        한/글의 액션을 설정하고 실행하기 위한 속성.
        GetDefalut, Execute, Run 등의 메서드를 가지고 있다.
        :return:
        """
        return self.hwp.HAction

    @property
    def HeadCtrl(self):
        """
        문서의 첫 번째 컨트롤을 리턴한다.
        거의 모든 경우 HeadCtrl은 구역 정의(Section Definition, secd)를 리턴한다.
        사용법은 아래와 같다.

        >>> # 문서에 첫 번째로 삽입된 표의 컨트롤을 탐색하여 선택하는 방법
        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()
        >>> ctrl = hwp.HeadCtrl
        >>> while True:
        >>>     if ctrl.UserDesc == "표":
        >>>         break
        >>>     ctrl = ctrl.Next
        >>> print("표가 선택되었습니다.")
        """
        return self.hwp.HeadCtrl

    @property
    def HParameterSet(self):
        """
        한/글에서 실행되는 대부분의 액션에 필요한
        다양한 파라미터셋을 제공해주는 속성.
        사용법은 아래와 같다.

        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()
        >>> pset = hwp.HParameterSet.HInsertText
        >>> pset.Text = "Hello world!"
        >>> hwp.HAction.Execute("InsertText", pset.HSet)

        :return:
        """
        return self.hwp.HParameterSet

    @property
    def IsEmpty(self) -> bool:
        """
        아무 내용도 들어있지 않은 빈 문서인지 여부를 나타낸다. 읽기전용
        """
        return self.hwp.IsEmpty

    @property
    def IsModified(self) -> bool:
        """
        최근 저장 또는 생성 이후 수정이 있는지 여부를 나타낸다. 읽기전용
        """
        return self.hwp.IsModified

    @property
    def IsPrivateInfoProtected(self):
        return self.hwp.IsPrivateInfoProtected

    @property
    def IsTrackChangePassword(self):
        return self.hwp.IsTrackChangePassword

    @property
    def IsTrackChange(self):
        return self.hwp.IsTrackChange

    @property
    def LastCtrl(self):
        """
        문서의 가장 마지막 컨트롤 객체를 리턴한다.
        연결리스트 타입이므로, HeadCtrl부터 LastCtrl까지 모두 연결되어 있고
        LastCtrl.Prev.Prev 또는 HeadCtrl.Next.Next 등으로 컨트롤 순차 탐색이 가능하다.
        :return:
        """
        return self.hwp.LastCtrl

    @property
    def PageCount(self):
        """
        현재 문서의 총 페이지 수를 리턴한다.
        :return:
        """
        return self.hwp.PageCount

    @property
    def ParaShape(self):
        """
        현재 캐럿이 위치한 문단의 문단모양 파라미터셋을 리턴하는 속성.
        :return:
        """
        return self.hwp.ParaShape

    @ParaShape.setter
    def ParaShape(self, prop):
        """
        문단모양 파라미터셋을 수정하기 위한 세터 프로퍼티
        :param prop:
        :return:
        """
        self.hwp.ParaShape = prop

    @property
    def ParentCtrl(self):
        """
        현재 선택되어 있거나, 캐럿이 들어있는 컨트롤을 포함하는 상위 컨트롤을 리턴한다.
        :return:
        """
        return self.hwp.ParentCtrl

    @property
    def Path(self):
        """
        현재 빈 문서가 아닌 경우, 열려 있는 문서의 파일명을 포함한 전체경로를 리턴한다.
        :return:
        """
        return self.hwp.Path

    @property
    def SelectionMode(self):
        """
        현재 선택모드가 어떤 상태인지 리턴한다.
        :return:
        """
        return self.hwp.SelectionMode

    @property
    def Version(self):
        """
        아래아한글 프로그램의 버전을 문자열로 리턴한다.
        :return:
        """
        return self.hwp.Version

    @property
    def ViewProperties(self):
        """
        현재 한/글 프로그램의 보기 속성 파라미터셋을 리턴한다.
        :return:
        """
        return self.hwp.ViewProperties

    @ViewProperties.setter
    def ViewProperties(self, prop):
        """
        현재 한/글 프로그램의 보기 속성 파라미터셋을 수정하는 세터 프로퍼티.
        :param prop:
        :return:
        """
        self.hwp.ViewProperties = prop

    @property
    def XHwpDocuments(self):
        """
        HwpApplication의 XHwpDocuments 객체를 리턴한다.
        :return:
        """
        return self.hwp.XHwpDocuments

    @property
    def XHwpMessageBox(self):
        """
        메시지박스 객체 리턴
        :return:
        """
        return self.hwp.XHwpMessageBox

    @property
    def XHwpODBC(self):
        return self.hwp.XHwpODBC

    @property
    def XHwpWindows(self):
        return self.hwp.XHwpWindows

    @property
    def ctrl_list(self):
        """
        문서 내 모든 ctrl를 리스트로 반환한다.
        단, 기본으로 삽입되고 선택 불가능한
        두 개의 컨트롤인 secd(섹션정의)와 cold(단정의) 두 개는
        ctrl_list에서 제외했다.
        (모든 컨트롤을 제거하는 등의 경우 편의를 위함)
        :return:
        """
        c_list = []
        ctrl = self.hwp.HeadCtrl.Next.Next
        while ctrl:
            c_list.append(ctrl)
            ctrl = ctrl.Next
        return c_list

    @property
    def current_page(self):
        """
        현재 페이지 번호를 리턴.
        1페이지에 있다면 1을 리턴한다.
        :return:
        """
        return self.KeyIndicator()[3]

    # 커스텀 메서드
    def get_selected_range(self):
        """
        선택한 범위의 셀주소를
        리스트로 리턴함
        """
        if not self.is_cell():
            raise AttributeError("캐럿이 표 안에 있어야 합니다.")
        pset = self.HParameterSet.HFieldCtrl
        self.HAction.GetDefault("TableFormula", pset.HSet)
        return pset.Command[2:-1].split(",")

    def fill_addr_field(self):
        if not self.is_cell():
            raise AttributeError("캐럿이 표 안에 있어야 합니다.")
        self.TableColBegin()
        self.TableColPageUp()
        self.set_cur_field_name("A1")
        while self.TableRightCell():
            self.set_cur_field_name(self.get_cell_addr())

    def unfill_addr_field(self):
        if not self.is_cell():
            raise AttributeError("캐럿이 표 안에 있어야 합니다.")
        self.TableColBegin()
        self.TableColPageUp()
        self.set_cur_field_name("")
        while self.TableRightCell():
            self.set_cur_field_name("")

    def resize_image(self, width:int=None, height:int=None, unit:Literal["mm", "hwpunit"]="mm"):
        """
        이미지 또는 그리기 개체의 크기를 조절하는 메서드.
        해당개체 선택 후 실행해야 함.
        """
        self.FindCtrl()
        prop = self.CurSelectedCtrl.Properties
        if width:
            prop.SetItem("Width", width if unit=="hwpunit" else self.MiliToHwpUnit(width))
        if height:
            prop.SetItem("Height", height if unit=="hwpunit" else self.MiliToHwpUnit(height))
        if width or height:
            self.CurSelectedCtrl.Properties = prop
            return True
        return False

    def save_image(self, path="./img.png", ctrl=""):
        path = os.path.abspath(path)
        if os.path.exists(path):
            raise FileExistsError("해당 이름의 파일이 이미 존재합니다.")
        if ctrl:
            self.move_to_ctrl(ctrl)
        self.find_ctrl()
        if not self.CurSelectedCtrl.CtrlID == "gso":
            return False
        pset = self.HParameterSet.HShapeObjSaveAsPicture
        self.HAction.GetDefault("PictureSave", pset.HSet)
        pset.Path = path
        pset.Ext = "BMP"
        try:
            self.HAction.Execute("PictureSave", pset.HSet)
        finally:
            self.Undo()
            if not os.path.exists(path):
                file_list = os.listdir(os.path.dirname(path))
                path_list = [os.path.join(os.path.dirname(path), i) for i in file_list]
                temp_file = \
                    sorted([i for i in path_list if i.startswith(os.path.splitext(path))], key=os.path.getmtime)[-1]
                Image.open(temp_file).save(path)
                os.remove(temp_file)
            print(f"image saved to {path}")

    def NewNumberModify(self, new_number: int,
                        num_type: Literal["Page", "Figure", "Footnote", "Table", "Endnote", "Equation"] = "Page"):
        """
        새 번호 조판을 수정할 수 있는 메서드.
        실행 전 [새 번호] 조판 옆에 캐럿이 위치해 있어야 하며,
        그렇지 않을 경우
        (쪽번호 외에도 그림, 각주, 표, 미주, 수식 등)
        다만, 주의할 점이 세 가지 있다.
        1. 기존에 쪽번호가 없는 문서에서는 작동하지 않으므로
           쪽번호가 정의되어 있어야 한다.
           (쪽번호 정의는 PageNumPos 메서드 참조)
        2. 새 번호를 지정한 페이지 및 이후 모든 페이지가
           영향을 받는다.
        3. NewNumber 실행시점의 캐럿위치 뒤쪽(해당 페이지 내)에
           NewNumber 조판이 있는 경우, 삽입한 조판은 무효가 된다.
           (페이지 맨 뒤쪽의 새 번호만 유효함)
        Todo: 페이지 내에 캐럿 뒤쪽으로 [새번호]조판이 있는 경우 지워버리기

        :param new_number:
            새 번호
        :param num_type:
            타입 지정
            "Page": 쪽(기본값)
            "Figure": 그림
            "Footnote": 각주
            "Table": 표
            "Endnote": 미주
            "Equation": 수식
        :return:
            성공시 True, 실패시 False를 리턴
        """
        current_pos = self.GetPos()
        current_page = self.PageCount
        ctrl_name = self.FindCtrl()
        if ctrl_name != "nwno" or self.PageCount != current_page:
            self.SetPos(*current_pos)
            return False
        pset = self.HParameterSet.HAutoNum
        self.HAction.GetDefault("NewNumberModify", pset.HSet)
        pset.NumType = self.AutoNumType(num_type)
        pset.NewNumber = new_number
        return self.HAction.Execute("NewNumberModify", pset.HSet)

    def NewNumber(self, new_number: int,
                  num_type: Literal["Page", "Figure", "Footnote", "Table", "Endnote", "Equation"] = "Page"):
        """
        새 번호를 매길 수 있는 메서드.
        (쪽번호 외에도 그림, 각주, 표, 미주, 수식 등)
        다만, 주의할 점이 세 가지 있다.
        1. 기존에 쪽번호가 없는 문서에서는 작동하지 않으므로
           쪽번호가 정의되어 있어야 한다.
           (쪽번호 정의는 PageNumPos 메서드 참조)
        2. 새 번호를 지정한 페이지 및 이후 모든 페이지가
           영향을 받는다.
        3. NewNumber 실행시점의 캐럿위치 뒤쪽(해당 페이지 내)에
           NewNumber 조판이 있는 경우, 삽입한 조판은 무효가 된다.
           (페이지 맨 뒤쪽의 새 번호만 유효함)

        :param new_number:
            새 번호
        :param num_type:
            타입 지정
            "Page": 쪽(기본값)
            "Figure": 그림
            "Footnote": 각주
            "Table": 표
            "Endnote": 미주
            "Equation": 수식
        :return:
            성공시 True, 실패시 False를 리턴
        """
        pset = self.HParameterSet.HAutoNum
        self.HAction.GetDefault("NewNumber", pset.HSet)
        pset.NumType = self.AutoNumType(num_type)
        pset.NewNumber = new_number
        return self.HAction.Execute("NewNumber", pset.HSet)

    def PageNumPos(self, global_start: int = 1, position: Literal[
        "TopLeft", "TopCenter", "TopRight", "BottomLeft", "BottomCenter", "BottomRight", "InsideTop", "OutsideTop", "InsideBottom", "OutsideBottom", "None"] = "BottomCenter",
                   number_format: Literal[
                       "Digit", "CircledDigit", "RomanCapital", "RomanSmall", "LatinCapital", "HangulSyllable", "Ideograph", "DecagonCircle", "DecagonCircleHanja"] = "Digit",
                   side_char=True):
        """
        문서 전체에 쪽번호를 삽입하는 메서드.
        :param global_start:
            시작번호를 지정할 수 있음(새 번호 아님. 새 번호는 hwp.NewNumber(n)을 사용할 것)
        :param position:
            쪽번호 위치를 지정하는 파라미터
            TopLeft, TopCenter, TopRight
            BottomLeft, BottomCenter(기본값), BottomRight
            InsideTop, OutsideTop, InsideBottom, OutsideBottom
            None(쪽번호숨김과 유사)
        :param number_format:
            쪽번호 서식을 지정하는 파라미터
	        "Digit": (1 2 3),
	        "CircledDigit": (① ② ③),
	        "RomanCapital":(I II III),
	        "RomanSmall": (i ii iii) ,
	        "LatinCapital": (A B C),
	        "HangulSyllable":(가 나 다),
	        "Ideograph": (一 二 三),
	        "DecagonCircle": (갑 을 병),
	        "DecagonCircleHanja": (甲 乙 丙),
        :param side_char:
            줄표 삽입 여부(bool)
            True : 줄표 삽입(기본값)
            False : 줄표 삽입하지 않음
        :return:
            성공시 True, 실패시 False를 리턴
        """
        pset = self.HParameterSet.HPageNumPos
        self.HAction.GetDefault("PageNumPos", pset.HSet)
        pset.DrawPos = self.PageNumPosition(position)
        pset.NumberFormat = self.NumberFormat(number_format)
        pset.NewNumber = global_start
        if side_char:
            pset.SideChar = 45
        else:
            pset.SideChar = 0
        return self.HAction.Execute("PageNumPos", pset.HSet)

    def table_to_string(self, rowsep="", colsep="\r\n"):
        if not self.is_cell():
            raise AssertionError("캐럿이 표 안에 있지 않습니다.")

        def extract_content_from_table(sep):
            pset = self.HParameterSet.HTableTblToStr
            self.HAction.GetDefault("TableTableToString", pset.HSet)
            # <pset.DelimiterType>
            # 0: hwp.Delimiter("Tab")
            # 1: hwp.Delimiter("SemiBreve") 콤마
            # 2: hwp.Delimiter("Space")
            # 3: hwp.Delimiter("LineSep")
            pset.DelimiterType = self.Delimiter("LineSep")
            pset.UserDefine = sep
            self.HAction.Execute("TableTableToString", pset.HSet)

        self.TableColPageDown()
        i = 1
        while self.TableSplitTable():
            self.MoveUp()
            i += 1
        for i in range(i):
            extract_content_from_table(colsep)
            self.insert_text(rowsep)
            self.SelectCtrlFront()

    def get_table_height(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"):
        """
        현재 캐럿이 속한 표의 너비(mm)를 리턴함
        :return: 표의 너비(mm)
        """
        if as_.lower() == "mm":
            return self.HwpUnitToMili(self.CellShape.Item("Height"))
        elif as_.lower() in ("hwpunit", "hu"):
            return self.CellShape.Item("Height")
        elif as_.lower() in ("point", "pt"):
            return self.HwpUnitToPoint(self.CellShape.Item("Height"))
        elif as_.lower() == "inch":
            return self.HwpUnitToInch(self.CellShape.Item("Height"))
        else:
            raise KeyError("mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다.")

    def get_row_num(self):
        """
        캐럿이 표 안에 있을 때,
        현재 표의 행의 갯수를 리턴
        (일부 행병합이 있는 경우, 최대 행번호를 리턴)
        * 단, 최대 행갯수가 최대 행번호와 다른 경우가 있으므로 유의할 것 *
        :return:
        """
        if not self.is_cell():
            raise AssertionError("현재 캐럿이 표 안에 있지 않습니다.")
        cur_pos = self.get_pos()
        self.TableColBegin()
        self.TableColPageDown()
        max_row_num = int(self.KeyIndicator()[-1][1:].split(")")[0][1:])
        while self.TableRightCell():
            max_row_num = max(max_row_num, int(self.KeyIndicator()[-1][1:].split(")")[0][1:]))
        self.set_pos(*cur_pos)
        return max_row_num

    def get_row_height(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"):
        """
        표 안에서 캐럿이 들어있는 행(row)의 높이를 리턴함.
        기본단위는 mm 이지만, HwpUnit이나 Point 등 보다 작은 단위를 사용할 수 있다.
        (메서드 내부에서는 HwpUnit으로 연산한다.)
        :param as_: 리턴하는 수치의 단위
        :return: 캐럿이 속한 행의 높이
        """
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        if as_.lower() == "mm":
            return self.HwpUnitToMili(pset.ShapeTableCell.Height)
        elif as_.lower() in ("hwpunit", "hu"):
            return pset.ShapeTableCell.Height
        elif as_.lower() in ("point", "pt"):
            return self.HwpUnitToPoint(pset.ShapeTableCell.Height)
        elif as_.lower() == "inch":
            return self.HwpUnitToInch(pset.ShapeTableCell.Height)
        else:
            raise KeyError("mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다.")

    def get_col_num(self):
        """
        캐럿이 표 안에 있을 때,
        현재 셀의 열번호, 즉 셀주소 문자열의 정수 부분을 리턴
        :return:
        """
        if not self.is_cell():
            raise AssertionError("현재 캐럿이 표 안에 있지 않습니다.")
        cur_pos = self.get_pos()
        self.TableColPageUp()
        self.TableColEnd()
        try:
            return ord(self.KeyIndicator()[-1][1:].split(")")[0][0]) - 64
        finally:
            self.set_pos(*cur_pos)

    def get_col_width(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"):
        """
        현재 캐럿이 위치한 셀(칼럼)의 너비를 리턴하는 메서드.
        기본 단위는 mm이지만, as_ 파라미터를 사용하여 단위를 hwpunit이나 point, inch 등으로 변경 가능하다.
        :param as_: 리턴값의 단위(mm, HwpUnit, Pt, Inch 등 4종류)
        :return:
        """
        if not self.is_cell():
            raise AssertionError("현재 캐럿이 표 안에 있지 않습니다.")
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        if as_.lower() == "mm":
            return self.HwpUnitToMili(pset.ShapeTableCell.Width)
        elif as_.lower() in ("hwpunit", "hu"):
            return pset.ShapeTableCell.Width
        elif as_.lower() in ("point", "pt"):
            return self.HwpUnitToPoint(pset.ShapeTableCell.Width)
        elif as_.lower() == "inch":
            return self.HwpUnitToInch(pset.ShapeTableCell.Width)
        else:
            raise KeyError("mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다.")

    def set_col_width(self, width: int | float | list | tuple, as_: Literal["mm", "ratio"] = "ratio"):
        """
        칼럼의 너비를 변경할 수 있는 메서드.
        정수(int)나 부동소수점수(float) 입력시 현재 칼럼의 너비가 변경되며,
        리스트나 튜플 등 iterable 타입 입력시에는 각 요소들의 비에 따라 칼럼들의 너비가 일괄변경된다.
        예를 들어 3행 3열의 표 안에서 set_col_width([1,2,3]) 을 실행하는 경우
        1열너비:2열너비:3열너비가 1:2:3으로 변경된다.
        (표 전체의 너비가 148mm라면, 각각 24mm : 48mm : 72mm로 변경된다는 뜻이다.)

        단, 열너비의 비가 아닌 "mm" 단위로 값을 입력하려면 as_="mm"로 파라미터를 수정하면 된다.
        이 때, width에 정수 또는 부동소수점수를 입력하는 경우 as_="ratio"를 사용할 수 없다.

        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()
        >>> hwp.create_table(3,3)
        >>> hwp.get_into_nth_table(0)
        >>> hwp.set_col_width([1,2,3])
        :param width: 열 너비
        :param as_:
        :return:
        """
        cur_pos = self.get_pos()
        if type(width) in (int, float):
            if as_ == "ratio":
                raise TypeError('width에 int나 float 입력시 as_ 파라미터는 "mm"로 설정해주세요.')
            self.TableColPageUp()
            self.TableCellBlock()
            self.TableCellBlockExtend()
            self.TableColPageDown()
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 1)
            pset.ShapeTableCell.Width = self.MiliToHwpUnit(width)
            try:
                return self.HAction.Execute("TablePropertyDialog", pset.HSet)
            finally:
                self.set_pos(*cur_pos)
        else:
            if as_ == "ratio":
                table_width = self.get_table_width()
                width = [i / sum(width) * table_width for i in width]
            self.TableColBegin()
            for i in width:
                self.TableColPageUp()
                self.TableCellBlock()
                self.TableCellBlockExtend()
                self.TableColPageDown()
                pset = self.HParameterSet.HShapeObject
                self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                pset.HSet.SetItem("ShapeType", 3)
                pset.HSet.SetItem("ShapeCellSize", 1)
                pset.ShapeTableCell.Width = self.MiliToHwpUnit(i)
                self.HAction.Execute("TablePropertyDialog", pset.HSet)
                self.TableRightCell()
            return self.set_pos(*cur_pos)

    def adjust_cellwidth(self, width: int | float | list | tuple, as_: Literal["mm", "ratio"] = "ratio"):
        """
        칼럼의 너비를 변경할 수 있는 메서드.
        정수(int)나 부동소수점수(float) 입력시 현재 칼럼의 너비가 변경되며,
        리스트나 튜플 등 iterable 타입 입력시에는 각 요소들의 비에 따라 칼럼들의 너비가 일괄변경된다.
        예를 들어 3행 3열의 표 안에서 set_col_width([1,2,3]) 을 실행하는 경우
        1열너비:2열너비:3열너비가 1:2:3으로 변경된다.
        (표 전체의 너비가 148mm라면, 각각 24mm : 48mm : 72mm로 변경된다는 뜻이다.)

        단, 열너비의 비가 아닌 "mm" 단위로 값을 입력하려면 as_="mm"로 파라미터를 수정하면 된다.
        이 때, width에 정수 또는 부동소수점수를 입력하는 경우 as_="ratio"를 사용할 수 없다.

        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()
        >>> hwp.create_table(3,3)
        >>> hwp.get_into_nth_table(0)
        >>> hwp.adjust_cellwidth([1,2,3])
        :param width: 열 너비
        :param as_:
        :return:
        """
        cur_pos = self.get_pos()
        if type(width) in (int, float):
            if as_ == "ratio":
                raise TypeError('width에 int나 float 입력시 as_ 파라미터는 "mm"로 설정해주세요.')
            self.TableColPageUp()
            self.TableCellBlock()
            self.TableCellBlockExtend()
            self.TableColPageDown()
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 1)
            pset.ShapeTableCell.Width = self.MiliToHwpUnit(width)
            try:
                return self.HAction.Execute("TablePropertyDialog", pset.HSet)
            finally:
                self.set_pos(*cur_pos)
        else:
            if as_ == "ratio":
                table_width = self.get_table_width()
                width = [i / sum(width) * table_width for i in width]
            self.TableColBegin()
            for i in width:
                self.TableColPageUp()
                self.TableCellBlock()
                self.TableCellBlockExtend()
                self.TableColPageDown()
                pset = self.HParameterSet.HShapeObject
                self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                pset.HSet.SetItem("ShapeType", 3)
                pset.HSet.SetItem("ShapeCellSize", 1)
                pset.ShapeTableCell.Width = self.MiliToHwpUnit(i)
                self.HAction.Execute("TablePropertyDialog", pset.HSet)
                self.TableRightCell()
            return self.set_pos(*cur_pos)

    def get_table_width(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"):
        """
        현재 캐럿이 속한 표의 너비(mm)를 리턴함.
        이 때 수치의 단위는 as_ 파라미터를 통해 변경 가능하며, "mm", "HwpUnit", "Pt", "Inch" 등을 쓸 수 있다.
        :return: 표의 너비(mm)
        """
        if as_.lower() == "mm":
            return self.HwpUnitToMili(self.CellShape.Item("Width"))
        elif as_.lower() in ("hwpunit", "hu"):
            return self.CellShape.Item("Width")
        elif as_.lower() in ("point", "pt"):
            return self.HwpUnitToPoint(self.CellShape.Item("Width"))
        elif as_.lower() == "inch":
            return self.HwpUnitToInch(self.CellShape.Item("Width"))
        else:
            raise KeyError("mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다.")

    def set_table_width(self, width: int = 0, as_: Literal["mm", "hwpunit", "hu"] = "mm"):
        """
        표 전체의 너비를 원래 열들의 비율을 유지하면서 조정하는 메서드.
        :param width: 너비(단위는 기본 mm이며, hwpunit으로 변경 가능)
        :param as_: 단위("mm" or "hwpunit")
        :return: 성공시 True
        """
        cur_pos = self.get_pos()
        while self.TableRightCell():
            if not self.get_cell_addr().endswith("1"):
                break

        if as_.lower() in ("hwpunit", "hu"):
            width = self.hwp_unit_to_mili(width)

        self.TableColBegin()
        if not width:
            sec_def = self.hwp.HParameterSet.HSecDef
            self.hwp.HAction.GetDefault("PageSetup", sec_def.HSet)
            width = sec_def.PageDef.PaperWidth - sec_def.PageDef.LeftMargin - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen - self.mili_to_hwp_unit(
                2)
            if as_ == "mm":
                width = self.HwpUnitToMili(width)
        table_width = self.get_table_width(as_=as_)
        cur_col_widths = []
        col_num = self.get_col_num()
        for i in range(col_num):
            cur_col_widths.append(self.get_col_width(as_=as_))
            self.TableRightCell()

        dst_col_widths = [i / table_width * width for i in cur_col_widths]

        self.TableColBegin()
        for i in dst_col_widths:
            self.TableColPageUp()
            self.TableCellBlock()
            self.TableCellBlockExtend()
            self.TableColPageDown()
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 1)
            pset.ShapeTableCell.Width = self.MiliToHwpUnit(i)
            self.HAction.Execute("TablePropertyDialog", pset.HSet)
            self.TableRightCell()
        return self.set_pos(*cur_pos)

    def save_pdf_as_image(self, path: str = "", img_format="bmp"):
        """
        문서보안이나 복제방지를 위해
        모든 페이지를 이미지로 변경 후
        PDF로 저장하는 메서드.
        아무 인수가 주어지지 않는 경우
        모든 페이지를 bmp로 저장한 후에
        현재 폴더에 {문서이름}.pdf로 저장한다.
        (만약 저장하지 않은 빈 문서의 경우에는 result.pdf로 저장한다.)

        :param path: 저장경로 및 파일명
        :param img_format: 이미지 변환 포맷
        :return:
        """
        if path == "" and self.Path:
            path = self.Path.rsplit(".hwp", maxsplit=1)[0] + ".pdf"
        elif path == "" and self.Path == "":
            path = os.path.abspath("result.pdf")
        else:
            if not os.path.exists(os.path.dirname(os.path.abspath(path))):
                os.mkdir(os.path.dirname(os.path.abspath(path)))
        temp_dir = tempfile.mkdtemp()
        self.create_page_image(os.path.join(temp_dir, f"img.{img_format}"))
        img_list = [os.path.join(temp_dir, i) for i in os.listdir(temp_dir)]
        img_list = [Image.open(i).convert("RGB") for i in img_list]
        img_list[0].save(path, save_all=True, append_images=img_list[1:])
        shutil.rmtree(temp_dir)
        return True

    def get_cell_addr(self, as_: Literal["str", "tuple"] = "str"):
        """
        현재 캐럿이 위치한 셀의 주소를 "A1" 또는 (0, 0)으로 리턴.
        캐럿이 표 안에 있지 않은 경우 False를 리턴함
        :param as_:
            "str"의 경우 엑셀처럼 "A1" 방식으로 리턴,
            "tuple"인 경우 (0,0) 방식으로 리턴.
        :return:
        """
        if not self.hwp.CellShape:
            return False
        result = self.KeyIndicator()[-1][1:].split(")")[0]
        if as_ == "str":
            return result
        else:
            return excel_address_to_tuple_zero_based(result)

    def save_all_pictures(self, save_path="./binData"):
        """
        현재 문서에 삽입된 모든 이미지들을
        삽입 당시 파일명으로 복원하여 저장.
        단, 문서 안에서 복사했거나 중복삽입한 이미지는 한 개만 저장됨.
        기본 저장폴더명은 ./binData이며
        기존에 save_path가 존재하는 경우,
        그 안의 파일들은 삭제되므로 유의해야 함.

        :param save_path:
            저장할 하위경로 이름
        :return:
        """
        current_path = self.Path
        if not current_path:
            raise FileNotFoundError("저장 후 진행해주시기 바랍니다.")
        self.save_as("temp.zip", format="HWPX")
        self.open(current_path)

        with zipfile.ZipFile("./temp.zip", 'r') as zf:
            zf.extractall(path="./temp")
        os.remove("./temp.zip")
        try:
            os.rename("./temp/binData", save_path)
        except FileExistsError:
            shutil.rmtree(save_path)
            os.rename("./temp/binData", save_path)
        with open("./temp/Contents/section0.xml", encoding="utf-8") as f:
            content = f.read()
        bin_list = re.findall(r'원본 그림의 이름: (.*?\..+?)\n', content)
        bin_list = rename_duplicates_in_list(bin_list)
        os.chdir(save_path)
        file_list = os.listdir()
        for i in file_list:
            idx = re.findall(r"\d+", i)[0]
            os.rename(i, i.replace(idx, f"{int(idx):04}"))

        for i, j in zip(os.listdir(), bin_list):
            os.rename(i, j)
        os.chdir("..")
        shutil.rmtree("./temp")
        return True

    def select_ctrl(self, ctrl, anchor_type:Literal[0,1,2]=0):
        """
        인수로 넣은 컨트롤 오브젝트를 선택하는 메서드.
        :param ctrl:
            선택하고자 하는 컨트롤
        :param anchor_type:
            컨트롤의 위치를 찾아갈 때 List, Para, Pos의 기준위치.
            (아주 특수한 경우를 제외하면 기본값을 쓰면 된다.)
            0: 바로 상위 리스트에서의 좌표(기본값)
            1: 탑레벨 리스트에서의 좌표
            2: 루트 리스트에서의 좌표
        :return:
        """
        cur_view_state = self.ViewProperties.Item("OptionFlag")
        if cur_view_state not in (2, 6):
            prop = self.ViewProperties
            prop.SetItem("OptionFlag", 6)
            self.ViewProperties = prop

        self.set_pos_by_set(ctrl.GetAnchorPos(anchor_type))
        try:
            if not self.SelectCtrlFront():
                return self.SelectCtrlReverse()
            else:
                return True
        finally:
            prop = self.ViewProperties
            prop.SetItem("OptionFlag", cur_view_state)
            self.ViewProperties = prop

    def move_to_ctrl(self, ctrl):
        """
        인수로 넣은 컨트롤 오브젝트의 조판 앞으로 이동하는 메서드
        :param ctrl:
        :return:
        """
        return self.set_pos_by_set(ctrl.GetAnchorPos(0))

    def set_visible(self, visible):
        """
        현재 조작중인 한/글 인스턴스의 백그라운드 숨김여부를 변경할 수 있다.

        :param visible:
            visible=False로 설정하면 현재 조작중인 한/글 인스턴스가 백그라운드로 숨겨진다.

        :return:
        """
        self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible

    def auto_spacing(self, init_spacing=0, init_ratio=100, max_spacing=40, min_spacing=40, verbose=True):
        """
        자동 자간조정 메서드(beta)
        라인 끝에 단어가 a와 b로 잘려 있는 경우 a>b인 경우 라인의 자간을 줄이고, a<b인 경우 자간을 넓혀
        잘린 단어가 합쳐질 때까지 자간조정을 계속한다.
        단, max_spacing이나 min_spacing을 넘어야 하는 경우에는 원상태로 되돌린 후
        해당 라인의 정보를 콘솔에 출력한다.
        (아주 너비가 작은 셀이나 글상자 등에서는 제대로 작동하지 않을 수 있음.)

        init_spacing과 init_ratio 파라미터를 통해
        자동자간조정을 실행하기 전에 모든 문서의 기본 자간장평을 설정할 수 있다.
        """

        def reset_para_spacing(init_spacing=init_spacing, init_ratio=init_ratio):
            self.MoveListEnd()
            self.MoveSelListBegin()
            try:
                return self.set_font(Spacing=init_spacing, Ratio=init_ratio)
            finally:
                self.Cancel()

        def get_spacing_dir():
            line_num = self.key_indicator()[5]
            self.MoveLineEnd()
            mid = self.get_pos()[2]
            self.Select()
            self.Select()
            _, _, _, head, _, _, tail = self.get_selected_pos()
            self.Cancel()
            if self.key_indicator()[5] == line_num:
                self.MoveNextChar()  # hwp.MoveNextParaBegin() 실행시 줄바꿈 엔터로 나뉜 구간을 무시해버림
            if (mid - head == 0) or (tail - mid == 0):  # 둘 중 하나가 0이면(잘렸으면)
                return 0  # 패스
            elif (mid - head) > (tail - mid):  # 앞이 뒤보다 길면
                return -1  # 자간 줄임
            else:  # 뒤가 더 길면
                return 1  # 자간 늘임

        def select_spacing_area(direction=0):
            if direction == 1:  # 늘여야 하면
                self.MoveLineBegin()  # hwp.MoveLineUp()으로 실행하면 위에 표가 있을 때 들어가버림
                self.MovePrevChar()
                self.MoveLineBegin()
                self.MoveSelLineEnd()
                self.MoveSelPrevWord()
            elif direction == -1:  # 줄여야 하면
                start_pos = self.get_pos()
                self.MoveLineBegin()
                self.MovePrevChar()
                self.MoveLineBegin()
                self.select_text_by_get_pos(self.get_pos(), start_pos)
                self.MoveSelPrevChar()

        def modify_spacing(direction=0):
            if direction == 0:
                return 0, ""
            loc_info = self.key_indicator()
            start_line_no = loc_info[5]
            string = f"{loc_info[3]}쪽 {loc_info[4]}단 {'' if self.get_pos()[0] == 0 else self.ParentCtrl.UserDesc}{start_line_no}줄({self.get_selected_text()})"
            if verbose:
                print(string, end=" : ")
            min_val = init_spacing
            max_val = init_spacing
            while self.key_indicator()[5] == start_line_no:
                if direction == -1:  # 줄여야 하면
                    self.CharShapeSpacingDecrease()
                    min_val -= 1
                elif direction == 1:  # 늘여야 하면
                    self.CharShapeSpacingIncrease()
                    max_val += 1
                if min_val == min_spacing or max_val == max_spacing:
                    self.set_font(Spacing=init_spacing)
                    if verbose:
                        print(f"[롤백]{string}\n")
                    break
            val = min_val if max_val == init_spacing else max_val
            if verbose:
                print(val)
            return val, string

        dd = defaultdict(list)
        self.MoveDocBegin()
        reset_para_spacing(init_spacing=init_spacing, init_ratio=init_ratio)
        self.MoveDocEnd()
        end_pos = self.get_pos()
        self.MoveDocBegin()
        while self.get_pos() != end_pos:
            direction = get_spacing_dir()
            select_spacing_area(direction)
            spacing, string = modify_spacing(direction)
            dd[spacing].append(string)

        area = 2
        while True:
            self.set_pos(area, 0, 0)
            if self.get_pos()[0] == 0:
                break
            reset_para_spacing(init_spacing=init_spacing, init_ratio=init_ratio)
            self.MoveListEnd()
            end_pos = self.get_pos()
            self.MoveListBegin()
            while self.get_pos() != end_pos:
                direction = get_spacing_dir()
                select_spacing_area(direction)
                spacing, string = modify_spacing(direction)
                dd[spacing].append(string)
            area += 1

        spacings = np.array(list(dd.keys()))
        if verbose:
            print("\n\n자간 평균 :", round(spacings.mean(), 1))
            print("자간 표준편차 :", round(spacings.std(), 1))
            print(f"자간 최대값 : {spacings.max()}({dd[spacings.max()]})")
            print(f"자간 최소값 : {spacings.min()}({dd[spacings.min()]})")
        return True

    def set_font(self,
                 Bold="",  # 진하게(True/False)
                 DiacSymMark="",  # 강조점(0~12)
                 Emboss="",  # 양각(True/False)
                 Engrave="",  # 음각(True/False)
                 FaceName="",  # 서체
                 FontType=1,  # 1(TTF),
                 Height="",  # 글자크기(pt, 0.1 ~ 4096)
                 Italic="",  # 이탤릭(True/False)
                 Offset="",  # 글자위치-상하오프셋(-100 ~ 100)
                 OutLineType="",  # 외곽선타입(0~6)
                 Ratio="",  # 장평(50~200)
                 ShadeColor="",
                 # 음영색(RGB, 0x000000 ~ 0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
                 ShadowColor="",  # 그림자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
                 ShadowOffsetX="",  # 그림자 X오프셋(-100 ~ 100)
                 ShadowOffsetY="",  # 그림자 Y오프셋(-100 ~ 100)
                 ShadowType="",  # 그림자 유형(0: 없음, 1: 비연속, 2:연속)
                 Size="",  # 글자크기 축소확대%(10~250)
                 SmallCaps="",  # 강조점
                 Spacing="",  # 자간(-50 ~ 50)
                 StrikeOutColor="",
                 # 취소선 색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
                 StrikeOutShape="",  # 취소선 모양(0~12, 0이 일반 취소선)
                 StrikeOutType="",  # 취소선 유무(True/False)
                 SubScript="",  # 아래첨자(True/False)
                 SuperScript="",  # 위첨자(True/False)
                 TextColor="",  # 글자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
                 UnderlineColor="",  # 밑줄색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
                 UnderlineShape="",  # 밑줄형태(0~12)
                 UnderlineType="",  # 밑줄위치(0:없음, 1:하단, 3:상단)
                 UseFontSpace="",  # 글꼴에 어울리는 빈칸(True/False)
                 UseKerning=""  # 커닝 적용(True/False) : 차이가 없다?
                 ):
        """
        글자모양을 메서드 형태로 수정할 수 있는 메서드.
        :param Bold:  # 진하게(True/False)
        :param DiacSymMark:  # 강조점(0~12)
        :param Emboss:  # 양각(True/False)
        :param Engrave:  # 음각(True/False)
        :param FaceName:  # 서체
        :param FontType:  # 1(TTF),
        :param Height:  # 글자크기(pt, 0.1 ~ 4096)
        :param Italic:  # 이탤릭(True/False)
        :param Offset:  # 글자위치-상하오프셋(-100 ~ 100)
        :param OutLineType:  # 외곽선타입(0~6)
        :param Ratio:   # 장평(50~200)
        :param ShadeColor:  # 음영색(RGB, 0x000000 ~ 0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
        :param ShadowColor:  # 그림자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
        :param ShadowOffsetX:  # 그림자 X오프셋(-100 ~ 100)
        :param ShadowOffsetY:  # 그림자 Y오프셋(-100 ~ 100)
        :param ShadowType:  # 그림자 유형(0: 없음, 1: 비연속, 2:연속)
        :param Size:  # 글자크기 축소확대%(10~250)
        :param SmallCaps:  # 강조점
        :param Spacing:  # 자간(-50 ~ 50)
        :param StrikeOutColor:  # 취소선 색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
        :param StrikeOutShape:  # 취소선 모양(0~12, 0이 일반 취소선)
        :param StrikeOutType:  # 취소선 유무(True/False)
        :param SubScript:  # 아래첨자(True/False)
        :param SuperScript:  # 위첨자(True/False)
        :param TextColor:  # 글자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
        :param UnderlineColor:  # 밑줄색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
        :param UnderlineShape:  # 밑줄형태(0~12)
        :param UnderlineType:  # 밑줄위치(0:없음, 1:하단, 3:상단)
        :param UseFontSpace:  # 글꼴에 어울리는 빈칸(True/False) : 차이가 나는 폰트를 못 찾았다...
        :param UseKerning: 커닝 적용(True/False) : 차이가 전혀 없다?
        :return:
        """
        d = {'Bold': Bold, 'DiacSymMark': DiacSymMark, 'Emboss': Emboss, 'Engrave': Engrave, 'FaceNameHangul': FaceName,
             'FaceNameHanja': FaceName, 'FaceNameJapanese': FaceName, 'FaceNameLatin': FaceName,
             'FaceNameOther': FaceName, 'FaceNameSymbol': FaceName, 'FaceNameUser': FaceName,
             'FontTypeHangul': FontType,
             'FontTypeHanja': FontType, 'FontTypeJapanese': FontType, 'FontTypeLatin': FontType,
             'FontTypeOther': FontType, 'FontTypeSymbol': FontType, 'FontTypeUser': FontType, 'Height': Height * 100,
             'Italic': Italic, 'OffsetHangul': Offset, 'OffsetHanja': Offset, 'OffsetJapanese': Offset,
             'OffsetLatin': Offset, 'OffsetOther': Offset, 'OffsetSymbol': Offset, 'OffsetUser': Offset,
             'OutLineType': OutLineType, 'RatioHangul': Ratio, 'RatioHanja': Ratio, 'RatioJapanese': Ratio,
             'RatioLatin': Ratio, 'RatioOther': Ratio, 'RatioSymbol': Ratio, 'RatioUser': Ratio,
             'ShadeColor': self.rgb_color(ShadeColor) if type(ShadeColor) == str and ShadeColor else ShadeColor,
             'ShadowColor': self.rgb_color(ShadowColor) if type(ShadowColor) == str and ShadowColor else ShadowColor,
             'ShadowOffsetX': ShadowOffsetX, 'ShadowOffsetY': ShadowOffsetY, 'ShadowType': ShadowType,
             'SizeHangul': Size, 'SizeHanja': Size, 'SizeJapanese': Size, 'SizeLatin': Size, 'SizeOther': Size,
             'SizeSymbol': Size, 'SizeUser': Size, 'SmallCaps': SmallCaps, 'SpacingHangul': Spacing,
             'SpacingHanja': Spacing, 'SpacingJapanese': Spacing, 'SpacingLatin': Spacing, 'SpacingOther': Spacing,
             'SpacingSymbol': Spacing, 'SpacingUser': Spacing, 'StrikeOutColor': StrikeOutColor,
             'StrikeOutShape': StrikeOutShape, 'StrikeOutType': StrikeOutType, 'SubScript': SubScript,
             'SuperScript': SuperScript,
             'TextColor': self.rgb_color(TextColor) if type(TextColor) == str and TextColor else TextColor,
             'UnderlineColor': self.rgb_color(UnderlineColor) if type(
                 UnderlineColor) == str and UnderlineColor else UnderlineColor, 'UnderlineShape': UnderlineShape,
             'UnderlineType': UnderlineType, 'UseFontSpace': UseFontSpace, 'UseKerning': UseKerning}
        pset = self.hwp.HParameterSet.HCharShape
        self.HAction.GetDefault("CharShape", pset.HSet)
        for key in d.keys():
            if d[key] != "":
                pset.__setattr__(key, d[key])
        return self.hwp.HAction.Execute("CharShape", pset.HSet)

    def cell_fill(self, face_color: tuple[int, int, int] = (217, 217, 217)):
        """
        선택한 셀에 색 채우기
        :param face_color:
        :return:
        """
        pset = self.hwp.HParameterSet.HCellBorderFill
        self.hwp.HAction.GetDefault("CellFill", pset.HSet)
        pset.FillAttr.type = self.hwp.BrushType("NullBrush|WinBrush")
        pset.FillAttr.WinBrushFaceColor = self.hwp.RGBColor(*face_color)
        pset.FillAttr.WinBrushHatchColor = self.hwp.RGBColor(153, 153, 153)
        pset.FillAttr.WinBrushFaceStyle = self.hwp.HatchStyle("None")
        pset.FillAttr.WindowsBrush = 1
        try:
            return self.hwp.HAction.Execute("CellFill", pset.HSet)
        finally:
            self.hwp.HAction.Run("Cancel")

    def fields_to_dict(self):
        """
        현재 문서에 저장된 필드명과 필드값을
        dict 타입으로 리턴하는 메서드.
        :return:
        """
        result = defaultdict(list)
        field_list = self.get_field_list(number=1)
        field_values = self.get_field_text(field_list)
        for i, j in zip(field_list.split("\x02"), field_values.split("\x02")):
            result[i.split("{")[0]].append(j)
        max_len = max([len(result[i]) for i in result])
        for i in result:
            if len(result[i]) == 1:
                result[i] *= max_len
        return result

    def get_into_nth_table(self, n=0, select=False):
        """
        문서 n번째 표의 첫 번째 셀로 이동하는 함수.
        첫 번째 표의 인덱스가 0이며, 음수인덱스 사용 가능.
        단, 표들의 인덱스 순서는 표의 위치 순서와 일치하지 않을 수도 있으므로 유의해야 한다.
        """
        if n >= 0:
            idx = 0
            ctrl = self.hwp.HeadCtrl
        else:
            idx = -1
            ctrl = self.hwp.LastCtrl
        if isinstance(n, type(ctrl)):
            # 정수인덱스 대신 ctrl 객체를 넣은 경우
            self.set_pos_by_set(n.GetAnchorPos(0))
            self.hwp.FindCtrl()
            self.ShapeObjTableSelCell()
            if not select:
                self.Cancel()
            return ctrl

        while ctrl:
            if ctrl.UserDesc == "표":
                if n in (0, -1):
                    self.set_pos_by_set(ctrl.GetAnchorPos(0))
                    self.hwp.FindCtrl()
                    self.ShapeObjTableSelCell()
                    if not select:
                        self.Cancel()
                    return ctrl
                else:
                    if idx == n:
                        self.set_pos_by_set(ctrl.GetAnchorPos(0))
                        self.hwp.FindCtrl()
                        self.ShapeObjTableSelCell()
                        if not select:
                            self.Cancel()
                        return ctrl
                    if n >= 0:
                        idx += 1
                    else:
                        idx -= 1
            if n >= 0:
                ctrl = ctrl.Next
            else:
                ctrl = ctrl.Prev
        return False  # raise IndexError(f"해당 인덱스의 표가 존재하지 않습니다."  #                  f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다.")

    def set_row_height(self, height: int | float, as_: Literal["mm", "hwpunit"] = "mm"):
        """
        캐럿이 표 안에 있는 경우
        캐럿이 위치한 행의 셀 높이를 조절하는 메서드(기본단위는 mm)
        :param height_mili:
        :return:
        """
        if not self.is_cell():
            raise AssertionError("캐럿이 표 안에 있지 않습니다. 표 안에서 실행해주세요.")
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.HSet.SetItem("ShapeType", 3)
        pset.HSet.SetItem("ShapeCellSize", 1)
        if as_.lower() == "mm":
            pset.ShapeTableCell.Height = self.hwp.MiliToHwpUnit(height)
        else:
            pset.ShapeTableCell.Height = height
        return self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

    def remove_background_picture(self):
        """
        표 안에 백그라운드 이미지가 삽입되어 있고,
        캐럿이 해당 셀 안에 들어있는 경우,
        이를 제거하는 메서드
        :return:
        """
        self.insert_background_picture("", border_type="SelectedCellDelete")

    def gradation_on_cell(self, color_list: list[tuple] | list[str] = [(0, 0, 0), (255, 255, 255)],
                          grad_type: Literal["Linear", "Radial", "Conical", "Square"] = "Linear", angle=0, xc=0, yc=0,
                          pos_list: list[int] = None, step_center=50, step=255, ):
        """
        셀에 그라데이션을 적용하는 메서드
        :param color_list:
        :param grad_type:
        :param angle:
        :param xc:
        :param yc:
        :param pos_list:
        :param step_center:
        :param step:
        :return:
        """
        if not self.is_cell():
            raise AssertionError("캐럿이 현재 표 안에 위치하지 않습니다. 표 안에서 다시 실행해주세요.")
        pset = self.hwp.HParameterSet.HCellBorderFill
        self.hwp.HAction.GetDefault("CellFill", pset.HSet)
        if not pset.FillAttr.type:
            pset.FillAttr.type = self.hwp.BrushType("NullBrush|GradBrush")
            pset.FillAttr.GradationType = 1
            pset.FillAttr.GradationCenterX = 0
            pset.FillAttr.GradationCenterY = 0
            pset.FillAttr.GradationAngle = 0
            pset.FillAttr.GradationStep = 1
            pset.FillAttr.GradationColorNum = 2
            pset.FillAttr.CreateItemArray("GradationIndexPos", 2)
            pset.FillAttr.GradationIndexPos.SetItem(0, 0)
            pset.FillAttr.GradationIndexPos.SetItem(1, 255)
            pset.FillAttr.GradationStepCenter = 50
            pset.FillAttr.CreateItemArray("GradationColor", 2)
            pset.FillAttr.GradationColor.SetItem(0, self.rgb_color(255, 255, 255))  # 시작색 ~ 끝색
            pset.FillAttr.GradationColor.SetItem(1, self.rgb_color(255, 255, 255))  # 시작색 ~ 끝색
            pset.FillAttr.GradationBrush = 1
            self.hwp.HAction.Execute("CellFill", pset.HSet)
        color_num = len(color_list)
        if color_num == 1:
            step = 1
        pset.FillAttr.type = self.hwp.BrushType("NullBrush|GradBrush")
        pset.FillAttr.GradationType = self.hwp.Gradation(grad_type)  # 0은 검정. Linear:1, Radial:2, Conical:3, Square:4
        pset.FillAttr.GradationCenterX = xc  # 가로중심
        pset.FillAttr.GradationCenterY = yc  # 세로중심
        pset.FillAttr.GradationAngle = angle  # 기울임
        pset.FillAttr.GradationStep = step  # 번짐정도(영역개수) 2~255 (0은 투명, 1은 시작색)
        pset.FillAttr.GradationColorNum = color_num  # ?
        pset.FillAttr.CreateItemArray("GradationIndexPos", color_num)

        if not pos_list and color_num > 1:
            pos_list = [round(i / (color_num - 1) * 255) for i in range(color_num)]
        elif color_num == 1:
            pos_list = [255]
        elif pos_list[-1] == 100:
            pos_list = [round(i * 2.55) for i in pos_list]
        for i in range(color_num):
            pset.FillAttr.GradationIndexPos.SetItem(i, pos_list[i])
        pset.FillAttr.GradationStepCenter = step_center  # 번짐중심(%), 중간이 50, 작을수록 시작점에, 클수록 끝점에 그라데이션 중간점이 가까워짐
        pset.FillAttr.CreateItemArray("GradationColor", color_num)
        for i in range(color_num):
            if type(color_list[i]) == str:
                pset.FillAttr.GradationColor.SetItem(i, self.rgb_color(color_list[i]))  # 시작색 ~ 끝색
            elif check_tuple_of_ints(color_list[i]):
                pset.FillAttr.GradationColor.SetItem(i, self.rgb_color(*color_list[i]))  # 시작색 ~ 끝색
        pset.FillAttr.GradationBrush = 1
        return self.hwp.HAction.Execute("CellFill", pset.HSet)

    def get_available_font(self) -> list:
        """
        현재 사용 가능한 폰트 리스트를 리턴.
        API 사용시 발생하는 오류로 인해 현재는 한글 폰트만 지원하고 있음.
        :return:
            현재 사용 가능한 폰트 리스트
        """
        result_list = []
        initial_font = self.CharShape.Item("FaceNameHangul")
        # for font_type in [
        #     'FaceNameHangul',
        #     'FaceNameHanja',
        #     'FaceNameJapanese',
        #     'FaceNameLatin',
        #     'FaceNameOther',
        #     'FaceNameSymbol',
        #     'FaceNameUser',
        #     'FaceNameHangul',
        # ]:
        cur_face = self.CharShape.Item("FaceNameHangul")
        while self.CharShapeNextFaceName():
            result_list.append(self.CharShape.Item('FaceNameHangul'))
            if cur_face == self.CharShape.Item('FaceNameHangul'):
                break
        return list(set(result_list))

    def get_charshape(self):
        pset = self.hwp.HParameterSet.HCharShape
        self.hwp.HAction.GetDefault("CharShape", pset.HSet)
        return pset

    def get_charshape_as_dict(self):
        result_dict = {}
        for key in self.HParameterSet.HCharShape._prop_map_get_.keys():
            result_dict[key] = self.CharShape.Item(key)
        return result_dict

    def set_charshape(self, pset):
        if isinstance(pset, dict):
            new_pset = self.hwp.HParameterSet.HCharShape
            for key in pset.keys():
                try:
                    new_pset.__setattr__(key, pset[key])
                except pythoncom.com_error:
                    print(key, pset[key])
        elif type(pset) == type(self.HParameterSet.HCharShape):
            new_pset = pset
        return self.hwp.HAction.Execute("CharShape", new_pset.HSet)

    def get_pagedef(self):
        """
        현재 페이지의 용지정보 파라미터셋을 리턴한다.
        리턴값은 set_pagedef 메서드를 통해
        새로운 문서에 적용할 수 있다.
        연관 메서드로, get_pagedef_as_dict는 보다 직관적으로
        밀리미터 단위로 변환된 dict를 리턴하므로,
        get_pagedef_as_dict 메서드를 추천한다.
        """
        pset = self.hwp.HParameterSet.HSecDef
        self.hwp.HAction.GetDefault("PageSetup", pset.HSet)
        return pset

    def get_pagedef_as_dict(self, as_: Literal["kor", "eng"] = "kor"):
        """
        현재 페이지의 용지정보를 dict 형태로 리턴한다.
        dict의 각 값은 밀리미터 단위로 변환된 값이며,
        set_pagedef 실행시 내부적으로 HWPUnit으로 자동변환하여 적용한다.
        (as_ 파라미터를 "eng"로 변경하면 원래 영문 아이템명의 사전을 리턴한다.)
        :return:
            현재 페이지의 용지정보(dict)
            각 키의 원래 아이템명은 아래와 같다.

            PaperWidth: 용지폭
            PaperHeight: 용지길이
            Landscape: 용지방향(0: 가로, 1:세로)
            GutterType: 제본타입(0: 한쪽, 1:맞쪽, 2:위쪽)
            TopMargin: 위쪽
            HeaderLen: 머리말
            LeftMargin: 왼쪽
            GutterLen: 제본여백
            RightMargin: 오른쪽
            FooterLen: 꼬리말
            BottomMargin: 아래쪽
        """
        code_to_desc = {'PaperWidth': "용지폭", 'PaperHeight': "용지길이", 'Landscape': "용지방향",  # 0: 가로, 1:세로
                        'GutterType': "제본타입",  # 0: 한쪽, 1:맞쪽, 2:위쪽
                        'TopMargin': "위쪽", 'HeaderLen': "머리말", 'LeftMargin': "왼쪽", 'GutterLen': "제본여백",
                        'RightMargin': "오른쪽",
                        'FooterLen': "꼬리말", 'BottomMargin': "아래쪽", }

        pset = self.hwp.HParameterSet.HSecDef
        self.hwp.HAction.GetDefault("PageSetup", pset.HSet)
        result_dict = {}
        for key in pset.PageDef._prop_map_get_.keys():
            if key == "HSet":
                pass
            elif key in ["Landscape", "GutterType"]:
                if as_ == "kor":
                    result_dict[code_to_desc[key]] = eval(f"pset.PageDef.{key}")
                else:
                    result_dict[key] = eval(f"pset.PageDef.{key}")
            else:
                if as_ == "kor":
                    result_dict[code_to_desc[key]] = self.hwp_unit_to_mili(eval(f"pset.PageDef.{key}"))
                else:
                    result_dict[key] = self.hwp_unit_to_mili(eval(f"pset.PageDef.{key}"))

        return result_dict

    def set_pagedef(self, pset, apply: Literal["cur", "all", "new"] = "cur"):
        """
        get_pagedef 또는 get_pagedef_as_dict를 통해 얻은 용지정보를
        새 문서에 적용하는 메서드이다.
        :param pset:
            파라미터셋 또는 dict. 용지정보를 담은 객체
        :return:
        """
        if isinstance(pset, dict):
            desc_to_code = {"용지폭": 'PaperWidth', "용지길이": 'PaperHeight', "용지방향": 'Landscape', "제본타입": 'GutterType',
                            "위쪽": 'TopMargin', "머리말": 'HeaderLen', "왼쪽": 'LeftMargin', "제본여백": 'GutterLen',
                            "오른쪽": 'RightMargin',
                            "꼬리말": 'FooterLen', "아래쪽": 'BottomMargin', }

            new_pset = self.hwp.HParameterSet.HSecDef
            for key in pset.keys():
                if key in desc_to_code.keys():  # 한글인 경우
                    if key in ["용지방향", "제본여백"]:
                        exec(f"new_pset.PageDef.{desc_to_code[key]} = {pset[key]}")
                    else:
                        exec(f"new_pset.PageDef.{desc_to_code[key]} = {self.mili_to_hwp_unit(pset[key])}")
                elif key in desc_to_code.values():  # 영문인 경우
                    if key in ["Landscape", "GutterLen"]:
                        exec(f"new_pset.PageDef.{key} = {pset[key]}")
                    else:
                        exec(f"new_pset.PageDef.{key} = {self.mili_to_hwp_unit(pset[key])}")

            # 적용범위
            if apply == "cur":
                new_pset.HSet.SetItem("ApplyTo", 2)
            elif apply == "all":
                new_pset.HSet.SetItem("ApplyTo", 3)
            elif apply == "new":
                new_pset.HSet.SetItem("ApplyTo", 4)
            return self.hwp.HAction.Execute("PageSetup", new_pset.HSet)

        elif type(pset) == type(self.hwp.HParameterSet.HSecDef):
            if apply == "cur":
                pset.HSet.SetItem("ApplyTo", 2)
            elif apply == "all":
                pset.HSet.SetItem("ApplyTo", 3)
            elif apply == "new":
                pset.HSet.SetItem("ApplyTo", 4)
            return self.hwp.HAction.Execute("PageSetup", pset.HSet)

    def save_block_as(self, path, format="HWP", attributes=1):
        if path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)
        pset = self.hwp.HParameterSet.HFileOpenSave
        self.hwp.HAction.GetDefault("FileSaveBlock_S", pset.HSet)
        pset.filename = path
        pset.Format = format
        pset.Attributes = attributes
        return self.hwp.HAction.Execute("FileSaveBlock_S", pset.HSet)

    def goto_page(self, page_num):
        pset = self.hwp.HParameterSet.HGotoE
        self.hwp.HAction.GetDefault("Goto", pset.HSet)
        pset.HSet.SetItem("DialogResult", page_num)
        pset.SetSelectionIndex = 1
        return self.hwp.HAction.Execute("Goto", pset.HSet)

    def table_from_data(self, data, transpose=False, header0="", treat_as_char=False, header=True, index=True,
                        cell_fill: bool | tuple[int, int, int] = False, header_bold=True):
        """
        dict, list 또는 csv나 xls, xlsx 및 json처럼 2차원 스프레드시트로 표현 가능한 데이터에 대해서,
        정확히는 pd.DataFrame으로 변환 가능한 데이터에 대해 아래아한글 표로 변환하는 작업을 한다.
        내부적으로 판다스 데이터프레임으로 변환하는 과정을 거친다.
        :param data: 테이블로 변환할 데이터
        :param transpose: 행/열 전환
        :param header0: index=True일 경우 (1,1) 셀에 들어갈 텍스트
        :param treat_as_char: 글자처럼 취급 여부
        :param header: 1행을 "제목행"으로 선택할지 여부
        :param header_bold: 1행의 텍스트에 bold를 적용할지 여부
        :return:
        """
        if type(data) in [dict, list]:
            df = pd.DataFrame(data)
        elif type(data) is str:  # 엑셀파일 경로 또는 json으로 간주
            if os.path.isfile(data):
                df = pd.read_excel(data) if ".xls" in data else pd.read_csv(data)
            else:
                df = pd.read_json(StringIO(data))
        else:
            df = data
        if transpose:
            df = df.T
        if index:
            idx_list = list(df.index)
            self.create_table(rows=len(df) + 1, cols=len(df.columns) + 1, treat_as_char=treat_as_char, header=header)
            self.insert_text(header0)
            self.TableRightCellAppend()
        else:
            self.create_table(rows=len(df) + 1, cols=len(df.columns), treat_as_char=treat_as_char, header=header)
        for i in df.columns:
            self.insert_text(i)
            self.TableRightCellAppend()
        for i in range(len(df)):
            if index:
                self.insert_text(idx_list.pop(0))
                self.TableRightCellAppend()
            for j in df.iloc[i]:
                self.insert_text(j)
                self.TableRightCell()
        self.TableColBegin()
        self.TableColPageUp()
        self.TableCellBlockExtendAbs()
        self.TableColEnd()
        if header_bold:
            self.CharShapeBold()
        if cell_fill:
            if isinstance(cell_fill, tuple):
                self.cell_fill(cell_fill)
            else:
                self.cell_fill()
        self.TableColBegin()
        self.Cancel()

    def count(self, word):
        return self.get_text_file().count(word)

    def delete_all_fields(self):
        start_pos = self.get_pos()
        ctrl = self.hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "%clk":
                self.hwp.DeleteCtrl(ctrl)
            ctrl = ctrl.Next
        for field in self.get_field_list().split("\x02"):
            self.rename_field(field, "")
        return self.set_pos(*start_pos)

    def delete_field_by_name(self, field_name, idx=-1):
        start_pos = self.get_pos()
        ctrl = self.hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "%clk":
                self.set_pos_by_set(ctrl.GetAnchorPos(1))
                if self.get_cur_field_name() == field_name:
                    self.hwp.DeleteCtrl(ctrl)
            ctrl = ctrl.Next
        try:
            if idx == -1:
                return self.rename_field(field_name, "")
            elif not field_name.endswith("}}"):
                return self.rename_field(field_name + f"{{{{{idx}}}}}", "")
        finally:
            self.set_pos(*start_pos)

    def markpen_on_selection(self, r=255, g=255, b=0):
        pset = self.hwp.HParameterSet.HMarkpenShape
        self.hwp.HAction.GetDefault("MarkPenShape", pset.HSet)
        pset.Color = self.rgb_color(r, g, b)
        return self.hwp.HAction.Execute("MarkPenShape", pset.HSet)

    def open_pdf(self, pdf_path, this_window=1):
        """
        pdf를 hwp문서로 변환하여 여는 함수.
        (최초 실행시 "다시 표시 안함ㅁ" 체크박스에 체크를 해야 한다.)

        :param pdf_path:
            pdf파일의 경로
        :param this_window:
            현재 창에 열고 싶으면 1, 새 창에 열고 싶으면 0.
            하지만 아직(2023.12.11.) 작동하지 않음.
        :return:
        """
        if pdf_path.lower()[1] != ":":
            pdf_path = os.path.join(os.getcwd(), pdf_path)
        pset = self.hwp.HParameterSet.HFileOpenSave
        self.hwp.HAction.Run("CallPDFConverter")
        self.hwp.HAction.GetDefault("FileOpenPDF", pset.HSet)
        pset.Attributes = 0
        pset.filename = pdf_path
        pset.OpenFlag = this_window
        return self.hwp.HAction.Execute("FileOpenPDF", pset.HSet)

    def msgbox(self, string, flag: int = 0):
        msgbox = self.hwp.XHwpMessageBox  # 메시지박스 생성
        msgbox.string = string
        msgbox.Flag = flag  # [확인] 버튼만 나타나게 설정
        msgbox.DoModal()  # 메시지박스 보이기
        return msgbox.Result

    def insert_file(self, filename, keep_section=1, keep_charshape=1, keep_parashape=1, keep_style=1,
                    move_doc_end=False):
        if filename.lower()[1] != ":":
            filename = os.path.join(os.getcwd(), filename)
        pset = self.hwp.HParameterSet.HInsertFile
        self.hwp.HAction.GetDefault("InsertFile", pset.HSet)
        pset.filename = filename
        pset.KeepSection = keep_section
        pset.KeepCharshape = keep_charshape
        pset.KeepParashape = keep_parashape
        pset.KeepStyle = keep_style
        try:
            return self.hwp.HAction.Execute("InsertFile", pset.HSet)
        finally:
            if move_doc_end:
                self.MoveDocEnd()

    def insert_memo(self, text, memo_type: Literal["revision", "memo"] = "memo"):
        """
        선택한 단어 범위에 메모고침표를 삽입하는 코드.
        한/글에서 일반 문자열을 삽입하는 코드와 크게 다르지 않다.
        선택모드가 아닌 경우 캐럿이 위치한 단어에 메모고침표를 삽입한다.
        :param text: str
        :return: None
        """
        if memo_type == "revision":
            self.InsertFieldRevisionChagne()  # 이 라인이 메모고침표 삽입하는 코드
        elif memo_type == "memo":
            self.InsertFieldMemo()
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

    def find_backward(self, src, regex=False):
        """
        문서 위쪽으로 find 메서드를 수행.
        해당 단어를 선택한 상태가 되며,
        문서 처음에 도달시 False 리턴

        :param src:
            찾을 단어
        :return:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴,
            단어가 더이상 없으면 False를 리턴
        """
        self.SetMessageBoxMode(0x2fff1)
        init_pos = str(self.KeyIndicator())
        pset = self.hwp.HParameterSet.HFindReplace
        pset.MatchCase = 1
        pset.SeveralWords = 1
        pset.UseWildCards = 1
        pset.AutoSpell = 1
        pset.Direction = self.find_dir("Backward")
        pset.FindString = src
        pset.IgnoreMessage = 0
        pset.HanjaFromHangul = 1
        pset.FindRegExp = regex
        try:
            return self.hwp.HAction.Execute("RepeatFind", pset.HSet)
        finally:
            self.SetMessageBoxMode(0xfffff)

    def find_forward(self, src, regex=False):
        """
        문서 아래쪽으로 find를 수행하는 메서드.
        해당 단어를 선택한 상태가 되며,
        문서 끝에 도달시 False 리턴.

        :param src:
            찾을 단어
        :return:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴,
            단어가 더이상 없으면 False를 리턴
        """
        self.SetMessageBoxMode(0x2fff1)
        init_pos = str(self.KeyIndicator())
        pset = self.hwp.HParameterSet.HFindReplace
        pset.MatchCase = 1
        pset.SeveralWords = 1
        pset.UseWildCards = 1
        pset.AutoSpell = 1
        pset.Direction = self.find_dir("Forward")
        pset.FindString = src
        pset.IgnoreMessage = 0
        pset.HanjaFromHangul = 1
        pset.FindRegExp = regex
        try:
            return self.hwp.HAction.Execute("RepeatFind", pset.HSet)
        finally:
            self.SetMessageBoxMode(0xfffff)

    def find(self, src, direction: Literal["Forward", "Backward", "AllDoc"] = "Forward", regex=False):
        """
        direction 방향으로 특정 단어를 찾아가는 메서드.
        해당 단어를 선택한 상태가 되며,
        탐색방향에 src 문자열이 없는 경우 False를 리턴

        :param src:
            찾을 단어
        :param direction:
            탐색방향
            "Forward": 아래쪽으로
            "Backward": 위쪽으로
            "AllDoc": 아래쪽 우선으로 찾고 문서끝 도달시 처음으로 돌아감.

        :return:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴,
            단어가 더이상 없으면 False를 리턴
        """
        self.SetMessageBoxMode(0x2fff1)
        init_pos = str(self.KeyIndicator())
        pset = self.hwp.HParameterSet.HFindReplace
        # self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.MatchCase = 1
        pset.SeveralWords = 1
        pset.UseWildCards = 1
        pset.AutoSpell = 1
        pset.Direction = self.find_dir(direction)
        pset.FindString = src
        pset.IgnoreMessage = 0
        pset.HanjaFromHangul = 1
        pset.FindRegExp = regex
        try:
            return self.hwp.HAction.Execute("RepeatFind", pset.HSet)
        finally:
            self.SetMessageBoxMode(0xfffff)

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
        self.MoveDocBegin()
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

        self.MoveDocBegin()
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

    def find_replace(self, src, dst, regex=False, direction: Literal["Backward", "Forward", "AllDoc"] = "Forward"):
        """
        아래아한글의 찾아바꾸기와 동일한 액션을 수항해지만,
        re=True로 설정하고 실행하면,
        문단별로 잘라서 문서 전체를 순회하며
        파이썬의 re.sub 함수를 실행한다.
        """
        self.SetMessageBoxMode(0x2fff1)
        if regex:
            whole_text = self.get_text_file()
            src_list = [i.group() for i in re.finditer(src, whole_text)]
            dst_list = [re.sub(src, dst, i) for i in src_list]
            for i, j in zip(src_list, dst_list):
                try:
                    return self.find_replace(i, j, direction=direction)
                finally:
                    self.SetMessageBoxMode(0xfffff)

        else:
            pset = self.hwp.HParameterSet.HFindReplace
            # self.hwp.HAction.GetDefault("AllReplace", pset.HSet)
            pset.Direction = self.hwp.FindDir(direction)
            pset.FindString = src  # "\\r\\n"
            pset.ReplaceString = dst  # "^n"
            pset.ReplaceMode = 1
            pset.IgnoreMessage = 0
            pset.HanjaFromHangul = 1
            pset.AutoSpell = 1
            pset.FindType = 1
            try:
                return self.hwp.HAction.Execute("ExecReplace", pset.HSet)
            finally:
                self.SetMessageBoxMode(0xfffff)

    def find_replace_all(self, src, dst, regex=False):
        """
        아래아한글의 찾아바꾸기와 동일한 액션을 수항해지만,
        re=True로 설정하고 실행하면,
        문단별로 잘라서 문서 전체를 순회하며
        파이썬의 re.sub 함수를 실행한다.
        """
        self.SetMessageBoxMode(0x2fff1)
        if regex:
            whole_text = self.get_text_file()
            src_list = [i.group() for i in re.finditer(src, whole_text)]
            dst_list = [re.sub(src, dst, i) for i in src_list]
            for i, j in zip(src_list, dst_list):
                self.find_replace_all(i, j)
        else:
            pset = self.hwp.HParameterSet.HFindReplace
            # self.hwp.HAction.GetDefault("AllReplace", pset.HSet)
            pset.Direction = self.hwp.FindDir("AllDoc")
            pset.FindString = src  # "\\r\\n"
            pset.ReplaceString = dst  # "^n"
            pset.ReplaceMode = 1
            pset.IgnoreMessage = 0
            pset.HanjaFromHangul = 1
            pset.AutoSpell = 1
            pset.FindType = 1
            try:
                return self.hwp.HAction.Execute("AllReplace", pset.HSet)
            finally:
                self.SetMessageBoxMode(0xfffff)

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
        result = result.replace("HAction.", "hwp.HAction.").replace("HParameterSet.", "hwp.HParameterSet.")
        result = re.sub(r"= (?!hwp\.)(\D)", r"= hwp.\g<1>", result)
        result = result.replace('hwp."', '"')
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

    def HwpUnitToMili(self, hwp_unit: int) -> float:
        """
        HwpUnit 값을 밀리미터로 변환한 값을 리턴한다.
        HwpUnit으로 리턴되었거나, 녹화된 코드의 HwpUnit값을 확인할 때 유용하게 사용할 수 있다.

        :return:
            HwpUnit을 7200으로 나눈 후 25.4를 곱하고 반올림한 값
        """
        return round(hwp_unit / 7200 * 25.4, 4)

    def create_table(self, rows, cols, treat_as_char: bool = True, width_type=0, height_type=0, header=True, height=0):
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
        total_width = (
                sec_def.PageDef.PaperWidth - sec_def.PageDef.LeftMargin - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen - self.mili_to_hwp_unit(
            2))

        pset.WidthValue = total_width  # 표 너비(근데 영향이 없는 듯)
        if height and height_type == 1:  # 표높이가 정의되어 있으면
            # 페이지 최대 높이 계산
            total_height = (
                    sec_def.PageDef.PaperHeight - sec_def.PageDef.TopMargin - sec_def.PageDef.BottomMargin - sec_def.PageDef.HeaderLen - sec_def.PageDef.FooterLen - self.mili_to_hwp_unit(
                2))
            pset.HeightValue = min(self.hwp.MiliToHwpUnit(height), total_height)  # 표 높이
            pset.CreateItemArray("RowHeight", rows)  # 행 m개 생성
            each_row_height = min((self.mili_to_hwp_unit(height) - self.mili_to_hwp_unit((0.5 + 0.5) * rows)) // rows,
                                  (total_height - self.mili_to_hwp_unit((0.5 + 0.5) * rows)) // rows)
            for i in range(rows):
                pset.RowHeight.SetItem(i, each_row_height)  # 1열
            pset.TableProperties.Height = min(self.MiliToHwpUnit(height),
                                              total_height - self.mili_to_hwp_unit((0.5 + 0.5) * rows))

        pset.CreateItemArray("ColWidth", cols)  # 열 n개 생성
        each_col_width = total_width - self.mili_to_hwp_unit(3.6 * cols)
        for i in range(cols):
            pset.ColWidth.SetItem(i, each_col_width)  # 1열
        # pset.TableProperties.TreatAsChar = treat_as_char  # 글자처럼 취급
        pset.TableProperties.Width = total_width  # self.hwp.MiliToHwpUnit(148)  # 표 너비
        self.hwp.HAction.Execute("TableCreate", pset.HSet)  # 위 코드 실행

        # 글자처럼 취급 여부 적용(treat_as_char)
        ctrl = self.hwp.CurSelectedCtrl or self.hwp.ParentCtrl
        pset = self.hwp.CreateSet("Table")
        pset.SetItem("TreatAsChar", treat_as_char)
        ctrl.Properties = pset

        # 제목 행 여부 적용(header)
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.ShapeTableCell.Header = header
        self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

    def get_selected_text(self, as_: Literal["list", "str"] = "str"):
        """
        한/글 문서 선택 구간의 텍스트를 리턴하는 메서드.
        :return:
            선택한 문자열
        """
        if self.SelectionMode == 0:
            if self.is_cell():
                self.TableCellBlock()
            else:
                self.Select()
                self.Select()
        if not self.hwp.InitScan(Range=0xff):
            return ""
        if as_ == "list":
            result = []
        else:
            result = ""
        state = 2
        while state not in [0, 1]:
            state, text = self.hwp.GetText()
            if as_ == "list":
                result.append(text)
            else:
                result += text
        self.hwp.ReleaseScan()
        return result if type(result) == str else result[:-1]

    def table_to_csv(self, n="", filename="result.csv", encoding="utf-8", startrow=0):
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
        ctrl = self.hwp.HeadCtrl
        if isinstance(n, type(ctrl)):
            # 정수인덱스 대신 ctrl 객체를 넣은 경우
            self.set_pos_by_set(n.GetAnchorPos(0))
            self.find_ctrl()
            self.ShapeObjTableSelCell()
        elif n == "" and self.is_cell():
            # 기본값은 현재위치의 표를 잡아오기
            self.TableCellBlock()
            self.TableColBegin()
            self.TableColPageUp()
        elif n == "" or isinstance(n, int):
            if n == "":
                n = 0
            if n >= 0:
                idx = 0
            else:
                idx = -1
                ctrl = self.hwp.LastCtrl

            while ctrl:
                if ctrl.UserDesc == "표":
                    if n in (0, -1):
                        self.set_pos_by_set(ctrl.GetAnchorPos(0))
                        self.hwp.FindCtrl()
                        self.ShapeObjTableSelCell()
                        break
                    else:
                        if idx == n:
                            self.set_pos_by_set(ctrl.GetAnchorPos(0))
                            self.hwp.FindCtrl()
                            self.ShapeObjTableSelCell()
                            break
                        if n >= 0:
                            idx += 1
                        else:
                            idx -= 1
                if n >= 0:
                    ctrl = ctrl.Next
                else:
                    ctrl = ctrl.Prev

            try:
                self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            except AttributeError:
                raise IndexError(f"해당 인덱스의 표가 존재하지 않습니다."
                                 f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다.")
            self.hwp.FindCtrl()
            self.ShapeObjTableSelCell()
        data = [self.get_selected_text()]
        col_count = 1
        start = False
        while self.TableRightCell():
            if not startrow:
                if re.match(r"\([A-Z]+1\)", self.hwp.KeyIndicator()[-1]):
                    col_count += 1
                data.append(self.get_selected_text())
            else:
                if re.match(rf"\([A-Z]+{1 + startrow}\)", self.hwp.KeyIndicator()[-1]):
                    col_count += 1
                    start = True
                if start:
                    data.append(self.get_selected_text())

        array = np.array(data).reshape(-1, col_count)
        df = pd.DataFrame(array[1:], columns=array[0])
        self.hwp.SetPos(*start_pos)
        df.to_csv(filename, index=False, encoding=encoding)
        self.hwp.SetPos(*start_pos)
        print(os.path.join(os.getcwd(), filename))
        return None

    def table_to_df_q(self, n="", startrow=0, columns=[]):
        """
        (2024. 3. 14. for문 추출 구조에서, 한 번에 추출하는 방식으로 변경->속도개선)
        한/글 문서의 n번째 표를 판다스 데이터프레임으로 리턴하는 메서드.
        n을 넣지 않는 경우, 캐럿이 셀에 있다면 해당 표를 df로,
        캐럿이 표 밖에 있다면 첫 번째 표를 df로 리턴한다.
        startrow는 표 제목에 일부 병합이 되어 있는 경우
        df로 변환시작할 행을 특정할 때 사용된다.
        :return:
            pd.DataFrame
        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> df = hwp.table_to_df(0)
        """
        start_pos = self.hwp.GetPos()
        ctrl = self.hwp.HeadCtrl
        if isinstance(n, type(ctrl)):
            # 정수인덱스 대신 ctrl 객체를 넣은 경우
            self.set_pos_by_set(n.GetAnchorPos(0))
            self.find_ctrl()
            self.ShapeObjTableSelCell()
        elif n == "" and self.is_cell():
            # 기본값은 현재위치의 표를 잡아오기
            self.TableCellBlock()
            self.TableColBegin()
            self.TableColPageUp()
        elif n == "" or isinstance(n, int):
            if n == "":
                n = 0
            if n >= 0:
                idx = 0
            else:
                idx = -1
                ctrl = self.hwp.LastCtrl

            while ctrl:
                if ctrl.UserDesc == "표":
                    if n in (0, -1):
                        self.set_pos_by_set(ctrl.GetAnchorPos(0))
                        self.hwp.FindCtrl()
                        self.ShapeObjTableSelCell()
                        break
                    else:
                        if idx == n:
                            self.set_pos_by_set(ctrl.GetAnchorPos(0))
                            self.hwp.FindCtrl()
                            self.ShapeObjTableSelCell()
                            break
                        if n >= 0:
                            idx += 1
                        else:
                            idx -= 1
                if n >= 0:
                    ctrl = ctrl.Next
                else:
                    ctrl = ctrl.Prev

            try:
                self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            except AttributeError:
                raise IndexError(f"해당 인덱스의 표가 존재하지 않습니다."
                                 f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다.")
            self.hwp.FindCtrl()
            self.ShapeObjTableSelCell()

        if startrow:
            while int(self.get_cell_addr()[1:]) - 1 != startrow:
                self.TableRightCell()
        self.TableCellBlock()
        self.TableCellBlockExtend()
        self.TableColPageDown()
        self.TableColEnd()
        # rows = int(re.sub(r"[A-Z]+", "", self.get_cell_addr()))
        rows = int(re.sub(r"[A-Z]+", "", self.get_cell_addr())) - startrow

        arr = np.array(self.get_selected_text(as_="list"), dtype=object).reshape(rows, -1)
        # if startrow:
        #     arr = arr[startrow:]
        if columns:
            if len(columns) != len(arr[0]):
                raise IndexError("columns의 길이가 열의 갯수와 맞지 않습니다.")
            df = pd.DataFrame(arr, columns=columns)
        else:
            df = pd.DataFrame(arr[1:], columns=arr[0])
        self.hwp.SetPos(*start_pos)
        return df

    def table_to_df(self, n="", startrow=0, columns=[]):
        """
        (2024. 3. 14. for문 추출 구조에서, 한 번에 추출하는 방식으로 변경->속도개선)
        한/글 문서의 n번째 표를 판다스 데이터프레임으로 리턴하는 메서드.
        n을 넣지 않는 경우, 캐럿이 셀에 있다면 해당 표를 df로,
        캐럿이 표 밖에 있다면 첫 번째 표를 df로 리턴한다.
        startrow는 표 제목에 일부 병합이 되어 있는 경우
        df로 변환시작할 행을 특정할 때 사용된다.
        :return:
            pd.DataFrame
        :example:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> df = hwp.table_to_df(0)
        """
        start_pos = self.hwp.GetPos()
        ctrl = self.hwp.HeadCtrl
        if isinstance(n, type(ctrl)):
            # 정수인덱스 대신 ctrl 객체를 넣은 경우
            self.set_pos_by_set(n.GetAnchorPos(0))
            self.find_ctrl()
            self.ShapeObjTableSelCell()
        elif n == "" and self.is_cell():
            # 기본값은 현재위치의 표를 잡아오기
            self.TableCellBlock()
            self.TableColBegin()
            self.TableColPageUp()
        elif n == "" or isinstance(n, int):
            if n == "":
                n = 0
            if n >= 0:
                idx = 0
            else:
                idx = -1
                ctrl = self.hwp.LastCtrl

            while ctrl:
                if ctrl.UserDesc == "표":
                    if n in (0, -1):
                        self.set_pos_by_set(ctrl.GetAnchorPos(0))
                        self.hwp.FindCtrl()
                        self.ShapeObjTableSelCell()
                        break
                    else:
                        if idx == n:
                            self.set_pos_by_set(ctrl.GetAnchorPos(0))
                            self.hwp.FindCtrl()
                            self.ShapeObjTableSelCell()
                            break
                        if n >= 0:
                            idx += 1
                        else:
                            idx -= 1
                if n >= 0:
                    ctrl = ctrl.Next
                else:
                    ctrl = ctrl.Prev

            try:
                self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            except AttributeError:
                raise IndexError(f"해당 인덱스의 표가 존재하지 않습니다."
                                 f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다.")
            self.hwp.FindCtrl()
            self.ShapeObjTableSelCell()

        self.TableCellBlock()
        self.TableCellBlockExtend()
        self.TableCellBlockExtend()
        rows = int(re.sub(r"[A-Z]+", "", self.get_cell_addr()))
        arr = np.array(self.get_selected_text(as_="list")).reshape(rows, -1)
        if startrow:
            arr = arr[startrow:]
        if columns:
            if len(columns) != len(arr[0]):
                raise IndexError("columns의 길이가 열의 갯수와 맞지 않습니다.")
            df = pd.DataFrame(arr, columns=columns)
        else:
            df = pd.DataFrame(arr[1:], columns=arr[0])
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
        try:
            self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
        finally:
            self.hwp.Run("Cancel")

    def insert_text(self, text):
        """
        한/글 문서 내 캐럿 위치에 문자열을 삽입하는 메서드.
        :return:
            삽입 성공시 True, 실패시 False를 리턴함.
        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_text('Hello world!')
            >>> hwp.BreakPara()
        """
        param = self.hwp.HParameterSet.HInsertText
        self.hwp.HAction.GetDefault("InsertText", param.HSet)
        param.Text = text
        return self.hwp.HAction.Execute("InsertText", param.HSet)

    def insert_lorem(self, para_num=1):
        api_url = f'https://api.api-ninjas.com/v1/loremipsum?paragraphs={para_num}'

        headers = {'X-Api-Key': "hzzbbAAy7mQjKyXSW5quRw==PbJStWB0ymMpGRH1"}

        req = request.Request(api_url, headers=headers)

        try:
            with request.urlopen(req) as response:
                response_text = json.loads(response.read().decode('utf-8'))["text"].replace("\n", "\r\n")
        except urllib.error.HTTPError as e:
            print("Error:", e.code, e.reason)
        except urllib.error.URLError as e:
            print("Error:", e.reason)
        return self.insert_text(response_text)

    def move_caption(self, location: Literal["Top", "Bottom", "Left", "Right"] = "Bottom",
                     align: Literal["Left", "Center", "Right", "Distribute", "Division", "Justify"] = "Justify"):
        """
        한/글 문서 내 모든 표의 주석 위치를 이동하는 메서드.
        """
        start_pos = self.hwp.GetPos()
        ctrl = self.HeadCtrl
        while ctrl:
            if ctrl.UserDesc == "번호 넣기":
                self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                if align == "Left":
                    self.ParagraphShapeAlignLeft()
                elif align == "Center":
                    self.ParagraphShapeAlignCenter()
                elif align == "Right":
                    self.ParagraphShapeAlignRight()
                elif align == "Distribute":
                    self.ParagraphShapeAlignDistribute()
                elif align == "Division":
                    self.ParagraphShapeAlignDivision()
                elif align == "Justify":
                    self.ParagraphShapeAlignJustify()
                param = self.hwp.HParameterSet.HShapeObject
                self.hwp.HAction.GetDefault("TablePropertyDialog", param.HSet)
                param.ShapeCaption.Side = self.hwp.SideType(location)
                self.hwp.HAction.Execute("TablePropertyDialog", param.HSet)
            ctrl = ctrl.Next
        self.hwp.SetPos(*start_pos)
        return True

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

    # 액션 파라미터용 함수

    def arc_type(self, arc_type):
        return self.hwp.ArcType(ArcType=arc_type)

    def ArcType(self, arc_type):
        return self.hwp.ArcType(ArcType=arc_type)

    def auto_num_type(self, autonum):
        return self.hwp.AutoNumType(autonum=autonum)

    def AutoNumType(self, autonum):
        return self.hwp.AutoNumType(autonum=autonum)

    def border_shape(self, border_type):
        return self.hwp.BorderShape(BorderType=border_type)

    def BorderShape(self, border_type):
        return self.hwp.BorderShape(BorderType=border_type)

    def break_word_latin(self, break_latin_word):
        return self.hwp.BreakWordLatin(BreakLatinWord=break_latin_word)

    def BreakWordLatin(self, break_latin_word):
        return self.hwp.BreakWordLatin(BreakLatinWord=break_latin_word)

    def brush_type(self, brush_type):
        return self.hwp.BrushType(BrushType=brush_type)

    def BrushType(self, brush_type):
        return self.hwp.BrushType(BrushType=brush_type)

    def canonical(self, canonical):
        return self.hwp.Canonical(Canonical=canonical)

    def Canonical(self, canonical):
        return self.hwp.Canonical(Canonical=canonical)

    def cell_apply(self, cell_apply):
        return self.hwp.CellApply(CellApply=cell_apply)

    def CellApply(self, cell_apply):
        return self.hwp.CellApply(CellApply=cell_apply)

    def char_shadow_type(self, shadow_type):
        return self.hwp.CharShadowType(ShadowType=shadow_type)

    def CharShadowType(self, shadow_type):
        return self.hwp.CharShadowType(ShadowType=shadow_type)

    def check_xobject(self, bstring):
        return self.hwp.CheckXObject(bstring=bstring)

    def CheckXObject(self, bstring):
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

    def Clear(self, option: int = 1):
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

    # Run 액션아이디의 Close와 중복되어 주석처리함. close로 실행가능
    # def Close(self, is_dirty: bool = False):
    #     return self.hwp.XHwpDocuments.Active_XHwpDocument.Close(isDirty=is_dirty)

    def col_def_type(self, col_def_type):
        return self.hwp.ColDefType(ColDefType=col_def_type)

    def ColDefType(self, col_def_type):
        return self.hwp.ColDefType(ColDefType=col_def_type)

    def col_layout_type(self, col_layout_type):
        return self.hwp.ColLayoutType(ColLayoutType=col_layout_type)

    def ColLayoutType(self, col_layout_type):
        return self.hwp.ColLayoutType(ColLayoutType=col_layout_type)

    def convert_pua_hangul_to_unicode(self, reverse):
        return self.hwp.ConvertPUAHangulToUnicode(Reverse=reverse)

    def ConvertPUAHangulToUnicode(self, reverse):
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
            >>> act = hwp.create_action("CharShape")
            >>> cs = act.CreateSet()  # equal to "cs = hwp.create_set(act)"
            >>> act.GetDefault(cs)
            >>> print(cs.Item("Height"))
            2800
            >>> # 현재 선택범위의 폰트 크기를 20pt로 변경하는 코드
            >>> act = hwp.create_action("CharShape")
            >>> cs = act.CreateSet()  # equal to "cs = hwp.create_set(act)"
            >>> act.GetDefault(cs)
            >>> cs.SetItem("Height", hwp.point_to_hwp_unit(20))
            >>> act.Execute(cs)
            True

        """
        return self.hwp.CreateAction(actidstr=actidstr)

    def CreateAction(self, actidstr: str):
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
            >>> act = hwp.create_action("CharShape")
            >>> cs = act.CreateSet()  # equal to "cs = hwp.create_set(act)"
            >>> act.GetDefault(cs)
            >>> print(cs.Item("Height"))
            2800
            >>> # 현재 선택범위의 폰트 크기를 20pt로 변경하는 코드
            >>> act = hwp.create_action("CharShape")
            >>> cs = act.CreateSet()  # equal to "cs = hwp.create_set(act)"
            >>> act.GetDefault(cs)
            >>> cs.SetItem("Height", hwp.point_to_hwp_unit(20))
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_field(direction="이름", memo="이름을 입력하는 필드", name="name")
            True
            >>> hwp.put_field_text("name", "일코")
        """
        return self.hwp.CreateField(Direction=direction, memo=memo, name=name)

    def CreateField(self, name: str, direction: str = "", memo: str = "") -> bool:
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_field(direction="이름", memo="이름을 입력하는 필드", name="name")
            True
            >>> hwp.put_field_text("name", "일코")
        """
        return self.hwp.CreateField(Direction=direction, memo=memo, name=name)

    def create_id(self, creation_id):
        return self.hwp.CreateID(CreationID=creation_id)

    def CreateId(self, creation_id):
        return self.hwp.CreateID(CreationID=creation_id)

    def create_mode(self, creation_mode):
        return self.hwp.CreateMode(CreationMode=creation_mode)

    def CreateMode(self, creation_mode):
        return self.hwp.CreateMode(CreationMode=creation_mode)

    def create_page_image(self, path: str, pgno: int = -1, resolution: int = 300, depth: int = 24,
                          format: str = "bmp") -> bool:
        """
        pgno로 지정한 페이지를 path 라는 파일명으로 저장한다.
        이 때 페이지번호는 1부터 시작하며,(1-index)
        pgno=0이면 현재 페이지, pgno=-1(기본값)이면 전체 페이지를 이미지로 저장한다.
        내부적으로 Pillow 모듈을 사용하여 변환하므로,
        사실상 Pillow에서 변환 가능한 모든 포맷으로 입력 가능하다.

        :param path:
            생성할 이미지 파일의 경로(전체경로로 입력해야 함)

        :param pgno:
            페이지 번호(1페이지 저장하려면 pgno=1).
            1부터 hwp.PageCount 사이에서 pgno 입력시 선택한 페이지만 저장한다.
            생략하면(기본값은 -1) 전체 페이지가 저장된다.
            이 때 path가 "img.jpg"라면 저장되는 파일명은
            "img001.jpg", "img002.jpg", "img003.jpg",..,"img099.jpg" 가 된다.

            현재 캐럿이 있는 페이지만 저장하고 싶을 때에는 pgno=0으로 설정하면 된다.



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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_page_image("c:/Users/User/Desktop/a.bmp")
            True
        """
        if pgno < -1 or pgno > self.PageCount:
            raise IndexError(f"pgno는 -1부터 {self.PageCount}까지 입력 가능합니다. (-1:전체 저장, 0:현재페이지 저장)")
        if path.lower()[1] != ":":
            path = os.path.abspath(path)
        if not os.path.exists(os.path.dirname(path)):
            os.mkdir(os.path.dirname(path))
        ext = path.rsplit(".", maxsplit=1)[-1]
        if pgno >= 0:
            if pgno == 0:
                pgno = self.current_page
            try:
                return self.hwp.CreatePageImage(Path=path, pgno=pgno - 1, resolution=resolution, depth=depth,
                                                Format=format)
            finally:
                if not ext.lower() in ("gif", "bmp"):
                    with Image.open(path.replace(ext, format)) as img:
                        img.save(path.replace(format, ext))
                    os.remove(path.replace(ext, format))
        elif pgno == -1:
            for i in range(1, self.PageCount + 1):
                path_ = os.path.join(os.path.dirname(path), os.path.basename(path).replace(f".{ext}", f"{i:03}.{ext}"))
                self.hwp.CreatePageImage(Path=path_, pgno=i - 1, resolution=resolution, depth=depth, Format=format)
                if not ext.lower() in ("gif", "bmp"):
                    with Image.open(path_.replace(ext, format)) as img:
                        img.save(path_.replace(format, ext))
                    os.remove(path_.replace(ext, format))
            return True

    def CreatePageImage(self, path: str, pgno: int = -1, resolution: int = 300, depth: int = 24,
                        format: str = "bmp") -> bool:
        """
        pgno로 지정한 페이지를 path 라는 파일명으로 저장한다.
        이 때 페이지번호는 1부터 시작하며,(1-index)
        pgno=0이면 현재 페이지, pgno=-1(기본값)이면 전체 페이지를 이미지로 저장한다.
        내부적으로 Pillow 모듈을 사용하여 변환하므로,
        사실상 Pillow에서 변환 가능한 모든 포맷으로 입력 가능하다.

        :param path:
            생성할 이미지 파일의 경로(전체경로로 입력해야 함)

        :param pgno:
            페이지 번호(1페이지 저장하려면 pgno=1).
            1부터 hwp.PageCount 사이에서 pgno 입력시 선택한 페이지만 저장한다.
            생략하면(기본값은 -1) 전체 페이지가 저장된다.
            이 때 path가 "img.jpg"라면 저장되는 파일명은
            "img001.jpg", "img002.jpg", "img003.jpg",..,"img099.jpg" 가 된다.

            현재 캐럿이 있는 페이지만 저장하고 싶을 때에는 pgno=0으로 설정하면 된다.


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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_page_image("c:/Users/User/Desktop/a.bmp")
            True
        """
        if pgno < -1 or pgno > self.PageCount:
            raise IndexError(f"pgno는 -1부터 {self.PageCount}까지 입력 가능합니다. (-1:전체 저장, 0:현재페이지 저장)")
        if path.lower()[1] != ":":
            path = os.path.abspath(path)
        if not os.path.exists(os.path.dirname(path)):
            os.mkdir(os.path.dirname(path))
        ext = path.rsplit(".", maxsplit=1)[-1]
        if pgno >= 0:
            if pgno == 0:
                pgno = self.current_page
            try:
                return self.hwp.CreatePageImage(Path=path, pgno=pgno - 1, resolution=resolution, depth=depth,
                                                Format=format)
            finally:
                if not ext.lower() in ("gif", "bmp"):
                    with Image.open(path.replace(ext, format)) as img:
                        img.save(path.replace(format, ext))
                    os.remove(path.replace(ext, format))
        elif pgno == -1:
            for i in range(1, self.PageCount + 1):
                path_ = os.path.join(os.path.dirname(path), os.path.basename(path).replace(f".{ext}", f"{i:03}.{ext}"))
                self.hwp.CreatePageImage(Path=path_, pgno=i - 1, resolution=resolution, depth=depth, Format=format)
                if not ext.lower() in ("gif", "bmp"):
                    with Image.open(path_.replace(ext, format)) as img:
                        img.save(path_.replace(format, ext))
                    os.remove(path_.replace(ext, format))
            return True

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

    def CreateSet(self, setidstr):
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

    def CrookedSlash(self, crooked_slash):
        return self.hwp.CrookedSlash(CrookedSlash=crooked_slash)

    def ds_mark(self, diac_sym_mark):
        return self.hwp.DSMark(DiacSymMark=diac_sym_mark)

    def DSMark(self, diac_sym_mark):
        return self.hwp.DSMark(DiacSymMark=diac_sym_mark)

    def dbf_code_type(self, dbf_code):
        return self.hwp.DbfCodeType(DbfCode=dbf_code)

    def DbfCodeType(self, dbf_code):
        return self.hwp.DbfCodeType(DbfCode=dbf_code)

    def delete_ctrl(self, ctrl) -> bool:
        """
        문서 내 컨트롤을 삭제한다.

        :param ctrl:
            삭제할 문서 내 컨트롤

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> ctrl = hwp.HeadCtrl.Next.Next
            >>> if ctrl.UserDesc == "표":
            ...     hwp.delete_ctrl(ctrl)
            ...
            True
        """
        return self.hwp.DeleteCtrl(ctrl=ctrl)

    def DeleteCtrl(self, ctrl) -> bool:
        """
        문서 내 컨트롤을 삭제한다.

        :param ctrl:
            삭제할 문서 내 컨트롤

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> ctrl = hwp.HeadCtrl.Next.Next
            >>> if ctrl.UserDesc == "표":
            ...     hwp.delete_ctrl(ctrl)
            ...
            True
        """
        return self.hwp.DeleteCtrl(ctrl=ctrl)

    def delimiter(self, delimiter):
        return self.hwp.Delimiter(Delimiter=delimiter)

    def Delimiter(self, delimiter):
        return self.hwp.Delimiter(Delimiter=delimiter)

    def draw_aspect(self, draw_aspect):
        return self.hwp.DrawAspect(DrawAspect=draw_aspect)

    def DrawAspect(self, draw_aspect):
        return self.hwp.DrawAspect(DrawAspect=draw_aspect)

    def draw_fill_image(self, fillimage):
        return self.hwp.DrawFillImage(fillimage=fillimage)

    def DrawFillImage(self, fillimage):
        return self.hwp.DrawFillImage(fillimage=fillimage)

    def draw_shadow_type(self, shadow_type):
        return self.hwp.DrawShadowType(ShadowType=shadow_type)

    def DrawShadowType(self, shadow_type):
        return self.hwp.DrawShadowType(ShadowType=shadow_type)

    def encrypt(self, encrypt):
        return self.hwp.Encrypt(Encrypt=encrypt)

    def Encrypt(self, encrypt):
        return self.hwp.Encrypt(Encrypt=encrypt)

    def end_size(self, end_size):
        return self.hwp.EndSize(EndSize=end_size)

    def EndSize(self, end_size):
        return self.hwp.EndSize(EndSize=end_size)

    def end_style(self, end_style):
        return self.hwp.EndStyle(EndStyle=end_style)

    def EndStyle(self, end_style):
        return self.hwp.EndStyle(EndStyle=end_style)

    def export_style(self, sty_filepath: str) -> bool:
        """
        현재 문서의 Style을 sty 파일로 Export한다.

        :param sty_filepath:
            Export할 sty 파일의 전체경로 문자열

        :return:
            성공시 True, 실패시 False

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.export_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if sty_filepath.lower()[1] != ":":
            sty_filepath = os.path.join(os.getcwd(), sty_filepath)

        style_set = self.hwp.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.hwp.ExportStyle(param=style_set.HSet)

    def ExportStyle(self, sty_filepath: str) -> bool:
        """
        현재 문서의 Style을 sty 파일로 Export한다.

        :param sty_filepath:
            Export할 sty 파일의 전체경로 문자열

        :return:
            성공시 True, 실패시 False

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.export_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if sty_filepath.lower()[1] != ":":
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

    def FieldExist(self, field):
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

    def FileTranslate(self, cur_lang, trans_lang):
        return self.hwp.FileTranslate(curLang=cur_lang, transLang=trans_lang)

    def fill_area_type(self, fill_area):
        return self.hwp.FillAreaType(FillArea=fill_area)

    def FillAreaType(self, fill_area):
        return self.hwp.FillAreaType(FillArea=fill_area)

    def find_ctrl(self):
        return self.hwp.FindCtrl()

    def FindCtrl(self):
        return self.hwp.FindCtrl()

    def find_dir(self, find_dir: Literal["Forward", "Backward", "AllDoc"] = "AllDoc"):
        return self.hwp.FindDir(FindDir=find_dir)

    def FindDir(self, find_dir: Literal["Forward", "Backward", "AllDoc"] = "AllDoc"):
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

    def FindPrivateInfo(self, private_type, private_string):
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

    def FontType(self, font_type):
        return self.hwp.FontType(FontType=font_type)

    def get_bin_data_path(self, binid):
        """
        Binary Data(Temp Image 등)의 경로를 가져온다.

        :param binid:
            바이너리 데이터의 ID 값 (1부터 시작)

        :return:
            바이너리 데이터의 경로

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> path = hwp.get_bin_data_path(2)
            >>> print(path)
            C:/Users/User/AppData/Local/Temp/Hnc/BinData/EMB00004dd86171.jpg
        """
        return self.hwp.GetBinDataPath(binid=binid)

    def GetBinDataPath(self, binid):
        """
        Binary Data(Temp Image 등)의 경로를 가져온다.

        :param binid:
            바이너리 데이터의 ID 값 (1부터 시작)

        :return:
            바이너리 데이터의 경로

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> path = hwp.get_bin_data_path(2)
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

    def GetCurFieldName(self, option=0):
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

    def GetCurMetatagName(self):
        return self.hwp.GetCurMetatagName()

    def get_field_list(self, number=1, option=0):
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

    def GetFieldList(self, number=1, option=0):
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

    def get_field_text(self, field: str | list | tuple | set, idx=0):
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
        if isinstance(field, str):
            if idx and "{{" not in field:
                return self.hwp.GetFieldText(Field=field + f"{{{{{idx}}}}}")
            else:
                return self.hwp.GetFieldText(Field=field)
        elif isinstance(field, list | tuple | set):
            return self.hwp.GetFieldText(Field="\x02".join(str(i) for i in field))

    def GetFieldText(self, field: str | list | tuple | set, idx=0):
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
        if isinstance(field, str):
            if idx and "{{" not in field:
                return self.hwp.GetFieldText(Field=field + f"{{{{{idx}}}}}")
            else:
                return self.hwp.GetFieldText(Field=field)
        elif isinstance(field, list | tuple | set):
            return self.hwp.GetFieldText(Field="\x02".join(str(i) for i in field))

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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.get_file_info("C:/Users/Administrator/Desktop/이력서.hwp")
            >>> print(pset.Item("Format"))
            >>> print(pset.Item("VersionStr"))
            >>> print(hex(pset.Item("VersionNum")))
            >>> print(pset.Item("Encrypted"))
            HWP
            5.1.1.0
            0x5010100
            0
        """
        if filename.lower()[1] != ":":
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.GetFileInfo(filename=filename)

    def GetFileInfo(self, filename):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.get_file_info("C:/Users/Administrator/Desktop/이력서.hwp")
            >>> print(pset.Item("Format"))
            >>> print(pset.Item("VersionStr"))
            >>> print(hex(pset.Item("VersionNum")))
            >>> print(pset.Item("Encrypted"))
            HWP
            5.1.1.0
            0x5010100
            0
        """
        if filename.lower()[1] != ":":
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.GetFileInfo(filename=filename)

    def get_font_list(self, langid=""):
        self.scan_font()
        return self.hwp.GetFontList(langid=langid)

    def GetFontList(self, langid=""):
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

    def GetHeadingString(self):
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

    def GetMessageBoxMode(self):
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

    def GetMetatagList(self, number, option):
        return self.hwp.GetMetatagList(Number=number, option=option)

    def get_metatag_name_text(self, tag):
        return self.hwp.GetMetatagNameText(tag=tag)

    def GetMetatagNameText(self, tag):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.get_mouse_pos(1, 1)
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

    def GetMousePos(self, x_rel_to=1, y_rel_to=1):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.get_mouse_pos(1, 1)
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

    def GetPageText(self, pgno: int = 0, option: hex = 0xffffffff) -> str:
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

    def GetPos(self) -> tuple[int]:
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.get_pos_by_set()  # 캐럿위치 저장
            >>> print(pset.Item("List"))
            6
            >>> print(pset.Item("Para"))
            3
            >>> print(pset.Item("Pos"))
            2
            >>> hwp.set_pos_by_set(pset)  # 캐럿위치 복원
            True
        """
        return self.hwp.GetPosBySet()

    def GetPosBySet(self):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.get_pos_by_set()  # 캐럿위치 저장
            >>> print(pset.Item("List"))
            6
            >>> print(pset.Item("Para"))
            3
            >>> print(pset.Item("Pos"))
            2
            >>> hwp.set_pos_by_set(pset)  # 캐럿위치 복원
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
        if filename.lower()[1] != ":":
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.GetScriptSource(filename=filename)

    def GetScriptSource(self, filename: str) -> str:
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
        if filename.lower()[1] != ":":
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_selected_pos()
            (True, 0, 0, 16, 0, 7, 16)
        """
        return self.hwp.GetSelectedPos()

    def GetSelectedPos(self):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_selected_pos()
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> sset = hwp.get_pos_by_set()
            >>> eset = hwp.get_pos_by_set()
            >>> hwp.get_selected_pos_by_set(sset, eset)
            >>> hwp.set_pos_by_set(eset)
            True
        """
        return self.hwp.GetSelectedPosBySet(sset=sset, eset=eset)

    def GetSelectedPosBySet(self, sset, eset):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> sset = hwp.get_pos_by_set()
            >>> eset = hwp.get_pos_by_set()
            >>> hwp.get_selected_pos_by_set(sset, eset)
            >>> hwp.set_pos_by_set(eset)
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.init_scan()
            >>> while True:
            ...     state, text = hwp.get_text()
            ...     print(state, text)
            ...     if state <= 1:
            ...         break
            ... hwp.release_scan()
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

    def GetText(self):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.init_scan()
            >>> while True:
            ...     state, text = hwp.get_text()
            ...     print(state, text)
            ...     if state <= 1:
            ...         break
            ... hwp.release_scan()
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_text_file()
            'ㅁㄴㅇㄹ\r\nㅁㄴㅇㄹ\r\nㅁㄴㅇㄹ\r\n\r\nㅂㅈㄷㄱ\r\nㅂㅈㄷㄱ\r\nㅂㅈㄷㄱ\r\n'
        """
        return self.hwp.GetTextFile(Format=format, option=option)

    def GetTextFile(self, format="UNICODE", option=""):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_text_file()
            'ㅁㄴㅇㄹ\r\nㅁㄴㅇㄹ\r\nㅁㄴㅇㄹ\r\n\r\nㅂㅈㄷㄱ\r\nㅂㅈㄷㄱ\r\nㅂㅈㄷㄱ\r\n'
        """
        return self.hwp.GetTextFile(Format=format, option=option)

    def get_translate_lang_list(self, cur_lang):
        return self.hwp.GetTranslateLangList(curLang=cur_lang)

    def GetTranslateLangList(self, cur_lang):
        return self.hwp.GetTranslateLangList(curLang=cur_lang)

    def get_user_info(self, user_info_id):
        return self.hwp.GetUserInfo(userInfoId=user_info_id)

    def GetUserInfo(self, user_info_id):
        return self.hwp.GetUserInfo(userInfoId=user_info_id)

    def gradation(self, gradation):
        return self.hwp.Gradation(Gradation=gradation)

    def Gradation(self, gradation):
        return self.hwp.Gradation(Gradation=gradation)

    def grid_method(self, grid_method):
        return self.hwp.GridMethod(GridMethod=grid_method)

    def GridMethod(self, grid_method):
        return self.hwp.GridMethod(GridMethod=grid_method)

    def grid_view_line(self, grid_view_line):
        return self.hwp.GridViewLine(GridViewLine=grid_view_line)

    def GridViewLine(self, grid_view_line):
        return self.hwp.GridViewLine(GridViewLine=grid_view_line)

    def gutter_method(self, gutter_type):
        return self.hwp.GutterMethod(GutterType=gutter_type)

    def GutterMethod(self, gutter_type):
        return self.hwp.GutterMethod(GutterType=gutter_type)

    def h_align(self, h_align):
        return self.hwp.HAlign(HAlign=h_align)

    def HAlign(self, h_align):
        return self.hwp.HAlign(HAlign=h_align)

    def handler(self, handler):
        return self.hwp.Handler(Handler=handler)

    def Handler(self, handler):
        return self.hwp.Handler(Handler=handler)

    def hash(self, hash):
        return self.hwp.Hash(Hash=hash)

    def Hash(self, hash):
        return self.hwp.Hash(Hash=hash)

    def hatch_style(self, hatch_style):
        return self.hwp.HatchStyle(HatchStyle=hatch_style)

    def HatchStyle(self, hatch_style):
        return self.hwp.HatchStyle(HatchStyle=hatch_style)

    def head_type(self, heading_type):
        return self.hwp.HeadType(HeadingType=heading_type)

    def HeadType(self, heading_type):
        return self.hwp.HeadType(HeadingType=heading_type)

    def height_rel(self, height_rel):
        return self.hwp.HeightRel(HeightRel=height_rel)

    def HeightRel(self, height_rel):
        return self.hwp.HeightRel(HeightRel=height_rel)

    def hiding(self, hiding):
        return self.hwp.Hiding(Hiding=hiding)

    def Hiding(self, hiding):
        return self.hwp.Hiding(Hiding=hiding)

    def horz_rel(self, horz_rel):
        return self.hwp.HorzRel(HorzRel=horz_rel)

    def HorzRel(self, horz_rel):
        return self.hwp.HorzRel(HorzRel=horz_rel)

    def hwp_line_type(self, line_type: Literal[
        "None", "Solid", "Dash", "Dot", "DashDot", "DashDotDot", "LongDash", "Circle", "DoubleSlim", "SlimThick", "ThickSlim", "SlimThickSlim"] = "Solid"):
        """
        "None": 없음(0)
        "Solid": 실선(1)
        "Dash": 파선(2)
        "Dot": 점선(3)
        "DashDot": 일점쇄선(4)
        "DashDotDot": 이점쇄선(5)
        "LongDash": 긴 파선(6)
        "Circle": 원형 점선(7)
        "DoubleSlim": 이중 실선(8)
        "SlimThick": 얇고 굵은 이중선(9)
        "ThickSlim": 굵고 얇은 이중선(10)
        "SlimThickSlim": 얇고 굵고 얇은 삼중선(11)
        """
        return self.hwp.HwpLineType(LineType=line_type)

    def HwpLineType(self, line_type: Literal[
        "None", "Solid", "Dash", "Dot", "DashDot", "DashDotDot", "LongDash", "Circle", "DoubleSlim", "SlimThick", "ThickSlim", "SlimThickSlim"] = "Solid"):
        """
        "None": 없음(0)
        "Solid": 실선(1)
        "Dash": 파선(2)
        "Dot": 점선(3)
        "DashDot": 일점쇄선(4)
        "DashDotDot": 이점쇄선(5)
        "LongDash": 긴 파선(6)
        "Circle": 원형 점선(7)
        "DoubleSlim": 이중 실선(8)
        "SlimThick": 얇고 굵은 이중선(9)
        "ThickSlim": 굵고 얇은 이중선(10)
        "SlimThickSlim": 얇고 굵고 얇은 삼중선(11)
        """
        return self.hwp.HwpLineType(LineType=line_type)

    def hwp_line_width(self, line_width: Literal[
        "0.1mm", "0.12mm", "0.15mm", "0.2mm", "0.25mm", "0.3mm", "0.4mm", "0.5mm", "0.6mm", "0.7mm", "1.0mm", "1.5mm", "2.0mm", "3.0mm", "4.0mm", "5.0mm"] = "0.1mm"):
        """
            "0.1mm"(0)
            "0.12mm"(1)
            "0.15mm"(2)
            "0.2mm"(3)
            "0.25mm"(4)
            "0.3mm"(5)
            "0.4mm"(6)
            "0.5mm"(7)
            "0.6mm"(8)
            "0.7mm"(9)
            "1.0mm"(10)
            "1.5mm"(11)
            "2.0mm"(12)
            "3.0mm"(13)
            "4.0mm"(14)
            "5.0mm"(15)
            """
        return self.hwp.HwpLineWidth(LineWidth=line_width)

    def HwpLineWidth(self, line_width: Literal[
        "0.1mm", "0.12mm", "0.15mm", "0.2mm", "0.25mm", "0.3mm", "0.4mm", "0.5mm", "0.6mm", "0.7mm", "1.0mm", "1.5mm", "2.0mm", "3.0mm", "4.0mm", "5.0mm"] = "0.1mm"):
        """
            "0.1mm"(0)
            "0.12mm"(1)
            "0.15mm"(2)
            "0.2mm"(3)
            "0.25mm"(4)
            "0.3mm"(5)
            "0.4mm"(6)
            "0.5mm"(7)
            "0.6mm"(8)
            "0.7mm"(9)
            "1.0mm"(10)
            "1.5mm"(11)
            "2.0mm"(12)
            "3.0mm"(13)
            "4.0mm"(14)
            "5.0mm"(15)
            """
        return self.hwp.HwpLineWidth(LineWidth=line_width)

    def hwp_outline_style(self, hwp_outline_style):
        return self.hwp.HwpOutlineStyle(HwpOutlineStyle=hwp_outline_style)

    def HwpOutlineStyle(self, hwp_outline_style):
        return self.hwp.HwpOutlineStyle(HwpOutlineStyle=hwp_outline_style)

    def hwp_outline_type(self, hwp_outline_type):
        return self.hwp.HwpOutlineType(HwpOutlineType=hwp_outline_type)

    def HwpOutlineType(self, hwp_outline_type):
        return self.hwp.HwpOutlineType(HwpOutlineType=hwp_outline_type)

    def hwp_underline_shape(self, hwp_underline_shape):
        return self.hwp.HwpUnderlineShape(HwpUnderlineShape=hwp_underline_shape)

    def HwpUnderlineShape(self, hwp_underline_shape):
        return self.hwp.HwpUnderlineShape(HwpUnderlineShape=hwp_underline_shape)

    def hwp_underline_type(self, hwp_underline_type):
        return self.hwp.HwpUnderlineType(HwpUnderlineType=hwp_underline_type)

    def HwpUnderlineType(self, hwp_underline_type):
        return self.hwp.HwpUnderlineType(HwpUnderlineType=hwp_underline_type)

    def hwp_zoom_type(self, zoom_type):
        return self.hwp.HwpZoomType(ZoomType=zoom_type)

    def HwpZoomType(self, zoom_type):
        return self.hwp.HwpZoomType(ZoomType=zoom_type)

    def image_format(self, image_format):
        return self.hwp.ImageFormat(ImageFormat=image_format)

    def ImageFormat(self, image_format):
        return self.hwp.ImageFormat(ImageFormat=image_format)

    def import_style(self, sty_filepath):
        """
        미리 저장된 특정 sty파일의 스타일을 임포트한다.

        :param sty_filepath:
            sty파일의 경로

        :return:
            성공시 True, 실패시 False

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.import_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if sty_filepath.lower()[1] != ":":
            sty_filepath = os.path.join(os.getcwd(), sty_filepath)

        style_set = self.hwp.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.hwp.ImportStyle(style_set.HSet)

    def ImportStyle(self, sty_filepath):
        """
        미리 저장된 특정 sty파일의 스타일을 임포트한다.

        :param sty_filepath:
            sty파일의 경로A

        :return:
            성공시 True, 실패시 False

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.import_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if sty_filepath.lower()[1] != ":":
            sty_filepath = os.path.join(os.getcwd(), sty_filepath)

        style_set = self.hwp.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.hwp.ImportStyle(style_set.HSet)

    def init_hparameter_set(self):
        return self.hwp.InitHParameterSet()

    def InitHParameterSet(self):
        return self.hwp.InitHParameterSet()

    def init_scan(self, option=0x07, range=0x77, spara=0, spos=0, epara=-1, epos=-1):
        """
        문서의 내용을 검색하기 위해 초기설정을 한다.
        문서의 검색 과정은 InitScan()으로 검색위한 준비 작업을 하고
        GetText()를 호출하여 본문의 텍스트를 얻어온다.
        GetText()를 반복호출하면 연속하여 본문의 텍스트를 얻어올 수 있다.
        검색이 끝나면 ReleaseScan()을 호출하여 관련 정보를 Release해야 한다.

        :param option: 기본값은 0x7(모든 컨트롤 대상)
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.init_scan(range=0xff)
            >>> _, text = hwp.get_text()
            >>> hwp.release_scan()
            >>> print(text)
            Hello, world!
        """
        return self.hwp.InitScan(option=option, Range=range, spara=spara, spos=spos, epara=epara, epos=epos)

    def InitScan(self, option=0x07, range=0x77, spara=0, spos=0, epara=-1, epos=-1):
        """
        문서의 내용을 검색하기 위해 초기설정을 한다.
        문서의 검색 과정은 InitScan()으로 검색위한 준비 작업을 하고
        GetText()를 호출하여 본문의 텍스트를 얻어온다.
        GetText()를 반복호출하면 연속하여 본문의 텍스트를 얻어올 수 있다.
        검색이 끝나면 ReleaseScan()을 호출하여 관련 정보를 Release해야 한다.

        :param option: 기본값은 0x7(모든 컨트롤 대상)
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.init_scan(range=0xff)
            >>> _, text = hwp.get_text()
            >>> hwp.release_scan()
            >>> print(text)
            Hello, world!
        """
        return self.hwp.InitScan(option=option, Range=range, spara=spara, spos=spos, epara=epara, epos=epos)

    def insert(self, path, format="", arg="", move_doc_end=False):
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
        if path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)
        try:
            return self.hwp.Insert(Path=path, Format=format, arg=arg)
        finally:
            if move_doc_end:
                self.MoveDocEnd()

    def insert_background_picture(self, path,
                                  border_type: Literal["SelectedCell", "SelectedCellDelete"] = "SelectedCell",
                                  embedded=True, filloption=5, effect=0, watermark=False, brightness=0,
                                  contrast=0) -> bool:
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_background_picture(path="C:/Users/User/Desktop/KakaoTalk_20230709_023118549.jpg")
            True
        """
        if path.startswith("http"):
            request.urlretrieve(path, os.path.join(os.getcwd(), "temp.jpg"))
            path = os.path.join(os.getcwd(), "temp.jpg")
        elif path and path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)

        try:
            return self.hwp.InsertBackgroundPicture(Path=path, BorderType=border_type, Embedded=embedded,
                                                    filloption=filloption, Effect=effect, watermark=watermark,
                                                    Brightness=brightness, Contrast=contrast)
        finally:
            if "temp.jpg" in os.listdir():
                os.remove(path)

    def InsertBackgroundPicture(self, path, border_type: Literal["SelectedCell", "SelectedCellDelete"] = "SelectedCell",
                                embedded=True, filloption=5, effect=0, watermark=False, brightness=0,
                                contrast=0) -> bool:
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_background_picture(path="C:/Users/User/Desktop/KakaoTalk_20230709_023118549.jpg")
            True
        """
        if path.startswith("http"):
            request.urlretrieve(path, os.path.join(os.getcwd(), "temp.jpg"))
            path = os.path.join(os.getcwd(), "temp.jpg")
        elif path and path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)

        try:
            return self.hwp.InsertBackgroundPicture(Path=path, BorderType=border_type, Embedded=embedded,
                                                    filloption=filloption, Effect=effect, watermark=watermark,
                                                    Brightness=brightness, Contrast=contrast)
        finally:
            if "temp.jpg" in os.listdir():
                os.remove(path)

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
            >>> from pyhwpx import Hwp
            >>> from time import sleep
            >>> hwp = Hwp()
            >>> tbset = hwp.create_set("TableCreation")
            >>> tbset.SetItem("Rows", 3)
            >>> tbset.SetItem("Cols", 5)
            >>> row_set = tbset.CreateItemArray("RowHeight", 3)
            >>> col_set = tbset.CreateItemArray("ColWidth", 5)
            >>> row_set.SetItem(0, hwp.mili_to_hwp_unit(10))
            >>> row_set.SetItem(1, hwp.mili_to_hwp_unit(10))
            >>> row_set.SetItem(2, hwp.mili_to_hwp_unit(10))
            >>> col_set.SetItem(0, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(1, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(2, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(3, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(4, hwp.mili_to_hwp_unit(26))
            >>> table = hwp.insert_ctrl("tbl", tbset)
            >>> sleep(3)  # 표 생성 3초 후 다시 표 삭제
            >>> hwp.delete_ctrl(table)


        """
        return self.hwp.InsertCtrl(CtrlID=ctrl_id, initparam=initparam)

    def InsertCtrl(self, ctrl_id, initparam):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> tbset = hwp.create_set("TableCreation")
            >>> tbset.SetItem("Rows", 3)
            >>> tbset.SetItem("Cols", 5)
            >>> row_set = tbset.CreateItemArray("RowHeight", 3)
            >>> col_set = tbset.CreateItemArray("ColWidth", 5)
            >>> row_set.SetItem(0, hwp.mili_to_hwp_unit(10))
            >>> row_set.SetItem(1, hwp.mili_to_hwp_unit(10))
            >>> row_set.SetItem(2, hwp.mili_to_hwp_unit(10))
            >>> col_set.SetItem(0, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(1, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(2, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(3, hwp.mili_to_hwp_unit(26))
            >>> col_set.SetItem(4, hwp.mili_to_hwp_unit(26))
            >>> table = hwp.insert_ctrl("tbl", tbset)
            >>> sleep(3)  # 표 생성 3초 후 다시 표 삭제
            >>> hwp.delete_ctrl(table)


        """
        return self.hwp.InsertCtrl(CtrlID=ctrl_id, initparam=initparam)

    def insert_picture(self, path, treat_as_char=True, embedded=True, sizeoption=0, reverse=False, watermark=False,
                       effect=0, width=0, height=0):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> path = "C:/Users/Administrator/Desktop/KakaoTalk_20230709_023118549.jpg"
            >>> ctrl = hwp.insert_picture(path)  # 삽입한 이미지 객체를 리턴함.
            >>> pset = ctrl.Properties  # == hwp.create_set("ShapeObject")
            >>> pset.SetItem("TreatAsChar", False)  # 글자처럼취급 해제
            >>> pset.SetItem("TextWrap", 2)  # 그림을 글 뒤로
            >>> ctrl.Properties = pset  # 설정한 값 적용(간단!)
        """
        if sizeoption == 1 and not all([width, height]) and not self.is_cell():
            raise ValueError("sizeoption이 1일 때에는 width와 height를 지정해주셔야 합니다.\n"
                             "단, 셀 안에 있는 경우에는 셀 너비에 맞게 이미지 크기를 자동으로 조절합니다.")

        if path.startswith("http"):
            temp_path = tempfile.TemporaryFile().name
            request.urlretrieve(path, temp_path)
            path = temp_path
            # request.urlretrieve(path, os.path.join(os.getcwd(), "temp.jpg"))
        elif path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)

        try:
            ctrl = self.hwp.InsertPicture(Path=path, Embedded=embedded, sizeoption=sizeoption, Reverse=reverse,
                                          watermark=watermark, Effect=effect, Width=width, Height=height)
            pic_prop = ctrl.Properties
            if not all([width, height]) and self.is_cell():
                cell_param = self.hwp.HParameterSet.HShapeObject
                self.hwp.HAction.GetDefault("TablePropertyDialog", cell_param.HSet)
                cell_width = cell_param.ShapeTableCell.Width
                dst_height = pic_prop.Item("Height") / pic_prop.Item("Width") * cell_width
                pic_prop.SetItem("Width", cell_width)
                pic_prop.SetItem("Height", round(dst_height))
            else:
                sec_def = self.HParameterSet.HSecDef
                self.HAction.GetDefault("PageSetup", sec_def.HSet)
                page_width = (
                        sec_def.PageDef.PaperWidth - sec_def.PageDef.LeftMargin - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen)
                page_height = (
                        sec_def.PageDef.PaperHeight - sec_def.PageDef.TopMargin - sec_def.PageDef.BottomMargin - sec_def.PageDef.HeaderLen - sec_def.PageDef.FooterLen)
                pic_width = pic_prop.Item("Width")
                pic_height = pic_prop.Item("Height")
                if pic_width > page_width or pic_height > page_height:
                    width_shrink_ratio = page_width / pic_width
                    height_shrink_ratio = page_height / pic_height
                    if width_shrink_ratio <= height_shrink_ratio:
                        pic_prop.SetItem("Width", page_width)
                        pic_prop.SetItem("Height", pic_height * width_shrink_ratio)
                    else:
                        pic_prop.SetItem("Width", pic_width * height_shrink_ratio)
                        pic_prop.SetItem("Height", page_height)
            pic_prop.SetItem("TreatAsChar", treat_as_char)
            ctrl.Properties = pic_prop
            return ctrl
        finally:
            if os.path.basename(path).startswith("tmp"):
                os.remove(path)

    def insert_picture(self, path, treat_as_char=True, embedded=True, sizeoption=0, reverse=False, watermark=False,
                       effect=0, width=0, height=0):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> path = "C:/Users/Administrator/Desktop/KakaoTalk_20230709_023118549.jpg"
            >>> ctrl = hwp.insert_picture(path)  # 삽입한 이미지 객체를 리턴함.
            >>> pset = ctrl.Properties  # == hwp.create_set("ShapeObject")
            >>> pset.SetItem("TreatAsChar", False)  # 글자처럼취급 해제
            >>> pset.SetItem("TextWrap", 2)  # 그림을 글 뒤로
            >>> ctrl.Properties = pset  # 설정한 값 적용(간단!)
        """
        if sizeoption == 1 and not all([width, height]) and not self.is_cell():
            raise ValueError("sizeoption이 1일 때에는 width와 height를 지정해주셔야 합니다.\n"
                             "단, 셀 안에 있는 경우에는 셀 너비에 맞게 이미지 크기를 자동으로 조절합니다.")

        if path.startswith("http"):
            temp_path = tempfile.TemporaryFile().name
            request.urlretrieve(path, temp_path)
            path = temp_path
            # request.urlretrieve(path, os.path.join(os.getcwd(), "temp.jpg"))
        elif path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)

        try:
            ctrl = self.hwp.InsertPicture(Path=path, Embedded=embedded, sizeoption=sizeoption, Reverse=reverse,
                                          watermark=watermark, Effect=effect, Width=width, Height=height)
            pic_prop = ctrl.Properties
            if not all([width, height]) and self.is_cell():
                cell_param = self.hwp.HParameterSet.HShapeObject
                self.hwp.HAction.GetDefault("TablePropertyDialog", cell_param.HSet)
                cell_width = cell_param.ShapeTableCell.Width
                dst_height = pic_prop.Item("Height") / pic_prop.Item("Width") * cell_width
                pic_prop.SetItem("Width", cell_width)
                pic_prop.SetItem("Height", round(dst_height))
            else:
                sec_def = self.HParameterSet.HSecDef
                self.HAction.GetDefault("PageSetup", sec_def.HSet)
                page_width = (
                        sec_def.PageDef.PaperWidth - sec_def.PageDef.LeftMargin - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen)
                page_height = (
                        sec_def.PageDef.PaperHeight - sec_def.PageDef.TopMargin - sec_def.PageDef.BottomMargin - sec_def.PageDef.HeaderLen - sec_def.PageDef.FooterLen)
                pic_width = pic_prop.Item("Width")
                pic_height = pic_prop.Item("Height")
                if pic_width > page_width or pic_height > page_height:
                    width_shrink_ratio = page_width / pic_width
                    height_shrink_ratio = page_height / pic_height
                    if width_shrink_ratio <= height_shrink_ratio:
                        pic_prop.SetItem("Width", page_width)
                        pic_prop.SetItem("Height", pic_height * width_shrink_ratio)
                    else:
                        pic_prop.SetItem("Width", pic_width * height_shrink_ratio)
                        pic_prop.SetItem("Height", page_height)
            pic_prop.SetItem("TreatAsChar", treat_as_char)
            ctrl.Properties = pic_prop
            return ctrl
        finally:
            if os.path.basename(path).startswith("tmp"):
                os.remove(path)

    def insert_random_picture(self, x: int = 200, y: int = 200):
        return self.insert_picture(f"https://picsum.photos/{x}/{y}")

    def is_action_enable(self, action_id):
        return self.hwp.IsActionEnable(actionID=action_id)

    def IsActionEnable(self, action_id):
        return self.hwp.IsActionEnable(actionID=action_id)

    def is_command_lock(self, action_id):
        """
        해당 액션이 잠겨있는지 확인한다.

        :param action_id: 액션 ID. (ActionIDTable.Hwp 참조)

        :return:
            잠겨있으면 True, 잠겨있지 않으면 False를 반환한다.
        """
        return self.hwp.IsCommandLock(actionID=action_id)

    def IsCommandLock(self, action_id):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.KeyIndicator()[-1][1:].split(")")[0]
            "A1"
        """
        return self.hwp.KeyIndicator()

    def KeyIndicator(self) -> tuple:
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.KeyIndicator()[-1][1:].split(")")[0]
            "A1"
        """
        return self.hwp.KeyIndicator()

    def line_spacing_method(self, line_spacing):
        return self.hwp.LineSpacingMethod(LineSpacing=line_spacing)

    def LineSpacingMethod(self, line_spacing):
        return self.hwp.LineSpacingMethod(LineSpacing=line_spacing)

    def line_wrap_type(self, line_wrap):
        return self.hwp.LineWrapType(LineWrap=line_wrap)

    def LineWrapType(self, line_wrap):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.LockCommand("Undo", True)
            >>> hwp.LockCommand("Redo", True)
        """
        return self.hwp.LockCommand(ActID=act_id, isLock=is_lock)

    def LockCommand(self, act_id, is_lock):
        """
        특정 액션이 실행되지 않도록 잠근다.

        :param act_id: 액션 ID. (ActionIDTable.Hwp 참조)

        :param is_lock:
            True이면 액션의 실행을 잠그고, False이면 액션이 실행되도록 한다.

        :return: None

        :example:
            >>> # Undo와 Redo 잠그기
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.LockCommand("Undo", True)
            >>> hwp.LockCommand("Redo", True)
        """
        return self.hwp.LockCommand(ActID=act_id, isLock=is_lock)

    def lunar_to_solar(self, l_year, l_month, l_day, l_leap, s_year, s_month, s_day):
        return self.hwp.LunarToSolar(lYear=l_year, lMonth=l_month, lDay=l_day, lLeap=l_leap, sYear=s_year,
                                     sMonth=s_month, sDay=s_day)

    def LunarToSolar(self, l_year, l_month, l_day, l_leap, s_year, s_month, s_day):
        return self.hwp.LunarToSolar(lYear=l_year, lMonth=l_month, lDay=l_day, lLeap=l_leap, sYear=s_year,
                                     sMonth=s_month, sDay=s_day)

    def lunar_to_solar_by_set(self, l_year, l_month, l_day, l_leap):
        return self.hwp.LunarToSolarBySet(lYear=l_year, lMonth=l_month, lLeap=l_leap)

    def LunarToSolarBySet(self, l_year, l_month, l_day, l_leap):
        return self.hwp.LunarToSolarBySet(lYear=l_year, lMonth=l_month, lLeap=l_leap)

    def macro_state(self, macro_state):
        return self.hwp.MacroState(MacroState=macro_state)

    def MacroState(self, macro_state):
        return self.hwp.MacroState(MacroState=macro_state)

    def mail_type(self, mail_type):
        return self.hwp.MailType(MailType=mail_type)

    def MailType(self, mail_type):
        return self.hwp.MailType(MailType=mail_type)

    def metatag_exist(self, tag):
        return self.hwp.MetatagExist(tag=tag)

    def MetatagExist(self, tag):
        return self.hwp.MetatagExist(tag=tag)

    def mili_to_hwp_unit(self, mili):
        return self.hwp.MiliToHwpUnit(mili=mili)

    def MiliToHwpUnit(self, mili):
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

    def ModifyFieldProperties(self, field, remove, add):
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

    def ModifyMetatagProperties(self, tag, remove, add):
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

    def MovePos(self, move_id=1, para=0, pos=0):
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

    def move_to_field(self, field, idx=0, text=True, start=True, select=False):
        """
        지정한 필드로 캐럿을 이동한다.

        :param field:
            필드이름. GetFieldText()/PutFieldText()와 같은 형식으로
            이름 뒤에 ‘{{#}}’로 번호를 지정할 수 있다.

        :param idx:
            동일명으로 여러 개의 필드가 존재하는 경우,
            idx번째 필드로 이동하고자 할 때 사용한다. 기본값은 0.
            idx를 지정하지 않아도, 필드 파라미터 뒤에 ‘{{#}}’를 추가하여 인덱스를 지정할 수 있다.
            이 경우 기본적으로 f스트링을 사용하며, f스트링 내부에 탈출문자열 \가 적용되지 않으므로
            중괄호를 다섯 겹 입력해야 한다. 예 : hwp.move_to_field(f"필드명{{{{{i}}}}}")

        :param text:
            필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(True)
            누름틀 코드로 이동할지(False)를 지정한다.
            누름틀이 아닌 필드일 경우 무시된다. 생략하면 True가 지정된다.

        :param start:
            필드의 처음(True)으로 이동할지 끝(False)으로 이동할지 지정한다.
            select를 True로 지정하면 무시된다. (캐럿이 처음에 위치해 있게 된다.)
            생략하면 True가 지정된다.

        :param select:
            필드 내용을 블록으로 선택할지(True), 캐럿만 이동할지(False) 지정한다.
            생략하면 False가 지정된다.
        :return:
        """
        if "{{" not in field:
            return self.hwp.MoveToField(Field=f"{field}{{{{{idx}}}}}", Text=text, start=start, select=select)
        else:
            return self.hwp.MoveToField(Field=field, Text=text, start=start, select=select)

    def MoveToField(self, field, idx=0, text=True, start=True, select=False):
        """
        지정한 필드로 캐럿을 이동한다.

        :param field:
            필드이름. GetFieldText()/PutFieldText()와 같은 형식으로
            이름 뒤에 ‘{{#}}’로 번호를 지정할 수 있다.

        :param idx:
            동일명으로 여러 개의 필드가 존재하는 경우,
            idx번째 필드로 이동하고자 할 때 사용한다. 기본값은 0.
            idx를 지정하지 않아도, 필드 파라미터 뒤에 ‘{{#}}’를 추가하여 인덱스를 지정할 수 있다.
            이 경우 기본적으로 f스트링을 사용하며, f스트링 내부에 탈출문자열 \가 적용되지 않으므로
            중괄호를 다섯 겹 입력해야 한다. 예 : hwp.move_to_field(f"필드명{{{{{i}}}}}")

        :param text:
            필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(True)
            누름틀 코드로 이동할지(False)를 지정한다.
            누름틀이 아닌 필드일 경우 무시된다. 생략하면 True가 지정된다.

        :param start:
            필드의 처음(True)으로 이동할지 끝(False)으로 이동할지 지정한다.
            select를 True로 지정하면 무시된다. (캐럿이 처음에 위치해 있게 된다.)
            생략하면 True가 지정된다.

        :param select:
            필드 내용을 블록으로 선택할지(True), 캐럿만 이동할지(False) 지정한다.
            생략하면 False가 지정된다.
        :return:
        """
        if "{{" not in field:
            return self.hwp.MoveToField(Field=f"{field}{{{{{idx}}}}}", Text=text, start=start, select=select)
        else:
            return self.hwp.MoveToField(Field=field, Text=text, start=start, select=select)

    def move_to_metatag(self, tag, text, start, select):
        return self.hwp.MoveToMetatag(tag=tag, Text=text, start=start, select=select)

    def MoveToMetatag(self, tag, text, start, select):
        return self.hwp.MoveToMetatag(tag=tag, Text=text, start=start, select=select)

    def number_format(self, num_format):
        return self.hwp.NumberFormat(NumFormat=num_format)

    def NumberFormat(self, num_format):
        return self.hwp.NumberFormat(NumFormat=num_format)

    def numbering(self, numbering):
        return self.hwp.Numbering(Numbering=numbering)

    def Numbering(self, numbering):
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
        if filename.startswith("http"):
            try:
                # url 문자열 중 hwp 파일명이 포함되어 있는지 체크해서 해당 파일명을 사용.
                hwp_name = [parse.unquote_plus(i) for i in re.split("[/?=&]", filename) if ".hwp" in i][0]
            except IndexError as e:
                # url 문자열 안에 hwp 파일명이 포함되어 있지 않은 경우에는 임시파일명 지정(temp.hwp)
                hwp_name = "temp.hwp"
            request.urlretrieve(filename, os.path.join(os.getcwd(), hwp_name))
            filename = os.path.join(os.getcwd(), hwp_name)
        elif filename.lower()[1] != ":":
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.Open(filename=filename, Format=format, arg=arg)

    def Open(self, filename, format="", arg=""):
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
        if filename.startswith("http"):
            try:
                # url 문자열 중 hwp 파일명이 포함되어 있는지 체크해서 해당 파일명을 사용.
                hwp_name = [parse.unquote_plus(i) for i in re.split("[/?=&]", filename) if ".hwp" in i][0]
            except IndexError as e:
                # url 문자열 안에 hwp 파일명이 포함되어 있지 않은 경우에는 임시파일명 지정(temp.hwp)
                hwp_name = "temp.hwp"
            request.urlretrieve(filename, os.path.join(os.getcwd(), hwp_name))
            filename = os.path.join(os.getcwd(), hwp_name)
        elif filename.lower()[1] != ":":
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.Open(filename=filename, Format=format, arg=arg)

    def page_num_position(self, pagenumpos: Literal[
        "TopLeft", "TopCenter", "TopRight", "BottomLeft", "BottomCenter", "BottomRight", "InsideTop", "OutsideTop", "InsideBottom", "OutsideBottom", "None"] = "BottomCenter"):
        return self.hwp.PageNumPosition(pagenumpos=pagenumpos)

    def PageNumPosition(self, pagenumpos: Literal[
        "TopLeft", "TopCenter", "TopRight", "BottomLeft", "BottomCenter", "BottomRight", "InsideTop", "OutsideTop", "InsideBottom", "OutsideBottom", "None"] = "BottomCenter"):
        return self.hwp.PageNumPosition(pagenumpos=pagenumpos)

    def page_type(self, page_type):
        return self.hwp.PageType(PageType=page_type)

    def PageType(self, page_type):
        return self.hwp.PageType(PageType=page_type)

    def para_head_align(self, para_head_align):
        return self.hwp.ParaHeadAlign(ParaHeadAlign=para_head_align)

    def ParaHeadAlign(self, para_head_align):
        return self.hwp.ParaHeadAlign(ParaHeadAlign=para_head_align)

    def pic_effect(self, pic_effect):
        return self.hwp.PicEffect(PicEffect=pic_effect)

    def PicEffect(self, pic_effect):
        return self.hwp.PicEffect(PicEffect=pic_effect)

    def placement_type(self, restart):
        return self.hwp.PlacementType(Restart=restart)

    def PlacementType(self, restart):
        return self.hwp.PlacementType(Restart=restart)

    def point_to_hwp_unit(self, point):
        return self.hwp.PointToHwpUnit(Point=point)

    def PointToHwpUnit(self, point):
        return self.hwp.PointToHwpUnit(Point=point)

    def hwp_unit_to_point(self, HwpUnit: int):
        return HwpUnit * 100

    def HwpUnitToPoint(self, HwpUnit: int):
        return HwpUnit * 100

    def hwp_unit_to_inch(self, HwpUnit):
        return HwpUnit / 7200

    def HwpUnitToInch(self, HwpUnit):
        return HwpUnit / 7200

    def inch_to_hwp_unit(self, inch):
        return inch * 7200

    def InchToHwpUnit(self, inch):
        return inch * 7200

    def present_effect(self, prsnteffect):
        return self.hwp.PresentEffect(prsnteffect=prsnteffect)

    def PresentEffect(self, prsnteffect):
        return self.hwp.PresentEffect(prsnteffect=prsnteffect)

    def print_device(self, print_device):
        return self.hwp.PrintDevice(PrintDevice=print_device)

    def PrintDevice(self, print_device):
        return self.hwp.PrintDevice(PrintDevice=print_device)

    def print_paper(self, print_paper):
        return self.hwp.PrintPaper(PrintPaper=print_paper)

    def PrintPaper(self, print_paper):
        return self.hwp.PrintPaper(PrintPaper=print_paper)

    def print_range(self, print_range):
        return self.hwp.PrintRange(PrintRange=print_range)

    def PrintRange(self, print_range):
        return self.hwp.PrintRange(PrintRange=print_range)

    def print_type(self, print_method):
        return self.hwp.PrintType(PrintMethod=print_method)

    def PrintType(self, print_method):
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

    def ProtectPrivateInfo(self, protecting_char, private_pattern_type):
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

    def put_field_text(self, field, text: Union[str, list, tuple, pd.Series] = "", idx=None):
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
            **단, field에 dict를 입력하는 경우에는 text 파라미터를 무시하고
            dict.keys를 필드명으로, dict.values를 필드값으로 입력한다.**


        :param text:
            필드에 채워 넣을 문자열의 리스트.
            형식은 필드 리스트와 동일하게 필드의 개수만큼
            텍스트를 0x02로 구분하여 지정한다.

        :return: None

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 현재 캐럿 위치에 zxcv 필드 생성
            >>> hwp.create_field("zxcv")
            >>> # zxcv 필드에 "Hello world!" 텍스트 삽입
            >>> hwp.put_field_text("zxcv", "Hello world!")
        """
        if isinstance(field, str) and (field.endswith(".xlsx") or field.endswith(".xls")):
            field = pd.read_excel(field)

        if isinstance(field, dict):  # dict 자료형의 경우에는 text를 생략하고
            field, text = list(zip(*list(field.items())))
            field_str = ""
            text_str = ""
            if isinstance(idx, int):
                for f_i, f in enumerate(field):
                    field_str += f"{f}{{{{{idx}}}}}\x02"
                    text_str += f"{text[f_i][idx]}\x02"  # for t_i, t in enumerate(text[f_i]):
            else:
                if isinstance(text[0], (list, tuple)):
                    for f_i, f in enumerate(field):
                        for t_i, t in enumerate(text[f_i]):
                            field_str += f"{f}{{{{{t_i}}}}}\x02"
                            text_str += f"{t}\x02"
                elif isinstance(text[0], (str, int, float)):
                    for f_i, f in enumerate(field):
                        field_str += f"{f}\x02"
                    text_str = "\x02".join(text)

            self.hwp.PutFieldText(Field=field_str, Text=text_str)
            return

        if isinstance(field, str) and type(text) in (list, tuple, pd.Series):
            field = [f"{field}{{{{{i}}}}}" for i in range(len(text))]

        if type(field) in [pd.Series]:  # 필드명 리스트를 파라미터로 넣은 경우
            if not text:  # text 파라미터가 입력되지 않았다면
                text_str = "\x02".join([field[i] for i in field.index])
                field_str = "\x02".join([str(i) for i in field.index])  # \x02로 병합
                self.hwp.PutFieldText(Field=field_str, Text=text_str)
                return
            elif type(text) in [list, tuple, pd.Series]:  # 필드 텍스트를 리스트나 배열로 넣은 경우에도
                text = "\x02".join([str(i) for i in text])  # \x02로 병합
            else:
                raise IOError("text parameter required.")

        if type(field) in [list, tuple]:

            # field와 text가 [[field0:str, list[text:str]], [field1:str, list[text:str]]] 타입인 경우
            if not text and isinstance(field[0][0], (str, int, float)) and not isinstance(field[0][1],
                                                                                          (str, int)) and len(
                field[0][1]) >= 1:
                text_str = ""
                field_str = "\x02".join(
                    [str(field[i][0]) + f"{{{{{j}}}}}" for j in range(len(field[0][1])) for i in range(len(field))])
                for i in range(len(field[0][1])):
                    text_str += "\x02".join([str(field[j][1][i]) for j in range(len(field))]) + "\x02"
                return self.hwp.PutFieldText(Field=field_str, Text=text_str)

            elif type(field) in (list, tuple, set) and type(text) in (list, tuple, set):
                # field와 text가 모두 배열로 만들어져 있는 경우
                field_str = "\x02".join([str(field[i]) for i in range(len(field))])
                text_str = "\x02".join([str(text[i]) for i in range(len(text))])
                return self.hwp.PutFieldText(Field=field_str, Text=text_str)
            else:
                # field와 text가 field타입 안에 [[field0:str, text0:str], [field1:str, text1:str]] 형태로 들어간 경우
                field_str = "\x02".join([str(field[i][0]) for i in range(len(field))])
                text_str = "\x02".join([str(field[i][1]) for i in range(len(field))])
                return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(field, pd.DataFrame):
            if isinstance(field.columns, pd.core.indexes.range.RangeIndex):
                field = field.T
            text_str = ""
            if isinstance(idx, int):
                field_str = "\x02".join([str(i) + f"{{{{{idx}}}}}" for i in field])  # \x02로 병합
                text_str += "\x02".join([str(t) for t in field.iloc[idx]]) + "\x02"
            else:
                field_str = "\x02".join([str(i) + f"{{{{{j}}}}}" for j in range(len(field)) for i in field])  # \x02로 병합
                for i in range(len(field)):
                    text_str += "\x02".join([str(t) for t in field.iloc[i]]) + "\x02"
            return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(text, pd.DataFrame):
            if not isinstance(text.columns, pd.core.indexes.range.RangeIndex):
                text = text.T
            text_str = ""
            if isinstance(idx, int):
                field_str = "\x02".join([i + f"{{{{{idx}}}}}" for i in field.split("\x02")])  # \x02로 병합
                text_str += "\x02".join([str(t) for t in text[idx]]) + "\x02"
            else:
                field_str = "\x02".join([str(i) + f"{{{{{j}}}}}" for i in field.split("\x02") for j in
                                         range(len(text.columns))])  # \x02로 병합
                for i in range(len(text)):
                    text_str += "\x02".join([str(t) for t in text.iloc[i]]) + "\x02"
            return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(idx, int):
            return self.hwp.PutFieldText(Field=field.replace("\x02", f"{{{{{idx}}}}}\x02") + f"{{{{{idx}}}}}",
                                         Text=text)
        else:
            return self.hwp.PutFieldText(Field=field, Text=text)

    def PutFieldText(self, field, text: Union[str, list, tuple, pd.Series] = "", idx=None):
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
            **단, field에 dict를 입력하는 경우에는 text 파라미터를 무시하고
            dict.keys를 필드명으로, dict.values를 필드값으로 입력한다.**


        :param text:
            필드에 채워 넣을 문자열의 리스트.
            형식은 필드 리스트와 동일하게 필드의 개수만큼
            텍스트를 0x02로 구분하여 지정한다.

        :return: None

        :example:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 현재 캐럿 위치에 zxcv 필드 생성
            >>> hwp.create_field("zxcv")
            >>> # zxcv 필드에 "Hello world!" 텍스트 삽입
            >>> hwp.put_field_text("zxcv", "Hello world!")
        """
        if isinstance(field, str) and (field.endswith(".xlsx") or field.endswith(".xls")):
            field = pd.read_excel(field)

        if isinstance(field, dict):  # dict 자료형의 경우에는 text를 생략하고
            field, text = list(zip(*list(field.items())))
            field_str = ""
            text_str = ""
            if isinstance(idx, int):
                for f_i, f in enumerate(field):
                    field_str += f"{f}{{{{{idx}}}}}\x02"
                    text_str += f"{text[f_i][idx]}\x02"  # for t_i, t in enumerate(text[f_i]):
            else:
                if isinstance(text[0], (list, tuple)):
                    for f_i, f in enumerate(field):
                        for t_i, t in enumerate(text[f_i]):
                            field_str += f"{f}{{{{{t_i}}}}}\x02"
                            text_str += f"{t}\x02"
                elif isinstance(text[0], (str, int, float)):
                    for f_i, f in enumerate(field):
                        field_str += f"{f}\x02"
                    text_str = "\x02".join(text)

            self.hwp.PutFieldText(Field=field_str, Text=text_str)
            return

        if isinstance(field, str) and type(text) in (list, tuple, pd.Series):
            field = [f"{field}{{{{{i}}}}}" for i in range(len(text))]

        if type(field) in [pd.Series]:  # 필드명 리스트를 파라미터로 넣은 경우
            if not text:  # text 파라미터가 입력되지 않았다면
                text_str = "\x02".join([field[i] for i in field.index])
                field_str = "\x02".join([str(i) for i in field.index])  # \x02로 병합
                self.hwp.PutFieldText(Field=field_str, Text=text_str)
                return
            elif type(text) in [list, tuple, pd.Series]:  # 필드 텍스트를 리스트나 배열로 넣은 경우에도
                text = "\x02".join([str(i) for i in text])  # \x02로 병합
            else:
                raise IOError("text parameter required.")

        if type(field) in [list, tuple]:

            # field와 text가 [[field0:str, list[text:str]], [field1:str, list[text:str]]] 타입인 경우
            if not text and isinstance(field[0][0], (str, int, float)) and not isinstance(field[0][1],
                                                                                          (str, int)) and len(
                field[0][1]) >= 1:
                text_str = ""
                field_str = "\x02".join(
                    [str(field[i][0]) + f"{{{{{j}}}}}" for j in range(len(field[0][1])) for i in range(len(field))])
                for i in range(len(field[0][1])):
                    text_str += "\x02".join([str(field[j][1][i]) for j in range(len(field))]) + "\x02"
                return self.hwp.PutFieldText(Field=field_str, Text=text_str)

            elif type(field) in (list, tuple, set) and type(text) in (list, tuple, set):
                # field와 text가 모두 배열로 만들어져 있는 경우
                field_str = "\x02".join([str(field[i]) for i in range(len(field))])
                text_str = "\x02".join([str(text[i]) for i in range(len(text))])
                return self.hwp.PutFieldText(Field=field_str, Text=text_str)

            else:
                # field와 text가 field타입 안에 [[field0:str, text0:str], [field1:str, text1:str]] 형태로 들어간 경우
                field_str = "\x02".join([str(field[i][0]) for i in range(len(field))])
                text_str = "\x02".join([str(field[i][1]) for i in range(len(field))])
                return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(field, pd.DataFrame):
            if isinstance(field.columns, pd.core.indexes.range.RangeIndex):
                field = field.T
            text_str = ""
            if isinstance(idx, int):
                field_str = "\x02".join([str(i) + f"{{{{{idx}}}}}" for i in field])  # \x02로 병합
                text_str += "\x02".join([str(t) for t in field.iloc[idx]]) + "\x02"
            else:
                field_str = "\x02".join([str(i) + f"{{{{{j}}}}}" for j in range(len(field)) for i in field])  # \x02로 병합
                for i in range(len(field)):
                    text_str += "\x02".join([str(t) for t in field.iloc[i]]) + "\x02"
            return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(text, pd.DataFrame):
            if not isinstance(text.columns, pd.core.indexes.range.RangeIndex):
                text = text.T
            text_str = ""
            if isinstance(idx, int):
                field_str = "\x02".join([i + f"{{{{{idx}}}}}" for i in field.split("\x02")])  # \x02로 병합
                text_str += "\x02".join([str(t) for t in text[idx]]) + "\x02"
            else:
                field_str = "\x02".join([str(i) + f"{{{{{j}}}}}" for i in field.split("\x02") for j in
                                         range(len(text.columns))])  # \x02로 병합
                for i in range(len(text)):
                    text_str += "\x02".join([str(t) for t in text.iloc[i]]) + "\x02"
            return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(idx, int):
            return self.hwp.PutFieldText(Field=field.replace("\x02", f"{{{{{idx}}}}}\x02") + f"{{{{{idx}}}}}",
                                         Text=text)
        else:
            return self.hwp.PutFieldText(Field=field, Text=text)

    def put_metatag_name_text(self, tag, text):
        return self.hwp.PutMetatagNameText(tag=tag, Text=text)

    def PutMetatagNameText(self, tag, text):
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

    def Quit(self):
        """
        한/글을 종료한다.
        단, 저장되지 않은 변경사항이 있는 경우 팝업이 뜨므로
        clear나 save 등의 메서드를 실행한 후에 quit을 실행해야 한다.
        :return:
        """
        self.hwp.Quit()
        del self.hwp

    def rgb_color(self, red_or_colorname: str | tuple, green=255, blue=255):
        color_palette = {"Red": (255, 0, 0), "Green": (0, 255, 0), "Blue": (0, 0, 255), "Yellow": (255, 255, 0),
                         "Cyan": (0, 255, 255), "Magenta": (255, 0, 255), "Black": (0, 0, 0), "White": (255, 255, 255),
                         "Gray": (128, 128, 128), "Orange": (255, 165, 0), "DarkBlue": (0, 0, 139),
                         "Purple": (128, 0, 128),
                         "Pink": (255, 192, 203), "Lime": (0, 255, 0), "SkyBlue": (135, 206, 235),
                         "Gold": (255, 215, 0),
                         "Silver": (192, 192, 192), "Mint": (189, 252, 201), "Tomato": (255, 99, 71),
                         "Olive": (128, 128, 0),
                         "Crimson": (220, 20, 60), "Navy": (0, 0, 128), "Teal": (0, 128, 128),
                         "Chocolate": (210, 105, 30), }
        if red_or_colorname in color_palette:
            return self.hwp.RGBColor(*color_palette[red_or_colorname])
        return self.hwp.RGBColor(red=red_or_colorname, green=green, blue=blue)

    def RGBColor(self, red_or_colorname: str | tuple, green=255, blue=255):
        color_palette = {"Red": (255, 0, 0), "Green": (0, 255, 0), "Blue": (0, 0, 255), "Yellow": (255, 255, 0),
                         "Cyan": (0, 255, 255), "Magenta": (255, 0, 255), "Black": (0, 0, 0), "White": (255, 255, 255),
                         "Gray": (128, 128, 128), "Orange": (255, 165, 0), "DarkBlue": (0, 0, 139),
                         "Purple": (128, 0, 128),
                         "Pink": (255, 192, 203), "Lime": (0, 255, 0), "SkyBlue": (135, 206, 235),
                         "Gold": (255, 215, 0),
                         "Silver": (192, 192, 192), "Mint": (189, 252, 201), "Tomato": (255, 99, 71),
                         "Olive": (128, 128, 0),
                         "Crimson": (220, 20, 60), "Navy": (0, 0, 128), "Teal": (0, 128, 128),
                         "Chocolate": (210, 105, 30), }
        if red_or_colorname in color_palette:
            return self.hwp.RGBColor(*color_palette[red_or_colorname])
        return self.hwp.RGBColor(red=red_or_colorname, green=green, blue=blue)

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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 사전에 레지스트리에 보안모듈이 등록되어 있어야 한다.
            >>> # 보다 자세한 설명은 공식문서 참조
            >>> hwp.register_module("FilePathChekDLL", "FilePathCheckerModule")
            True
        """
        self.register_regedit()
        return self.hwp.RegisterModule(ModuleType=module_type, ModuleData=module_data)

    def RegisterModule(self, module_type="FilePathCheckDLL", module_data="FilePathCheckerModule"):
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
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 사전에 레지스트리에 보안모듈이 등록되어 있어야 한다.
            >>> # 보다 자세한 설명은 공식문서 참조
            >>> hwp.register_module("FilePathChekDLL", "FilePathCheckerModule")
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
                        subprocess.check_output(['pip', 'show', 'pyhwpx'], stderr=subprocess.DEVNULL).decode(
                            encoding="cp949").split("\r\n") if i.startswith("Location: ")][0]
        except:
            try:
                location = [i.split(": ")[1] for i in subprocess.check_output(['pip', 'show', 'pyhwpx'],
                                                                              stderr=subprocess.DEVNULL).decode().split(
                    "\r\n") if i.startswith("Location: ")][0]
            except subprocess.CalledProcessError as e:
                # FilePathCheckerModule.dll을 못 찾는 경우에는 아래 분기 중 하나를 실행
                #

                # 1. pyinstaller로 컴파일했고,
                #    --add-binary="FilePathCheckerModule.dll:." 옵션을 추가한 경우
                location = ""
                for dirpath, dirnames, filenames in os.walk(pyinstaller_path):
                    for filename in filenames:
                        if filename == "FilePathCheckerModule.dll":
                            location = dirpath

                # 2. "FilePathCheckerModule.dll" 파일을 실행파일과 같은 경로에 둔 경우
                if "FilePathCheckerModule.dll" in os.listdir(os.getcwd()):
                    location = os.getcwd()
                elif os.path.exists(os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModule.dll")):
                    location = os.environ["USERPROFILE"]

                # 3. 위의 두 경우가 아닐 때, 인터넷에 연결되어 있는 경우에는
                #    사용자 폴더(예: c:\\users\\user)에
                #    FilePathCheckerModule.dll을 다운로드하기.
                if not location:
                    print("not location")
                    # pyhwpx가 설치되어 있지 않은 PC에서는,
                    # 공식사이트에서 다운을 받게 하자.
                    from zipfile import ZipFile
                    print("downloading FilePathCheckerModule.dll to User Profile Folder")
                    f = request.urlretrieve(
                        "https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/%EB%B3%B4%EC%95%88%EB%AA%A8%EB%93%88(Automation).zip",
                        filename=os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModule.zip"))
                    with ZipFile(f[0]) as zf:
                        zf.extract(
                            "FilePathCheckerModuleExample.dll",
                            os.path.join(os.environ["USERPROFILE"]))
                    os.remove(os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModule.zip"))
                    if not os.path.exists(os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModule.dll")):
                        os.rename(os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModuleExample.dll"),
                                  os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModule.dll"))
                    location = os.environ["USERPROFILE"]
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
            >>> from pyhwpx import Hwp()
            >>> hwp = Hwp()
            >>>
            >>> hwp.register_private_info_pattern(0x01, "NNNN-NNNN;NN-NN-NNNN-NNNN")  # 전화번호패턴
        """
        return self.hwp.RegisterPrivateInfoPattern(PrivateType=private_type, PrivatePattern=private_pattern)

    def RegisterPrivateInfoPattern(self, private_type, private_pattern):
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
            >>> from pyhwpx import Hwp()
            >>> hwp = Hwp()
            >>>
            >>> hwp.register_private_info_pattern(0x01, "NNNN-NNNN;NN-NN-NNNN-NNNN")  # 전화번호패턴
        """
        return self.hwp.RegisterPrivateInfoPattern(PrivateType=private_type, PrivatePattern=private_pattern)

    def release_action(self, action):
        return self.hwp.ReleaseAction(action=action)

    def ReleaseAction(self, action):
        return self.hwp.ReleaseAction(action=action)

    def release_scan(self):
        """
        InitScan()으로 설정된 초기화 정보를 해제한다.
        텍스트 검색작업이 끝나면 반드시 호출하여 설정된 정보를 해제해야 한다.

        :return: None
        """
        return self.hwp.ReleaseScan()

    def ReleaseScan(self):
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
            >>> hwp.create_field("asdf")  # "asdf" 필드 생성
            >>> hwp.rename_field("asdf", "zxcv")  # asdf 필드명을 "zxcv"로 변경
            >>> hwp.put_field_text("zxcv", "Hello world!")  # zxcv 필드에 텍스트 삽입
        """
        return self.hwp.RenameField(oldname=oldname, newname=newname)

    def rename_metatag(self, oldtag, newtag):
        return self.hwp.RenameMetatag(oldtag=oldtag, newtag=newtag)

    def RenameMetatag(self, oldtag, newtag):
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
        >>> hwp.replace_action("Cut", "Cut")

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

    def ReplaceAction(self, old_action_id, new_action_id):
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
        >>> hwp.replace_action("Cut", "Cut")

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

    def ReplaceFont(self, langid, des_font_name, des_font_type, new_font_name, new_font_type):
        return self.hwp.ReplaceFont(langid=langid, desFontName=des_font_name, desFontType=des_font_type,
                                    newFontName=new_font_name, newFontType=new_font_type)

    def revision(self, revision):
        return self.hwp.Revision(Revision=revision)

    def Revision(self, revision):
        return self.hwp.Revision(Revision=revision)

    # Run 액션

    def Run(self, act_id):
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

    def Cut(self, remove_cell=True):
        """
        잘라내기. Copy 액션과 유사하지만, 복사 대신 잘라내기 기능을 수행한다. 자주 쓰이는 메서드이다.
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
        한 글자 뒤로 이동. 현재 리스트만을 대상으로 동작한다.
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
        다음 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextParaBegin")
        if self.get_pos()[1] != cwd[1]:
            return True
        else:
            return False

    def MoveNextPos(self):
        """
        한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveNextPosEx(self):
        """
        한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함)
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextPosEx")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveNextWord(self):
        """
        한 단어 뒤로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveNextWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePageBegin(self):
        """
        현재 페이지의 시작점으로 이동한다.. 만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 상단으로 페이지뷰잉이 맞춰진다.
        """
        return self.hwp.HAction.Run("MovePageBegin")

    def MovePageDown(self):
        """
        앞 페이지의 시작으로 이동. 현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePageDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePageEnd(self):
        """
        현재 페이지의 끝점으로 이동한다.. 만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 하단으로 페이지뷰잉이 맞춰진다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePageEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePageUp(self):
        """
        뒤 페이지의 시작으로 이동. 현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다.
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
        한 레벨 상위/탑레벨/루트 리스트로 이동한다.. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다.
        """
        return self.hwp.HAction.Run("MoveParentList")

    def MovePrevChar(self):
        """
        한 글자 앞 이동. 현재 리스트만을 대상으로 동작한다.
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
        앞 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevParaBegin")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevParaEnd(self):
        """
        앞 문단의 끝으로 이동. 현재 리스트만을 대상으로 동작한다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevParaEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevPos(self):
        """
        한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevPosEx(self):
        """
        한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함)
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MovePrevPosEx")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MovePrevWord(self):
        """
        한 단어 앞으로 이동. 현재 리스트만을 대상으로 동작한다.
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
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSectionDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSectionUp(self):
        """
        앞 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다.
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSectionUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

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
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelLeft(self):
        """
        셀렉션: 캐럿을 (논리적 방향) 왼쪽으로 이동
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelLeft")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelLineBegin(self):
        """
        셀렉션: 줄 처음
        """
        return self.hwp.HAction.Run("MoveSelLineBegin")

    def MoveSelLineDown(self):
        """
        셀렉션: 한줄 아래
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelLineDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelLineEnd(self):
        """
        셀렉션: 줄 끝
        """
        return self.hwp.HAction.Run("MoveSelLineEnd")

    def MoveSelLineUp(self):
        """
        셀렉션: 한줄 위
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelLineUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

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
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextChar")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelNextParaBegin(self):
        """
        셀렉션: 다음 문단 처음
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextParaBegin")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelNextPos(self):
        """
        셀렉션: 다음 위치
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelNextWord(self):
        """
        셀렉션: 다음 단어
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelNextWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPageDown(self):
        """
        셀렉션: 페이지다운
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPageDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPageUp(self):
        """
        셀렉션: 페이지 업
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPageUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

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
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevChar")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevParaBegin(self):
        """
        셀렉션: 이전 문단 시작
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevParaBegin")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevParaEnd(self):
        """
        셀렉션: 이전 문단 끝
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevParaEnd")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevPos(self):
        """
        셀렉션: 이전 위치
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevPos")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelPrevWord(self):
        """
        셀렉션: 이전 단어
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelPrevWord")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelRight(self):
        """
        셀렉션: 캐럿을 (논리적 방향) 오른쪽으로 이동
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelRight")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

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
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelViewDown(self):
        """
        셀렉션: 아래
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelViewDown")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

    def MoveSelViewUp(self):
        """
        셀렉션: 위
        """
        cwd = self.get_pos()
        self.hwp.HAction.Run("MoveSelViewUp")
        if self.get_pos()[0] != cwd[0] or self.get_pos()[1:] != cwd[1:]:
            return True
        else:
            return False

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

    def paste(self, option: Literal[0, 1, 2, 3, 4, 5, 6] = 4):
        """
        붙여넣기 확장메서드. (참고로 paste가 아닌 Paste는 API 그대로 작동한다.)
        option 파라미터에 할당할 수 있는 값은 모두 7가지로,
        0: (셀) 왼쪽에 끼워넣기
        1: 오른쪽에 끼워넣기
        2: 위쪽에 끼워넣기
        3: 아래쪽에 끼워넣기
        4: 덮어쓰기
        5: 내용만 덮어쓰기
        6: 셀 안에 표로 넣기
        """
        pset = self.hwp.HParameterSet.HSelectionOpt
        self.hwp.HAction.GetDefault("Paste", pset.HSet)
        pset.option = option
        self.hwp.HAction.Execute("Paste", pset.HSet)

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
        # return self.hwp.HAction.Run("Select")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("Select", pset.HSet)
        return self.HAction.Execute("Select", pset.HSet)

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
        # return self.hwp.HAction.Run("ShapeObjTableSelCell")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("ShapeObjTableSelCell", pset.HSet)
        return self.HAction.Execute("ShapeObjTableSelCell", pset.HSet)

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
        # return self.hwp.HAction.Run("TableCellBlock")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("TableCellBlock", pset.HSet)
        return self.HAction.Execute("TableCellBlock", pset.HSet)

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

    def TableDeleteCell(self, remain_cell=False):
        """
        셀 삭제
        """
        if remain_cell:
            self.set_message_box_mode(0x1000)
        else:
            self.set_message_box_mode(0x2000)
        try:
            return self.hwp.HAction.Run("TableDeleteCell")
        finally:
            self.set_message_box_mode(0xF000)

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
        self.Cancel()
        result = self.hwp.HAction.Run("TableMergeTable")
        if result:
            return result
        else:
            self.set_message_box_mode(0x1)
            sleep(0.1)
            self.set_message_box_mode(0xf)
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
        # return self.hwp.HAction.Run("TableRightCell")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("TableRightCell", pset.HSet)
        return self.HAction.Execute("TableRightCell", pset.HSet)

    def TableRightCellAppend(self):
        """
        셀 이동: 셀 오른쪽에 이어서
        """
        # return self.hwp.HAction.Run("TableRightCellAppend")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("TableRightCellAppend", pset.HSet)
        return self.HAction.Execute("TableRightCellAppend", pset.HSet)

    def TableSplitCell(self, Rows=2, Cols=0, DistributeHeight=0, Merge=0):
        """
        셀 나누기. Run메서드 같아 보이지만,
        엄연히 파라미터셋이 필수인 정통액션이다.

        Rows: 나눌 행 수(기본값:2)
        Cols: 나눌 열 수(기본값:0)
        DistributeHeight: 줄 높이를 같게 나누기(0 or 1)
        Merge: 셀을 합친 후 나누기(0 or 1)
        """
        pset = self.HParameterSet.HTableSplitCell
        pset.Rows = Rows
        pset.Cols = Cols
        pset.DistributeHeight = DistributeHeight
        pset.Merge = Merge
        return self.HAction.Execute("TableSplitCell", pset.HSet)

    def TableSplitTable(self):
        """
        표 나누기
        """
        if self.get_cell_addr("tuple")[0] == 0:
            return False
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
            >>> hwp.run_script_macro("OnDocument_New", u_macro_type=1)
            True
            >>> hwp.run_script_macro("OnScriptMacro_중국어1성")
            True
        """
        return self.hwp.RunScriptMacro(FunctionName=function_name, uMacroType=u_macro_type, uScriptType=u_script_type)

    def RunScriptMacro(self, function_name, u_macro_type=0, u_script_type=0):
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
            >>> hwp.run_script_macro("OnDocument_New", u_macro_type=1)
            True
            >>> hwp.run_script_macro("OnScriptMacro_중국어1성")
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

    def Save(self, save_if_dirty=True):
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
        if path.lower()[1] != ":":
            path = os.path.abspath(path)
        ext = path.rsplit(".", maxsplit=1)[-1]
        if ext.lower() == "pdf":
            pset = self.HParameterSet.HFileOpenSave
            self.HAction.GetDefault("FileSaveAsPdf", pset.HSet)
            self.HParameterSet.HFileOpenSave.filename = path
            self.HParameterSet.HFileOpenSave.Format = "PDF"
            self.HParameterSet.HFileOpenSave.Attributes = 16384
            return self.HAction.Execute("FileSaveAsPdf", pset.HSet)
        else:
            return self.hwp.SaveAs(Path=path, Format=format, arg=arg)

    def SaveAs(self, path, format="HWP", arg=""):
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
        if path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)
        return self.hwp.SaveAs(Path=path, Format=format, arg=arg)

    def scan_font(self):
        return self.hwp.ScanFont()

    def ScanFont(self):
        return self.hwp.ScanFont()

    def select_text_by_get_pos(self, s_getpos, e_getpos):
        self.set_pos(s_getpos[0], 0, 0)
        return self.hwp.SelectText(spara=s_getpos[1], spos=s_getpos[2], epara=e_getpos[1], epos=e_getpos[2])

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
            _, slist, spara, spos, elist, epara, epos = spara
        self.set_pos(slist, 0, 0)
        if epos == -1:
            self.hwp.SelectText(spara=spara, spos=spos, epara=epara, epos=0)
            return self.MoveSelParaEnd()
        else:
            return self.hwp.SelectText(spara=spara, spos=spos, epara=epara, epos=epos)

    def SelectText(self, spara: Union[int, list, tuple] = 0, spos=0, epara=0, epos=0, slist=0):
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
            _, slist, spara, spos, elist, epara, epos = spara
        self.set_pos(slist, 0, 0)
        if epos == -1:
            self.hwp.SelectText(spara=spara, spos=spos, epara=epara, epos=0)
            return self.MoveSelParaEnd()
        else:
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
        if lp_image_path.lower()[1] != ":":
            lp_image_path = os.path.join(os.getcwd(), lp_image_path)
        return self.hwp.SetBarCodeImage(lpImagePath=lp_image_path, pgno=pgno, index=index, X=x, Y=y, Width=width,
                                        Height=height)

    def SetBarCodeImage(self, lp_image_path, pgno, index, x, y, width, height):
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
        if lp_image_path.lower()[1] != ":":
            lp_image_path = os.path.join(os.getcwd(), lp_image_path)
        return self.hwp.SetBarCodeImage(lpImagePath=lp_image_path, pgno=pgno, index=index, X=x, Y=y, Width=width,
                                        Height=height)

    def set_cur_field_name(self, field, option=0, direction="", memo=""):
        """
        현재 캐럿이 위치하는 곳의 필드이름을 설정한다.
        GetFieldList()의 옵션 중에 4(hwpFieldSelection) 옵션은 사용하지 않는다.
        (표의 셀에 셀필드를 매기고 싶은 경우 사용한다.)

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
        if not self.is_cell():
            raise AssertionError("캐럿이 표 안에 있지 않습니다.")
        if self.SelectionMode == 0x13:
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 0)
            pset.ShapeTableCell.CellCtrlData.name = field
            return self.HAction.Execute("TablePropertyDialog", pset.HSet)
        else:
            return self.hwp.SetCurFieldName(Field=field, option=option, Direction=direction, memo=memo)

    def SetCurFieldName(self, field, option=0, direction="", memo=""):
        """
        현재 캐럿이 위치하는 곳의 필드이름을 설정한다.
        GetFieldList()의 옵션 중에 4(hwpFieldSelection) 옵션은 사용하지 않는다.
        (표의 셀에 셀필드를 매기고 싶은 경우 사용한다.)

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
        if not self.is_cell():
            raise AssertionError("캐럿이 표 안에 있지 않습니다.")
        if self.SelectionMode == 0x13:
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 0)
            pset.ShapeTableCell.CellCtrlData.name = field
            return self.HAction.Execute("TablePropertyDialog", pset.HSet)
        else:
            return self.hwp.SetCurFieldName(Field=field, option=option, Direction=direction, memo=memo)

    def set_cur_metatag_name(self, tag):
        return self.hwp.SetCurMetatagName(tag=tag)

    def SetCurMetatagName(self, tag):
        return self.hwp.SetCurMetatagName(tag=tag)

    def set_drm_authority(self, authority):
        return self.hwp.SetDRMAuthority(authority=authority)

    def SetDRMAuthority(self, authority):
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

    def SetFieldViewOption(self, option):
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

    def SetMessageBoxMode(self, mode):
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
        self.hwp.SetPos(List=list, Para=para, pos=pos)
        if (list, para) == self.get_pos()[:2]:
            return True
        else:
            return False

    def SetPos(self, list, para, pos):
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
        self.hwp.SetPos(List=list, Para=para, pos=pos)
        if para == self.get_pos()[1]:
            return True
        else:
            return False

    def set_pos_by_set(self, disp_val):
        """
        캐럿을 ParameterSet으로 얻어지는 위치로 옮긴다.

        :param disp_val:
            캐럿을 옮길 위치에 대한 ParameterSet 정보

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> start_pos = hwp.GetPosBySet()  # 현재 위치를 저장하고,
            >>> hwp.set_pos_by_set(start_pos)  # 특정 작업 후에 저장위치로 재이동
        """
        return self.hwp.SetPosBySet(dispVal=disp_val)

    def SetPosBySet(self, disp_val):
        """
        캐럿을 ParameterSet으로 얻어지는 위치로 옮긴다.

        :param disp_val:
            캐럿을 옮길 위치에 대한 ParameterSet 정보

        :return:
            성공하면 True, 실패하면 False

        :example:
            >>> start_pos = hwp.GetPosBySet()  # 현재 위치를 저장하고,
            >>> hwp.set_pos_by_set(start_pos)  # 특정 작업 후에 저장위치로 재이동
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

    def SetPrivateInfoPassword(self, password):
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

    def set_text_file(self, data: str, format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X",
                      option="insertfile"):
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

    def SetTextFile(self, data: str, format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X",
                    option="insertfile"):
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

    def SetTitleName(self, title):
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

    def SetUserInfo(self, user_info_id, value):
        return self.hwp.SetUserInfo(userInfoId=user_info_id, Value=value)

    def side_type(self, side_type):
        return self.hwp.SideType(SideType=side_type)

    def SideType(self, side_type):
        return self.hwp.SideType(SideType=side_type)

    def signature(self, signature):
        return self.hwp.Signature(Signature=signature)

    def Signature(self, signature):
        return self.hwp.Signature(Signature=signature)

    def slash(self, slash):
        return self.hwp.Slash(Slash=slash)

    def Slash(self, slash):
        return self.hwp.Slash(Slash=slash)

    def solar_to_lunar(self, s_year, s_month, s_day, l_year, l_month, l_day, l_leap):
        return self.hwp.SolarToLunar(sYear=s_year, sMonth=s_month, sDay=s_day, lYear=l_year, lMonth=l_month, lDay=l_day,
                                     lLeap=l_leap)

    def SolarToLunar(self, s_year, s_month, s_day, l_year, l_month, l_day, l_leap):
        return self.hwp.SolarToLunar(sYear=s_year, sMonth=s_month, sDay=s_day, lYear=l_year, lMonth=l_month, lDay=l_day,
                                     lLeap=l_leap)

    def solar_to_lunar_by_set(self, s_year, s_month, s_day):
        return self.hwp.SolarToLunarBySet(sYear=s_year, sMonth=s_month, sDay=s_day)

    def SolarToLunarBySet(self, s_year, s_month, s_day):
        return self.hwp.SolarToLunarBySet(sYear=s_year, sMonth=s_month, sDay=s_day)

    def sort_delimiter(self, sort_delimiter):
        return self.hwp.SortDelimiter(SortDelimiter=sort_delimiter)

    def SortDelimiter(self, sort_delimiter):
        return self.hwp.SortDelimiter(SortDelimiter=sort_delimiter)

    def strike_out(self, strike_out_type):
        return self.hwp.StrikeOut(StrikeOutType=strike_out_type)

    def StrikeOut(self, strike_out_type):
        return self.hwp.StrikeOut(StrikeOutType=strike_out_type)

    def style_type(self, style_type):
        return self.hwp.StyleType(StyleType=style_type)

    def StyleType(self, style_type):
        return self.hwp.StyleType(StyleType=style_type)

    def subt_pos(self, subt_pos):
        return self.hwp.SubtPos(SubtPos=subt_pos)

    def SubtPos(self, subt_pos):
        return self.hwp.SubtPos(SubtPos=subt_pos)

    def table_break(self, page_break):
        return self.hwp.TableBreak(PageBreak=page_break)

    def TableBreak(self, page_break):
        return self.hwp.TableBreak(PageBreak=page_break)

    def table_format(self, table_format):
        return self.hwp.TableFormat(TableFormat=table_format)

    def TableFormat(self, table_format):
        return self.hwp.TableFormat(TableFormat=table_format)

    def table_swap_type(self, tableswap):
        return self.hwp.TableSwapType(tableswap=tableswap)

    def TableSwapType(self, tableswap):
        return self.hwp.TableSwapType(tableswap=tableswap)

    def table_target(self, table_target):
        return self.hwp.TableTarget(TableTarget=table_target)

    def TableTarget(self, table_target):
        return self.hwp.TableTarget(TableTarget=table_target)

    def text_align(self, text_align):
        return self.hwp.TextAlign(TextAlign=text_align)

    def TextAlign(self, text_align):
        return self.hwp.TextAlign(TextAlign=text_align)

    def text_art_align(self, text_art_align):
        return self.hwp.TextArtAlign(TextArtAlign=text_art_align)

    def TextArtAlign(self, text_art_align):
        return self.hwp.TextArtAlign(TextArtAlign=text_art_align)

    def text_dir(self, text_direction):
        return self.hwp.TextDir(TextDirection=text_direction)

    def TextDir(self, text_direction):
        return self.hwp.TextDir(TextDirection=text_direction)

    def text_flow_type(self, text_flow):
        return self.hwp.TextFlowType(TextFlow=text_flow)

    def TextFlowType(self, text_flow):
        return self.hwp.TextFlowType(TextFlow=text_flow)

    def text_wrap_type(self, text_wrap):
        return self.hwp.TextWrapType(TextWrap=text_wrap)

    def TextWrapType(self, text_wrap):
        return self.hwp.TextWrapType(TextWrap=text_wrap)

    def un_select_ctrl(self):
        return self.hwp.UnSelectCtrl()

    def UnSelectCtrl(self):
        return self.hwp.UnSelectCtrl()

    def v_align(self, v_align):
        return self.hwp.VAlign(VAlign=v_align)

    def VAlign(self, v_align):
        return self.hwp.VAlign(VAlign=v_align)

    def vert_rel(self, vert_rel):
        return self.hwp.VertRel(VertRel=vert_rel)

    def VertRel(self, vert_rel):
        return self.hwp.VertRel(VertRel=vert_rel)

    def view_flag(self, view_flag):
        return self.hwp.ViewFlag(ViewFlag=view_flag)

    def ViewFlag(self, view_flag):
        return self.hwp.ViewFlag(ViewFlag=view_flag)

    def watermark_brush(self, watermark_brush):
        return self.hwp.WatermarkBrush(WatermarkBrush=watermark_brush)

    def WatermarkBrush(self, watermark_brush):
        return self.hwp.WatermarkBrush(WatermarkBrush=watermark_brush)

    def width_rel(self, width_rel):
        return self.hwp.WidthRel(WidthRel=width_rel)

    def WidthRel(self, width_rel):
        return self.hwp.WidthRel(WidthRel=width_rel)
