from importlib.resources import files
import ctypes
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import urllib.error
import xml.etree.ElementTree as ET
import zipfile
from functools import wraps

from collections import defaultdict
from io import StringIO
from time import sleep
from typing import Literal, Union, Any
from urllib import request, parse
from winreg import QueryValueEx

import numpy as np
import pandas as pd
import pyperclip as cb
from PIL import Image

if sys.platform == 'win32':
    import pythoncom
    import win32api
    import win32con
    import win32gui

    # CircularImport 오류 출력안함
    devnull = open(os.devnull, 'w')
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    sys.stdout = devnull
    sys.stderr = devnull

    try:
        import win32com.client as win32
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        devnull.close()

# for pyinstaller
_ = files("pyhwpx").joinpath("FilePathCheckerModule.dll")

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
def com_initialized(func):
    """
    이용준님께서 기여해주셨습니다. (https://github.com/YongJun-Lee-98)
    이 데코레이터는 함수 실행 전에 COM 라이브러리를 초기화하고,
    실행 후에 COM 라이브러리를 해제합니다.

    Python의 GC가 COM 객체를 자동으로 제거하더라도,
    CoUninitialize()를 호출하지 않으면
    COM 라이브러리가 해당 스레드에서 완전히 해제되지 않을 수 있기 때문입니다.
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()

    return wrapper


def addr_to_tuple(cell_address: str) -> tuple[int, int]:
    """
    엑셀주소를 튜플로 변환하는 헬퍼함수

    엑셀 셀 주소("A1", "B2", "ASD100000" 등)를 `(row, col)` 튜플로 변환하는 헬퍼함수입니다.
    예를 들어 `addr_to_tuple("C3")`을 실행하면 `(3, 3)`을 리턴하는 식입니다.
    `pyhwpx` 일부 메서드의 내부 연산에 사용됩니다.

    Args:
        cell_address: 엑셀 방식의 "셀주소" 문자열

    Returns:
        (row, column) 형식의 주소 튜플

    Examples:
        >>> from pyhwpx import addr_to_tuple
        >>> print(addr_to_tuple("C3"))
        (3, 3)
        >>> print(addr_to_tuple("AB10"))
        (10, 28)
        >>> print(addr_to_tuple("AAA100000"))
        (100000, 703)
    """

    # 정규표현식을 이용해 문자 부분(열), 숫자 부분(행)을 분리
    match = re.match(r"^([A-Z]+)(\d+)$", cell_address.upper())
    if not match:
        raise ValueError(f"잘못된 셀 주소 형식입니다: {cell_address}")

    col_letters, row_str = match.groups()

    # 문자 부분 -> 열 번호(col)로 변환
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord('A') + 1)

    # 숫자 부분 -> 행 번호(row)로 변환
    row = int(row_str)

    return row, col


def tuple_to_addr(col: int, row: int) -> str:
    """
    (컬럼번호, 행번호)를 인자로 받아 엑셀 셀 주소 문자열(예: `"AAA3"`)을 반환합니다.

    `hwp.goto_addr(addr)` 메서드 내부에서 활용됩니다. 직접 사용하지 않습니다.

    Args:
        col: 열(칼럼) 번호(1부터 시작)
        row: 행(로우) 번호(1부터 시작)

    Returns:
        str: 엑셀 형식의 주소 문자열(예: `"A1"`, `"VVS1004"`)

    Examples:
        >>> from pyhwpx import tuple_to_addr
        >>> print(tuple_to_addr(1, 2))
        B1
    """
    letters = []
    # 컬럼번호(col)를 "A"~"Z", "AA"~"ZZ", ... 형태로 변환
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters.append(chr(remainder + ord('A')))
    letters.reverse()  # 스택처럼 뒤집어 넣었으므로 최종 결과는 reverse() 후 합침
    col_str = "".join(letters)

    return f"{col_str}{row}"


def _open_dialog(hwnd, key="M", delay=0.2) -> None:
    win32gui.SetForegroundWindow(hwnd)
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
    win32api.keybd_event(ord(key), 0, 0, 0)
    win32api.keybd_event(ord(key), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(delay)


def _get_edit_text(hwnd: int, delay: int = 0.2) -> str:
    """
    수식컨트롤 관련 헬퍼함수. 직접 사용하지 말 것.
    """
    sleep(delay)
    length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH) + 1
    buffer = win32gui.PyMakeBuffer(length * 2)
    win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length, buffer)
    text = buffer[:length * 2].tobytes().decode('utf-16')[:-1]
    return text


def _refresh_eq(hwnd: int, delay: int = 0.1) -> None:
    """
    수식 새로고침을 위한 키 전송 함수.
    hwnd로 수식 편집기 창을 찾아놓은 후 실행하면 Ctrl-(Tab-Tab)을 전송한다.
    EquationCreate 및 EquationModify에서
    수식을 정리하기 위해 만들어놓은 헬퍼함수.
    (이런 기능은 넣는 게 아니었어ㅜㅜㅜ)

    """
    sleep(delay)
    win32gui.SetForegroundWindow(hwnd)
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, 2, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, 2, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, 2, 0)
    sleep(delay)


def _eq_create(visible: bool) -> bool:
    """
    멀티스레드 형태로 새 수식편집기를 실행하는 헬퍼함수. 직접 사용하지 말 것.

    Args:
        visible: 아래아한글을 백그라운드에서 실행할지(False), 혹은 화면에 보이게 할지(True) 결정하는 파라미터

    Returns:
        무조건 True를 리턴함
    """
    pythoncom.CoInitialize()
    hwp = Hwp(visible=visible)
    hwp.HAction.Run("EquationCreate")
    pythoncom.CoUninitialize()
    return True


def _eq_modify(visible) -> bool:
    """
    멀티스레드 형태로 기존 수식에 대한 수식편집기를 실행하는 헬퍼함수. 직접 사용하지 말 것.

    Args:
        visible: 아래아한글을 백그라운드에서 실행할지(False), 혹은 화면에 보이게 할지(True) 결정하는 파라미터

    Returns:
        무조건 True를 리턴
    """
    pythoncom.CoInitialize()
    hwp = Hwp(visible=visible)
    hwp.hwp.HAction.Run("EquationModify")
    pythoncom.CoUninitialize()
    return True


def _close_eqedit(save: bool = False, delay: float = 0.1) -> bool:
    """
    멀티스레드로 열린 수식편집기를 억지로 찾아 닫는 헬퍼함수. 직접 사용하지 말 것.

    Args:
        save: 수식편집기를 닫기 전에 저장할지 결정
        delay: 실행 지연시간

    Returns:
        성공시 True, 실패시 False

    """
    hwnd = 0
    while not hwnd:
        hwnd = win32gui.FindWindow(None, "수식 편집기")
        sleep(delay)
    win32gui.SendMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    sleep(delay)
    hwnd = win32gui.FindWindow(None, "수식")
    if hwnd:
        if save:
            win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, ord("Y"), 0)
            win32gui.SendMessage(hwnd, win32con.WM_KEYUP, ord("Y"), 0)
        else:
            win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, ord("N"), 0)
            win32gui.SendMessage(hwnd, win32con.WM_KEYUP, ord("N"), 0)
        return True
    else:
        return False


def crop_data_from_selection(data, selection) -> list[str]:
    # 리스트 a의 셀 주소를 바탕으로 데이터 범위를 추출하는 함수.
    # pyhwpx 내부적으로만 사용됨
    if not selection:
        return []

    # 셀 주소를 행과 열 인덱스로 변환
    indices = [addr_to_tuple(cell) for cell in selection]

    # 범위 계산
    min_row = min(idx[0] for idx in indices)
    max_row = max(idx[0] for idx in indices)
    min_col = min(idx[1] for idx in indices)
    max_col = max(idx[1] for idx in indices)

    # 범위 추출
    result = []
    for row in range(min_row, max_row + 1):
        result.append(data[row][min_col:max_col + 1])

    return result


def check_registry_key(key_name:str="FilePathCheckerModule") -> bool:
    """
    아래아한글의 보안모듈 FilePathCheckerModule의 레지스트리에 등록여부 체크

    Args:
        key_name: 아래아한글 보안모듈 키 이름. 기본값은 "FilePathCheckerModule"

    Returns:
        등록되어 있는 경우 True, 미등록인 경우 False
    """
    from winreg import ConnectRegistry, HKEY_CURRENT_USER, OpenKey, CloseKey
    winup_path = r"Software\HNC\HwpAutomation\Modules"
    alt_winup_path = r"Software\Hnc\HwpUserAction\Modules"
    reg_handle = ConnectRegistry(None, HKEY_CURRENT_USER)

    for path in [winup_path, alt_winup_path]:
        try:
            from winreg import KEY_READ
            key = OpenKey(reg_handle, path, 0, KEY_READ)
            try:
                value, regtype = QueryValueEx(key, key_name)
                if value and os.path.exists(value):
                    CloseKey(key)
                    return True
            except FileNotFoundError:
                pass
            CloseKey(key)
        except FileNotFoundError:
            pass
    return False


def rename_duplicates_in_list(file_list: list[str]) -> list[str]:
    """
    문서 내 이미지를 파일로 저장할 때, 동일한 이름의 파일 뒤에 (2), (3).. 붙여주는 헬퍼함수

    Args:
        file_list: 문서 내 이미지 파일명 목록

    Returns:
        중복된 이름이 두 개 이상 있는 경우 뒤에 "(2)", "(3)"을 붙인 새로운 문자열 리스트
    """
    counts = {}

    for i, item in enumerate(file_list):
        if item in counts:
            counts[item] += 1
            new_item = f"{os.path.splitext(item)[0]}({counts[item]}){os.path.splitext(item)[1]}"
        else:
            counts[item] = 0
            new_item = item

        file_list[i] = new_item

    return file_list


def check_tuple_of_ints(var: tuple) -> bool:
    """
    변수가 튜플이고 모든 요소가 int인지 확인하는 헬퍼함수

    Args:
        var: 이터러블 자료형. 일반적으로 튜플.

    Returns:
        튜플이면서, 요소들이 모두 int인 경우 True를 리턴, 그렇지 않으면 False

    """
    if isinstance(var, tuple):  # 먼저 변수가 튜플인지 확인
        return all(isinstance(item, int) for item in var)  # 모든 요소가 int인지 확인
    return False  # 변수가 튜플이 아니면 False 반환


def excel_address_to_tuple_zero_based(address: str) -> tuple[int | Any, int | Any]:
    """
    엑셀 셀 주소를 튜플로 변환하는 헬퍼함수

    """
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


# 아래아한글의 ctrl 래퍼 클래스 정의
class Ctrl:
    """
    아래아한글의 모든 개체(표, 그림, 글상자 및 각주/미주 등)를 다루기 위한 클래스.
    """
    def __init__(self, com_obj):
        self._com_obj = com_obj  # 원래 COM 객체

    def __repr__(self):
        return f"<CtrlCode: CtrlID={self.CtrlID}, CtrlCH={self.CtrlCh}, UserDesc={self.UserDesc}>"

    def GetCtrlInstID(self) -> str:
        """
        **[한글2024전용]** 컨트롤의 고유 아이디를 정수형태 문자열로 리턴하는 메서드

        한글2024부터 제공하는 기능으로 정확하게 컨트롤을 선택하기 위한 새로운 수단이다.
        기존의 ``FindCtrl()``, ``hwp.SelectCtrlFront()``나
        ``hwp.SelectCtrlReverse()`` 등 인접 컨트롤을 선택하는 방법에는
        문제의 소지가 있었다. 대표적인 예로, 이미지가 들어있는 셀 안에서
        표 컨트롤을 선택하려고 하면, 어떤 방법을 쓰든 이미지가 선택돼버리기 때문에
        이미지를 선택하지 않는 여러 꼼수를 생각해내야 했다.
        하지만 ctrl.GetCtrlInstID()와 hwp.SelectCtrl()을
        같이 사용하면 그럴 걱정이 전혀 없게 된다.

        다만 사용시 주의할 점이 하나 있는데,

        `Get`/`SetTextFile`이나 `save_block_as` 등의 메서드 혹은
        `Cut`/`Paste` 사용시에는, 문서상에서 컨트롤이 지워졌다 다시 씌어지는 시점에
        `CtrlInstID`가 바뀌게 된다. (다만, 마우스로 드래그해 옮길 땐 아이디가 바뀌지 않는다.)

        Returns:
            10자리 정수 형태의 문자열로 구성된 `CtrlInstID`를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_random_picture()
            >>> hwp.insert_random_picture()
            >>> hwp.insert_random_picture()
            >>> for ctrl in hwp.ctrl_list:
            ...     print(ctrl.GetCtrlInstID())
            ...
            1816447703
            1816447705
            1816447707
            >>> hwp.hwp.SelectCtrl("")
        """
        return self._com_obj.GetCtrlInstID()

    def GetAnchorPos(self, type_: int = 0) -> "Hwp.HParameterSet":
        """
        해당 컨트롤의 앵커(조판부호)의 위치를 반환한다.

        Args:
            type_:
                기준위치

                - 0: 바로 상위 리스트에서의 앵커 위치(기본값)
                - 1: 탑레벨 리스트에서의 앵커 위치
                - 2: 루트 리 스트에서의 앵커 위치

        Returns:
            성공했을 경우 ListParaPos ParameterSet이 반환된다. 실패했을 때는 None을 리턴함.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 표만 두 개 있는 문서에 연결됨.
            >>> for ctrl in hwp.ctrl_list:
            ...     print(ctrl.GetAnchorPos().Item("List"), end=" ")
            ...     print(ctrl.GetAnchorPos().Item("Para"), end=" ")
            ...     print(ctrl.GetAnchorPos().Item("Pos"))
            0 0 16
            0 2 0
        """
        return self._com_obj.GetAnchorPos(type=type_)

    @property
    def CtrlCh(self) -> int:
        """
        선택한 개체(Ctrl)의 타입 확인할 수 있는 컨트롤 문자를 리턴

        일반적으로 컨트롤 ID를 사용해 컨트롤의 종류를 판별하지만,
        이보다 더 포괄적인 범주를 나타내는 컨트롤 문자로 판별할 수도 있다.
        예를 들어 각주와 미주는 ID는 다르지만, 컨트롤 문자는 17로 동일하다.
        컨트롤 문자는 1부터 31사이의 값을 사용한다.
        (그럼에도, CtrlCh는 개인적으로 잘 사용하지 않는다.)

        Returns:
            1~31의 정수

                - 1: 예약
                - 2: 구역/단 정의
                - 3: 필드 시작
                - 4: 필드 끝
                - 5: 예약
                - 6: 예약
                - 7: 예약
                - 8: 예약
                - 9: 탭
                - 10: 강제 줄 나눔
                - 11: 그리기 개체 / 표
                - 12: 예약
                - 13: 문단 나누기
                - 14: 예약
                - 15: 주석
                - 16: 머리말 / 꼬리말
                - 17: 각주 / 미주
                - 18: 자동 번호
                - 19: 예약
                - 20: 예약
                - 21: 쪽바뀜
                - 22: 책갈피 / 찾아보기 표시
                - 23: 덧말 / 글자 겹침
                - 24: 하이픈
                - 25: 예약
                - 26: 예약
                - 27: 예약
                - 28: 예약
                - 29: 예약
                - 30: 묶음 빈칸
                - 31: 고정 폭 빈칸

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 표 두 개가 들어있는 문서에서
            >>> for ctrl in hwp.ctrl_list:
            ...     print(ctrl.CtrlCh)
            ...
            11
            11
        """
        return self._com_obj.CtrlCh

    @property
    def CtrlID(self) -> str:
        """
        컨트롤 아이디

        컨트롤 ID는 컨트롤의 종류를 나타내기 위해 할당된 ID로서, 최대 4개의 문자로 구성된 문자열이다.
        예를 들어 표는 "tbl", 각주는 "fn"이다. 이와 비슷하게 CtrlCh는 정수로, UserDesc는 한글 문자열로 리턴한다.

        한/글에서 현재까지 지원되는 모든 컨트롤의 ID는 아래 Returns 참조.

        Returns:
            해당 컨트롤의 컨트롤아이디

                - "cold" : (ColDef) 단
                - "secd" : (SecDef) 구역
                - "fn" : (FootnoteShape) 각주
                - "en" : (FootnoteShape) 미주
                - "tbl" : (TableCreation) 표
                - "eqed" : (EqEdit) 수식
                - "gso" : (ShapeObject) 그리기 개체
                - "atno" : (AutoNum) 번호 넣기
                - "nwno" : (AutoNum) 새 번호로
                - "pgct" : (PageNumCtrl) 페이지 번호 제어(97의 홀수 쪽에서 시작)
                - "pghd" : (PageHiding) 감추기
                - "pgnp" : (PageNumPos) 쪽 번호 위치
                - "head" : (HeaderFooter) 머리말
                - "foot" : (HeaderFooter) 꼬리말
                - "%dte" : (FieldCtrl) 현재의 날짜/시간 필드
                - "%ddt" : (FieldCtrl) 파일 작성 날짜/시간 필드
                - "%pat" : (FieldCtrl) 문서 경로 필드
                - "%bmk" : (FieldCtrl) 블록 책갈피
                - "%mmg" : (FieldCtrl) 메일 머지
                - "%xrf" : (FieldCtrl) 상호 참조
                - "%fmu" : (FieldCtrl) 계산식
                - "%clk" : (FieldCtrl) 누름틀
                - "%smr" : (FieldCtrl) 문서 요약 정보 필드
                - "%usr" : (FieldCtrl) 사용자 정보 필드
                - "%hlk" : (FieldCtrl) 하이퍼링크
                - "bokm" : (TextCtrl) 책갈피
                - "idxm" : (IndexMark) 찾아보기
                - "tdut" : (Dutmal) 덧말
                - "tcmt" : (None) 주석

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 2x2 표의 각 셀 안에 이미지가 총 4장 들어있는 문서
            >>> for ctrl in hwp.ctrl_list:
            ...     print(ctrl.CtrlID)
            ...
            tbl
            gso
            gso
            gso
            gso
        """
        return self._com_obj.CtrlID

    @property
    def HasList(self):
        return self._com_obj.HasList

    @property
    def Next(self) -> "Ctrl":
        """
        다음 컨트롤.

        문서 중의 모든 컨트롤(표, 그림 등의 특수 문자들)은 linked list로 서로 연결되어 있는데, list 중 다음 컨트롤을 나타낸다.

        Returns:
            현재 컨트롤의 다음 컨트롤

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 표 하나만 들어 있는 문서에서
            >>> print(hwp.HeadCtrl.Next.Next.UserDesc)
            표
        """
        return Ctrl(self._com_obj.Next)

    @property
    def Prev(self) -> "Ctrl":
        """
        앞 컨트롤.

        문서 중의 모든 컨트롤(표, 그림 등의 특수 문자들)은 linked list로 서로 연결되어 있는데, list 중 앞 컨트롤을 나타낸다.

        Returns:
            현재 컨트롤의 이전 컨트롤

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 빈 표, 그 아래에 그림 한 개 삽입된 문서에서
            >>> print(hwp.LastCtrl.Prev.UserDesc)
            표
        """
        return Ctrl(self._com_obj.Prev)

    @property
    def Properties(self):
        """
        컨트롤의 속성을 나타낸다.

        모든 컨트롤은 대응하는 parameter set으로 속성을 읽고 쓸 수 있다.

        Examples:
            >>> # 문서의 모든 그림의 너비, 높이를 각각 절반으로 줄이는 코드
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> for ctrl in hwp.ctrl_list:
            ...     if ctrl.UserDesc == "그림":
            ...         prop = ctrl.Properties
            ...         width = prop.Item("Width")
            ...         height = prop.Item("Height")
            ...         prop.SetItem("Width", width // 2)
            ...         prop.SetItem("Height", height // 2)
            ...         ctrl.Properties = prop
        """
        return self._com_obj.Properties

    @Properties.setter
    def Properties(self, prop: Any) -> None:
        self._com_obj.Properties = prop

    @property
    def UserDesc(self):
        """
        컨트롤의 종류를 사용자에게 보여줄 수 있는 localize된 문자열로 나타낸다.
        """
        return self._com_obj.UserDesc


class XHwpDocuments:
    """
    아래아한글의 문서 오브젝트를 조작하기 위한 XHwpDocuments 래퍼 클래스. (작성중)
    """
    def __init__(self, com_obj):
        self._com_obj = com_obj

    def __repr__(self):
        return f"<XHwpDocuments com_obj={self._com_obj}>"

    def __getitem__(self, index):
        count = len(self)
        if isinstance(index, int):
            if index < 0:
                index += count  # 음수 인덱스를 양수로 변환
            if 0 <= index < count:
                return XHwpDocument(self._com_obj.Item(index))
            else:
                raise IndexError("Index out of range")
        else:
            raise TypeError("Index must be an integer")

    def __iter__(self):
        for i in range(len(self)):
            yield XHwpDocument(self[i])

    def __len__(self):
        return self._com_obj.Count

    @property
    def Active_XHwpDocument(self):
        return XHwpDocument(self._com_obj.Active_XHwpDocument)

    @property
    def Application(self):
        return self._com_obj.Application

    @property
    def CLSID(self):
        return self._com_obj.CLSID

    @property
    def Count(self):
        return self._com_obj.Count

    def Add(self, isTab:bool=False) -> "XHwpDocument":
        """
        문서 추가

        Args:
            isTab: 탭으로 열 건지(True), 문서로 열 건지(False, 기본값) 결정

        Returns:
            문서 오브젝트(XHwpDocument) 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.XHwpDocuments.Add(True)

        """
        return XHwpDocument(self._com_obj.Add(isTab=isTab))

    def Close(self, isDirty=False) -> None:
        """
        문서창 닫기
        """
        return self._com_obj.Close(isDirty=isDirty)

    def FindItem(self, lDocID:int) -> "XHwpDocument":
        """
        해당 DocumentID의 문서 오브젝트가 있는지 탐색

        Args:
            lDocID: 찾고자 하는 문서오브젝트의 ID

        Returns:
            int: 해당 아이디의 XHwpDocument 리턴
            None: 없는 경우 None 리턴
        """
        return XHwpDocument(self._com_obj.FindItem(lDocID))


class XHwpDocument:
    def __init__(self, com_obj):
        self._com_obj = com_obj

    @property
    def Application(self):
        return self._com_obj.Application

    @property
    def CLSID(self):
        return self._com_obj.CLSID

    def Clear(self, option:bool=False) -> None:
        return self._com_obj.Clear(option=option)

    def Close(self, isDirty:bool=False) -> None:
        return self._com_obj.Close(isDirty=isDirty)

    @property
    def DocumentID(self):
        return self._com_obj.DocumentID

    @property
    def EditMode(self):
        return self._com_obj.EditMode

    @property
    def Format(self):
        return self._com_obj.Format

    @property
    def FullName(self):
        """
        문서의 전체경로 문자열. 저장하지 않은 빈 문서인 경우에는 빈 문자열 ''
        """
        return self._com_obj.FullName

    @property
    def Modified(self) -> int:
        return self._com_obj.Modified

    def Open(self, filename:str, Format:str, arg:str):
        return self._com_obj.Open(filename=filename, Format=Format, arg=arg)

    @property
    def Path(self) -> str:
        return self._com_obj.Path

    def Redo(self, Count:int):
        return self._com_obj.Redo(Count=Count)

    def Save(self, save_if_dirty:bool):
        return self._com_obj.Save(save_if_dirty=save_if_dirty)

    def SaveAs(self, Path:str, Format:str, arg:str):
        return self._com_obj.SaveAs(Path=Path, Format=Format, arg=arg)

    def SendBrowser(self):
        return self._com_obj.SendBrowser()

    def SetActive_XHwpDocument(self):
        return self._com_obj.SetActive_XHwpDocument()

    def Undo(self, Count:int):
        return self._com_obj.Undo(Count=Count)

    @property
    def XHwpCharacterShape(self):
        return self._com_obj.XHwpCharacterShape

    @property
    def XHwpDocumentInfo(self):
        return self._com_obj.XHwpDocumentInfo

    @property
    def XHwpFind(self):
        return self._com_obj.XHwpFind

    @property
    def XHwpFormCheckButtons(self):
        return self._com_obj.XHwpFormCheckButtons

    @property
    def XHwpFormComboBoxs(self):
        return self._com_obj.XHwpFormComboBoxs

    @property
    def XHwpFormEdits(self):
        return self._com_obj.XHwpFormEdits

    @property
    def XHwpFormPushButtons(self):
        return self._com_obj.XHwpFormPushButtons

    @property
    def XHwpFormRadioButtons(self):
        return self._com_obj.XHwpFormRadioButtons

    @property
    def XHwpParagraphShape(self):
        return self._com_obj.XHwpParagraphShape

    @property
    def XHwpPrint(self):
        return self._com_obj.XHwpPrint

    @property
    def XHwpRange(self):
        return self._com_obj.XHwpRange

    @property
    def XHwpSelection(self):
        return self._com_obj.XHwpSelection

    @property
    def XHwpSendMail(self):
        return self._com_obj.XHwpSendMail

    @property
    def XHwpSummaryInfo(self):
        return self._com_obj.XHwpSummaryInfo


# 아래아한글 오토메이션 클래스 정의
class Hwp:
    """
    아래아한글 인스턴스를 실행합니다.

    실행방법은 간단합니다. `from pyhwpx import Hwp`로 `Hwp` 클래스를 임포트한 후,
    `hwp = Hwp()` 명령어를 실행하면 아래아한글이 자동으로 열립니다.
    만약 기존에 아래아한글 창이 하나 이상 열려 있다면, 가장 마지막에 접근했던 아래아한글 창과 연결됩니다.

    Args:
        new (bool):
            `new=True` 인 경우, 기존에 열려 있는 한/글 인스턴스와 무관한 새 인스턴스를 생성하게 됩니다.
            `new=False` (기본값)인 경우, 우선적으로 기존에 열려 있는 한/글 창에 연결을 시도합니다. (연결되지 않기도 합니다.)
        visible (bool):
            한/글 인스턴스를 백그라운드에서 실행할지, 화면에 나타낼지 선택합니다.
            기본값은 `True` 이며, 한/글 창이 화면에 나타나게 됩니다.
            `visible=False` 파라미터를 추가할 경우 한/글 창이 보이지 않는 상태로 백그라운드에서 작업할 수 있습니다.
        register_module (bool):
            보안모듈을 Hwp 클래스에서 직접 실행하게 허용합니다. 기본값은 `True` 입니다.
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") 메서드를 직접 실행하는 것과 동일합니다.

    Examples:
        >>> from pyhwpx import Hwp
        >>> hwp = Hwp()
        >>> hwp.insert_text("Hello world!")
        True
        >>> hwp.save_as("./hello.hwp")
        True
        >>> hwp.clear()
        >>> hwp.quit()
    """

    def __repr__(self):
        return f"<Hwp: DocumentID={self.XHwpDocuments.Active_XHwpDocument.DocumentID}, Title=\"{self.get_title()}\", FullName=\"{self.XHwpDocuments.Active_XHwpDocument.FullName or None}\">"

    def __init__(self, new: bool = False, visible: bool = True, register_module: bool = True):
        self.hwp = 0
        self.htf_fonts = {
            "명조": {
                "FaceNameHangul": "명조",
                "FaceNameLatin": "명조",
                "FaceNameHanja": "명조",
                "FaceNameJapanese": "명조",
                "FaceNameOther": "명조",
                "FaceNameSymbol": "명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "고딕": {
                "FaceNameHangul": "고딕",
                "FaceNameLatin": "고딕",
                "FaceNameHanja": "명조",
                "FaceNameJapanese": "고딕",
                "FaceNameOther": "명조",
                "FaceNameSymbol": "명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "샘물": {
                "FaceNameHangul": "샘물",
                "FaceNameLatin": "산세리프",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "필기": {
                "FaceNameHangul": "필기",
                "FaceNameLatin": "필기",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명조": {
                "FaceNameHangul": "한양신명조",
                "FaceNameLatin": "한양신명조",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "견명조": {
                "FaceNameHangul": "한양견명조",
                "FaceNameLatin": "한양견명조",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명조 약자": {
                "FaceNameHangul": "한양신명조",
                "FaceNameLatin": "한양신명조",
                "FaceNameHanja": "신명조 약자",
                "FaceNameJapanese": "한양신명조V",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명조 간자": {
                "FaceNameHangul": "한양신명조",
                "FaceNameLatin": "한양신명조",
                "FaceNameHanja": "신명조 간자",
                "FaceNameJapanese": "한양신명조V",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "중고딕": {
                "FaceNameHangul": "한양중고딕",
                "FaceNameLatin": "한양중고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "중고딕 약자": {
                "FaceNameHangul": "한양중고딕",
                "FaceNameLatin": "한양중고딕",
                "FaceNameHanja": "중고딕 약자",
                "FaceNameJapanese": "한양중고딕V",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "중고딕 간자": {
                "FaceNameHangul": "한양중고딕",
                "FaceNameLatin": "한양중고딕",
                "FaceNameHanja": "중고딕 간자",
                "FaceNameJapanese": "한양중고딕V",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "견고딕": {
                "FaceNameHangul": "한양견고딕",
                "FaceNameLatin": "한양견고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "그래픽": {
                "FaceNameHangul": "한양그래픽",
                "FaceNameLatin": "한양그래픽",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "궁서": {
                "FaceNameHangul": "한양궁서",
                "FaceNameLatin": "한양궁서",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "해서 약자": {
                "FaceNameHangul": "한양궁서",
                "FaceNameLatin": "한양궁서",
                "FaceNameHanja": "해서 약자",
                "FaceNameJapanese": "한양신명조V",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "해서 간자": {
                "FaceNameHangul": "한양궁서",
                "FaceNameLatin": "한양궁서",
                "FaceNameHanja": "해서 간자",
                "FaceNameJapanese": "한양신명조V",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "가는공한": {
                "FaceNameHangul": "가는공한",
                "FaceNameLatin": "가는공한",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "중간공한": {
                "FaceNameHangul": "중간공한",
                "FaceNameLatin": "중간공한",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "굵은공한": {
                "FaceNameHangul": "굵은공한",
                "FaceNameLatin": "굵은공한",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "가는한": {
                "FaceNameHangul": "가는한",
                "FaceNameLatin": "가는한",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "중간한": {
                "FaceNameHangul": "중간한",
                "FaceNameLatin": "중간한",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "굵은한": {
                "FaceNameHangul": "굵은한",
                "FaceNameLatin": "굵은한",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "펜흘림": {
                "FaceNameHangul": "펜흘림",
                "FaceNameLatin": "펜흘림",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "복숭아": {
                "FaceNameHangul": "복숭아",
                "FaceNameLatin": "한양중고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "옥수수": {
                "FaceNameHangul": "옥수수",
                "FaceNameLatin": "옥수수",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "오이": {
                "FaceNameHangul": "오이",
                "FaceNameLatin": "오이",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "가지": {
                "FaceNameHangul": "가지",
                "FaceNameLatin": "오이",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "강낭콩": {
                "FaceNameHangul": "강낭콩",
                "FaceNameLatin": "오이",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "딸기": {
                "FaceNameHangul": "딸기",
                "FaceNameLatin": "옥수수",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "타이프": {
                "FaceNameHangul": "타이프",
                "FaceNameLatin": "타이프",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "헤드라인": {
                "FaceNameHangul": "태 헤드라인T",
                "FaceNameLatin": "한양견고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "가는헤드라인": {
                "FaceNameHangul": "태 가는 헤드라인T",
                "FaceNameLatin": "HCI Hollyhock",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "헤드라인D": {
                "FaceNameHangul": "태 헤드라인D",
                "FaceNameLatin": "Blippo Blk BT",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "가는헤드라인D": {
                "FaceNameHangul": "태 가는 헤드라인D",
                "FaceNameLatin": "Hobo BT",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "태 나무": {
                "FaceNameHangul": "태 나무",
                "FaceNameLatin": "태 나무",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 다운명조M": {
                "FaceNameHangul": "양재 다운명조M",
                "FaceNameLatin": "양재 다운명조M",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 본목각M": {
                "FaceNameHangul": "양재 본목각M",
                "FaceNameLatin": "양재 본목각M",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 소슬": {
                "FaceNameHangul": "양재 소슬",
                "FaceNameLatin": "양재 소슬",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 튼튼B": {
                "FaceNameHangul": "양재 튼튼B",
                "FaceNameLatin": "양재 튼튼B",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 참숯B": {
                "FaceNameHangul": "양재 참숯B",
                "FaceNameLatin": "양재 참숯B",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 둘기": {
                "FaceNameHangul": "양재 둘기",
                "FaceNameLatin": "양재 둘기",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 매화": {
                "FaceNameHangul": "양재 매화",
                "FaceNameLatin": "양재 매화",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 샤넬": {
                "FaceNameHangul": "양재 샤넬",
                "FaceNameLatin": "양재 샤넬",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 와당": {
                "FaceNameHangul": "양재 와당",
                "FaceNameLatin": "양재 와당",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "양재 이니셜": {
                "FaceNameHangul": "양재 이니셜",
                "FaceNameLatin": "양재 이니셜",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼명조": {
                "FaceNameHangul": "휴먼명조",
                "FaceNameLatin": "HCI Poppy",
                "FaceNameHanja": "한양신명조",
                "FaceNameJapanese": "한양신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼고딕": {
                "FaceNameHangul": "휴먼고딕",
                "FaceNameLatin": "HCI Hollyhock",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "가는안상수체": {
                "FaceNameHangul": "가는안상수체",
                "FaceNameLatin": "가는안상수체영문",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "중간안상수체": {
                "FaceNameHangul": "중간안상수체",
                "FaceNameLatin": "중간안상수체영문",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "굵은안상수체": {
                "FaceNameHangul": "굵은안상수체",
                "FaceNameLatin": "굵은안상수체영문",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼가는샘체": {
                "FaceNameHangul": "휴먼가는샘체",
                "FaceNameLatin": "중간한",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼중간샘체": {
                "FaceNameHangul": "휴먼중간샘체",
                "FaceNameLatin": "HCI Hollyhock Narrow",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼굵은샘체": {
                "FaceNameHangul": "휴먼굵은샘체",
                "FaceNameLatin": "한양견고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼가는팸체": {
                "FaceNameHangul": "휴먼가는팸체",
                "FaceNameLatin": "중간한",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼중간팸체": {
                "FaceNameHangul": "휴먼중간팸체",
                "FaceNameLatin": "HCI Hollyhock Narrow",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼굵은팸체": {
                "FaceNameHangul": "휴먼굵은팸체",
                "FaceNameLatin": "한양견고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "휴먼옛체": {
                "FaceNameHangul": "휴먼옛체",
                "FaceNameLatin": "한양궁서",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "한양중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 세명조": {
                "FaceNameHangul": "신명 세명조",
                "FaceNameLatin": "신명 세명조",
                "FaceNameHanja": "신명 세명조",
                "FaceNameJapanese": "신명 신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 신명조": {
                "FaceNameHangul": "신명 신명조",
                "FaceNameLatin": "신명 신명조",
                "FaceNameHanja": "신명 세명조",
                "FaceNameJapanese": "신명 신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 신신명조": {
                "FaceNameHangul": "신명 신신명조",
                "FaceNameLatin": "신명 신신명조",
                "FaceNameHanja": "신명 세명조",
                "FaceNameJapanese": "신명 신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 중명조": {
                "FaceNameHangul": "신명 중명조",
                "FaceNameLatin": "신명 중명조",
                "FaceNameHanja": "신명 중명조",
                "FaceNameJapanese": "신명 신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 태명조": {
                "FaceNameHangul": "신명 태명조",
                "FaceNameLatin": "신명 태명조",
                "FaceNameHanja": "신명 태명조",
                "FaceNameJapanese": "신명 태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "신명 견명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 견명조": {
                "FaceNameHangul": "신명 견명조",
                "FaceNameLatin": "신명 견명조",
                "FaceNameHanja": "신명 견명조",
                "FaceNameJapanese": "신명 견명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "신명 견명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 신문명조": {
                "FaceNameHangul": "신명 신문명조",
                "FaceNameLatin": "신명 신문명조",
                "FaceNameHanja": "신명 신문명조",
                "FaceNameJapanese": "신명 신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 순명조": {
                "FaceNameHangul": "신명 순명조",
                "FaceNameLatin": "신명 순명조",
                "FaceNameHanja": "신명 중명조",
                "FaceNameJapanese": "신명 태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 세고딕": {
                "FaceNameHangul": "신명 세고딕",
                "FaceNameLatin": "신명 세고딕",
                "FaceNameHanja": "신명 세고딕",
                "FaceNameJapanese": "신명 중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 중고딕": {
                "FaceNameHangul": "신명 중고딕",
                "FaceNameLatin": "신명 중고딕",
                "FaceNameHanja": "신명 중고딕",
                "FaceNameJapanese": "신명 중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 태고딕": {
                "FaceNameHangul": "신명 태고딕",
                "FaceNameLatin": "신명 태고딕",
                "FaceNameHanja": "신명 태고딕",
                "FaceNameJapanese": "신명 태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "신명 태그래픽",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 견고딕": {
                "FaceNameHangul": "신명 견고딕",
                "FaceNameLatin": "신명 견고딕",
                "FaceNameHanja": "신명 견고딕",
                "FaceNameJapanese": "신명 태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "신명 견고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 세나루": {
                "FaceNameHangul": "신명 세나루",
                "FaceNameLatin": "신명 세나루",
                "FaceNameHanja": "신명 중고딕",
                "FaceNameJapanese": "신명 중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 디나루": {
                "FaceNameHangul": "신명 디나루",
                "FaceNameLatin": "신명 디나루",
                "FaceNameHanja": "신명 중고딕",
                "FaceNameJapanese": "신명 중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 신그래픽": {
                "FaceNameHangul": "신명 신그래픽",
                "FaceNameLatin": "신명 신그래픽",
                "FaceNameHanja": "신명 중고딕",
                "FaceNameJapanese": "신명 중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 태그래픽": {
                "FaceNameHangul": "신명 태그래픽",
                "FaceNameLatin": "신명 태그래픽",
                "FaceNameHanja": "신명 중고딕",
                "FaceNameJapanese": "신명 태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "신명 태그래픽",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "신명 궁서": {
                "FaceNameHangul": "신명 궁서",
                "FaceNameLatin": "신명 궁서",
                "FaceNameHanja": "신명 궁서",
                "FaceNameJapanese": "신명 신명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "한양신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#세명조": {
                "FaceNameHangul": "#세명조",
                "FaceNameLatin": "#세명조",
                "FaceNameHanja": "#신명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#세명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신명조": {
                "FaceNameHangul": "#신명조",
                "FaceNameLatin": "#신명조",
                "FaceNameHanja": "#신명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#중명조": {
                "FaceNameHangul": "#중명조",
                "FaceNameLatin": "#중명조",
                "FaceNameHanja": "#중명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#중명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신중명조": {
                "FaceNameHangul": "#신중명조",
                "FaceNameLatin": "#신중명조",
                "FaceNameHanja": "#중명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#중명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#화명조A": {
                "FaceNameHangul": "#화명조A",
                "FaceNameLatin": "#화명조A",
                "FaceNameHanja": "#중명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#중명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#화명조B": {
                "FaceNameHangul": "#화명조B",
                "FaceNameLatin": "#화명조B",
                "FaceNameHanja": "#중명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#중명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#태명조": {
                "FaceNameHangul": "#태명조",
                "FaceNameLatin": "#태명조",
                "FaceNameHanja": "#태명조",
                "FaceNameJapanese": "#태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#태명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신태명조": {
                "FaceNameHangul": "#신태명조",
                "FaceNameLatin": "#신태명조",
                "FaceNameHanja": "#태명조",
                "FaceNameJapanese": "#태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신태명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#태신명조": {
                "FaceNameHangul": "#태신명조",
                "FaceNameLatin": "#태신명조",
                "FaceNameHanja": "#태명조",
                "FaceNameJapanese": "#태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#태신명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#견명조": {
                "FaceNameHangul": "#견명조",
                "FaceNameLatin": "#견명조",
                "FaceNameHanja": "#견명조",
                "FaceNameJapanese": "#태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#견명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신문명조": {
                "FaceNameHangul": "#신문명조",
                "FaceNameLatin": "#신문명조",
                "FaceNameHanja": "#신문명조",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신문명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신문태명": {
                "FaceNameHangul": "#신문태명",
                "FaceNameLatin": "#신문태명",
                "FaceNameHanja": "#태명조",
                "FaceNameJapanese": "#태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#태명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신문견명": {
                "FaceNameHangul": "#신문견명",
                "FaceNameLatin": "#신문견명",
                "FaceNameHanja": "#신문견명",
                "FaceNameJapanese": "#태명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#견명조",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#세고딕": {
                "FaceNameHangul": "#세고딕",
                "FaceNameLatin": "#세고딕",
                "FaceNameHanja": "#신세고딕",
                "FaceNameJapanese": "#세고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#세고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신세고딕": {
                "FaceNameHangul": "#신세고딕",
                "FaceNameLatin": "#신세고딕",
                "FaceNameHanja": "#신세고딕",
                "FaceNameJapanese": "#세고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신세고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#중고딕": {
                "FaceNameHangul": "#중고딕",
                "FaceNameLatin": "#중고딕",
                "FaceNameHanja": "#중고딕",
                "FaceNameJapanese": "#중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#중고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#태고딕": {
                "FaceNameHangul": "#태고딕",
                "FaceNameLatin": "#태고딕",
                "FaceNameHanja": "#태고딕",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#태고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#견고딕": {
                "FaceNameHangul": "#견고딕",
                "FaceNameLatin": "#견고딕",
                "FaceNameHanja": "#견고딕",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#견고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신문고딕": {
                "FaceNameHangul": "#신문고딕",
                "FaceNameLatin": "#신문고딕",
                "FaceNameHanja": "#신문고딕",
                "FaceNameJapanese": "#중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신문고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신문태고": {
                "FaceNameHangul": "#신문태고",
                "FaceNameLatin": "#신문태고",
                "FaceNameHanja": "#태고딕",
                "FaceNameJapanese": "#중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#태고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신문견고": {
                "FaceNameHangul": "#신문견고",
                "FaceNameLatin": "#신문견고",
                "FaceNameHanja": "#신문견고",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#견고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#세나루": {
                "FaceNameHangul": "#세나루",
                "FaceNameLatin": "#세나루",
                "FaceNameHanja": "#신세나루",
                "FaceNameJapanese": "#세고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#세나루",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신세나루": {
                "FaceNameHangul": "#신세나루",
                "FaceNameLatin": "#신세나루",
                "FaceNameHanja": "#신세나루",
                "FaceNameJapanese": "#세고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신세나루",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#디나루": {
                "FaceNameHangul": "#디나루",
                "FaceNameLatin": "#디나루",
                "FaceNameHanja": "#신디나루",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#디나루",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신디나루": {
                "FaceNameHangul": "#신디나루",
                "FaceNameLatin": "#신디나루",
                "FaceNameHanja": "#신디나루",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신디나루",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#그래픽": {
                "FaceNameHangul": "#그래픽",
                "FaceNameLatin": "#그래픽",
                "FaceNameHanja": "#신세고딕",
                "FaceNameJapanese": "#중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#그래픽",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#신그래픽": {
                "FaceNameHangul": "#신그래픽",
                "FaceNameLatin": "#신그래픽",
                "FaceNameHanja": "#중고딕",
                "FaceNameJapanese": "#중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#신그래픽",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#태그래픽": {
                "FaceNameHangul": "#태그래픽",
                "FaceNameLatin": "#태그래픽",
                "FaceNameHanja": "#태고딕",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#태그래픽",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#궁서": {
                "FaceNameHangul": "#궁서",
                "FaceNameLatin": "#궁서",
                "FaceNameHanja": "#궁서",
                "FaceNameJapanese": "#세명조",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#궁서",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#공작": {
                "FaceNameHangul": "#공작",
                "FaceNameLatin": "#공작",
                "FaceNameHanja": "#중고딕",
                "FaceNameJapanese": "#중고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#공작",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#수암A": {
                "FaceNameHangul": "#수암A",
                "FaceNameLatin": "#수암A",
                "FaceNameHanja": "#태고딕",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#수암A",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#수암B": {
                "FaceNameHangul": "#수암B",
                "FaceNameLatin": "#수암B",
                "FaceNameHanja": "#태고딕",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#수암A",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "#빅": {
                "FaceNameHangul": "#빅",
                "FaceNameLatin": "#빅",
                "FaceNameHanja": "#견고딕",
                "FaceNameJapanese": "#태고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "#빅",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "시스템": {
                "FaceNameHangul": "시스템",
                "FaceNameLatin": "시스템",
                "FaceNameHanja": "시스템",
                "FaceNameJapanese": "시스템",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "시스템",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "시스템 약자": {
                "FaceNameHangul": "시스템",
                "FaceNameLatin": "시스템",
                "FaceNameHanja": "시스템 약자",
                "FaceNameJapanese": "시스템",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "시스템",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "시스템 간자": {
                "FaceNameHangul": "시스템",
                "FaceNameLatin": "시스템",
                "FaceNameHanja": "시스템 간자",
                "FaceNameJapanese": "시스템",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "시스템",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            },
            "HY둥근고딕": {
                "FaceNameHangul": "HY둥근고딕",
                "FaceNameLatin": "HY둥근고딕",
                "FaceNameHanja": "한양중고딕",
                "FaceNameJapanese": "HY둥근고딕",
                "FaceNameOther": "한양신명조",
                "FaceNameSymbol": "HY둥근고딕",
                "FaceNameUser": "명조",
                "FontTypeHangul": 2,
                "FontTypeHanja": 2,
                "FontTypeJapanese": 2,
                "FontTypeLatin": 2,
                "FontTypeOther": 2,
                "FontTypeSymbol": 2,
                "FontTypeUser": 2
            }
        }
        context = pythoncom.CreateBindCtx(0)
        pythoncom.CoInitialize()  # 이걸 꼭 실행해야 하는가? 왜 Pycharm이나 주피터에서는 괜찮고, vscode에서는 CoInitialize 오류가 나는지?
        running_coms = pythoncom.GetRunningObjectTable()
        monikers = running_coms.EnumRunning()

        if not new:
            for moniker in monikers:
                name = moniker.GetDisplayName(context, moniker)
                if name.startswith('!HwpObject.'):
                    obj = running_coms.GetObject(moniker)
                    self.hwp = win32.gencache.EnsureDispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch))
        if not self.hwp:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try:
            self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible
        except Exception as e:
            # print(e)
            sleep(0.01)
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible

        if register_module:  # and not check_registry_key():
            try:
                self.register_module()
            except Exception as e:
                print(e, "RegisterModule 액션을 실행할 수 없음. 개발자에게 문의해주세요.")

    @property
    def Application(self) -> "Hwp.Application":
        """
        저수준의 아래아한글 오토메이션API에 직접 접근하기 위한 속성입니다.

        `hwp.Application.~~~` 로 실행 가능한 모든 속성은, 간단히 `hwp.~~~` 로 실행할 수도 있지만
        pyhwpx와 API의 작동방식을 동일하게 하기 위해 구현해 두었습니다.

        Returns:
        저수준의 HwpApplication 객체

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.Application.XHwpWindows.Item(0).Visible = True
        """
        return self.hwp.Application

    @property
    def CellShape(self) -> Any:
        """
        셀(또는 표) 모양을 관리하는 파라미터셋 속성입니다.

        Returns:
        CellShape 파라미터셋

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.CellShape.Item("Height")  # 현재 표의 높이를 HwpUnit 단위로 리턴
            6410
            >>> hwp.HwpUnitToMili(hwp.CellShape.Item("Height"))
            22.6131
            >>> hwp.get_table_height()  # 위와 동일한 값을 리턴함
            22.6131
            >>> hwp.get_row_height()  # 현재 셀의 높이를 밀리미터 단위로 리턴
            4.5226
        """
        return self.hwp.CellShape

    @CellShape.setter
    def CellShape(self, prop: Any) -> None:
        self.hwp.CellShape = prop

    @property
    def CharShape(self) -> "Hwp.CharShape":
        """
        글자모양 파라미터셋을 조회하거나 업데이트할 수 있는 파라미터셋 속성.

        여러 속성값을 조회하고 싶은 경우에는 hwp.CharShape 대신
        `hwp.get_charshape_as_dict()` 메서드를 사용하면 편리합니다.
        CharShape 속성을 변경할 때는 아래 예시처럼
        hwp.set_font() 함수를 사용하는 것을 추천합니다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>>
            >>> # 현재 캐럿위치 또는 선택영역의 글자크기를 포인트단위로 출력
            >>> hwp.HwpUnitToPoint(hwp.CharShape.Item("Height"))
            10.0
            >>> # 여러 속성값을 확인하고 싶은 경우에는 아래처럼~
            >>> hwp.get_charshape_as_dict()
            {'Bold': 0,
             'BorderFill': <win32com.gen_py.HwpObject 1.0 Type Library.IDHwpParameterSet instance at 0x2267681890512>,
             'DiacSymMark': 0,
             'Emboss': 0,
             'Engrave': 0,
             'FaceNameHangul': '함초롬바탕',
             'FaceNameHanja': '함초롬바탕',
             'FaceNameJapanese': '함초롬바탕',
             'FaceNameLatin': '함초롬바탕',
             'FaceNameOther': '함초롬바탕',
             'FaceNameSymbol': '함초롬바탕',
             'FaceNameUser': '함초롬바탕',
             'FontTypeHangul': 1,
             'FontTypeHanja': 1,
             'FontTypeJapanese': 1,
             'FontTypeLatin': 1,
             'FontTypeOther': 1,
             'FontTypeSymbol': 1,
             'FontTypeUser': 1,
             'HSet': None,
             'Height': 2000,
             'Italic': 0,
             'OffsetHangul': 0,
             'OffsetHanja': 0,
             'OffsetJapanese': 0,
            ...
             'UnderlineColor': 0,
             'UnderlineShape': 0,
             'UnderlineType': 0,
             'UseFontSpace': 0,
             'UseKerning': 0}
            >>>
            >>> # 속성을 변경하는 예시
            >>> prop = hwp.CharShape  # 글자속성 개체를 복사한 후
            >>> prop.SetItem("Height", hwp.PointToHwpUnit(20))  # 파라미터 아이템 변경 후
            >>> hwp.CharShape = prop  # 글자속성을 prop으로 업데이트
            >>>
            >>> # 위 세 줄의 코드는 간단히 아래 단축메서드로도 실행가능
            >>> hwp.set_font(Height=30)  # 글자크기를 30으로 변경
        """
        return self.hwp.CharShape

    @CharShape.setter
    def CharShape(self, prop: Any) -> None:
        self.hwp.CharShape = prop

    @property
    def CLSID(self):
        """
        파라미터셋의 CLSID(클래스아이디)를 조회함. 읽기전용 속성이며, 사용할 일이 없음..

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.CharShape.CLSID
            IID('{599CBB08-7780-4F3B-8ADA-7F2ECFB57181}')
        """
        return self.hwp.CLSID

    @property
    def coclass_clsid(self):
        """
        coclass의 clsid를 리턴하는 읽기전용 속성. 사용하지 않음.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.coclass_clsid
            IID('{2291CF00-64A1-4877-A9B4-68CFE89612D6}')
        """
        return self.hwp.coclass_clsid

    @property
    def CurFieldState(self):
        """
        현재 캐럿이 들어있는 영역의 상태를 조회할 수 있는 속성.

        필드 안에 들어있지 않으면(본문, 캡션이나 주석 포함) 0을 리턴하며,
        셀 안이면 1, 글상자 안이면 4를 리턴합니다.
        셀필드 안에 있으면 17, 누름틀 안에 있으면 18을 리턴합니다.
        셀필드 안의 누름틀 안에서도 누름틀과 동일하게 18을 리턴하는 점에 유의하세요.
        정수값에 따라 현재 캐럿의 위치를 파악할 수 있기 때문에 다양하게 활용할 수 있습니다.
        예를 들어 필드와 무관하게 "캐럿이 셀 안에 있는가"를 알고 싶은 경우에도
        `hwp.CurFieldState` 가 1을 리턴하는지 확인하는 방식을 사용할 수 있습니다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 캐럿이 현재 표 안에 들어있는지 확인하고 싶은 경우
            >>> if hwp.CurFieldState == 1:
            ...     print("캐럿이 셀 안에 들어있습니다.")
            ... else:
            ...     print("캐럿이 셀 안에 들어있지 않습니다.")
            캐럿이 셀 안에 들어있습니다.
        """
        return self.hwp.CurFieldState

    @property
    def CurMetatagState(self) -> int:
        """
        (한글2024 이상) 현재 캐럿이 들어가 있는 메타태그 상태를 조회할 수 있는 속성.

        Returns:
        1: 셀 메타태그 영역에 들어있음
            4: 메타태그가 부여된 글상자 또는 그리기개체 컨트롤 내부의 텍스트 공간에 있음
            8: 메타태그가 부여된 이미지 또는 글맵시, 글상자 등의 컨트롤 선택상태임
            16: 메타태그가 부여된 표 컨트롤 선택 상태임
            32: 메타태그 영역에 들어있지 않음
            40: 컨트롤을 선택하고 있긴 한데, 메타태그는 지정되어 있지 않은 상태(8+32)
            64: 본문 메타태그 영역에 들어있음

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> if hwp.CurMetatagState == 1:
            ...     print("현재 캐럿이 셀 메타태그 영역에 들어있습니다.")
        """
        return self.hwp.CurMetatagState

    @property
    def CurSelectedCtrl(self):
        """
        현재 선택된 오브젝트의 컨트롤을 리턴하는 속성

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 문서의 첫 번째 표를 선택하고 "글자처럼 취급" 속성 켜기
            >>> hwp.get_into_nth_table()  # 문서 첫 번째 표의 A1 셀로 이동
            >>> hwp.SelectCtrlFront()  # 표 오브젝트 선택
            >>> ctrl = hwp.CurSelectedCtrl  # <-- 표 오브젝트의 컨트롤정보 변수지정
            >>> prop = ctrl.Properties  # 컨트롤정보의 속성(일종의 파라미터셋) 변수지정
            >>> prop.SetItem("TreatAsChar", True)  # 복사한 파라미터셋의 글자처럼취급 아이템값을 True로 변경
            >>> ctrl.Properties = prop  # 파라미터셋 속성을 표 오브젝트 컨트롤에 적용
            >>> hwp.Cancel()  # 적용을 마쳤으면 표선택 해제(권장)
        """
        return Ctrl(self.hwp.CurSelectedCtrl)

    @property
    def EditMode(self) -> int:
        """
        현재 편집모드(a.k.a. 읽기전용)를 리턴하는 속성.
        일반적으로 자동화에 쓸 일이 없으므로 무시해도 됩니다.
        편집모드로 변경하고 싶으면 1, 읽기전용으로 변경하고 싶으면 0 대입합니다.

        Returns:
        편집모드는 1을, 읽기전용인 경우 0을 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.EditMode
            1
            >>> hwp.EditMode = 0  # 읽기 전용으로 변경됨
        """
        return self.hwp.EditMode

    @EditMode.setter
    def EditMode(self, prop: int):
        self.hwp.EditMode = prop

    @property
    def EngineProperties(self):
        return self.hwp.EngineProperties

    @property
    def HAction(self):
        """
        한/글의 액션을 설정하고 실행하기 위한 속성.

        GetDefalut, Execute, Run 등의 메서드를 가지고 있습니다.
        저수준의 액션과 파라미터셋을 조합하여 기능을 실행할 때에 필요합니다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # "Hello world!" 문자열을 입력하는 액션
            >>> pset = hwp.HParameterSet.HInsertText
            >>> act_id = "InsertText"
            >>> pset.Text = "Hello world!\\r\\n"  # 줄바꿈 포함
            >>> hwp.HAction.Execute(act_id, pset.HSet)
            >>> # 위 네 줄의 명령어는 아래 방법으로도 실행 가능
            >>> hwp.insert_text("Hello world!")
            >>> hwp.BreakPara()  # 줄바꿈 메서드
            True
        """
        return self.hwp.HAction

    @property
    def HeadCtrl(self):
        """
        문서의 첫 번째 컨트롤을 리턴한다.

        문서의 첫 번째, 두 번째 컨트롤은 항상 "구역 정의"와 "단 정의"이다. (이 둘은 숨겨져 있음)
        그러므로 `hwp.HeadCtrl` 은 항상 구역정의(secd: section definition)이며,
        `hwp.HeadCtrl.Next` 는 단 정의(cold: column definition)이다.

        사용자가 삽입한 첫 번째 컨트롤은 항상 `hwp.HeadCtrl.Next.Next` 이다.

        HeadCtrl과 반대로 문서의 가장 마지막 컨트롤은 hwp.LastCtrl이며, 이전 컨트롤로 순회하려면
        `.Next` 대신 `.Prev` 를 사용하면 된다.
        hwp.HeadCtrl의 기본적인 사용법은 아래와 같다.

        Examples:
            >>> # 문서에 삽입된 모든 표의 "글자처럼 취급" 속성을 해제하는 코드
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> ctrl = hwp.HeadCtrl
            >>> while ctrl:
            ...     if ctrl.UserDesc == "표":  # 이제 ctrl 변수가 해당 표 컨트롤을 가리키고 있으므로
            ...         prop = ctrl.Properties
            ...         prop.SetItem("TreatAsChar", True)
            ...         ctrl.Properties = prop
            ...     ctrl = ctrl.Next
            >>> print("모든 표의 글자처럼 취급 속성 해제작업이 완료되었습니다.")
            모든 표의 글자처럼 취급 속성 해제작업이 완료되었습니다.
        """
        return Ctrl(self.hwp.HeadCtrl)

    @property
    def HParameterSet(self):
        """
        한/글에서 실행되는 대부분의 액션을 설정하는 데 필요한 파라미터셋들이 들어있는 속성.

        HAction과 HParameterSet을 조합하면 어떤 복잡한 동작이라도 구현해낼 수 있지만
        공식 API 문서를 읽으며 코딩하기보다는, 해당 동작을 한/글 내에서 스크립트매크로로 녹화하고
        녹화된 매크로에서 액션아이디와 파라미터셋을 참고하는 방식이 훨씬 효율적이다.
        HParameterSet을 활용하는 예시코드는 아래와 같다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.HParameterSet.HInsertText
            >>> pset.Text = "Hello world!"
            >>> hwp.HAction.Execute("InsertText", pset.HSet)
            True
        """
        return self.hwp.HParameterSet

    @property
    def IsEmpty(self) -> bool:
        """
        아무 내용도 들어있지 않은 빈 문서인지 여부를 나타낸다. 읽기전용임

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 특정 문서를 열고, 비어있는지 확인
            >>> hwp.open("./example.hwpx")
            >>> if hwp.IsEmpty:
            ...     print("빈 문서입니다.")
            ... else:
            ...     print("빈 문서가 아닙니다.")
            빈 문서가 아닙니다.
        """
        return self.hwp.IsEmpty

    @property
    def IsModified(self) -> bool:
        """
        최근 저장 또는 생성 이후 수정이 있는지 여부를 나타낸다. 읽기전용이며,
        자동화에 활용하는 경우는 거의 없다. 패스~

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.open("./example.hwpx")
            >>> # 신나게 작업을 마친 후, 종료하기 직전, 혹시 하는 마음에
            >>> # 수정사항이 있으면 저장하고 끄기 & 수정사항이 없으면 그냥 한/글 종료하기
            >>> if hwp.IsModified:
            ...     hwp.save()
            ... hwp.quit()
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

        연결리스트 타입으로, HeadCtrl부터 LastCtrl까지 모두 연결되어 있고
        LastCtrl.Prev.Prev 또는 HeadCtrl.Next.Next 등으로 컨트롤 순차 탐색이 가능하다.
        혹자는 `hwp.HeadCtrl` 만 있으면 되는 거 아닌가 생각할 수 있지만,
        특정 조건의 컨트롤을 삭제!!하는 경우 삭제한 컨트롤 이후의 모든 컨트롤의 인덱스가 변경되어버리므로
        이런 경우에는 LastCtrl에서 역순으로 진행해야 한다. (HeadCtrl부터 Next 작업을 하면 인덱스 꼬임)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 문서 내의 모든 그림 삭제하기
            >>> ctrl = hwp.LastCtrl  # <---
            >>> while ctrl:
            ...     if ctrl.UserDesc == "그림":
            ...         hwp.DeleteCtrl(ctrl)
            ...     ctrl = ctrl.Prev
            ... print("모든 그림을  삭제하였습니다.")
            모든 그림을 삭제하였습니다.
            >>> # 아래처럼 for문과 hwp.ctrl_list로도 구현할 수 있음
            >>> for ctrl in [i for i in hwp.ctrl_list if i.UserDesc == "그림"][::-1]:  # 역순 아니어도 무관.
            ...     hwp.DeleteCtrl(ctrl)
        """
        return Ctrl(self.hwp.LastCtrl)

    @property
    def PageCount(self) -> int:
        """
        현재 문서의 총 페이지 수를 리턴.

        Returns:
            현재 문서의 총 페이지 수

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.open("./example.hwpx")
            >>> print(f"현재 이 문서의 총 페이지 수는 {hwp.PageCount}입니다.")
            현재 이 문서의 총 페이지 수는 20입니다.
        """
        return self.hwp.PageCount

    @property
    def ParaShape(self):
        """
        CharShape, CellShape과 함께 가장 많이 사용되는 단축Shape 삼대장 중 하나.

        현재 캐럿이 위치한, 혹은 선택한 문단(블록)의 문단모양 파라미터셋을 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.open("./example.hwp")
            >>> # 현재 캐럿위치의 줄간격 조회
            >>> val = hwp.ParaShape.Item("LineSpacing")
            >>> print(f"현재 문단의 줄간격은 {val}%입니다.")
            현재 문단의 줄간격은 160%입니다.
            >>>
            >>> # 본문 모든 문단의 줄간격을 200%로 수정하기
            >>> hwp.SelectAll()  # 전체선택
            >>> prop = hwp.ParaShape
            >>> prop.SetItem("LineSpacing", 200)
            >>> hwp.ParaShape = prop
            >>> print("본문 전체의 줄간격을 200%로 수정하였습니다.")
            본문 전체의 줄간격을 200%로 수정하였습니다.
        """
        return self.hwp.ParaShape

    @ParaShape.setter
    def ParaShape(self, prop):
        self.hwp.ParaShape = prop

    @property
    def ParentCtrl(self) -> Ctrl:
        """
        현재 선택되어 있거나, 캐럿이 들어있는 컨트롤을 포함하는 상위 컨트롤을 리턴한다.

        Returns:
            상위 컨트롤(Ctrl)
        """
        return Ctrl(self.hwp.ParentCtrl)

    @property
    def Path(self) -> str:
        """
        현재 빈 문서가 아닌 경우, 열려 있는 문서의 파일명을 포함한 전체경로를 리턴한다.

        Returns:
        현재 문서의 전체경로

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.open("./example.hwpx")
            >>> hwp.Path
            C:/Users/User/desktop/example.hwpx
        """
        return self.hwp.Path

    @property
    def SelectionMode(self) -> int:
        """
        현재 선택모드가 어떤 상태인지 리턴한다.

        Returns:
        """
        return self.hwp.SelectionMode

    @property
    def Title(self) -> str:
        """
        현재 연결된 아래아한글 창의 제목표시줄 타이틀을 리턴한다.

        Returns:
            현재 연결된 아래아한글 창의 제목표시줄 타이틀

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> print(hwp.Title)
            빈 문서 1 - 한글
        """
        return self.get_title()

    @property
    def Version(self) -> list[int]:
        """
        아래아한글 프로그램의 버전을 리스트로 리턴한다.

        Returns:
            아래아한글 프로그램의 버전(문서 버전이 아님)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.Version
            [13, 0, 0, 2151]
        """
        return [int(i) for i in self.hwp.Version.split(", ")]

    @property
    def ViewProperties(self):
        """
        현재 한/글 프로그램의 보기 속성 파라미터셋을 리턴한다.

        Returns:
        """
        return self.hwp.ViewProperties

    @ViewProperties.setter
    def ViewProperties(self, prop):
        self.hwp.ViewProperties = prop

    @property
    def XHwpDocuments(self):
        """
        HwpApplication의 XHwpDocuments 객체를 리턴한다.

        Returns:
        """
        return self.hwp.XHwpDocuments

    @property
    def XHwpMessageBox(self) -> "Hwp.XHwpMessageBox":
        """
        메시지박스 객체 리턴

        Returns:
        """
        return self.hwp.XHwpMessageBox

    @property
    def XHwpODBC(self) -> "Hwp.XHwpODBC":
        return self.hwp.XHwpODBC

    @property
    def XHwpWindows(self) -> "Hwp.XHwpWindows":
        return self.hwp.XHwpWindows

    @property
    def ctrl_list(self) -> list:
        """
        문서 내 모든 ctrl를 리스트로 반환한다.

        단, 기본으로 삽입되어 있는 두 개의 컨트롤인
        secd(섹션정의)와 cold(단정의) 두 개는 어차피 선택불가하므로
        ctrl_list에서 제외했다.
        (모든 컨트롤을 제거하는 등의 경우, 편의를 위함)

        Returns:
            문서 내 모든 컨트롤의 리스트. 단, HeadCtrl(secd), HeadCtrl.Next(cold)는 포함하지 않는다.
        """
        c_list = []
        ctrl = self.hwp.HeadCtrl.Next.Next
        while ctrl:
            c_list.append(ctrl)
            ctrl = ctrl.Next
        return [Ctrl(i) for i in c_list]

    @property
    def current_page(self) -> int:
        """
        새쪽번호나 구역과 무관한 현재 쪽의 순서를 리턴.

        1페이지에 있다면 1을 리턴한다.
        새쪽번호가 적용되어 있어도
        페이지의 인덱스를 리턴한다.

        Returns:
            현재 쪽번호
        """
        return self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.CurrentPage + 1

    @property
    def current_printpage(self) -> int:
        """
        페이지인덱스가 아닌, 종이에 표시되는 쪽번호를 리턴.

        1페이지에 있다면 1을 리턴한다.
        새쪽번호가 적용되어 있다면
        수정된 쪽번호를 리턴한다.

        Returns:
        """
        return self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.CurrentPrintPage

    @property
    def current_font(self):
        charshape = self.get_charshape_as_dict()  # hwp.CharShape
        if charshape["FontTypeHangul"] == 1:
            return charshape["FaceNameHangul"]
        elif charshape["FontTypeHangul"] == 2:
            sub_dict = {key: value for key, value in charshape.items() if key.startswith('F')}
            for key, value in self.htf_fonts.items():
                if value == sub_dict:
                    return key

    # 커스텀 메서드
    def get_ctrl_pos(self, ctrl: Any = None, option: Literal[0, 1] = 0, as_tuple: bool = True) -> tuple[int, int, int]:
        """
        특정 컨트롤의 앵커(빨간 조판부호) 좌표를 리턴하는 메서드. 한글2024 미만의 버전에서, 컨트롤의 정확한 위치를 파악하기 위함

        Args:
            ctrl: 컨트롤 오브젝트. 특정하지 않으면 현재 선택된 컨트롤의 좌표를 리턴
            option:
                "표안의 표"처럼 컨트롤이 중첩된 경우에 어느 좌표를 리턴할지 결정할 수 있음

                    - 0: 현재 컨트롤이 포함된 리스트 기준으로 좌표 리턴
                    - 1: 현재 컨트롤을 포함하는 최상위 컨트롤 기준의 좌표 리턴

            as_tuple:
                리턴값을 (List, Para, Pos) 형태의 튜플로 리턴할지 여부. 기본값은 True.
                `as_tuple=False` 의 경우에는 ListParaPos 파라미터셋 자체를 리턴

        Returns:
            기본적으로 (List, Para, Pos) 형태의 튜플로 리턴하며, as_tuple=False 옵션 추가시에는 해당 ListParaPos 파라미터셋 자체를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 2x2표의 A1셀 안에 2x2표를 삽입하고, 표안의 표를 선택한 상태에서
            >>> # 컨트롤이 포함된 영역의 좌표를 리턴하려면(가장 많이 쓰임)
            >>> hwp.get_ctrl_pos()
            (3, 0, 0)
            >>> # 현재컨트롤을 포함한 최상위컨트롤의 본문기준 좌표를 리턴하려면
            >>> hwp.get_ctrl_pos(option=1)
            (0, 0, 16)
            >>> # 특정 컨트롤의 위치를 저장해 뒀다가 해당 위치로 이동하고 싶은 경우
            >>> pos = hwp.get_ctrl_pos(hwp.CurSelectedCtrl)  # 좌표 저장
            >>> # 모종의 작업으로 컨트롤 위치가 바뀌더라도, 컨트롤을 찾아갈 수 있음
            >>> hwp.set_pos(*pos)  # 해당 컨트롤 앞으로 이동함
            True
            >>> # 특정 컨트롤 위치 앞으로 이동하기 액션은 아래처럼도 실행 가능
            >>> hwp.move_to_ctrl(hwp.ctrl_list[-1])
            True
        """
        if ctrl is None:  # 컨트롤을 지정하지 않으면
            ctrl = self.CurSelectedCtrl  # 현재 선택중인 컨트롤
        if as_tuple:
            return (
                ctrl.GetAnchorPos(option).Item("List"),
                ctrl.GetAnchorPos(option).Item("Para"),
                ctrl.GetAnchorPos(option).Item("Pos"),
            )
        else:
            return ctrl.GetAnchorPos(option)

    def get_linespacing(self, method: Literal["Fixed", "Percent", "BetweenLines", "AtLeast"] = "Percent") -> int | float:
        """
        현재 캐럿 위치의 줄간격(%) 리턴.

        ![get_linespacing](assets/get_linespacing.gif){ loading=lazy }

        단, 줄간격 기준은 "글자에 따라(%)" 로 설정되어 있어야 하며,
        "글자에 따라"가 아닌 경우에는 method 파라미터를 실제 옵션과 일치시켜야 함.

        Args:
            method:
                줄간격 단위기준. 일치하지 않아도 값은 출력되지만, 단위를 모르게 됨..

                    - "Fixed": 고정값(포인트 단위)
                    - "Percent": 글자에 따라(기본값, %)
                    - "BetweenLines": 여백만 지정(포인트 단위)
                    - "AtLeast": 최소(포인트 단위)

        Returns:
            현재 캐럿이 위치한 문단의 줄간격(% 또는 Point). method에 따라 값이 바뀌므로 주의.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_linespacing()
            160
            >>> # 줄간격을 "최소" 기준 17.0point로 설정했다면, 아래처럼 실행해야 함
            >>> hwp.get_linespacing("AtLeast")
            170

        """
        act = "ParagraphShape"
        pset = self.hwp.HParameterSet.HParaShape
        self.hwp.HAction.GetDefault(act, pset.HSet)
        if pset.LineSpacingType == self.hwp.LineSpacingMethod(method):
            return pset.LineSpacing
        else:  # 어찌됐든 포인트 단위로 리턴
            return self.HwpUnitToPoint(pset.LineSpacing / 2)  # 이상하게 1/2 곱해야 맞다.

    def set_linespacing(self, value: int | float = 160,
                        method: Literal["Fixed", "Percent", "BetweenLines", "AtLeast"] = "Percent") -> bool:
        """
        현재 캐럿 위치의 문단 또는 선택 블록의 줄간격(%) 설정

        Args:
            value: 줄간격 값("Percent"인 경우에는 %, 그 외에는 point 값으로 적용됨). 기본값은 160(%)
            method:
                줄간격 단위기준. method가 일치해야 정상적으로 적용됨.

                    - "Fixed": 고정값(포인트 단위)
                    - "Percent": 글자에 따라(기본값, %)
                    - "BetweenLines": 여백만 지정(포인트 단위)
                    - "AtLeast": 최소(포인트 단위)


        Returns:
        성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.SelectAll()  # 전체선택
            >>> hwp.set_linespacing(160)  # 본문 모든 문단의 줄간격을 160%로 변경
            True
            >>> hwp.set_linespacing(20, method="BetweenLines")  # 본문 모든 문단의 줄간격을 "여백만 지정"으로 20pt 적용
            True
        """
        act = "ParagraphShape"
        pset = self.hwp.HParameterSet.HParaShape
        self.hwp.HAction.GetDefault(act, pset.HSet)
        pset.LineSpacingType = self.hwp.LineSpacingMethod(method)  # Percent
        if method == "Percent":
            pset.LineSpacing = value  # %값 그대로 넣으면 됨
        else:  # # 그 외에는 입력한 값을 아래아한글이 HwpUnit 단위라고 간주하므로
            pset.LineSpacing = value * 200  # HwpUnit 단위로 변환 후
        return self.hwp.HAction.Execute(act, pset.HSet)

    def is_empty_para(self) -> bool:
        """
        본문의 문단을 순회하면서 특정 서식을 적용할 때
        빈 문단에서 MoveNext~ 또는 MovePrev~ 등의 액션이 오작동하므로 이를 방지하기 위한 개발자용 헬퍼메서드.
        단독으로는 활용하지 말 것.

        Returns:
        빈 문단일 경우 제자리에서 True, 비어있지 않은 경우 False를 리턴

        """
        self.MoveSelNextChar()
        if self.get_pos()[2] == 0:  # 빈 문단이면?
            self.Cancel()
            self.MovePrevParaEnd()
            return True
        else:
            self.MoveParaBegin()
            return False

    def goto_addr(self, addr: str|int = "A1", col: int=0, select_cell: bool=False) -> bool:
        """
        셀 주소를 문자열로 입력받아 해당 주소로 이동하는 메서드.

        !!! warning "Deprecated"
            이 기능은 더 이상 사용되지 않으며, 다음 마이너 업데이트에서 제거될 예정입니다.
            대신 작업할 표 안에서 직접 `hwp.fill_addr_field()` 메서드를 실행하여
            셀주소 셀필드를 채운 후 `hwp.move_to_field()`를 통해 이동하는 방식을 사용해 주시기 바랍니다.
            사용 후에는 `hwp.unfill_addr_field()` 메서드를 통해 초기화를 해주셔야 합니다.

        셀 주소는 "C3"처럼 문자열로 입력하거나, 행번호, 열번호를 입력할 수 있음. 시작값은 1.

        Args:
            addr: 셀 주소 문자열 또는 행번호(1부터)
            col: 셀 주소를 정수로 입력하는 경우 열번호(1부터)
            select_cell: 이동 후 셀블록 선택 여부

        Returns:
           이동 성공 여부(성공시 True/실패시 False)
        """
        if not self.is_cell():
            return False  # 표 안에 있지 않으면 False 리턴(종료)
        if type(addr) == int and col:  # "A1" 대신 (1, 1) 처럼 tuple(int, int) 방식일 경우
            addr = tuple_to_addr(addr, col)  # 문자열 "A1" 방식으로 우선 변환

        refresh = False

        # 우선 A1 셀로 이동 시도
        self.HAction.Run("TableColBegin")
        self.HAction.Run("TableColPageUp")

        if addr.upper() == "A1":  # A1 셀이 맞으면!
            if select_cell:
                self.HAction.Run("TableCellBlock")
            return True

        init = self.get_pos()[0]  # 무조건 A1임.
        try:
            if self.addr_info[0] == init:
                pass
            else:
                refresh = True
                self.addr_info = [init, ["A1"]]
        except AttributeError:
            refresh = True
            self.addr_info = [init, ["A1"]]

        if refresh:
            i = 1
            while self.set_pos(init + i, 0, 0):
                cur_addr = self.KeyIndicator()[-1][1:].split(")")[0]
                if cur_addr == "A1":
                    break
                self.addr_info[1].append(cur_addr)
                i += 1
        try:
            self.set_pos(init + self.addr_info[1].index(addr.upper()), 0, 0)
            if select_cell:
                self.HAction.Run("TableCellBlock")
            return True
        except ValueError:
            self.set_pos(init, 0, 0)
            return False

    def get_field_info(self) -> list[dict]:
        """
        문서 내의 모든 누름틀의 정보(지시문 및 메모)를 추출하는 메서드.

        셀필드는 지시문과 메모가 없으므로 이 메서드에서는 추출하지 않는다.
        만약 셀필드를 포함하여 모든 필드의 이름만 추출하고 싶다면
        ``hwp.get_field_list().split("\\r\\n")`` 메서드를 쓰면 된다.

        Returns:
            [{'name': 'zxcv', 'direction': 'adsf', 'memo': 'qwer'}] 형식의 사전 리스트

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_field_info()
            [{'name': '누름틀1', 'direction': '안내문1', 'memo': '메모1'},
            {'name': '누름틀2', 'direction': '안내문2', 'memo': '메모2'}]

        """
        txt = self.GetTextFile("HWPML2X")
        try:
            root = ET.fromstring(txt)
            results = []
            for field in root.findall(".//FIELDBEGIN"):
                name_value = field.attrib.get("Name")
                command = re.split(r"(Clickhere:set:\d+:Direction:wstring:\d+:)|( HelpState:wstring:\d+:)",
                                   field.attrib.get("Command")[:-2])
                results.append({"name": name_value, "direction": command[3], "memo": command[-1]})
            return results
        except ET.ParseError as e:
            print("XML 파싱 오류:", e)
            return False
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다.")
            return False

    def get_image_info(self, ctrl: Any = None) -> dict[str:str, str:list[int, int]]:
        """
        이미지 컨트롤의 원본 그림의 이름과
        원본 그림의 크기 정보를 추출하는 메서드

        Args:
            ctrl: 아래아한글의 이미지 컨트롤. ctrl을 지정하지 않으면 현재 선택된 이미지의 정보를 추출

        Returns:
        해당 이미지의 삽입 전 파일명과, [Width, Height] 리스트

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 이미지 선택 상태에서
            >>> hwp.get_image_info()
            {'name': 'tmpmj2md6uy', 'size': [200, 200]}
            >>> # 문서 마지막그림 정보
            >>> ctrl = [i for i in hwp.ctrl_list if i.UserDesc == "그림"][-1]
            >>> hwp.get_image_info(ctrl)
            {'name': 'tmpxk_5noth', 'size': [1920, 1080]}
        """
        if ctrl is None:
            ctrl = self.CurSelectedCtrl

        if not ctrl or ctrl.UserDesc != "그림":
            return False
        self.select_ctrl(ctrl)
        block = self.GetTextFile("HWPML2X", option="saveblock:true")
        self.add_tab()
        self.SetTextFile(block, "HWPML2X")
        self.save_as("temp.xml", "HWPML2X")
        self.clear()
        self.FileClose()
        tree = ET.parse('temp.xml')
        root = tree.getroot()

        for shapeobject in root.findall('.//SHAPEOBJECT'):
            shapecmt = shapeobject.find('SHAPECOMMENT')
            if shapecmt is not None and shapecmt.text:
                info = shapecmt.text.split("\n")[1:]
                try:
                    return {"name": info[0].split(": ")[1],
                            "size": [int(i) for i in info[1][14:-5].split("pixel, 세로 ")]}
                finally:
                    os.remove("temp.xml")
        return False

    def goto_style(self, style: Union[int, str]) -> bool:
        """
        특정 스타일이 적용된 위치로 이동하는 메서드.

        탐색은 문서아랫방향으로만 수행하며 현재위치 이후 해당 스타일이 없거나,
        스타일이름/인덱스번호가 잘못된 경우 False를 리턴
        참고로, API의 Goto는 1부터 시작하므로 메서드 내부에서 인덱스에 1을 더하고 있음

        Args:
            style: 스타일이름(str) 또는 스타일번호(첫 번째 스타일이 0)

        Returns:
            성공시 True, 실패시 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.goto_style(0)  # 캐럿 뒤의 "바탕글" 스타일로 이동
            True
            >>> hwp.goto_style("개요 1")  # 캐럿 뒤의 "개요 1" 스타일로 이동
            True
        """
        if type(style) == int:
            style_idx = style + 1
        elif type(style) == str:
            style_dict = self.get_style_dict(as_="dict")
            if style in [style_dict[i]["name"] for i in style_dict]:
                style_idx = [i for i in style_dict if style_dict[i]["name"] == style][0] + 1
            else:
                return False
        pset = self.hwp.HParameterSet.HGotoE
        self.hwp.HAction.GetDefault("Goto", pset.HSet)
        pset.HSet.SetItem("DialogResult", style_idx)
        pset.SetSelectionIndex = 4
        cur_messagebox_mode = self.hwp.GetMessageBoxMode()
        self.hwp.SetMessageBoxMode(0x20000)
        try:
            if style == "바탕글" or 0:
                cur_pos = self.hwp.GetPos()
                self.hwp.HAction.Execute("Goto", pset.HSet)
                if self.hwp.GetPos() == cur_pos:
                    self.hwp.SetPos(*cur_pos)
                    return False
                else:
                    return True
            else:
                return self.hwp.HAction.Execute("Goto", pset.HSet)
        finally:
            self.hwp.SetMessageBoxMode(cur_messagebox_mode)

    def get_into_table_caption(self) -> bool:
        """
        표 캡션(정확히는 표번호가 있는 리스트공간)으로 이동하는 메서드.

        (추후 개선예정 : 캡션 스타일로 찾아가기 기능 추가할 것)

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        pset = self.hwp.HParameterSet.HGotoE
        pset.HSet.SetItem("DialogResult", 56)  # 표번호
        pset.SetSelectionIndex = 5  # 조판부호
        return self.hwp.HAction.Execute("Goto", pset.HSet)

    def shape_copy_paste(self, Type: Literal["font", "para", "both"] = "both", cell_attr: bool = False,
                         cell_border: bool = False, cell_fill: bool = False, cell_only: int = 0) -> bool:
        """
        모양복사 메서드

        ![introduce](assets/shape_copy_paste.gif){ loading=lazy }

        Args:
            Type: 글자("font"), 문단("para"), 글자&문단("both") 중에서 택일
            cell_attr: 셀 속성 복사여부(True / False)
            cell_border: 셀 선 복사여부(True / False)
            cell_fill: 셀 음영 복사여부(True / False)
            cell_only: 셀만 복사할지, 내용도 복사할지 여부(0 / 1)

        Returns:
            성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_into_nth_table(0)  # 문서 첫 번째 셀로 이동
            >>> hwp.shape_copy_paste()  # 모양복사 (무엇을 붙여넣을 건지는 "붙여넣기" 시점에 결정함)
            >>> hwp.TableCellBlockExtendAbs()
            >>> hwp.TableCellBlockExtend()  # 셀 전체 선택
            >>> hwp.shape_copy_paste(cell_fill=True)  # 글자&문단모양, 셀음영만 붙여넣기

        """
        pset = self.hwp.HParameterSet.HShapeCopyPaste
        pset.type = ["font", "para", "both"].index(Type)
        if self.is_cell():
            self.hwp.HParameterSet.HShapeCopyPaste.CellAttr = cell_attr
            self.hwp.HParameterSet.HShapeCopyPaste.CellBorder = cell_border
            self.hwp.HParameterSet.HShapeCopyPaste.CellFill = cell_fill
            self.hwp.HParameterSet.HShapeCopyPaste.TypeBodyAndCellOnly = cell_only
        return self.hwp.HAction.Execute("ShapeCopyPaste", pset.HSet)

    def export_mathml(self, mml_path, delay=0.2):
        """
        MathML 포맷의 수식문서 파일경로를 입력하면

        아래아한글 수식으로 삽입하는 함수
        """

        if self.SelectionMode != 4:
            raise AssertionError("추출할 수식을 먼저 선택해주세요.")
        self.EquationRefresh()
        self.EquationModify(thread=True)

        mml_path = os.path.abspath(mml_path)
        if not os.path.exists(os.path.dirname(mml_path)):
            os.mkdir(os.path.dirname(mml_path))
        if os.path.exists(mml_path):
            os.remove(mml_path)

        hwnd1 = 0
        while not hwnd1:
            hwnd1 = win32gui.FindWindow(None, "수식 편집기")
            win32gui.ShowWindow(hwnd1, win32con.SW_HIDE)
            sleep(delay)  # 제거예정
        # _refresh_eq(hwnd1, delay=delay)
        _open_dialog(hwnd=hwnd1, key="S")
        sleep(delay)

        hwnd2 = 0
        while not hwnd2:
            _open_dialog(hwnd1)
            hwnd2 = win32gui.FindWindow(None, "MathML 형식으로 내보내기")
            win32gui.ShowWindow(hwnd2, win32con.SW_HIDE)
            sleep(delay)

        sleep(delay)
        child_hwnds = []
        win32gui.EnumChildWindows(hwnd2, lambda hwnd, param: param.append(hwnd), child_hwnds)
        for chwnd in child_hwnds:
            class_name = win32gui.GetClassName(chwnd)
            if class_name == "Edit":
                while True:
                    win32gui.SendMessage(chwnd, win32con.WM_SETTEXT, None, mml_path)
                    if _get_edit_text(chwnd) == mml_path:
                        win32api.keybd_event(win32con.VK_EXECUTE, 0, 0, 0)
                        sleep(delay)
                        break
                    sleep(delay)
                break
        self.EquationClose(save=True)
        self.Cancel()

    def import_mathml(self, mml_path, delay=0.2):
        """
        MathML 포맷의 수식문서 파일경로를 입력하면

        아래아한글 수식으로 삽입하는 함수
        """
        if os.path.exists(os.path.abspath(mml_path)):
            mml_path = os.path.abspath(mml_path)
        else:
            raise AssertionError("mathml 파일을 찾을 수 없습니다. 경로를 다시 확인해주세요.")

        self.Cancel()
        self.EquationCreate(thread=True)

        hwnd1 = 0
        while not hwnd1:
            hwnd1 = win32gui.FindWindow(None, "수식 편집기")
        _open_dialog(hwnd=hwnd1, key="M")
        sleep(delay)

        hwnd2 = 0
        while not hwnd2:
            _open_dialog(hwnd1)
            hwnd2 = win32gui.FindWindow(None, "MathML 파일 불러오기")
            win32gui.ShowWindow(hwnd2, win32con.SW_HIDE)
            sleep(delay)

        sleep(delay)
        child_hwnds = []
        win32gui.EnumChildWindows(hwnd2, lambda hwnd, param: param.append(hwnd), child_hwnds)
        for chwnd in child_hwnds:
            class_name = win32gui.GetClassName(chwnd)
            if class_name == "Edit":
                while True:
                    win32gui.SendMessage(chwnd, win32con.WM_SETTEXT, None, mml_path)
                    if _get_edit_text(chwnd) == mml_path:
                        win32api.keybd_event(win32con.VK_EXECUTE, 0, 0, 0)
                        sleep(delay)
                        break
                    sleep(delay)
                break
        self.EquationClose(save=True)
        self.hwp.HAction.Run("SelectCtrlReverse")
        self.EquationRefresh()
        return True

    def maximize_window(self) -> int:
        """현재 창 최대화"""
        win32gui.ShowWindow(
            self.XHwpWindows.Active_XHwpWindow.WindowHandle, 3)

    def minimize_window(self) -> int:
        """현재 창 최소화"""
        win32gui.ShowWindow(
            self.XHwpWindows.Active_XHwpWindow.WindowHandle, 6)

    def delete_style_by_name(self, src: int | str, dst: int | str) -> bool:
        """
        **주의사항**

        매번 메서드를 호출할 때마다 문서를 저장함(구현 편의를 위해ㅜ)!!!
        다소 번거롭더라도 StyleDelete 액션을 직접 실행하는 것을 추천함.

        특정 스타일을 이름 (또는 인덱스번호)로 삭제하고
        대체할 스타일 또한 이름 (또는 인덱스번호)로 지정해주는 메서드.
        """
        style_dict = self.get_style_dict(as_="dict")
        pset = self.HParameterSet.HStyleDelete
        self.HAction.GetDefault("StyleDelete", pset.HSet)
        if type(src) == int:
            pset.Target = src
        elif src in [style_dict[i]["name"] for i in style_dict]:
            pset.Target = [i for i in style_dict if style_dict[i]["name"] == src][0]
        else:
            raise IndexError("해당 스타일이름을 찾을 수 없습니다.")
        if type(dst) == int:
            pset.Alternation = dst
        elif dst in [style_dict[i]["name"] for i in style_dict]:
            pset.Alternation = [i for i in style_dict if style_dict[i]["name"] == dst][0]
        else:
            raise IndexError("해당 스타일이름을 찾을 수 없습니다.")
        return self.HAction.Execute("StyleDelete", pset.HSet)

    # def register_font_ui(self):  # 느리고 불안정한 관계로 도입 보류
    #     """
    #     FontNameComboImpl 요소의 폰트 이름을 추출하는 함수
    #     Returns:
    #        폰트 이름 (str), 찾지 못하면 None 반환
    #     """
    #     from comtypes import CoCreateInstance
    #     from comtypes.gen.UIAutomationClient import CUIAutomation, IUIAutomation
    #
    #     # UI Automation 객체 초기화
    #     uia = CoCreateInstance(
    #         CUIAutomation._reg_clsid_,
    #         interface=IUIAutomation,
    #     )
    #
    #     # 주어진 핸들로부터 시작하는 요소 가져오기
    #     element = uia.ElementFromHandle(self.XHwpWindows.Active_XHwpWindow.WindowHandle)
    #
    #     # FontNameComboImpl 요소 찾기
    #     font_cond = uia.CreatePropertyCondition(30012, "FontNameComboImpl")  # 30012는 UIA_ClassNamePropertyId
    #     self.cur_font_ui = element.FindFirst(4, font_cond)  # 4는 TreeScope.Descendants에 해당

    def get_style_dict(self, as_: Literal["list", "dict"] = "list") -> list|dict:
        """
        스타일 목록을 사전 데이터로 리턴하는 메서드.

        (도움 주신 kosohn님께 아주 큰 감사!!!)
        """
        cur_pos = self.get_pos()
        if not self.MoveSelRight():
            self.MoveSelLeft()
        if not self.save_block_as("temp.xml", format="HWPML2X"):
            self.SelectAll()
            self.save_block_as("temp.xml", format="HWPML2X")
            self.Cancel()
        self.set_pos(*cur_pos)

        tree = ET.parse("temp.xml")
        root = tree.getroot()
        if as_ == "list":
            styles = [
                {
                    'index': int(style.get("Id")),
                    'type': style.get('Type'),
                    'name': style.get('Name'),
                    'engName': style.get('EngName')
                }
                for style in root.findall('.//STYLE')
            ]
        elif as_ == "dict":
            styles = {
                int(style.get('Id')): {
                    'type': style.get('Type'),
                    'name': style.get('Name'),
                    'engName': style.get('EngName')
                }
                for style in root.findall('.//STYLE')
            }
        else:
            raise TypeError("as_ 파라미터는 'list'또는 'dict' 중 하나로 설정해주세요. 기본값은 'list'입니다.")
        os.remove("temp.xml")
        return styles

    def get_style(self) -> dict:
        """
        현재 캐럿이 위치한 문단의 스타일정보를 사전 형태로 리턴한다.

        Returns:
            스타일 목록 딕셔너리
        """
        style_dict = self.get_style_dict(as_="list")
        pset = self.HParameterSet.HStyle
        self.HAction.GetDefault("Style", pset.HSet)
        return style_dict[pset.Apply]

    def set_style(self, style:int|str) -> bool:
        """
        현재 캐럿이 위치한 문단의 스타일을 변경한다.

        스타일 입력은 style 인수로 정수값(스타일번호) 또는 문자열(스타일이름)을 넣으면 된다.

        Args:
            style: 현재 문단에 적용할 스타일 번호(int) 또는 스타일이름(str)

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        pset = self.HParameterSet.HStyle
        if type(style) != int:
            style_dict = self.get_style_dict(as_="dict")
            key = None
            for key, value in style_dict.items():
                if value.get('name') == style:
                    style = key
                    break
                else:
                    continue
            if style != key:
                raise KeyError("해당하는 스타일이 없습니다.")
        self.HAction.GetDefault("Style", pset.HSet)
        pset.Apply = style
        return self.HAction.Execute("Style", pset.HSet)

    def get_selected_range(self) -> list[str]:
        """
        선택한 범위의 셀주소를 리스트로 리턴함

        캐럿이 표 안에 있어야 함
        """
        if not self.is_cell():
            raise AttributeError("캐럿이 표 안에 있어야 합니다.")
        pset = self.HParameterSet.HFieldCtrl
        self.HAction.GetDefault("TableFormula", pset.HSet)
        return pset.Command[2:-1].split(",")

    def fill_addr_field(self):
        """
        현재 표 안에서 모든 셀에 엑셀 셀주소 스타일("A1")의 셀필드를 채우는 메서드

        """
        if not self.is_cell():
            raise AttributeError("캐럿이 표 안에 있어야 합니다.")
        self.TableColBegin()
        self.TableColPageUp()
        self.set_cur_field_name("A1")
        while self.TableRightCell():
            self.set_cur_field_name(self.get_cell_addr())

    def unfill_addr_field(self):
        """
        현재 캐럿이 들어있는 표의 셀필드를 모두 제거하는 메서드

        """
        if not self.is_cell():
            raise AttributeError("캐럿이 표 안에 있어야 합니다.")
        self.TableColBegin()
        self.TableColPageUp()
        self.set_cur_field_name("")
        while self.TableRightCell():
            self.set_cur_field_name("")

    def resize_image(self, width: int = None, height: int = None, unit: Literal["mm", "hwpunit"] = "mm"):
        """
        이미지 또는 그리기 개체의 크기를 조절하는 메서드.

        해당개체 선택 후 실행해야 함.
        """
        self.FindCtrl()
        prop = self.CurSelectedCtrl.Properties
        if width:
            prop.SetItem("Width", width if unit == "hwpunit" else self.MiliToHwpUnit(width))
        if height:
            prop.SetItem("Height", height if unit == "hwpunit" else self.MiliToHwpUnit(height))
        if width or height:
            self.CurSelectedCtrl.Properties = prop
            return True
        return False

    def save_image(self, path: str = "./img.png", ctrl: Any = None):
        path = os.path.abspath(path)
        if os.path.exists(path):
            os.remove(path)
        if ctrl:
            self.select_ctrl(ctrl)
        else:
            self.find_ctrl()
        if not self.CurSelectedCtrl.CtrlID == "gso":
            return False
        pset = self.HParameterSet.HShapeObjSaveAsPicture
        self.HAction.GetDefault("PictureSave", pset.HSet)
        pset.Path = path
        pset.Ext = "BMP"
        try:
            return self.HAction.Execute("PictureSave", pset.HSet)
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

    def NewNumberModify(self, new_number: int, num_type: Literal["Page", "Figure", "Footnote", "Table", "Endnote", "Equation"] = "Page") -> bool:
        """
        새 번호 조판을 수정할 수 있는 메서드.

        실행 전 [새 번호] 조판 옆에 캐럿이 위치해 있어야 하며,
        그렇지 않을 경우
        (쪽번호 외에도 그림, 각주, 표, 미주, 수식 등)
        다만, 주의할 점이 세 가지 있다.

            1. 기존에 쪽번호가 없는 문서에서는 작동하지 않으므로 쪽번호가 정의되어 있어야 한다. (쪽번호 정의는 `PageNumPos` 메서드 참조)
            2. 새 번호를 지정한 페이지 및 이후 모든 페이지가 영향을 받는다.
            3. `NewNumber` 실행시점의 캐럿위치 뒤쪽(해당 페이지 내)에 `NewNumber` 조판이 있는 경우, 삽입한 조판은 무효가 된다. (페이지 맨 뒤쪽의 새 번호만 유효함)

        Args:
            new_number: 새 번호
            num_type:
                타입 지정

                    - "Page": 쪽(기본값)
                    - "Figure": 그림
                    - "Footnote": 각주
                    - "Table": 표
                    - "Endnote": 미주
                    - "Equation": 수식

        Returns:
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

    def NewNumber(self, new_number: int, num_type: Literal["Page", "Figure", "Footnote", "Table", "Endnote", "Equation"] = "Page") -> bool:
        """
        새 번호를 매길 수 있는 메서드.

        (쪽번호 외에도 그림, 각주, 표, 미주, 수식 등)
        다만, 주의할 점이 세 가지 있다.
        1. 기존에 쪽번호가 없는 문서에서는 작동하지 않으므로
           쪽번호가 정의되어 있어야 한다.
           (쪽번호 정의는 PageNumPos 메서드 참조)
        2. 새 번호를 지정한 페이지 및 이후 모든 페이지가 영향을 받는다.
        3. NewNumber 실행시점의 캐럿위치 뒤쪽(해당 페이지 내)에
           NewNumber 조판이 있는 경우, 삽입한 조판은 무효가 된다.
           (페이지 맨 뒤쪽의 새 번호만 유효함)

        Args:
            new_number: 새 번호
            num_type: 타입 지정

                - "Page": 쪽(기본값)
                - "Figure": 그림
                - "Footnote": 각주
                - "Table": 표
                - "Endnote": 미주
                - "Equation": 수식

        Returns:
            성공시 True, 실패시 False를 리턴

        Examples:
            >>> # 쪽번호가 있는 문서에서
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.NewNumber(5)  # 현재 페이지번호가 5로 바뀜
        """
        pset = self.HParameterSet.HAutoNum
        self.HAction.GetDefault("NewNumber", pset.HSet)
        pset.NumType = self.AutoNumType(num_type)
        pset.NewNumber = new_number
        return self.HAction.Execute("NewNumber", pset.HSet)

    def PageNumPos(self, global_start: int = 1, position: Literal["TopLeft", "TopCenter", "TopRight", "BottomLeft", "BottomCenter", "BottomRight", "InsideTop", "OutsideTop", "InsideBottom", "OutsideBottom", "None"] = "BottomCenter",
                   number_format: Literal["Digit", "CircledDigit", "RomanCapital", "RomanSmall", "LatinCapital", "HangulSyllable", "Ideograph", "DecagonCircle", "DecagonCircleHanja"] = "Digit",
                   side_char:bool=True) -> bool:
        """
        문서 전체에 쪽번호를 삽입하는 메서드.

        Args:
            global_start: 시작번호를 지정할 수 있음(새 번호 아님. 새 번호는 hwp.NewNumber(n)을 사용할 것)
            position:
                쪽번호 위치를 지정하는 파라미터

                    - TopLeft
                    - TopCenter
                    - TopRight
                    - BottomLeft
                    - BottomCenter(기본값)
                    - BottomRight
                    - InsideTop
                    - OutsideTop
                    - InsideBottom
                    - OutsideBottom
                    - None(쪽번호숨김과 유사)

            number_format:
                쪽번호 서식을 지정하는 파라미터

                    - "Digit": (1 2 3),
                    - "CircledDigit": (① ② ③),
                    - "RomanCapital":(I II III),
                    - "RomanSmall": (i ii iii) ,
                    - "LatinCapital": (A B C),
                    - "HangulSyllable": (가 나 다),
                    - "Ideograph": (一 二 三),
                    - "DecagonCircle": (갑 을 병),
                    - "DecagonCircleHanja": (甲 乙 丙),
            side_char:
                줄표 삽입 여부(bool)

                    - True : 줄표 삽입(기본값)
                    - False : 줄표 삽입하지 않음

        Returns:
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

    def set_cell_margin(self, left: float = 1.8, right: float = 1.8, top: float = 0.5, bottom: float = 0.5, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        표 중 커서가 위치한 셀 또는 다중선택한 모든 셀의 안 여백을 지정하는 메서드.

        표 안에서만 실행가능하며, 전체 셀이 아닌 표 자체를 선택한 상태에서는 여백이 적용되지 않음.
        차례대로 왼쪽, 오른쪽, 상단, 하단의 여백을 밀리미터(기본값) 또는 HwpUnit 단위로 지정.

        Args:
            left: 셀의 좌측 안여백
            right: 셀의 우측 안여백
            top: 셀의 상단 안여백
            bottom: 셀의 하단 안여백
            as_: 입력단위. ["mm", "hwpunit"] 중 기본값은 "mm"

        Returns:
            성공시 True, 실패시 False를 리턴함
        """
        if not self.is_cell():
            return False
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.HSet.SetItem("ShapeType", 3)
        pset.HSet.SetItem("ShapeCellSize", 0)
        pset.ShapeTableCell.HasMargin = 1
        if as_ == "mm":
            pset.ShapeTableCell.MarginLeft = self.hwp.MiliToHwpUnit(left)
            pset.ShapeTableCell.MarginRight = self.hwp.MiliToHwpUnit(right)
            pset.ShapeTableCell.MarginTop = self.hwp.MiliToHwpUnit(top)
            pset.ShapeTableCell.MarginBottom = self.hwp.MiliToHwpUnit(bottom)
        elif as_.lower() == "hwpunit":
            pset.ShapeTableCell.MarginLeft = left
            pset.ShapeTableCell.MarginRight = right
            pset.ShapeTableCell.MarginTop = top
            pset.ShapeTableCell.MarginBottom = bottom
        return self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

    def get_cell_margin(self, as_: Literal["mm", "hwpunit"] = "mm") -> None | dict[str, int] | bool | dict[str, float]:
        """
        표 중 커서가 위치한 셀 또는 다중선택한 모든 셀의 안 여백을 조회하는 메서드.

        표 안에서만 실행가능하며, 전체 셀이 아닌 표 자체를 선택한 상태에서는 여백이 조회되지 않음.

        Args:
            as_: 리턴값의 단위("mm" 또는 "hwpunit" 중 지정가능. 기본값은 "mm")

        Returns:
            모든 셀의 안여백을 dict로 리턴. 표 안에 있지 않으면 False를 리턴
        """
        if not self.is_cell():
            return False
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        if as_ == "mm":
            return {
                "left": self.hwp_unit_to_mili(pset.ShapeTableCell.MarginLeft),
                "right": self.hwp_unit_to_mili(pset.ShapeTableCell.MarginRight),
                "top": self.hwp_unit_to_mili(pset.ShapeTableCell.MarginTop),
                "bottom": self.hwp_unit_to_mili(pset.ShapeTableCell.MarginBottom),
            }
        elif as_.lower() == "hwpunit":
            return {
                "left": pset.ShapeTableCell.MarginLeft,
                "right": pset.ShapeTableCell.MarginRight,
                "top": pset.ShapeTableCell.MarginTop,
                "bottom": pset.ShapeTableCell.MarginBottom,
            }

    def set_table_inside_margin(self, left: float = 1.8, right: float = 1.8, top: float = 0.5, bottom: float = 0.5, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        표 내부 모든 셀의 안여백을 일괄설정하는 메서드.

        표 전체를 선택하지 않고 표 내부에 커서가 있기만 하면 모든 셀에 적용됨.

        Args:
            left: 모든 셀의 좌측여백(mm)
            right: 모든 셀의 우측여백(mm)
            top: 모든 셀의 상단여백(mm)
            bottom: 모든 셀의 하단여백(mm)
            as_: 입력단위. "mm", "hwpunit" 중 택일. 기본값은 "mm"

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        if not self.is_cell():
            return False
        if as_.lower() == "mm":
            left, right, top, bottom = [self.hwp.MiliToHwpUnit(i) for i in [left, right, top, bottom]]
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.CellMarginLeft = left
        pset.CellMarginRight = right
        pset.CellMarginTop = top
        pset.CellMarginBottom = bottom
        return self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

    def get_table_inside_margin(self, as_: Literal["mm", "hwpunit"] = "mm") -> None | dict[str, int] | bool | dict[str, float]:
        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        prop = self.CurSelectedCtrl.Properties
        margin_left = prop.Item("CellMarginLeft")
        margin_right = prop.Item("CellMarginRight")
        margin_top = prop.Item("CellMarginTop")
        margin_bottom = prop.Item("CellMarginBottom")
        self.set_pos(*cur_pos)
        if as_ == "mm":
            return {
                "left": self.hwp_unit_to_mili(margin_left),
                "right": self.hwp_unit_to_mili(margin_right),
                "top": self.hwp_unit_to_mili(margin_top),
                "bottom": self.hwp_unit_to_mili(margin_bottom),
            }
        elif as_.lower() == "hwpunit":
            return {
                "left": margin_left,
                "right": margin_right,
                "top": margin_top,
                "bottom": margin_bottom,
            }

    def get_table_outside_margin(self, as_: Literal["mm", "hwpunit"] = "mm") -> None | dict[str, int] | bool | dict[str, float]:
        """
        표의 바깥 여백을 딕셔너리로 한 번에 리턴하는 메서드

        Args:
            as_:
                리턴하는 여백값의 단위

                    - "mm": 밀리미터(기본값)
                    - "hwpunit": HwpUnit

        Returns:
            표의 상하좌우 바깥여백값을 담은 딕셔너리. 표 안에서 실행하지 않은 경우에는 False를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_table_outside_margin()
            {'left': 4.0, 'right': 3.0, 'top': 2.0, 'bottom': 1.0}

        """
        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        self.TableCellBlock()
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        margin_left = pset.OutsideMarginLeft
        margin_right = pset.OutsideMarginRight
        margin_top = pset.OutsideMarginTop
        margin_bottom = pset.OutsideMarginBottom
        self.set_pos(*cur_pos)
        if as_ == "mm":
            return {
                "left": round(self.hwp_unit_to_mili(margin_left), 2),
                "right": round(self.hwp_unit_to_mili(margin_right), 2),
                "top": round(self.hwp_unit_to_mili(margin_top), 2),
                "bottom": round(self.hwp_unit_to_mili(margin_bottom), 2),
            }
        elif as_.lower() == "hwpunit":
            return {
                "left": margin_left,
                "right": margin_right,
                "top": margin_top,
                "bottom": margin_bottom,
            }

    def get_table_outside_margin_left(self, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        표의 바깥 왼쪽 여백값을 리턴하는 메서드

        Args:
            as_:
                리턴하는 여백값의 단위

                    - "mm": 밀리미터(기본값)
                    - "hwpunit": HwpUnit

        Returns:
            표의 좌측 바깥여백값. 단위에 따라 int|float을 리턴하며, 표 안에서 실행하지 않은 경우에는 False를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_table_outside_margin_left()
            4.0

        """
        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        self.TableCellBlock()
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        margin = pset.OutsideMarginLeft
        self.set_pos(*cur_pos)
        return round(self.hwp_unit_to_mili(margin), 2) if as_ == "mm" else margin

    def get_table_outside_margin_right(self, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        표의 바깥 오른쪽 여백값을 리턴하는 메서드

        Args:
            as_:
                리턴하는 여백값의 단위

                    - "mm": 밀리미터(기본값)
                    - "hwpunit": HwpUnit

        Returns:
            표의 우측 바깥여백값. 단위에 따라 int|float을 리턴하며, 표 안에서 실행하지 않은 경우에는 False를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_table_outside_margin_left()
            3.0

        """

        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        self.TableCellBlock()
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        margin = pset.OutsideMarginRight
        self.set_pos(*cur_pos)
        return round(self.hwp_unit_to_mili(margin), 2) if as_ == "mm" else margin

    def get_table_outside_margin_top(self, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        표의 바깥 상단 여백값을 리턴하는 메서드

        Args:
            as_:
                리턴하는 여백값의 단위

                    - "mm": 밀리미터(기본값)
                    - "hwpunit": HwpUnit

        Returns:
            표의 위쪽 바깥여백값. 단위에 따라 int|float을 리턴하며, 표 안에서 실행하지 않은 경우에는 False를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_table_outside_margin_top()
            2.0

        """
        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        self.TableCellBlock()
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        margin = pset.OutsideMarginTop
        self.set_pos(*cur_pos)
        return round(self.hwp_unit_to_mili(margin), 2) if as_ == "mm" else margin

    def get_table_outside_margin_bottom(self, as_: Literal["mm", "hwpunit"] = "mm") -> int|float|bool:
        """
        표의 바깥 하단 여백값을 리턴하는 메서드

        Args:
            as_:
                리턴하는 여백값의 단위

                    - "mm": 밀리미터(기본값)
                    - "hwpunit": HwpUnit

        Returns:
            표의 아랫쪽 바깥여백값. 단위에 따라 int|float을 리턴하며, 표 안에서 실행하지 않은 경우에는 False를 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_table_outside_margin_bottom()
            1.0

        """
        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        self.TableCellBlock()
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        margin = pset.OutsideMarginBottom
        self.set_pos(*cur_pos)
        return round(self.hwp_unit_to_mili(margin), 2) if as_ == "mm" else margin

    def set_table_outside_margin(self, left: float=-1.0, right: float=-1.0, top: float=-1.0, bottom: float=-1.0, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        표의 바깥여백을 변경하는 메서드.

        기본 입력단위는 "mm"이며, "HwpUnit" 단위로 변경 가능.

        Args:
            left: 표의 좌측 바깥여백
            right: 표의 우측 바깥여백
            top: 표의 상단 바깥여백
            bottom: 표의 하단 바깥여백
            as_: 입력단위. ["mm", "hwpunit"] 중 기본값은 "mm"

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        if not self.is_cell():
            return False
        cur_pos = self.get_pos()
        pset = self.HParameterSet.HShapeObject
        self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        if as_ == "mm":
            if left >= 0:
                pset.OutsideMarginLeft = self.MiliToHwpUnit(left)
            if right >= 0:
                pset.OutsideMarginRight = self.MiliToHwpUnit(right)
            if top >= 0:
                pset.OutsideMarginTop = self.MiliToHwpUnit(top)
            if bottom >= 0:
                pset.OutsideMarginBottom = self.MiliToHwpUnit(bottom)
        elif as_.lower() == "hwpunit":
            if left >= 0:
                pset.OutsideMarginLeft = left
            if right >= 0:
                pset.OutsideMarginRight = right
            if top >= 0:
                pset.OutsideMarginTop = top
            if bottom >= 0:
                pset.OutsideMarginBottom = bottom
        try:
            return self.HAction.Execute("TablePropertyDialog", pset.HSet)
        finally:
            self.set_pos(*cur_pos)

    def set_table_outside_margin_left(self, val, as_: Literal["mm", "hwpunit"] = "mm"):
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        prop = self.CurSelectedCtrl.Properties
        if as_ == "mm":
            val = self.mili_to_hwp_unit(val)
        prop.SetItem("OutsideMarginLeft", val)
        self.CurSelectedCtrl.Properties = prop
        return self.set_pos(*cur_pos)

    def set_table_outside_margin_right(self, val, as_: Literal["mm", "hwpunit"] = "mm"):
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        prop = self.CurSelectedCtrl.Properties
        if as_ == "mm":
            val = self.mili_to_hwp_unit(val)
        prop.SetItem("OutsideMarginRight", val)
        self.CurSelectedCtrl.Properties = prop
        return self.set_pos(*cur_pos)

    def set_table_outside_margin_top(self, val, as_: Literal["mm", "hwpunit"] = "mm"):
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        prop = self.CurSelectedCtrl.Properties
        if as_ == "mm":
            val = self.mili_to_hwp_unit(val)
        prop.SetItem("OutsideMarginTop", val)
        self.CurSelectedCtrl.Properties = prop
        return self.set_pos(*cur_pos)

    def set_table_outside_margin_bottom(self, val, as_: Literal["mm", "hwpunit"] = "mm"):
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        prop = self.CurSelectedCtrl.Properties
        if as_ == "mm":
            val = self.mili_to_hwp_unit(val)
        prop.SetItem("OutsideMarginBottom", val)
        self.CurSelectedCtrl.Properties = prop
        return self.set_pos(*cur_pos)

    def get_table_height(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"):
        """
        현재 캐럿이 속한 표의 높이(mm)를 리턴함

        Returns:
        표의 높이(mm)
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

        현재 표의 행의 최대갯수를 리턴

        Returns:
        최대 행갯수:int
        """
        if not self.is_cell():
            raise AssertionError("현재 캐럿이 표 안에 있지 않습니다.")
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        t = self.GetTextFile("HWPML2X", "saveblock")
        root = ET.fromstring(t)
        table = root.find('.//TABLE')
        row_count = int(table.get('RowCount'))
        self.set_pos(*cur_pos)
        return row_count

    def get_row_height(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm") -> float|int:
        """
        표 안에서 캐럿이 들어있는 행(row)의 높이를 리턴함.

        기본단위는 mm 이지만, HwpUnit이나 Point 등 보다 작은 단위를 사용할 수 있다. (메서드 내부에서는 HwpUnit으로 연산한다.)

        Args:
            as_: 리턴하는 수치의 단위

        Returns:
            캐럿이 속한 행의 높이
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

        Returns:
        """
        if not self.is_cell():
            raise AssertionError("현재 캐럿이 표 안에 있지 않습니다.")
        cur_pos = self.get_pos()
        self.SelectCtrlFront()
        t = self.GetTextFile("HWPML2X", "saveblock")
        root = ET.fromstring(t)
        table = root.find('.//TABLE')
        col_count = int(table.get('ColCount'))
        self.set_pos(*cur_pos)
        return col_count

    def get_col_width(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm") -> int|float:
        """
        현재 캐럿이 위치한 셀(칼럼)의 너비를 리턴하는 메서드.

        기본 단위는 mm이지만, as_ 파라미터를 사용하여 단위를 hwpunit이나 point, inch 등으로 변경 가능하다.

        Args:
            as_: 리턴값의 단위(mm, HwpUnit, Pt, Inch 등 4종류)

        Returns:
            현재 칼럼의 너비(기본단위는 mm)
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

    def set_col_width(self, width: int | float | list | tuple, as_: Literal["mm", "ratio"] = "ratio") -> bool:
        """
        칼럼의 너비를 변경하는 메서드.

        정수(int)나 부동소수점수(float) 입력시 현재 칼럼의 너비가 변경되며,
        리스트나 튜플 등 iterable 타입 입력시에는 각 요소들의 비에 따라 칼럼들의 너비가 일괄변경된다.
        예를 들어 3행 3열의 표 안에서 set_col_width([1,2,3]) 을 실행하는 경우
        1열너비:2열너비:3열너비가 1:2:3으로 변경된다.
        (표 전체의 너비가 148mm라면, 각각 24mm : 48mm : 72mm로 변경된다는 뜻이다.)

        단, 열너비의 비가 아닌 "mm" 단위로 값을 입력하려면 as_="mm"로 파라미터를 수정하면 된다.
        이 때, width에 정수 또는 부동소수점수를 입력하는 경우 as_="ratio"를 사용할 수 없다.

        Args:
            width: 열 너비
            as_: 단위

        Returns:
            성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_table(3,3)
            True
            >>> hwp.get_into_nth_table(0)
            True
            >>> hwp.set_col_width([1,2,3])
            True
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

    def adjust_cellwidth(self, width: int | float | list | tuple, as_: Literal["mm", "ratio"] = "ratio") -> bool:
        """
        칼럼의 너비를 변경할 수 있는 메서드.

        정수(int)나 부동소수점수(float) 입력시 현재 칼럼의 너비가 변경되며,
        리스트나 튜플 등 iterable 타입 입력시에는 각 요소들의 비에 따라 칼럼들의 너비가 일괄변경된다.
        예를 들어 3행 3열의 표 안에서 set_col_width([1,2,3]) 을 실행하는 경우
        1열너비:2열너비:3열너비가 1:2:3으로 변경된다.
        (표 전체의 너비가 148mm라면, 각각 24mm : 48mm : 72mm로 변경된다는 뜻이다.)

        단, 열너비의 비가 아닌 "mm" 단위로 값을 입력하려면 as_="mm"로 파라미터를 수정하면 된다.
        이 때, width에 정수 또는 부동소수점수를 입력하는 경우 as_="ratio"를 사용할 수 없다.

        Args:
            width: 열 너비
            as_: 단위

        Returns:
            성공시 True

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_table(3,3)
            >>> hwp.get_into_nth_table(0)
            >>> hwp.adjust_cellwidth([1,2,3])
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

    def get_table_width(self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm") -> float:
        """
        현재 캐럿이 속한 표의 너비(mm)를 리턴함.

        이 때 수치의 단위는 as_ 파라미터를 통해 변경 가능하며, "mm", "HwpUnit", "Pt", "Inch" 등을 쓸 수 있다.

        Returns:
            표의 너비(mm)
        """
        if not self.is_cell():
            raise IndexError("캐럿이 표 안에 있어야 합니다.")
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

    def set_table_width(self, width: int = 0, as_: Literal["mm", "hwpunit", "hu"] = "mm") -> bool:
        """
        표 전체의 너비를 원래 열들의 비율을 유지하면서 조정하는 메서드.

        내부적으로 xml 파싱을 사용하는 방식으로 변경.

        Args:
            width: 너비(단위는 기본 mm이며, hwpunit으로 변경 가능)
            as_: 단위("mm" or "hwpunit")

        Returns:
            성공시 True

        Examples:
            >>> # 모든 표의 너비를 본문여백(용지너비 - 좌측여백 - 우측여백 - 제본여백 - 표 좌우 바깥여백)에 맞추기
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> i = 0
            >>> while hwp.get_into_nth_table(i):
            ...     hwp.set_table_width()
            True

        """
        if not width:
            sec_def = self.hwp.HParameterSet.HSecDef
            self.hwp.HAction.GetDefault("PageSetup", sec_def.HSet)
            if sec_def.PageDef.Landscape == 0:
                width = (
                        sec_def.PageDef.PaperWidth - sec_def.PageDef.LeftMargin - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen
                        - self.get_table_outside_margin_left(as_="hwpunit") - self.get_table_outside_margin_right(
                    as_="hwpunit"))
            elif sec_def.PageDef.Landscape == 1:
                width = (
                        sec_def.PageDef.PaperHeight - sec_def.PageDef.LeftMargin - sec_def.PageDef.RightMargin - sec_def.PageDef.GutterLen
                        - self.get_table_outside_margin_left(as_="hwpunit") - self.get_table_outside_margin_right(
                    as_="hwpunit"))
        elif as_ == "mm":
            width = self.mili_to_hwp_unit(width)
        ratio = width / self.get_table_width(as_="hwpunit")
        cur_pos = self.get_pos()
        while True:
            self.SelectCtrlFront()
            ctrl = self.CurSelectedCtrl
            if ctrl.UserDesc == "표":
                break
        t = self.GetTextFile("HWPML2X", "saveblock")
        root = ET.fromstring(t)
        table = root.find('.//TABLE')

        if table is not None:
            for cell in table.findall('.//CELL'):
                width = cell.get('Width')
                if width:
                    cell.set('Width', str(int(width) * ratio))
        t = ET.tostring(root, encoding='UTF-16').decode('utf-16')
        cur_view_state = self.ViewProperties.Item("OptionFlag")
        if cur_view_state not in (2, 6):
            prop = self.ViewProperties
            prop.SetItem("OptionFlag", 6)
            self.ViewProperties = prop
        self.move_to_ctrl(ctrl)
        self.MoveSelRight()
        self.HAction.Run("Delete")
        self.SetTextFile(t, format="HWPML2X", option="insertfile")
        prop = self.ViewProperties
        prop.SetItem("OptionFlag", cur_view_state)
        self.ViewProperties = prop
        self.set_pos(*cur_pos)

    def save_pdf_as_image(self, path: str = "", img_format:str="bmp") -> bool:
        """
        문서보안이나 복제방지를 위해 모든 페이지를 이미지로 변경 후 PDF로 저장하는 메서드.

        아무 인수가 주어지지 않는 경우
        모든 페이지를 bmp로 저장한 후에
        현재 폴더에 {문서이름}.pdf로 저장한다.
        (만약 저장하지 않은 빈 문서의 경우에는 result.pdf로 저장한다.)

        Args:
            path: 저장경로 및 파일명
            img_format: 이미지 변환 포맷

        Returns:
            성공시 True
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

    def get_cell_addr(self, as_: Literal["str", "tuple"] = "str") -> tuple[int]|bool:
        """
        현재 캐럿이 위치한 셀의 주소를 "A1" 또는 (0, 0)으로 리턴.

        캐럿이 표 안에 있지 않은 경우 False를 리턴함

        Args:
            as_: `"str"`의 경우 엑셀처럼 `"A1"` 방식으로 리턴, `"tuple"`인 경우 (0,0) 방식으로 리턴.

        Returns:
        """
        if not self.hwp.CellShape:
            return False
        result = self.KeyIndicator()[-1][1:].split(")")[0]
        if as_ == "str":
            return result
        else:
            return excel_address_to_tuple_zero_based(result)

    def save_all_pictures(self, save_path:str="./binData") -> bool:
        """
        현재 문서에 삽입된 모든 이미지들을

        삽입 당시 파일명으로 복원하여 저장.
        단, 문서 안에서 복사했거나 중복삽입한 이미지는 한 개만 저장됨.
        기본 저장폴더명은 ./binData이며
        기존에 save_path가 존재하는 경우,
        그 안의 파일들은 삭제되므로 유의해야 함.

        Args:
            save_path: 저장할 하위경로 이름

        Returns:
            bool: 성공시 True
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

    def SelectCtrl(self, ctrllist: str | int, option: Literal[0, 1] = 1) -> bool:
        """
        한글2024 이상의 버전에서 사용 가능한 API 기반의 신규 메서드.

        가급적 hwp.select_ctrl(ctrl)을 실행할 것을 추천.

        Args:
            ctrllist:
                특정 컨트롤의 인스턴스 아이디(11자리 정수 또는 문자열).
                인스턴스아이디는 `ctrl.GetCtrlInstID()` 로 구할 수 있으며
                이는 한/글 2024부터 반영된 개념(2022 이하에서는 제공하지 않음)
            option:
                특정 컨트롤(들)을 선택하고 있는 상태에서, 추가선택할 수 있는 옵션.

                    - 0: 추가선택
                    - 1: 기존 선택해제 후 컨트롤 선택

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 한글2024 이상의 버전
            >>> # 문서 마지막 컨트롤 선택하기
            >>> hwp.SelectCtrl(hwp.LastCtrl.GetCtrlInstID(), 0)

        """
        if int(self.Version[0]) >= 13:  # 한/글2024 이상이면
            return self.hwp.SelectCtrl(ctrllist=ctrllist, option=option)
        else:
            raise NotImplementedError("아래아한글 버전이 2024 미만입니다. hwp.select_ctrl()을 대신 사용하셔야 합니다.")

    def select_ctrl(self, ctrl:Ctrl, anchor_type: Literal[0, 1, 2] = 0, option: int = 1) -> bool:
        """
        인수로 넣은 컨트롤 오브젝트를 선택하는 pyhwpx 전용 메서드.

        Args:
            ctrl: 선택하고자 하는 컨트롤
            anchor_type:
                컨트롤의 위치를 찾아갈 때 List, Para, Pos의 기준위치. (아주 특수한 경우를 제외하면 기본값을 쓰면 된다.)

                    - 0: 바로 상위 리스트에서의 좌표(기본값)
                    - 1: 탑레벨 리스트에서의 좌표
                    - 2: 루트 리스트에서의 좌표

        Returns:
            성공시 True
        """
        if int(self.Version[0]) >= 13:  # 한/글2024 이상이면
            self.hwp.SelectCtrl(ctrl.GetCtrlInstID(), option=option)
        else:  # 이하 버전이면
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

    def move_to_ctrl(self, ctrl: Any, option: Literal[0, 1, 2] = 0) -> bool:
        """
        메서드에 넣은 ctrl의 조판부호 앞으로 이동하는 메서드.

        Args:
            ctrl: 이동하고자 하는 컨트롤

        Returns:
            성공시 True, 실패시 False 리턴
        """
        return self.set_pos_by_set(ctrl.GetAnchorPos(option))

    def set_visible(self, visible:bool) -> None:
        """
        현재 조작중인 한/글 인스턴스의 백그라운드 숨김여부를 변경할 수 있다.

        Args:
            visible: `visible=False`로 설정하면 현재 조작중인 한/글 인스턴스가 백그라운드로 숨겨진다.

        Returns:
            None
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
                 Bold:str|bool="",  # 진하게(True/False)
                 DiacSymMark:str|int="",  # 강조점(0~12)
                 Emboss:str|bool="",  # 양각(True/False)
                 Engrave:str|bool="",  # 음각(True/False)
                 FaceName:str="",  # 서체
                 FontType:int=1,  # 1(TTF), 2(HTF)
                 Height:str|float="",  # 글자크기(pt, 0.1 ~ 4096)
                 Italic:str|bool="",  # 이탤릭(True/False)
                 Offset:str|int="",  # 글자위치-상하오프셋(-100 ~ 100)
                 OutLineType:str|int="",  # 외곽선타입(0~6)
                 Ratio:str|int="",  # 장평(50~200)
                 ShadeColor:str|int="",  # 음영색(RGB, 0x000000 ~ 0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
                 ShadowColor:str|int="",  # 그림자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
                 ShadowOffsetX:str|int="",  # 그림자 X오프셋(-100 ~ 100)
                 ShadowOffsetY:str|int="",  # 그림자 Y오프셋(-100 ~ 100)
                 ShadowType:str|int="",  # 그림자 유형(0: 없음, 1: 비연속, 2:연속)
                 Size:str|int="",  # 글자크기 축소확대%(10~250)
                 SmallCaps:str|bool="",  # 강조점
                 Spacing:str|int="",  # 자간(-50 ~ 50)
                 StrikeOutColor:str|int="",
                 # 취소선 색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
                 StrikeOutShape:str|int="",  # 취소선 모양(0~12, 0이 일반 취소선)
                 StrikeOutType:str|bool="",  # 취소선 유무(True/False)
                 SubScript:str|bool="",  # 아래첨자(True/False)
                 SuperScript:str|bool="",  # 위첨자(True/False)
                 TextColor:str|int="",  # 글자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
                 UnderlineColor:str|int="",  # 밑줄색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
                 UnderlineShape:str|int="",  # 밑줄형태(0~12)
                 UnderlineType:str|int="",  # 밑줄위치(0:없음, 1:하단, 3:상단)
                 UseFontSpace:str|bool="",  # 글꼴에 어울리는 빈칸(True/False)
                 UseKerning:str|bool=""  # 커닝 적용(True/False) : 차이가 없다?
                 ) -> bool:
        """
        글자모양을 메서드 형태로 수정할 수 있는 메서드.

        Args:
            Bold: 진하게(True/False)
            DiacSymMark: 1행 선택강조점(0~12)
            Emboss: hwp.TableCellBlockRow()양각(True/False)
            Engrave: hwp.set_font("D2Coding")음각(True/False)
            FaceName: 서체
            FontType: 1(TTF) 고정
            Height: 글자크기(pt, 0.1 ~ 4096)
            Italic: 이탤릭(True/False)
            Offset: 글자위치-상하오프셋(-100 ~ 100)
            OutLineType: 외곽선타입(0~6)
            Ratio: 장평(50~200)
            ShadeColor: 음영색(RGB, 0x000000 ~ 0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
            ShadowColor: 그림자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
            ShadowOffsetX: 그림자 X오프셋(-100 ~ 100)
            ShadowOffsetY: 그림자 Y오프셋(-100 ~ 100)
            ShadowType: 그림자 유형(0: 없음, 1: 비연속, 2:연속)
            Size: 글자크기 축소확대%(10~250)
            SmallCaps: 강조점
            Spacing: 자간(-50 ~ 50)
            StrikeOutColor: 취소선 색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
            StrikeOutShape: 취소선 모양(0~12, 0이 일반 취소선)
            StrikeOutType: 취소선 유무(True/False)
            SubScript: 아래첨자(True/False)
            SuperScript: 위첨자(True/False)
            TextColor: 글자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
            UnderlineColor: 밑줄색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
            UnderlineShape: 밑줄형태(0~12)
            UnderlineType: 밑줄위치(0:없음, 1:하단, 3:상단)
            UseFontSpace: 글꼴에 어울리는 빈칸(True/False) : 차이가 나는 폰트를 못 찾았다...
            UseKerning: 커닝 적용(True/False) : 차이가 전혀 없다?

        Returns:
            성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.SelectAll()  # 전체선택
            >>> hwp.set_font(FaceName="D2Coding", TextColor="Orange")

        """
        d = {'Bold': Bold, 'DiacSymMark': DiacSymMark, 'Emboss': Emboss, 'Engrave': Engrave,
             "FaceNameUser": FaceName, "FaceNameSymbol": FaceName, "FaceNameOther": FaceName,
             "FaceNameJapanese": FaceName, "FaceNameHanja": FaceName, "FaceNameLatin": FaceName,
             "FaceNameHangul": FaceName,
             "FontTypeUser": 1, "FontTypeSymbol": 1, "FontTypeOther": 1, "FontTypeJapanese": 1, "FontTypeHanja": 1,
             "FontTypeLatin": 1, "FontTypeHangul": 1,
             'Height': Height * 100, 'Italic': Italic, 'OffsetHangul': Offset, 'OffsetHanja': Offset,
             'OffsetJapanese': Offset, 'OffsetLatin': Offset, 'OffsetOther': Offset,
             'OffsetSymbol': Offset, 'OffsetUser': Offset,
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

        if FaceName in self.htf_fonts.keys():
            d |= self.htf_fonts[FaceName]

        pset = self.hwp.HParameterSet.HCharShape
        self.HAction.GetDefault("CharShape", pset.HSet)
        for key in d.keys():
            if d[key] != "":
                pset.__setattr__(key, d[key])
        return self.hwp.HAction.Execute("CharShape", pset.HSet)

    def cell_fill(self, face_color: tuple[int, int, int] = (217, 217, 217)):
        """
        선택한 셀에 색 채우기

        Args:
            face_color: (red, green, blue) 형태의 튜플. 각 정수는 0~255까지이며, 만약

        Returns:
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

        Returns:
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

    def get_into_nth_table(self, n=0, select_cell=False):
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
            if select_cell:
                self.ShapeObjTableSelCell()
            else:
                self.ShapeObjTextBoxEdit()
            return ctrl

        while ctrl:
            if ctrl.UserDesc == "표":
                if n in (0, -1):
                    self.set_pos_by_set(ctrl.GetAnchorPos(0))
                    self.hwp.FindCtrl()
                    self.ShapeObjTableSelCell()
                    if not select_cell:
                        self.Cancel()
                    return ctrl
                else:
                    if idx == n:
                        self.set_pos_by_set(ctrl.GetAnchorPos(0))
                        self.hwp.FindCtrl()
                        self.ShapeObjTableSelCell()
                        if not select_cell:
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

    def set_row_height(self, height: int | float, as_: Literal["mm", "hwpunit"] = "mm") -> bool:
        """
        캐럿이 표 안에 있는 경우

        캐럿이 위치한 행의 셀 높이를 조절하는 메서드(기본단위는 mm)

        Args:
            height: 현재 행의 높이 설정값(기본단위는 mm)

        Returns:
            성공시 True, 실패시 False 리턴
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

        Returns:
        """
        self.insert_background_picture("", border_type="SelectedCellDelete")

    def gradation_on_cell(self, color_list: list[tuple[int, int, int]]|list[str, str] = [(0, 0, 0), (255, 255, 255)],
                          grad_type: Literal["Linear", "Radial", "Conical", "Square"] = "Linear", angle:int=0, xc:int=0, yc:int=0,
                          pos_list: list[int] = None, step_center:int=50, step:int=255) -> bool:
        """
        셀에 그라데이션을 적용하는 메서드

        ![gradation_on_cell](assets/gradation_on_cell.gif){ loading:lazy }

        Args:
            color_list: 시작RGB 튜플과 종료RGB 튜플의 리스트
            grad_type: 그라데이션 형태(선형, 방사형, 콘형, 사각형)
            angle: 그라데이션 각도
            xc:  x 중심점
            yc:  y 중심점
            pos_list:  변곡점 목록
            step_center:  그라데이션 단계의 중심점
            step:  그라데이션 단계 수

        Returns:
            성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> for i in range(0, 256, 2):
            ...     hwp.gradation_on_cell(
            ...         color_list=[(255,i,i), (i,255,i)],
            ...         grad_type="Square",
            ...         xc=40, yc=60,
            ...         pos_list=[20,80],
            ...         step_center=int(i/255*100),
            ...         step=i,
            ...     )
            ...
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

        Returns:
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

    def get_markpen_color(self):
        """
        현재 선택된 영역의 형광펜 색(RGB)을 튜플로 리턴하는 메서드

        Returns:
        """
        if self.hwp.SelectionMode == 0:
            self.hwp.HAction.Run("MoveNextChar")
            self.hwp.HAction.Run("MoveSelPrevChar")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as temp_file:
            temp_filename = temp_file.name
            self.save_block_as(temp_filename, format="HWPX")
        self.HAction.Run("Cancel")

        # 임시 디렉토리 생성
        with tempfile.TemporaryDirectory() as temp_dir:
            # ZIP 파일을 임시 디렉토리에 추출
            with zipfile.ZipFile(temp_filename, 'r') as zf:
                zf.extractall(path=temp_dir)

            # section0.xml 파일 경로
            section0_path = os.path.join(temp_dir, "Contents/section0.xml")

            # XML 파일 열기
            with open(section0_path, encoding="utf-8") as f:
                content = f.read()

            # 색상 정보 추출
            hex_text = re.findall(r'<hp:markpenBegin color="#([A-F0-9]{6})"', content)
            if not hex_text:
                raise AssertionError("형광펜 속성을 찾을 수 없습니다.")

            r = int("0x" + hex_text[0][:2], base=16)
            g = int("0x" + hex_text[0][2:4], base=16)
            b = int("0x" + hex_text[0][4:], base=16)
            try:
                return r, g, b
            finally:
                # 임시 ZIP 파일 삭제
                os.remove(temp_filename)

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

        Returns:
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

    def set_pagedef(self, pset:Any, apply: Literal["cur", "all", "new"] = "cur") -> bool:
        """
        get_pagedef 또는 get_pagedef_as_dict를 통해 얻은 용지정보를 현재구역에 적용하는 메서드

        Args:
            pset: 파라미터셋 또는 dict. 용지정보를 담은 객체

        Returns:
            성공시 True, 실패시 False를 리턴
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
        if self.hwp.SelectionMode == 0:
            return False
        if path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)
        pset = self.hwp.HParameterSet.HFileOpenSave
        self.hwp.HAction.GetDefault("FileSaveBlock_S", pset.HSet)
        pset.filename = path
        pset.Format = format
        pset.Attributes = attributes
        return self.hwp.HAction.Execute("FileSaveBlock_S", pset.HSet)

    def goto_printpage(self, page_num: int = 1) -> bool:
        """
        인쇄페이지 기준으로 해당 페이지로 이동

        1페이지의 page_num은 1이다.

        Args:
            page_num: 이동할 페이지번호

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        pset = self.hwp.HParameterSet.HGotoE
        self.hwp.HAction.GetDefault("Goto", pset.HSet)
        pset.HSet.SetItem("DialogResult", page_num)
        pset.SetSelectionIndex = 1
        return self.hwp.HAction.Execute("Goto", pset.HSet)

    def goto_page(self, page_index: int | str = 1) -> tuple[int, int]:
        """
        새쪽번호와 관계없이 페이지 순서를 통해

        특정 페이지를 찾아가는 메서드. 1이 1페이지임.

        Args:
            page_index: 찾아갈 페이지(시작페이지는 1)

        Returns:
            tuple(인쇄기준페이지, 페이지인덱스)
        """
        if int(page_index) > self.hwp.PageCount:
            return False
            # raise ValueError("입력한 페이지 인덱스가 문서 총 페이지보다 큽니다.")
        elif int(page_index) < 1:
            return False
            # raise ValueError("1 이상의 값을 입력해야 합니다.")
        self.goto_printpage(page_index)
        cur_page = self.current_page
        if page_index == cur_page:
            pass
        elif page_index < cur_page:
            for _ in range(cur_page - page_index):
                self.MovePageUp()
        else:
            for _ in range(page_index - cur_page):
                self.MovePageDown()
        return self.current_printpage, self.current_page

    def table_from_data(self, data:pd.DataFrame|dict|list|str, transpose:bool=False, header0:str="", treat_as_char:bool=False, header:bool=True, index:bool=True,
                        cell_fill: bool | tuple[int, int, int] = False, header_bold:bool=True) -> None:
        """
        dict, list 또는 csv나 xls, xlsx 및 json처럼 2차원 스프레드시트로 표현 가능한 데이터에 대해서,

        정확히는 pd.DataFrame으로 변환 가능한 데이터에 대해 아래아한글 표로 변환하는 작업을 한다.
        내부적으로 판다스 데이터프레임으로 변환하는 과정을 거친다.

        Args:
            data: 테이블로 변환할 데이터
            transpose: 행/열 전환
            header0: index=True일 경우 (1,1) 셀에 들어갈 텍스트
            treat_as_char: 글자처럼 취급 여부
            header: 1행을 "제목행"으로 선택할지 여부
            header_bold: 1행의 텍스트에 bold를 적용할지 여부

        Returns:
            None
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
        for field in self.get_field_list().split("\x02")[::-1]:
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

    def open_pdf(self, pdf_path:str, this_window:int=1) -> bool:
        """
        pdf를 hwp문서로 변환하여 여는 함수.

        (최초 실행시 "다시 표시 안함ㅁ" 체크박스에 체크를 해야 한다.)

        Args:
            pdf_path: pdf파일의 경로
            this_window: 현재 창에 열고 싶으면 1, 새 창에 열고 싶으면 0. 하지만 아직(2023.12.11.) 작동하지 않음.

        Returns:
            성공시 True, 실패시 False 리턴
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

    def insert_memo(self, text:str="", memo_type: Literal["revision", "memo"] = "memo") -> None:
        """
        선택한 단어 범위에 메모고침표를 삽입하는 코드.

        한/글에서 일반 문자열을 삽입하는 코드와 크게 다르지 않다.
        선택모드가 아닌 경우 캐럿이 위치한 단어에 메모고침표를 삽입한다.

        Args:
            text: 삽입할 메모 내용

        Returns:
            None
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

        Returns:
        표 안에 있으면 True, 그렇지 않으면 False를 리턴
        """
        if self.key_indicator()[-1].startswith("("):
            return True
        else:
            return False

    def find_backward(self, src:str, regex:bool=False) -> bool:
        """
        문서 위쪽으로 find 메서드를 수행.

        해당 단어를 선택한 상태가 되며,
        문서 처음에 도달시 False 리턴

        Args:
            src: 찾을 단어

        Returns:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴, 단어가 더이상 없으면 False를 리턴
        """
        self.SetMessageBoxMode(0x2fff1)
        init_pos = str(self.KeyIndicator())
        pset = self.hwp.HParameterSet.HFindReplace
        pset.MatchCase = 1
        pset.SeveralWords = 1
        pset.UseWildCards = 1
        pset.AutoSpell = 1
        pset.Direction = self.FindDir("Backward")
        pset.FindString = src
        pset.IgnoreMessage = 0
        pset.HanjaFromHangul = 1
        pset.FindRegExp = regex
        try:
            return self.hwp.HAction.Execute("RepeatFind", pset.HSet)
        finally:
            self.SetMessageBoxMode(0xfffff)

    def find_forward(self, src:str, regex:bool=False) -> bool:
        """
        문서 아래쪽으로 find를 수행하는 메서드.

        해당 단어를 선택한 상태가 되며,
        문서 끝에 도달시 False 리턴.

        Args:
            src: 찾을 단어

        Returns:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴, 단어가 더이상 없으면 False를 리턴
        """
        self.SetMessageBoxMode(0x2fff1)
        init_pos = str(self.KeyIndicator())
        pset = self.hwp.HParameterSet.HFindReplace
        pset.MatchCase = 1
        pset.SeveralWords = 1
        pset.UseWildCards = 1
        pset.AutoSpell = 1
        pset.Direction = self.FindDir("Forward")
        pset.FindString = src
        pset.IgnoreMessage = 0
        pset.HanjaFromHangul = 1
        pset.FindRegExp = regex
        try:
            return self.hwp.HAction.Execute("RepeatFind", pset.HSet)
        finally:
            self.SetMessageBoxMode(0xfffff)

    def find(self, src:str, direction: Literal["Forward", "Backward", "AllDoc"] = "Forward", regex:bool=False, MatchCase:int=1,
             SeveralWords:int=1, UseWildCards:int=1, WholeWordOnly:int=0, AutoSpell:int=1, HanjaFromHangul:int=1, AllWordForms:int=0,
             FindStyle:str="", ReplaceStyle:str="", FindJaso:int=0, FindType:int=1) -> bool:
        """
        direction 방향으로 특정 단어를 찾아가는 메서드.

        해당 단어를 선택한 상태가 되며,
        탐색방향에 src 문자열이 없는 경우 False를 리턴

        Args:
            src: 찾을 단어
            direction:
                탐색방향

                    - "Forward": 아래쪽으로(기본값)
                    - "Backward": 위쪽으로
                    - "AllDoc": 아래쪽 우선으로 찾고 문서끝 도달시 처음으로 돌아감.

            regex: 정규식 탐색(기본값 False)
            MatchCase: 대소문자 구분(기본값 1)
            SeveralWords: 여러 단어 찾기
            UseWildCards: 아무개 문자(1),
            WholeWordOnly: 온전한 낱말(0),
            AutoSpell:
            HanjaFromHangul: 한글로 한자 찾기(1),
            AllWordForms:
            FindStyle: 찾을 글자모양
            ReplaceStyle: 바꿀 글자모양
            FindJaso: 자소 단위 찾기(0),
            FindType:

        Returns:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴,
            단어가 더이상 없으면 False를 리턴
        """
        self.SetMessageBoxMode(0x2fff1)
        pset = self.hwp.HParameterSet.HFindReplace
        pset.MatchCase = MatchCase
        pset.SeveralWords = SeveralWords
        pset.UseWildCards = UseWildCards
        pset.WholeWordOnly = WholeWordOnly
        pset.AutoSpell = AutoSpell
        pset.Direction = self.FindDir(direction)
        pset.FindString = src
        pset.IgnoreMessage = 0
        pset.HanjaFromHangul = HanjaFromHangul
        pset.AllWordForms = AllWordForms
        pset.FindJaso = FindJaso
        pset.FindStyle = FindStyle
        pset.ReplaceStyle = ReplaceStyle
        pset.FindRegExp = regex
        pset.FindType = FindType
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

        Returns:
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

    def find_replace(self, src, dst, regex=False, direction: Literal["Backward", "Forward", "AllDoc"] = "Forward",
                     MatchCase=1, AllWordForms=0, SeveralWords=1, UseWildCards=1, WholeWordOnly=0, AutoSpell=1,
                     IgnoreFindString=0, IgnoreReplaceString=0, ReplaceMode=1, HanjaFromHangul=1,
                     FindJaso=0, FindStyle="", ReplaceStyle="", FindType=1):
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
                    return self.find_replace(i, j, direction=direction, MatchCase=MatchCase, AllWordForms=AllWordForms,
                                             SeveralWords=SeveralWords, UseWildCards=UseWildCards,
                                             WholeWordOnly=WholeWordOnly, AutoSpell=AutoSpell,
                                             IgnoreFindString=IgnoreFindString, IgnoreReplaceString=IgnoreReplaceString,
                                             ReplaceMode=ReplaceMode, HanjaFromHangul=HanjaFromHangul,
                                             FindJaso=FindJaso,
                                             FindStyle=FindStyle, ReplaceStyle=ReplaceStyle, FindType=FindType)
                finally:
                    self.SetMessageBoxMode(0xfffff)

        else:
            pset = self.hwp.HParameterSet.HFindReplace
            # self.hwp.HAction.GetDefault("ExecReplace", pset.HSet)
            pset.MatchCase = MatchCase
            pset.AllWordForms = AllWordForms
            pset.SeveralWords = SeveralWords
            pset.UseWildCards = UseWildCards
            pset.WholeWordOnly = WholeWordOnly
            pset.AutoSpell = AutoSpell
            pset.Direction = self.hwp.FindDir(direction)
            pset.IgnoreFindString = IgnoreFindString
            pset.IgnoreReplaceString = IgnoreReplaceString
            pset.FindString = src  # "\\r\\n"
            pset.ReplaceString = dst  # "^n"
            pset.ReplaceMode = ReplaceMode
            pset.IgnoreMessage = 0
            pset.HanjaFromHangul = HanjaFromHangul
            pset.FindJaso = FindJaso
            pset.FindRegExp = 0
            pset.FindStyle = FindStyle
            pset.ReplaceStyle = ReplaceStyle
            pset.FindType = FindType
            try:
                return self.hwp.HAction.Execute("ExecReplace", pset.HSet)
            finally:
                self.SetMessageBoxMode(0xfffff)

    def find_replace_all(self, src, dst, regex=False, MatchCase=1, AllWordForms=0, SeveralWords=1, UseWildCards=1,
                         WholeWordOnly=0, AutoSpell=1, IgnoreFindString=0, IgnoreReplaceString=0, ReplaceMode=1,
                         HanjaFromHangul=1, FindJaso=0, FindStyle="", ReplaceStyle="", FindType=1):
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
                self.find_replace_all(i, j, MatchCase=MatchCase, AllWordForms=AllWordForms, SeveralWords=SeveralWords,
                                      UseWildCards=UseWildCards, WholeWordOnly=WholeWordOnly, AutoSpell=AutoSpell,
                                      IgnoreFindString=IgnoreFindString, IgnoreReplaceString=IgnoreReplaceString,
                                      ReplaceMode=ReplaceMode, HanjaFromHangul=HanjaFromHangul, FindJaso=FindJaso,
                                      FindStyle=FindStyle, ReplaceStyle=ReplaceStyle, FindType=FindType)
        else:
            pset = self.hwp.HParameterSet.HFindReplace
            # self.hwp.HAction.GetDefault("AllReplace", pset.HSet)
            pset.MatchCase = MatchCase
            pset.AllWordForms = AllWordForms
            pset.SeveralWords = SeveralWords
            pset.UseWildCards = UseWildCards
            pset.WholeWordOnly = WholeWordOnly
            pset.AutoSpell = AutoSpell
            pset.Direction = self.hwp.FindDir("AllDoc")
            pset.IgnoreFindString = IgnoreFindString
            pset.IgnoreReplaceString = IgnoreReplaceString
            pset.FindString = src  # "\\r\\n"
            pset.ReplaceString = dst  # "^n"
            pset.ReplaceMode = ReplaceMode
            pset.IgnoreMessage = 0
            pset.HanjaFromHangul = HanjaFromHangul
            pset.FindJaso = FindJaso
            pset.FindRegExp = 0
            pset.FindStyle = FindStyle
            pset.ReplaceStyle = ReplaceStyle
            pset.FindType = FindType
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

    def switch_to(self, num:int) -> bool:
        """
        여러 개의 hwp인스턴스가 열려 있는 경우 해당 인덱스의 문서창 인스턴스를 활성화한다.

        Args:
            num: 인스턴스 번호
        """
        try:
            self.hwp.XHwpDocuments.Item(num).SetActive_XHwpDocument()
            return True
        except pythoncom.com_error as e:
            return False

    def add_tab(self) -> "Hwp.XHwpDocuments":
        """
        새 문서를 현재 창의 새 탭에 추가한다.

        백그라운드 상태에서 새 창을 만들 때 윈도우에 나타나는 경우가 있는데,
        add_tab() 함수를 사용하면 백그라운드 작업이 보장된다.
        탭 전환은 switch_to() 메서드로 가능하다.

        새 창을 추가하고 싶은 경우는 add_tab 대신 hwp.FileNew()나 hwp.add_doc()을 실행하면 된다.
        """
        return self.hwp.XHwpDocuments.Add(1)  # 0은 새 창, 1은 새 탭

    def add_doc(self) -> "Hwp.XHwpDocuments":
        """
        새 문서를 추가한다.

        원래 창이 백그라운드로 숨겨져 있어도 추가된 문서는 보이는 상태가 기본값이다.
        숨기려면 `hwp.set_visible(False)`를 실행해야 한다.
        새 탭을 추가하고 싶은 경우는 `add_doc` 대신 `add_tab`을 실행하면 된다.
        """
        return self.hwp.XHwpDocuments.Add(0)  # 0은 새 창, 1은 새 탭

    def hwp_unit_to_mili(self, hwp_unit:int) -> float:
        """
        HwpUnit 값을 밀리미터로 변환한 값을 리턴한다.

        HwpUnit으로 리턴되었거나, 녹화된 코드의 HwpUnit값을 확인할 때 유용하게 사용할 수 있다.

        Returns:
            HwpUnit을 7200으로 나눈 후 25.4를 곱하고 소숫점 셋째자리에서 반올림한 값
        """
        if hwp_unit == 0:
            return 0
        else:
            return round(hwp_unit / 7200 * 25.4, 2)

    def HwpUnitToMili(self, hwp_unit:int) -> float:
        """
        HwpUnit 값을 밀리미터로 변환한 값을 리턴한다.

        HwpUnit으로 리턴되었거나, 녹화된 코드의 HwpUnit값을 확인할 때 유용하게 사용할 수 있다.

        Returns:
            HwpUnit을 7200으로 나눈 후 25.4를 곱하고 소숫점 셋째 자리에서 반올림한 값
        """
        if hwp_unit == 0:
            return 0
        else:
            return round(hwp_unit / 7200 * 25.4, 2)

    def create_table(self, rows, cols, treat_as_char: bool = True, width_type=0, height_type=0, header=True,
                     height=0) -> bool:
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

        Returns:
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
        each_col_width = round((total_width - self.mili_to_hwp_unit(3.6 * cols)) / cols)
        for i in range(cols):
            pset.ColWidth.SetItem(i, each_col_width)  # 1열
        if self.Version[0] == "8":
            pset.TableProperties.TreatAsChar = treat_as_char  # 글자처럼 취급
        pset.TableProperties.Width = total_width  # self.hwp.MiliToHwpUnit(148)  # 표 너비
        try:
            return self.hwp.HAction.Execute("TableCreate", pset.HSet)  # 위 코드 실행
        finally:
            # 글자처럼 취급 여부 적용(treat_as_char)
            if self.Version[0] != "8":
                ctrl = self.hwp.CurSelectedCtrl or self.hwp.ParentCtrl
                pset = self.hwp.CreateSet("Table")
                pset.SetItem("TreatAsChar", treat_as_char)
                ctrl.Properties = pset

            # 제목 행 여부 적용(header)
            if header:
                pset = self.hwp.HParameterSet.HShapeObject
                self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                pset.ShapeTableCell.Header = header
                self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

    def get_selected_text(self, as_: Literal["list", "str"] = "str"):
        """
        한/글 문서 선택 구간의 텍스트를 리턴하는 메서드.

        Returns:
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

    def table_to_csv(self, n="", filename="result.csv", encoding="utf-8", startrow=0) -> None:
        """
        한/글 문서의 idx번째 표를 현재 폴더에 filename으로 csv포맷으로 저장한다.

        filename을 지정하지 않는 경우 "./result.csv"가 기본값이다.

        Returns:
        None을 리턴하고, 표데이터를 "./result.csv"에 저장한다.

        Examples:
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

    def table_to_df_q(self, n="", startrow=0, columns=[]) -> pd.DataFrame:
        """
        (2024. 3. 14. for문 추출 구조에서, 한 번에 추출하는 방식으로 변경->속도개선)

        한/글 문서의 n번째 표를 판다스 데이터프레임으로 리턴하는 메서드.
        n을 넣지 않는 경우, 캐럿이 셀에 있다면 해당 표를 df로,
        캐럿이 표 밖에 있다면 첫 번째 표를 df로 리턴한다.
        startrow는 표 제목에 일부 병합이 되어 있는 경우
        df로 변환시작할 행을 특정할 때 사용된다.

        Returns:
        아래아한글 표 데이터를 가진 판다스 데이터프레임 인스턴스

        Examples:
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

    def table_to_df(self, n="", cols=0, selected_range=None, start_pos=None) -> pd.DataFrame:
        """
        (2025. 3. 3. RowSpan이랑 ColSpan을 이용해서, 중복되는 값은 그냥 모든 셀에 넣어버림

        한/글 문서의 n번째 표를 판다스 데이터프레임으로 리턴하는 메서드.
        n을 넣지 않는 경우, 캐럿이 셀에 있다면 해당 표를 df로,
        캐럿이 표 밖에 있다면 첫 번째 표를 df로 리턴한다.

        Returns:
        아래아한글 표 데이터를 가진 판다스 데이터프레임 인스턴스

        Examples:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> df = hwp.table_to_df()  # 현재 캐럿이 들어가 있는 표 전체를 df로(1행을 df의 칼럼으로)
            >>> df = hwp.table_to_df(0, cols=2)  # 문서의 첫 번째 표를 df로(2번인덱스행(3행)을 칼럼명으로, 그 아래(4행부터)를 값으로)
        """
        if self.SelectionMode != 19:
            start_pos = self.hwp.GetPos()
            ctrl = self.hwp.HeadCtrl
            if isinstance(n, type(ctrl)):
                # 정수인덱스 대신 ctrl 객체를 넣은 경우
                self.set_pos_by_set(n.GetAnchorPos(0))
                self.find_ctrl()
            elif n == "" and self.is_cell():
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
                            break
                        else:
                            if idx == n:
                                self.set_pos_by_set(ctrl.GetAnchorPos(0))
                                self.hwp.FindCtrl()
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
        else:
            selected_range = self.get_selected_range()
        xml_data = self.GetTextFile("HWPML2X", option="saveblock")
        root = ET.fromstring(xml_data)

        # TABLE 태그에 RowCount, ColCount가 있으면 사용하고, 없으면 ROW, CELL 수로 결정
        table_el = root.find('.//TABLE')
        if table_el is not None:
            row_count = int(table_el.attrib.get("RowCount", "0"))
            col_count = int(table_el.attrib.get("ColCount", "0"))
        else:
            rows = root.findall('.//ROW')
            row_count = len(rows)
            col_count = max(len(row.findall('.//CELL')) for row in rows)

        # 결과를 저장할 2차원 리스트 초기화 (빈 문자열로 채움)
        result = [["" for _ in range(col_count)] for _ in range(row_count)]

        row_index = 0
        for row in root.findall('.//ROW'):
            col_index = 0
            for cell in row.findall('.//CELL'):
                # 이미 값이 채워진 셀이 있으면 건너뛰고 다음 빈 칸 찾기
                while col_index < col_count and result[row_index][col_index] != "":
                    col_index += 1
                if col_index >= col_count:
                    break

                # CELL 내 텍스트 추출 (CHAR 태그의 텍스트 연결)
                cell_text = ''
                for text in cell.findall('.//TEXT'):
                    for char in text.findall('.//CHAR'):
                        if char.text:
                            cell_text += char.text
                    cell_text += "\r\n"
                if cell_text.endswith("\r\n"):
                    cell_text = cell_text[:-2]

                # RowSpan과 ColSpan 값 읽기 (기본값은 1)
                row_span = int(cell.attrib.get("RowSpan", "1"))
                col_span = int(cell.attrib.get("ColSpan", "1"))

                # 행과 열로 병합된 영역에 대해 값을 채워줌
                for i in range(row_span):
                    for j in range(col_span):
                        if row_index + i < row_count and col_index + j < col_count:
                            result[row_index + i][col_index + j] = cell_text
                # 현재 셀이 차지한 열 수만큼 col_index를 이동
                col_index += col_span
            row_index += 1

        # 선택 영역이 있을 경우 후처리
        if self.SelectionMode == 19:
            result = crop_data_from_selection(result, selected_range)

        # DataFrame 생성: cols가 int면 해당 인덱스 행을 header로 사용
        if type(cols) == int:
            columns = result[cols]
            data = result[cols + 1:]
            df = pd.DataFrame(data, columns=columns)
        elif type(cols) in (list, tuple):
            df = pd.DataFrame(result, columns=cols)

        try:
            return df
        finally:
            if self.SelectionMode != 19:
                self.set_pos(*start_pos)

    def table_to_bottom(self, offset:float=0.) -> bool:
        """
        표 앞에 캐럿을 둔 상태 또는 캐럿이 표 안에 있는 상태에서 위 함수 실행시

        표를 (페이지 기준) 하단으로 위치시킨다.

        Args:
            offset: 페이지 하단 기준 오프셋(mm)

        Returns:
            성공시 True
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

        Returns:
        삽입 성공시 True, 실패시 False를 리턴함.
        Examples:
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

    def move_all_caption(self, location: Literal["Top", "Bottom", "Left", "Right"] = "Bottom",
                         align: Literal["Left", "Center", "Right", "Distribute", "Division", "Justify"] = "Justify"):
        """
        한/글 문서 내 모든 표, 그림의 주석 위치를 일괄 변경하는 메서드.

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

    def CheckXObject(self, bstring):
        return self.hwp.CheckXObject(bstring=bstring)

    def clear(self, option: int = 1) -> None:
        """
        현재 편집중인 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다.

        Args:
            option: 편집중인 문서의 내용에 대한 처리 방법, 생략하면 1(hwpDiscard)가 선택된다.

                - 0: 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다. (hwpAskSave)
                - 1: 문서의 내용을 버린다. (hwpDiscard, 기본값)
                - 2: 문서가 변경된 경우 저장한다. (hwpSaveIfDirty)
                - 3: 무조건 저장한다. (hwpSave)

        Returns:
            None

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.clear()
        """
        return self.hwp.XHwpDocuments.Active_XHwpDocument.Clear(option=option)

    def Clear(self, option: int = 1) -> None:
        """
        현재 편집중인 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다.

        Args:
            option:
                편집중인 문서의 내용에 대한 처리 방법, 생략하면 1(hwpDiscard)가 선택된다.

                    - 0: 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다. (hwpAskSave)
                    - 1: 문서의 내용을 버린다. (hwpDiscard, 기본값)
                    - 2: 문서가 변경된 경우 저장한다. (hwpSaveIfDirty)
                    - 3: 무조건 저장한다. (hwpSave)

        Returns:
            None

        Examples:
            >>> from pyhwpx import Hwp
            >>>
            >>> hwp = Hwp()
            >>> hwp.clear()
        """
        return self.hwp.XHwpDocuments.Active_XHwpDocument.Clear(option=option)

    def close(self, is_dirty: bool = False, interval:float=0.01) -> bool:
        """
        문서를 버리고 닫은 후, 새 문서창을 여는 메서드.

        굳이 새 문서파일이 필요한 게 아니라면 ``hwp.close()`` 대신 ``hwp.clear()``를 사용할 것.

        Args:
            is_dirty: True인 경우 변경사항이 있을 때 문서를 닫지 않는다. False일 때는 변경사항을 버리고 문서를 닫음

        Returns:
            문서창을 닫으면 True, 문서창 닫기에 실패하면 False 리턴
        """
        while True:
            try:
                return self.hwp.XHwpDocuments.Active_XHwpDocument.Close(isDirty=is_dirty)
            except AttributeError:
                sleep(interval)

    def create_action(self, actidstr: str) -> Any:
        """
        Action 객체를 생성한다.

        액션에 대한 세부적인 제어가 필요할 때 사용한다.
        예를 들어 기능을 수행하지 않고 대화상자만을 띄운다든지,
        대화상자 없이 지정한 옵션에 따라 기능을 수행하는 등에 사용할 수 있다.

        Args:
            actidstr: 액션 ID (ActionIDTable.hwp 참조)

        Returns:
            Action object

        Examples:
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

    def CreateAction(self, actidstr: str) -> "Hwp.HAction":
        """
        Action 객체를 생성한다.

        액션에 대한 세부적인 제어가 필요할 때 사용한다.
        예를 들어 기능을 수행하지 않고 대화상자만을 띄운다든지,
        대화상자 없이 지정한 옵션에 따라 기능을 수행하는 등에 사용할 수 있다.

        Args:
            actidstr: 액션 ID (ActionIDTable.hwp 참조)

        Returns:
            Action object

        Examples:
            >>> from pyhwpx import Hwp
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

        Args:
            name: 누름틀 필드에 대한 필드 이름(중요)
            direction: 누름틀에 입력이 안 된 상태에서 보이는 안내문/지시문.
            memo: 누름틀에 대한 설명/도움말

        Returns:
            성공이면 True, 실패면 False

        Examples:
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

        Args:
            name: 누름틀 필드에 대한 필드 이름(중요)
            direction: 누름틀에 입력이 안 된 상태에서 보이는 안내문/지시문.
            memo: 누름틀에 대한 설명/도움말

        Returns:
            성공이면 True, 실패면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_field(direction="이름", memo="이름을 입력하는 필드", name="name")
            True
            >>> hwp.put_field_text("name", "일코")
        """
        return self.hwp.CreateField(Direction=direction, memo=memo, name=name)

    def CreateId(self, creation_id):
        return self.hwp.CreateID(CreationID=creation_id)

    def CreateMode(self, creation_mode):
        return self.hwp.CreateMode(CreationMode=creation_mode)

    def create_page_image(self, path: str, pgno: int = -1, resolution: int = 300, depth: int = 24,
                          format: str = "bmp") -> bool:
        """
        ``pgno``로 지정한 페이지를 ``path`` 라는 파일명으로 저장한다.
        이 때 페이지번호는 1부터 시작하며,(1-index)
        ``pgno=0``이면 현재 페이지, ``pgno=-1``(기본값)이면 전체 페이지를 이미지로 저장한다.
        내부적으로 pillow 모듈을 사용하여 변환하므로,
        사실상 pillow에서 변환 가능한 모든 포맷으로 입력 가능하다.

        Args:
            path: 생성할 이미지 파일의 경로(전체경로로 입력해야 함)
            pgno:
                페이지 번호(1페이지 저장하려면 pgno=1).
                1부터 hwp.PageCount 사이에서 pgno 입력시 선택한 페이지만 저장한다.
                생략하면(기본값은 -1) 전체 페이지가 저장된다.
                이 때 path가 "img.jpg"라면 저장되는 파일명은
                "img001.jpg", "img002.jpg", "img003.jpg",..,"img099.jpg" 가 된다.
                현재 캐럿이 있는 페이지만 저장하고 싶을 때에는 ``pgno=0``으로 설정하면 된다.
            resolution:
                이미지 해상도. DPI단위(96, 300, 1200 등)로 지정한다.
                생략하면 300이 사용된다.
            depth: 이미지파일의 Color Depth(1, 4, 8, 24)를 지정한다. 생략하면 24
            format: 이미지파일의 포맷. "bmp", "gif"중의 하나. 생략하면 "bmp"가 사용된다.

        Returns:
            성공하면 True, 실패하면 False

        Examples:
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
        ``pgno``로 지정한 페이지를 ``path`` 라는 파일명으로 저장한다.
        이 때 페이지번호는 1부터 시작하며,(1-index)
        ``pgno=0``이면 현재 페이지, ``pgno=-1``(기본값)이면 전체 페이지를 이미지로 저장한다.
        내부적으로 pillow 모듈을 사용하여 변환하므로,
        사실상 pillow에서 변환 가능한 모든 포맷으로 입력 가능하다.

        Args:
            path: 생성할 이미지 파일의 경로(전체경로로 입력해야 함)
            pgno:
                페이지 번호(1페이지 저장하려면 pgno=1).
                1부터 hwp.PageCount 사이에서 pgno 입력시 선택한 페이지만 저장한다.
                생략하면(기본값은 -1) 전체 페이지가 저장된다.
                이 때 path가 "img.jpg"라면 저장되는 파일명은
                "img001.jpg", "img002.jpg", "img003.jpg",..,"img099.jpg" 가 된다.
                현재 캐럿이 있는 페이지만 저장하고 싶을 때에는 ``pgno=0``으로 설정하면 된다.
            resolution:
                이미지 해상도. DPI단위(96, 300, 1200 등)로 지정한다.
                생략하면 300이 사용된다.
            depth: 이미지파일의 Color Depth(1, 4, 8, 24)를 지정한다. 생략하면 24
            format: 이미지파일의 포맷. "bmp", "gif"중의 하나. 생략하면 "bmp"가 사용된다.

        Returns:
            성공하면 True, 실패하면 False

        Examples:
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

    def CreateSet(self, setidstr: str) -> "Hwp.HParameterSet":
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

        Args:
            setidstr: 생성할 ParameterSet의 ID (ParameterSet Table.hwp 참고)

        Returns:
            생성된 ParameterSet Object
        """
        return self.hwp.CreateSet(setidstr=setidstr)

    def delete_ctrl(self, ctrl: Ctrl) -> bool:
        """
        문서 내 컨트롤을 삭제한다.

        Args:
            ctrl: 삭제할 문서 내 컨트롤

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> ctrl = hwp.HeadCtrl.Next.Next
            >>> if ctrl.UserDesc == "표":
            ...     hwp.delete_ctrl(ctrl)
            ...
            True
        """
        return self.hwp.DeleteCtrl(ctrl=ctrl)

    def DeleteCtrl(self, ctrl: Ctrl) -> bool:
        """
        문서 내 컨트롤을 삭제한다.

        Args:
            ctrl: 삭제할 문서 내 컨트롤

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> ctrl = hwp.HeadCtrl.Next.Next
            >>> if ctrl.UserDesc == "표":
            ...     hwp.delete_ctrl(ctrl)
            ...
            True
        """
        return self.hwp.DeleteCtrl(ctrl=ctrl)

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

    def EquationCreate(self, thread=False):
        """
        수식만들기 창을 직접 제어하기 위한 스레드 함수. 직접 사용하지 말 것.
        """
        visible = self.hwp.XHwpWindows.Active_XHwpWindow.Visible
        if thread:
            if win32gui.FindWindow(None, "수식 편집기"):
                return False
            t = threading.Thread(target=_eq_create, args=(visible,), name="eq_create")
            t.start()
            t.join(timeout=0)
            return True
        else:
            return self.hwp.HAction.Run("EquationCreate")

    def EquationClose(self, save=False, delay=0.1):
        """
        ``EquationCreate(thread=True)``로 생성한 수식창을 닫기 위한 메서드.
        """
        return _close_eqedit(save, delay)

    def EquationModify(self, thread=False):
        """
        ``EquationCreate(thread=True)``로 생성한 수식창에서 편집작업을 하기 위한 메서드.
        """
        visible = self.hwp.XHwpWindows.Active_XHwpWindow.Visible
        if thread:
            if win32gui.FindWindow(None, "수식 편집기"):
                return False
            t = threading.Thread(target=_eq_modify, args=(visible,), name="eq_modify")
            t.start()
            t.join(timeout=0)
            return True
        else:
            return self.hwp.HAction.Run("EquationModify")

    def EquationRefresh(self) -> bool:  # , delay=0.2):
        """
        수식을 정형화함. (kosohn님께서 도움 주셔서 만든 메서드.)

        Returns:
            성공시 True, 실패시 False를 리턴함.

        """
        pset = self.hwp.HParameterSet.HEqEdit
        self.hwp.HAction.GetDefault("EquationModify", pset.HSet)
        pset.string = pset.VisualString
        pset.Version = "Equation Version 60"
        return self.hwp.HAction.Execute("EquationModify", pset.HSet)

    def export_style(self, sty_filepath: str) -> bool:
        """
        현재 문서의 Style을 sty 파일로 Export한다. #스타일 #내보내기

        Args:
            sty_filepath: Export할 sty 파일의 전체경로 문자열

        Returns:
            성공시 True, 실패시 False

        Examples:
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
        현재 문서의 Style을 sty 파일로 Export한다.  #스타일 #내보내기

        Args:
            sty_filepath: Export할 sty 파일의 전체경로 문자열

        Returns:
            성공시 True, 실패시 False

        Examples:
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

    def field_exist(self, field: str) -> bool:
        """
        문서에 해당 이름의 데이터 필드가 존재하는지 검사한다.

        Args:
            field: 필드이름

        Returns:
            필드가 존재하면 True, 존재하지 않으면 False
        """
        return self.hwp.FieldExist(Field=field)

    def FieldExist(self, field:str) -> bool:
        """
        문서에 해당 이름의 데이터 필드(누름틀, 셀필드)가 존재하는지 검사한다.

        Args:
            field: 필드이름

        Returns:
            필드가 존재하면 True, 존재하지 않으면 False
        """
        return self.hwp.FieldExist(Field=field)

    def FileTranslate(self, cur_lang:str="ko", trans_lang:str="en") -> bool:
        """
        문서를 번역함(Ctrl-Z 안 됨.) 한 달 10,000자 무료

        Args:
            cur_lang: 현재 문서 언어(예 - ko)
            trans_lang: 목표언어(예 - en)

        Returns:
            성공 후 True 리턴(실패하면 프로그램 종료됨ㅜ)

        Examples:



        """
        return self.hwp.FileTranslate(curLang=cur_lang, transLang=trans_lang)

    def FillAreaType(self, fill_area):
        return self.hwp.FillAreaType(FillArea=fill_area)

    def FindDir(self, find_dir: Literal["Forward", "Backward", "AllDoc"] = "Forward"):
        return self.hwp.FindDir(FindDir=find_dir)

    def find_ctrl(self):
        """컨트롤 선택하기"""
        return self.hwp.FindCtrl()

    def FindCtrl(self):
        """컨트롤 선택하기"""
        return self.hwp.FindCtrl()

    def find_private_info(self, private_type:int, private_string:str) -> int:
        """
        개인정보를 찾는다. (비밀번호 설정 등의 이유, 현재 비활성화된 것으로 추정)

        Args:
            private_type:
                보호할 개인정보 유형. 다음의 값을 하나이상 조합한다.

                    - 0x0001: 전화번호
                    - 0x0002: 주민등록번호
                    - 0x0004: 외국인등록번호
                    - 0x0008: 전자우편
                    - 0x0010: 계좌번호
                    - 0x0020: 신용카드번호
                    - 0x0040: IP 주소
                    - 0x0080: 생년월일
                    - 0x0100: 주소
                    - 0x0200: 사용자 정의
                    - 0x0400: 기타

            private_string: 기타 문자열. 0x0400 유형이 존재할 경우에만 유효하므로, 생략가능하다. (예: "신한카드")

        Returns:
            찾은 개인정보의 유형 값. 개인정보가 없는 경우에는 0을 반환한다. 또한, 검색 중 문서의 끝(end of document)을 만나면 –1을 반환한다. 이는 함수가 무한히 반복하는 것을 막아준다. 구체적으로는 아래와 같다.

                - 0x0001 : 전화번호
                - 0x0002 : 주민등록번호
                - 0x0004 : 외국인등록번호
                - 0x0008 : 전자우편
                - 0x0010 : 계좌번호
                - 0x0020 : 신용카드번호
                - 0x0040 : IP 주소
                - 0x0080 : 생년월일
                - 0x0100 : 주소
                - 0x0200 : 사용자 정의
                - 0x0400 : 기타
        """
        return self.hwp.FindPrivateInfo(PrivateType=private_type, PrivateString=private_string)

    def FindPrivateInfo(self, private_type:int, private_string:str) -> int:
        """
        개인정보를 찾는다. (비밀번호 설정 등의 이유, 현재 비활성화된 것으로 추정)

        Args:
            private_type:
                보호할 개인정보 유형. 다음의 값을 하나이상 조합한다.

                    - 0x0001: 전화번호
                    - 0x0002: 주민등록번호
                    - 0x0004: 외국인등록번호
                    - 0x0008: 전자우편
                    - 0x0010: 계좌번호
                    - 0x0020: 신용카드번호
                    - 0x0040: IP 주소
                    - 0x0080: 생년월일
                    - 0x0100: 주소
                    - 0x0200: 사용자 정의
                    - 0x0400: 기타

            private_string: 기타 문자열. 0x0400 유형이 존재할 경우에만 유효하므로, 생략가능하다. (예: "신한카드")

        Returns:
            찾은 개인정보의 유형 값. 개인정보가 없는 경우에는 0을 반환한다. 또한, 검색 중 문서의 끝(end of document)을 만나면 –1을 반환한다. 이는 함수가 무한히 반복하는 것을 막아준다. 구체적으로는 아래와 같다.

                - 0x0001 : 전화번호
                - 0x0002 : 주민등록번호
                - 0x0004 : 외국인등록번호
                - 0x0008 : 전자우편
                - 0x0010 : 계좌번호
                - 0x0020 : 신용카드번호
                - 0x0040 : IP 주소
                - 0x0080 : 생년월일
                - 0x0100 : 주소
                - 0x0200 : 사용자 정의
                - 0x0400 : 기타
        """
        return self.hwp.FindPrivateInfo(PrivateType=private_type, PrivateString=private_string)

    def get_bin_data_path(self, binid: int) -> str:
        """
        Binary Data(Temp Image 등)의 경로를 가져온다.

        Args:
            binid: 바이너리 데이터의 ID 값 (1부터 시작)

        Returns:
            바이너리 데이터의 경로

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> path = hwp.get_bin_data_path(2)
            >>> print(path)
            C:/Users/User/AppData/Local/Temp/Hnc/BinData/EMB00004dd86171.jpg
        """
        return self.hwp.GetBinDataPath(binid=binid)

    def GetBinDataPath(self, binid: int) -> str:
        """
        Binary Data(Temp Image 등)의 경로를 가져온다.

        Args:
            binid: 바이너리 데이터의 ID 값 (1부터 시작)

        Returns:
            바이너리 데이터의 경로

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> path = hwp.get_bin_data_path(2)
            >>> print(path)
            C:/Users/User/AppData/Local/Temp/Hnc/BinData/EMB00004dd86171.jpg
        """
        return self.hwp.GetBinDataPath(binid=binid)

    def get_cur_field_name(self, option:int=0) -> str:
        """
        현재 캐럿이 위치하는 곳의 필드이름을 구한다.
        이 함수를 통해 현재 필드가 셀필드인지 누름틀필드인지 구할 수 있다.
        참고로, 필드 좌측에 커서가 붙어있을 때는 이름을 구할 수 있지만,
        우측에 붙어 있을 때는 작동하지 않는다.
        GetFieldList()의 옵션 중에 hwpFieldSelection(=4)옵션은 사용하지 않는다.


        Args:
            option: 다음과 같은 옵션을 지정할 수 있다.

                - 0: 모두 off. 생략하면 0이 지정된다.
                - 1: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다.(hwpFieldCell)
                - 2: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)

        Returns:
            필드이름이 돌아온다. 필드이름이 없는 경우 빈 문자열이 돌아온다.
        """
        return self.hwp.GetCurFieldName(option=option)

    def GetCurFieldName(self, option:int=0) -> str:
        """
        현재 캐럿이 위치하는 곳의 필드이름을 구한다.
        이 함수를 통해 현재 필드가 셀필드인지 누름틀필드인지 구할 수 있다.
        참고로, 필드 좌측에 커서가 붙어있을 때는 이름을 구할 수 있지만,
        우측에 붙어 있을 때는 작동하지 않는다.
        GetFieldList()의 옵션 중에 hwpFieldSelection(=4)옵션은 사용하지 않는다.


        Args:
            option: 다음과 같은 옵션을 지정할 수 있다.

                - 0: 모두 off. 생략하면 0이 지정된다.
                - 1: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다.(hwpFieldCell)
                - 2: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)

        Returns:
            필드이름이 돌아온다. 필드이름이 없는 경우 빈 문자열이 돌아온다.
        """
        return self.hwp.GetCurFieldName(option=option)

    def get_cur_metatag_name(self) -> str:
        """
        현재 캐럿위치의 메타태그 이름을 리턴하는 메서드.

        Returns:
            str: 현재 위치의 메타태그 이름

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # "#test"라는 메타태그 이름이 부여된 표를 선택한 상태에서
            >>> hwp.get_cur_metatag_name()
            #test

        """
        try:
            return self.hwp.GetCurMetatagName()
        except pythoncom.com_error as e:
            print(e, "메타태그명을 출력할 수 없습니다.")

    def GetCurMetatagName(self):
        """
        현재 캐럿위치의 메타태그 이름을 리턴하는 메서드.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # "#test"라는 메타태그 이름이 부여된 표를 선택한 상태에서
            >>> hwp.GetCurMetatagName()
            #test
        """
        try:
            return self.hwp.GetCurMetatagName()
        except pythoncom.com_error as e:
            # 가끔 com_error 발생한다. (대부분 최초 실행시?)
            print(e, "메타태그명을 출력할 수 없습니다.")

    def get_field_list(self, number: int = 1, option: int = 0) -> str:
        """
        문서에 존재하는 필드의 목록을 구한다.

        문서 중에 동일한 이름의 필드가 여러 개 존재할 때는
        number에 지정한 타입에 따라 3 가지의 서로 다른 방식 중에서 선택할 수 있다.
        예를 들어 문서 중 title, body, title, body, footer 순으로
        5개의 필드가 존재할 때, hwpFieldPlain, hwpFieldNumber, HwpFieldCount
        세 가지 형식에 따라 다음과 같은 내용이 돌아온다.

            - hwpFieldPlain: "title\x02body\x02title\x02body\x02footer"
            - hwpFieldNumber: "title{{0}}\x02body{{0}}\x02title{{1}}\x02body{{1}}\x02footer{{0}}"
            - hwpFieldCount: "title{{2}}\x02body{{2}}\x02footer{{1}}"

        Args:
            number:
                문서 내에서 동일한 이름의 필드가 여러 개 존재할 경우 이를 구별하기 위한 식별방법을 지정한다. 생략하면 0(hwpFieldPlain)이 지정된다.

                    - 0: 아무 기호 없이 순서대로 필드의 이름을 나열한다.(hwpFieldPlain)
                    - 1: 필드이름 뒤에 일련번호가 ``{{#}}``과 같은 형식으로 붙는다.(hwpFieldNumber)
                    - 2: 필드이름 뒤에 그 필드의 개수가 ``{{#}}``과 같은 형식으로 붙는다.(hwpFieldCount)

            option:
                다음과 같은 옵션을 조합할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다.

                    - 0x01: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere과는 함께 지정할 수 없다.(hwpFieldCell)
                    - 0x02: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)
                    - 0x04: 선택된 내용 안에 존재하는 필드 리스트를 구한다.(HwpFieldSelection)

        Returns:
            각 필드 사이를 문자코드 0x02로 구분하여 다음과 같은 형식으로 리턴 한다.
            (가장 마지막 필드에는 0x02가 붙지 않는다.)
            "필드이름#1\\x02필드이름#2\\x02...필드이름#n"
        """
        return self.hwp.GetFieldList(Number=number, option=option)

    def GetFieldList(self, number:int=1, option:int=0) -> str:
        """
        문서에 존재하는 필드의 목록을 구한다.

        문서 중에 동일한 이름의 필드가 여러 개 존재할 때는
        number에 지정한 타입에 따라 3가지의 서로 다른 방식 중에서 선택할 수 있다.

        예를 들어 문서 중 title, body, title, body, footer 순으로
        5개의 필드가 존재할 때, 0(hwpFieldPlain), 1(hwpFieldNumber), 2(HwpFieldCount)
        세 가지 형식에 따라 다음과 같은 내용이 돌아온다.

            - 0 (hwpFieldPlain): "title\\x02body\\x02title\\x02body\\x02footer"
            - 1 (hwpFieldNumber): "title{{0}}\\x02body{{0}}\\x02title{{1}}\\x02body{{1}}\\x02footer{{0}}"
            - 2 (hwpFieldCount): "title{{2}}\\x02body{{2}}\\x02footer{{1}}"

        Args:
            number:
                문서 내에서 동일한 이름의 필드가 여러 개 존재할 경우 이를 구별하기 위한 식별방법을 지정한다. 생략하면 0(hwpFieldPlain)이 지정된다.

                    - 0: 아무 기호 없이 순서대로 필드의 이름을 나열한다.(hwpFieldPlain)
                    - 1: 필드이름 뒤에 일련번호가 ``{{#}}``과 같은 형식으로 붙는다.(hwpFieldNumber)
                    - 2: 필드이름 뒤에 그 필드의 개수가 ``{{#}}``과 같은 형식으로 붙는다.(hwpFieldCount)

            option:
                다음과 같은 옵션을 조합할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다.

                    - 0x01: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere과는 함께 지정할 수 없다.(hwpFieldCell)
                    - 0x02: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)
                    - 0x04: 선택된 내용 안에 존재하는 필드 리스트를 구한다.(HwpFieldSelection)

        Returns:
            각 필드 사이를 문자코드 0x02로 구분하여 다음과 같은 형식으로 리턴 한다. (가장 마지막 필드에는 0x02가 붙지 않는다.)
            ``"필드이름#1\\x02필드이름#2\\x02...필드이름#n"``
        """
        return self.hwp.GetFieldList(Number=number, option=option)

    def get_field_text(self, field: str | list | tuple | set, idx: int = 0) -> str:
        """
        지정한 필드에서 문자열을 구한다.

        Args:
            field:
                텍스트를 구할 필드 이름의 리스트.
                다음과 같이 필드 사이를 문자 코드 0x02로 구분하여
                한 번에 여러 개의 필드를 지정할 수 있다.
                "필드이름#1\\x02필드이름#2\\x02...필드이름#n"
                지정한 필드 이름이 문서 중에 두 개 이상 존재할 때의 표현 방식은 다음과 같다.
                "필드이름": 이름의 필드 중 첫 번째
                "필드이름{{n}}": 지정한 이름의 필드 중 n 번째
                예를 들어 "제목{{1}}\\x02본문\\x02이름{{0}}" 과 같이 지정하면
                '제목'이라는 이름의 필드 중 두 번째,
                '본문'이라는 이름의 필드 중 첫 번째,
                '이름'이라는 이름의 필드 중 첫 번째를 각각 지정한다.
                즉, '필드이름'과 '필드이름{{0}}'은 동일한 의미로 해석된다.
            idx:
                특정 필드가 여러 개이고, 각각의 값이 다를 때, ``field{{n}}`` 대신 ``hwp.get_field_text(field, idx=n)``라고 작성할 수 있다.

        Returns:
            텍스트 데이터가 돌아온다.
            텍스트에서 탭은 '\\t'(0x9),
            문단 바뀜은 CR/LF(0x0D/0x0A == \\r\\n)로 표현되며,
            이외의 특수 코드는 포함되지 않는다.
            필드 텍스트의 끝은 0x02(\\x02)로 표현되며,
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

    def GetFieldText(self, field: str | list | tuple | set, idx: int = 0) -> str:
        """
        지정한 필드에서 문자열을 구한다.

        Args:
            field:
                텍스트를 구할 필드 이름의 리스트.
                다음과 같이 필드 사이를 문자 코드 0x02로 구분하여
                한 번에 여러 개의 필드를 지정할 수 있다.
                "필드이름#1\\x02필드이름#2\\x02...필드이름#n"
                지정한 필드 이름이 문서 중에 두 개 이상 존재할 때의 표현 방식은 다음과 같다.
                "필드이름": 이름의 필드 중 첫 번째
                "필드이름{{n}}": 지정한 이름의 필드 중 n 번째
                예를 들어 "제목{{1}}\\x02본문\\x02이름{{0}}" 과 같이 지정하면
                '제목'이라는 이름의 필드 중 두 번째,
                '본문'이라는 이름의 필드 중 첫 번째,
                '이름'이라는 이름의 필드 중 첫 번째를 각각 지정한다.
                즉, '필드이름'과 '필드이름{{0}}'은 동일한 의미로 해석된다.
            idx:
                특정 필드가 여러 개이고, 각각의 값이 다를 때, ``field{{n}}`` 대신 ``hwp.GetFieldText(field, idx=n)``라고 작성할 수 있다.

        Returns:
            텍스트 데이터가 돌아온다.
            텍스트에서 탭은 '\\t'(0x9),
            문단 바뀜은 CR/LF(0x0D/0x0A == \\r\\n)로 표현되며,
            이외의 특수 코드는 포함되지 않는다.
            필드 텍스트의 끝은 0x02(\\x02)로 표현되며,
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

    def get_file_info(self, filename: str) -> "Hwp.HParameterSet":
        """
        파일 정보를 알아낸다.

        한글 문서를 열기 전에 암호가 걸린 문서인지 확인할 목적으로 만들어졌다.
        (현재 한/글2022 기준으로 hwpx포맷에 대해서는 파일정보를 파악할 수 없다.)

        Args:
            filename: 정보를 구하고자 하는 hwp 파일의 전체 경로

        Returns:
            "FileInfo" ParameterSet이 반환된다. 파라미터셋의 ItemID는 아래와 같다.

                - Format(string) : 파일의 형식.(HWP : 한/글 파일, UNKNOWN : 알 수 없음.)
                - VersionStr(string) : 파일의 버전 문자열. ex)5.0.0.3
                - VersionNum(unsigned int) : 파일의 버전. ex) 0x05000003
                - Encrypted(int) : 암호 여부. 현재는 파일 버전 3.0.0.0 이후 문서-한/글97, 한/글 워디안 및 한/글 2002 이상의 버전-에 대해서만 판단한다. (-1: 판단할 수 없음, 0: 암호가 걸려 있지 않음, 양수: 암호가 걸려 있음.)

        Examples:
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

    def GetFileInfo(self, filename:str) -> "Hwp.HParameterSet":
        """
        파일 정보를 알아낸다.

        한글 문서를 열기 전에 암호가 걸린 문서인지 확인할 목적으로 만들어졌다.
        (현재 한/글2022 기준으로 hwpx포맷에 대해서는 파일정보를 파악할 수 없다.)

        Args:
            filename: 정보를 구하고자 하는 hwp 파일의 전체 경로

        Returns:
            "FileInfo" ParameterSet이 반환된다. 파라미터셋의 ItemID는 아래와 같다.

                - Format(string) : 파일의 형식.(HWP : 한/글 파일, UNKNOWN : 알 수 없음.)
                - VersionStr(string) : 파일의 버전 문자열. ex)5.0.0.3
                - VersionNum(unsigned int) : 파일의 버전. ex) 0x05000003
                - Encrypted(int) : 암호 여부. 현재는 파일 버전 3.0.0.0 이후 문서-한/글97, 한/글 워디안 및 한/글 2002 이상의 버전-에 대해서만 판단한다. (-1: 판단할 수 없음, 0: 암호가 걸려 있지 않음, 양수: 암호가 걸려 있음.)

        Examples:
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

    def get_font_list(self, langid: str = "") -> list[str]:
        """
        현재 문서에 사용되고 있는 폰트 목록 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_font_list()
            ['D2Coding,R', 'Pretendard Variable Thin,R', '나눔명조,R', '함초롬바탕,R']
        """
        self.scan_font()
        return [i.rsplit(",", maxsplit=1)[0] for i in self.hwp.GetFontList(langid=langid).split("\x02")]

    def GetFontList(self, langid: str = "") -> list:
        """
        현재 문서에 사용되고 있는 폰트 목록 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_font_list()
            ['D2Coding,R', 'Pretendard Variable Thin,R', '나눔명조,R', '함초롬바탕,R']
        """
        self.scan_font()
        return self.hwp.GetFontList(langid=langid)

    def get_heading_string(self) -> str:
        """
        현재 커서가 위치한 문단 시작부분의 글머리표/문단번호/개요번호를 추출한다.

        글머리표/문단번호/개요번호가 있는 경우, 해당 문자열을 얻어올 수 있다.
        문단에 글머리표/문단번호/개요번호가 없는 경우, 빈 문자열이 추출된다.

        Returns:
            (글머리표/문단번호/개요번호가 있다면) 해당 문자열이 반환된다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_heading_string()
            '1.'
        """
        return self.hwp.GetHeadingString()

    def GetHeadingString(self) -> str:
        """
        현재 커서가 위치한 문단의 글머리표/문단번호/개요번호를 추출한다.

        글머리표/문단번호/개요번호가 있는 경우, 해당 문자열을 얻어올 수 있다.
        문단에 글머리표/문단번호/개요번호가 없는 경우, 빈 문자열이 추출된다.

        Returns:
            (글머리표/문단번호/개요번호가 있다면) 해당 문자열이 반환된다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.GetHeadingString()
            '1.'
        """
        return self.hwp.GetHeadingString()

    def get_message_box_mode(self) -> int:
        """
        현재 메시지 박스의 Mode를 ``int``로 얻어온다.

        set_message_box_mode와 함께 쓰인다. 6개의 대화상자에서 각각 확인/취소/종료/재시도/무시/예/아니오 버튼을 자동으로 선택할 수 있게 설정할 수 있으며 조합 가능하다.
        리턴하는 정수의 의미는 ``set_message_box_mode``를 참고한다.
        """
        return self.hwp.GetMessageBoxMode()

    def GetMessageBoxMode(self):
        """
        현재 메시지 박스의 Mode를 int로 얻어온다.

        set_message_box_mode와 함께 쓰인다.
        6개의 대화상자에서 각각 확인/취소/종료/재시도/무시/예/아니오 버튼을
        자동으로 선택할 수 있게 설정할 수 있으며 조합 가능하다.
        리턴하는 정수의 의미는 ``set_message_box_mode``를 참고한다.
        """
        return self.hwp.GetMessageBoxMode()

    def get_metatag_list(self, number, option):
        """메타태그리스트 가져오기"""
        return self.hwp.GetMetatagList(Number=number, option=option)

    def GetMetatagList(self, number, option):
        """메타태그리스트 가져오기"""
        return self.hwp.GetMetatagList(Number=number, option=option)

    def get_metatag_name_text(self, tag):
        """메타태그이름 문자열 가져오기"""
        return self.hwp.GetMetatagNameText(tag=tag)

    def GetMetatagNameText(self, tag):
        """메타태그이름 문자열 가져오기"""
        return self.hwp.GetMetatagNameText(tag=tag)

    def get_mouse_pos(self, x_rel_to:int=1, y_rel_to:int=1) -> "Hwp.HParameterSet":
        """
        마우스의 현재 위치를 얻어온다.

        단위가 HWPUNIT임을 주의해야 한다.
        (1 inch = 7200 HWPUNIT, 1mm = 283.465 HWPUNIT)

        Args:
            x_rel_to: X좌표계의 기준 위치(기본값은 1:쪽기준)

                - 0: 종이 기준으로 좌표를 가져온다.
                - 1: 쪽 기준으로 좌표를 가져온다.

            y_rel_to: Y좌표계의 기준 위치(기본값은 1:쪽기준)

                - 0: 종이 기준으로 좌표를 가져온다.
                - 1: 쪽 기준으로 좌표를 가져온다.

        Returns:
            "MousePos" ParameterSet이 반환된다.
            아이템ID는 아래와 같다.

                - XRelTo(unsigned long): 가로 상대적 기준(0: 종이, 1: 쪽)
                - YRelTo(unsigned long): 세로 상대적 기준(0: 종이, 1: 쪽)
                - Page(unsigned long): 페이지 번호(0-based)
                - X(long): 가로 클릭한 위치(HWPUNIT)
                - Y(long): 세로 클릭한 위치(HWPUNIT)

        Examples:
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

    def GetMousePos(self, x_rel_to:int=1, y_rel_to:int=1) -> "Hwp.HParameterSet":
        """
        마우스의 현재 위치를 얻어온다.

        단위가 HWPUNIT임을 주의해야 한다.
        (1 inch = 7200 HWPUNIT, 1mm = 283.465 HWPUNIT)

        Args:
            x_rel_to: X좌표계의 기준 위치(기본값은 1:쪽기준)

                - 0: 종이 기준으로 좌표를 가져온다.
                - 1: 쪽 기준으로 좌표를 가져온다.

            y_rel_to: Y좌표계의 기준 위치(기본값은 1:쪽기준)

                - 0: 종이 기준으로 좌표를 가져온다.
                - 1: 쪽 기준으로 좌표를 가져온다.

        Returns:
            "MousePos" ParameterSet이 반환된다.
            아이템ID는 아래와 같다.

                - XRelTo(unsigned long): 가로 상대적 기준(0: 종이, 1: 쪽)
                - YRelTo(unsigned long): 세로 상대적 기준(0: 종이, 1: 쪽)
                - Page(unsigned long): 페이지 번호(0-based)
                - X(long): 가로 클릭한 위치(HWPUNIT)
                - Y(long): 세로 클릭한 위치(HWPUNIT)

        Examples:
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

        Args:
            pgno: 텍스트를 추출 할 페이지의 번호(0부터 시작)

            option: 추출 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다. 생략(또는 0xffffffff)하면 모든 텍스트를 추출한다.

                - 0x00: 본문 텍스트만 추출한다.(maskNormal)
                - 0x01: 표에대한 텍스트를 추출한다.(maskTable)
                - 0x02: 글상자 텍스트를 추출한다.(maskTextbox)
                - 0x04: 캡션 텍스트를 추출한다. (표, ShapeObject)(maskCaption)

        Returns:
            해당 페이지의 텍스트가 추출된다.
            글머리는 추출하지만, 표번호는 추출하지 못한다.
        """
        return self.hwp.GetPageText(pgno=pgno, option=option)

    def GetPageText(self, pgno: int = 0, option: hex = 0xffffffff) -> str:
        """
        페이지 단위의 텍스트 추출.

        일반 텍스트(글자처럼 취급 도형 포함)를 우선적으로 추출하고,
        도형(표, 글상자) 내의 텍스트를 추출한다.

            - 팁1: get_text로는 글머리를 추출하지 않지만, get_page_text는 추출한다.
            - 팁2: 아무리 get_page_text라도 유일하게 표번호는 추출하지 못한다. 표번호는 XML태그 속성 안에 저장되기 때문이다.

        Args:
            pgno: 텍스트를 추출 할 페이지의 번호(0부터 시작)

            option: 추출 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다. 생략(또는 0xffffffff)하면 모든 텍스트를 추출한다.

                - 0x00: 본문 텍스트만 추출한다.(maskNormal)
                - 0x01: 표에대한 텍스트를 추출한다.(maskTable)
                - 0x02: 글상자 텍스트를 추출한다.(maskTextbox)
                - 0x04: 캡션 텍스트를 추출한다. (표, ShapeObject)(maskCaption)

        Returns:
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

        Returns:
            (List, Para, Pos) 튜플.

                - List: 캐럿이 위치한 문서 내 list ID(본문이 0)
                - Para: 캐럿이 위치한 문단 ID(0부터 시작)
                - Pos: 캐럿이 위치한 문단 내 글자 위치(0부터 시작)

        """
        return self.hwp.GetPos()

    def GetPos(self) -> tuple[int]:
        """
        캐럿의 위치를 얻어온다.

        파라미터 중 리스트는, 문단과 컨트롤들이 연결된 한/글 문서 내 구조를 뜻한다.
        리스트 아이디는 문서 내 위치 정보 중 하나로서 SelectText에 넘겨줄 때 사용한다.
        (파이썬 자료형인 list가 아님)

        Returns:
            (List, Para, Pos) 튜플.

                - List: 캐럿이 위치한 문서 내 list ID(본문이 0)
                - Para: 캐럿이 위치한 문단 ID(0부터 시작)
                - Pos: 캐럿이 위치한 문단 내 글자 위치(0부터 시작)

        """
        return self.hwp.GetPos()

    def get_pos_by_set(self) -> "Hwp.HParameterSet":
        """
        현재 캐럿의 위치 정보를 ParameterSet으로 얻어온다.

        해당 파라미터셋은 set_pos_by_set에 직접 집어넣을 수 있어 간편히 사용할 수 있다.

        Returns:
            캐럿 위치에 대한 ParameterSet
            해당 파라미터셋의 아이템은 아래와 같다.
            "List": 캐럿이 위치한 문서 내 list ID(본문이 0)
            "Para": 캐럿이 위치한 문단 ID(0부터 시작)
            "Pos": 캐럿이 위치한 문단 내 글자 위치(0부터 시작)

        Examples:
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

    def GetPosBySet(self) -> "Hwp.HParameterSet":
        """
        현재 캐럿의 위치 정보를 ParameterSet으로 얻어온다.

        해당 파라미터셋은 set_pos_by_set에 직접 집어넣을 수 있어 간편히 사용할 수 있다.

        Returns:
            캐럿 위치에 대한 ParameterSet. 해당 파라미터셋의 아이템은 아래와 같다.

                - "List": 캐럿이 위치한 문서 내 list ID(본문이 0)
                - "Para": 캐럿이 위치한 문단 ID(0부터 시작)
                - "Pos": 캐럿이 위치한 문단 내 글자 위치(0부터 시작)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> pset = hwp.GetPosBySet()  # 캐럿위치 저장
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

        형태로 비어있는 상태이며, OnDocument_New와 OnDocument_Open 두 개의 함수에 한해서만 코드를 추가하고 실행할 수 있다.

        Args:
            filename: 매크로 소스를 가져올 한/글 문서의 전체경로

        Returns:
            (문서에 포함된) 스크립트의 소스코드

        Examples:
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

        Args:
            filename: 매크로 소스를 가져올 한/글 문서의 전체경로

        Returns:
            (문서에 포함된) 스크립트의 소스코드

        Examples:
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

    def get_selected_pos(self) -> tuple[bool, str, str, int, str, str, int]:
        """
        현재 설정된 블록의 위치정보를 얻어온다.

        Returns:
            블록상태여부, 시작과 끝위치 인덱스인 6개 정수 등 7개 요소의 튜플을 리턴(is_block, slist, spara, spos, elist, epara, epos)

                - is_block: 현재 블록선택상태 여부(블록상태이면 True)
                - slist: 설정된 블록의 시작 리스트 아이디.
                - spara: 설정된 블록의 시작 문단 아이디.
                - spos: 설정된 블록의 문단 내 시작 글자 단위 위치.
                - elist: 설정된 블록의 끝 리스트 아이디.
                - epara: 설정된 블록의 끝 문단 아이디.
                - epos: 설정된 블록의 문단 내 끝 글자 단위 위치.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_selected_pos()
            (True, 0, 0, 16, 0, 7, 16)
        """
        return self.hwp.GetSelectedPos()

    def GetSelectedPos(self) -> tuple[bool, str, str, int, str, str, int]:
        """
        현재 설정된 블록의 위치정보를 얻어온다.

        Returns:
            블록상태여부, 시작과 끝위치 인덱스인 6개 정수 등 7개 요소의 튜플을 리턴(is_block, slist, spara, spos, elist, epara, epos)

                - is_block: 현재 블록선택상태 여부(블록상태이면 True)
                - slist: 설정된 블록의 시작 리스트 아이디.
                - spara: 설정된 블록의 시작 문단 아이디.
                - spos: 설정된 블록의 문단 내 시작 글자 단위 위치.
                - elist: 설정된 블록의 끝 리스트 아이디.
                - epara: 설정된 블록의 끝 문단 아이디.
                - epos: 설정된 블록의 문단 내 끝 글자 단위 위치.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.GetSelectedPos()
            (True, 0, 0, 16, 0, 7, 16)
        """
        return self.hwp.GetSelectedPos()

    def get_selected_pos_by_set(self, sset:Any, eset:Any) -> bool:
        """
        현재 설정된 블록의 위치정보를 얻어온다.

        (GetSelectedPos의 ParameterSet버전)
        실행 전 GetPos 형태의 파라미터셋 두 개를 미리 만들어서
        인자로 넣어줘야 한다.

        Args:
            sset: 설정된 블록의 시작 파라메터셋 (ListParaPos)
            eset: 설정된 블록의 끝 파라메터셋 (ListParaPos)

        Returns:
            성공하면 True, 실패하면 False. (실행시 sset과 eset의 아이템 값이 업데이트된다.)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 선택할 블록의 시작 위치에 캐럿을 둔 상태로
            >>> sset = hwp.get_pos_by_set()
            >>> # 블록 끝이 될 부분으로 이동한 후에
            >>> eset = hwp.get_pos_by_set()
            >>> hwp.get_selected_pos_by_set(sset, eset)
            >>> hwp.set_pos_by_set(eset)
            True
        """
        return self.hwp.GetSelectedPosBySet(sset=sset, eset=eset)

    def GetSelectedPosBySet(self, sset:"Hwp.HParameterSet", eset:"Hwp.HParameterSet") -> bool:
        """
        현재 설정된 블록의 위치정보를 얻어온다.

        (GetSelectedPos의 ParameterSet버전)
        실행 전 GetPos 형태의 파라미터셋 두 개를 미리 만들어서
        인자로 넣어줘야 한다.

        Args:
            sset: 설정된 블록의 시작 파라메터셋 (ListParaPos)
            eset: 설정된 블록의 끝 파라메터셋 (ListParaPos)

        Returns:
            성공하면 True, 실패하면 False. (실행시 sset과 eset의 아이템 값이 업데이트된다.)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 선택할 블록의 시작 위치에 캐럿을 둔 상태로
            >>> sset = hwp.get_pos_by_set()
            >>> # 블록 끝이 될 부분으로 이동한 후에
            >>> eset = hwp.get_pos_by_set()
            >>> hwp.GetSelectedPosBySet(sset, eset)
            >>> hwp.set_pos_by_set(eset)
            True
        """
        return self.hwp.GetSelectedPosBySet(sset=sset, eset=eset)

    def get_text(self) -> tuple[int, str]:
        """
        문서 내에서 텍스트를 얻어온다.

        줄바꿈 기준으로 텍스트를 얻어오므로 반복실행해야 한다.
        get_text()의 사용이 끝나면 release_scan()을 반드시 호출하여
        관련 정보를 초기화 해주어야 한다.
        get_text()로 추출한 텍스트가 있는 문단으로 캐럿을 이동 시키려면
        move_pos(201)을 실행하면 된다.

        Returns:
            (state: int, text: str) 형태의 튜플을 리턴한다. text는 추출한 텍스트 데이터이다. 텍스트에서 탭은 '\\t'(0x9), 문단 바뀜은 '\\r\\n'(0x0D/0x0A)로 표현되며, 이외의 특수 코드는 포함되지 않는다.

            state의 의미는 아래와 같다.

                - 0: 텍스트 정보 없음
                - 1: 리스트의 끝
                - 2: 일반 텍스트
                - 3: 다음 문단
                - 4: 제어문자 내부로 들어감
                - 5: 제어문자를 빠져나옴
                - 101: 초기화 안 됨(init_scan() 실패 또는 init_scan()을 실행하지 않은 경우)
                - 102: 텍스트 변환 실패

        Examples:
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

    def GetText(self) -> tuple[int, str]:
        """
        문서 내에서 텍스트를 얻어온다.

        줄바꿈 기준으로 텍스트를 얻어오므로 반복실행해야 한다.
        get_text()의 사용이 끝나면 release_scan()을 반드시 호출하여
        관련 정보를 초기화 해주어야 한다.
        get_text()로 추출한 텍스트가 있는 문단으로 캐럿을 이동 시키려면
        move_pos(201)을 실행하면 된다.

        Returns:
            (state: int, text: str) 형태의 튜플을 리턴한다. text는 추출한 텍스트 데이터이다. 텍스트에서 탭은 '\\t'(0x9), 문단 바뀜은 '\\r\\n'(0x0D/0x0A)로 표현되며, 이외의 특수 코드는 포함되지 않는다.
            state의 의미는 아래와 같다.

                - 0: 텍스트 정보 없음
                - 1: 리스트의 끝
                - 2: 일반 텍스트
                - 3: 다음 문단
                - 4: 제어문자 내부로 들어감
                - 5: 제어문자를 빠져나옴
                - 101: 초기화 안 됨(init_scan() 실패 또는 init_scan()을 실행하지 않은 경우)
                - 102: 텍스트 변환 실패

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.InitScan()
            >>> while True:
            ...     state, text = hwp.GetText()
            ...     print(state, text)
            ...     if state <= 1:
            ...         break
            ... hwp.ReleaseScan()
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

    def get_text_file(self, format:Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"]="UNICODE", option:str="saveblock:true") -> str:
        """
        현재 열린 문서 전체 또는 선택한 범위를 문자열로 리턴한다.

        이 함수는 JScript나 VBScript와 같이
        직접적으로 local disk를 접근하기 힘든 언어를 위해 만들어졌으므로
        disk를 접근할 수 있는 언어에서는 사용하지 않기를 권장.
        disk를 접근할 수 있다면, Save나 SaveBlockAction을 사용할 것.
        이 함수 역시 내부적으로는 save나 SaveBlockAction을 호출하도록 되어있고
        텍스트로 저장된 파일이 메모리에서 3~4번 복사되기 때문에 느리고, 메모리를 낭비함.

        팁1: ``hwp.Copy()``, ``hwp.Paste()`` 대신 get_text_file/set_text_file을 사용하기 추천.

        팁2: ``format="HTML"``로 추출시 표번호가 유지된다.

        Args:
            format:
                파일의 형식. 기본값은 "UNICODE". 내부적으로 str.lower() 메서드가 포함되어 있으므로 소문자로 입력해도 된다.

                    - "HWP": HWP native format, BASE64로 인코딩되어 있다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다.ver:0x0505010B
                    - "HWPML2X": HWP 형식과 호환. 문서의 모든 정보를 유지
                    - "HTML": 인터넷 문서 HTML 형식. 한/글 고유의 서식은 손실된다.
                    - "UNICODE": 유니코드 텍스트, 서식정보가 없는 텍스트만 저장.
                    - "TEXT": 일반 텍스트. 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다.

            option:
                option 파라미터에 "saveblock"을 지정하면 선택된 블록만 저장한다. 개체 선택 상태에서는 동작하지 않는다.

        Returns:
            지정된 포맷에 맞춰 파일을 문자열로 변환한 값을 반환한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_text_file()
            'ㅁㄴㅇㄹ\\r\\nㅁㄴㅇㄹ\\r\\nㅁㄴㅇㄹ\\r\\n\\r\\nㅂㅈㄷㄱ\\r\\nㅂㅈㄷㄱ\\r\\nㅂㅈㄷㄱ\\r\\n'
        """
        return self.hwp.GetTextFile(Format=format, option=option)

    def GetTextFile(self, format:Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"]="UNICODE", option:str="saveblock:true") -> str:
        """
        현재 열린 문서 전체 또는 선택한 범위를 문자열로 리턴한다.

        이 함수는 JScript나 VBScript와 같이
        직접적으로 local disk를 접근하기 힘든 언어를 위해 만들어졌으므로
        disk를 접근할 수 있는 언어에서는 사용하지 않기를 권장.
        disk를 접근할 수 있다면, Save나 SaveBlockAction을 사용할 것.
        이 함수 역시 내부적으로는 save나 SaveBlockAction을 호출하도록 되어있고
        텍스트로 저장된 파일이 메모리에서 3~4번 복사되기 때문에 느리고, 메모리를 낭비함.

        팁1: ``hwp.Copy()``, ``hwp.Paste()`` 대신 get_text_file/set_text_file을 사용하기 추천.

        팁2: ``format="HTML"``로 추출시 표번호가 유지된다.

        Args:
            format:
                파일의 형식. 기본값은 "UNICODE". 내부적으로 str.lower() 메서드가 포함되어 있으므로 소문자로 입력해도 된다.

                    - "HWP": HWP native format, BASE64로 인코딩되어 있다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다.ver:0x0505010B
                    - "HWPML2X": HWP 형식과 호환. 문서의 모든 정보를 유지
                    - "HTML": 인터넷 문서 HTML 형식. 한/글 고유의 서식은 손실된다.
                    - "UNICODE": 유니코드 텍스트, 서식정보가 없는 텍스트만 저장.
                    - "TEXT": 일반 텍스트. 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다.

            option:
                option 파라미터에 "saveblock"을 지정하면 선택된 블록만 저장한다. 개체 선택 상태에서는 동작하지 않는다.

        Returns:
            지정된 포맷에 맞춰 파일을 문자열로 변환한 값을 반환한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.GetTextFile()
            'ㅁㄴㅇㄹ\\r\\nㅁㄴㅇㄹ\\r\\nㅁㄴㅇㄹ\\r\\n\\r\\nㅂㅈㄷㄱ\\r\\nㅂㅈㄷㄱ\\r\\nㅂㅈㄷㄱ\\r\\n'
        """
        return self.hwp.GetTextFile(Format=format, option=option)

    def import_style(self, sty_filepath: str) -> bool:
        """
        미리 저장된 특정 sty파일의 스타일을 임포트한다.

        Args:
            sty_filepath: sty파일의 경로

        Returns:
            성공시 True, 실패시 False

        Examples:
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

    def ImportStyle(self, sty_filepath: str) -> bool:
        """
        미리 저장된 특정 sty파일의 스타일을 임포트한다.

        Args:
            sty_filepath: sty파일의 경로

        Returns:
            성공시 True, 실패시 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.ImportStyle("C:/Users/User/Desktop/new_style.sty")
            True
        """
        if sty_filepath.lower()[1] != ":":
            sty_filepath = os.path.join(os.getcwd(), sty_filepath)

        style_set = self.hwp.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.hwp.ImportStyle(style_set.HSet)

    def InitHParameterSet(self):
        return self.hwp.InitHParameterSet()

    def init_scan(self, option:int=0x07, range:int=0x77, spara:int=0, spos:int=0, epara:int=-1, epos:int=-1) -> bool:
        """
        문서의 내용을 검색하기 위해 초기설정을 한다.

        문서의 검색 과정은 InitScan()으로 검색위한 준비 작업을 하고
        GetText()를 호출하여 본문의 텍스트를 얻어온다.
        GetText()를 반복호출하면 연속하여 본문의 텍스트를 얻어올 수 있다.
        검색이 끝나면 ReleaseScan()을 호출하여 관련 정보를 Release해야 한다.

        Args:
            option:
                기본값은 0x7(모든 컨트롤 대상) 찾을 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다.
                생략하면 모든 컨트롤을 찾을 대상으로 한다.

                    - 0x00: 본문을 대상으로 검색한다.(서브리스트를 검색하지 않는다.) - maskNormal
                    - 0x01: char 타입 컨트롤 마스크를 대상으로 한다.(강제줄나눔, 문단 끝, 하이픈, 묶움빈칸, 고정폭빈칸, 등...) - maskChar
                    - 0x02: inline 타입 컨트롤 마스크를 대상으로 한다.(누름틀 필드 끝, 등...) - maskInline
                    - 0x04: extende 타입 컨트롤 마스크를 대상으로 한다.(바탕쪽, 프레젠테이션, 다단, 누름틀 필드 시작, Shape Object, 머리말, 꼬리말, 각주, 미주, 번호관련 컨트롤, 새 번호 관련 컨트롤, 감추기, 찾아보기, 글자 겹침, 등...) - maskCtrl

            range:
                검색의 범위를 다음과 같은 옵션을 조합(sum)하여 지정할 수 있다.
                생략하면 "문서 시작부터 - 문서의 끝까지" 검색 범위가 지정된다.

                    - 0x0000: 캐럿 위치부터. (시작 위치) - scanSposCurrent
                    - 0x0010: 특정 위치부터. (시작 위치) - scanSposSpecified
                    - 0x0020: 줄의 시작부터. (시작 위치) - scanSposLine
                    - 0x0030: 문단의 시작부터. (시작 위치) - scanSposParagraph
                    - 0x0040: 구역의 시작부터. (시작 위치) - scanSposSection
                    - 0x0050: 리스트의 시작부터. (시작 위치) - scanSposList
                    - 0x0060: 컨트롤의 시작부터. (시작 위치) - scanSposControl
                    - 0x0070: 문서의 시작부터. (시작 위치) - scanSposDocument
                    - 0x0000: 캐럿 위치까지. (끝 위치) - scanEposCurrent
                    - 0x0001: 특정 위치까지. (끝 위치) - scanEposSpecified
                    - 0x0002: 줄의 끝까지. (끝 위치) - scanEposLine
                    - 0x0003: 문단의 끝까지. (끝 위치) - scanEposParagraph
                    - 0x0004: 구역의 끝까지. (끝 위치) - scanEposSection
                    - 0x0005: 리스트의 끝까지. (끝 위치) - scanEposList
                    - 0x0006: 컨트롤의 끝까지. (끝 위치) - scanEposControl
                    - 0x0007: 문서의 끝까지. (끝 위치) - scanEposDocument
                    - 0x00ff: 검색의 범위를 블록으로 제한. - scanWithinSelection
                    - 0x0000: 정뱡향. (검색 방향) - scanForward
                    - 0x0100: 역방향. (검색 방향) - scanBackward

            spara: 검색 시작 위치의 문단 번호. scanSposSpecified 옵션이 지정되었을 때만 유효하다. (예: range=0x0011)
            spos: 검색 시작 위치의 문단 중에서 문자의 위치. scanSposSpecified 옵션이 지정되었을 때만 유효하다. (예: range=0x0011)
            epara: 검색 끝 위치의 문단 번호. scanEposSpecified 옵션이 지정되었을 때만 유효하다.
            epos: 검색 끝 위치의 문단 중에서 문자의 위치. scanEposSpecified 옵션이 지정되었을 때만 유효하다.

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.init_scan(range=0xff)
            >>> _, text = hwp.get_text()
            >>> hwp.release_scan()
            >>> print(text)
            Hello, world!
        """
        return self.hwp.InitScan(option=option, Range=range, spara=spara, spos=spos, epara=epara, epos=epos)

    def InitScan(self, option:int=0x07, range:int=0x77, spara:int=0, spos:int=0, epara:int=-1, epos:int=-1) -> bool:
        """
        문서의 내용을 검색하기 위해 초기설정을 한다.

        문서의 검색 과정은 InitScan()으로 검색위한 준비 작업을 하고
        GetText()를 호출하여 본문의 텍스트를 얻어온다.
        GetText()를 반복호출하면 연속하여 본문의 텍스트를 얻어올 수 있다.
        검색이 끝나면 ReleaseScan()을 호출하여 관련 정보를 Release해야 한다.

        Args:
            option:
                기본값은 0x7(모든 컨트롤 대상) 찾을 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다.
                생략하면 모든 컨트롤을 찾을 대상으로 한다.

                    - 0x00: 본문을 대상으로 검색한다.(서브리스트를 검색하지 않는다.) - maskNormal
                    - 0x01: char 타입 컨트롤 마스크를 대상으로 한다.(강제줄나눔, 문단 끝, 하이픈, 묶움빈칸, 고정폭빈칸, 등...) - maskChar
                    - 0x02: inline 타입 컨트롤 마스크를 대상으로 한다.(누름틀 필드 끝, 등...) - maskInline
                    - 0x04: extende 타입 컨트롤 마스크를 대상으로 한다.(바탕쪽, 프레젠테이션, 다단, 누름틀 필드 시작, Shape Object, 머리말, 꼬리말, 각주, 미주, 번호관련 컨트롤, 새 번호 관련 컨트롤, 감추기, 찾아보기, 글자 겹침, 등...) - maskCtrl

            range:
                검색의 범위를 다음과 같은 옵션을 조합(sum)하여 지정할 수 있다.
                생략하면 "문서 시작부터 - 문서의 끝까지" 검색 범위가 지정된다.

                    - 0x0000: 캐럿 위치부터. (시작 위치) - scanSposCurrent
                    - 0x0010: 특정 위치부터. (시작 위치) - scanSposSpecified
                    - 0x0020: 줄의 시작부터. (시작 위치) - scanSposLine
                    - 0x0030: 문단의 시작부터. (시작 위치) - scanSposParagraph
                    - 0x0040: 구역의 시작부터. (시작 위치) - scanSposSection
                    - 0x0050: 리스트의 시작부터. (시작 위치) - scanSposList
                    - 0x0060: 컨트롤의 시작부터. (시작 위치) - scanSposControl
                    - 0x0070: 문서의 시작부터. (시작 위치) - scanSposDocument
                    - 0x0000: 캐럿 위치까지. (끝 위치) - scanEposCurrent
                    - 0x0001: 특정 위치까지. (끝 위치) - scanEposSpecified
                    - 0x0002: 줄의 끝까지. (끝 위치) - scanEposLine
                    - 0x0003: 문단의 끝까지. (끝 위치) - scanEposParagraph
                    - 0x0004: 구역의 끝까지. (끝 위치) - scanEposSection
                    - 0x0005: 리스트의 끝까지. (끝 위치) - scanEposList
                    - 0x0006: 컨트롤의 끝까지. (끝 위치) - scanEposControl
                    - 0x0007: 문서의 끝까지. (끝 위치) - scanEposDocument
                    - 0x00ff: 검색의 범위를 블록으로 제한. - scanWithinSelection
                    - 0x0000: 정뱡향. (검색 방향) - scanForward
                    - 0x0100: 역방향. (검색 방향) - scanBackward

            spara: 검색 시작 위치의 문단 번호. scanSposSpecified 옵션이 지정되었을 때만 유효하다. (예: range=0x0011)
            spos: 검색 시작 위치의 문단 중에서 문자의 위치. scanSposSpecified 옵션이 지정되었을 때만 유효하다. (예: range=0x0011)
            epara: 검색 끝 위치의 문단 번호. scanEposSpecified 옵션이 지정되었을 때만 유효하다.
            epos: 검색 끝 위치의 문단 중에서 문자의 위치. scanEposSpecified 옵션이 지정되었을 때만 유효하다.

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.init_scan(range=0xff)
            >>> _, text = hwp.get_text()
            >>> hwp.release_scan()
            >>> print(text)
            Hello, world!
        """
        return self.hwp.InitScan(option=option, Range=range, spara=spara, spos=spos, epara=epara, epos=epos)

    def insert(self, path: str, format: str = "", arg: str = "", move_doc_end: bool = False) -> bool:
        """
        현재 캐럿 위치에 문서파일을 삽입한다.

        `format`, `arg` 파라미터에 대한 자세한 설명은 `open` 참조

        Args:
            path: 문서파일의 경로
            format:
                문서형식. **빈 문자열을 지정하면 자동으로 선택한다.**
                생략하면 빈 문자열이 지정된다.
                아래에 쓰여 있는 대로 대문자로만 써야 한다.

                    - "HWPX": 한/글 hwpx format
                    - "HWP": 한/글 native format
                    - "HWP30": 한/글 3.X/96/97
                    - "HTML": 인터넷 문서
                    - "TEXT": 아스키 텍스트 문서
                    - "UNICODE": 유니코드 텍스트 문서
                    - "HWP20": 한글 2.0
                    - "HWP21": 한글 2.1/2.5
                    - "HWP15": 한글 1.X
                    - "HWPML1X": HWPML 1.X 문서 (Open만 가능)
                    - "HWPML2X": HWPML 2.X 문서 (Open / SaveAs 가능)
                    - "RTF": 서식 있는 텍스트 문서
                    - "DBF": DBASE II/III 문서
                    - "HUNMIN": 훈민정음 3.0/2000
                    - "MSWORD": 마이크로소프트 워드 문서
                    - "DOCRTF": MS 워드 문서 (doc)
                    - "OOXML": MS 워드 문서 (docx)
                    - "HANA": 하나워드 문서
                    - "ARIRANG": 아리랑 문서
                    - "ICHITARO": 一太郞 문서 (일본 워드프로세서)
                    - "WPS": WPS 문서
                    - "DOCIMG": 인터넷 프레젠테이션 문서(SaveAs만 가능)
                    - "SWF": Macromedia Flash 문서(SaveAs만 가능)
            arg:
                세부옵션. 의미는 format에 지정한 파일형식에 따라 다르다.
                조합 가능하며, 생략하면 빈 문자열이 지정된다.

                공통

                    - "setcurdir:FALSE;" :로드한 후 해당 파일이 존재하는 폴더로 현재 위치를 변경한다. hyperlink 정보가 상대적인 위치로 되어 있을 때 유용하다.

                HWP/HWPX

                    - "lock:TRUE;": 로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
                    - "notext:FALSE;": 텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)
                    - "template:FALSE;": 새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다. 이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.
                    - "suspendpassword:FALSE;": TRUE로 지정하면 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.
                    - "forceopen:FALSE;": TRUE로 지정하면 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다.
                    - "versionwarning:FALSE;": TRUE로 지정하면 문서가 상위버전일 경우 메시지 박스를 띄우게 된다.

                HTML

                    - "code"(string, codepage): 문서변환 시 사용되는 코드 페이지를 지정할 수 있으며 code키가 존재할 경우 필터사용 시 사용자 다이얼로그를  띄우지 않는다.
                    - (코드페이지 종류는 아래와 같다.)
                    - ("utf8" : UTF8)
                    - ("unicode": 유니코드)
                    - ("ks":  한글 KS 완성형)
                    - ("acp" : Active Codepage 현재 시스템의 코드 페이지)
                    - ("kssm": 한글 조합형)
                    - ("sjis" : 일본)
                    - ("gb" : 중국 간체)
                    - ("big5" : 중국 번체)
                    - "textunit:(string, pixel);": Export될 Text의 크기의 단위 결정.pixel, point, mili 지정 가능.
                    - "formatunit:(string, pixel);": Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능

                DOCIMG

                    - "asimg:FALSE;": 저장할 때 페이지를 image로 저장
                    - "ashtml:FALSE;": 저장할 때 페이지를 html로 저장

                TEXT

                    - "code:(string, codepage);": 문서 변환 시 사용되는 코드 페이지를 지정할 수 있으며
                    - code키가 존재할 경우 필터 사용 시 사용자 다이얼로그를  띄우지 않는다.

        Returns:
            성공하면 True, 실패하면 False
        """
        if path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)
        try:
            return self.hwp.Insert(Path=path, Format=format, arg=arg)
        finally:
            if move_doc_end:
                self.MoveDocEnd()

    def insert_background_picture(self, path:str,
                                  border_type: Literal["SelectedCell", "SelectedCellDelete"] = "SelectedCell",
                                  embedded:bool=True, filloption:int=5, effect:int=0, watermark:bool=False, brightness:int=0,
                                  contrast:int=0) -> bool:
        """
        **셀**에 배경이미지를 삽입한다.

        CellBorderFill의 SetItem 중 FillAttr 의 SetItem FileName 에
        이미지의 binary data를 지정해 줄 수가 없어서 만든 함수다.
        기타 배경에 대한 다른 조정은 Action과 ParameterSet의 조합으로 가능하다.

        Args:
            path: 삽입할 이미지 파일
            border_type:
                배경 유형을 문자열로 지정(파라미터 이름과는 다르게 삽입/제거 기능이다.)

                    - "SelectedCell": 현재 선택된 표의 셀의 배경을 변경한다.
                    - "SelectedCellDelete": 현재 선택된 표의 셀의 배경을 지운다.

                단, 배경 제거시 반드시 셀이 선택되어 있어야함.
                커서가 위치하는 것만으로는 동작하지 않음.

            embedded: 이미지 파일을 문서 내에 포함할지 여부 (True/False). 생략하면 True
            filloption:
                삽입할 그림의 크기를 지정하는 옵션

                    - 0: 바둑판식으로 - 모두
                    - 1: 바둑판식으로 - 가로/위
                    - 2: 바둑판식으로 - 가로/아로
                    - 3: 바둑판식으로 - 세로/왼쪽
                    - 4: 바둑판식으로 - 세로/오른쪽
                    - 5: 크기에 맞추어(기본값)
                    - 6: 가운데로
                    - 7: 가운데 위로
                    - 8: 가운데 아래로
                    - 9: 왼쪽 가운데로
                    - 10: 왼쪽 위로
                    - 11: 왼쪽 아래로
                    - 12: 오른쪽 가운데로
                    - 13: 오른쪽 위로
                    - 14: 오른쪽 아래로

            effect:
                이미지효과

                    - 0: 원래 그림(기본값)
                    - 1: 그레이 스케일
                    - 2: 흑백으로

            watermark: watermark효과 유무(True/False). 기본값은 False. 이 옵션이 True이면 brightness 와 contrast 옵션이 무시된다.
            brightness: 밝기 지정(-100 ~ 100), 기본 값은 0
            contrast: 선명도 지정(-100 ~ 100), 기본 값은 0

        Returns:
            성공했을 경우 True, 실패했을 경우 False

        Examples:
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

    def InsertBackgroundPicture(self, path:str, border_type: Literal["SelectedCell", "SelectedCellDelete"] = "SelectedCell",
                                embedded:bool=True, filloption:int=5, effect:int=0, watermark:bool=False, brightness:int=0,
                                contrast:int=0) -> bool:
        """
        **셀**에 배경이미지를 삽입한다.

        CellBorderFill의 SetItem 중 FillAttr 의 SetItem FileName 에
        이미지의 binary data를 지정해 줄 수가 없어서 만든 함수다.
        기타 배경에 대한 다른 조정은 Action과 ParameterSet의 조합으로 가능하다.

        Args:
            path: 삽입할 이미지 파일
            border_type:
                배경 유형을 문자열로 지정(파라미터 이름과는 다르게 삽입/제거 기능이다.)

                    - "SelectedCell": 현재 선택된 표의 셀의 배경을 변경한다.
                    - "SelectedCellDelete": 현재 선택된 표의 셀의 배경을 지운다.

                단, 배경 제거시 반드시 셀이 선택되어 있어야함.
                커서가 위치하는 것만으로는 동작하지 않음.

            embedded: 이미지 파일을 문서 내에 포함할지 여부 (True/False). 생략하면 True
            filloption:
                삽입할 그림의 크기를 지정하는 옵션

                    - 0: 바둑판식으로 - 모두
                    - 1: 바둑판식으로 - 가로/위
                    - 2: 바둑판식으로 - 가로/아로
                    - 3: 바둑판식으로 - 세로/왼쪽
                    - 4: 바둑판식으로 - 세로/오른쪽
                    - 5: 크기에 맞추어(기본값)
                    - 6: 가운데로
                    - 7: 가운데 위로
                    - 8: 가운데 아래로
                    - 9: 왼쪽 가운데로
                    - 10: 왼쪽 위로
                    - 11: 왼쪽 아래로
                    - 12: 오른쪽 가운데로
                    - 13: 오른쪽 위로
                    - 14: 오른쪽 아래로

            effect:
                이미지효과

                    - 0: 원래 그림(기본값)
                    - 1: 그레이 스케일
                    - 2: 흑백으로

            watermark: watermark효과 유무(True/False). 기본값은 False. 이 옵션이 True이면 brightness 와 contrast 옵션이 무시된다.
            brightness: 밝기 지정(-100 ~ 100), 기본 값은 0
            contrast: 선명도 지정(-100 ~ 100), 기본 값은 0

        Returns:
            성공했을 경우 True, 실패했을 경우 False

        Examples:
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

    def insert_ctrl(self, ctrl_id:str, initparam:Any) -> Ctrl:
        """
        현재 캐럿 위치에 컨트롤을 삽입한다.

        ctrlid에 지정할 수 있는 컨트롤 ID는 HwpCtrl.CtrlID가 반환하는 ID와 동일하다.
        자세한 것은  Ctrl 오브젝트 Properties인 CtrlID를 참조.
        initparam에는 컨트롤의 초기 속성을 지정한다.
        대부분의 컨트롤은 Ctrl.Properties와 동일한 포맷의 parameter set을 사용하지만,
        컨트롤 생성 시에는 다른 포맷을 사용하는 경우도 있다.
        예를 들어 표의 경우 Ctrl.Properties에는 "Table" 셋을 사용하지만,
        생성 시 initparam에 지정하는 값은 "TableCreation" 셋이다.

        Args:
            ctrl_id: 삽입할 컨트롤 ID
            initparam: 컨트롤 초기속성. 생략하면 default 속성으로 생성한다.

        Returns:
            생성된 컨트롤 object

        Examples:
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
            >>> table = hwp.insert_ctrl("tbl", tbset)  # <---
            >>> sleep(3)  # 표 생성 3초 후 다시 표 삭제
            >>> hwp.delete_ctrl(table)
        """
        return self.hwp.InsertCtrl(CtrlID=ctrl_id, initparam=initparam)

    def InsertCtrl(self, ctrl_id:str, initparam:"Hwp.HParameterSet") -> Ctrl:
        """
        현재 캐럿 위치에 컨트롤을 삽입한다.

        ctrlid에 지정할 수 있는 컨트롤 ID는 `HwpCtrl.CtrlID`가 반환하는 ID와 동일하다.
        자세한 것은  `Ctrl` 오브젝트 Properties인 `CtrlID`를 참조.
        `initparam`에는 컨트롤의 초기 속성을 지정한다.
        대부분의 컨트롤은 `Ctrl.Properties`와 동일한 포맷의 parameter set을 사용하지만,
        컨트롤 생성 시에는 다른 포맷을 사용하는 경우도 있다.
        예를 들어 표의 경우 Ctrl.Properties에는 `"Table"` 셋을 사용하지만,
        생성 시 `initparam`에 지정하는 값은 `"TableCreation"` 셋이다.

        Args:
            ctrl_id: 삽입할 컨트롤 ID
            initparam: 컨트롤 초기속성. 생략하면 default 속성으로 생성한다.

        Returns:
            생성된 컨트롤 object

        Examples:
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
            >>> table = hwp.InsertCtrl("tbl", tbset)  # <---
            >>> sleep(3)  # 표 생성 3초 후 다시 표 삭제
            >>> hwp.delete_ctrl(table)
        """
        return self.hwp.InsertCtrl(CtrlID=ctrl_id, initparam=initparam)

    def insert_picture(self, path: str, treat_as_char:bool=True, embedded:bool=True, sizeoption:int=0, reverse:bool=False, watermark:bool=False, effect:int=0, width:int=0, height:int=0) -> Ctrl:
        """
        현재 캐럿의 위치에 그림을 삽입한다.

        다만, 그림의 종횡비를 유지한 채로 셀의 높이만 키워주는 옵션이 없다.
        이런 작업을 원하는 경우에는 그림을 클립보드로 복사하고,
        Ctrl-V로 붙여넣기를 하는 수 밖에 없다.
        또한, 셀의 크기를 조절할 때 이미지의 크기도 따라 변경되게 하고 싶다면
        insert_background_picture 함수를 사용하는 것도 좋다.

        Args:
            path: 삽입할 이미지 파일의 전체경로
            embedded: 이미지 파일을 문서 내에 포함할지 여부 (True/False). 생략하면 True
            sizeoption:
                삽입할 그림의 크기를 지정하는 옵션. 기본값은 2

                    - 0: 이미지 원래의 크기로 삽입한다. width와 height를 지정할 필요 없다.(realSize)
                    - 1: width와 height에 지정한 크기로 그림을 삽입한다.(specificSize)
                    - 2: 현재 캐럿이 표의 셀 안에 있을 경우, 셀의 크기에 맞게 자동 조절하여 삽입한다. (종횡비 유지안함)(cellSize) 캐럿이 셀 안에 있지 않으면 이미지의 원래 크기대로 삽입된다.
                    - 3: 현재 캐럿이 표의 셀 안에 있을 경우, 셀의 크기에 맞추어 원본 이미지의 가로 세로의 비율이 동일하게 확대/축소하여 삽입한다.(cellSizeWithSameRatio)

            reverse: 이미지의 반전 유무 (True/False). 기본값은 False
            watermark: watermark효과 유무 (True/False). 기본값은 False
            effect:
                그림 효과

                    - 0: 실제 이미지 그대로
                    - 1: 그레이 스케일
                    - 2: 흑백효과

            width: 그림의 가로 크기 지정. 단위는 mm(HWPUNIT 아님!)
            height: 그림의 높이 크기 지정. 단위는 mm

        Returns:
            생성된 컨트롤 object.

        Examples:
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

                pset = self.HParameterSet.HShapeObject
                self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                if pset.ShapeTableCell.HasMargin == 1:  # 1이면
                    # 특정 셀 안여백
                    cell_pset = self.HParameterSet.HShapeObject
                    self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                    margin = round(cell_pset.ShapeTableCell.MarginLeft + cell_pset.ShapeTableCell.MarginRight, 2)
                else:
                    # 전역 셀 안여백
                    margin = round(pset.CellMarginLeft + pset.CellMarginRight, 2)

                cell_width = pset.ShapeTableCell.Width - margin
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
            return Ctrl(ctrl)
        finally:
            if os.path.basename(path).startswith("tmp"):
                os.remove(path)

    def insert_random_picture(self, x: int = 200, y: int = 200) -> Ctrl:
        """
        랜덤 이미지를 삽입한다.

        내부적으로 ``https://picsum.photos/`` API를 사용한다.

        Args:
            x: 너비 해상도
            y: 높이 해상도

        Returns:
            삽입한 이미지 컨트롤

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_random_picture(640, 480)
            <CtrlCode: CtrlID=gso, CtrlCH=11, UserDesc=그림>
        """
        return self.insert_picture(f"https://picsum.photos/{x}/{y}")

    def IsActionEnable(self, action_id:str) -> bool:
        """
        액션 실행 가능한지 여부를 bool로 리턴

        액션 관련해서는 기존 버전체크보다 이걸 사용하는 게
        훨씬 안정적일 것 같기는 하지만(예: CopyPage, PastePage, DeletePage 및 메타태그액션 등)
        신규 메서드(SelectCtrl 등) 지원여부는 체크해주지 못한다ㅜ
        """
        return self.hwp.IsActionEnable(actionID=action_id)

    def is_command_lock(self, action_id: str) -> bool:
        """
        해당 액션이 잠겨있는지 확인한다.

        Args:
            action_id: 액션 ID. (ActionIDTable.Hwp 참조)

        Returns:
            잠겨있으면 True, 잠겨있지 않으면 False를 반환한다.
        """
        return self.hwp.IsCommandLock(actionID=action_id)

    def IsCommandLock(self, action_id:str) -> bool:
        """
        해당 액션이 잠겨있는지 확인한다.

        Args:
            action_id: 액션 ID. (ActionIDTable.Hwp 참조)

        Returns:
            잠겨있으면 True, 잠겨있지 않으면 False를 반환한다.
        """
        return self.hwp.IsCommandLock(actionID=action_id)

    def key_indicator(self) -> tuple:
        """
        상태 바의 정보를 얻어온다.

        Returns:
            튜플(succ, seccnt, secno, prnpageno, colno, line, pos, over, ctrlname)

                - succ: 성공하면 True, 실패하면 False (항상 True임..)
                - seccnt: 총 구역
                - secno: 현재 구역
                - prnpageno: 쪽
                - colno: 단
                - line: 줄
                - pos: 칸
                - over: 삽입모드 (True: 수정, False: 삽입)
                - ctrlname: 캐럿이 위치한 곳의 컨트롤이름

        Examples:
            >>> # 현재 셀 주소(표 안에 있을 때)
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.key_indicator()[-1][1:].split(")")[0]
            "A1"
        """
        return self.hwp.KeyIndicator()

    def KeyIndicator(self) -> tuple:
        """
        상태 바의 정보를 얻어온다.

        Returns:
            튜플(succ, seccnt, secno, prnpageno, colno, line, pos, over, ctrlname)

                - succ: 성공하면 True, 실패하면 False (항상 True임..)
                - seccnt: 총 구역
                - secno: 현재 구역
                - prnpageno: 쪽
                - colno: 단
                - line: 줄
                - pos: 칸
                - over: 삽입모드 (True: 수정, False: 삽입)
                - ctrlname: 캐럿이 위치한 곳의 컨트롤이름

        Examples:
            >>> # 현재 셀 주소(표 안에 있을 때)
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.KeyIndicator()[-1][1:].split(")")[0]
            "A1"
        """
        return self.hwp.KeyIndicator()

    def lock_command(self, act_id: str, is_lock: bool) -> None:
        """
        특정 액션이 실행되지 않도록 잠근다.

        Args:
            act_id: 액션 ID. (ActionIDTable.Hwp 참조)
            is_lock: True이면 액션의 실행을 잠그고, False이면 액션이 실행되도록 한다.

        Returns:
            None

        Examples:
            >>> # Undo와 Redo 잠그기
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.lock_command("Undo", True)
            >>> hwp.lock_command("Redo", True)
        """
        return self.hwp.LockCommand(ActID=act_id, isLock=is_lock)

    def LockCommand(self, act_id:str, is_lock:bool) -> None:
        """
        특정 액션이 실행되지 않도록 잠근다.

        Args:
            act_id: 액션 ID. (ActionIDTable.Hwp 참조)
            is_lock: True이면 액션의 실행을 잠그고, False이면 액션이 실행되도록 한다.

        Returns:
            None

        Examples:
            >>> # Undo와 Redo 잠그기
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.LockCommand("Undo", True)
            >>> hwp.LockCommand("Redo", True)
        """
        return self.hwp.LockCommand(ActID=act_id, isLock=is_lock)

    def MarkPenNext(self):
        """
        다음 형광펜 삽입 위치로 이동한다.
        """
        return self.hwp.HAction.Run("MarkPenNext")

    def MarkPenPrev(self):
        """
        이전 형광펜 삽입 위치로 이동한다.
        """
        return self.hwp.HAction.Run("MarkPenPrev")

    def MetatagExist(self, tag):
        """특정 이름의 메타태그가 존재하는지?"""
        return self.hwp.MetatagExist(tag=tag)

    def modify_field_properties(self, field: str, remove: bool, add: bool):
        # """
        # 지정한 필드의 속성을 바꾼다. (사용안함)
        #
        # 양식모드에서 편집가능/불가 여부를 변경하는 메서드지만,
        # 현재 양식모드에서 어떤 속성이라도 편집가능하다..
        # 혹시 필드명이나 메모, 지시문을 수정하고 싶다면
        # set_cur_field_name 메서드를 사용하자.
        #
        # Args:
        #     field:
        #     remove:
        #     add:
        #
        # Returns:
        # """
        return self.hwp.ModifyFieldProperties(Field=field, remove=remove, Add=add)

    def ModifyFieldProperties(self, field, remove, add) -> None:
        """필드 속성 변경"""
        return self.hwp.ModifyFieldProperties(Field=field, remove=remove, Add=add)

    def modify_metatag_properties(self, tag, remove, add):
        """메타태그 속성 변경"""
        return self.hwp.ModifyMetatagProperties(tag=tag, remove=remove, Add=add)

    def ModifyMetatagProperties(self, tag, remove, add):
        """메타태그 속성 변경"""
        return self.hwp.ModifyMetatagProperties(tag=tag, remove=remove, Add=add)

    def move_pos(self, move_id: int = 1, para: int = 0, pos: int = 0) -> bool:
        """
        캐럿의 위치를 옮긴다.

        move_id를 200(moveScrPos)으로 지정한 경우에는
        스크린 좌표로 마우스 커서의 (x,y)좌표를 그대로 넘겨주면 된다.
        201(moveScanPos)는 문서를 검색하는 중 캐럿을 이동시키려 할 경우에만 사용이 가능하다.
        (솔직히 200 사용법은 잘 모르겠다;)

        Args:
            move_id:
                아래와 같은 값을 지정할 수 있다. 생략하면 1(moveCurList)이 지정된다.

                    - 0: 루트 리스트의 특정 위치.(para pos로 위치 지정) moveMain
                    - 1: 현재 리스트의 특정 위치.(para pos로 위치 지정) moveCurList
                    - 2: 문서의 시작으로 이동. moveTopOfFile
                    - 3: 문서의 끝으로 이동. moveBottomOfFile
                    - 4: 현재 리스트의 시작으로 이동 moveTopOfList
                    - 5: 현재 리스트의 끝으로 이동 moveBottomOfList
                    - 6: 현재 위치한 문단의 시작으로 이동 moveStartOfPara
                    - 7: 현재 위치한 문단의 끝으로 이동 moveEndOfPara
                    - 8: 현재 위치한 단어의 시작으로 이동.(현재 리스트만을 대상으로 동작한다.) moveStartOfWord
                    - 9: 현재 위치한 단어의 끝으로 이동.(현재 리스트만을 대상으로 동작한다.) moveEndOfWord
                    - 10: 다음 문단의 시작으로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextPara
                    - 11: 앞 문단의 끝으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevPara
                    - 12: 한 글자 뒤로 이동.(서브 리스트를 옮겨 다닐 수 있다.) moveNextPos
                    - 13: 한 글자 앞으로 이동.(서브 리스트를 옮겨 다닐 수 있다.) movePrevPos
                    - 14: 한 글자 뒤로 이동.(서브 리스트를 옮겨 다닐 수 있다. 머리말/꼬리말, 각주/미주, 글상자 포함.) moveNextPosEx
                    - 15: 한 글자 앞으로 이동.(서브 리스트를 옮겨 다닐 수 있다. 머리말/꼬리말, 각주/미주, 글상자 포함.) movePrevPosEx
                    - 16: 한 글자 뒤로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextChar
                    - 17: 한 글자 앞으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevChar
                    - 18: 한 단어 뒤로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextWord
                    - 19: 한 단어 앞으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevWord
                    - 20: 한 줄 아래로 이동. moveNextLine
                    - 21: 한 줄 위로 이동. movePrevLine
                    - 22: 현재 위치한 줄의 시작으로 이동. moveStartOfLine
                    - 23: 현재 위치한 줄의 끝으로 이동. moveEndOfLine
                    - 24: 한 레벨 상위로 이동한다. moveParentList
                    - 25: 탑레벨 리스트로 이동한다. moveTopLevelList
                    - 26: 루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 반환한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다. moveRootList
                    - 27: 현재 캐럿이 위치한 곳으로 이동한다. (캐럿 위치가 뷰의 맨 위쪽으로 올라간다.) moveCurrentCaret
                    - 100: 현재 캐럿이 위치한 셀의 왼쪽 moveLeftOfCell
                    - 101: 현재 캐럿이 위치한 셀의 오른쪽 moveRightOfCell
                    - 102: 현재 캐럿이 위치한 셀의 위쪽 moveUpOfCell
                    - 103: 현재 캐럿이 위치한 셀의 아래쪽 moveDownOfCell
                    - 104: 현재 캐럿이 위치한 셀에서 행(row)의 시작 moveStartOfCell
                    - 105: 현재 캐럿이 위치한 셀에서 행(row)의 끝 moveEndOfCell
                    - 106: 현재 캐럿이 위치한 셀에서 열(column)의 시작 moveTopOfCell
                    - 107: 현재 캐럿이 위치한 셀에서 열(column)의 끝 moveBottomOfCell
                    - 200: 한/글 문서창에서의 screen 좌표로서 위치를 설정 한다. moveScrPos
                    - 201: GetText() 실행 후 위치로 이동한다. moveScanPos
            para:
                이동할 문단의 번호. 0(moveMain) 또는 1(moveCurList)가 지정되었을 때만 사용된다. 200(moveScrPos)가 지정되었을 때는 문단번호가 아닌 스크린 좌표로 해석된다. (스크린 좌표 : LOWORD = x좌표, HIWORD = y좌표)

            pos: 이동할 문단 중에서 문자의 위치. 0(moveMain) 또는 1(moveCurList)가 지정되었을 때만 사용된다.

        Returns:
            성공하면 True, 실패하면 False
        """
        return self.hwp.MovePos(moveID=move_id, Para=para, pos=pos)

    def MovePos(self, move_id:int=1, para:int=0, pos:int=0) -> bool:
        """
        캐럿의 위치를 옮긴다.

        move_id를 200(moveScrPos)으로 지정한 경우에는
        스크린 좌표로 마우스 커서의 (x,y)좌표를 그대로 넘겨주면 된다.
        201(moveScanPos)는 문서를 검색하는 중 캐럿을 이동시키려 할 경우에만 사용이 가능하다.
        (솔직히 200 사용법은 잘 모르겠다;)

        Args:
            move_id:
                아래와 같은 값을 지정할 수 있다. 생략하면 1(moveCurList)이 지정된다.

                    - 0: 루트 리스트의 특정 위치.(para pos로 위치 지정) moveMain
                    - 1: 현재 리스트의 특정 위치.(para pos로 위치 지정) moveCurList
                    - 2: 문서의 시작으로 이동. moveTopOfFile
                    - 3: 문서의 끝으로 이동. moveBottomOfFile
                    - 4: 현재 리스트의 시작으로 이동 moveTopOfList
                    - 5: 현재 리스트의 끝으로 이동 moveBottomOfList
                    - 6: 현재 위치한 문단의 시작으로 이동 moveStartOfPara
                    - 7: 현재 위치한 문단의 끝으로 이동 moveEndOfPara
                    - 8: 현재 위치한 단어의 시작으로 이동.(현재 리스트만을 대상으로 동작한다.) moveStartOfWord
                    - 9: 현재 위치한 단어의 끝으로 이동.(현재 리스트만을 대상으로 동작한다.) moveEndOfWord
                    - 10: 다음 문단의 시작으로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextPara
                    - 11: 앞 문단의 끝으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevPara
                    - 12: 한 글자 뒤로 이동.(서브 리스트를 옮겨 다닐 수 있다.) moveNextPos
                    - 13: 한 글자 앞으로 이동.(서브 리스트를 옮겨 다닐 수 있다.) movePrevPos
                    - 14: 한 글자 뒤로 이동.(서브 리스트를 옮겨 다닐 수 있다. 머리말/꼬리말, 각주/미주, 글상자 포함.) moveNextPosEx
                    - 15: 한 글자 앞으로 이동.(서브 리스트를 옮겨 다닐 수 있다. 머리말/꼬리말, 각주/미주, 글상자 포함.) movePrevPosEx
                    - 16: 한 글자 뒤로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextChar
                    - 17: 한 글자 앞으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevChar
                    - 18: 한 단어 뒤로 이동.(현재 리스트만을 대상으로 동작한다.) moveNextWord
                    - 19: 한 단어 앞으로 이동.(현재 리스트만을 대상으로 동작한다.) movePrevWord
                    - 20: 한 줄 아래로 이동. moveNextLine
                    - 21: 한 줄 위로 이동. movePrevLine
                    - 22: 현재 위치한 줄의 시작으로 이동. moveStartOfLine
                    - 23: 현재 위치한 줄의 끝으로 이동. moveEndOfLine
                    - 24: 한 레벨 상위로 이동한다. moveParentList
                    - 25: 탑레벨 리스트로 이동한다. moveTopLevelList
                    - 26: 루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 반환한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다. moveRootList
                    - 27: 현재 캐럿이 위치한 곳으로 이동한다. (캐럿 위치가 뷰의 맨 위쪽으로 올라간다.) moveCurrentCaret
                    - 100: 현재 캐럿이 위치한 셀의 왼쪽 moveLeftOfCell
                    - 101: 현재 캐럿이 위치한 셀의 오른쪽 moveRightOfCell
                    - 102: 현재 캐럿이 위치한 셀의 위쪽 moveUpOfCell
                    - 103: 현재 캐럿이 위치한 셀의 아래쪽 moveDownOfCell
                    - 104: 현재 캐럿이 위치한 셀에서 행(row)의 시작 moveStartOfCell
                    - 105: 현재 캐럿이 위치한 셀에서 행(row)의 끝 moveEndOfCell
                    - 106: 현재 캐럿이 위치한 셀에서 열(column)의 시작 moveTopOfCell
                    - 107: 현재 캐럿이 위치한 셀에서 열(column)의 끝 moveBottomOfCell
                    - 200: 한/글 문서창에서의 screen 좌표로서 위치를 설정 한다. moveScrPos
                    - 201: GetText() 실행 후 위치로 이동한다. moveScanPos
            para:
                이동할 문단의 번호. 0(moveMain) 또는 1(moveCurList)가 지정되었을 때만 사용된다. 200(moveScrPos)가 지정되었을 때는 문단번호가 아닌 스크린 좌표로 해석된다. (스크린 좌표 : LOWORD = x좌표, HIWORD = y좌표)

            pos: 이동할 문단 중에서 문자의 위치. 0(moveMain) 또는 1(moveCurList)가 지정되었을 때만 사용된다.

        Returns:
            성공하면 True, 실패하면 False
        """
        return self.hwp.MovePos(moveID=move_id, Para=para, pos=pos)

    def move_to_field(self, field:str, idx:int=0, text:bool=True, start:bool=True, select:bool=False) -> bool:
        """
        지정한 필드로 캐럿을 이동한다.

        Args:
            field: 필드이름. GetFieldText()/PutFieldText()와 같은 형식으로 이름 뒤에 ‘{{#}}’로 번호를 지정할 수 있다.
            idx: 동일명으로 여러 개의 필드가 존재하는 경우, idx번째 필드로 이동하고자 할 때 사용한다. 기본값은 0. idx를 지정하지 않아도, 필드 파라미터 뒤에 ‘{{#}}’를 추가하여 인덱스를 지정할 수 있다. 이 경우 기본적으로 f스트링을 사용하며, f스트링 내부에 탈출문자열 \\가 적용되지 않으므로 중괄호를 다섯 겹 입력해야 한다. 예 : hwp.move_to_field(f"필드명{{{{{i}}}}}")
            text: 필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(True) 누름틀 코드로 이동할지(False)를 지정한다. 누름틀이 아닌 필드일 경우 무시된다. 생략하면 True가 지정된다.
            start: 필드의 처음(True)으로 이동할지 끝(False)으로 이동할지 지정한다. select를 True로 지정하면 무시된다. (캐럿이 처음에 위치해 있게 된다.) 생략하면 True가 지정된다.
            select: 필드 내용을 블록으로 선택할지(True), 캐럿만 이동할지(False) 지정한다. 생략하면 False가 지정된다.

        Returns:
            성공시 True, 실패시 False를 리턴한다.
        """
        if "{{" not in field:
            return self.hwp.MoveToField(Field=f"{field}{{{{{idx}}}}}", Text=text, start=start, select=select)
        else:
            return self.hwp.MoveToField(Field=field, Text=text, start=start, select=select)

    def MoveToField(self, field:str, idx:int=0, text:bool=True, start:bool=True, select:bool=False) -> bool:
        """
        지정한 필드로 캐럿을 이동한다.

        Args:
            field: 필드이름. `hwp.GetFieldText()` / `hwp.PutFieldText()`와 같은 형식으로 이름 뒤에 `'{{#}}'`로 번호를 지정할 수 있다.
            idx: 동일명으로 여러 개의 필드가 존재하는 경우, idx번째 필드로 이동하고자 할 때 사용한다. 기본값은 `0`. idx를 지정하지 않아도, 필드 파라미터 뒤에 `'{{#}}'`를 추가하여 인덱스를 지정할 수 있다. 이 경우 기본적으로 f스트링을 사용하며, f스트링 내부에 탈출문자열 `\\`가 적용되지 않으므로 중괄호를 다섯 겹 입력해야 한다. 예 : `hwp.move_to_field(f"필드명{{{{{i}}}}}")`
            text: 필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(`True`) 누름틀 코드로 이동할지(`False`)를 지정한다. 누름틀이 아닌 필드일 경우 무시된다. 생략하면 `True`가 지정된다.
            start: 필드의 처음(`True`)으로 이동할지 끝(`False`)으로 이동할지 지정한다. `select`를 `True`로 지정하면 무시된다. (캐럿이 처음에 위치해 있게 된다.) 생략하면 `True`가 지정된다.
            select: 필드 내용을 블록으로 선택할지(`True`), 캐럿만 이동할지(`False`) 지정한다. 생략하면 `False`가 지정된다.

        Returns:
            bool: 성공시 `True`, 실패시 `False`를 리턴한다.
        """
        if "{{" not in field:
            return self.hwp.MoveToField(Field=f"{field}{{{{{idx}}}}}", Text=text, start=start, select=select)
        else:
            return self.hwp.MoveToField(Field=field, Text=text, start=start, select=select)

    def move_to_metatag(self, tag, text, start, select):
        """특정 메타태그로 이동"""
        return self.hwp.MoveToMetatag(tag=tag, Text=text, start=start, select=select)

    def MoveToMetatag(self, tag, text, start, select):
        """특정 메타태그로 이동"""
        return self.hwp.MoveToMetatag(tag=tag, Text=text, start=start, select=select)

    def open(self, filename: str, format: str = "", arg: str = "") -> bool:
        """
        문서를 연다.

        Args:
            filename: 문서 파일의 전체경로
            format:
                문서 형식. 빈 문자열을 지정하면 자동으로 인식한다. 생략하면 빈 문자열이 지정된다.

                    - "HWP": 한/글 native format
                    - "HWP30": 한/글 3.X/96/97
                    - "HTML": 인터넷 문서
                    - "TEXT": 아스키 텍스트 문서
                    - "UNICODE": 유니코드 텍스트 문서
                    - "HWP20": 한글 2.0
                    - "HWP21": 한글 2.1/2.5
                    - "HWP15": 한글 1.X
                    - "HWPML1X": HWPML 1.X 문서 (Open만 가능)
                    - "HWPML2X": HWPML 2.X 문서 (Open / SaveAs 가능)
                    - "RTF": 서식 있는 텍스트 문서
                    - "DBF": DBASE II/III 문서
                    - "HUNMIN": 훈민정음 3.0/2000
                    - "MSWORD": 마이크로소프트 워드 문서
                    - "DOCRTF": MS 워드 문서 (doc)
                    - "OOXML": MS 워드 문서 (docx)
                    - "HANA": 하나워드 문서
                    - "ARIRANG": 아리랑 문서
                    - "ICHITARO": 一太郞 문서 (일본 워드프로세서)
                    - "WPS": WPS 문서
                    - "DOCIMG": 인터넷 프레젠테이션 문서(SaveAs만 가능)
                    - "SWF": Macromedia Flash 문서(SaveAs만 가능)

            arg:
                세부 옵션. 의미는 `format`에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.

                `arg`에 지정할 수 있는 옵션의 의미는 필터가 정의하기에 따라 다르지만,
                syntax는 다음과 같이 공통된 형식을 사용한다.

                `"key:value;key:value;..."`

                * key는 A-Z, a-z, 0-9, _ 로 구성된다.
                * value는 타입에 따라 다음과 같은 3 종류가 있다.

                    - boolean: ex) `fullsave:true` (== `fullsave`)
                    - integer: ex) `type:20`
                    - string:  ex) `prefix:_This_`

                * value는 생략 가능하며, 이때는 콜론도 생략한다.
                * arg에 지정할 수 있는 옵션

                모든 파일포맷

                    - `setcurdir` (boolean, true/false): 로드한 후 해당 파일이 존재하는 폴더로 현재 위치를 변경한다. hyperlink 정보가 상대적인 위치로 되어 있을 때 유용하다.

                HWP(HWPX)

                    - `lock` (boolean, TRUE): 로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
                    - `notext` (boolean, FALSE): 텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)
                    - template (boolean, FALSE): 새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다. 이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.
                    - suspendpassword (boolean, FALSE): TRUE로 지정하면 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.
                    - forceopen (boolean, FALSE): TRUE로 지정하면 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다.
                    - versionwarning (boolean, FALSE): TRUE로 지정하면 문서가 상위버전일 경우 메시지 박스를 띄우게 된다.

                HTML

                    - code(string, codepage): 문서변환 시 사용되는 코드 페이지를 지정할 수 있으며 code키가 존재할 경우 필터사용 시 사용자 다이얼로그를  띄우지 않는다.
                    - textunit(boolean, pixel): Export될 Text의 크기의 단위 결정.pixel, point, mili 지정 가능.
                    - formatunit(boolean, pixel): Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능

                ※ [codepage 종류]

                    - ks :  한글 KS 완성형
                    - kssm: 한글 조합형
                    - sjis : 일본
                    - utf8 : UTF8
                    - unicode: 유니코드
                    - gb : 중국 간체
                    - big5 : 중국 번체
                    - acp : Active Codepage 현재 시스템의 코드 페이지

                DOCIMG

                    - asimg(boolean, FALSE): 저장할 때 페이지를 image로 저장
                    - ashtml(boolean, FALSE): 저장할 때 페이지를 html로 저장

        Returns:
            성공하면 True, 실패하면 False
        """
        if filename and filename.startswith("http"):
            try:
                # url 문자열 중 hwp 파일명이 포함되어 있는지 체크해서 해당 파일명을 사용.
                hwp_name = [parse.unquote_plus(i) for i in re.split("[/?=&]", filename) if ".hwp" in i][0]
            except IndexError as e:
                # url 문자열 안에 hwp 파일명이 포함되어 있지 않은 경우에는 임시파일명 지정(temp.hwp)
                hwp_name = "temp.hwp"
            request.urlretrieve(filename, os.path.join(os.getcwd(), hwp_name))
            filename = os.path.join(os.getcwd(), hwp_name)
        elif filename.lower()[1] != ":" and os.path.exists(os.path.join(os.getcwd(), filename)):
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.Open(filename=filename, Format=format, arg=arg)

    def Open(self, filename: str, format: str = "", arg: str = "") -> bool:
        """
        문서를 연다.

        Args:
            filename: 문서 파일의 전체경로
            format:
                문서 형식. 빈 문자열을 지정하면 자동으로 인식한다. 생략하면 빈 문자열이 지정된다.

                    - "HWP": 한/글 native format
                    - "HWP30": 한/글 3.X/96/97
                    - "HTML": 인터넷 문서
                    - "TEXT": 아스키 텍스트 문서
                    - "UNICODE": 유니코드 텍스트 문서
                    - "HWP20": 한글 2.0
                    - "HWP21": 한글 2.1/2.5
                    - "HWP15": 한글 1.X
                    - "HWPML1X": HWPML 1.X 문서 (Open만 가능)
                    - "HWPML2X": HWPML 2.X 문서 (Open / SaveAs 가능)
                    - "RTF": 서식 있는 텍스트 문서
                    - "DBF": DBASE II/III 문서
                    - "HUNMIN": 훈민정음 3.0/2000
                    - "MSWORD": 마이크로소프트 워드 문서
                    - "DOCRTF": MS 워드 문서 (doc)
                    - "OOXML": MS 워드 문서 (docx)
                    - "HANA": 하나워드 문서
                    - "ARIRANG": 아리랑 문서
                    - "ICHITARO": 一太郞 문서 (일본 워드프로세서)
                    - "WPS": WPS 문서
                    - "DOCIMG": 인터넷 프레젠테이션 문서(SaveAs만 가능)
                    - "SWF": Macromedia Flash 문서(SaveAs만 가능)

            arg:
                세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.
                arg에 지정할 수 있는 옵션의 의미는 필터가 정의하기에 따라 다르지만,
                syntax는 다음과 같이 공통된 형식을 사용한다.

                `"key:value;key:value;..."`

                * key는 A-Z, a-z, 0-9, _ 로 구성된다.
                * value는 타입에 따라 다음과 같은 3 종류가 있다.
                boolean: ex) `fullsave:true (== fullsave)`
                integer: ex) `type:20`
                string:  ex) `prefix:_This_`
                * value는 생략 가능하며, 이때는 콜론도 생략한다.
                * arg에 지정할 수 있는 옵션

                모든 파일포맷

                    - setcurdir(boolean, true/false): 로드한 후 해당 파일이 존재하는 폴더로 현재 위치를 변경한다. hyperlink 정보가 상대적인 위치로 되어 있을 때 유용하다.

                HWP(HWPX)

                    - lock (boolean, TRUE): 로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
                    - notext (boolean, FALSE): 텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)
                    - template (boolean, FALSE): 새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다. 이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.
                    - suspendpassword (boolean, FALSE): TRUE로 지정하면 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.
                    - forceopen (boolean, FALSE): TRUE로 지정하면 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다.
                    - versionwarning (boolean, FALSE): TRUE로 지정하면 문서가 상위버전일 경우 메시지 박스를 띄우게 된다.

                HTML

                    - code(string, codepage): 문서변환 시 사용되는 코드 페이지를 지정할 수 있으며 code키가 존재할 경우 필터사용 시 사용자 다이얼로그를  띄우지 않는다.
                    - textunit(boolean, pixel): Export될 Text의 크기의 단위 결정.pixel, point, mili 지정 가능.
                    - formatunit(boolean, pixel): Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능

                ※ [codepage 종류]

                    - ks :  한글 KS 완성형
                    - kssm: 한글 조합형
                    - sjis : 일본
                    - utf8 : UTF8
                    - unicode: 유니코드
                    - gb : 중국 간체
                    - big5 : 중국 번체
                    - acp : Active Codepage 현재 시스템의 코드 페이지

                DOCIMG

                    - asimg(boolean, FALSE): 저장할 때 페이지를 image로 저장
                    - ashtml(boolean, FALSE): 저장할 때 페이지를 html로 저장

        Returns:
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

    def point_to_hwp_unit(self, point: float) -> int:
        """
        글자에 쓰이는 포인트 단위를 HwpUnit으로 변환
        """
        return self.hwp.PointToHwpUnit(Point=point)

    def PointToHwpUnit(self, point: float) -> int:
        """
        글자에 쓰이는 포인트 단위를 HwpUnit으로 변환
        """
        return self.hwp.PointToHwpUnit(Point=point)

    def hwp_unit_to_point(self, HwpUnit: int) -> float:
        """
        HwpUnit을 포인트 단위로 변환
        """
        return HwpUnit / 100

    def HwpUnitToPoint(self, HwpUnit: int) -> float:
        """
        HwpUnit을 포인트 단위로 변환
        """
        return HwpUnit / 100

    def hwp_unit_to_inch(self, HwpUnit: int) -> float:
        """
        HwpUnit을 인치로 변환
        """
        if HwpUnit == 0:
            return 0
        else:
            return HwpUnit / 7200

    def HwpUnitToInch(self, HwpUnit: int) -> float:
        """
        HwpUnit을 인치로 변환
        """
        if HwpUnit == 0:
            return 0
        else:
            return HwpUnit / 7200

    def inch_to_hwp_unit(self, inch) -> int:
        """
        인치 단위를 HwpUnit으로 변환
        """
        return round(inch * 7200, 0)

    def InchToHwpUnit(self, inch) -> int:
        """
        인치 단위를 HwpUnit으로 변환
        """
        return round(inch * 7200, 0)

    def protect_private_info(self, protecting_char:str, private_pattern_type:int) -> bool:
        """
        개인정보를 보호한다.

        한/글의 경우 “찾아서 보호”와 “선택 글자 보호”를 다른 기능으로 구현하였지만,
        API에서는 하나의 함수로 구현한다.

        Args:
            protecting_char: 보호문자. 개인정보는 해당문자로 가려진다.
            private_pattern_type: 보호유형. 개인정보 유형마다 설정할 수 있는 값이 다르다. 0값은 기본 보호유형으로 모든 개인정보를 보호문자로 보호한다.

        Returns:
            개인정보를 보호문자로 치환한 경우에 true를 반환한다.
	        개인정보를 보호하지 못할 경우 false를 반환한다.
	        문자열이 선택되지 않은 상태이거나, 개체가 선택된 상태에서는 실패한다.
	        또한, 보호유형이 잘못된 설정된 경우에도 실패한다.
	        마지막으로 보호암호가 설정되지 않은 경우에도 실패하게 된다.
        """
        return self.hwp.ProtectPrivateInfo(PotectingChar=protecting_char, PrivatePatternType=private_pattern_type)

    def ProtectPrivateInfo(self, protecting_char:str, private_pattern_type:int) -> bool:
        """
        개인정보를 보호한다.

        한/글의 경우 “찾아서 보호”와 “선택 글자 보호”를 다른 기능으로 구현하였지만,
        API에서는 하나의 함수로 구현한다.

        Args:
            protecting_char: 보호문자. 개인정보는 해당문자로 가려진다.
            private_pattern_type: 보호유형. 개인정보 유형마다 설정할 수 있는 값이 다르다. 0값은 기본 보호유형으로 모든 개인정보를 보호문자로 보호한다.

        Returns:
            개인정보를 보호문자로 치환한 경우에 true를 반환한다.
	        개인정보를 보호하지 못할 경우 false를 반환한다.
	        문자열이 선택되지 않은 상태이거나, 개체가 선택된 상태에서는 실패한다.
	        또한, 보호유형이 잘못된 설정된 경우에도 실패한다.
	        마지막으로 보호암호가 설정되지 않은 경우에도 실패하게 된다.
        """
        return self.hwp.ProtectPrivateInfo(PotectingChar=protecting_char, PrivatePatternType=private_pattern_type)

    def put_field_text(self, field: Any = "", text: Union[str, list, tuple, pd.Series] = "", idx=None) -> None:
        """
        지정한 필드의 내용을 채운다.

        현재 필드에 입력되어 있는 내용은 지워진다.
        채워진 내용의 글자모양은 필드에 지정해 놓은 글자모양을 따라간다.
        fieldlist의 필드 개수와, textlist의 텍스트 개수는 동일해야 한다.
        존재하지 않는 필드에 대해서는 무시한다.

        Args:
            field: 내용을 채울 필드 이름의 리스트. 한 번에 여러 개의 필드를 지정할 수 있으며, 형식은 GetFieldText와 동일하다. 다만 필드 이름 뒤에 "{{#}}"로 번호를 지정하지 않으면 해당 이름을 가진 모든 필드에 동일한 텍스트를 채워 넣는다. 즉, PutFieldText에서는 ‘필드이름’과 ‘필드이름{{0}}’의 의미가 다르다. **단, field에 dict를 입력하는 경우에는 text 파라미터를 무시하고 dict.keys를 필드명으로, dict.values를 필드값으로 입력한다.**
            text: 필드에 채워 넣을 문자열의 리스트. 형식은 필드 리스트와 동일하게 필드의 개수만큼 텍스트를 0x02로 구분하여 지정한다.

        Returns:
            None

        Examples:
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

    def PutFieldText(self, field: Any = "", text: Union[str, list, tuple, pd.Series] = "", idx=None) -> None:
        """
        지정한 필드의 내용을 채운다.

        현재 필드에 입력되어 있는 내용은 지워진다.
        채워진 내용의 글자모양은 필드에 지정해 놓은 글자모양을 따라간다.
        fieldlist의 필드 개수와, textlist의 텍스트 개수는 동일해야 한다.
        존재하지 않는 필드에 대해서는 무시한다.

        Args:
            field: 내용을 채울 필드 이름의 리스트. 한 번에 여러 개의 필드를 지정할 수 있으며, 형식은 GetFieldText와 동일하다. 다만 필드 이름 뒤에 "{{#}}"로 번호를 지정하지 않으면 해당 이름을 가진 모든 필드에 동일한 텍스트를 채워 넣는다. 즉, PutFieldText에서는 ‘필드이름’과 ‘필드이름{{0}}’의 의미가 다르다. **단, field에 dict를 입력하는 경우에는 text 파라미터를 무시하고 dict.keys를 필드명으로, dict.values를 필드값으로 입력한다.**
            text: 필드에 채워 넣을 문자열의 리스트. 형식은 필드 리스트와 동일하게 필드의 개수만큼 텍스트를 "\\x02"로 구분하여 지정한다.

        Returns:
            None

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 현재 캐럿 위치에 zxcv 필드 생성
            >>> hwp.CreateField("zxcv")
            >>> # zxcv 필드에 "Hello world!" 텍스트 삽입
            >>> hwp.PutFieldText("zxcv", "Hello world!")
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
        """메타태그에 텍스트 삽입"""
        return self.hwp.PutMetatagNameText(tag=tag, Text=text)

    def PutMetatagNameText(self, tag, text):
        """메타태그에 텍스트 삽입"""
        return self.hwp.PutMetatagNameText(tag=tag, Text=text)

    def PutParaNumber(self):
        """
        문단번호 삽입/제거 토글

        """
        return self.hwp.HAction.Run("PutParaNumber")

    def PutOutlinleNumber(self):
        """
        개요번호 삽입/제거 토글

        """
        return self.hwp.HAction.Run("PutOutlineNumber")

    def quit(self, save:bool=False) -> None:
        """
        한/글을 종료한다.

        단, 저장되지 않은 변경사항이 있는 경우 팝업이 뜨므로
        clear나 save 등의 메서드를 실행한 후에 quit을 실행해야 한다.

        Args:
            save: 변경사항이 있는 경우 저장할지 여부. 기본값은 저장안함(False)

        Returns:
            None
        """
        if self.Path == "":  # 빈 문서인 경우
            self.clear()
        elif save: # 빈 문서가 아닌 경우
            self.save()
        self.hwp.Quit()

    def Quit(self, save:bool=False) -> None:
        """
        한/글을 종료한다.

        단, 저장되지 않은 변경사항이 있는 경우 팝업이 뜨므로
        clear나 save 등의 메서드를 실행한 후에 Quit을 실행해야 한다.

        Args:
            save: 변경사항이 있는 경우 저장할지 여부. 기본값은 저장안함(False)

        Returns:
            None
        """
        if self.Path == "":  # 빈 문서인 경우
            self.clear()
        elif save:  # 빈 문서가 아닌 경우
            self.save()
        self.hwp.Quit()

    def rgb_color(self, red_or_colorname: str | tuple | int, green: int = 255, blue: int = 255) -> int:
        """
        RGB값을 한/글이 인식하는 정수 형태로 변환해주는 헬퍼 메서드.

        ![rgb_color / RGBColor](assets/rgb_color.gif){ loading=lazy }

        주로 글자색이나, 셀 색깔을 적용할 때 사용한다.
        RGB값을 세 개의 정수로 입력하는 것이 기본적인 사용방법이지만,
        자주 사용되는 아래의 24가지 색깔은 문자열로 입력 가능하다.

        Args:
            red_or_colorname:
                R값(0~255)을 입력하거나, 혹은 아래 목록의 색깔 문자열을 직접 입력할 수 있다.

                    - "Red": (255, 0, 0)
                    - "Green": (0, 255, 0)
                    - "Blue": (0, 0, 255)
                    - "Yellow": (255, 255, 0)
                    - "Cyan": (0, 255, 255)
                    - "Magenta": (255, 0, 255)
                    - "Black": (0, 0, 0)
                    - "White": (255, 255, 255)
                    - "Gray": (128, 128, 128)
                    - "Orange": (255, 165, 0)
                    - "DarkBlue": (0, 0, 139)
                    - "Purple": (128, 0, 128)
                    - "Pink": (255, 192, 203)
                    - "Lime": (0, 255, 0)
                    - "SkyBlue": (135, 206, 235)
                    - "Gold": (255, 215, 0)
                    - "Silver": (192, 192, 192)
                    - "Mint": (189, 252, 201)
                    - "Tomato": (255, 99, 71)
                    - "Olive": (128, 128, 0)
                    - "Crimson": (220, 20, 60)
                    - "Navy": (0, 0, 128)
                    - "Teal": (0, 128, 128)
                    - "Chocolate": (210, 105, 30)

            green: G값(0~255)
            blue: B값(0~255)

        Returns:
            아래아한글이 인식하는 정수 형태의 RGB값.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.set_font(TextColor=hwp.RGBColor("Red"))  # 글자를 빨강으로
            >>> hwp.insert_text("빨간 글자색\\r\\n")
            >>> hwp.set_font(ShadeColor=hwp.RGBColor(0, 255, 0))  # 음영을 녹색으로
            >>> hwp.insert_text("초록 음영색")
        """
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

    def RGBColor(self, red_or_colorname: str | tuple | int, green: int = 255, blue: int = 255) -> int:
        """
        RGB값을 한/글이 인식하는 정수 형태로 변환해주는 메서드.

        자주 쓰이는 24가지 색깔은 문자열로 입력 가능하다.

        Args:
            red_or_colorname: R값(0~255) 또는 색깔 문자열
            green: G값(0~255)
            blue: B값(0~255)

        Returns:
            아래아한글이 인식하는 정수 형태의 RGB값.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.set_font(TextColor=hwp.RGBColor("Red"))  # 글자색 빨강
            >>> hwp.insert_text("빨간 글자색\\r\\n")
            >>> hwp.set_font(ShadeColor=hwp.RGBColor(0, 255, 0))  # 음영색 초록
            >>> hwp.insert_text("초록 음영색")
        """
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

    def register_module(self, module_type: str = "FilePathCheckDLL",
                        module_data: str = "FilePathCheckerModule") -> bool:
        """
        (인스턴스 생성시 자동으로 실행된다.)

        한/글 컨트롤에 부가적인 모듈을 등록한다.
        사용자가 모르는 사이에 파일이 수정되거나 서버로 전송되는 것을 막기 위해
        한/글 오토메이션은 파일을 불러오거나 저장할 때 사용자로부터 승인을 받도록 되어있다.
        그러나 이미 검증받은 웹페이지이거나,
        이미 사용자의 파일 시스템에 대해 강력한 접근 권한을 갖는 응용프로그램의 경우에는
        이러한 승인절차가 아무런 의미가 없으며 오히려 불편하기만 하다.
        이런 경우 register_module을 통해 보안승인모듈을 등록하여 승인절차를 생략할 수 있다.

        Args:
            module_type: 모듈의 유형. 기본값은 "FilePathCheckDLL"이다. 파일경로 승인모듈을 DLL 형태로 추가한다.
            module_data: Registry에 등록된 DLL 모듈 ID

        Returns:
            추가모듈등록에 성공하면 True를, 실패하면 False를 반환한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 사전에 레지스트리에 보안모듈이 등록되어 있어야 한다.
            >>> # 보다 자세한 설명은 공식문서 참조
            >>> hwp.register_module("FilePathChekDLL", "FilePathCheckerModule")
            True
        """
        if not check_registry_key():
            self.register_regedit()
        return self.hwp.RegisterModule(ModuleType=module_type, ModuleData=module_data)

    def RegisterModule(self, module_type: str = "FilePathCheckDLL", module_data: str = "FilePathCheckerModule") -> bool:
        """
        한/글 컨트롤에 부가적인 모듈을 등록한다. 기본동작은 "보안모듈 등록"

        (인스턴스 생성시 자동으로 실행된다.)
        사용자가 모르는 사이에 파일이 수정되거나 서버로 전송되는 것을 막기 위해
        한/글 오토메이션은 파일을 불러오거나 저장할 때 사용자로부터 승인을 받도록 되어있다.
        그러나 이미 검증받은 웹페이지이거나,
        이미 사용자의 파일 시스템에 대해 강력한 접근 권한을 갖는 응용프로그램의 경우에는
        이러한 승인절차가 아무런 의미가 없으며 오히려 불편하기만 하다.
        이런 경우 register_module을 통해 보안승인모듈을 등록하여 승인절차를 생략할 수 있다.

        Args:
            module_type: 모듈의 유형 문자열. 기본값은 보안모듈인 "FilePathCheckDLL"이다. 파일경로 승인모듈을 DLL 형태로 추가한다.
            module_data: Registry에 등록된 DLL 모듈 ID

        Returns:
            추가모듈등록에 성공하면 True를, 실패하면 False를 반환한다.
        """
        if not check_registry_key(module_data):
            self.register_regedit()
        return self.hwp.RegisterModule(ModuleType=module_type, ModuleData=module_data)

    @staticmethod
    def register_regedit(dll_name: str = "FilePathCheckerModule.dll") -> None:
        """
        레지스트리 에디터에 한/글 보안모듈을 자동등록하는 메서드.

        가장 먼저 파이썬과 pyhwpx 모듈이 설치된 상태라고 가정하고 'site-packages/pyhwpx' 폴더에서
        'FilePathCheckerModule.dll' 파일을 찾는다.
        두 번째로는 pyinstaller로 컴파일했다고 가정하고, MEIPASS 하위폴더를 탐색한다.
        이후로, 차례대로 실행파일과 동일한 경로, 사용자 폴더를 탐색한 후에도 보안모듈 dll파일을 찾지 못하면
        아래아한글 깃헙 저장소에서 직접 보안모듈을 다운받아 사용자 폴더에 설치하고, 레지스트리를 수정한다.

        Args:
            dll_name: 보안모듈 dll 파일명. 관례적으로 FilePathCheckerModule.dll을 쓴다.

        Returns:
            None

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()  # 이 때 내부적으로 register_regedit를 실행한다.
        """
        import os
        import subprocess
        from winreg import ConnectRegistry, HKEY_CURRENT_USER, OpenKey, KEY_WRITE, SetValueEx, REG_SZ, CloseKey

        try:
            # pyhwpx가 설치된 파이썬 환경 또는 pyinstaller로 컴파일한 환경에서 pyhwpx 경로 찾기
            # 살펴본 결과, FilePathCheckerModule.dll 파일은 pyinstaller 컴파일시 자동포함되지는 않는 것으로 확인..
            location = [i.split(": ")[1] for i in
                        subprocess.check_output(['pip', 'show', 'pyhwpx'], stderr=subprocess.DEVNULL).decode(
                            encoding="cp949").split("\r\n") if i.startswith("Location: ")][0]
            location = os.path.join(location, "pyhwpx")
        except UnicodeDecodeError:
            location = [i.split(": ")[1] for i in
                        subprocess.check_output(['pip', 'show', 'pyhwpx'], stderr=subprocess.DEVNULL).decode().split(
                            "\r\n") if i.startswith("Location: ")][0]
            location = os.path.join(location, "pyhwpx")
        print("default dll :", os.path.join(location, dll_name))
        if not os.path.exists(os.path.join(location, dll_name)):
            print("위 폴더에서 보안모듈을 찾을 수 없음..")
            location = ""
            # except subprocess.CalledProcessError as e:
            # FilePathCheckerModule.dll을 못 찾는 경우에는 아래 분기 중 하나를 실행
            #

            # 1. pyinstaller로 컴파일했고,
            #    --add-binary="FilePathCheckerModule.dll:." 옵션을 추가한 경우
            for dirpath, dirnames, filenames in os.walk(pyinstaller_path):
                for filename in filenames:
                    if filename.lower() == dll_name.lower():
                        location = dirpath
                        print(location, "에서 보안모듈을 찾았습니다.")
                        break
            print("pyinstaller 하위경로에 보안모듈 없음..")

            # 2. "FilePathCheckerModule.dll" 파일을 실행파일과 같은 경로에 둔 경우

            if dll_name.lower() in [i.lower() for i in os.listdir(os.getcwd())]:
                print("실행파일 경로에서 보안모듈을 찾았습니다.")
                location = os.getcwd()
                print("보안모듈 경로 :", location)
            # elif os.path.exists(os.path.join(os.environ["USERPROFILE"], "FilePathCheckerModule.dll")):
            elif dll_name.lower in [i.lower() for i in os.listdir(os.path.join(os.environ["USERPROFILE"]))]:
                print("사용자 폴더에서 보안모듈을 찾았습니다.")
                location = os.environ["USERPROFILE"]
                print("보안모듈 경로 :", location)
            # 3. 위의 두 경우가 아닐 때, 인터넷에 연결되어 있는 경우에는
            #    사용자 폴더(예: c:\\users\\user)에
            #    FilePathCheckerModule.dll을 다운로드하기.
            if location == "":
                # pyhwpx가 설치되어 있지 않은 PC에서는,
                # 공식사이트에서 다운을 받게 하자.
                from zipfile import ZipFile
                print("https://github.com/hancom-io에서 보안모듈 다운로드를 시도합니다.")
                try:
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
                                  os.path.join(os.environ["USERPROFILE"], dll_name))
                    location = os.environ["USERPROFILE"]
                    print("사용자폴더", location, "에 보안모듈을 설치하였습니다.")
                except urllib.error.URLError as e:
                    # URLError를 처리합니다.
                    print(f"내부망에서는 보안모듈을 다운로드할 수 없습니다. 보안모듈을 직접 다운받아 설치하여 주시기 바랍니다.: \n{e.reason}")
                except Exception as e:
                    # 기타 예외를 처리합니다.
                    print(f"예기치 못한 오류가 발생했습니다. 아래 오류를 개발자에게 문의해주시기 바랍니다: \n{str(e)}")
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
        SetValueEx(key, "FilePathCheckerModule", 0, REG_SZ, os.path.join(location, dll_name))
        CloseKey(key)

    def register_private_info_pattern(self, private_type:int, private_pattern:str) -> bool:
        """
        개인정보의 패턴을 등록한다.

        (현재 작동하지 않는다.)

        Args:
            private_type:
                등록할 개인정보 유형. 다음의 값 중 하나다.

                    - 0x0001: 전화번호
                    - 0x0002: 주민등록번호
                    - 0x0004: 외국인등록번호
                    - 0x0008: 전자우편
                    - 0x0010: 계좌번호
                    - 0x0020: 신용카드번호
                    - 0x0040: IP 주소
                    - 0x0080: 생년월일
                    - 0x0100: 주소
                    - 0x0200: 사용자 정의

            private_pattern:
                등록할 개인정보 패턴. 예를 들면 이런 형태로 입력한다.
                (예) 주민등록번호 - "NNNNNN-NNNNNNN"
                한/글이 이미 정의한 패턴은 정의하면 안 된다.
                함수를 여러 번 호출하는 것을 피하기 위해 패턴을 “;”기호로 구분
                반속해서 입력할 수 있도록 한다.

        Returns:
            등록이 성공하였으면 True, 실패하였으면 False

        Examples:
            >>> from pyhwpx import Hwp()
            >>> hwp = Hwp()
            >>>
            >>> hwp.register_private_info_pattern(0x01, "NNNN-NNNN;NN-NN-NNNN-NNNN")  # 전화번호패턴
        """
        return self.hwp.RegisterPrivateInfoPattern(PrivateType=private_type, PrivatePattern=private_pattern)

    def RegisterPrivateInfoPattern(self, private_type: int, private_pattern: int) -> bool:
        """
        개인정보의 패턴을 등록한다.

        (현재 작동하지 않는다.)

        Args:
            private_type:
                등록할 개인정보 유형. 다음의 값 중 하나다.

                    - 0x0001: 전화번호
                    - 0x0002: 주민등록번호
                    - 0x0004: 외국인등록번호
                    - 0x0008: 전자우편
                    - 0x0010: 계좌번호
                    - 0x0020: 신용카드번호
                    - 0x0040: IP 주소
                    - 0x0080: 생년월일
                    - 0x0100: 주소
                    - 0x0200: 사용자 정의

            private_pattern:
                등록할 개인정보 패턴. 예를 들면 이런 형태로 입력한다.

			    (예) 주민등록번호 - "NNNNNN-NNNNNNN"

			    한/글이 이미 정의한 패턴은 정의하면 안 된다.
			    함수를 여러 번 호출하는 것을 피하기 위해 패턴을 “;”기호로 구분
			    반복해서 입력할 수 있도록 한다.

        Returns:
            등록이 성공하였으면 True, 실패하였으면 False

        Examples:
            >>> from pyhwpx import Hwp()
            >>> hwp = Hwp()
            >>> hwp.register_private_info_pattern(0x01, "NNNN-NNNN;NN-NN-NNNN-NNNN")  # 전화번호패턴
        """
        return self.hwp.RegisterPrivateInfoPattern(PrivateType=private_type, PrivatePattern=private_pattern)

    def ReleaseAction(self, action: str):
        return self.hwp.ReleaseAction(action=action)

    def release_scan(self) -> None:
        """
        InitScan()으로 설정된 초기화 정보를 해제한다.

        텍스트 검색작업이 끝나면 반드시 호출하여 설정된 정보를 해제해야 한다.

        Returns:
            None
        """
        return self.hwp.ReleaseScan()

    def ReleaseScan(self) -> None:
        """
        InitScan()으로 설정된 초기화 정보를 해제한다.

        텍스트 검색작업이 끝나면 반드시 호출하여 설정된 정보를 해제해야 한다.

        Returns:
            None
        """
        return self.hwp.ReleaseScan()

    def rename_field(self, oldname: str, newname: str) -> bool:
        """
        지정한 필드의 이름을 바꾼다.

        예를 들어 oldname에 "title{{0}}\\x02title{{1}}",
        newname에 "tt1\\x02tt2로 지정하면 첫 번째 title은 tt1로, 두 번째 title은 tt2로 변경된다.
        oldname의 필드 개수와, newname의 필드 개수는 동일해야 한다.
        존재하지 않는 필드에 대해서는 무시한다.

        Args:
            oldname: 이름을 바꿀 필드 이름의 리스트. 형식은 PutFieldText와 동일하게 "\\x02"로 구분한다.
            newname: 새로운 필드 이름의 리스트. oldname과 동일한 개수의 필드 이름을 "\\x02"로 구분하여 지정한다.

        Returns:
            None

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_field("asdf")  # "asdf" 필드 생성
            >>> hwp.rename_field("asdf", "zxcv")  # asdf 필드명을 "zxcv"로 변경
            >>> hwp.put_field_text("zxcv", "Hello world!")  # zxcv 필드에 텍스트 삽입
        """
        return self.hwp.RenameField(oldname=oldname, newname=newname)

    def rename_metatag(self, oldtag, newtag):
        """메타태그 이름 변경"""
        return self.hwp.RenameMetatag(oldtag=oldtag, newtag=newtag)

    def RenameMetatag(self, oldtag, newtag):
        """메타태그 이름 변경"""
        return self.hwp.RenameMetatag(oldtag=oldtag, newtag=newtag)

    def replace_action(self, old_action_id: str, new_action_id: str) -> bool:
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

        Args:
            old_action_id: 변경될 원본 Action ID. 한/글 컨트롤에서 사용할 수 있는 Action ID는 ActionTable.hwp(별도문서)를 참고한다.
            new_action_id: 변경할 대체 Action ID. 기존의 Action ID와 UserAction ID(ver:0x07050206) 모두 사용가능하다.

        Returns:
            Action을 바꾸면 True를 바꾸지 못했다면 False를 반환한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.replace_action("Cut", "Cut")
        """

        return self.hwp.ReplaceAction(OldActionID=old_action_id, NewActionID=new_action_id)

    def ReplaceAction(self, old_action_id: str, new_action_id: str) -> bool:
        """
        특정 Action을 다른 Action으로 대체한다.

        이는 메뉴나 단축키로 호출되는 Action을 대체할 뿐,
        CreateAction()이나, Run() 등의 함수를 이용할 때에는 아무런 영향을 주지 않는다.
        즉, ReplaceAction(“Cut", "Copy")을 호출하여
        ”오려내기“ Action을 ”복사하기“ Action으로 교체하면
        Ctrl+X 단축키나 오려내기 메뉴/툴바 기능을 수행하더라도 복사하기 기능이 수행되지만,
        코드 상에서 Run("Cut")을 실행하면 오려내기 Action이 실행된다.
        또한, 대체된 Action을 원래의 Action으로 되돌리기 위해서는
        NewActionID의 값을 원래의 Action으로 설정한 뒤 호출한다.

        Args:
            old_action_id: 변경될 원본 Action ID. 한/글 컨트롤에서 사용할 수 있는 Action ID는 ActionTable.hwp(별도문서)를 참고한다.
            new_action_id: 변경할 대체 Action ID. 기존의 Action ID와 UserAction ID(ver:0x07050206) 모두 사용가능하다.

        Returns:
            Action을 바꾸면 True를, 바꾸지 못했다면 False를 반환한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.replace_action("Cut", "Cut")
        """

        return self.hwp.ReplaceAction(OldActionID=old_action_id, NewActionID=new_action_id)

    def replace_font(self, langid, des_font_name, des_font_type, new_font_name, new_font_type):
        return self.hwp.ReplaceFont(langid=langid, desFontName=des_font_name, desFontType=des_font_type,
                                    newFontName=new_font_name, newFontType=new_font_type)

    def ReplaceFont(self, langid, des_font_name, des_font_type, new_font_name, new_font_type):
        return self.hwp.ReplaceFont(langid=langid, desFontName=des_font_name, desFontType=des_font_type,
                                    newFontName=new_font_name, newFontType=new_font_type)

    def Revision(self, revision):
        return self.hwp.Revision(Revision=revision)

    # Run 액션

    def Run(self, act_id: str) -> bool:
        """
        액션을 실행한다.

        ActionTable.hwp 액션 리스트 중에서 "별도의 파라미터가 필요하지 않은" 단순 액션을 hwp.Run(액션아이디)으로 호출할 수 있다. 단, `hwp.Run("BreakPara")` 처럼 실행하는 대신 `hwp.BreakPara()` 방식으로 실행 가능하다.

        Args:
            act_id: 액션 ID (ActionIDTable.hwp 참조)

        Returns:
            성공시 True, 실패시 False를 반환한다.
        """
        return self.hwp.HAction.Run(act_id)

    def ASendBrowserText(self):
        """
        웹브라우저로 보내기
        """
        return self.hwp.HAction.Run("ASendBrowserText")

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

    def BottomTabFrameClose(self):
        """아래쪽 작업창 감추기"""
        return self.hwp.HAction.Run("BottomTabFrameClose")

    def BreakColDef(self):
        """
        다단 레이아웃을 사용하는 경우의 "단 정의 삽입 액션(Ctrl-Alt-Enter)".

        단 정의 삽입 위치를 기점으로 구분된 다단을 하나 추가한다.
        다단이 아닌 경우에는 일반 "문단나누기(Enter)"와 동일하다.
        """
        return self.hwp.HAction.Run("BreakColDef")

    def BreakColumn(self):
        """
        다단 레이아웃을 사용하는 경우 "단 나누기[배분다단] 액션(Ctrl-Shift-Enter)".

        단 정의 삽입 위치를 기점으로 구분된 다단을 하나 추가한다.
        다단이 아닌 경우에는 일반 "문단나누기(Enter)"와 동일하다.

        """
        return self.hwp.HAction.Run("BreakColumn")

    def BreakLine(self):
        """
        라인나누기 액션(Shift-Enter).

        들여쓰기나 내어쓰기 등 문단속성이 적용되어 있는 경우에
        속성을 유지한 채로 줄넘김만 삽입한다. 이 단축키를 모르고 보고서를 작성하면,
        들여쓰기를 맞추기 위해 스페이스를 여러 개 삽입했다가,
        앞의 문구를 수정하는 과정에서 스페이스 뭉치가 문단 중간에 들어가버리는 대참사가 자주 발생할 수 있다.
        """
        return self.hwp.HAction.Run("BreakLine")

    def BreakPage(self):
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

    def ChangeSkin(self):
        """스킨 바꾸기"""
        return self.hwp.HAction.Run("ChangeSkin")

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

    def CharShapeLang(self):
        """글자 언어"""
        return self.hwp.HAction.Run("CharShapeLang")

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

    def CharShapePrevFaceName(self):
        """
        이전 글꼴 ALT+SHIFT+G
        """
        return self.hwp.HAction.Run("CharShapePrevFaceName")

    def CharShapeShadow(self):
        """
        선택한 텍스트 글자모양 중 그림자 속성을 토글한다.

        """
        return self.hwp.HAction.Run("CharShapeShadow")

    def CharShapeSpacing(self):
        """
        글자모양(alt-L) 창을 열고, 자간 값에 포커스를 옮긴다.
        """
        return self.hwp.HAction.Run("CharShapeSpacing")

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

    def CharShapTypeface(self):
        """글자 모양 (실행 안됨)"""
        return self.hwp.HAction.Run("CharShapTypeface")

    def CharShapeUnderline(self):
        """
        선택한 텍스트에 밑줄 속성을 토글한다. 대소문자에 유의해야 한다. (UnderLine이 아니다.)

        """
        return self.hwp.HAction.Run("CharShapeUnderline")

    def CharShapeWidth(self):
        """글자 모양(Alt-L) 창에서 글자 장평에 포커스를 둔다."""
        return self.hwp.HAction.Run("CharShapeWidth")

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
        현재 리스트를 닫고 (최)상위 리스트로 이동하는 액션.

        대표적인 예로, 메모나 각주 등을 작성한 후 본문으로 빠져나올 때, 혹은 여러 겹의 표 안에 있을 때 한 번에 표 밖으로 캐럿을 옮길 때 사용한다. 굉장히 자주 쓰이는 액션이며, 경우에 따라 Close가 아니라 CloseEx를 써야 하는 경우도 있다.
        (레퍼런스 포인트가 등록되어 있으면 그 포인트로, 없으면 루트 리스트로 이동한다. 나머지 특성은 MoveRootList와 동일)
        """
        cur_pos = self.GetPos()
        self.hwp.HAction.Run("Close")
        for _ in range(5):
            if self.GetPos() != cur_pos:
                break
            else:
                sleep(0.05)

    def CloseEx(self):
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

    def CommentDelete(self):
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

    def ComposeChars(self, Chars: str | int = "", CharSize: int = -3, CheckCompose: int = 0, CircleType: int = 0,
                     **kwargs) -> bool:
        """
        글자 겹치기 메서드(원문자 만들기)

        캐럿 위치의 서체를 따라가지만, 임의의 키워드로 폰트 수정 가능(예: Bold=True, Italic=True, TextColor=hwp.RGBColor(255,0,0) 등)

        Args:
            Chars: 겹칠 글자(정수도 문자열로 인식)
            CharSize: 글자확대(2:150%, 1:140%, 0:130%, -1:120%, -2:110%, -3:100%, -4:90%, -5:80%, -6:70%, -7:60%, -8:50%)
            CheckCompose: 모양 안에 글자 겹치기 여부(1이면 글자들끼리도 겹침)
            CircleType: 테두리 모양(0:없음, 1:원, 2:반전원, 3:사각, 4:반전사각, 5:삼각, 6:반전삼각, 7:해, 8:마름모, 9:반전마름모, 10:뭉툭사각, 11:재활용빈화살표, 12:재활용화살표, 13:재활용채운화살표)

        Returns:
            성공하면 True, 실패하면 False를 리턴
        """
        pset = self.HParameterSet.HChCompose
        self.HAction.GetDefault("ComposeChars", pset.HSet)
        pset.Chars = Chars
        pset.CharSize = CharSize
        pset.CheckCompose = CheckCompose
        pset.CircleType = CircleType  # 0~13
        for key in kwargs:
            if kwargs[key] != "":
                setattr(pset.CharShapes.CircleCharShape, key, kwargs[key])
        return self.HAction.Execute("ComposeChars", pset.HSet)

    def Copy(self):
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

    def CopyPage(self):
        """
        쪽 복사

        한글2014 이하의 버전에서는 사용할 수 없다.
        """
        return self.hwp.HAction.Run("CopyPage")

    def CustCopyBtn(self):
        """툴바 버튼 복사하기"""
        return self.hwp.HAction.Run("CustCopyBtn")

    def CustCutBtn(self):
        """툴바 버튼 오려두기"""
        return self.hwp.HAction.Run("CustCutBtn")

    def CustEraseBtn(self):
        """툴바 버튼 지우기"""
        return self.hwp.HAction.Run("CustEraseBtn")

    def CustInsSepBtn(self):
        """툴바 버튼에 구분선 넣기"""
        return self.hwp.HAction.Run("CustInsSepBtn")

    def CustomizeToolbar(self):
        """도구상자 사용자 설정"""
        return self.hwp.HAction.Run("CustomizeToolbar")

    def CustPasteBtn(self):
        """툴바 버튼 붙여기"""
        return self.hwp.HAction.Run("CustPasteBtn")

    def CustRenameBtn(self):
        """툴바 버튼 이름 바꾸기"""
        return self.hwp.HAction.Run("CustRenameBtn")

    def CustRestBtn(self):
        """툴바 버튼 처음 상태로 되돌리기"""
        return self.hwp.HAction.Run("CustRestBtn")

    def CustViewIconBtn(self):
        """툴바 버튼 아이콘만 보이기"""
        return self.hwp.HAction.Run("CustViewIconBtn")

    def CustViewIconNameBtn(self):
        """툴바 버튼 이름과 아이콘 보이기"""
        return self.hwp.HAction.Run("CustViewIconNameBtn")

    def CustViewNameBtn(self):
        """툴바 버튼 이름만 보이기"""
        return self.hwp.HAction.Run("CustViewNameBtn")

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

    def DeleteBack(self, delete_ctrl:bool=True) -> bool:
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

    def DeleteFieldMemo(self):
        """
        메모 지우기. 누름틀 지우기와 유사하다. 메모 누름틀에 붙어있거나, 메모 안에 들어가 있는 경우 위 액션 실행시 해당 메모가 삭제된다.

        """
        return self.hwp.HAction.Run("DeleteFieldMemo")

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

    def DeletePage(self):
        """
        쪽 지우기

        한글2014 이하의 버전에서는 사용할 수 없다.

        """
        return self.hwp.HAction.Run("DeletePage")

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

    def DrawObjCancelOneStep(self):
        """
        다각형(곡선) 그리는 중 이전 선 지우기.

        현재 사용 안함(?)

        """
        return self.hwp.HAction.Run("DrawObjCancelOneStep")

    def DrawObjEditDetail(self):
        """
        그리기 개체 중 다각형 점편집 액션.

        다각형이 선택된 상태에서만 실행가능.

        """
        return self.hwp.HAction.Run("DrawObjEditDetail")

    def DrawObjOpenClosePolygon(self):
        """
        닫힌 다각형 열기 또는 열린 다각형 닫기 토글.

        ①다각형 개체 선택상태가 아니라 편집상태에서만 위 명령어가 실행된다.

        ②닫힌 다각형을 열 때는 마지막으로 봉합된 점에서 아주 조금만 열린다.

        ③아주 조금만 열린 상태에서 닫으면 노드(꼭지점)가 추가되지 않지만, 적절한 거리를 벌리고 닫기를 하면 추가됨.

        """
        return self.hwp.HAction.Run("DrawObjOpenClosePolygon")

    def DrawObjTemplateSave(self):
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

    def FileClose(self):
        """
        문서 닫기.

        저장 이후 변경사항이 있으면 팝업이 뜨므로 주의
        """
        return self.hwp.HAction.Run("FileClose")

    def FileFind(self):
        """문서 찾기"""
        return self.hwp.HAction.Run("FileFind")

    def FileNew(self):
        """
        새 문서 창을 여는 명령어.

        참고로 현재 창에서 새 탭을 여는 명령어는 ``hwp.FileNewTab()``

        여담이지만 한/글2020 기준으로 새 창은 30개까지 열 수 있다.
        그리고 한 창에는 탭을 30개까지 열 수 있다.
        즉, (리소스만 충분하다면) 동시에 열어서 자동화를 돌릴 수 있는
        문서 갯수는 900개!
        """
        return self.hwp.HAction.Run("FileNew")

    def FileNewTab(self):
        """
        새 탭을 여는 명령어.
        """
        return self.hwp.HAction.Run("FileNewTab")

    def FileNextVersionDiff(self):
        """버전 비교 :　앞으로 이동"""
        return self.hwp.HAction.Run("FileNextVersionDiff")

    def FilePrevVersionDiff(self):
        """버전 비교 : 뒤로 이동"""
        return self.hwp.HAction.Run("FilePrevVersionDiff")

    def FileOpen(self):
        """
        문서를 여는 명령어.

        단 파일선택 팝업이 뜨므로,
        자동화작업시에는 이 명령어를 사용하지 않는다.
        대신 hwp.open(파일명)을 사용해야 한다.
        """
        return self.hwp.HAction.Run("FileOpen")

    def FileOpenMRU(self):
        """
        최근 작업문서 열기

        현재는 FileOpen과 동일한 동작을 하는 것으로 보임.
        사용자입력을 요구하는 팝업이 뜨므로
        자동화에 사용하지 않으며, hwp.open(Path)을 써야 한다.

        """
        return self.hwp.HAction.Run("FileOpenMRU")

    def FilePreview(self):
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

    def FileQuit(self):
        """
        한/글 프로그램을 종료한다.

        단, 저장 이후 문서수정이 있는 경우에는 팝업이 뜨므로,
        ①저장하거나 ②수정내용을 버리는 메서드를 활용해야 한다.

        """
        return self.hwp.HAction.Run("FileQuit")

    def FileSave(self):
        """
        문서 저장(Alt-S).

        가급적 ``hwp.save()``를 사용하자.
        ``hwp.save()``와 ``hwp.FileSave()``에 차이가 있는데
        ``hwp.save()``는 실제 변경이 없으면 저장을 수행하지 않지만
        ``hwp.FileSave()``는 변경이 없어도 저장을 수행하므로 수정일자가 바뀐다.
        """
        return self.hwp.HAction.Run("FileSave")

    def FileSaveAs(self):
        """
        다른 이름으로 저장(Alt-V).

        사용자입력을 필요로 하므로 이 액션은 사용하지 않는다.
        대신 hwp.save_as(Path)를 사용하면 된다.

        """
        return self.hwp.HAction.Run("FileSaveAs")

    def FileSaveAsDRM(self):
        """배포용 문서로 저장하기"""
        return self.hwp.HAction.Run("FileSaveAsDRM")

    def FileVersionDiffChangeAlign(self):
        """버전 비교 : 비교화면 배열 변경 (좌우↔상하)"""
        return self.hwp.HAction.Run("FileVersionDiffChangeAlign")

    def FileVersionDiffSameAlign(self):
        """버전 비교 : 비교화면 다시 정렬"""
        return self.hwp.HAction.Run("FileVersionDiffSameAlign")

    def FileVersionDiffSyncScroll(self):
        """버전 비교 : 비교화면 동시에 이동"""
        return self.hwp.HAction.Run("FileVersionDiffSyncScroll")

    def FillColorShadeDec(self):
        """면색 음영 비율 감소"""
        return self.hwp.HAction.Run("FillColorShadeDec")

    def FillColorShadeInc(self):
        """면색 음영 비율 증가"""
        return self.hwp.HAction.Run("FillColorShadeInc")

    def FindForeBackBookmark(self):
        """
        책갈피 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackBookmark")

    def FindForeBackCtrl(self):
        """
        조판부호 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.
        참고로 ``hwp.FindForeBackSelectCtrl``은 선택.
        """
        return self.hwp.HAction.Run("FindForeBackCtrl")

    def FindForeBackFind(self):
        """
        찾기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackFind")

    def FindForeBackLine(self):
        """
        줄 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackLine")

    def FindForeBackPage(self):
        """
        쪽 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackPage")

    def FindForeBackSection(self):
        """
        구역 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackSection")

    def FindForeBackSelectCtrl(self):
        """앞뒤로 찾아가기 : 조판 부호 찾기 (선택)"""
        return self.hwp.HAction.Run("FindForeBackSelectCtrl")

    def FindForeBackStyle(self):
        """
        스타일 찾아가기.

        사용자 입력을 요구하므로 자동화에는 사용하지 않는다.

        """
        return self.hwp.HAction.Run("FindForeBackStyle")

    def FormDesignMode(self):
        """양식 개체 디자인 모드 변경"""
        return self.hwp.HAction.Run("FormDesignMode")

    def FormObjCreatorCheckButton(self):
        """양식 개체 체크 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorCheckButton")

    def FormObjCreatorComboBox(self):
        """양식 개체 콤보 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorComboBox")

    def FormObjCreatorEdit(self):
        """양식 개체 에디트 박스 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorEdit")

    def FormObjCreatorPushButton(self):
        """양식 개체 푸쉬 버튼 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorPushButton")

    def FormObjCreatorRadioButton(self):
        """양식 개체 라디오 버튼 넣기"""
        return self.hwp.HAction.Run("FormObjCreatorRadioButton")

    def FormObjRadioGroup(self):
        """양식 개체 라디오 버튼 그룹 묶기"""
        return self.hwp.HAction.Run("FormObjRadioGroup")

    def FrameFullScreen(self):
        """
        한/글 프로그램창 전체화면(창 최대화 아님).

        전체화면 해제는 hwp.FrameFullScreenEnd() 또는 hwp.CloseEx()
        """
        return self.hwp.HAction.Run("FrameFullScreen")

    def FrameFullScreenEnd(self):
        """전체 화면 닫기"""
        return self.hwp.HAction.Run("FrameFullScreenEnd")

    def FrameHRuler(self):
        """가로축 눈금자 보이기/감추기"""
        return self.hwp.HAction.Run("FrameHRuler")

    def FrameStatusBar(self):
        """
        한/글 프로그램 하단의 상태바 보이기/숨기기 토글

        """
        return self.hwp.HAction.Run("FrameStatusBar")

    def FrameViewZoomRibon(self):
        """화면 확대/축소"""
        return self.hwp.HAction.Run("FrameViewZoomRibon")

    def FrameVRuler(self):
        """세로축 눈금자 보이기/감추기"""
        return self.hwp.HAction.Run("FrameVRuler")

    def HancomRoom(self):
        """한컴 계약방"""
        return self.hwp.HAction.Run("HancomRoom")

    def HanThDIC(self):
        """
        한/글에 내장되어 있는 "유의어/반의어 사전"을 여는 액션.

        """
        return self.hwp.HAction.Run("HanThDIC")

    def HeaderFooterDelete(self):
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

    def HeaderFooterModify(self):
        """
        머리말/꼬리말 고치기.

        마우스를 쓰지 않고 머리말/꼬리말 편집상태로 들어갈 수 있다.
        단, 커서가 머리말/꼬리말 컨트롤에 닿아 있는 상태에서 실행해야 한다.
        """
        return self.hwp.HAction.Run("HeaderFooterModify")

    def HeaderFooterToNext(self):
        """
        다음 머리말/꼬리말.

        당장은 사용방법을 모르겠다..
        """
        return self.hwp.HAction.Run("HeaderFooterToNext")

    def HeaderFooterToPrev(self):
        """
        이전 머리말.

        당장은 사용방법을 모르겠다..
        """
        return self.hwp.HAction.Run("HeaderFooterToPrev")

    def HelpContents(self):
        """내용"""
        return self.hwp.HAction.Run("HelpContents")

    def HelpIndex(self):
        """찾아보기"""
        return self.hwp.HAction.Run("HelpIndex")

    def HelpWeb(self):
        """온라인 고객 지원"""
        return self.hwp.HAction.Run("HelpWeb")

    def HiddenCredits(self):
        """
        인터넷 정보.

        사용방법을 모르겠다.

        """
        return self.hwp.HAction.Run("HiddenCredits")

    def HideTitle(self):
        """
        차례 숨기기(Ctrl-K-S)

        ([도구 - 차례/색인 - 차례 숨기기] 메뉴에 대응.
        실행한 개요라인을 자동생성되는 제목차례에서 숨긴다.
        즉시 변경되지 않으며,
        ``hwp.UpdateAllContents()``(모든 차례 새로고침, Ctrl-KA) 실행시
        제목차례가 업데이트된다.
        """
        return self.hwp.HAction.Run("HideTitle")

    def HimConfig(self):
        """
        입력기 언어별 환경설정.

        현재는 실행되지 않는 듯 하다.
        대신 ``hwp.HimKbdChange()``로 환경설정창을 띄울 수 있다.
        자동화에는 쓰이지 않는다.

        """
        return self.hwp.HAction.Run("Him Config")

    def HimKbdChange(self):
        """
        입력기 언어별 환경설정.

        """
        return self.hwp.HAction.Run("HimKbdChange")

    def HorzScrollbar(self):
        """가로축 스크롤바 보이기/감추기"""
        return self.hwp.HAction.Run("HorzScrollbar")

    def HwpCtrlEquationCreate97(self):
        """
        한/글97버전 수식 만들기

        실행되지 않는 듯 하다.

        """
        return self.hwp.HAction.Run("HwpCtrlEquationCreate97")

    def HwpCtrlFileNew(self):
        """
        한글컨트롤 전용 새문서.

        실행되지 않는 듯 하다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileNew")

    def HwpCtrlFileOpen(self):
        """
        한글컨트롤 전용 파일 열기.

        실행되지 않는 듯 하다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileOpen")

    def HwpCtrlFileSave(self):
        """
        한글컨트롤 전용 파일 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSave")

    def HwpCtrlFileSaveAs(self):
        """
        한글컨트롤 전용 다른 이름으로 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAs")

    def HwpCtrlFileSaveAsAutoBlock(self):
        """
        한글컨트롤 전용 다른이름으로 블록 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAsAutoBlock")

    def HwpCtrlFileSaveAutoBlock(self):
        """
        한/글 컨트롤 전용 블록 저장.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFileSaveAutoBlock")

    def HwpCtrlFindDlg(self):
        """
        한/글 컨트롤 전용 찾기 대화상자.

        실행되지 않는다.

        """
        return self.hwp.HAction.Run("HwpCtrlFindDlg")

    def HwpCtrlReplaceDlg(self):
        """
        한/글 컨트롤 전용 바꾸기 대화상자

        """
        return self.hwp.HAction.Run("HwpCtrlReplaceDlg")

    def HwpDic(self):
        """
        한컴 사전(F12).

        현재 캐럿이 닿아 있거나, 블록선택한 구간을 검색어에 자동으로 넣는다.

        """
        return self.hwp.HAction.Run("HwpDic")

    def HwpTabViewAction(self):
        """빠른 실행 작업창"""
        return self.hwp.HAction.Run("HwpTabViewAction")

    def HwpTabViewAttribute(self):
        """양식 개체 속성 작업창"""
        return self.hwp.HAction.Run("HwpTabViewAttribute")

    def HwpTabViewClipboard(self):
        """클립보드 작업창"""
        return self.hwp.HAction.Run("HwpTabViewClipboard")

    def HwpTabViewDistant(self):
        """쪽모양 보기 작업창"""
        return self.hwp.HAction.Run("HwpTabViewDistant")

    def HwpTabViewHwpDic(self):
        """사전 검색 작업창"""
        return self.hwp.HAction.Run("HwpTabViewHwpDic")

    def HwpTabViewMasterPage(self):
        """바탕쪽 보기 작업창"""
        return self.hwp.HAction.Run("HwpTabViewMasterPage")

    def HwpTabViewOutline(self):
        """개요 보기 작업창"""
        return self.hwp.HAction.Run("HwpTabViewOutline")

    def HwpTabViewScript(self):
        """스크립트 작업창"""
        return self.hwp.HAction.Run("HwpTabViewScript")

    def HwpViewType(self):
        """문서창 모양 설정"""
        return self.hwp.HAction.Run("HwpViewType")

    def HwpWSDic(self):
        """사전 검색 작업창 (Shift + F12)"""
        return self.hwp.HAction.Run("HwpWSDic")

    def HyperlinkBackward(self):
        """
        하이퍼링크 뒤로.

        하이퍼링크를 통해서 문서를 탐색하여 페이지나 캐럿을 이동한 경우, (브라우저의 "뒤로가기"처럼) 이동 전의 위치로 돌아간다.

        """
        return self.hwp.HAction.Run("HyperlinkBackward")

    def HyperlinkForward(self):
        """
        하이퍼링크 앞으로.

        ``hwp.HyperlinkBackward()`` 에 상반되는 명령어로, 브라우저의 "앞으로 가기"나 한/글의 재실행과 유사하다. 하이퍼링크 등으로 이동한 후에 뒤로가기를 눌렀다면, 캐럿이 뒤로가기 전 위치로 다시 이동한다.

        """
        return self.hwp.HAction.Run("HyperlinkForward")

    def ImageFindPath(self):
        """
        그림 경로 찾기.

        현재는 실행되지 않는 듯.

        """
        return self.hwp.HAction.Run("ImageFindPath")

    def InputCodeChange(self):
        """
        문자/코드 변환

        현재 캐럿의 바로 앞 문자를 찾아서 문자이면 코드로, 코드이면 문자로 변환해준다.(변환 가능한 코드영역 0x0020 ~ 0x10FFFF 까지)

        """
        return self.hwp.HAction.Run("InputCodeChange")

    def InputHanja(self):
        """
        한자로 바꾸기 창을 띄워준다.

        추가입력이 필요하여 자동화에는 쓰이지 않음.

        """
        return self.hwp.HAction.Run("InputHanja")

    def InputHanjaBusu(self):
        """
        부수로 입력.

        자동화에는 쓰이지 않음.

        """
        return self.hwp.HAction.Run("InputHanjaBusu")

    def InputHanjaMean(self):
        """
        한자 새김 입력창 띄우기.

        뜻과 음을 입력하면 적절한 한자를 삽입해준다.입력시 뜻과 음은 붙여서 입력. (예)하늘천

        """
        return self.hwp.HAction.Run("InputHanjaMean")

    def InsertAutoNum(self):
        """
        번호 다시 넣기(?)

        실행이 안되는 듯.

        """
        return self.hwp.HAction.Run("InsertAutoNum")

    def InsertCpNo(self):
        """
        현재 쪽번호(상용구) 삽입.

        쪽번호와 마찬가지로, 문자열이 실시간으로 변경된다.

        ※유의사항 : 이 쪽번호는 찾기, 찾아바꾸기, GetText 및 누름틀 안에 넣고 GetFieldText나 복붙 등 그 어떤 방법으로도 추출되지 않는다.
        한 마디로 눈에는 보이는 것 같지만 실재하지 않는 숫자임. 참고로 표번호도 그렇다. 값이 아니라 속성이라서 그렇다.

        """
        return self.hwp.HAction.Run("InsertCpNo")

    def InsertCpTpNo(self):
        """
        상용구 코드 넣기(현재 쪽/전체 쪽).

        실시간으로 변경된다.

        """
        return self.hwp.HAction.Run("InsertCpTpNo")

    def InsertDateCode(self):
        """
        상용구 코드 넣기(만든 날짜).

        현재날짜가 아님에 유의.

        """
        return self.hwp.HAction.Run("InsertDateCode")

    def InsertDocInfo(self):
        """
        상용구 코드 넣기

        (만든 사람, 현재 쪽, 만든 날짜)

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
        메모고침표 넣기

        (현재 한/글메뉴에 없음, 메모와 동일한 기능)

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

    def LeftShiftBlock(self):
        """텍스트 블록 상태에서 블록 왼쪽에 있는 탭 또는 공백을 지운다."""
        return self.hwp.HAction.Run("LeftShiftBlock")

    def LeftTabFrameClose(self):
        """왼쪽 작업창 감추기"""
        return self.hwp.HAction.Run("LeftTabFrameClose")

    def LinkTextBox(self):
        """
        글상자 연결.

        글상자가 선택되지 않았거나, 캐럿이 글상자 내부에 있지 않으면 동작하지 않는다.

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
        한영 수동 전환

        현재 커서위치 또는 문단나누기 이전에 입력된 내용에 대해서 강제적으로 한/영 전환을 한다.

        """
        return self.hwp.HAction.Run("ManualChangeHangul")

    def MarkPenColor(self):
        """형광펜 색"""
        return self.hwp.HAction.Run("MarkPenColor")

    def MarkTitle(self):
        """
        제목 차례 표시([도구-차례/찾아보기-제목 차례 표시]메뉴에 대응).

        차례 코드가 삽입되어 나중에 차례 만들기에서 사용할 수 있다.
        적용여부는 Ctrl+G,C를 이용해 조판부호를 확인하면 알 수 있다.

        """
        return self.hwp.HAction.Run("MarkTitle")

    def MasterPage(self):
        """
        바탕쪽 진입

        """
        return self.hwp.HAction.Run("MasterPage")

    def MasterPageDuplicate(self):
        """
        기존 바탕쪽과 겹침.

        바탕쪽 편집상태가 활성화되어 있으며 [구역 마지막쪽], [구역임의 쪽]일 경우에만 사용 가능하다.

        """
        return self.hwp.HAction.Run("MasterPageDuplicate")

    def MasterPageExcept(self):
        """
        첫 쪽 제외

        """
        return self.hwp.HAction.Run("MasterPageExcept")

    def MasterPageFront(self):
        """
        바탕쪽 앞으로 보내기.

        바탕쪽 편집모드일 경우에만 동작한다.

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

    def MasterPageType(self):
        """바탕쪽 종류"""
        return self.hwp.HAction.Run("MasterPageType")

    def MasterWsItemOnOff(self):
        """바탕쪽 작업창 보이기/감추기"""
        return self.hwp.HAction.Run("MasterWsItemOnOff")

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
        고치기(채우기 속성 탭으로).

        만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다.
        채우기 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다.

        """
        return self.hwp.HAction.Run("ModifyFillProperty")

    def ModifyLineProperty(self):
        """
        고치기(선/테두리 속성 탭으로).

        만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다.
        선/테두리 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다.

        """
        return self.hwp.HAction.Run("ModifyLineProperty")

    def ModifyShapeObject(self):
        """
        고치기 - 개체 속성

        """
        return self.hwp.HAction.Run("ModifyShapeObject")

    def MoveColumnBegin(self):
        """
        단의 시작점으로 이동

        단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다.

        """
        return self.hwp.HAction.Run("MoveColumnBegin")

    def MoveColumnEnd(self):
        """
        단의 끝점으로 이동한다.

        단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다.

        """
        return self.hwp.HAction.Run("MoveColumnEnd")

    def MoveDocBegin(self):
        """
        문서의 시작으로 이동

        만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다.
        현재 서브 리스트 내에 있으면 빠져나간다. 자동화에 아주 많이 사용된다.

        """
        return self.hwp.HAction.Run("MoveDocBegin")

    def MoveDocEnd(self):
        """
        문서의 끝으로 이동

        만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다.
        현재 서브 리스트 내에 있으면 빠져나간다.

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

    def ParagraphShapeWithNext(self):
        """
        다음 문단과 함께

        """
        return self.hwp.HAction.Run("ParagraphShapeWithNext")

    def ParaShapeLineSpace(self):
        """문단 모양"""
        return self.hwp.HAction.Run("ParaShapeLineSpace")

    def paste(self, option: Literal[0, 1, 2, 3, 4, 5, 6] = 4):
        """
        붙여넣기 확장메서드.

        (참고로 paste가 아닌 Paste는 API 그대로 작동한다.)

        Args:
            option:
                파라미터에 할당할 수 있는 값은 모두 7가지로,

                    - 0: (셀) 왼쪽에 끼워넣기
                    - 1: 오른쪽에 끼워넣기
                    - 2: 위쪽에 끼워넣기
                    - 3: 아래쪽에 끼워넣기
                    - 4: 덮어쓰기
                    - 5: 내용만 덮어쓰기
                    - 6: 셀 안에 표로 넣기
        """
        pset = self.hwp.HParameterSet.HSelectionOpt
        self.hwp.HAction.GetDefault("Paste", pset.HSet)
        pset.option = option
        self.hwp.HAction.Execute("Paste", pset.HSet)

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

    def ShapeObjTableSelCell(self):
        """
        표의 첫 번째 셀 선택

        테이블 선택상태에서 실행해야 한다.
        비슷한 명령어로 ``hwp.ShapeObjTextBoxEdit()``가 있다.
        ``hwp.ShapeObjTableSelCell()``이 셀블록상태인데 반해
        ``hwp.ShapeObjTextBoxEdit()``는 편집상태이다.
        """
        # return self.hwp.HAction.Run("ShapeObjTableSelCell")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("ShapeObjTableSelCell", pset.HSet)
        return self.HAction.Execute("ShapeObjTableSelCell", pset.HSet)

    def ShapeObjTextBoxEdit(self):
        """
        표나 글상자 선택상태에서 편집모드로 들어가기

        표를 선택하고 있는 경우 A1 셀 안으로 이동한다.

        """
        return self.hwp.HAction.Run("ShapeObjTextBoxEdit")

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

    def TableDeleteCell(self, remain_cell:bool=False) -> bool:
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
        # return self.hwp.HAction.Run("TableRightCell")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("TableRightCell", pset.HSet)
        return self.HAction.Execute("TableRightCell", pset.HSet)

    def TableRightCellAppend(self):
        """
        셀 이동: 셀 오른쪽에 이어서

        우측 셀로 이동하다 끝에 도달하면 다음 행의 첫 번째 셀로 이동.
        그리고 다음 행이 없는 경우에는 새 행을 아래 추가하고 첫 번째 셀로 이동.

        """
        # return self.hwp.HAction.Run("TableRightCellAppend")
        pset = self.HParameterSet.HInsertText
        self.HAction.GetDefault("TableRightCellAppend", pset.HSet)
        return self.HAction.Execute("TableRightCellAppend", pset.HSet)

    def TableSplitCell(self, Rows:int=2, Cols:int=0, DistributeHeight:int=0, Merge:int=0) -> bool:
        """
        셀 나누기.

        Args:
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

    def run_script_macro(self, function_name:str, u_macro_type:int=0, u_script_type:int=0) -> bool:
        """
        한/글 문서 내에 존재하는 매크로를 실행한다.

        문서매크로, 스크립트매크로 모두 실행 가능하다.
        재미있는 점은 한/글 내에서 문서매크로 실행시
        New, Open 두 개의 함수 밖에 선택할 수 없으므로
        별도의 함수를 정의하더라도 이 두 함수 중 하나에서 호출해야 하지만,
        (진입점이 되어야 함)
        self.hwp.run_script_macro 명령어를 통해서는 제한없이 실행할 수 있다.

        Args:
            function_name: 실행할 매크로 함수이름(전체이름)
            u_macro_type:
                매크로의 유형. 밑의 값 중 하나이다.

                    - 0: 스크립트 매크로(전역 매크로-HWP_GLOBAL_MACRO_TYPE, 기본값)
                    - 1: 문서 매크로(해당문서에만 저장/적용되는 매크로-HWP_DOCUMENT_MACRO_TYPE)

            u_script_type:
                스크립트의 유형. 현재는 javascript만을 유일하게 지원한다.
                아무 정수나 입력하면 된다. (기본값: 0)

        Returns:
            무조건 True를 반환(매크로의 실행여부와 상관없음)

        Examples:
            >>> hwp.run_script_macro("OnDocument_New", u_macro_type=1)
            True
            >>> hwp.run_script_macro("OnScriptMacro_중국어1성")
            True
        """
        return self.hwp.RunScriptMacro(FunctionName=function_name, uMacroType=u_macro_type, uScriptType=u_script_type)

    def RunScriptMacro(self, function_name:str, u_macro_type:int=0, u_script_type:int=0) -> bool:
        """
        한/글 문서 내에 존재하는 매크로를 실행한다.

        문서매크로, 스크립트매크로 모두 실행 가능하다.
        재미있는 점은 한/글 내에서 문서매크로 실행시
        New, Open 두 개의 함수 밖에 선택할 수 없으므로
        별도의 함수를 정의하더라도 이 두 함수 중 하나에서 호출해야 하지만,
        (진입점이 되어야 함)
        self.hwp.run_script_macro 명령어를 통해서는 제한없이 실행할 수 있다.

        Args:
            function_name: 실행할 매크로 함수이름(전체이름)
            u_macro_type:
                매크로의 유형. 밑의 값 중 하나이다.

                    - 0: 스크립트 매크로(전역 매크로-HWP_GLOBAL_MACRO_TYPE, 기본값)
                    - 1: 문서 매크로(해당문서에만 저장/적용되는 매크로-HWP_DOCUMENT_MACRO_TYPE)

            u_script_type:
                스크립트의 유형. 현재는 javascript만을 유일하게 지원한다.
                아무 정수나 입력하면 된다. (기본값: 0)

        Returns:
            무조건 True를 반환(매크로의 실행여부와 상관없음)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.run_script_macro("OnDocument_New", u_macro_type=1)
            True
            >>> hwp.run_script_macro("OnScriptMacro_중국어1성")
            True
        """
        return self.hwp.RunScriptMacro(FunctionName=function_name, uMacroType=u_macro_type, uScriptType=u_script_type)

    def save(self, save_if_dirty:bool=True) -> bool:
        """
        현재 편집중인 문서를 저장한다.

        문서의 경로가 지정되어있지 않으면 “새 이름으로 저장” 대화상자가 뜬다.

        Args:
            save_if_dirty:
                True를 지정하면 문서가 변경된 경우에만 저장한다.
                False를 지정하면 변경여부와 상관없이 무조건 저장한다.
                생략하면 True가 지정된다.

        Returns:
            성공하면 True, 실패하면 False
        """
        return self.hwp.Save(save_if_dirty=save_if_dirty)

    def Save(self, save_if_dirty:bool=True) -> bool:
        """
        현재 편집중인 문서를 저장한다.

        문서의 경로가 지정되어있지 않으면 “새 이름으로 저장” 대화상자가 뜬다.

        Args:
            save_if_dirty:
                True를 지정하면 문서가 변경된 경우에만 저장한다.
                False를 지정하면 변경여부와 상관없이 무조건 저장한다.
                생략하면 True가 지정된다.

        Returns:
            성공하면 True, 실패하면 False
        """
        return self.hwp.Save(save_if_dirty=save_if_dirty)

    def save_as(self, path:str, format:str="HWP", arg:str="", split_page:bool=False) -> bool:
        """
        현재 편집중인 문서를 지정한 이름으로 저장한다.

        format, arg의 일반적인 개념에 대해서는 Open()참조.

        Args:
            path: 문서 파일의 전체경로
            format: 문서 형식. 생략하면 "HWP"가 지정된다.
            arg:
                세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.
                여러 개를 한꺼번에 할 경우에는 세미콜론으로 구분하여 연속적으로 사용할 수 있다.

                "lock:TRUE;backup:FALSE;prvtext:1"

                    - "lock:true": 저장한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
                    - "backup:false": 백업 파일 생성 여부
                    - "compress:true": 압축 여부
                    - "fullsave:false": 스토리지 파일을 완전히 새로 생성하여 저장
                    - "prvimage:2": 미리보기 이미지 (0=off, 1=BMP, 2=GIF)
                    - "prvtext:1": 미리보기 텍스트 (0=off, 1=on)
                    - "autosave:false": 자동저장 파일로 저장할 지 여부 (TRUE: 자동저장, FALSE: 지정 파일로 저장)
                    - "export": 다른 이름으로 저장하지만 열린 문서는 바꾸지 않는다.(lock:false와 함께 설정되어 있을 시 동작)

            split_page: html+ 포맷으로 저장할 때, 페이지 나누기 여부

        Returns:
            성공하면 True, 실패하면 False
        """

        if path.lower()[1] != ":":
            path = os.path.abspath(path)
        ext = path.rsplit(".", maxsplit=1)[-1]
        if format.lower() == "html+":  # 서식 있는 인터넷 문서
            # 키 코드 상수
            VK_SHIFT = 0x10  # Shift 키
            VK_CONTROL = 0x11  # Ctrl 키
            VK_MENU = 0x12  # Alt 키
            VK_D = 0x44  # D 키
            VK_LEFT = 0x25  # 왼쪽 화살표 키
            VK_UP = 0x26  # 위쪽 화살표 키
            VK_RIGHT = 0x27  # 오른쪽 화살표 키
            VK_DOWN = 0x28  # 아래쪽 화살표 키

            # SendInput 관련 구조체 정의는 이전 코드와 동일
            PUL = ctypes.POINTER(ctypes.c_ulong)

            class KeyBdInput(ctypes.Structure):
                _fields_ = [("wVk", ctypes.c_ushort),
                            ("wScan", ctypes.c_ushort),
                            ("dwFlags", ctypes.c_ulong),
                            ("time", ctypes.c_ulong),
                            ("dwExtraInfo", PUL)]

            class HardwareInput(ctypes.Structure):
                _fields_ = [("uMsg", ctypes.c_ulong),
                            ("wParamL", ctypes.c_short),
                            ("wParamH", ctypes.c_ushort)]

            class MouseInput(ctypes.Structure):
                _fields_ = [("dx", ctypes.c_long),
                            ("dy", ctypes.c_long),
                            ("mouseData", ctypes.c_ulong),
                            ("dwFlags", ctypes.c_ulong),
                            ("time", ctypes.c_ulong),
                            ("dwExtraInfo", PUL)]

            class Input_I(ctypes.Union):
                _fields_ = [("ki", KeyBdInput),
                            ("mi", MouseInput),
                            ("hi", HardwareInput)]

            class Input(ctypes.Structure):
                _fields_ = [("type", ctypes.c_ulong),
                            ("ii", Input_I)]

            # 키를 누르는 함수
            def press_key(hexKeyCode):
                extra = ctypes.c_ulong(0)
                ii_ = Input_I()
                ii_.ki = KeyBdInput(hexKeyCode, 0, 0, 0, ctypes.pointer(extra))
                x = Input(ctypes.c_ulong(1), ii_)
                ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

            # 키를 떼는 함수
            def release_key(hexKeyCode):
                extra = ctypes.c_ulong(0)
                ii_ = Input_I()
                ii_.ki = KeyBdInput(hexKeyCode, 0, 0x0002, 0, ctypes.pointer(extra))
                x = Input(ctypes.c_ulong(1), ii_)
                ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

            def find_window_and_send_key(window_name, key_code, retries=5, delay=0.1):
                for attempt in range(retries):
                    try:
                        hwnd = win32gui.FindWindow(None, window_name)
                        if hwnd == 0:
                            raise Exception(f"{window_name} 창을 찾을 수 없습니다.")

                        # 창을 포커스로 설정
                        win32gui.SetForegroundWindow(hwnd)
                        sleep(delay)  # 포커스 설정 후 약간의 지연

                        # 키 입력 전송
                        press_key(key_code)
                        sleep(0.05)
                        release_key(key_code)
                        sleep(0.05)

                        return True

                    except Exception as e:
                        # print(f"Attempt {attempt + 1} failed: {e}")
                        sleep(delay)  # 다음 시도 전 지연

                return False

            def find_window_and_confirm(window_name, retries=5, delay=0.1):
                key = "D"
                for attempt in range(retries):
                    try:
                        hwnd = win32gui.FindWindow(None, window_name)
                        if hwnd == 0:
                            raise Exception("Window not found")

                        # 창을 포커스로 설정
                        win32gui.SetForegroundWindow(hwnd)
                        sleep(delay)  # 포커스 설정 후 약간의 지연

                        # 키 입력 전송
                        for char in key:
                            vk_code = ord(char.upper())  # 가상 키 코드로 변환
                            press_key(vk_code)
                            sleep(0.05)
                            release_key(vk_code)
                            sleep(0.05)

                        return True

                    except Exception as e:
                        # print(f"Attempt {attempt + 1} failed: {e}")
                        sleep(delay)  # 다음 시도 전 지연

                return False

            def save_as_html_plus(path, visible=True):
                pythoncom.CoInitialize()
                hwp = Hwp(visible=visible)
                pset = hwp.HParameterSet.HFileOpenSave
                hwp.HAction.GetDefault("FileSaveAs_S", pset.HSet)
                pset.filename = path
                pset.Format = "HTML+"
                hwp.HAction.Execute("FileSaveAs_S", pset.HSet)
                find_window_and_send_key("서식 있는 인터넷 문서 종류", VK_UP)
                pythoncom.CoUninitialize()
                return True

            t = threading.Thread(target=save_as_html_plus, args=(path, True))
            t.start()
            t.join(timeout=0)
            if split_page:
                find_window_and_send_key("서식 있는 인터넷 문서 종류", VK_UP)
            find_window_and_confirm("서식 있는 인터넷 문서 종류")
            return True

        if ext.lower() == "pdf" or format.lower() == "pdf":
            pset = self.HParameterSet.HFileOpenSave
            self.HAction.GetDefault("FileSaveAs_S", pset.HSet)
            pset.filename = path
            pset.Format = "PDF"
            pset.Attributes = 0
            if not self.HAction.Execute("FileSaveAs_S", pset.HSet):
                pset = self.HParameterSet.HFileOpenSave
                self.HAction.GetDefault("FileSaveAsPdf", pset.HSet)
                self.HParameterSet.HFileOpenSave.filename = path
                self.HParameterSet.HFileOpenSave.Format = "PDF"
                self.HParameterSet.HFileOpenSave.Attributes = 16384
                return self.HAction.Execute("FileSaveAsPdf", pset.HSet)
            else:
                return True
        elif ext.lower() == "hwpx" or format.lower() == "hwpx":
            return self.hwp.SaveAs(Path=path, Format="HWPX", arg=arg)
        else:
            return self.hwp.SaveAs(Path=path, Format=format, arg=arg)

    def SaveAs(self, path:str, format:str="HWP", arg:str="") -> bool:
        """
        현재 편집중인 문서를 지정한 이름으로 저장한다.

        format, arg의 일반적인 개념에 대해서는 Open()참조.

        Args:
            path: 문서 파일의 전체경로
            format: 문서 형식. 생략하면 "HWP"가 지정된다.
            arg:
                세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.
                여러 개를 한꺼번에 할 경우에는 세미콜론으로 구분하여 연속적으로 사용할 수 있다.

                "lock:TRUE;backup:FALSE;prvtext:1"

                    - "lock:true": 저장한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부
                    - "backup:false": 백업 파일 생성 여부
                    - "compress:true": 압축 여부
                    - "fullsave:false": 스토리지 파일을 완전히 새로 생성하여 저장
                    - "prvimage:2": 미리보기 이미지 (0=off, 1=BMP, 2=GIF)
                    - "prvtext:1": 미리보기 텍스트 (0=off, 1=on)
                    - "autosave:false": 자동저장 파일로 저장할 지 여부 (TRUE: 자동저장, FALSE: 지정 파일로 저장)
                    - "export": 다른 이름으로 저장하지만 열린 문서는 바꾸지 않는다.(lock:false와 함께 설정되어 있을 시 동작)

        Returns:
            성공하면 True, 실패하면 False
        """
        if path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)
        return self.hwp.SaveAs(Path=path, Format=format, arg=arg)

    def scan_font(self):
        return self.hwp.ScanFont()

    def ScanFont(self):
        return self.hwp.ScanFont()

    def select_text_by_get_pos(self, s_getpos:tuple, e_getpos:tuple) -> bool:
        """
        hwp.get_pos()로 얻은 두 튜플 사이의 텍스트를 선택하는 메서드.
        """
        self.set_pos(s_getpos[0], 0, 0)
        return self.hwp.SelectText(spara=s_getpos[1], spos=s_getpos[2], epara=e_getpos[1], epos=e_getpos[2])

    def select_text(self, spara:Union[int, list, tuple]=0, spos:int=0, epara:int=0, epos:int=0, slist:int=0) -> bool:
        """
        특정 범위의 텍스트를 블록선택한다.

        epos가 가리키는 문자는 포함되지 않는다.

        Args:
            spara: 블록 시작 위치의 문단 번호.
            spos: 블록 시작 위치의 문단 중에서 문자의 위치.
            epara: 블록 끝 위치의 문단 번호.
            epos: 블록 끝 위치의 문단 중에서 문자의 위치.

        Returns:
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

    def SelectText(self, spara: Union[int, list, tuple] = 0, spos:int=0, epara:int=0, epos:int=0, slist:int=0) -> bool:
        """
        특정 범위의 텍스트를 블록선택한다.

        epos가 가리키는 문자는 포함되지 않는다.

        Args:
            spara: 블록 시작 위치의 문단 번호.
            spos: 블록 시작 위치의 문단 중에서 문자의 위치.
            epara: 블록 끝 위치의 문단 번호.
            epos: 블록 끝 위치의 문단 중에서 문자의 위치.

        Returns:
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

    def set_cur_field_name(self, field: str = "", direction: str = "", memo: str = "", option: int = 0) -> bool:
        """
        표 안에서 현재 캐럿이 위치하는 셀, 또는 블록선택한 셀들의 필드이름을 설정한다.

        GetFieldList()의 옵션 중에 4(hwpFieldSelection) 옵션은 사용하지 않는다.

        셀필드가 아닌 누름틀 생성은 `create_field` 메서드를 이용해야 한다.

        Args:
            field: 데이터 필드 이름
            direction: 누름틀 필드의 안내문. 누름틀 필드일 때만 유효하다.
            memo: 누름틀 필드의 메모. 누름틀 필드일 때만 유효하다.
            option:
                다음과 같은 옵션을 지정할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다.

                    - 1: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다.(hwpFieldCell)
                    - 2: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_table(5, 5, True)  # 3행3열의 글자처럼 취급한 표 생성(A1셀로 이동)
            >>> hwp.TableCellBlockExtendAbs()
            >>> hwp.TableCellBlockExtend()  # 셀 전체 선택
            >>> hwp.set_cur_field_name("target_table")  # 모든 셀의 셀필드 이름을 "target_table"로 바꿈
            >>> hwp.put_field_text("target_table", list(range(1, 26)))  # 각 셀에 1~25까지의 정수를 넣음
            >>> hwp.set_cur_field_name("")  # 셀필드 초기화
            >>> hwp.Cancel()  # 셀블록 선택취소
        """
        if not self.is_cell():
            raise AssertionError("캐럿이 표 안에 있지 않습니다.")

        if self.SelectionMode in [0, 3, 19]:
            # 셀 안에서 아무것도 선택하지 않았거나, 하나 이상의 셀블록 상태일 때
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 0)
            pset.ShapeTableCell.CellCtrlData.name = field
            return self.HAction.Execute("TablePropertyDialog", pset.HSet)
        else:
            return self.hwp.SetCurFieldName(Field=field, option=option, Direction=direction, memo=memo)

    def SetCurFieldName(self, field: str = "", direction: str = "", memo: str = "", option: int = 0) -> bool:
        """
        표 안에서 현재 캐럿이 위치하는 셀, 또는 블록선택한 셀들의 필드이름을 설정한다.

        GetFieldList()의 옵션 중에 4(hwpFieldSelection) 옵션은 사용하지 않는다.

        셀필드가 아닌 누름틀 생성은 `create_field` 메서드를 이용해야 한다.

        Args:
            field: 데이터 필드 이름
            direction: 누름틀 필드의 안내문. 누름틀 필드일 때만 유효하다.
            memo: 누름틀 필드의 메모. 누름틀 필드일 때만 유효하다.
            option:
                다음과 같은 옵션을 지정할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다.

                    - 1: 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다.(hwpFieldCell)
                    - 2: 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다.(hwpFieldClickHere)

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.create_table(5, 5, True)  # 3행3열의 글자처럼 취급한 표 생성(A1셀로 이동)
            >>> hwp.TableCellBlockExtendAbs()
            >>> hwp.TableCellBlockExtend()  # 셀 전체 선택
            >>> hwp.set_cur_field_name("target_table")  # 모든 셀의 셀필드 이름을 "target_table"로 바꿈
            >>> hwp.put_field_text("target_table", list(range(1, 26)))  # 각 셀에 1~25까지의 정수를 넣음
            >>> hwp.set_cur_field_name("")  # 셀필드 초기화
            >>> hwp.Cancel()  # 셀블록 선택취소
        """
        if not self.is_cell():
            raise AssertionError("캐럿이 표 안에 있지 않습니다.")
        if self.SelectionMode in [0, 3, 19]:
            # 셀 안에서 아무것도 선택하지 않았거나, 하나 이상의 셀블록 상태일 때
            pset = self.HParameterSet.HShapeObject
            self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", 0)
            pset.ShapeTableCell.CellCtrlData.name = field
            return self.HAction.Execute("TablePropertyDialog", pset.HSet)
        else:
            return self.hwp.SetCurFieldName(Field=field, option=option, Direction=direction, memo=memo)

    def set_field_view_option(self, option:int) -> int:
        """
        양식모드와 읽기전용모드일 때 현재 열린 문서의 필드의 겉보기 속성(『』표시)을 바꾼다.

        EditMode와 비슷하게 현재 열려있는 문서에 대한 속성이다. 따라서 저장되지 않는다.
        (작동하지 않음)

        Args:
            option:
                겉보기 속성 bit

                - 1: 누름틀의 『』을 표시하지 않음, 기타필드의 『』을 표시하지 않음
                - 2: 누름틀의 『』을 빨간색으로 표시, 기타필드의 『』을 흰색으로 표시(기본값)
                - 3: 누름틀의 『』을 흰색으로 표시, 기타필드의 『』을 흰색으로 표시

        Returns:
            설정된 속성이 반환된다. 에러일 경우 0이 반환된다.
        """
        return self.hwp.SetFieldViewOption(option=option)

    def SetFieldViewOption(self, option:int) -> bool:
        """
        양식모드와 읽기전용모드일 때 현재 열린 문서의 필드의 겉보기 속성(『』표시)을 바꾼다.

        EditMode와 비슷하게 현재 열려있는 문서에 대한 속성이다. 따라서 저장되지 않는다.
        (작동하지 않음)

        Args:
            option:
                겉보기 속성 bit

                - 1: 누름틀의 『』을 표시하지 않음, 기타필드의 『』을 표시하지 않음
                - 2: 누름틀의 『』을 빨간색으로 표시, 기타필드의 『』을 흰색으로 표시(기본값)
                - 3: 누름틀의 『』을 흰색으로 표시, 기타필드의 『』을 흰색으로 표시

        Returns:
            설정된 속성이 반환된다. 에러일 경우 0이 반환된다.
        """
        return self.hwp.SetFieldViewOption(option=option)

    def set_message_box_mode(self, mode:int) -> int:
        """
        메시지박스 버튼 자동클릭

        한/글에서 쓰는 다양한 메시지박스가 뜨지 않고,

        자동으로 특정 버튼을 클릭한 효과를 주기 위해 사용한다.
        한/글에서 한/글이 로드된 후 SetMessageBoxMode()를 호출해서 사용한다.
        SetMessageBoxMode는 하나의 파라메터를 받으며,
        해당 파라메터는 자동으로 스킵할 버튼의 값으로 설정된다.
        예를 들어, MB_OK_IDOK (0x00000001)값을 주면,
        MB_OK형태의 메시지박스에서 OK버튼이 눌린 효과를 낸다.

        Args:
            mode:
                메시지 박스 자동선택 종류

                0. 모든 자동설정 해제: 0xFFFFFF

                1. 확인 버튼만 있는 팝업의 경우

                    - 확인 자동누르기: 0x1
                    - 확인 자동누르기 해제: 0xF

                2. 확인/취소 버튼이 있는 팝업의 경우

                    - 확인 자동누르기: 0x10
                    - 취소 자동누르기: 0x20
                    - 확인/취소 옵션 해제: 0xF0

                3. 종료/재시도/무시 팝업의 경우

                    - 종료 자동누르기: 0x100
                    - 재시도 자동누르기: 0x200
                    - 무시 자동누르기: 0x400
                    - 종료/재시도/무시 옵션 해제: 0xF00

                4. 예/아니오/취소 팝업의 경우

                    - 예 자동누르기: 0x1000
                    - 아니오 자동누르기: 0x2000
                    - 취소 자동누르기: 0x4000
                    - 예/아니오/취소 옵션 해제: 0xF000

                5. 예/아니오 팝업의 경우

                    - 예 자동누르기: 0x10000
                    - 아니오 자동누르기: 0x20000
                    - 예/아니오 옵션 해제: 0xF0000

                6. 재시도/취소 팝업의 경우

                    - 재시도 자동누르기: 0x100000
                    - 취소 자동누르기: 0x200000
                    - 재시도/취소 옵션 해제: 0xF00000

        Returns:
            실행 직전의 MessageBoxMode(현재 값이 아님에 주의)
        """
        return self.hwp.SetMessageBoxMode(Mode=mode)

    def SetMessageBoxMode(self, mode:int) -> int:
        """
        메시지박스 버튼 자동클릭

        한/글에서 쓰는 다양한 메시지박스가 뜨지 않고,

        자동으로 특정 버튼을 클릭한 효과를 주기 위해 사용한다.
        한/글에서 한/글이 로드된 후 SetMessageBoxMode()를 호출해서 사용한다.
        SetMessageBoxMode는 하나의 파라메터를 받으며,
        해당 파라메터는 자동으로 스킵할 버튼의 값으로 설정된다.
        예를 들어, MB_OK_IDOK (0x00000001)값을 주면,
        MB_OK형태의 메시지박스에서 OK버튼이 눌린 효과를 낸다.

        Args:
            mode:
                메시지 박스 자동선택 종류

                0. 모든 자동설정 해제: 0xFFFFFF

                1. 확인 버튼만 있는 팝업의 경우

                    - 확인 자동누르기: 0x1
                    - 확인 자동누르기 해제: 0xF

                2. 확인/취소 버튼이 있는 팝업의 경우

                    - 확인 자동누르기: 0x10
                    - 취소 자동누르기: 0x20
                    - 확인/취소 옵션 해제: 0xF0

                3. 종료/재시도/무시 팝업의 경우

                    - 종료 자동누르기: 0x100
                    - 재시도 자동누르기: 0x200
                    - 무시 자동누르기: 0x400
                    - 종료/재시도/무시 옵션 해제: 0xF00

                4. 예/아니오/취소 팝업의 경우

                    - 예 자동누르기: 0x1000
                    - 아니오 자동누르기: 0x2000
                    - 취소 자동누르기: 0x4000
                    - 예/아니오/취소 옵션 해제: 0xF000

                5. 예/아니오 팝업의 경우

                    - 예 자동누르기: 0x10000
                    - 아니오 자동누르기: 0x20000
                    - 예/아니오 옵션 해제: 0xF0000

                6. 재시도/취소 팝업의 경우

                    - 재시도 자동누르기: 0x100000
                    - 취소 자동누르기: 0x200000
                    - 재시도/취소 옵션 해제: 0xF00000

        Returns:
            실행 직전의 MessageBoxMode(현재 값이 아님에 주의)
        """
        return self.hwp.SetMessageBoxMode(Mode=mode)

    def set_pos(self, List: int, para: int, pos: int) -> bool:
        """
        캐럿을 문서 내 특정 위치로 옮기기

        지정된 좌표로 캐럿을 옮겨준다.

        Args:
            List: 캐럿이 위치한 문서 내 list ID
            para: 캐럿이 위치한 문단 ID. 음수거나, 범위를 넘어가면 문서의 시작으로 이동하며, pos는 무시한다.
            pos: 캐럿이 위치한 문단 내 글자 위치. -1을 주면 해당문단의 끝으로 이동한다. 단 para가 범위 밖일 경우 pos는 무시되고 문서의 시작으로 캐럿을 옮긴다.

        Returns:
            성공하면 True, 실패하면 False
        """
        self.hwp.SetPos(List=List, Para=para, pos=pos)
        if (List, para) == self.get_pos()[:2]:
            return True
        else:
            return False

    def SetPos(self, List: int, para: int, pos: int) -> bool:
        """
        캐럿을 문서 내 특정 위치로 옮기기

        지정된 위치로 캐럿을 옮겨준다.

        Args:
            List: 캐럿이 위치한 문서 내 list ID
            para: 캐럿이 위치한 문단 ID. 음수거나, 범위를 넘어가면 문서의 시작으로 이동하며, pos는 무시한다.
            pos: 캐럿이 위치한 문단 내 글자 위치. -1을 주면 해당문단의 끝으로 이동한다. 단 para가 범위 밖일 경우 pos는 무시되고 문서의 시작으로 캐럿을 옮긴다.

        Returns:
            성공하면 True, 실패하면 False
        """
        self.hwp.SetPos(List=List, Para=para, pos=pos)
        if (List, para) == self.get_pos()[:2]:
            return True
        else:
            return False

    def set_pos_by_set(self, disp_val:Any) -> bool:
        """
        캐럿을 ParameterSet으로 얻어지는 위치로 옮긴다.

        Args:
            disp_val: 캐럿을 옮길 위치에 대한 ParameterSet 정보

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> start_pos = hwp.GetPosBySet()  # 현재 위치를 체크포인트처럼 저장하고,
            >>> # 특정 작업(이동 및 입력작업) 후에
            >>> hwp.set_pos_by_set(start_pos)  # 저장했던 위치로 되돌아가기
        """
        return self.hwp.SetPosBySet(dispVal=disp_val)

    def SetPosBySet(self, disp_val:"Hwp.HParameterSet") -> bool:
        """
        캐럿을 ParameterSet으로 얻어지는 위치로 옮긴다.

        Args:
            disp_val: 캐럿을 옮길 위치에 대한 ParameterSet 정보

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> start_pos = hwp.GetPosBySet()  # 현재 위치를 저장하고,
            >>> hwp.set_pos_by_set(start_pos)  # 특정 작업 후에 저장위치로 재이동
        """
        return self.hwp.SetPosBySet(dispVal=disp_val)

    def set_private_info_password(self, password:str) -> bool:
        """
        개인정보보호를 위한 암호를 등록한다.

        개인정보 보호를 설정하기 위해서는
        우선 개인정보 보호 암호를 먼저 설정해야 한다.
        그러므로 개인정보 보호 함수를 실행하기 이전에
        반드시 이 함수를 호출해야 한다.
        (현재 작동하지 않는다.)

        Args:
            password: 새 암호

        Returns:
            정상적으로 암호가 설정되면 True를 반환한다.
            암호설정에 실패하면 false를 반환한다. false를 반환하는 경우는 다음과 같다
            - 암호의 길이가 너무 짧거나 너무 길 때 (영문 5~44자, 한글 3~22자)
            - 암호가 이미 설정되었음. 또는 암호가 이미 설정된 문서임
        """
        return self.hwp.SetPrivateInfoPassword(Password=password)

    def SetPrivateInfoPassword(self, password:str) -> bool:
        """
        개인정보보호를 위한 암호를 등록한다.

        개인정보 보호를 설정하기 위해서는
        우선 개인정보 보호 암호를 먼저 설정해야 한다.
        그러므로 개인정보 보호 함수를 실행하기 이전에
        반드시 이 함수를 호출해야 한다.
        (현재 작동하지 않는다.)

        Args:
            password: 새 암호

        Returns:
            정상적으로 암호가 설정되면 true를 반환한다.
            암호설정에 실패하면 false를 반환한다. false를 반환하는 경우는 다음과 같다
            - 암호의 길이가 너무 짧거나 너무 길 때 (영문 5~44자, 한글 3~22자)
            - 암호가 이미 설정되었음. 또는 암호가 이미 설정된 문서임
        """
        return self.hwp.SetPrivateInfoPassword(Password=password)

    def set_text_file(self, data: str, format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X", option:str="insertfile") -> int:
        """
        GetTextFile로 저장한 문자열 정보를 문서에 삽입

        Args:
            data: 문자열로 변경된 text 파일
            format:
                파일의 형식

                    - "HWP": HWP native format. BASE64 로 인코딩되어 있어야 한다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다.ver:0x0505010B
                    - "HWPML2X": HWP 형식과 호환. 문서의 모든 정보를 유지
                    - "HTML": 인터넷 문서 HTML 형식. 한/글 고유의 서식은 손실된다.
                    - "UNICODE": 유니코드 텍스트, 서식정보가 없는 텍스트만 저장
                    - "TEXT": 일반 텍스트, 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다.
            option: "insertfile"을 지정하면 현재커서 이후에 지정된 파일을 삽입(기본값)

        Returns:
            성공이면 1을, 실패하면 0을 반환한다.
        """
        return self.hwp.SetTextFile(data=data, Format=format, option=option)

    def SetTextFile(self, data: str, format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X", option:str="insertfile") -> int:
        """
        GetTextFile로 저장한 문자열 정보를 문서에 삽입

        Args:
            data: 문자열로 변경된 text 파일
            format:
                파일의 형식

                    - "HWP": HWP native format. BASE64 로 인코딩되어 있어야 한다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다.ver:0x0505010B
                    - "HWPML2X": HWP 형식과 호환. 문서의 모든 정보를 유지
                    - "HTML": 인터넷 문서 HTML 형식. 한/글 고유의 서식은 손실된다.
                    - "UNICODE": 유니코드 텍스트, 서식정보가 없는 텍스트만 저장
                    - "TEXT": 일반 텍스트, 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다.
            option: "insertfile"을 지정하면 현재커서 이후에 지정된 파일을 삽입(기본값)

        Returns:
            성공이면 1을, 실패하면 0을 반환한다.
        """
        return self.hwp.SetTextFile(data=data, Format=format, option=option)

    def get_title(self) -> str:
        """
        한/글 프로그램의 타이틀을 조회한다. 내부적으로 윈도우핸들을 이용한다.

        SetTitleName이라는 못난 이름의 API가 있는데, 차마 get_title_name이라고 따라짓지는 못했다ㅜ
        (파일명을 조회하려면 title 대신 Path나 FullName 등을 조회하면 된다.)

        Returns:
            한/글 창의 상단 타이틀. Path와 달리 빈 문서 상태라도 "빈 문서 1 - 한글" 문자열을 리턴한다.

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> print(hwp.get_title())
            빈 문서 1 - 한글
        """
        return win32gui.GetWindowText(self.hwp.XHwpWindows.Active_XHwpWindow.WindowHandle)

    def set_title_name(self, title: str = "") -> bool:
        """
        한/글 프로그램의 타이틀을 변경한다.

        파일명과 무관하게 설정할 수 있으며, 이모지 등 모든 특수문자를 허용한다. 단, 끝에는 항상 " - 한글"이 붙는다.
        타이틀을 빈 문자열로 만들면 자동으로 원래 타이틀로 돌아간다.

        Args:
            title: 변경할 타이틀 문자열

        Returns:
            성공시 True

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.open("asdf.hwp")
            >>> hwp.get_title()
            asdf.hwp [c:\\Users\\user\\desktop\\] - 한글
            >>> hwp.set_title_name("😘")
            >>> hwp.get_title()
            😘 - 한글
        """
        return self.hwp.SetTitleName(Title=title)

    def SetTitleName(self, title: str = "") -> bool:
        """
        한/글 프로그램의 타이틀을 변경한다.

        파일명과 무관하게 설정할 수 있으며, 이모지 등 모든 특수문자를 허용한다. 단, 끝에는 항상 " - 한글"이 붙는다.
        타이틀을 빈 문자열로 만들면 자동으로 원래 타이틀로 돌아간다.

        Args:
            title: 변경할 타이틀 문자열

        Returns:
            성공시 True

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.open("asdf.hwp")
            >>> hwp.get_title()
            asdf.hwp [c:\\Users\\user\\desktop\\] - 한글
            >>> hwp.SetTitleName("😘")
            >>> hwp.get_title()
            😘 - 한글
        """
        return self.hwp.SetTitleName(Title=title)

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

    #### 파라미터 헬퍼메서드 : 별도의 동작은 하지 않고, 파라미터 변환, 연산 등을 돕는다. ####

    def BorderShape(self, border_type):
        return self.hwp.BorderShape(BorderType=border_type)

    def ArcType(self, arc_type):
        return self.hwp.ArcType(ArcType=arc_type)

    def AutoNumType(self, autonum):
        return self.hwp.AutoNumType(autonum=autonum)

    def BreakWordLatin(self, break_latin_word):
        return self.hwp.BreakWordLatin(BreakLatinWord=break_latin_word)

    def BrushType(self, brush_type):
        return self.hwp.BrushType(BrushType=brush_type)

    def Canonical(self, canonical):
        return self.hwp.Canonical(Canonical=canonical)

    def CellApply(self, cell_apply):
        return self.hwp.CellApply(CellApply=cell_apply)

    def CharShadowType(self, shadow_type):
        return self.hwp.CharShadowType(ShadowType=shadow_type)

    def ColDefType(self, col_def_type):
        return self.hwp.ColDefType(ColDefType=col_def_type)

    def ColLayoutType(self, col_layout_type):
        return self.hwp.ColLayoutType(ColLayoutType=col_layout_type)

    def ConvertPUAHangulToUnicode(self, reverse):
        return self.hwp.ConvertPUAHangulToUnicode(Reverse=reverse)

    def CrookedSlash(self, crooked_slash):
        return self.hwp.CrookedSlash(CrookedSlash=crooked_slash)

    def DSMark(self, diac_sym_mark):
        return self.hwp.DSMark(DiacSymMark=diac_sym_mark)

    def DbfCodeType(self, dbf_code):
        return self.hwp.DbfCodeType(DbfCode=dbf_code)

    def Delimiter(self, delimiter):
        return self.hwp.Delimiter(Delimiter=delimiter)

    def DrawAspect(self, draw_aspect):
        return self.hwp.DrawAspect(DrawAspect=draw_aspect)

    def DrawFillImage(self, fillimage):
        return self.hwp.DrawFillImage(fillimage=fillimage)

    def DrawShadowType(self, shadow_type):
        return self.hwp.DrawShadowType(ShadowType=shadow_type)

    def Encrypt(self, encrypt):
        return self.hwp.Encrypt(Encrypt=encrypt)

    def EndSize(self, end_size):
        return self.hwp.EndSize(EndSize=end_size)

    def EndStyle(self, end_style):
        return self.hwp.EndStyle(EndStyle=end_style)

    def FontType(self, font_type):
        return self.hwp.FontType(FontType=font_type)

    def GetTranslateLangList(self, cur_lang):
        return self.hwp.GetTranslateLangList(curLang=cur_lang)

    def GetUserInfo(self, user_info_id):
        return self.hwp.GetUserInfo(userInfoId=user_info_id)

    def Gradation(self, gradation):
        return self.hwp.Gradation(Gradation=gradation)

    def GridMethod(self, grid_method):
        return self.hwp.GridMethod(GridMethod=grid_method)

    def GridViewLine(self, grid_view_line):
        return self.hwp.GridViewLine(GridViewLine=grid_view_line)

    def GutterMethod(self, gutter_type):
        return self.hwp.GutterMethod(GutterType=gutter_type)

    def HAlign(self, h_align):
        return self.hwp.HAlign(HAlign=h_align)

    def Handler(self, handler):
        return self.hwp.Handler(Handler=handler)

    def Hash(self, hash):
        return self.hwp.Hash(Hash=hash)

    def HatchStyle(self, hatch_style):
        return self.hwp.HatchStyle(HatchStyle=hatch_style)

    def HeadType(self, heading_type):
        return self.hwp.HeadType(HeadingType=heading_type)

    def HeightRel(self, height_rel):
        return self.hwp.HeightRel(HeightRel=height_rel)

    def Hiding(self, hiding):
        return self.hwp.Hiding(Hiding=hiding)

    def HorzRel(self, horz_rel):
        return self.hwp.HorzRel(HorzRel=horz_rel)

    def HwpLineType(self, line_type: Literal["None", "Solid", "Dash", "Dot", "DashDot", "DashDotDot", "LongDash", "Circle", "DoubleSlim", "SlimThick", "ThickSlim", "SlimThickSlim"] = "Solid"):
        """
        한/글에서 표나 개체의 선 타입을 결정하는 헬퍼메서드. 단순히 문자열을 정수로 변환한다.

        Args:
            line_type:
                문자열 파라미터. 종류는 아래와 같다.

                    - "None": 없음(0)
                    - "Solid": 실선(1)
                    - "Dash": 파선(2)
                    - "Dot": 점선(3)
                    - "DashDot": 일점쇄선(4)
                    - "DashDotDot": 이점쇄선(5)
                    - "LongDash": 긴 파선(6)
                    - "Circle": 원형 점선(7)
                    - "DoubleSlim": 이중 실선(8)
                    - "SlimThick": 얇고 굵은 이중선(9)
                    - "ThickSlim": 굵고 얇은 이중선(10)
                    - "SlimThickSlim": 얇고 굵고 얇은 삼중선(11)
        """
        return self.hwp.HwpLineType(LineType=line_type)

    def HwpLineWidth(self, line_width: Literal["0.1mm", "0.12mm", "0.15mm", "0.2mm", "0.25mm", "0.3mm", "0.4mm", "0.5mm", "0.6mm", "0.7mm", "1.0mm", "1.5mm", "2.0mm", "3.0mm", "4.0mm", "5.0mm"] = "0.1mm") -> int:
        """
        선 너비를 정해주는 헬퍼 메서드.

        목록은 아래와 같다.

        Args:
            line_width:

                - "0.1mm": 0
                - "0.12mm": 1
                - "0.15mm": 2
                - "0.2mm": 3
                - "0.25mm": 4
                - "0.3mm": 5
                - "0.4mm": 6
                - "0.5mm": 7
                - "0.6mm": 8
                - "0.7mm": 9
                - "1.0mm": 10
                - "1.5mm": 11
                - "2.0mm": 12
                - "3.0mm": 13
                - "4.0mm": 14
                - "5.0mm": 15

        Returns:
            hwp가 인식하는 선굵기 정수(0~15)
            """
        return self.hwp.HwpLineWidth(LineWidth=line_width)

    def HwpOutlineStyle(self, hwp_outline_style):
        return self.hwp.HwpOutlineStyle(HwpOutlineStyle=hwp_outline_style)

    def HwpOutlineType(self, hwp_outline_type):
        return self.hwp.HwpOutlineType(HwpOutlineType=hwp_outline_type)

    def HwpUnderlineShape(self, hwp_underline_shape):
        return self.hwp.HwpUnderlineShape(HwpUnderlineShape=hwp_underline_shape)

    def HwpUnderlineType(self, hwp_underline_type):
        return self.hwp.HwpUnderlineType(HwpUnderlineType=hwp_underline_type)

    def HwpZoomType(self, zoom_type):
        return self.hwp.HwpZoomType(ZoomType=zoom_type)

    def ImageFormat(self, image_format):
        return self.hwp.ImageFormat(ImageFormat=image_format)

    def LineSpacingMethod(self, line_spacing):
        return self.hwp.LineSpacingMethod(LineSpacing=line_spacing)

    def LineWrapType(self, line_wrap):
        return self.hwp.LineWrapType(LineWrap=line_wrap)

    def LunarToSolar(self, l_year, l_month, l_day, l_leap, s_year, s_month, s_day):
        return self.hwp.LunarToSolar(lYear=l_year, lMonth=l_month, lDay=l_day, lLeap=l_leap, sYear=s_year,
                                     sMonth=s_month, sDay=s_day)

    def LunarToSolarBySet(self, l_year, l_month, l_day, l_leap):
        return self.hwp.LunarToSolarBySet(lYear=l_year, lMonth=l_month, lLeap=l_leap)

    def MacroState(self, macro_state):
        return self.hwp.MacroState(MacroState=macro_state)

    def MailType(self, mail_type):
        return self.hwp.MailType(MailType=mail_type)

    def mili_to_hwp_unit(self, mili: float) -> int:
        return self.hwp.MiliToHwpUnit(mili=mili)

    def MiliToHwpUnit(self, mili: float) -> int:
        return self.hwp.MiliToHwpUnit(mili=mili)

    def NumberFormat(self, num_format):
        return self.hwp.NumberFormat(NumFormat=num_format)

    def Numbering(self, numbering):
        return self.hwp.Numbering(Numbering=numbering)

    def PageNumPosition(self, pagenumpos: Literal[
        "TopLeft", "TopCenter", "TopRight", "BottomLeft", "BottomCenter", "BottomRight", "InsideTop", "OutsideTop", "InsideBottom", "OutsideBottom", "None"] = "BottomCenter"):
        return self.hwp.PageNumPosition(pagenumpos=pagenumpos)

    def PageType(self, page_type):
        return self.hwp.PageType(PageType=page_type)

    def ParaHeadAlign(self, para_head_align):
        return self.hwp.ParaHeadAlign(ParaHeadAlign=para_head_align)

    def PicEffect(self, pic_effect):
        return self.hwp.PicEffect(PicEffect=pic_effect)

    def PlacementType(self, restart):
        return self.hwp.PlacementType(Restart=restart)

    def PresentEffect(self, prsnteffect):
        return self.hwp.PresentEffect(prsnteffect=prsnteffect)

    def PrintDevice(self, print_device):
        return self.hwp.PrintDevice(PrintDevice=print_device)

    def PrintPaper(self, print_paper):
        return self.hwp.PrintPaper(PrintPaper=print_paper)

    def PrintRange(self, print_range):
        return self.hwp.PrintRange(PrintRange=print_range)

    def PrintType(self, print_method):
        return self.hwp.PrintType(PrintMethod=print_method)

    def SetCurMetatagName(self, tag):
        return self.hwp.SetCurMetatagName(tag=tag)

    def SetDRMAuthority(self, authority):
        return self.hwp.SetDRMAuthority(authority=authority)

    def SetUserInfo(self, user_info_id, value):
        return self.hwp.SetUserInfo(userInfoId=user_info_id, Value=value)

    def SideType(self, side_type):
        return self.hwp.SideType(SideType=side_type)

    def Signature(self, signature):
        return self.hwp.Signature(Signature=signature)

    def Slash(self, slash):
        return self.hwp.Slash(Slash=slash)

    def SolarToLunar(self, s_year, s_month, s_day, l_year, l_month, l_day, l_leap):
        return self.hwp.SolarToLunar(sYear=s_year, sMonth=s_month, sDay=s_day, lYear=l_year, lMonth=l_month, lDay=l_day,
                                     lLeap=l_leap)

    def SolarToLunarBySet(self, s_year, s_month, s_day):
        return self.hwp.SolarToLunarBySet(sYear=s_year, sMonth=s_month, sDay=s_day)

    def SortDelimiter(self, sort_delimiter):
        return self.hwp.SortDelimiter(SortDelimiter=sort_delimiter)

    def StrikeOut(self, strike_out_type):
        return self.hwp.StrikeOut(StrikeOutType=strike_out_type)

    def StyleType(self, style_type):
        return self.hwp.StyleType(StyleType=style_type)

    def SubtPos(self, subt_pos):
        return self.hwp.SubtPos(SubtPos=subt_pos)

    def TableBreak(self, page_break):
        return self.hwp.TableBreak(PageBreak=page_break)

    def TableFormat(self, table_format):
        return self.hwp.TableFormat(TableFormat=table_format)

    def TableSwapType(self, tableswap):
        return self.hwp.TableSwapType(tableswap=tableswap)

    def TableTarget(self, table_target):
        return self.hwp.TableTarget(TableTarget=table_target)

    def TextAlign(self, text_align):
        return self.hwp.TextAlign(TextAlign=text_align)

    def TextArtAlign(self, text_art_align):
        return self.hwp.TextArtAlign(TextArtAlign=text_art_align)

    def TextDir(self, text_direction):
        return self.hwp.TextDir(TextDirection=text_direction)

    def TextFlowType(self, text_flow):
        return self.hwp.TextFlowType(TextFlow=text_flow)

    def TextWrapType(self, text_wrap):
        return self.hwp.TextWrapType(TextWrap=text_wrap)

    def VAlign(self, v_align):
        return self.hwp.VAlign(VAlign=v_align)

    def VertRel(self, vert_rel):
        return self.hwp.VertRel(VertRel=vert_rel)

    def ViewFlag(self, view_flag):
        return self.hwp.ViewFlag(ViewFlag=view_flag)

    def WatermarkBrush(self, watermark_brush):
        return self.hwp.WatermarkBrush(WatermarkBrush=watermark_brush)

    def WidthRel(self, width_rel):
        return self.hwp.WidthRel(WidthRel=width_rel)