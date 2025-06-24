from __future__ import annotations
from .param_helpers import ParamHelpers
from .run_methods import RunMethods
from .fonts import fonts
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
from typing import Literal, Union, Any, Optional, Tuple, List, Dict
from urllib import request, parse
from winreg import QueryValueEx

import numpy as np
import pandas as pd
import pyperclip as cb
from PIL import Image

if sys.platform == "win32":
    import pythoncom
    import win32api
    import win32con
    import win32gui

    # CircularImport 오류 출력안함
    devnull = open(os.devnull, "w")
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

if getattr(sys, "frozen", False):
    pyinstaller_path = sys._MEIPASS
else:
    pyinstaller_path = os.path.dirname(os.path.abspath(__file__))

# temp 폴더 삭제
try:
    shutil.rmtree(os.path.join(os.environ["USERPROFILE"], "AppData/Local/Temp/gen_py"))
except FileNotFoundError as e:
    pass

# Type Library 파일 재생성
win32.gencache.EnsureModule("{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}", 0, 1, 0)


__all__ = ["Hwp", "com_initialized"]


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


def log_error(method):
    @wraps(method)
    def wrapper(*args, **kwargs):
        try:
            return method(*args, **kwargs)
        except Exception as e:
            print(f"오류 발생: {e} \n 다음 함수를 확인하세요 : [{method.__name__}]")
            raise  # 예외를 다시 전파 (없애고 싶으면 여기만 수정하면 됨)

    return wrapper


def addr_to_tuple(cell_address: str) -> Tuple[int, int]:
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
        col = col * 26 + (ord(ch) - ord("A") + 1)

    # 숫자 부분 -> 행 번호(row)로 변환
    row = int(row_str)

    return row, col


def tuple_to_addr(row: int, col: int) -> str:
    """
    (행번호, 칼럼번호)를 인자로 받아 엑셀 셀 주소 문자열(예: `"AAA3"`)을 반환합니다.

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
        letters.append(chr(remainder + ord("A")))
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
    text = buffer[: length * 2].tobytes().decode("utf-16")[:-1]
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


def crop_data_from_selection(data, selection) -> List[str]:
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
        result.append(data[row][min_col : max_col + 1])

    return result


def check_registry_key(key_name: str = "FilePathCheckerModule") -> bool:
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


def rename_duplicates_in_list(file_list: List[str]) -> List[str]:
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


def excel_address_to_tuple_zero_based(address: str) -> Tuple[Union[int, Any], Union[int, Any]]:
    """
    엑셀 셀 주소를 튜플로 변환하는 헬퍼함수

    """
    column = 0
    row = 0
    for char in address:
        if char.isalpha():
            column = column * 26 + (ord(char.upper()) - ord("A"))
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
        if hasattr(self, "_com_obj"):
            return f"<CtrlCode: CtrlID={self.CtrlID!r}, CtrlCH={self.CtrlCh!r}, UserDesc={self.UserDesc!r}>"
        return None

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
    def Next(self) -> Optional[Ctrl]:
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
        try:
            next_ctrl = self._com_obj.Next
        except AttributeError:
            return None
        return Ctrl(next_ctrl) if next_ctrl is not None else None

    @property
    def Prev(self) -> Optional[Ctrl]:
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
        try:
            prev_ctrl = self._com_obj.Prev
        except AttributeError:
            return None
        return Ctrl(prev_ctrl) if prev_ctrl is not None else None

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
    한글과컴퓨터 문서(HWP) COM 객체의 컬렉션을 표현하는 클래스입니다.

    이 클래스는 COM 객체로 표현된 문서 컬렉션과 상호작용하기 위한 속성과 메서드를 제공합니다.
    문서 컬렉션에 대한 반복, 인덱싱, 길이 조회 등의 컬렉션 유사 동작을 지원하며,
    문서를 추가하고, 닫고, 특정 문서 오브젝트를 검색하는 메서드를 포함합니다.

    Attributes:
        Active_XHwpDocument (XHwpDocument): 컬렉션의 활성 문서
        Application: COM 객체와 연결된 애플리케이션을 반환
        CLSID: COM 객체의 CLSID
        Count (int): 컬렉션 내 문서 개수

    Methods:
        Add(isTab: bool = False) -> XHwpDocument:
            컬렉션에 새 문서를 추가합니다.
        Close(isDirty: bool = False) -> None:
            활성 문서 창을 닫습니다.
        FindItem(lDocID: int) -> XHwpDocument:
            주어진 문서 ID에 해당하는 문서 객체를 찾아 반환합니다.
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

    def Add(self, isTab: bool = False) -> "XHwpDocument":
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

    def FindItem(self, lDocID: int) -> "XHwpDocument":
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
    """
    한글과컴퓨터 문서(HWP)의 COM 객체를 래핑한 클래스입니다.

    이 클래스는 HwpDocument COM 객체와 상호작용하기 위한 속성과 메서드를 제공합니다.
    문서 속성, 데이터 조작 기능, 상태 제어 등 다양한 기능을 제공하며,
    한글과컴퓨터의 문서 관리 시스템을 파이썬스럽게 다룰 수 있도록 설계되었습니다.

    속성:
        Application: 문서와 연결된 어플리케이션 객체
        CLSID: 문서의 클래스 ID
        DocumentID: 문서의 고유 식별자
        EditMode: 현재 문서의 편집 모드
        Format: 문서의 형식
        FullName: 문서의 전체 경로를 문자열로 반환. 저장되지 않은 문서는 빈 문자열 반환
        Modified: 문서의 수정 여부를 나타냄
        Path: 문서가 저장된 폴더 경로
        XHwpCharacterShape: 문서의 글자모양 설정에 접근
        XHwpDocumentInfo: 문서의 상세 정보에 접근
        XHwpFind: 텍스트 찾기 기능에 접근
        XHwpFormCheckButtons: 체크박스 양식 요소에 접근
        XHwpFormComboBoxs: 콤보박스 양식 요소에 접근
        XHwpFormEdits: 편집 양식 요소에 접근
        XHwpFormPushButtons: 버튼 양식 요소에 접근
        XHwpFormRadioButtons: 라디오버튼 양식 요소에 접근
        XHwpParagraphShape: 문단 모양 설정에 접근
        XHwpPrint: 문서 인쇄 제어에 접근
        XHwpRange: 문서 내 범위 설정에 접근
        XHwpSelection: 현재 선택된 텍스트나 개체에 접근
        XHwpSendMail: 메일 보내기 기능에 접근
        XHwpSummaryInfo: 문서 요약 정보에 접근
    """

    def __repr__(self):
        return f'<Doc: DocumentID={self.DocumentID}, FullName="{self.FullName or None}", Modified="{True if self.Modified else False}">'

    def __init__(self, com_obj):
        self._com_obj = com_obj

    @property
    def Application(self):
        return self._com_obj.Application

    @property
    def CLSID(self):
        return self._com_obj.CLSID

    def Clear(self, option: bool = False) -> None:
        return self._com_obj.Clear(option=option)

    def Close(self, isDirty: bool = False) -> None:
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

    def Open(self, filename: str, Format: str, arg: str):
        return self._com_obj.Open(filename=filename, Format=Format, arg=arg)

    @property
    def Path(self) -> str:
        return self._com_obj.Path

    def Redo(self, Count: int):
        return self._com_obj.Redo(Count=Count)

    def Save(self, save_if_dirty: bool):
        return self._com_obj.Save(save_if_dirty=save_if_dirty)

    def SaveAs(self, Path: str, Format: str, arg: str):
        return self._com_obj.SaveAs(Path=Path, Format=Format, arg=arg)

    def SendBrowser(self):
        return self._com_obj.SendBrowser()

    def SetActive_XHwpDocument(self):
        return self._com_obj.SetActive_XHwpDocument()

    def Undo(self, Count: int):
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
class Hwp(ParamHelpers, RunMethods):
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
        return f'<Hwp: DocumentID={self.XHwpDocuments.Active_XHwpDocument.DocumentID}, Title="{self.get_title()}", FullName="{self.XHwpDocuments.Active_XHwpDocument.FullName or None}">'

    def __init__(
        self,
        new: bool = False,
        visible: bool = True,
        register_module: bool = True,
        on_quit: bool = False,
    ):
        self.hwp = 0
        self.on_quit = on_quit
        self.htf_fonts = fonts
        context = pythoncom.CreateBindCtx(0)
        pythoncom.CoInitialize()  # 이걸 꼭 실행해야 하는가? 왜 Pycharm이나 주피터에서는 괜찮고, vscode에서는 CoInitialize 오류가 나는지?
        running_coms = pythoncom.GetRunningObjectTable()
        monikers = running_coms.EnumRunning()

        if not new:
            for moniker in monikers:
                name = moniker.GetDisplayName(context, moniker)
                if name.startswith("!HwpObject."):
                    obj = running_coms.GetObject(moniker)
                    self.hwp = win32.gencache.EnsureDispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch)
                    )
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
                print(
                    e, "RegisterModule 액션을 실행할 수 없음. 개발자에게 문의해주세요."
                )

    def __del__(self):
        if self.on_quit:
            try:
                self.quit(save=False)
            except:
                pass
        pythoncom.CoUninitialize()

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
    def Version(self) -> List[int]:
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
        return (
            self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.CurrentPage + 1
        )

    @property
    def current_printpage(self) -> int:
        """
        페이지인덱스가 아닌, 종이에 표시되는 쪽번호를 리턴.

        1페이지에 있다면 1을 리턴한다.
        새쪽번호가 적용되어 있다면
        수정된 쪽번호를 리턴한다.

        Returns:
        """
        return (
            self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.CurrentPrintPage
        )

    @property
    def current_font(self):
        charshape = self.get_charshape_as_dict()  # hwp.CharShape
        if charshape["FontTypeHangul"] == 1:
            return charshape["FaceNameHangul"]
        elif charshape["FontTypeHangul"] == 2:
            sub_dict = {
                key: value for key, value in charshape.items() if key.startswith("F")
            }
            for key, value in self.htf_fonts.items():
                if value == sub_dict:
                    return key

    # 커스텀 메서드
    def get_ctrl_pos(
        self, ctrl: Any = None, option: Literal[0, 1] = 0, as_tuple: bool = True
    ) -> Tuple[int, int, int]:
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

    def get_linespacing(
        self, method: Literal["Fixed", "Percent", "BetweenLines", "AtLeast"] = "Percent"
    ) -> Union[int, float]:
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
            return self.HwpUnitToPoint(
                pset.LineSpacing / 2
            )  # 이상하게 1/2 곱해야 맞다.

    def set_linespacing(
        self,
        value: Union[int, float] = 160,
        method: Literal["Fixed", "Percent", "BetweenLines", "AtLeast"] = "Percent",
    ) -> bool:
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
        before_pos = self.get_pos()
        self.MoveRight()
        after_pos = self.get_pos()
        self.set_pos(*before_pos)
        if before_pos[2] == after_pos[2] == 0 and before_pos[0] == after_pos[0]:
            return True
        else:
            return False

    def goto_addr(
        self, addr: Union[str, int] = "A1", col: int = 0, select_cell: bool = False
    ) -> bool:
        """
        셀 주소를 문자열로 입력받아 해당 주소로 이동하는 메서드.
        셀 주소는 "C3"처럼 문자열로 입력하거나, 행번호, 열번호를 입력할 수 있음. 시작값은 1.

        Args:
            addr: 셀 주소 문자열("A1") 또는 행번호(1부터).
            col: 셀 주소를 정수로 입력하는 경우 열번호(1부터)
            select_cell: 이동 후 셀블록 선택 여부

        Returns:
           이동 성공 여부(성공시 True/실패시 False)
        """
        if not self.is_cell():
            return False  # 표 안에 있지 않으면 False 리턴(종료)
        cur_pos = self.get_pos()

        if (
            type(addr) == int and col
        ):  # "A1" 대신 (1, 1) 처럼 tuple[int, int] 방식일 경우
            addr = tuple_to_addr(addr, col)  # 문자열 "A1" 방식으로 우선 변환

        refresh = False

        # 우선 맨 끝 셀로 이동
        self.HAction.Run("TableColEnd")
        self.HAction.Run("TableColPageDown")
        t_end = self.get_pos()[0]  # 마지막 셀의 구역번호(List) 저장
        self.HAction.Run("TableColBegin")
        self.HAction.Run("TableColPageUp")
        t_init = self.get_pos()[0]  # 시작 셀의 구역번호(List) 저장

        try:
            if self.addr_info[0] == t_end:
                pass
            else:
                refresh = True
                self.addr_info = [
                    t_end,
                    ["A1"],
                ]  # 로컬변수가 아닌 인스턴스변수로 저장(재실행 때 활용하기 위함)
        except AttributeError:
            refresh = True
            self.addr_info = [t_end, ["A1"]]

        if refresh:
            i = 1
            while self.set_pos(t_init + i, 0, 0):
                cur_addr = self.KeyIndicator()[-1][1:].split(")")[0]
                if cur_addr == "A1":
                    temp_pos = self.get_pos()
                    self.CloseEx()
                    if self.is_cell():
                        self.set_pos(*temp_pos)
                        self.TableCellBlockExtendAbs()
                        self.TableCellBlockExtend()
                        subt_len = len(self.get_selected_range())
                        i += subt_len
                        for _ in range(subt_len):
                            self.addr_info[1].append("")
                        self.CloseEx()
                        self.TableRightCell()
                        continue
                    else:
                        self.set_pos(*temp_pos)
                    break
                if not self.is_cell():
                    self.addr_info[1].append("")
                    i += 1
                    continue
                if self.get_pos()[0] == 0:
                    break
                self.addr_info[1].append(cur_addr)
                i += 1
        try:
            self.set_pos(t_init + self.addr_info[1].index(addr.upper()), 0, 0)
            if select_cell:
                self.HAction.Run("TableCellBlock")
            return True
        except ValueError:
            self.set_pos(*cur_pos)
            return False

    def get_field_info(self) -> List[dict]:
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
                command = re.split(
                    r"(Clickhere:set:\d+:Direction:wstring:\d+:)|( HelpState:wstring:\d+:)",
                    field.attrib.get("Command")[:-2],
                )
                results.append(
                    {"name": name_value, "direction": command[3], "memo": command[-1]}
                )
            return results
        except ET.ParseError as e:
            print("XML 파싱 오류:", e)
            return False
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다.")
            return False

    def get_image_info(self, ctrl: Any = None) -> Dict[str, List[int]]:
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
        tree = ET.parse("temp.xml")
        root = tree.getroot()

        for shapeobject in root.findall(".//SHAPEOBJECT"):
            shapecmt = shapeobject.find("SHAPECOMMENT")
            if shapecmt is not None and shapecmt.text:
                info = shapecmt.text.split("\n")[1:]
                try:
                    return {
                        "name": info[0].split(": ")[1],
                        "size": [int(i) for i in info[1][14:-5].split("pixel, 세로 ")],
                    }
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
            if style in [style_dict[i]["Name"] for i in style_dict]:
                style_idx = [i for i in style_dict if style_dict[i]["Name"] == style][
                    0
                ] + 1
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

    def shape_copy_paste(
        self,
        Type: Literal["font", "para", "both"] = "both",
        cell_attr: bool = False,
        cell_border: bool = False,
        cell_fill: bool = False,
        cell_only: int = 0,
    ) -> bool:
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
        win32gui.EnumChildWindows(
            hwnd2, lambda hwnd, param: param.append(hwnd), child_hwnds
        )
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
            raise AssertionError(
                "mathml 파일을 찾을 수 없습니다. 경로를 다시 확인해주세요."
            )

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
        win32gui.EnumChildWindows(
            hwnd2, lambda hwnd, param: param.append(hwnd), child_hwnds
        )
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

    def maximize_window(self) -> None:
        """현재 창 최대화"""
        win32gui.ShowWindow(self.XHwpWindows.Active_XHwpWindow.WindowHandle, 3)

    def minimize_window(self) -> None:
        """현재 창 최소화"""
        win32gui.ShowWindow(self.XHwpWindows.Active_XHwpWindow.WindowHandle, 6)

    def delete_style_by_name(self, src: Union[int, str, List[Union[int, str]]], dst: Union[int, str]) -> bool:
        """
        특정 스타일을 이름 (또는 인덱스번호)로 삭제하고
        대체할 스타일 또한 이름 (또는 인덱스번호)로 지정해주는 메서드.
        """
        style_dict: dict = self.get_style_dict(as_="dict")
        if type(src) != list:
            src = [src]

        for idx, s in enumerate(src):
            pset = self.HParameterSet.HStyleDelete
            self.HAction.GetDefault("StyleDelete", pset.HSet)
            if type(s) == int:
                pset.Target = s
            elif s in [style_dict[i]["Name"] for i in style_dict]:
                pset.Target = [i for i in style_dict if style_dict[i]["Name"] == s][0]
            else:
                raise IndexError("해당 스타일이름을 찾을 수 없습니다.")
            if type(dst) == int:
                pset.Alternation = dst
            elif dst in [style_dict[i]["Name"] for i in style_dict]:
                pset.Alternation = [i for i in style_dict if style_dict[i]["Name"] == dst][0]
            else:
                raise IndexError("해당 스타일이름을 찾을 수 없습니다.")
            self.HAction.Execute("StyleDelete", pset.HSet)
        return True

    def get_style_dict(self, as_: Literal["list", "dict"] = "list") -> Union[list, dict]:
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
                    "CharShape": style.get("CharShape"),
                    "EngName": style.get("EngName"),
                    "Id": int(style.get("Id")),
                    "LangId": style.get("LangId"),
                    "LockForm": style.get("LockForm"),
                    "Name": style.get("Name"),
                    "NextStyle": style.get("NextStyle"),
                    "ParaShape": style.get("ParaShape"),
                    "Type": style.get("Type"),
                }
                for style in root.findall(".//STYLE")
            ]
        elif as_ == "dict":
            styles = {
                int(style.get("Id")): {
                    "CharShape": style.get("CharShape"),
                    "EngName": style.get("EngName"),
                    "Id": int(style.get("Id")),
                    "LangId": style.get("LangId"),
                    "LockForm": style.get("LockForm"),
                    "Name": style.get("Name"),
                    "NextStyle": style.get("NextStyle"),
                    "ParaShape": style.get("ParaShape"),
                    "Type": style.get("Type"),
                }
                for style in root.findall(".//STYLE")
            }
        else:
            raise TypeError(
                "as_ 파라미터는 'list'또는 'dict' 중 하나로 설정해주세요. 기본값은 'list'입니다."
            )
        os.remove("temp.xml")
        return styles

    def get_used_style_dict(self, as_: Literal["list", "dict"] = "list") -> Union[list, dict]:
        """
        현재 문서에서 사용된 스타일 목록만 list[dict] 또는 dict[dict] 데이터로 리턴하는 메서드.
        """

        cur_pos = self.get_pos()
        if not self.MoveSelRight():
            self.MoveSelLeft()
        self.SelectAll()
        self.save_block_as("temp.xml", format="HWPML2X")
        self.Cancel()
        self.set_pos(*cur_pos)

        tree = ET.parse("temp.xml")
        root = tree.getroot()
        if as_ == "list":
            styles = [
                {
                    "index": int(style.get("Id")),
                    "type": style.get("Type"),
                    "name": style.get("Name"),
                    "engName": style.get("EngName"),
                }
                for style in root.findall(".//STYLE")
            ]
        elif as_ == "dict":
            styles = {
                int(style.get("Id")): {
                    "type": style.get("Type"),
                    "name": style.get("Name"),
                    "engName": style.get("EngName"),
                }
                for style in root.findall(".//STYLE")
            }
        else:
            raise TypeError(
                "as_ 파라미터는 'list'또는 'dict' 중 하나로 설정해주세요. 기본값은 'list'입니다."
            )
        used_style_index = {int(p.get('Style')) for p in root.findall('.//P') if p.get('Style') is not None}
        os.remove("temp.xml")
        return [i for i in styles if i["index"] in used_style_index] \
            if as_ == "list" else {i: styles[i] for i in styles if i in used_style_index}

    def remove_unused_styles(self, alt=0):
        """
        문서 내에 정의만 되어 있고 실제 사용되지 않은 모든 스타일을 일괄제거하는 메서드.
        사용에 주의할 것.
        """
        self.MoveDocBegin()
        self.SelectAll()
        used_styles = self.get_used_style_dict("dict").keys()
        self.Cancel()
        all_styles = self.get_style_dict("dict").keys()
        return self.delete_style_by_name(list(all_styles - used_styles)[::-1], alt)

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

    def set_style(self, style: Union[int, str]) -> bool:
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
                if value.get("Name") == style:
                    style = key
                    break
                else:
                    continue
            if style != key:
                raise KeyError("해당하는 스타일이 없습니다.")
        self.HAction.GetDefault("Style", pset.HSet)
        pset.Apply = style
        return self.HAction.Execute("Style", pset.HSet)

    def get_selected_range(self) -> List[str]:
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

    def resize_image(
        self,
        width: int = None,
        height: int = None,
        unit: Literal["mm", "hwpunit"] = "mm",
    ):
        """
        이미지 또는 그리기 개체의 크기를 조절하는 메서드.

        해당개체 선택 후 실행해야 함.
        """
        self.FindCtrl()
        prop = self.CurSelectedCtrl.Properties
        if width:
            prop.SetItem(
                "Width", width if unit == "hwpunit" else self.MiliToHwpUnit(width)
            )
        if height:
            prop.SetItem(
                "Height", height if unit == "hwpunit" else self.MiliToHwpUnit(height)
            )
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
                temp_file = sorted(
                    [i for i in path_list if i.startswith(os.path.splitext(path))],
                    key=os.path.getmtime,
                )[-1]
                Image.open(temp_file).save(path)
                os.remove(temp_file)
            print(f"image saved to {path}")

    def new_number_modify(
        self,
        new_number: int,
        num_type: Literal[
            "Page", "Figure", "Footnote", "Table", "Endnote", "Equation"
        ] = "Page",
    ) -> bool:
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

    def NewNumberModify(
        self,
        new_number: int,
        num_type: Literal[
            "Page", "Figure", "Footnote", "Table", "Endnote", "Equation"
        ] = "Page",
    ) -> bool:
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

    def new_number(
        self,
        new_number: int,
        num_type: Literal[
            "Page", "Figure", "Footnote", "Table", "Endnote", "Equation"
        ] = "Page",
    ) -> bool:
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

    def NewNumber(
        self,
        new_number: int,
        num_type: Literal[
            "Page", "Figure", "Footnote", "Table", "Endnote", "Equation"
        ] = "Page",
    ) -> bool:
        pset = self.HParameterSet.HAutoNum
        self.HAction.GetDefault("NewNumber", pset.HSet)
        pset.NumType = self.AutoNumType(num_type)
        pset.NewNumber = new_number
        return self.HAction.Execute("NewNumber", pset.HSet)

    def page_num_pos(
        self,
        global_start: int = 1,
        position: Literal[
            "TopLeft",
            "TopCenter",
            "TopRight",
            "BottomLeft",
            "BottomCenter",
            "BottomRight",
            "InsideTop",
            "OutsideTop",
            "InsideBottom",
            "OutsideBottom",
            "None",
        ] = "BottomCenter",
        number_format: Literal[
            "Digit",
            "CircledDigit",
            "RomanCapital",
            "RomanSmall",
            "LatinCapital",
            "HangulSyllable",
            "Ideograph",
            "DecagonCircle",
            "DecagonCircleHanja",
        ] = "Digit",
        side_char: bool = True,
    ) -> bool:
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

    def PageNumPos(
        self,
        global_start: int = 1,
        position: Literal[
            "TopLeft",
            "TopCenter",
            "TopRight",
            "BottomLeft",
            "BottomCenter",
            "BottomRight",
            "InsideTop",
            "OutsideTop",
            "InsideBottom",
            "OutsideBottom",
            "None",
        ] = "BottomCenter",
        number_format: Literal[
            "Digit",
            "CircledDigit",
            "RomanCapital",
            "RomanSmall",
            "LatinCapital",
            "HangulSyllable",
            "Ideograph",
            "DecagonCircle",
            "DecagonCircleHanja",
        ] = "Digit",
        side_char: bool = True,
    ) -> bool:
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

    def set_cell_margin(
        self,
        left: float = 1.8,
        right: float = 1.8,
        top: float = 0.5,
        bottom: float = 0.5,
        as_: Literal["mm", "hwpunit"] = "mm",
    ) -> bool:
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

    def get_cell_margin(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> Union[None, Dict[str, int], Dict[str, float], bool]:
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
        else:
            return False

    def set_table_inside_margin(
        self,
        left: float = 1.8,
        right: float = 1.8,
        top: float = 0.5,
        bottom: float = 0.5,
        as_: Literal["mm", "hwpunit"] = "mm",
    ) -> bool:
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
            left, right, top, bottom = [
                self.hwp.MiliToHwpUnit(i) for i in [left, right, top, bottom]
            ]
        pset = self.hwp.HParameterSet.HShapeObject
        self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        pset.CellMarginLeft = left
        pset.CellMarginRight = right
        pset.CellMarginTop = top
        pset.CellMarginBottom = bottom
        return self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)

    def get_table_inside_margin(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> Union[None, Dict[str, int], bool, Dict[str, float]]:
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
        else:
            return False

    def get_table_outside_margin(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> Union[None, Dict[str, int], bool, Dict[str, float]]:
        """
        표의 바깥 여백을 딕셔너리로 한 번에 리턴하는 메서드

        Args:
            as_: 리턴하는 여백값의 단위

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

    def get_table_outside_margin_left(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> bool:
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

    def get_table_outside_margin_right(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> bool:
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

    def get_table_outside_margin_top(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> bool:
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

    def get_table_outside_margin_bottom(
        self, as_: Literal["mm", "hwpunit"] = "mm"
    ) -> Union[int, float, bool]:
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

    def set_table_outside_margin(
        self,
        left: float = -1.0,
        right: float = -1.0,
        top: float = -1.0,
        bottom: float = -1.0,
        as_: Literal["mm", "hwpunit"] = "mm",
    ) -> bool:
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

    def set_table_outside_margin_bottom(
        self, val, as_: Literal["mm", "hwpunit"] = "mm"
    ):
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
            raise KeyError(
                "mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다."
            )

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
        table = root.find(".//TABLE")
        row_count = int(table.get("RowCount"))
        self.set_pos(*cur_pos)
        return row_count

    def get_row_height(
        self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"
    ) -> Union[float, int]:
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
            raise KeyError(
                "mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다."
            )

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
        table = root.find(".//TABLE")
        col_count = int(table.get("ColCount"))
        self.set_pos(*cur_pos)
        return col_count

    def get_col_width(
        self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"
    ) -> Union[int, float]:
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
            raise KeyError(
                "mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다."
            )

    def set_col_width(
        self, width: Union[int, float, list, tuple], as_: Literal["mm", "ratio"] = "ratio"
    ) -> bool:
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
                raise TypeError(
                    'width에 int나 float 입력시 as_ 파라미터는 "mm"로 설정해주세요.'
                )
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
                self.Cancel()
                self.TableRightCell()
            return self.set_pos(*cur_pos)

    def adjust_cellwidth(
        self, width: Union[int, float, list, tuple], as_: Literal["mm", "ratio"] = "ratio"
    ) -> bool:
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
                raise TypeError(
                    'width에 int나 float 입력시 as_ 파라미터는 "mm"로 설정해주세요.'
                )
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

    def get_table_width(
        self, as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"
    ) -> float:
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
            raise KeyError(
                "mm, hwpunit, hu, point, pt, inch 중 하나를 입력하셔야 합니다."
            )

    def set_table_width(
        self, width: int = 0, as_: Literal["mm", "hwpunit", "hu"] = "mm"
    ) -> bool:
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
                    sec_def.PageDef.PaperWidth
                    - sec_def.PageDef.LeftMargin
                    - sec_def.PageDef.RightMargin
                    - sec_def.PageDef.GutterLen
                    - self.get_table_outside_margin_left(as_="hwpunit")
                    - self.get_table_outside_margin_right(as_="hwpunit")
                )
            elif sec_def.PageDef.Landscape == 1:
                width = (
                    sec_def.PageDef.PaperHeight
                    - sec_def.PageDef.LeftMargin
                    - sec_def.PageDef.RightMargin
                    - sec_def.PageDef.GutterLen
                    - self.get_table_outside_margin_left(as_="hwpunit")
                    - self.get_table_outside_margin_right(as_="hwpunit")
                )
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
        table = root.find(".//TABLE")

        if table is not None:
            for cell in table.findall(".//CELL"):
                width = cell.get("Width")
                if width:
                    cell.set("Width", str(int(width) * ratio))
        t = ET.tostring(root, encoding="UTF-16").decode("utf-16")
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

    def save_pdf_as_image(self, path: str = "", img_format: str = "bmp") -> bool:
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

    def get_cell_addr(self, as_: Literal["str", "tuple"] = "str") -> Union[Tuple[int, ...], bool]:
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

    def save_all_pictures(self, save_path: str = "./binData") -> bool:
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

        with zipfile.ZipFile("./temp.zip", "r") as zf:
            zf.extractall(path="./temp")
        os.remove("./temp.zip")
        try:
            os.rename("./temp/binData", save_path)
        except FileExistsError:
            shutil.rmtree(save_path)
            os.rename("./temp/binData", save_path)
        with open("./temp/Contents/section0.xml", encoding="utf-8") as f:
            content = f.read()
        bin_list = re.findall(r"원본 그림의 이름: (.*?\..+?)\n", content)
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

    def SelectCtrl(self, ctrllist: Union[str, int], option: Literal[0, 1] = 1) -> bool:
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
            raise NotImplementedError(
                "아래아한글 버전이 2024 미만입니다. hwp.select_ctrl()을 대신 사용하셔야 합니다."
            )

    def select_ctrl(
        self, ctrl: Ctrl, anchor_type: Literal[0, 1, 2] = 0, option: int = 1
    ) -> bool:
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

    def set_visible(self, visible: bool) -> None:
        """
        현재 조작중인 한/글 인스턴스의 백그라운드 숨김여부를 변경할 수 있다.

        Args:
            visible: `visible=False`로 설정하면 현재 조작중인 한/글 인스턴스가 백그라운드로 숨겨진다.

        Returns:
            None
        """
        self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible

    def auto_spacing(
        self,
        init_spacing=0,
        init_ratio=100,
        max_spacing=40,
        min_spacing=40,
        verbose=True,
    ):
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
            string = f"{loc_info[3]}쪽 {loc_info[4]}단 {'' if self.get_pos()[0] == 0 else self.ParentCtrl.UserDesc}{start_line_no}줄({self.get_selected_text(keep_select=True)})"
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

    def set_font(
        self,
        Bold: Union[str, bool] = "",  # 진하게(True/False)
        DiacSymMark: Union[str, int] = "",  # 강조점(0~12)
        Emboss: Union[str, bool] = "",  # 양각(True/False)
        Engrave: Union[str, bool] = "",  # 음각(True/False)
        FaceName: str = "",  # 서체
        FontType: int = 1,  # 1(TTF), 2(HTF)
        Height: Union[str, float] = "",  # 글자크기(pt, 0.1 ~ 4096)
        Italic: Union[str, bool] = "",  # 이탤릭(True/False)
        Offset: Union[str, int] = "",  # 글자위치-상하오프셋(-100 ~ 100)
        OutLineType: Union[str, int] = "",  # 외곽선타입(0~6)
        Ratio: Union[str, int] = "",  # 장평(50~200)
        ShadeColor: (
            Union[str, int]
        ) = "",  # 음영색(RGB, 0x000000 ~ 0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
        ShadowColor: (
            Union[str, int]
        ) = "",  # 그림자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
        ShadowOffsetX: Union[str, int] = "",  # 그림자 X오프셋(-100 ~ 100)
        ShadowOffsetY: Union[str, int] = "",  # 그림자 Y오프셋(-100 ~ 100)
        ShadowType: Union[str, int] = "",  # 그림자 유형(0: 없음, 1: 비연속, 2:연속)
        Size: Union[str, int] = "",  # 글자크기 축소확대%(10~250)
        SmallCaps: Union[str, bool] = "",  # 강조점
        Spacing: Union[str, int] = "",  # 자간(-50 ~ 50)
        StrikeOutColor: Union[str, int] = "",
        # 취소선 색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 취소는 0xffffffff(4294967295)
        StrikeOutShape: Union[str, int] = "",  # 취소선 모양(0~12, 0이 일반 취소선)
        StrikeOutType: Union[str, bool] = "",  # 취소선 유무(True/False)
        SubScript: Union[str, bool] = "",  # 아래첨자(True/False)
        SuperScript: Union[str, bool] = "",  # 위첨자(True/False)
        TextColor: (
            Union[str, int]
        ) = "",  # 글자색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
        UnderlineColor: (
            Union[str, int]
        ) = "",  # 밑줄색(RGB, 0x0~0xffffff) ~= hwp.rgb_color(255,255,255), 기본값은 0xffffffff(4294967295)
        UnderlineShape: Union[str, int] = "",  # 밑줄형태(0~12)
        UnderlineType: Union[str, int] = "",  # 밑줄위치(0:없음, 1:하단, 3:상단)
        UseFontSpace: Union[str, bool] = "",  # 글꼴에 어울리는 빈칸(True/False)
        UseKerning: Union[str, bool] = "",  # 커닝 적용(True/False) : 차이가 없다?
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
        d = {
            "Bold": Bold,
            "DiacSymMark": DiacSymMark,
            "Emboss": Emboss,
            "Engrave": Engrave,
            "FaceNameUser": FaceName,
            "FaceNameSymbol": FaceName,
            "FaceNameOther": FaceName,
            "FaceNameJapanese": FaceName,
            "FaceNameHanja": FaceName,
            "FaceNameLatin": FaceName,
            "FaceNameHangul": FaceName,
            "FontTypeUser": 1,
            "FontTypeSymbol": 1,
            "FontTypeOther": 1,
            "FontTypeJapanese": 1,
            "FontTypeHanja": 1,
            "FontTypeLatin": 1,
            "FontTypeHangul": 1,
            "Height": Height * 100,
            "Italic": Italic,
            "OffsetHangul": Offset,
            "OffsetHanja": Offset,
            "OffsetJapanese": Offset,
            "OffsetLatin": Offset,
            "OffsetOther": Offset,
            "OffsetSymbol": Offset,
            "OffsetUser": Offset,
            "OutLineType": OutLineType,
            "RatioHangul": Ratio,
            "RatioHanja": Ratio,
            "RatioJapanese": Ratio,
            "RatioLatin": Ratio,
            "RatioOther": Ratio,
            "RatioSymbol": Ratio,
            "RatioUser": Ratio,
            "ShadeColor": (
                self.rgb_color(ShadeColor)
                if type(ShadeColor) == str and ShadeColor
                else ShadeColor
            ),
            "ShadowColor": (
                self.rgb_color(ShadowColor)
                if type(ShadowColor) == str and ShadowColor
                else ShadowColor
            ),
            "ShadowOffsetX": ShadowOffsetX,
            "ShadowOffsetY": ShadowOffsetY,
            "ShadowType": ShadowType,
            "SizeHangul": Size,
            "SizeHanja": Size,
            "SizeJapanese": Size,
            "SizeLatin": Size,
            "SizeOther": Size,
            "SizeSymbol": Size,
            "SizeUser": Size,
            "SmallCaps": SmallCaps,
            "SpacingHangul": Spacing,
            "SpacingHanja": Spacing,
            "SpacingJapanese": Spacing,
            "SpacingLatin": Spacing,
            "SpacingOther": Spacing,
            "SpacingSymbol": Spacing,
            "SpacingUser": Spacing,
            "StrikeOutColor": StrikeOutColor,
            "StrikeOutShape": StrikeOutShape,
            "StrikeOutType": StrikeOutType,
            "SubScript": SubScript,
            "SuperScript": SuperScript,
            "TextColor": (
                self.rgb_color(TextColor)
                if type(TextColor) == str and TextColor
                else TextColor
            ),
            "UnderlineColor": (
                self.rgb_color(UnderlineColor)
                if type(UnderlineColor) == str and UnderlineColor
                else UnderlineColor
            ),
            "UnderlineShape": UnderlineShape,
            "UnderlineType": UnderlineType,
            "UseFontSpace": UseFontSpace,
            "UseKerning": UseKerning,
        }

        if FaceName in self.htf_fonts.keys():
            d |= self.htf_fonts[FaceName]

        pset = self.hwp.HParameterSet.HCharShape
        self.HAction.GetDefault("CharShape", pset.HSet)
        for key in d.keys():
            if d[key] != "":
                pset.__setattr__(key, d[key])
        return self.hwp.HAction.Execute("CharShape", pset.HSet)

    def cell_fill(self, face_color: Tuple[int, int, int] = (217, 217, 217)):
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

    def set_row_height(
        self, height: Union[int, float], as_: Literal["mm", "hwpunit"] = "mm"
    ) -> bool:
        """
        캐럿이 표 안에 있는 경우

        캐럿이 위치한 행의 셀 높이를 조절하는 메서드(기본단위는 mm)

        Args:
            height: 현재 행의 높이 설정값(기본단위는 mm)

        Returns:
            성공시 True, 실패시 False 리턴
        """
        if not self.is_cell():
            raise AssertionError(
                "캐럿이 표 안에 있지 않습니다. 표 안에서 실행해주세요."
            )
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

    def gradation_on_cell(
        self,
        color_list: Union[List[Tuple[int, int, int]], List[str]] = [
            (0, 0, 0),
            (255, 255, 255),
        ],
        grad_type: Literal["Linear", "Radial", "Conical", "Square"] = "Linear",
        angle: int = 0,
        xc: int = 0,
        yc: int = 0,
        pos_list: List[int] = None,
        step_center: int = 50,
        step: int = 255,
    ) -> bool:
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
            raise AssertionError(
                "캐럿이 현재 표 안에 위치하지 않습니다. 표 안에서 다시 실행해주세요."
            )
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
            pset.FillAttr.GradationColor.SetItem(
                0, self.rgb_color(255, 255, 255)
            )  # 시작색 ~ 끝색
            pset.FillAttr.GradationColor.SetItem(
                1, self.rgb_color(255, 255, 255)
            )  # 시작색 ~ 끝색
            pset.FillAttr.GradationBrush = 1
            self.hwp.HAction.Execute("CellFill", pset.HSet)
        color_num = len(color_list)
        if color_num == 1:
            step = 1
        pset.FillAttr.type = self.hwp.BrushType("NullBrush|GradBrush")
        pset.FillAttr.GradationType = self.hwp.Gradation(
            grad_type
        )  # 0은 검정. Linear:1, Radial:2, Conical:3, Square:4
        pset.FillAttr.GradationCenterX = xc  # 가로중심
        pset.FillAttr.GradationCenterY = yc  # 세로중심
        pset.FillAttr.GradationAngle = angle  # 기울임
        pset.FillAttr.GradationStep = (
            step  # 번짐정도(영역개수) 2~255 (0은 투명, 1은 시작색)
        )
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
                pset.FillAttr.GradationColor.SetItem(
                    i, self.rgb_color(color_list[i])
                )  # 시작색 ~ 끝색
            elif check_tuple_of_ints(color_list[i]):
                pset.FillAttr.GradationColor.SetItem(
                    i, self.rgb_color(*color_list[i])
                )  # 시작색 ~ 끝색
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
            result_list.append(self.CharShape.Item("FaceNameHangul"))
            if cur_face == self.CharShape.Item("FaceNameHangul"):
                break
        return list(set(result_list))

    def get_charshape(self):
        """
        현재 캐럿의 글자모양 파라미터셋을 리턴하는 메서드.
        변수로 저장해 두고, set_charshape을 통해
        특정 선택영역에 이 글자모양을 적용할 수 있다.
        """
        pset = self.hwp.HParameterSet.HCharShape
        self.hwp.HAction.GetDefault("CharShape", pset.HSet)
        return pset

    def get_charshape_as_dict(self) -> dict:
        """
        현재 캐럿의 글자모양 파라미터셋을 (보기좋게) dict로 리턴하는 메서드.
        get_charshape와 동일하게 set_charshape에 이 dict를 사용할 수도 있다.
        """
        result_dict = {}
        for key in self.HParameterSet.HCharShape._prop_map_get_.keys():
            result_dict[key] = self.CharShape.Item(key)
        return result_dict

    def set_charshape(self, pset):
        """
        get_charshape 또는 get_charshape_as_dict를 통해 저장된 파라미터셋을 통해
        캐럿위치 또는 특정 선택영역에 해당 글자모양을 적용할 수 있다.

        Args:
            pset: hwp.HParameterSet.HCharShape 파라미터셋 또는 get_charshape_as_dict 결과
        """
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

    def get_parashape(self):
        """
        현재 캐럿이 위치한 문단모양 파라미터셋을 리턴하는 메서드.
        변수로 저장해 두고, set_parashape을 통해
        특정 선택영역에 이 문단모양을 적용할 수 있다.
        """
        pset = self.hwp.HParameterSet.HParaShape
        self.hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
        return pset

    def get_parashape_as_dict(self):
        """
        현재 캐럿의 문단모양 파라미터셋을 (보기좋게) dict로 리턴하는 메서드.
        get_parashape와 동일하게 set_parashape에 이 dict를 사용할 수도 있다.
        """
        result_dict = {}
        for key in self.hwp.HParameterSet.HParaShape._prop_map_get_.keys():
            result_dict[key] = self.ParaShape.Item(key)
        return result_dict

    def set_parashape(self, pset):
        """
        get_parashape 또는 get_parashape_as_dict를 통해 저장된 파라미터셋을 통해
        캐럿이 포함된 문단 또는 블록선택한 문단 전체에 해당 문단모양을 적용할 수 있다.
        """
        if isinstance(pset, dict):
            new_pset = self.hwp.HParameterSet.HParaShape
            for key in pset.keys():
                try:
                    new_pset.__setattr__(key, pset[key])
                except pythoncom.com_error:
                    print(key, pset[key])
        elif type(pset) == type(self.HParameterSet.HParaShape):
            new_pset = pset
        return self.hwp.HAction.Execute("ParagraphShape", new_pset.HSet)

    def set_para(
            self,
            AlignType: Optional[Literal["Justify", "Left", "Center", "Right", "Distribute", "DistributeSpace"]] = None,
            BreakNonLatinWord: Optional[Literal[0, 1]] = None,
            LineSpacing: Optional[int] = None,
            Condense: Optional[int] = None,
            SnapToGrid: Optional[Literal[0, 1]] = None,
            NextSpacing: Optional[float] = None,
            PrevSpacing: Optional[float] = None,
            Indentation: Optional[float] = None,
            RightMargin: Optional[float] = None,
            LeftMargin: Optional[float] = None,
            PagebreakBefore: Optional[Literal[0, 1]] = None,
            KeepLinesTogether: Optional[Literal[0, 1]] = None,
            KeepWithNext: Optional[Literal[0, 1]] = None,
            WidowOrphan: Optional[Literal[0, 1]] = None,
            AutoSpaceEAsianNum: Optional[Literal[0, 1]] = None,
            AutoSpaceEAsianEng: Optional[Literal[0, 1]] = None,
            LineWrap: Optional[Literal[0, 1]] = None,
            FontLineHeight: Optional[Literal[0, 1]] = None,
            TextAlignment: Optional[Literal[0, 1, 2, 3]] = None,
    ):
        """
        문단 모양을 설정하는 단축메서드. set_font와 유사하게 함수처럼 문단 모양을 설정할 수 있다.
        미리 정의된 별도의 파라미터셋을 통해 문단모양을 적용하고 싶다면 set_parashape 메서드를 사용한다.

        Args:
            AlignType: 문단의 정렬 유형을 지정

                - "Justify": 양쪽 정렬
                - "Left": 왼쪽 정렬
                - "Center": 가운데 정렬
                - "Right": 오른쪽 정렬
                - "Distribute": 배분 정렬
                - "DistributeSpace": 나눔 정렬

            BreakNonLatinWord: 한글 단위 줄 나눔 기준

                - 0: 줄 끝에서 어절 단위로 나눔
                - 1: 줄 끝에서 글자 단위로 나눔

            LineSpacing: 단락 내 줄 간격 설정(0~500)
            Condense: 줄 나눔 기준 최소 공백(25~100)
            SnapToGrid: 편집 용지의 줄 격자 사용 여부.
            NextSpacing: 문단 아래 간격(0.0~841.8포인트 범위)
            PrevSpacing: 문단 위 간격(0.0~841.8포인트 범위)
            Indentation: 첫 줄 들여쓰기/내어쓰기(여백 없는 A4용지 기준 -570.2~570.2포인트 범위)
            RightMargin: 오른쪽 여백(포인트)
            LeftMargin: 왼쪽 여백(포인트)
            PagebreakBefore: 문단 앞에서 항상 쪽 나눔 여부(0~1)
            KeepLinesTogether: 문단 보호 여부(0~1)
            KeepWithNext: 다음 문단과 함께
            WidowOrphan: 외톨이줄 보호(한 페이지에 최소 두 줄을 유지)여부
            AutoSpaceEAsianNum: 한글과 숫자 간격 자동 조절 여부
            AutoSpaceEAsianEng: 한글과 영어 간격 자동 조절 여부
            LineWrap: 한 줄로 입력 여부
            FontLineHeight: 글꼴에 어울리는 줄 높이 여부
            TextAlignment: 세로 정렬 방식

                - 0: 글꼴 기준
                - 1: 위쪽 기준
                - 2: 가운데 기준
                - 3: 아래쪽 기준

        Returns:
            주어진 파라미터 세트로 "ParagraphShape" 작업을 실행한 결과
        """
        pset = self.hwp.HParameterSet.HParaShape
        self.hwp.HAction.GetDefault("ParagraphShape", pset.HSet)

        setters = {
            "AlignType": lambda v: setattr(pset, "AlignType", self.HAlign(v) if isinstance(v, str) else v),
            "BreakNonLatinWord": lambda v: setattr(pset, "BreakNonLatinWord",
                                                   0 if v == -1 and 1 <= pset.AlignType <= 3 else (
                                                       1 if v == -1 else v)),
            "LineSpacing": lambda v: setattr(pset, "LineSpacing", v),
            "Condense": lambda v: setattr(pset, "Condense", 100 - v),
            "SnapToGrid": lambda v: setattr(pset, "SnapToGrid", v),
            "NextSpacing": lambda v: setattr(pset, "NextSpacing", self.PointToHwpUnit(v * 2)),
            "PrevSpacing": lambda v: setattr(pset, "PrevSpacing", self.PointToHwpUnit(v * 2)),
            "Indentation": lambda v: setattr(pset, "Indentation", self.PointToHwpUnit(v * 2)),
            "RightMargin": lambda v: setattr(pset, "RightMargin", self.PointToHwpUnit(v * 2)),
            "LeftMargin": lambda v: setattr(pset, "LeftMargin", self.PointToHwpUnit(v * 2)),
            "PagebreakBefore": lambda v: setattr(pset, "PagebreakBefore", v),
            "KeepLinesTogether": lambda v: setattr(pset, "KeepLinesTogether", v),
            "KeepWithNext": lambda v: setattr(pset, "KeepWithNext", v),
            "WidowOrphan": lambda v: setattr(pset, "WidowOrphan", v),
            "AutoSpaceEAsianNum": lambda v: setattr(pset, "AutoSpaceEAsianNum", v),
            "AutoSpaceEAsianEng": lambda v: setattr(pset, "AutoSpaceEAsianEng", v),
            "LineWrap": lambda v: setattr(pset, "LineWrap", v),
            "FontLineHeight": lambda v: setattr(pset, "FontLineHeight", v),
            "TextAlignment": lambda v: setattr(pset, "TextAlignment", v),
        }
        for name, setter in setters.items():
            val = locals()[name]
            if val is not None:
                setter(val)
        return self.HAction.Execute("ParagraphShape", pset.HSet)

    def apply_parashape(self, para_dict: dict) -> bool:
        """
        hwp.get_parashape_as_dict() 메서드를 통해 저장한 문단모양을 다른 문단에 적용할 수 있는 메서드.
        아직은 모든 속성을 지원하지 않는다. 저장 및 적용 가능한 파라미터아이템은 아래 19가지이다.
        특정 파라미터 아이템만 적용하고 싶은 경우 apply_parashape 메서드를 복사하여 커스텀해도 되지만,
        가급적 set_para 메서드를 직접 사용하는 것을 추천한다.

        "AlignType", "BreakNonLatinWord", "LineSpacing", "Condense", "SnapToGrid",
        "NextSpacing", "PrevSpacing", "Indentation", "RightMargin", "LeftMargin",
        "PagebreakBefore", "KeepLinesTogether", "KeepWithNext", "WidowOrphan",
        "AutoSpaceEAsianNum", "AutoSpaceEAsianEng", "LineWrap", "FontLineHeight",
        "TextAlignment"

        Args:
            para_dict: 미리 hwp.get_parashape_as_dict() 메서드로 추출해놓은 문단속성 딕셔너리.

        Returns:
            성공시 True, 실패시 False를 리턴(사실 모든 경우 True를 리턴하는 셈이다.)

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 모양을 복사하고자 하는 문단에서
            >>> para_dict = hwp.get_parashape_as_dict()
            >>> # 모양을 붙여넣기하려는 문단으로 이동 후
            >>> hwp.apply_parashape(para_dict)
            True
            >>>
            >>> # Border나 Heading 속성 등 모든 문단모양을 적용하기 위해서는
            >>> # 아래와 같이 get_parashape과 set_parashape을 사용하면 된다.
            >>>
            >>> # 모양을 복사하고자 하는 문단에서
            >>> parashape = hwp.get_parashape()
            >>> # 모양을 붙여넣기하려는 문단으로 이동 후
            >>> hwp.set_parashape(parashape)
            True
            >>>
            >>> # 마지막으로, 특정 문단속성만 지정하여 변경하고자 할 때에는
            >>> # hwp.set_para 메서드가 가장 간편하다. (파라미터셋의 아이템과 이름은 동일하다.)
            >>> hwp.set_para(AlignType="Justify", LineSpacing=160, LeftMargin=0)
            True
        """
        param_names = [
            "AlignType", "BreakNonLatinWord", "LineSpacing", "Condense", "SnapToGrid",
            "NextSpacing", "PrevSpacing", "Indentation", "RightMargin", "LeftMargin",
            "PagebreakBefore", "KeepLinesTogether", "KeepWithNext", "WidowOrphan",
            "AutoSpaceEAsianNum", "AutoSpaceEAsianEng", "LineWrap", "FontLineHeight",
            "TextAlignment",
        ]

        kwargs = {
            name: para_dict[name]
            for name in param_names
            if name in para_dict and isinstance(para_dict[name], (int, float))
        }

        return self.set_para(**kwargs)

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
            with zipfile.ZipFile(temp_filename, "r") as zf:
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
        code_to_desc = {
            "PaperWidth": "용지폭",
            "PaperHeight": "용지길이",
            "Landscape": "용지방향",  # 0: 가로, 1:세로
            "GutterType": "제본타입",  # 0: 한쪽, 1:맞쪽, 2:위쪽
            "TopMargin": "위쪽",
            "HeaderLen": "머리말",
            "LeftMargin": "왼쪽",
            "GutterLen": "제본여백",
            "RightMargin": "오른쪽",
            "FooterLen": "꼬리말",
            "BottomMargin": "아래쪽",
        }

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
                    result_dict[code_to_desc[key]] = self.hwp_unit_to_mili(
                        eval(f"pset.PageDef.{key}")
                    )
                else:
                    result_dict[key] = self.hwp_unit_to_mili(
                        eval(f"pset.PageDef.{key}")
                    )

        return result_dict

    def set_pagedef(
        self, pset: Any, apply: Literal["cur", "all", "new"] = "cur"
    ) -> bool:
        """
        get_pagedef 또는 get_pagedef_as_dict를 통해 얻은 용지정보를 현재구역에 적용하는 메서드

        Args:
            pset: 파라미터셋 또는 dict. 용지정보를 담은 객체

        Returns:
            성공시 True, 실패시 False를 리턴
        """
        if isinstance(pset, dict):
            desc_to_code = {
                "용지폭": "PaperWidth",
                "용지길이": "PaperHeight",
                "용지방향": "Landscape",
                "제본타입": "GutterType",
                "위쪽": "TopMargin",
                "머리말": "HeaderLen",
                "왼쪽": "LeftMargin",
                "제본여백": "GutterLen",
                "오른쪽": "RightMargin",
                "꼬리말": "FooterLen",
                "아래쪽": "BottomMargin",
            }

            new_pset = self.hwp.HParameterSet.HSecDef
            for key in pset.keys():
                if key in desc_to_code.keys():  # 한글인 경우
                    if key in ["용지방향", "제본여백"]:
                        exec(f"new_pset.PageDef.{desc_to_code[key]} = {pset[key]}")
                    else:
                        exec(
                            f"new_pset.PageDef.{desc_to_code[key]} = {self.mili_to_hwp_unit(pset[key])}"
                        )
                elif key in desc_to_code.values():  # 영문인 경우
                    if key in ["Landscape", "GutterLen"]:
                        exec(f"new_pset.PageDef.{key} = {pset[key]}")
                    else:
                        exec(
                            f"new_pset.PageDef.{key} = {self.mili_to_hwp_unit(pset[key])}"
                        )

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

    def goto_page(self, page_index: Union[int, str] = 1) -> Tuple[int, int]:
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

    def table_from_data(
        self,
        data: Union[pd.DataFrame, dict, list, str],
        transpose: bool = False,
        header0: str = "",
        treat_as_char: bool = False,
        header: bool = True,
        index: bool = True,
        cell_fill: Union[bool, Tuple[int, int, int]] = False,
        header_bold: bool = True,
    ) -> None:
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
            self.create_table(
                rows=len(df) + 1,
                cols=len(df.columns) + 1,
                treat_as_char=treat_as_char,
                header=header,
            )
            self.insert_text(header0)
            self.TableRightCellAppend()
        else:
            self.create_table(
                rows=len(df) + 1,
                cols=len(df.columns),
                treat_as_char=treat_as_char,
                header=header,
            )
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

    def open_pdf(self, pdf_path: str, this_window: int = 1) -> bool:
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

    def insert_file(
        self,
        filename,
        keep_section=1,
        keep_charshape=1,
        keep_parashape=1,
        keep_style=1,
        move_doc_end=False,
    ):
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

    def insert_memo(
        self, text: str = "", memo_type: Literal["revision", "memo"] = "memo"
    ) -> None:
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

    def find_backward(self, src: str, regex: bool = False) -> bool:
        """
        문서 위쪽으로 find 메서드를 수행.

        해당 단어를 선택한 상태가 되며,
        문서 처음에 도달시 False 리턴

        Args:
            src: 찾을 단어

        Returns:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴, 단어가 더이상 없으면 False를 리턴
        """
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("FindDlg", pset.HSet)
        self.hwp.HAction.Execute("FindDlg", pset.HSet)
        self.SetMessageBoxMode(0x2FFF1)
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
            self.SetMessageBoxMode(0xFFFFF)

    def find_forward(self, src: str, regex: bool = False) -> bool:
        """
        문서 아래쪽으로 find를 수행하는 메서드.

        해당 단어를 선택한 상태가 되며,
        문서 끝에 도달시 False 리턴.

        Args:
            src: 찾을 단어

        Returns:
            단어를 찾으면 찾아가서 선택한 후 True를 리턴, 단어가 더이상 없으면 False를 리턴
        """
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("FindDlg", pset.HSet)
        self.hwp.HAction.Execute("FindDlg", pset.HSet)
        self.SetMessageBoxMode(0x2FFF1)
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
            self.SetMessageBoxMode(0xFFFFF)

    def find(
            self,
            src: str = "",
            direction: Literal["Forward", "Backward", "AllDoc"] = "Forward",
            regex: bool = False,
            TextColor: Optional[int] = None,
            Height: Optional[int|float] = None,
            MatchCase: int = 1,
            SeveralWords: int = 0,
            UseWildCards: int = 1,
            WholeWordOnly: int = 0,
            AutoSpell: int = 1,
            HanjaFromHangul: int = 1,
            AllWordForms: int = 0,
            FindStyle: str = "",
            ReplaceStyle: str = "",
            FindJaso: int = 0,
            FindType: int = 1,
    ) -> bool:
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
            TextColor: 글자색(hwp.RGBColor)
            Height: 글자크기(hwp.PointToHwpUnit)
            MatchCase: 대소문자 구분(기본값 1)
            SeveralWords: 여러 단어 찾기(콤마로 구분하여 or연산 실시, 기본값 0)
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
        self.SetMessageBoxMode(0x2FFF1)
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("FindDlg", pset.HSet)
        self.hwp.HAction.Execute("FindDlg", pset.HSet)
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
        if TextColor is not None:
            pset.FindCharShape.TextColor = TextColor
        if Height is not None:
            pset.FindCharShape.Height = self.PointToHwpUnit(Height)
        try:
            return self.hwp.HAction.Execute("RepeatFind", pset.HSet)
        finally:
            self.SetMessageBoxMode(0xFFFFF)

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
        cur_loc = self.get_pos()
        self.MoveDocBegin()
        while self.find(r"\{\{[^(\}\})]+\}\}", regex=True):
            field_name = self.get_selected_text(keep_select=True)[2:-2]
            if ":" in field_name:
                field_name, direction = field_name.split(":", maxsplit=1)
                if ":" in direction:
                    direction, memo = direction.split(":", maxsplit=1)
                else:
                    memo = ""
            else:
                direction = field_name
                memo = ""
            if self.get_selected_text(keep_select=True).endswith("\r\n"):
                raise Exception("필드를 닫는 중괄호가 없습니다.")
            self.hwp.HAction.Run("Delete")
            self.create_field(field_name, direction, memo)

        self.MoveDocBegin()
        while self.find(r"\[\[[^\]\]]+\]\]", regex=True):
            field_name = self.get_selected_text(keep_select=True)[2:-2]
            if ":" in field_name:
                field_name, direction = field_name.split(":", maxsplit=1)
                if ":" in direction:
                    direction, memo = direction.split(":", maxsplit=1)
                else:
                    memo = ""
            else:
                direction = field_name
                memo = ""
            if self.get_selected_text(keep_select=True).endswith("\r\n"):
                raise Exception("필드를 닫는 중괄호가 없습니다.")
            self.hwp.HAction.Run("Delete")
            if self.is_cell():
                self.set_cur_field_name(
                    field_name, option=1, direction=direction, memo=memo
                )
            else:
                pass
        self.set_pos(*cur_loc)

    def find_replace(
        self,
        src,
        dst,
        regex=False,
        direction: Literal["Backward", "Forward", "AllDoc"] = "Forward",
        MatchCase=1,
        AllWordForms=0,
        SeveralWords=1,
        UseWildCards=1,
        WholeWordOnly=0,
        AutoSpell=1,
        IgnoreFindString=0,
        IgnoreReplaceString=0,
        ReplaceMode=1,
        HanjaFromHangul=1,
        FindJaso=0,
        FindStyle="",
        ReplaceStyle="",
        FindType=1,
    ):
        """
        아래아한글의 찾아바꾸기와 동일한 액션을 수항해지만,

        re=True로 설정하고 실행하면,
        문단별로 잘라서 문서 전체를 순회하며
        파이썬의 re.sub 함수를 실행한다.
        """
        self.SetMessageBoxMode(0x2FFF1)
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("FindDlg", pset.HSet)
        self.hwp.HAction.Execute("FindDlg", pset.HSet)
        if regex:
            whole_text = self.get_text_file()
            src_list = [i.group() for i in re.finditer(src, whole_text)]
            dst_list = [re.sub(src, dst, i) for i in src_list]
            for i, j in zip(src_list, dst_list):
                try:
                    return self.find_replace(
                        i,
                        j,
                        direction=direction,
                        MatchCase=MatchCase,
                        AllWordForms=AllWordForms,
                        SeveralWords=SeveralWords,
                        UseWildCards=UseWildCards,
                        WholeWordOnly=WholeWordOnly,
                        AutoSpell=AutoSpell,
                        IgnoreFindString=IgnoreFindString,
                        IgnoreReplaceString=IgnoreReplaceString,
                        ReplaceMode=ReplaceMode,
                        HanjaFromHangul=HanjaFromHangul,
                        FindJaso=FindJaso,
                        FindStyle=FindStyle,
                        ReplaceStyle=ReplaceStyle,
                        FindType=FindType,
                    )
                finally:
                    self.SetMessageBoxMode(0xFFFFF)

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
                self.SetMessageBoxMode(0xFFFFF)

    def find_replace_all(
        self,
        src,
        dst,
        regex=False,
        MatchCase=1,
        AllWordForms=0,
        SeveralWords=1,
        UseWildCards=1,
        WholeWordOnly=0,
        AutoSpell=1,
        IgnoreFindString=0,
        IgnoreReplaceString=0,
        ReplaceMode=1,
        HanjaFromHangul=1,
        FindJaso=0,
        FindStyle="",
        ReplaceStyle="",
        FindType=1,
    ):
        """
        아래아한글의 찾아바꾸기와 동일한 액션을 수항해지만,

        re=True로 설정하고 실행하면,
        문단별로 잘라서 문서 전체를 순회하며
        파이썬의 re.sub 함수를 실행한다.
        """
        self.SetMessageBoxMode(0x2FFF1)
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("FindDlg", pset.HSet)
        self.hwp.HAction.Execute("FindDlg", pset.HSet)
        if regex:
            whole_text = self.get_text_file()
            src_list = [i.group() for i in re.finditer(src, whole_text)]
            dst_list = [re.sub(src, dst, i) for i in src_list]
            for i, j in zip(src_list, dst_list):
                self.find_replace_all(
                    i,
                    j,
                    MatchCase=MatchCase,
                    AllWordForms=AllWordForms,
                    SeveralWords=SeveralWords,
                    UseWildCards=UseWildCards,
                    WholeWordOnly=WholeWordOnly,
                    AutoSpell=AutoSpell,
                    IgnoreFindString=IgnoreFindString,
                    IgnoreReplaceString=IgnoreReplaceString,
                    ReplaceMode=ReplaceMode,
                    HanjaFromHangul=HanjaFromHangul,
                    FindJaso=FindJaso,
                    FindStyle=FindStyle,
                    ReplaceStyle=ReplaceStyle,
                    FindType=FindType,
                )
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
                self.SetMessageBoxMode(0xFFFFF)

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
            inner_param = (
                text.split("{\r\n")[1]
                .split("\r\n}")[0]
                .replace("        ", f"    {pset_name}")
            )
            result = (
                f"def script_macro():\r\n    pset = {pset_name}\r\n    "
                + text.replace("    ", "")
                .split("with")[0]
                .replace(pset_name, "pset")
                .replace("\r\n", "\r\n    ")
                + inner_param.replace(pset_name, "pset.")
                .replace("    ", "")
                .replace("}\r\n", "")
                .replace("..", ".")
                .replace("\r\n", "\r\n    ")
            )
        else:
            pset_name = text.split(", ")[1].split(".HSet")[0]
            result = (
                f"def script_macro():\r\n    pset = {pset_name}\r\n    "
                + text.replace("    ", "")
                .replace(pset_name, "pset")
                .replace("\r\n", "\r\n    ")
            )
        result = result.replace("HAction.", "hwp.HAction.").replace(
            "HParameterSet.", "hwp.HParameterSet."
        )
        result = re.sub(r"= (?!hwp\.)(\D)", r"= hwp.\g<1>", result)
        result = result.replace('hwp."', '"')
        print(result)
        cb.copy(result)

    def clear_field_text(self):
        for i in self.hwp.GetFieldList(1).split("\x02"):
            self.hwp.PutFieldText(i, "")

    @property
    def doc_list(self) -> List[XHwpDocument]:
        return self.XHwpDocuments

    def switch_to(self, num: int) -> Optional[XHwpDocument]:
        """
        여러 개의 hwp인스턴스가 열려 있는 경우 해당 인덱스의 문서창 인스턴스를 활성화한다.

        Args:
            num: 전환할 문서 인스턴스 아이디(1부터 시작)

        Returns:
            문서 전환에 성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.add_doc()
            >>> hwp.add_tab()
            >>> hwp.switch_to(0)
        """
        if num < self.hwp.XHwpDocuments.Count:
            self.hwp.XHwpDocuments[num].SetActive_XHwpDocument()
            return XHwpDocument(self.XHwpDocuments[num])
        else:
            return False

    def add_tab(self) -> XHwpDocument:
        """
        새 문서를 현재 창의 새 탭에 추가한다.

        백그라운드 상태에서 새 창을 만들 때 윈도우에 나타나는 경우가 있는데,
        add_tab() 함수를 사용하면 백그라운드 작업이 보장된다.
        탭 전환은 switch_to() 메서드로 가능하다.

        새 창을 추가하고 싶은 경우는 add_tab 대신 hwp.FileNew()나 hwp.add_doc()을 실행하면 된다.
        """
        return XHwpDocument(self.hwp.XHwpDocuments.Add(1))  # 0은 새 창, 1은 새 탭

    def add_doc(self) -> XHwpDocument:
        """
        새 문서를 추가한다.

        원래 창이 백그라운드로 숨겨져 있어도 추가된 문서는 보이는 상태가 기본값이다.
        숨기려면 `hwp.set_visible(False)`를 실행해야 한다.
        새 탭을 추가하고 싶은 경우는 `add_doc` 대신 `add_tab`을 실행하면 된다.

        Returns:
            XHwpDocument: 생성한 문서 오브젝트

        """
        return XHwpDocument(self.hwp.XHwpDocuments.Add(0))  # 0은 새 창, 1은 새 탭

    def create_table(
        self,
        rows: int = 1,
        cols: int = 1,
        treat_as_char: bool = True,
        width_type: int = 0,
        height_type: int = 0,
        header: bool = True,
        height: int = 0,
    ) -> bool:
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

        Args:
            rows: 행 수
            cols: 열 수
            treat_as_char: 글자처럼 취급 여부
            width_type: 너비 정의 형식
            height_type: 높이 정의 형식
            header: 1행을 제목행으로 설정할지 여부
            height:

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
            sec_def.PageDef.PaperWidth
            - sec_def.PageDef.LeftMargin
            - sec_def.PageDef.RightMargin
            - sec_def.PageDef.GutterLen
            - self.mili_to_hwp_unit(2)
        )

        pset.WidthValue = total_width  # 표 너비(근데 영향이 없는 듯)
        if height and height_type == 1:  # 표높이가 정의되어 있으면
            # 페이지 최대 높이 계산
            total_height = (
                sec_def.PageDef.PaperHeight
                - sec_def.PageDef.TopMargin
                - sec_def.PageDef.BottomMargin
                - sec_def.PageDef.HeaderLen
                - sec_def.PageDef.FooterLen
                - self.mili_to_hwp_unit(2)
            )
            pset.HeightValue = min(
                self.hwp.MiliToHwpUnit(height), total_height
            )  # 표 높이
            pset.CreateItemArray("RowHeight", rows)  # 행 m개 생성
            each_row_height = min(
                (
                    self.mili_to_hwp_unit(height)
                    - self.mili_to_hwp_unit((0.5 + 0.5) * rows)
                )
                // rows,
                (total_height - self.mili_to_hwp_unit((0.5 + 0.5) * rows)) // rows,
            )
            for i in range(rows):
                pset.RowHeight.SetItem(i, each_row_height)  # 1열
            pset.TableProperties.Height = min(
                self.MiliToHwpUnit(height),
                total_height - self.mili_to_hwp_unit((0.5 + 0.5) * rows),
            )

        pset.CreateItemArray("ColWidth", cols)  # 열 n개 생성
        each_col_width = round((total_width - self.mili_to_hwp_unit(3.6 * cols)) / cols)
        for i in range(cols):
            pset.ColWidth.SetItem(i, each_col_width)  # 1열
        if self.Version[0] == "8":
            pset.TableProperties.TreatAsChar = treat_as_char  # 글자처럼 취급
        pset.TableProperties.Width = (
            total_width  # self.hwp.MiliToHwpUnit(148)  # 표 너비
        )
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

    def get_selected_text(self, as_: Literal["list", "str"] = "str", keep_select: bool = False):
        """
        한/글 문서 선택 구간의 텍스트를 리턴하는 메서드.
        표 안에 있을 때는 셀의 문자열을, 본문일 때는 선택영역 또는 현재 단어를 리턴.

        Args:
            as_: 문자열 형태로 리턴할지("str"), 리스트 형태로 리턴할지("list") 결정.
        Returns:
            선택한 문자열 또는 셀 문자열
        """
        if self.SelectionMode == 0:
            if self.is_cell():
                self.TableCellBlock()
            else:
                self.Select()
                self.Select()
        if not self.hwp.InitScan(Range=0xFF):
            self.Cancel()
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
        if not keep_select:
            self.Cancel()
        return result if type(result) == str else result[:-1]

    def table_to_csv(
        self, n="", filename="result.csv", encoding="utf-8", startrow=0
    ) -> None:
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
                raise IndexError(
                    f"해당 인덱스의 표가 존재하지 않습니다."
                    f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다."
                )
            self.hwp.FindCtrl()
            self.ShapeObjTableSelCell()
        data = [self.get_selected_text(keep_select=True)]
        col_count = 1
        start = False
        while self.TableRightCell():
            if not startrow:
                if re.match(r"\([A-Z]+1\)", self.hwp.KeyIndicator()[-1]):
                    col_count += 1
                data.append(self.get_selected_text(keep_select=True))
            else:
                if re.match(rf"\([A-Z]+{1 + startrow}\)", self.hwp.KeyIndicator()[-1]):
                    col_count += 1
                    start = True
                if start:
                    data.append(self.get_selected_text(keep_select=True))

        array = np.array(data).reshape(-1, col_count)
        df = pd.DataFrame(array[1:], columns=array[0])
        self.hwp.SetPos(*start_pos)
        df.to_csv(filename, index=False, encoding=encoding)
        self.hwp.SetPos(*start_pos)
        print(os.path.join(os.getcwd(), filename))
        return None

    def table_to_df_q(
        self, n: str = "", startrow: int = 0, columns: list = []
    ) -> pd.DataFrame:
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
                raise IndexError(
                    f"해당 인덱스의 표가 존재하지 않습니다."
                    f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다."
                )
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

        arr = np.array(self.get_selected_text(as_="list", keep_select=True), dtype=object).reshape(
            rows, -1
        )
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

    def table_to_df(
        self, n="", cols=0, selected_range=None, start_pos=None
    ) -> pd.DataFrame:
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
                    raise IndexError(
                        f"해당 인덱스의 표가 존재하지 않습니다."
                        f"현재 문서에는 표가 {abs(int(idx + 0.1))}개 존재합니다."
                    )
                self.hwp.FindCtrl()
        else:
            selected_range = self.get_selected_range()
        xml_data = self.GetTextFile("HWPML2X", option="saveblock")
        root = ET.fromstring(xml_data)

        # TABLE 태그에 RowCount, ColCount가 있으면 사용하고, 없으면 ROW, CELL 수로 결정
        table_el = root.find(".//TABLE")
        if table_el is not None:
            row_count = int(table_el.attrib.get("RowCount", "0"))
            col_count = int(table_el.attrib.get("ColCount", "0"))
        else:
            rows = root.findall(".//ROW")
            row_count = len(rows)
            col_count = max(len(row.findall(".//CELL")) for row in rows)

        # 결과를 저장할 2차원 리스트 초기화 (빈 문자열로 채움)
        result = [["" for _ in range(col_count)] for _ in range(row_count)]

        row_index = 0
        for row in root.findall(".//ROW"):
            col_index = 0
            for cell in row.findall(".//CELL"):
                # 이미 값이 채워진 셀이 있으면 건너뛰고 다음 빈 칸 찾기
                while col_index < col_count and result[row_index][col_index] != "":
                    col_index += 1
                if col_index >= col_count:
                    break

                # CELL 내 텍스트 추출 (CHAR 태그의 텍스트 연결)
                cell_text = ""
                for text in cell.findall(".//TEXT"):
                    for char in text.findall(".//CHAR"):
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
            data = result[cols + 1 :]
            df = pd.DataFrame(data, columns=columns)
        elif type(cols) in (list, tuple):
            df = pd.DataFrame(result, columns=cols)

        try:
            return df
        finally:
            if self.SelectionMode != 19:
                self.set_pos(*start_pos)

    def table_to_bottom(self, offset: float = 0.0) -> bool:
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
            return self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
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

    def insert_lorem(self, para_num: int = 1) -> bool:
        """
        Lorem Ipsum을 캐럿 위치에 작성한다.

        Args:
            para_num: 삽입할 Lorem Ipsum 문단 갯수

        Returns:
             성공시 True, 실패시 False를 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.insert_lorem(3)
            True
        """
        api_url = f"https://api.api-ninjas.com/v1/loremipsum?paragraphs={para_num}"

        headers = {"X-Api-Key": "hzzbbAAy7mQjKyXSW5quRw==PbJStWB0ymMpGRH1"}

        req = request.Request(api_url, headers=headers)

        try:
            with request.urlopen(req) as response:
                response_text = json.loads(response.read().decode("utf-8"))[
                    "text"
                ].replace("\n", "\r\n")
        except urllib.error.HTTPError as e:
            print("Error:", e.code, e.reason)
        except urllib.error.URLError as e:
            print("Error:", e.reason)
        return self.insert_text(response_text)

    def move_all_caption(
        self,
        location: Literal["Top", "Bottom", "Left", "Right"] = "Bottom",
        align: Literal[
            "Left", "Center", "Right", "Distribute", "Division", "Justify"
        ] = "Justify",
    ):
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

    @property
    def is_empty(self) -> bool:
        """
        아무 내용도 들어있지 않은 빈 문서인지 여부를 나타낸다. 읽기전용
        """
        return self.hwp.IsEmpty

    @property
    def is_modified(self) -> bool:
        """
        최근 저장 또는 생성 이후 수정이 있는지 여부를 나타낸다. 읽기전용
        """
        return self.hwp.IsModified

    # 액션 파라미터용 함수

    def check_xobject(self, bstring):
        return self.hwp.CheckXObject(bstring=bstring)

    def CheckXObject(self, bstring):
        return self.check_xobject(bstring)

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
        return self.clear(option)

    def close(self, is_dirty: bool = False, interval: float = 0.01) -> bool:
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
                return self.hwp.XHwpDocuments.Active_XHwpDocument.Close(
                    isDirty=is_dirty
                )
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
        return self.create_action(actidstr)

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
        return self.create_field(name, direction, memo)

    def create_id(self, creation_id):
        return self.hwp.CreateID(CreationID=creation_id)

    def CreateId(self, creation_id):
        return self.create_id(creation_id)

    def create_mode(self, creation_mode):
        return self.hwp.CreateMode(CreationMode=creation_mode)

    def CreateMode(self, creation_mode):
        return self.create_mode(creation_mode)

    def create_page_image(
        self,
        path: str,
        pgno: int = -1,
        resolution: int = 300,
        depth: int = 24,
        format: str = "bmp",
    ) -> bool:
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
            raise IndexError(
                f"pgno는 -1부터 {self.PageCount}까지 입력 가능합니다. (-1:전체 저장, 0:현재페이지 저장)"
            )
        if path.lower()[1] != ":":
            path = os.path.abspath(path)
        if not os.path.exists(os.path.dirname(path)):
            os.mkdir(os.path.dirname(path))
        ext = path.rsplit(".", maxsplit=1)[-1]
        if pgno >= 0:
            if pgno == 0:
                pgno = self.current_page
            try:
                return self.hwp.CreatePageImage(
                    Path=path,
                    pgno=pgno - 1,
                    resolution=resolution,
                    depth=depth,
                    Format=format,
                )
            finally:
                if not ext.lower() in ("gif", "bmp"):
                    with Image.open(path.replace(ext, format)) as img:
                        img.save(path.replace(format, ext))
                    os.remove(path.replace(ext, format))
        elif pgno == -1:
            for i in range(1, self.PageCount + 1):
                path_ = os.path.join(
                    os.path.dirname(path),
                    os.path.basename(path).replace(f".{ext}", f"{i:03}.{ext}"),
                )
                self.hwp.CreatePageImage(
                    Path=path_,
                    pgno=i - 1,
                    resolution=resolution,
                    depth=depth,
                    Format=format,
                )
                if not ext.lower() in ("gif", "bmp"):
                    with Image.open(path_.replace(ext, format)) as img:
                        img.save(path_.replace(format, ext))
                    os.remove(path_.replace(ext, format))
            return True

    def CreatePageImage(
        self,
        path: str,
        pgno: int = -1,
        resolution: int = 300,
        depth: int = 24,
        format: str = "bmp",
    ) -> bool:
        return self.create_page_image(path, pgno, resolution, depth, format)

    def create_set(self, setidstr: str) -> "Hwp.HParameterSet":
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

    def CreateSet(self, setidstr: str) -> "Hwp.HParameterSet":
        return self.create_set(setidstr)

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
        return self.delete_ctrl(ctrl)

    def EquationCreate(self, thread=False):
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
        return _close_eqedit(save, delay)

    def EquationModify(self, thread=False):
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
        return self.export_style(sty_filepath)

    def field_exist(self, field: str) -> bool:
        """
        문서에 해당 이름의 데이터 필드가 존재하는지 검사한다.

        Args:
            field: 필드이름

        Returns:
            필드가 존재하면 True, 존재하지 않으면 False
        """
        return self.hwp.FieldExist(Field=field)

    def FieldExist(self, field: str) -> bool:
        return self.field_exist(field)

    def file_translate(self, cur_lang: str = "ko", trans_lang: str = "en") -> bool:
        """
        문서를 번역함(Ctrl-Z 안 됨.) 한 달 10,000자 무료

        Args:
            cur_lang: 현재 문서 언어(예 - ko)
            trans_lang: 목표언어(예 - en)

        Returns:
            성공 후 True 리턴(실패하면 프로그램 종료됨ㅜ)
        """
        return self.hwp.FileTranslate(curLang=cur_lang, transLang=trans_lang)

    def FileTranslate(self, cur_lang: str = "ko", trans_lang: str = "en") -> bool:
        return self.file_translate(cur_lang, trans_lang)

    def find_ctrl(self):
        """컨트롤 선택하기"""
        return self.hwp.FindCtrl()

    def FindCtrl(self):
        return self.find_ctrl()

    def find_private_info(self, private_type: int, private_string: str) -> int:
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
        return self.hwp.FindPrivateInfo(
            PrivateType=private_type, PrivateString=private_string
        )

    def FindPrivateInfo(self, private_type: int, private_string: str) -> int:
        return self.find_private_info(private_type, private_string)

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
        return self.get_bin_data_path(binid)

    def get_cur_field_name(self, option: int = 0) -> str:
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

    def GetCurFieldName(self, option: int = 0) -> str:
        return self.get_cur_field_name(option)

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
        return self.get_cur_metatag_name()

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

    def GetFieldList(self, number: int = 1, option: int = 0) -> str:
        return self.get_field_list(number, option)

    def get_field_text(self, field: Union[str, list, tuple, set], idx: int = 0) -> str:
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
        elif isinstance(field, Union[list, tuple, set]):
            return self.hwp.GetFieldText(Field="\x02".join(str(i) for i in field))

    def GetFieldText(self, field: Union[str, list, tuple, set], idx: int = 0) -> str:
        return self.get_field_text(field, idx)

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

    def GetFileInfo(self, filename: str) -> "Hwp.HParameterSet":
        return self.get_file_info(filename)

    def get_font_list(self, langid: str = "") -> List[str]:
        """
        현재 문서에 사용되고 있는 폰트 목록 리턴

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> hwp.get_font_list()
            ['D2Coding,R', 'Pretendard Variable Thin,R', '나눔명조,R', '함초롬바탕,R']
        """
        self.scan_font()
        return [
            i.rsplit(",", maxsplit=1)[0]
            for i in self.hwp.GetFontList(langid=langid).split("\x02")
        ]

    def GetFontList(self, langid: str = "") -> list:
        return self.get_font_list(langid)

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
        return self.get_heading_string()

    def get_message_box_mode(self) -> int:
        """
        현재 메시지 박스의 Mode를 ``int``로 얻어온다.

        set_message_box_mode와 함께 쓰인다. 6개의 대화상자에서 각각 확인/취소/종료/재시도/무시/예/아니오 버튼을 자동으로 선택할 수 있게 설정할 수 있으며 조합 가능하다.
        리턴하는 정수의 의미는 ``set_message_box_mode``를 참고한다.
        """
        return self.hwp.GetMessageBoxMode()

    def GetMessageBoxMode(self):
        return self.get_message_box_mode()

    def get_metatag_list(self, number, option):
        """메타태그리스트 가져오기"""
        return self.hwp.GetMetatagList(Number=number, option=option)

    def GetMetatagList(self, number, option):
        return self.get_metatag_list(number, option)

    def get_metatag_name_text(self, tag):
        """메타태그이름 문자열 가져오기"""
        return self.hwp.GetMetatagNameText(tag=tag)

    def GetMetatagNameText(self, tag):
        return self.get_metatag_name_text(tag)

    def get_mouse_pos(
        self, x_rel_to: int = 1, y_rel_to: int = 1
    ) -> "Hwp.HParameterSet":
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

    def GetMousePos(self, x_rel_to: int = 1, y_rel_to: int = 1) -> "Hwp.HParameterSet":
        return self.get_mouse_pos(x_rel_to, y_rel_to)

    def get_page_text(self, pgno: int = 0, option: hex = 0xFFFFFFFF) -> str:
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

    def GetPageText(self, pgno: int = 0, option: hex = 0xFFFFFFFF) -> str:
        return self.get_page_text(pgno, option)

    def get_pos(self) -> Tuple[int]:
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

    def GetPos(self) -> Tuple[int]:
        return self.get_pos()

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
        return self.get_pos_by_set()

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
        return self.get_script_source(filename)

    def get_selected_pos(self) -> Tuple[bool, str, str, int, str, str, int]:
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
            >>> marked_area = hwp.get_selected_pos()
            >>> # 임의의 영역으로 이동 후, 저장한 구간을 선택하기
            >>> hwp.select_text(marked_area)
            True
        """
        return self.hwp.GetSelectedPos()

    def GetSelectedPos(self) -> Tuple[bool, str, str, int, str, str, int]:
        return self.get_selected_pos()

    def get_selected_pos_by_set(self, sset: Any, eset: Any) -> bool:
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

    def GetSelectedPosBySet(
        self, sset: "Hwp.HParameterSet", eset: "Hwp.HParameterSet"
    ) -> bool:
        return self.get_selected_pos_by_set(sset, eset)

    def get_text(self) -> Tuple[int, str]:
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

    def GetText(self) -> Tuple[int, str]:
        return self.get_text()

    def get_text_file(
        self,
        format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "UNICODE",
        option: str = "saveblock:true",
    ) -> str:
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

    def GetTextFile(
        self,
        format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "UNICODE",
        option: str = "saveblock:true",
    ) -> str:
        return self.get_text_file(format, option)

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
        return self.import_style(sty_filepath)

    def init_hparameterset(self):
        return self.hwp.InitHParameterSet()

    def InitHParameterSet(self):
        return self.init_hparameterset()

    def init_scan(
        self,
        option: int = 0x07,
        range: int = 0x77,
        spara: int = 0,
        spos: int = 0,
        epara: int = -1,
        epos: int = -1,
    ) -> bool:
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
        return self.hwp.InitScan(
            option=option, Range=range, spara=spara, spos=spos, epara=epara, epos=epos
        )

    def InitScan(
        self,
        option: int = 0x07,
        range: int = 0x77,
        spara: int = 0,
        spos: int = 0,
        epara: int = -1,
        epos: int = -1,
    ) -> bool:
        return self.init_scan(option, range, spara, spos, epara, epos)

    def insert(
        self, path: str, format: str = "", arg: str = "", move_doc_end: bool = False
    ) -> bool:
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

    def Insert(
        self, path: str, format: str = "", arg: str = "", move_doc_end: bool = False
    ) -> bool:
        return self.insert(path, format, arg, move_doc_end)

    def insert_background_picture(
        self,
        path: str,
        border_type: Literal["SelectedCell", "SelectedCellDelete"] = "SelectedCell",
        embedded: bool = True,
        filloption: int = 5,
        effect: int = 0,
        watermark: bool = False,
        brightness: int = 0,
        contrast: int = 0,
    ) -> bool:
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
            return self.hwp.InsertBackgroundPicture(
                Path=path,
                BorderType=border_type,
                Embedded=embedded,
                filloption=filloption,
                Effect=effect,
                watermark=watermark,
                Brightness=brightness,
                Contrast=contrast,
            )
        finally:
            if "temp.jpg" in os.listdir():
                os.remove(path)

    def InsertBackgroundPicture(
        self,
        path: str,
        border_type: Literal["SelectedCell", "SelectedCellDelete"] = "SelectedCell",
        embedded: bool = True,
        filloption: int = 5,
        effect: int = 0,
        watermark: bool = False,
        brightness: int = 0,
        contrast: int = 0,
    ) -> bool:
        return self.insert_background_picture(
            path,
            border_type,
            embedded,
            filloption,
            effect,
            watermark,
            brightness,
            contrast,
        )

    def insert_ctrl(self, ctrl_id: str, initparam: Any) -> Ctrl:
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

    def InsertCtrl(self, ctrl_id: str, initparam: "Hwp.HParameterSet") -> Ctrl:
        return self.insert_ctrl(ctrl_id, initparam)

    def insert_picture(
        self,
        path: str,
        treat_as_char: bool = True,
        embedded: bool = True,
        sizeoption: int = 0,
        reverse: bool = False,
        watermark: bool = False,
        effect: int = 0,
        width: int = 0,
        height: int = 0,
    ) -> Ctrl:
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
            raise ValueError(
                "sizeoption이 1일 때에는 width와 height를 지정해주셔야 합니다.\n"
                "단, 셀 안에 있는 경우에는 셀 너비에 맞게 이미지 크기를 자동으로 조절합니다."
            )

        if path.startswith("http"):
            temp_path = tempfile.TemporaryFile().name
            request.urlretrieve(path, temp_path)
            path = temp_path
            # request.urlretrieve(path, os.path.join(os.getcwd(), "temp.jpg"))
        elif path.lower()[1] != ":":
            path = os.path.join(os.getcwd(), path)

        try:
            ctrl = self.hwp.InsertPicture(
                Path=path,
                Embedded=embedded,
                sizeoption=sizeoption,
                Reverse=reverse,
                watermark=watermark,
                Effect=effect,
                Width=width,
                Height=height,
            )
            pic_prop = ctrl.Properties
            if not all([width, height]) and self.is_cell():

                pset = self.HParameterSet.HShapeObject
                self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                if pset.ShapeTableCell.HasMargin == 1:  # 1이면
                    # 특정 셀 안여백
                    cell_pset = self.HParameterSet.HShapeObject
                    self.HAction.GetDefault("TablePropertyDialog", pset.HSet)
                    margin = round(
                        cell_pset.ShapeTableCell.MarginLeft
                        + cell_pset.ShapeTableCell.MarginRight,
                        2,
                    )
                else:
                    # 전역 셀 안여백
                    margin = round(pset.CellMarginLeft + pset.CellMarginRight, 2)

                cell_width = pset.ShapeTableCell.Width - margin
                dst_height = (
                    pic_prop.Item("Height") / pic_prop.Item("Width") * cell_width
                )
                pic_prop.SetItem("Width", cell_width)
                pic_prop.SetItem("Height", round(dst_height))
            else:
                sec_def = self.HParameterSet.HSecDef
                self.HAction.GetDefault("PageSetup", sec_def.HSet)
                page_width = (
                    sec_def.PageDef.PaperWidth
                    - sec_def.PageDef.LeftMargin
                    - sec_def.PageDef.RightMargin
                    - sec_def.PageDef.GutterLen
                )
                page_height = (
                    sec_def.PageDef.PaperHeight
                    - sec_def.PageDef.TopMargin
                    - sec_def.PageDef.BottomMargin
                    - sec_def.PageDef.HeaderLen
                    - sec_def.PageDef.FooterLen
                )
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

    def InsertPicture(
        self,
        path: str,
        treat_as_char: bool = True,
        embedded: bool = True,
        sizeoption: int = 0,
        reverse: bool = False,
        watermark: bool = False,
        effect: int = 0,
        width: int = 0,
        height: int = 0,
    ) -> Ctrl:
        return self.insert_picture(
            path,
            treat_as_char,
            embedded,
            sizeoption,
            reverse,
            watermark,
            effect,
            width,
            height,
        )

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

    def is_action_enable(self, action_id: str) -> bool:
        """
        액션 실행 가능한지 여부를 bool로 리턴

        액션 관련해서는 기존 버전체크보다 이걸 사용하는 게
        훨씬 안정적일 것 같기는 하지만(예: CopyPage, PastePage, DeletePage 및 메타태그액션 등)
        신규 메서드(SelectCtrl 등) 지원여부는 체크해주지 못한다ㅜ
        """
        return self.hwp.IsActionEnable(actionID=action_id)

    def IsActionEnable(self, action_id: str) -> bool:
        return self.is_action_enable(action_id)

    def is_command_lock(self, action_id: str) -> bool:
        """
        해당 액션이 잠겨있는지 확인한다.

        Args:
            action_id: 액션 ID. (ActionIDTable.Hwp 참조)

        Returns:
            잠겨있으면 True, 잠겨있지 않으면 False를 반환한다.
        """
        return self.hwp.IsCommandLock(actionID=action_id)

    def IsCommandLock(self, action_id: str) -> bool:
        return self.is_command_lock(action_id)

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
        return self.key_indicator()

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

    def LockCommand(self, act_id: str, is_lock: bool) -> None:
        return self.lock_command(act_id, is_lock)

    def modify_field_properties(self, field: str, remove: bool, add: bool):
        return self.hwp.ModifyFieldProperties(Field=field, remove=remove, Add=add)

    def ModifyFieldProperties(self, field, remove, add) -> None:
        return self.modify_field_properties(field, remove, add)

    def modify_metatag_properties(self, tag, remove, add):
        return self.hwp.ModifyMetatagProperties(tag=tag, remove=remove, Add=add)

    def ModifyMetatagProperties(self, tag, remove, add):
        return self.modify_metatag_properties(tag, remove, add)

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

    def MovePos(self, move_id: int = 1, para: int = 0, pos: int = 0) -> bool:
        return self.move_pos(move_id, para, pos)

    def move_to_field(
        self,
        field: str,
        idx: int = 0,
        text: bool = True,
        start: bool = True,
        select: bool = False,
    ) -> bool:
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
            return self.hwp.MoveToField(
                Field=f"{field}{{{{{idx}}}}}", Text=text, start=start, select=select
            )
        else:
            return self.hwp.MoveToField(
                Field=field, Text=text, start=start, select=select
            )

    def MoveToField(
        self,
        field: str,
        idx: int = 0,
        text: bool = True,
        start: bool = True,
        select: bool = False,
    ) -> bool:
        return self.move_to_field(field, idx, text, start, select)

    def move_to_metatag(self, tag, text, start, select):
        """특정 메타태그로 이동"""
        return self.hwp.MoveToMetatag(tag=tag, Text=text, start=start, select=select)

    def MoveToMetatag(self, tag, text, start, select):
        return self.move_to_metatag(tag, text, start, select)

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
                hwp_name = [
                    parse.unquote_plus(i)
                    for i in re.split("[/?=&]", filename)
                    if ".hwp" in i
                ][0]
            except IndexError as e:
                # url 문자열 안에 hwp 파일명이 포함되어 있지 않은 경우에는 임시파일명 지정(temp.hwp)
                hwp_name = "temp.hwp"
            request.urlretrieve(filename, os.path.join(os.getcwd(), hwp_name))
            filename = os.path.join(os.getcwd(), hwp_name)
        elif filename.lower()[1] != ":" and os.path.exists(
            os.path.join(os.getcwd(), filename)
        ):
            filename = os.path.join(os.getcwd(), filename)
        return self.hwp.Open(filename=filename, Format=format, arg=arg)

    def Open(self, filename: str, format: str = "", arg: str = "") -> bool:
        return self.open(filename, format, arg)

    def point_to_hwp_unit(self, point: float) -> int:
        """
        글자에 쓰이는 포인트 단위를 HwpUnit으로 변환
        """
        return self.hwp.PointToHwpUnit(Point=point)

    def PointToHwpUnit(self, point: float) -> int:
        return self.point_to_hwp_unit(point)

    @staticmethod
    def hwp_unit_to_point(HwpUnit: int) -> float:
        """
        HwpUnit을 포인트 단위로 변환
        """
        return HwpUnit / 100

    def HwpUnitToPoint(self, HwpUnit: int) -> float:
        return self.hwp_unit_to_point(HwpUnit)

    @staticmethod
    def hwp_unit_to_inch(HwpUnit: int) -> float:
        """
        HwpUnit을 인치로 변환
        """
        if HwpUnit == 0:
            return 0
        else:
            return HwpUnit / 7200

    def HwpUnitToInch(self, HwpUnit: int) -> float:
        return self.hwp_unit_to_inch(HwpUnit)

    @staticmethod
    def inch_to_hwp_unit(inch) -> int:
        """
        인치 단위를 HwpUnit으로 변환
        """
        return round(inch * 7200, 0)

    def InchToHwpUnit(self, inch: float) -> int:
        return self.inch_to_hwp_unit(inch)

    def protect_private_info(
        self, protecting_char: str, private_pattern_type: int
    ) -> bool:
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
        return self.hwp.ProtectPrivateInfo(
            PotectingChar=protecting_char, PrivatePatternType=private_pattern_type
        )

    def ProtectPrivateInfo(
        self, protecting_char: str, private_pattern_type: int
    ) -> bool:
        return self.protect_private_info(protecting_char, private_pattern_type)

    def put_field_text(
        self, field: Any = "", text: Union[str, list, tuple, pd.Series] = "", idx=None
    ) -> None:
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
        if isinstance(field, str) and (
            field.endswith(".xlsx") or field.endswith(".xls")
        ):
            field = pd.read_excel(field)

        if isinstance(field, dict):  # dict 자료형의 경우에는 text를 생략하고
            field, text = list(zip(*list(field.items())))
            field_str = ""
            text_str = ""
            if isinstance(idx, int):
                for f_i, f in enumerate(field):
                    field_str += f"{f}{{{{{idx}}}}}\x02"
                    text_str += (
                        f"{text[f_i][idx]}\x02"  # for t_i, t in enumerate(text[f_i]):
                    )
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
                text_str = "\x02".join([str(field[i]) for i in field.index])
                field_str = "\x02".join([str(i) for i in field.index])  # \x02로 병합
                self.hwp.PutFieldText(Field=field_str, Text=text_str)
                return
            elif type(text) in [
                list,
                tuple,
                pd.Series,
            ]:  # 필드 텍스트를 리스트나 배열로 넣은 경우에도
                text = "\x02".join([str(i) for i in text])  # \x02로 병합
            else:
                raise IOError("text parameter required.")

        if type(field) in [list, tuple]:

            # field와 text가 [[field0:str, list[text:str]], [field1:str, list[text:str]]] 타입인 경우
            if (
                not text
                and isinstance(field[0][0], (str, int, float))
                and not isinstance(field[0][1], (str, int))
                and len(field[0][1]) >= 1
            ):
                text_str = ""
                field_str = "\x02".join(
                    [
                        str(field[i][0]) + f"{{{{{j}}}}}"
                        for j in range(len(field[0][1]))
                        for i in range(len(field))
                    ]
                )
                for i in range(len(field[0][1])):
                    text_str += (
                        "\x02".join([str(field[j][1][i]) for j in range(len(field))])
                        + "\x02"
                    )
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
                field_str = "\x02".join(
                    [str(i) + f"{{{{{idx}}}}}" for i in field]
                )  # \x02로 병합
                text_str += "\x02".join([str(t) for t in field.iloc[idx]]) + "\x02"
            else:
                field_str = "\x02".join(
                    [str(i) + f"{{{{{j}}}}}" for j in range(len(field)) for i in field]
                )  # \x02로 병합
                for i in range(len(field)):
                    text_str += "\x02".join([str(t) for t in field.iloc[i]]) + "\x02"
            return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(text, pd.DataFrame):
            if not isinstance(text.columns, pd.core.indexes.range.RangeIndex):
                text = text.T
            text_str = ""
            if isinstance(idx, int):
                field_str = "\x02".join(
                    [i + f"{{{{{idx}}}}}" for i in field.split("\x02")]
                )  # \x02로 병합
                text_str += "\x02".join([str(t) for t in text[idx]]) + "\x02"
            else:
                field_str = "\x02".join(
                    [
                        str(i) + f"{{{{{j}}}}}"
                        for i in field.split("\x02")
                        for j in range(len(text.columns))
                    ]
                )  # \x02로 병합
                for i in range(len(text)):
                    text_str += "\x02".join([str(t) for t in text.iloc[i]]) + "\x02"
            return self.hwp.PutFieldText(Field=field_str, Text=text_str)

        if isinstance(idx, int):
            return self.hwp.PutFieldText(
                Field=field.replace("\x02", f"{{{{{idx}}}}}\x02") + f"{{{{{idx}}}}}",
                Text=text,
            )
        else:
            return self.hwp.PutFieldText(Field=field, Text=text)

    def PutFieldText(
        self, field: Any = "", text: Union[str, list, tuple, pd.Series] = "", idx=None
    ) -> None:
        return self.put_field_text(field, text, idx)

    def put_metatag_name_text(self, tag: str, text: str):
        """메타태그에 텍스트 삽입"""
        return self.hwp.PutMetatagNameText(tag=tag, Text=text)

    def PutMetatagNameText(self, tag: str, text: str):
        return self.put_metatag_name_text(tag, text)

    def quit(self, save: bool = False) -> None:
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
        elif save:  # 빈 문서가 아닌 경우
            self.save()
        else:
            self.clear()
        self.hwp.Quit()

    def Quit(self, save: bool = False) -> None:
        return self.quit(save)

    def rgb_color(
        self, red_or_colorname: Union[str, tuple, int], green: int = 255, blue: int = 255
    ) -> int:
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
        color_palette = {
            "Red": (255, 0, 0),
            "Green": (0, 255, 0),
            "Blue": (0, 0, 255),
            "Yellow": (255, 255, 0),
            "Cyan": (0, 255, 255),
            "Magenta": (255, 0, 255),
            "Black": (0, 0, 0),
            "White": (255, 255, 255),
            "Gray": (128, 128, 128),
            "Orange": (255, 165, 0),
            "DarkBlue": (0, 0, 139),
            "Purple": (128, 0, 128),
            "Pink": (255, 192, 203),
            "Lime": (0, 255, 0),
            "SkyBlue": (135, 206, 235),
            "Gold": (255, 215, 0),
            "Silver": (192, 192, 192),
            "Mint": (189, 252, 201),
            "Tomato": (255, 99, 71),
            "Olive": (128, 128, 0),
            "Crimson": (220, 20, 60),
            "Navy": (0, 0, 128),
            "Teal": (0, 128, 128),
            "Chocolate": (210, 105, 30),
        }
        if red_or_colorname in color_palette:
            return self.hwp.RGBColor(*color_palette[red_or_colorname])
        return self.hwp.RGBColor(red=red_or_colorname, green=green, blue=blue)

    def RGBColor(
        self, red_or_colorname: Union[str, tuple, int], green: int = 255, blue: int = 255
    ) -> int:
        return self.rgb_color(red_or_colorname, green, blue)

    def register_module(
        self,
        module_type: str = "FilePathCheckDLL",
        module_data: str = "FilePathCheckerModule",
    ) -> bool:
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

    def RegisterModule(
        self,
        module_type: str = "FilePathCheckDLL",
        module_data: str = "FilePathCheckerModule",
    ) -> bool:
        return self.register_module(module_type, module_data)

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
        from winreg import (
            ConnectRegistry,
            HKEY_CURRENT_USER,
            OpenKey,
            KEY_WRITE,
            SetValueEx,
            REG_SZ,
            CloseKey,
        )

        try:
            # pyhwpx가 설치된 파이썬 환경 또는 pyinstaller로 컴파일한 환경에서 pyhwpx 경로 찾기
            # 살펴본 결과, FilePathCheckerModule.dll 파일은 pyinstaller 컴파일시 자동포함되지는 않는 것으로 확인..
            location = [
                i.split(": ")[1]
                for i in subprocess.check_output(
                    ["pip", "show", "pyhwpx"], stderr=subprocess.DEVNULL
                )
                .decode(encoding="cp949")
                .split("\r\n")
                if i.startswith("Location: ")
            ][0]
            location = os.path.join(location, "pyhwpx")
        except UnicodeDecodeError:
            location = [
                i.split(": ")[1]
                for i in subprocess.check_output(
                    ["pip", "show", "pyhwpx"], stderr=subprocess.DEVNULL
                )
                .decode()
                .split("\r\n")
                if i.startswith("Location: ")
            ][0]
            location = os.path.join(location, "pyhwpx")
        except Exception:
            pass
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
            elif dll_name.lower in [
                i.lower() for i in os.listdir(os.path.join(os.environ["USERPROFILE"]))
            ]:
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

                print(
                    "https://github.com/hancom-io에서 보안모듈 다운로드를 시도합니다."
                )
                try:
                    f = request.urlretrieve(
                        "https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/%EB%B3%B4%EC%95%88%EB%AA%A8%EB%93%88(Automation).zip",
                        filename=os.path.join(
                            os.environ["USERPROFILE"], "FilePathCheckerModule.zip"
                        ),
                    )
                    with ZipFile(f[0]) as zf:
                        zf.extract(
                            "FilePathCheckerModuleExample.dll",
                            os.path.join(os.environ["USERPROFILE"]),
                        )
                    os.remove(
                        os.path.join(
                            os.environ["USERPROFILE"], "FilePathCheckerModule.zip"
                        )
                    )
                    if not os.path.exists(
                        os.path.join(
                            os.environ["USERPROFILE"], "FilePathCheckerModule.dll"
                        )
                    ):
                        os.rename(
                            os.path.join(
                                os.environ["USERPROFILE"],
                                "FilePathCheckerModuleExample.dll",
                            ),
                            os.path.join(os.environ["USERPROFILE"], dll_name),
                        )
                    location = os.environ["USERPROFILE"]
                    print("사용자폴더", location, "에 보안모듈을 설치하였습니다.")
                except urllib.error.URLError as e:
                    # URLError를 처리합니다.
                    print(
                        f"내부망에서는 보안모듈을 다운로드할 수 없습니다. 보안모듈을 직접 다운받아 설치하여 주시기 바랍니다.: \n{e.reason}"
                    )
                except Exception as e:
                    # 기타 예외를 처리합니다.
                    print(
                        f"예기치 못한 오류가 발생했습니다. 아래 오류를 개발자에게 문의해주시기 바랍니다: \n{str(e)}"
                    )
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
        SetValueEx(
            key, "FilePathCheckerModule", 0, REG_SZ, os.path.join(location, dll_name)
        )
        CloseKey(key)

    def register_private_info_pattern(
        self, private_type: int, private_pattern: str
    ) -> bool:
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
        return self.hwp.RegisterPrivateInfoPattern(
            PrivateType=private_type, PrivatePattern=private_pattern
        )

    def RegisterPrivateInfoPattern(
        self, private_type: int, private_pattern: int
    ) -> bool:
        return self.register_private_info_pattern(private_type, private_pattern)

    def release_action(self, action: str):
        return self.hwp.ReleaseAction(action=action)

    def ReleaseAction(self, action: str):
        return self.release_action(action)

    def release_scan(self) -> None:
        """
        InitScan()으로 설정된 초기화 정보를 해제한다.

        텍스트 검색작업이 끝나면 반드시 호출하여 설정된 정보를 해제해야 한다.

        Returns:
            None
        """
        return self.hwp.ReleaseScan()

    def ReleaseScan(self) -> None:
        return self.release_scan()

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

    def RenameField(self, oldname: str, newname: str) -> bool:
        return self.rename_field(oldname, newname)

    def rename_metatag(self, oldtag, newtag):
        """메타태그 이름 변경"""
        return self.hwp.RenameMetatag(oldtag=oldtag, newtag=newtag)

    def RenameMetatag(self, oldtag, newtag):
        return self.rename_metatag(oldtag, newtag)

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

        return self.hwp.ReplaceAction(
            OldActionID=old_action_id, NewActionID=new_action_id
        )

    def ReplaceAction(self, old_action_id: str, new_action_id: str) -> bool:
        return self.replace_action(old_action_id, new_action_id)

    def replace_font(
        self, langid, des_font_name, des_font_type, new_font_name, new_font_type
    ):
        return self.hwp.ReplaceFont(
            langid=langid,
            desFontName=des_font_name,
            desFontType=des_font_type,
            newFontName=new_font_name,
            newFontType=new_font_type,
        )

    def ReplaceFont(
        self, langid, des_font_name, des_font_type, new_font_name, new_font_type
    ):
        return self.replace_font(
            langid, des_font_name, des_font_type, new_font_name, new_font_type
        )

    def revision(self, revision):
        return self.hwp.Revision(Revision=revision)

    def Revision(self, revision):
        return self.revision(revision)

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

    def compose_chars(
        self,
        Chars: Union[str, int] = "",
        CharSize: int = -3,
        CheckCompose: int = 0,
        CircleType: int = 0,
        **kwargs,
    ) -> bool:
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

    def ComposeChars(
        self,
        Chars: Union[str, int] = "",
        CharSize: int = -3,
        CheckCompose: int = 0,
        CircleType: int = 0,
        **kwargs,
    ) -> bool:
        return self.compose_chars(Chars, CharSize, CheckCompose, CircleType, **kwargs)

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

    def run_script_macro(
        self, function_name: str, u_macro_type: int = 0, u_script_type: int = 0
    ) -> bool:
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
        return self.hwp.RunScriptMacro(
            FunctionName=function_name,
            uMacroType=u_macro_type,
            uScriptType=u_script_type,
        )

    def RunScriptMacro(
        self, function_name: str, u_macro_type: int = 0, u_script_type: int = 0
    ) -> bool:
        return self.run_script_macro(function_name, u_macro_type, u_script_type)

    def save(self, save_if_dirty: bool = True) -> bool:
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

    def Save(self, save_if_dirty: bool = True) -> bool:
        return self.save(save_if_dirty)

    def save_as(
        self, path: str, format: str = "HWP", arg: str = "", split_page: bool = False
    ) -> bool:
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
                _fields_ = [
                    ("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", PUL),
                ]

            class HardwareInput(ctypes.Structure):
                _fields_ = [
                    ("uMsg", ctypes.c_ulong),
                    ("wParamL", ctypes.c_short),
                    ("wParamH", ctypes.c_ushort),
                ]

            class MouseInput(ctypes.Structure):
                _fields_ = [
                    ("dx", ctypes.c_long),
                    ("dy", ctypes.c_long),
                    ("mouseData", ctypes.c_ulong),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", PUL),
                ]

            class Input_I(ctypes.Union):
                _fields_ = [
                    ("ki", KeyBdInput),
                    ("mi", MouseInput),
                    ("hi", HardwareInput),
                ]

            class Input(ctypes.Structure):
                _fields_ = [("type", ctypes.c_ulong), ("ii", Input_I)]

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

    def SaveAs(self, path: str, format: str = "HWP", arg: str = "") -> bool:
        return self.save_as(path, format, arg)

    def scan_font(self):
        return self.hwp.ScanFont()

    def ScanFont(self):
        return self.scan_font()

    def select_text_by_get_pos(self, s_getpos: tuple, e_getpos: tuple) -> bool:
        """
        hwp.get_pos()로 얻은 두 튜플 사이의 텍스트를 선택하는 메서드.

        Args:
            s_getpos: 선택 시작점 좌표의 get_pos값
            e_getpos: 선택 끝점 좌표의 get_pos값

        Returns:
            선택 성공시 True, 실패시 False를 리턴함

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> s_pos = hwp.get_pos()  # 선택 시작부분 저장
            >>> # 이래저래 좌표 이동 후
            >>> e_pos = hwp.get_pos()  # 선택 끝부분 저장
            >>> # 이래저래 또 좌표 이동 후
            >>> hwp.select_text_by_get_pos(s_pos, e_pos)
            True
        """
        self.set_pos(s_getpos[0], 0, 0)
        return self.hwp.SelectText(
            spara=s_getpos[1], spos=s_getpos[2], epara=e_getpos[1], epos=e_getpos[2]
        )

    def select_text(
        self,
        spara: Union[int, list, tuple] = 0,
        spos: int = 0,
        epara: int = 0,
        epos: int = 0,
        slist: int = 0,
    ) -> bool:
        """
        특정 범위의 텍스트를 블록선택한다.
        (epos가 가리키는 문자는 포함되지 않는다.)
        hwp.get_selected_pos()를 통해 저장한 위치로 돌아가는 데에도 사용된다.

        Args:
            spara: 블록 시작 위치의 문단 번호. (또는 hwp.get_selected_pos() 리턴값)
            spos: 블록 시작 위치의 문단 중에서 문자의 위치.
            epara: 블록 끝 위치의 문단 번호.
            epos: 블록 끝 위치의 문단 중에서 문자의 위치.

        Returns:
            성공하면 True, 실패하면 False

        Examples:
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp()
            >>> # 본문의 세 번째 문단 전체 선택하기(기본 사용법)
            >>> hwp.select_text(2, 0, 2, -1, 0)
            True
            >>> # 임의의 영역으로 이동 후 저장한 위치로 되돌아가기
            >>> selected_range = hwp.get_selected_pos()
            >>> hwp.select_text(selected_range)
            True
        """
        if type(spara) in [list, tuple]:
            _, slist, spara, spos, elist, epara, epos = spara
        self.set_pos(slist, 0, 0)
        if epos == -1:
            self.hwp.SelectText(spara=spara, spos=spos, epara=epara, epos=0)
            return self.MoveSelParaEnd()
        else:
            return self.hwp.SelectText(spara=spara, spos=spos, epara=epara, epos=epos)

    def SelectText(
        self,
        spara: Union[int, list, tuple] = 0,
        spos: int = 0,
        epara: int = 0,
        epos: int = 0,
        slist: int = 0,
    ) -> bool:
        return self.select_text(spara, spos, epara, epos, slist)

    def set_cur_field_name(
        self, field: str = "", direction: str = "", memo: str = "", option: int = 0
    ) -> bool:
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
            return self.hwp.SetCurFieldName(
                Field=field, option=option, Direction=direction, memo=memo
            )

    def SetCurFieldName(
        self, field: str = "", direction: str = "", memo: str = "", option: int = 0
    ) -> bool:
        return self.set_cur_field_name(field, direction, memo, option)

    def set_field_view_option(self, option: int) -> int:
        # """
        # 양식모드와 읽기전용모드일 때 현재 열린 문서의 필드의 겉보기 속성(『』표시)을 바꾼다.
        #
        # EditMode와 비슷하게 현재 열려있는 문서에 대한 속성이다. 따라서 저장되지 않는다.
        # (작동하지 않음)
        #
        # Args:
        #     option:
        #         겉보기 속성 bit
        #
        #         - 1: 누름틀의 『』을 표시하지 않음, 기타필드의 『』을 표시하지 않음
        #         - 2: 누름틀의 『』을 빨간색으로 표시, 기타필드의 『』을 흰색으로 표시(기본값)
        #         - 3: 누름틀의 『』을 흰색으로 표시, 기타필드의 『』을 흰색으로 표시
        #
        # Returns:
        #     설정된 속성이 반환된다. 에러일 경우 0이 반환된다.
        # """
        return self.hwp.SetFieldViewOption(option=option)

    def SetFieldViewOption(self, option: int) -> bool:
        return self.set_field_view_option(option)

    def set_message_box_mode(self, mode: int) -> int:
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

    def SetMessageBoxMode(self, mode: int) -> int:
        return self.set_message_box_mode(mode)

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
        return self.set_pos(List, para, pos)

    def set_pos_by_set(self, disp_val: Any) -> bool:
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

    def SetPosBySet(self, disp_val: "Hwp.HParameterSet") -> bool:
        return self.set_pos_by_set(disp_val)

    def set_private_info_password(self, password: str) -> bool:
        # """
        # 개인정보보호를 위한 암호를 등록한다.
        #
        # 개인정보 보호를 설정하기 위해서는
        # 우선 개인정보 보호 암호를 먼저 설정해야 한다.
        # 그러므로 개인정보 보호 함수를 실행하기 이전에
        # 반드시 이 함수를 호출해야 한다.
        # (현재 작동하지 않는다.)
        #
        # Args:
        #     password: 새 암호
        #
        # Returns:
        #     정상적으로 암호가 설정되면 True를 반환한다.
        #     암호설정에 실패하면 false를 반환한다. false를 반환하는 경우는 다음과 같다
        #     - 암호의 길이가 너무 짧거나 너무 길 때 (영문 5~44자, 한글 3~22자)
        #     - 암호가 이미 설정되었음. 또는 암호가 이미 설정된 문서임
        # """
        return self.hwp.SetPrivateInfoPassword(Password=password)

    def SetPrivateInfoPassword(self, password: str) -> bool:
        return self.set_private_info_password(password)

    def set_text_file(
        self,
        data: str,
        format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X",
        option: str = "insertfile",
    ) -> int:
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

    def SetTextFile(
        self,
        data: str,
        format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X",
        option: str = "insertfile",
    ) -> int:
        return self.set_text_file(data, format, option)

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
        return win32gui.GetWindowText(
            self.hwp.XHwpWindows.Active_XHwpWindow.WindowHandle
        )

    def set_title(self, title: str = "") -> bool:
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
            >>> hwp.set_title("😘")
            >>> hwp.get_title()
            😘 - 한글
        """
        return self.hwp.SetTitleName(Title=title)

    def SetTitle(self, title: str = "") -> bool:
        return self.set_title(title)
