import os
import re

import numpy as np
import pandas as pd
import pythoncom
import win32com.client as win32


class Hwp:
    def __init__(self):
        hwp = ""
        context = pythoncom.CreateBindCtx(0)

        # 현재 실행중인 프로세스를 가져옵니다.
        running_coms = pythoncom.GetRunningObjectTable()
        monikers = running_coms.EnumRunning()

        for moniker in monikers:
            name = moniker.GetDisplayName(context, moniker);
            # moniker의 DisplayName을 통해 한글을 가져옵니다
            # 한글의 경우 HwpObject.버전으로 각 버전별 실행 이름을 설정합니다.
            if name.startswith('!HwpObject.'):
                # 120은 한글 2022의 경우입니다.
                # 현재 moniker를 통해 ROT에서 한글의 object를 가져옵니다.
                obj = running_coms.GetObject(moniker)
                # 가져온 object를 Dispatch를 통해 사용할수 있는 객체로 변환시킵니다.
                hwp = win32.gencache.EnsureDispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                # 그이후는 오토메이션 api를 사용할수 있습니다
        if not hwp:
            hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
        hwp.XHwpWindows.Item(0).Visible = True

        self.Application = hwp.Application
        self.ArcType = hwp.ArcType
        self.AutoNumType = hwp.AutoNumType
        self.BorderShape = hwp.BorderShape
        self.BreakWordLatin = hwp.BreakWordLatin
        self.BrushType = hwp.BrushType
        self.CLSID = hwp.CLSID
        self.Canonical = hwp.Canonical
        self.CellApply = hwp.CellApply
        self.CellShape = hwp.CellShape
        self.CharShadowType = hwp.CharShadowType
        self.CharShape = hwp.CharShape
        self.CheckXObject = hwp.CheckXObject
        self.Clear = hwp.Clear
        self.ColDefType = hwp.ColDefType
        self.ColLayoutType = hwp.ColLayoutType
        self.ConvertPUAHangulToUnicode = hwp.ConvertPUAHangulToUnicode
        self.CreateAction = hwp.CreateAction
        self.CreateField = hwp.CreateField
        self.CreateID = hwp.CreateID
        self.CreateMode = hwp.CreateMode
        self.CreatePageImage = hwp.CreatePageImage
        self.CreateSet = hwp.CreateSet
        self.CrookedSlash = hwp.CrookedSlash
        self.CurFieldState = hwp.CurFieldState
        self.CurMetatagState = hwp.CurMetatagState
        self.CurSelectedCtrl = hwp.CurSelectedCtrl
        self.DSMark = hwp.DSMark
        self.DbfCodeType = hwp.DbfCodeType
        self.DeleteCtrl = hwp.DeleteCtrl
        self.Delimiter = hwp.Delimiter
        self.DrawAspect = hwp.DrawAspect
        self.DrawFillImage = hwp.DrawFillImage
        self.DrawShadowType = hwp.DrawShadowType
        self.EditMode = hwp.EditMode
        self.Encrypt = hwp.Encrypt
        self.EndSize = hwp.EndSize
        self.EndStyle = hwp.EndStyle
        self.EngineProperties = hwp.EngineProperties
        self.ExportStyle = hwp.ExportStyle
        self.FieldExist = hwp.FieldExist
        self.FileTranslate = hwp.FileTranslate
        self.FillAreaType = hwp.FillAreaType
        self.FindCtrl = hwp.FindCtrl
        self.FindDir = hwp.FindDir
        self.FindPrivateInfo = hwp.FindPrivateInfo
        self.FontType = hwp.FontType
        self.GetBinDataPath = hwp.GetBinDataPath
        self.GetCurFieldName = hwp.GetCurFieldName
        self.GetCurMetatagName = hwp.GetCurMetatagName
        self.GetFieldList = hwp.GetFieldList
        self.GetFieldText = hwp.GetFieldText
        self.GetFileInfo = hwp.GetFileInfo
        self.GetFontList = hwp.GetFontList
        self.GetHeadingString = hwp.GetHeadingString
        self.GetMessageBoxMode = hwp.GetMessageBoxMode
        self.GetMetatagList = hwp.GetMetatagList
        self.GetMetatagNameText = hwp.GetMetatagNameText
        self.GetMousePos = hwp.GetMousePos
        self.GetPageText = hwp.GetPageText
        self.GetPos = hwp.GetPos
        self.GetPosBySet = hwp.GetPosBySet
        self.GetScriptSource = hwp.GetScriptSource
        self.GetSelectedPos = hwp.GetSelectedPos
        self.GetSelectedPosBySet = hwp.GetSelectedPosBySet
        self.GetText = hwp.GetText
        self.GetTextFile = hwp.GetTextFile
        self.GetTranslateLangList = hwp.GetTranslateLangList
        self.GetUserInfo = hwp.GetUserInfo
        self.Gradation = hwp.Gradation
        self.GridMethod = hwp.GridMethod
        self.GridViewLine = hwp.GridViewLine
        self.GutterMethod = hwp.GutterMethod
        self.HAction = hwp.HAction
        self.HAlign = hwp.HAlign
        self.HParameterSet = hwp.HParameterSet
        self.Handler = hwp.Handler
        self.Hash = hwp.Hash
        self.HatchStyle = hwp.HatchStyle
        self.HeadCtrl = hwp.HeadCtrl
        self.HeadType = hwp.HeadType
        self.HeightRel = hwp.HeightRel
        self.Hiding = hwp.Hiding
        self.HorzRel = hwp.HorzRel
        self.HwpLineType = hwp.HwpLineType
        self.HwpLineWidth = hwp.HwpLineWidth
        self.HwpOutlineStyle = hwp.HwpOutlineStyle
        self.HwpOutlineType = hwp.HwpOutlineType
        self.HwpUnderlineShape = hwp.HwpUnderlineShape
        self.HwpUnderlineType = hwp.HwpUnderlineType
        self.HwpZoomType = hwp.HwpZoomType
        self.ImageFormat = hwp.ImageFormat
        self.ImportStyle = hwp.ImportStyle
        self.InitHParameterSet = hwp.InitHParameterSet
        self.InitScan = hwp.InitScan
        self.Insert = hwp.Insert
        self.InsertBackgroundPicture = hwp.InsertBackgroundPicture
        self.InsertCtrl = hwp.InsertCtrl
        self.InsertPicture = hwp.InsertPicture
        self.IsActionEnable = hwp.IsActionEnable
        self.IsCommandLock = hwp.IsCommandLock
        self.IsEmpty = hwp.IsEmpty
        self.IsModified = hwp.IsModified
        self.IsPrivateInfoProtected = hwp.IsPrivateInfoProtected
        self.IsTrackChange = hwp.IsTrackChange
        self.IsTrackChangePassword = hwp.IsTrackChangePassword
        self.KeyIndicator = hwp.KeyIndicator
        self.LastCtrl = hwp.LastCtrl
        self.LineSpacingMethod = hwp.LineSpacingMethod
        self.LineWrapType = hwp.LineWrapType
        self.LockCommand = hwp.LockCommand
        self.LunarToSolar = hwp.LunarToSolar
        self.LunarToSolarBySet = hwp.LunarToSolarBySet
        self.MacroState = hwp.MacroState
        self.MailType = hwp.MailType
        self.MetatagExist = hwp.MetatagExist
        self.MiliToHwpUnit = hwp.MiliToHwpUnit
        self.ModifyFieldProperties = hwp.ModifyFieldProperties
        self.ModifyMetatagProperties = hwp.ModifyMetatagProperties
        self.MovePos = hwp.MovePos
        self.MoveToField = hwp.MoveToField
        self.MoveToMetatag = hwp.MoveToMetatag
        self.NumberFormat = hwp.NumberFormat
        self.Numbering = hwp.Numbering
        self.Open = hwp.Open
        self.PageCount = hwp.PageCount
        self.PageNumPosition = hwp.PageNumPosition
        self.PageType = hwp.PageType
        self.ParaHeadAlign = hwp.ParaHeadAlign
        self.ParaShape = hwp.ParaShape
        self.ParentCtrl = hwp.ParentCtrl
        self.Path = hwp.Path
        self.PicEffect = hwp.PicEffect
        self.PlacementType = hwp.PlacementType
        self.PointToHwpUnit = hwp.PointToHwpUnit
        self.PresentEffect = hwp.PresentEffect
        self.PrintDevice = hwp.PrintDevice
        self.PrintPaper = hwp.PrintPaper
        self.PrintRange = hwp.PrintRange
        self.PrintType = hwp.PrintType
        self.ProtectPrivateInfo = hwp.ProtectPrivateInfo
        self.PutFieldText = hwp.PutFieldText
        self.PutMetatagNameText = hwp.PutMetatagNameText
        self.Quit = hwp.Quit
        self.RGBColor = hwp.RGBColor
        self.RegisterModule = hwp.RegisterModule
        self.RegisterPrivateInfoPattern = hwp.RegisterPrivateInfoPattern
        self.ReleaseAction = hwp.ReleaseAction
        self.ReleaseScan = hwp.ReleaseScan
        self.RenameField = hwp.RenameField
        self.RenameMetatag = hwp.RenameMetatag
        self.ReplaceAction = hwp.ReplaceAction
        self.ReplaceFont = hwp.ReplaceFont
        self.Revision = hwp.Revision
        self.Run = hwp.Run
        self.RunScriptMacro = hwp.RunScriptMacro
        self.Save = hwp.Save
        self.SaveAs = hwp.SaveAs
        self.ScanFont = hwp.ScanFont
        self.SelectText = hwp.SelectText
        self.SelectionMode = hwp.SelectionMode
        self.SetBarCodeImage = hwp.SetBarCodeImage
        self.SetCurFieldName = hwp.SetCurFieldName
        self.SetCurMetatagName = hwp.SetCurMetatagName
        self.SetDRMAuthority = hwp.SetDRMAuthority
        self.SetFieldViewOption = hwp.SetFieldViewOption
        self.SetMessageBoxMode = hwp.SetMessageBoxMode
        self.SetPos = hwp.SetPos
        self.SetPosBySet = hwp.SetPosBySet
        self.SetPrivateInfoPassword = hwp.SetPrivateInfoPassword
        self.SetTextFile = hwp.SetTextFile
        self.SetTitleName = hwp.SetTitleName
        self.SetUserInfo = hwp.SetUserInfo
        self.SideType = hwp.SideType
        self.Signature = hwp.Signature
        self.Slash = hwp.Slash
        self.SolarToLunar = hwp.SolarToLunar
        self.SolarToLunarBySet = hwp.SolarToLunarBySet
        self.SortDelimiter = hwp.SortDelimiter
        self.StrikeOut = hwp.StrikeOut
        self.StyleType = hwp.StyleType
        self.SubtPos = hwp.SubtPos
        self.TableBreak = hwp.TableBreak
        self.TableFormat = hwp.TableFormat
        self.TableSwapType = hwp.TableSwapType
        self.TableTarget = hwp.TableTarget
        self.TextAlign = hwp.TextAlign
        self.TextArtAlign = hwp.TextArtAlign
        self.TextDir = hwp.TextDir
        self.TextFlowType = hwp.TextFlowType
        self.TextWrapType = hwp.TextWrapType
        self.UnSelectCtrl = hwp.UnSelectCtrl
        self.VAlign = hwp.VAlign
        self.Version = hwp.Version
        self.VertRel = hwp.VertRel
        self.ViewFlag = hwp.ViewFlag
        self.ViewProperties = hwp.ViewProperties
        self.WatermarkBrush = hwp.WatermarkBrush
        self.WidthRel = hwp.WidthRel
        self.XHwpDocuments = hwp.XHwpDocuments
        self.XHwpMessageBox = hwp.XHwpMessageBox
        self.XHwpODBC = hwp.XHwpODBC
        self.XHwpWindows = hwp.XHwpWindows

    def get_sel_text(self):
        self.InitScan(Range=0xff)
        total_text = ""
        state = 2
        while state not in [0, 1]:
            state, text = self.GetText()
            total_text += text
        self.ReleaseScan()
        return total_text

    def table_to_csv(self, idx=1, filename="result.csv"):
        start_pos = self.GetPos()
        table_num = 0
        ctrl = self.HeadCtrl
        while ctrl.Next:
            if ctrl.UserDesc == "표":
                table_num += 1
            if table_num == idx:
                break
            ctrl = ctrl.Next

        self.SetPosBySet(ctrl.GetAnchorPos(0))
        self.FindCtrl()
        self.HAction.Run("ShapeObjTableSelCell")
        data = list(self.get_sel_text())
        col_count = 1
        while self.HAction.Run("TableRightCell"):
            # a.append(get_text().replace("\r\n", "\n"))
            if re.match("\([A-Z]1\)", self.KeyIndicator()[-1]):
                col_count += 1
            data.append(self.get_sel_text())

        array = np.array(data).reshape(col_count, -1)
        df = pd.DataFrame(array[1:], columns=array[0])
        df.to_csv(filename, index=False)
        self.SetPos(*start_pos)
        print(os.path.join(os.getcwd(), filename))
        return None

    def table_to_df(self, idx=1):
        start_pos = self.GetPos()
        table_num = 0
        ctrl = self.HeadCtrl
        while ctrl.Next:
            if ctrl.UserDesc == "표":
                table_num += 1
            if table_num == idx:
                break
            ctrl = ctrl.Next

        self.SetPosBySet(ctrl.GetAnchorPos(0))
        self.FindCtrl()
        self.HAction.Run("ShapeObjTableSelCell")
        data = list(self.get_sel_text())
        col_count = 1
        while self.HAction.Run("TableRightCell"):
            # a.append(get_text().replace("\r\n", "\n"))
            if re.match("\([A-Z]1\)", self.KeyIndicator()[-1]):
                col_count += 1
            data.append(self.get_sel_text())

        array = np.array(data).reshape(col_count, -1)
        df = pd.DataFrame(array[1:], columns=array[0])
        self.SetPos(*start_pos)
        return df

    def insert_text(self, text):
        param = self.HParameterSet.HInsertText
        self.HAction.GetDefault("InsertText", param.HSet)
        param.Text = text
        self.HAction.Execute("InsertText", param.HSet)

    def move_caption(self, location="Bottom"):
        start_pos = self.GetPos()
        ctrl = self.HeadCtrl
        while ctrl:
            if ctrl.UserDesc == "번호 넣기":
                self.SetPosBySet(ctrl.GetAnchorPos(0))
                self.HAction.Run("ParagraphShapeAlignCenter")
                param = self.HParameterSet.HShapeObject
                self.HAction.GetDefault("TablePropertyDialog", param.HSet)
                param.ShapeCaption.Side = self.SideType(location)
                self.HAction.Execute("TablePropertyDialog", param.HSet)
            ctrl = ctrl.Next
        self.SetPos(*start_pos)
        return None

    def is_empty(self) -> bool:
        """
        아무 내용도 들어있지 않은 빈 문서인지 여부를 나타낸다. 읽기전용
        """
        return self.IsEmpty

    def arc_type(self, arc_type):
        return self.ArcType(ArcType=arc_type)

    def auto_num_type(self, autonum):
        return self.AutoNumType(autonum=autonum)

    def border_shape(self, border_type):
        return self.BorderShape(BorderType=border_type)

    def break_word_latin(self, break_latin_word):
        return self.BreakWordLatin(BreakLatinWord=break_latin_word)

    def brush_type(self, brush_type):
        return self.BrushType(BrushType=brush_type)

    def canonical(self, canonical):
        return self.Canonical(Canonical=canonical)

    def cell_apply(self, cell_apply):
        return self.CellApply(CellApply=cell_apply)

    def char_shadow_type(self, shadow_type):
        return self.CharShadowType(ShadowType=shadow_type)

    def check_xobject(self, bstring):
        return self.CheckXObject(bstring=bstring)

    def clear(self, option: int = 1):
        """
        현재 편집중인 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다.

        :param option:
            편집중인 문서의 내용에 대한 처리 방법, 생략하면 1(hwpDiscard)가 선택된다.
            0: 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다. (hwpAskSave)
            1: 문서의 내용을 버린다. (hwpDiscard)
            2: 문서가 변경된 경우 저장한다. (hwpSaveIfDirty)
            3: 무조건 저장한다. (hwpSave)

        :return:
            None

        :examples:
            >>> hwp.clear(1)
            True

        """
        return self.Clear(option=option)

    def col_def_type(self, col_def_type):
        return self.ColDefType(ColDefType=col_def_type)

    def col_layout_type(self, col_layout_type):
        return self.ColLayoutType(ColLayoutType=col_layout_type)

    def convert_pua_hangul_to_unicode(self, reverse):
        return self.ConvertPUAHangulToUnicode(Reverse=reverse)

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

        :examples:
            >>> # 현재 커서의 폰트 크기(Height)를 구하는 코드
            >>> act = hwp.CreateAction("CharShape")
            >>> cs = act.CreateSet()  # == cs = hwp.CreateSet(act)
            >>> act.GetDefault(cs)
            >>> print(cs.Item("Height"))
            2800

            >>> # 현재 선택범위의 폰트 크기를 20pt로 변경하는 코드
            >>> act = hwp.CreateAction("CharShape")
            >>> cs = act.CreateSet()  # == cs = hwp.CreateSet(act)
            >>> act.GetDefault(cs)
            >>> cs.SetItem("Height", hwp.PointToHwpUnit(20))
            >>> act.Execute(cs)
            True

        """
        return self.CreateAction(actidstr=actidstr)

    def create_field(self, direction: str, memo: str, name: str) -> bool:
        """
        캐럿의 현재 위치에 누름틀을 생성한다.

        :param direction:
            누름틀에 입력이 안 된 상태에서 보이는 안내문/지시문.

        :param memo:
            누름틀에 대한 설명/도움말

        :param name:
            누름틀 필드에 대한 필드 이름(중요)

        :return:
            성공이면 True, 실패면 False

        :examples:
            >>> hwp.create_field(direction="이름", memo="이름을 입력하는 필드", name="name")
            True
            >>> hwp.PutFieldText("name", "일코")
        """
        return self.CreateField(Direction=direction, memo=memo, name=name)

    def create_id(self, creation_id):
        return self.CreateID(CreationID=creation_id)

    def create_mode(self, creation_mode):
        return self.CreateMode(CreationMode=creation_mode)

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

        examples:
            >>> hwp.create_page_image("c:/Users/User/Desktop/a.bmp")
            True
        """
        return self.CreatePageImage(Path=path, pgno=pgno, resolution=resolution, depth=depth, Format=format)

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
        return self.CreateSet(setidstr=setidstr)

    def crooked_slash(self, crooked_slash):
        return self.CrookedSlash(CrookedSlash=crooked_slash)

    def ds_mark(self, diac_sym_mark):
        return self.DSMark(DiacSymMark=diac_sym_mark)

    def dbf_code_type(self, dbf_code):
        return self.DbfCodeType(DbfCode=dbf_code)

    def delete_ctrl(self, ctrl) -> bool:
        """
        문서 내 컨트롤을 삭제한다.

        :param ctrl:
            삭제할 문서 내 컨트롤

        :return:
            성공하면 True, 실패하면 False

        examples:
            >>> ctrl = hwp.HeadCtrl.Next.Next
            >>> if ctrl.UserDesc == "표":
            ...     hwp.delete_ctrl(ctrl)
            ...
            True
        """
        return self.DeleteCtrl(ctrl=ctrl)

    def delimiter(self, delimiter):
        return self.Delimiter(Delimiter=delimiter)

    def draw_aspect(self, draw_aspect):
        return self.DrawAspect(DrawAspect=draw_aspect)

    def draw_fill_image(self, fillimage):
        return self.DrawFillImage(fillimage=fillimage)

    def draw_shadow_type(self, shadow_type):
        return self.DrawShadowType(ShadowType=shadow_type)

    def encrypt(self, encrypt):
        return self.Encrypt(Encrypt=encrypt)

    def end_size(self, end_size):
        return self.EndSize(EndSize=end_size)

    def end_style(self, end_style):
        return self.EndStyle(EndStyle=end_style)

    def export_style(self, sty_filepath: str) -> bool:
        """
        현재 문서의 Style을 sty 파일로 Export한다.

        :param sty_filepath:
            Export할 sty 파일의 전체경로 문자열

        :return:
            성공시 True, 실패시 False

        :Examples
            >>> hwp.export_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        style_set = self.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.ExportStyle(param=style_set.HSet)

    def field_exist(self, field):
        """
        문서에 지정된 데이터 필드가 존재하는지 검사한다.

        :param field:
            필드이름

        :return:
            필드가 존재하면 True, 존재하지 않으면 False
        """
        return self.FieldExist(Field=field)

    def file_translate(self, cur_lang, trans_lang):
        return self.FileTranslate(curLang=cur_lang, transLang=trans_lang)

    def fill_area_type(self, fill_area):
        return self.FillAreaType(FillArea=fill_area)

    def find_ctrl(self):
        return self.FindCtrl()

    def find_dir(self, find_dir):
        return self.FindDir(FindDir=find_dir)

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
        return self.FindPrivateInfo(PrivateType=private_type, PrivateString=private_string)

    def font_type(self, font_type):
        return self.FontType(FontType=font_type)

    def get_bin_data_path(self, binid):
        """
        Binary Data(Temp Image 등)의 경로를 가져온다.

        :param binid:
            바이너리 데이터의 ID 값 (1부터 시작)

        :return:
            바이너리 데이터의 경로

        Examples:
            >>> path = hwp.GetBinDataPath(2)
            >>> print(path)
            C:/Users/User/AppData/Local/Temp/Hnc/BinData/EMB00004dd86171.jpg
        """
        return self.GetBinDataPath(binid=binid)

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
        return self.GetCurFieldName(option=option)

    def get_cur_metatag_name(self):
        return self.GetCurMetatagName()

    def get_field_list(self, number=0, option=0):
        """
        문서에 존재하는 필드의 목록을 구한다.
        문서 중에 동일한 이름의 필드가 여러 개 존재할 때는
        number에 지정한 타입에 따라 3 가지의 서로 다른 방식 중에서 선택할 수 있다.
        예를 들어 문서 중 title, body, title, body, footer 순으로
        5개의 필드가 존재할 때, hwpFieldPlain, hwpFieldNumber, HwpFieldCount
        세 가지 형식에 따라 다음과 같은 내용이 돌아온다.
        hwpFieldPlain: "title\x2body\x2title\x2body\x2footer"
        hwpFieldNumber: "title{{0}}\x2body{{0}}\x2title{{1}}\x2body{{1}}\x2footer{{0}}"
        hwpFieldCount: "title{{2}}\x2body{{2}}\x2footer{{1}}"

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
        return self.GetFieldList(Number=number, option=option)

    def get_field_text(self, field):
        return self.GetFieldText(Field=field)

    def get_file_info(self, filename):
        return self.GetFileInfo(filename=filename)

    def get_font_list(self, langid):
        return self.GetFontList(langid=langid)

    def get_heading_string(self):
        return self.GetHeadingString()

    def get_message_box_mode(self):
        return self.GetMessageBoxMode()

    def get_metatag_list(self, number, option):
        return self.GetMetatagList(Number=number, option=option)

    def get_metatag_name_text(self, tag):
        return self.GetMetatagNameText(tag=tag)

    def get_mouse_pos(self, x_rel_to, y_rel_to):
        pass

    def get_page_text(self, pgno, option):
        pass

    def get_pos(self, list, para, pos):
        pass

    def get_pos_by_set(self):
        pass

    def get_script_source(self, filename):
        pass

    def get_selected_pos(self, slist, spara, spos, elist, epara, epos):
        pass

    def get_selected_pos_by_set(self, sset, eset):
        pass

    def get_text(self, text):
        pass

    def get_text_file(self, format, option):
        pass

    def get_translate_lang_list(self, cur_lang):
        pass

    def get_user_info(self, user_info_id):
        pass

    def gradation(self, gradation):
        pass

    def grid_method(self, grid_method):
        pass

    def grid_view_line(self, grid_view_line):
        pass

    def gutter_method(self, gutter_type):
        pass

    def h_align(self, h_align):
        pass

    def handler(self, handler):
        pass

    def hash(self, hash):
        pass

    def hatch_style(self, hatch_style):
        pass

    def head_type(self, heading_type):
        pass

    def height_rel(self, height_rel):
        pass

    def hiding(self, hiding):
        pass

    def horz_rel(self, horz_rel):
        pass

    def hwp_line_type(self, line_type):
        pass

    def hwp_line_width(self, line_width):
        pass

    def hwp_outline_style(self, hwp_outline_style):
        pass

    def hwp_outline_type(self, hwp_outline_type):
        pass

    def hwp_underline_shape(self, hwp_underline_shape):
        pass

    def hwp_underline_type(self, hwp_underline_type):
        pass

    def hwp_zoom_type(self, zoom_type):
        pass

    def image_format(self, image_format):
        pass

    def import_style(self, sty_filepath):
        """
        미리 저장된 특정 sty파일의 스타일을 임포트한다.

        :param sty_filepath:
            sty파일의 경로

        :return:
            성공시 True, 실패시 False

        :Examples
            >>> hwp.import_style("C:/Users/User/Desktop/new_style.sty")
            True
        """
        style_set = self.HParameterSet.HStyleTemplate
        style_set.filename = sty_filepath
        return self.ImportStyle(style_set.HSet)

    def init_hparameter_set(self):
        pass

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

        Examples:
            >>> hwp.init_scan(range=0xff)
            >>> _, text = hwp.get_text()
            >>> hwp.release_scan()
            >>> print(text)
            Hello, world!
        """
        return self.InitScan(option=option, Range=range, spara=spara,
                             spos=spos, epara=epara, epos=epos)

    def insert(self, path, format="", arg=""):
        """
        현재 캐럿 위치에 문서파일을 삽입한다.
        format, arg에 대해서는 hwp.open 참조

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
        return self.Insert(Path=path, Format=format, arg=arg)

    def insert_background_picture(self, path, border_type="SelectedCell",
                                  embedded=True, filloption=5, effect=1,
                                  watermark=False, brightness=0, contrast=0) -> bool:
        """
        셀에 배경이미지를 삽입한다.
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

        Examples:
            >>> hwp.insert_background_picture(path="C:/Users/User/Desktop/KakaoTalk_20230709_023118549.jpg")
            True
        """
        return self.InsertBackgroundPicture(Path=path, BorderType=border_type,
                                            Embedded=embedded, filloption=filloption,
                                            Effect=effect, watermark=watermark,
                                            Brightness=brightness, Contrast=contrast)

    def insert_ctrl(self, ctrl_id, initparam):
        pass

    def insert_picture(self, path, embedded, sizeoption, reverse, watermark, effect, width, height):
        pass

    def is_action_enable(self, action_id):
        pass

    def is_command_lock(self, action_id):
        pass

    def key_indicator(self, seccnt, secno, prnpageno, colno, line, pos, over, ctrlname):
        pass

    def line_spacing_method(self, line_spacing):
        pass

    def line_wrap_type(self, line_wrap):
        pass

    def lock_command(self, act_id, is_lock):
        pass

    def lunar_to_solar(self, l_year, l_month, l_day, l_leap, s_year, s_month, s_day):
        pass

    def lunar_to_solar_by_set(self, l_year, l_month, l_day, l_leap):
        pass

    def macro_state(self, macro_state):
        pass

    def mail_type(self, mail_type):
        pass

    def metatag_exist(self, tag):
        pass

    def mili_to_hwp_unit(self, mili):
        pass

    def modify_field_properties(self, field, remove, add):
        pass

    def modify_metatag_properties(self, tag, remove, add):
        pass

    def move_pos(self, move_id, para, pos):
        pass

    def move_to_field(self, field, text, start, select):
        pass

    def move_to_metatag(self, tag, text, start, select):
        pass

    def number_format(self, num_format):
        pass

    def numbering(self, numbering):
        pass

    def open(self, filename, format, arg):
        pass

    def page_num_position(self, pagenumpos):
        pass

    def page_type(self, page_type):
        pass

    def para_head_align(self, para_head_align):
        pass

    def pic_effect(self, pic_effect):
        pass

    def placement_type(self, restart):
        pass

    def point_to_hwp_unit(self, point):
        pass

    def present_effect(self, prsnteffect):
        pass

    def print_device(self, print_device):
        pass

    def print_paper(self, print_paper):
        pass

    def print_range(self, print_range):
        pass

    def print_type(self, print_method):
        pass

    def protect_private_info(self, potecting_char, private_pattern_type):
        pass

    def put_field_text(self, field, text):
        pass

    def put_metatag_name_text(self, tag, text):
        pass

    def rgb_color(self, red, green, blue):
        pass

    def register_module(self, module_type, module_data):
        pass

    def register_private_info_pattern(self, private_type, private_pattern):
        pass

    def release_action(self, action):
        pass

    def release_scan(self):
        pass

    def rename_field(self, oldname, newname):
        pass

    def rename_metatag(self, oldtag, newtag):
        pass

    def replace_font(self, langid, des_font_name, des_font_type, new_font_name, new_font_type):
        pass

    def revision(self, revision):
        pass

    def run(self, act_id):
        pass

    def run_script_macro(self, function_name, u_macro_type, u_script_type):
        pass

    def save(self, save_if_dirty):
        pass

    def save_as(self, path, format, arg):
        pass

    def scan_font(self):
        pass

    def select_text(self, spara, spos, epara, epos):
        pass

    def set_bar_code_image(self, lp_image_path, pgno, index, x, y, width, height):
        pass

    def set_cur_field_name(self, field, option, direction, memo):
        pass

    def set_cur_metatag_name(self, tag):
        pass

    def set_drm_authority(self, authority):
        pass

    def set_field_view_option(self, option):
        pass

    def set_message_box_mode(self, mode):
        pass

    def set_pos(self, list, para, pos):
        pass

    def set_pos_by_set(self, disp_val):
        pass

    def set_private_info_password(self, password):
        pass

    def set_text_file(self, data, format, option):
        pass

    def set_title_name(self, title):
        pass

    def set_user_info(self, user_info_id, value):
        pass

    def side_type(self, side_type):
        pass

    def signature(self, signature):
        pass

    def slash(self, slash):
        pass

    def solar_to_lunar(self, s_year, s_month, s_day, l_year, l_month, l_day, l_leap):
        pass

    def solar_to_lunar_by_set(self, s_year, s_month, s_day):
        pass

    def sort_delimiter(self, sort_delimiter):
        pass

    def strike_out(self, strike_out_type):
        pass

    def style_type(self, style_type):
        pass

    def subt_pos(self, subt_pos):
        pass

    def table_break(self, page_break):
        pass

    def table_format(self, table_format):
        pass

    def table_swap_type(self, tableswap):
        pass

    def table_target(self, table_target):
        pass

    def text_align(self, text_align):
        pass

    def text_art_align(self, text_art_align):
        pass

    def text_dir(self, text_direction):
        pass

    def text_flow_type(self, text_flow):
        pass

    def text_wrap_type(self, text_wrap):
        pass

    def un_select_ctrl(self):
        pass

    def v_align(self, v_align):
        pass

    def vert_rel(self, vert_rel):
        pass

    def view_flag(self, view_flag):
        pass

    def watermark_brush(self, watermark_brush):
        pass

    def width_rel(self, width_rel):
        pass
