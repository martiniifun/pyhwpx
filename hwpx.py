import win32com.client as win32
import pythoncom
import pandas as pd
import numpy as np
import re
import os


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
