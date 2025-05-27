from typing import Any, Literal


class ParamHelpers:
    """
    파라미터 헬퍼메서드 : 별도의 동작은 하지 않고, 파라미터 변환, 연산 등을 돕는다.
    """

    hwp: Any

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

    def FillAreaType(self, fill_area):
        return self.hwp.FillAreaType(FillArea=fill_area)

    def FindDir(self, find_dir: Literal["Forward", "Backward", "AllDoc"] = "Forward"):
        return self.hwp.FindDir(FindDir=find_dir)

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

    def HeadType(self, heading_type: Literal["None", "Outline", "Number", "Bullet"]) -> int:
        """
        문단 종류를 결정할 때 사용하는 헬퍼함수

        현재 문단의 머리에 '개요 번호'나 '문단 번호', '그럼리표' 등을 넣어 문단 종류를 바꿀 것인지,
        '없음'을 선택해 보통 모양의 문단으로 놓아둘 것인지를 선택.

        Args:
            heading_type: 문단 종류

                - "None": 없음(보통 모양의 문단)
                - "Outline": 개요 문단
                - "Number": 번호 문단
                - "Bullet": 글머리표 문단

        Returns:
            int: 옵션에 해당하는 정수를 리턴
        """
        return self.hwp.HeadType(HeadingType=heading_type)

    def HeightRel(self, height_rel):
        return self.hwp.HeightRel(HeightRel=height_rel)

    def Hiding(self, hiding):
        return self.hwp.Hiding(Hiding=hiding)

    def HorzRel(self, horz_rel):
        return self.hwp.HorzRel(HorzRel=horz_rel)

    def HwpLineType(
        self,
        line_type: Literal[
            "None",
            "Solid",
            "Dash",
            "Dot",
            "DashDot",
            "DashDotDot",
            "LongDash",
            "Circle",
            "DoubleSlim",
            "SlimThick",
            "ThickSlim",
            "SlimThickSlim",
        ] = "Solid",
    ):
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

    def HwpLineWidth(
        self,
        line_width: Literal[
            "0.1mm",
            "0.12mm",
            "0.15mm",
            "0.2mm",
            "0.25mm",
            "0.3mm",
            "0.4mm",
            "0.5mm",
            "0.6mm",
            "0.7mm",
            "1.0mm",
            "1.5mm",
            "2.0mm",
            "3.0mm",
            "4.0mm",
            "5.0mm",
        ] = "0.1mm",
    ) -> int:
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
        return self.hwp.LunarToSolar(
            lYear=l_year,
            lMonth=l_month,
            lDay=l_day,
            lLeap=l_leap,
            sYear=s_year,
            sMonth=s_month,
            sDay=s_day,
        )

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

    @staticmethod
    def hwp_unit_to_mili(hwp_unit: int) -> float:
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

    def HwpUnitToMili(self, hwp_unit: int) -> float:
        return self.hwp_unit_to_mili(hwp_unit)

    def NumberFormat(self, num_format: Literal[
        "Digit",  # 123
        "CircledDigit",  # ①
        "RomanCapital",  # I
        "RomanSmall",  # i
        "LatinCapital",  # A
        "LatinSmall",  # a
        "CircledLatinCapital",  # Ⓐ
        "CircledLatinSmall",  # ⓐ
        "HangulSyllable",  # 가나다
        "CircledHangulSyllable",  # ㉯
        "HangulJamo",  # ㄱㄴㄷ
        "CircledHangulJamo",  # ㉠
        "HangulPhonetic",  # 일이삼
        "Ideograph",  # 一
        "CircledIdeograph",  # ㊀
        "DecagonCircle",  # 갑을병
        "DecagonCircleHanja",  # 甲
    ]):
        """
        개요번호 사용자 정의를 위해 미리 정의된 포맷 모음
        
        Args:
            num_format(str):
                포맷 종류.

                    - "Digit": 123
                    - "CircledDigit": ①
                    - "RomanCapital": I
                    - "RomanSmall": i
                    - "LatinCapital": A
                    - "LatinSmall": a
                    - "CircledLatinCapital": Ⓐ
                    - "CircledLatinSmall": ⓐ
                    - "HangulSyllable": 가나다
                    - "CircledHangulSyllable": ㉯
                    - "HangulJamo": ㄱㄴㄷ
                    - "CircledHangulJamo": ㉠
                    - "HangulPhonetic": 일이삼
                    - "Ideograph": 一
                    - "CircledIdeograph": ㊀
                    - "DecagonCircle": 갑을병
                    - "DecagonCircleHanja": 甲

        Returns:
            int: 해당 정수로 치환됨(Digit=0, CircledDigit=1, ... DecagonCircleHanja=16)

        Examples:
            >>> # 개요번호 사용자 정의
            >>> from pyhwpx import Hwp
            >>> hwp = Hwp(new=True)
            >>> pset = hwp.HParameterSet.HSecDef
            >>> hwp.HAction.GetDefault("OutlineNumber", pset.HSet)
            >>> pset.OutlineShape.StrFormatLevel0 = "^1."
            >>> pset.OutlineShape.NumFormatLevel0 = hwp.NumberFormat("RomanCapital")  # <---
            >>> pset.OutlineShape.StartNumber0 = 1
            >>> pset.OutlineShape.NewList = 0
            >>> pset.HSet.SetItem("ApplyClass", 24)  # 앞 구역의 개요 번호에 이어서
            >>> pset.HSet.SetItem("ApplyTo", 3)  # 적용범위(2:현재구역, 3:문서 전체, 4:새 구역으로)
            >>> hwp.HAction.Execute("OutlineNumber", pset.HSet)
            True
        """
        return self.hwp.NumberFormat(NumFormat=num_format)

    def Numbering(self, numbering):
        return self.hwp.Numbering(Numbering=numbering)

    def PageNumPosition(
        self,
        pagenumpos: Literal[
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
    ):
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
        return self.hwp.SolarToLunar(
            sYear=s_year,
            sMonth=s_month,
            sDay=s_day,
            lYear=l_year,
            lMonth=l_month,
            lDay=l_day,
            lLeap=l_leap,
        )

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
