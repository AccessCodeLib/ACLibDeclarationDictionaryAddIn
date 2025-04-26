﻿Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =2
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9072
    DatasheetFontHeight =11
    ItemSuffix =4
    Left =7620
    Top =3045
    Right =20775
    Bottom =14775
    RecSrcDt = Begin
        0x9487851a5859e640
    End
    Caption ="API Information"
    DatasheetFontName ="Calibri"
    FetchDefaults =0
    FilterOnLoad =0
    OrderByOnLoad =0
    FetchDefaults =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =6804
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Width =9072
                    Height =6804
                    Name ="InfoReportAPI"
                    SourceObject ="Report.InfoReportAPI"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedWidth =9072
                    LayoutCachedHeight =6804
                End
            End
        End
    End
End
CodeBehindForm
' See "InfoFormAPI.cls"
