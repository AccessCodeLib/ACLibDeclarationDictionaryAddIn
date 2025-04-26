Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7386
    DatasheetFontHeight =11
    ItemSuffix =335
    Left =7620
    Top =3045
    Right =20775
    Bottom =14775
    RecSrcDt = Begin
        0x956642cd6e4ee640
    End
    Caption ="Install Add-in"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
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
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7278
            Name ="Detailbereich"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =3137
                    Width =4740
                    Height =300
                    TabIndex =5
                    Name ="txtAddInTitle"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =3137
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =3437
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =3137
                            Width =1440
                            Height =300
                            Name ="lbltxtAddInName"
                            Caption ="Title"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =3137
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =3437
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =574
                    Width =4740
                    Height =300
                    TabIndex =1
                    Name ="txtFileName"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =574
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =874
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =574
                            Width =1440
                            Height =300
                            Name ="lblFileName"
                            Caption ="File name"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =574
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =874
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =3617
                    Width =4740
                    Height =300
                    TabIndex =6
                    Name ="txtAddInAuthor"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =3617
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =3917
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =3617
                            Width =1440
                            Height =300
                            Name ="lblAddInAuthor"
                            Caption ="Author"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =3617
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =3917
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =4097
                    Width =4740
                    Height =300
                    TabIndex =7
                    Name ="txtAddInCompany"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =4097
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =4397
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =4097
                            Width =1440
                            Height =300
                            Name ="lblAddInCompany"
                            Caption ="Company"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =4097
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =4397
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =4577
                    Width =4740
                    Height =1123
                    TabIndex =8
                    Name ="txtAddInComment"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    VerticalAnchor =2
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =4577
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =5700
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =4577
                            Width =1440
                            Height =1123
                            Name ="lblAddInComment"
                            Caption ="Comment"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            VerticalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =4577
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =5700
                            RowStart =9
                            RowEnd =9
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =574
                    Top =6240
                    Width =6240
                    Height =450
                    TabIndex =10
                    Name ="cmdInstallAddIn"
                    Caption ="Install Add-in"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =57
                    RightPadding =567
                    BottomPadding =567
                    HorizontalAnchor =2

                    LayoutCachedLeft =574
                    LayoutCachedTop =6240
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =6690
                    RowStart =11
                    RowEnd =11
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =2340
                    Width =4740
                    Height =291
                    TabIndex =4
                    Name ="txtAddInStartFunction"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =2631
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =2340
                            Width =1440
                            Height =291
                            Name ="lblAddInStartFunction"
                            Caption ="Start Function"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =2340
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =2631
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =1860
                    Width =4740
                    Height =300
                    TabIndex =3
                    Name ="txtAddInRegPathName"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =2160
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =1860
                            Width =1440
                            Height =300
                            Name ="Bezeichnungsfeld105"
                            Caption ="Name"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =1860
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =2160
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =574
                    Top =1534
                    Width =6240
                    Height =300
                    FontWeight =700
                    Name ="Bezeichnungsfeld112"
                    Caption ="USysRegInfo"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =57
                    RightPadding =567
                    BottomPadding =0
                    HorizontalAnchor =2
                    LayoutCachedLeft =574
                    LayoutCachedTop =1534
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =1834
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =574
                    Top =2811
                    Width =6240
                    Height =300
                    FontWeight =700
                    Name ="Bezeichnungsfeld150"
                    Caption ="Database properties"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =57
                    RightPadding =567
                    BottomPadding =0
                    HorizontalAnchor =2
                    LayoutCachedLeft =574
                    LayoutCachedTop =2811
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =3111
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2074
                    Top =1054
                    Width =4740
                    Height =300
                    TabIndex =2
                    Name ="txtAppTitle"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2074
                    LayoutCachedTop =1054
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =1354
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =1054
                            Width =1440
                            Height =300
                            Name ="Label247"
                            Caption ="AppTitle"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =1054
                            LayoutCachedWidth =2014
                            LayoutCachedHeight =1354
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =85
                    Width =0
                    Height =0
                    Name ="sysFirst"
                    Caption ="-"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3017
                    Top =5880
                    Width =3797
                    Height =300
                    TabIndex =9
                    Name ="cbCompileAddIn"
                    DefaultValue ="False"
                    GroupTable =1
                    RightPadding =567

                    LayoutCachedLeft =3017
                    LayoutCachedTop =5880
                    LayoutCachedWidth =6814
                    LayoutCachedHeight =6180
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =574
                            Top =5880
                            Width =2417
                            Height =300
                            ForeColor =0
                            Name ="Label325"
                            Caption ="Install Add-in as accde"
                            GroupTable =1
                            LeftPadding =57
                            RightPadding =0
                            HorizontalAnchor =2
                            LayoutCachedLeft =574
                            LayoutCachedTop =5880
                            LayoutCachedWidth =2991
                            LayoutCachedHeight =6180
                            RowStart =10
                            RowEnd =10
                            ColumnEnd =1
                            LayoutGroup =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =114
                    Top =6916
                    Width =7139
                    Height =223
                    FontSize =8
                    Name ="lblVersionInfo"
                    HorizontalAnchor =2
                    LayoutCachedLeft =114
                    LayoutCachedTop =6916
                    LayoutCachedWidth =7253
                    LayoutCachedHeight =7139
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6179
                    Top =113
                    Width =1026
                    Height =397
                    TabIndex =11
                    Name ="cmdAPI"
                    Caption ="  API"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a27b17d4a8d ,
                        0xb17d4acfb17d4affb17d4affb17d4acfb17d4a8db17d4a270000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a72b17d4af6b17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4af6b17d4a7200000000 ,
                        0x000000000000000000000000b17d4a06b17d4ab7b17d4affb17d4affb17d4aff ,
                        0xb17d4affffffffffffffffffb17d4affb17d4affb17d4affb17d4affb17d4ab7 ,
                        0xb17d4a060000000000000000b17d4a93b17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affffffffffffffffffb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4a9000000000b17d4a2db17d4afcb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affc1976effc1976effb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4afcb17d4a2db17d4a93b17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xc1976effffffffffe9daccffb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4a90b17d4adbb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb8895bfffefdfdfff9f4f0ffba8c5fffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4ad5b17d4af9b17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affdac2aafffffffffff4ede5ffb98b5dffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4af3b17d4af9b17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affe0cbb7fffffffffff3ebe3ffb8895bffb17d4affb17d4aff ,
                        0xb17d4affb17d4af0b17d4ad8b17d4affb17d4affb17d4affb17d4affbf946aff ,
                        0xb17d4affb17d4affb17d4affe3d0bdffffffffffdbc3acffb17d4affb17d4aff ,
                        0xb17d4affb17d4ad5b17d4a90b17d4affb17d4affb27f4cfff9f6f2ffffffffff ,
                        0xc1976effb17d4affb17d4affd4b79bffffffffffe0cbb7ffb17d4affb17d4aff ,
                        0xb17d4affb17d4a8db17d4a2db17d4afcb17d4affb17d4affd9c0a8ffffffffff ,
                        0xf5eee8ffd2b497ffd8bda3fffbf9f6fffdfcfbffc1976effb17d4affb17d4aff ,
                        0xb17d4afcb17d4a2a00000000b17d4a90b17d4affb17d4affb27f4cffd9c0a8ff ,
                        0xfefdfdfffffffffffffffffff7f1ecffc7a27dffb17d4affb17d4affb17d4aff ,
                        0xb17d4a8d0000000000000000b17d4a06b17d4ab7b17d4affb17d4affb17d4aff ,
                        0xb78859ffc7a27dffc1976effb17d4affb17d4affb17d4affb17d4affb17d4ab7 ,
                        0xb17d4a0600000000000000000000000000000000b17d4a72b17d4af6b17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4af6b17d4a7200000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a27b17d4a8d ,
                        0xb17d4accb17d4afcb17d4afcb17d4accb17d4a8db17d4a270000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6179
                    LayoutCachedTop =113
                    LayoutCachedWidth =7205
                    LayoutCachedHeight =510
                    PictureCaptionArrangement =5
                End
            End
        End
    End
End
CodeBehindForm
' See "InstallAddInForm.cls"
