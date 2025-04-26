Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =10148
    DatasheetFontHeight =11
    ItemSuffix =81
    Left =7620
    Top =3045
    Right =20775
    Bottom =14775
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x1b36415d9252e640
    End
    Caption ="Declarations"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =4210752
            FontName ="Calibri"
            GridlineColor =10921638
            ForeTint =75.0
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackColor =14136213
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =14136213
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =15060409
            HoverTint =40.0
            PressedColor =9592887
            PressedShade =75.0
            HoverForeColor =4210752
            HoverForeTint =75.0
            PressedForeColor =4210752
            PressedForeTint =75.0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BackThemeColorIndex =1
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            Height =6236
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =789
                    Width =5416
                    Height =4542
                    FontSize =9
                    TabIndex =11
                    Name ="lbDictData"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2835"
                    AfterUpdate ="[Event Procedure]"
                    HorizontalAnchor =2
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =789
                    LayoutCachedWidth =5476
                    LayoutCachedHeight =5331
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =57
                    Width =5103
                    Height =456
                    TabIndex =1
                    Name ="filtDiff"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    HorizontalAnchor =2

                    LayoutCachedLeft =57
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =456
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =120
                            Top =60
                            Width =840
                            Height =345
                            Name ="labFilterSelection"
                            Caption ="Show"
                            LayoutCachedLeft =120
                            LayoutCachedTop =60
                            LayoutCachedWidth =960
                            LayoutCachedHeight =405
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1020
                            Top =80
                            OptionValue =1
                            Name ="Option4"

                            LayoutCachedLeft =1020
                            LayoutCachedTop =80
                            LayoutCachedWidth =1280
                            LayoutCachedHeight =320
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1250
                                    Top =50
                                    Width =1200
                                    Height =315
                                    Name ="Label5"
                                    Caption ="Differences "
                                    LayoutCachedLeft =1250
                                    LayoutCachedTop =50
                                    LayoutCachedWidth =2450
                                    LayoutCachedHeight =365
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2608
                            Top =80
                            TabIndex =1
                            OptionValue =0
                            Name ="Option6"

                            LayoutCachedLeft =2608
                            LayoutCachedTop =80
                            LayoutCachedWidth =2868
                            LayoutCachedHeight =320
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2838
                                    Top =50
                                    Width =825
                                    Height =315
                                    Name ="Label7"
                                    Caption ="Full list"
                                    LayoutCachedLeft =2838
                                    LayoutCachedTop =50
                                    LayoutCachedWidth =3663
                                    LayoutCachedHeight =365
                                End
                            End
                        End
                    End
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5674
                    Top =1911
                    Width =4380
                    Height =1354
                    TabIndex =8
                    ForeColor =0
                    Name ="lbVariations"
                    RowSourceType ="Value List"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    HorizontalAnchor =1
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =5674
                    LayoutCachedTop =1911
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =3265
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5674
                    Top =574
                    Width =4380
                    Height =317
                    FontWeight =700
                    TabIndex =6
                    Name ="txtWord"
                    ControlSource ="=[lbDictData]"
                    Format ="@;;\"(select item)\""
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0xbfbfbf00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740057006f00720064005d0020004900730020004e0075006c00 ,
                        0x6c0000000000
                    End
                    GroupTable =1
                    HorizontalAnchor =1

                    LayoutCachedLeft =5674
                    LayoutCachedTop =574
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =891
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000bfbfbf00ffffff00110000005b00 ,
                        0x74007800740057006f00720064005d0020004900730020004e0075006c006c00 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5494
                    Top =60
                    Height =397
                    TabIndex =3
                    Name ="cmdUpdateDict"
                    Caption ="update data"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5494
                    LayoutCachedTop =60
                    LayoutCachedWidth =7195
                    LayoutCachedHeight =457
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3741
                    Top =56
                    Width =1248
                    Height =300
                    TabIndex =2
                    Name ="txtDictInfo"
                    HorizontalAnchor =2

                    LayoutCachedLeft =3741
                    LayoutCachedTop =56
                    LayoutCachedWidth =4989
                    LayoutCachedHeight =356
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =85
                    Width =0
                    Height =0
                    Name ="Command15"
                    Caption ="sysFirst"

                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =5669
                    Top =4081
                    Width =2595
                    Height =1402
                    Name ="Label17"
                    Caption ="Test steps:\015\012   1. [update data]\015\012   2. change lettercase\015\012   "
                        "3. [update data]\015\012   4. show differences "
                    HorizontalAnchor =1
                    VerticalAnchor =1
                    LayoutCachedLeft =5669
                    LayoutCachedTop =4081
                    LayoutCachedWidth =8264
                    LayoutCachedHeight =5483
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5674
                    Top =1251
                    Width =4380
                    Height =300
                    TabIndex =7
                    Name ="txtCurrentLetterCase"
                    Format ="@;;---"
                    ConditionalFormat = Begin
                        0x01000000dc000000020000000100000000000000000000001200000001000000 ,
                        0xbfbfbf00ffffff000100000000000000130000003d00000001010000ba141900 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740057006f00720064005d0020004900730020004e0075006c00 ,
                        0x6c000000000053007400720043006f006d00700028005b007400780074005700 ,
                        0x6f00720064005d002c005b00740078007400430075007200720065006e007400 ,
                        0x560061006c00750065005d002c00300029003c003e00300000000000
                    End
                    GroupTable =1
                    TopPadding =0
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    HorizontalAnchor =1

                    LayoutCachedLeft =5674
                    LayoutCachedTop =1251
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =1551
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000bfbfbf00ffffff00110000005b00 ,
                        0x74007800740057006f00720064005d0020004900730020004e0075006c006c00 ,
                        0x0000000000000000000000000000000000000000000100000000000000010100 ,
                        0x00ba141900ffffff002900000053007400720043006f006d00700028005b0074 ,
                        0x007800740057006f00720064005d002c005b0074007800740043007500720072 ,
                        0x0065006e007400560061006c00750065005d002c00300029003c003e00300000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5674
                    Top =951
                    Width =4380
                    Height =300
                    Name ="Label19"
                    Caption ="Current lettercase:"
                    GroupTable =1
                    BottomPadding =0
                    HorizontalAnchor =1
                    LayoutCachedLeft =5674
                    LayoutCachedTop =951
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =1251
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5674
                    Top =1611
                    Width =4380
                    Height =300
                    Name ="Label20"
                    Caption ="Variations:"
                    GroupTable =1
                    BottomPadding =0
                    HorizontalAnchor =1
                    LayoutCachedLeft =5674
                    LayoutCachedTop =1611
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =1911
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5674
                    Top =3626
                    Width =3403
                    Height =343
                    TabIndex =9
                    Name ="txtSelectedLetterCase"
                    ControlSource ="=[lbVariations]"
                    Format ="@;;---"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0xbfbfbf00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740057006f00720064005d0020004900730020004e0075006c00 ,
                        0x6c0000000000
                    End
                    GroupTable =1
                    TopPadding =0
                    HorizontalAnchor =1

                    LayoutCachedLeft =5674
                    LayoutCachedTop =3626
                    LayoutCachedWidth =9077
                    LayoutCachedHeight =3969
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000bfbfbf00ffffff00110000005b00 ,
                        0x74007800740057006f00720064005d0020004900730020004e0075006c006c00 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5674
                    Top =3291
                    Width =4380
                    Height =300
                    Name ="Label45"
                    Caption ="Change to:"
                    GroupTable =1
                    TopPadding =0
                    HorizontalAnchor =1
                    LayoutCachedLeft =5674
                    LayoutCachedTop =3291
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =3591
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =9137
                    Top =3626
                    Width =917
                    Height =343
                    TabIndex =10
                    Name ="cmdChangeLetterCase"
                    Caption ="Commit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    HorizontalAnchor =1

                    LayoutCachedLeft =9137
                    LayoutCachedTop =3626
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =3969
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3754
                    Top =5443
                    Width =1697
                    Height =300
                    TabIndex =13
                    Name ="cmdSaveToTable"
                    Caption ="Export to table"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =3754
                    LayoutCachedTop =5443
                    LayoutCachedWidth =5451
                    LayoutCachedHeight =5743
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1997
                    Top =5443
                    Width =1697
                    Height =300
                    TabIndex =12
                    Name ="cmdLoadFromTable"
                    Caption ="Load from table"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =1997
                    LayoutCachedTop =5443
                    LayoutCachedWidth =3694
                    LayoutCachedHeight =5743
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =5443
                            Width =1875
                            Height =300
                            FontSize =10
                            Name ="Label56"
                            Caption ="Table: USysDeclDict"
                            GroupTable =2
                            LayoutCachedLeft =60
                            LayoutCachedTop =5443
                            LayoutCachedWidth =1935
                            LayoutCachedHeight =5743
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =5503
                    Top =5563
                    Width =4525
                    Height =540
                    LeftMargin =57
                    Name ="lblTableRecInfo"
                    HorizontalAnchor =2
                    LayoutCachedLeft =5503
                    LayoutCachedTop =5563
                    LayoutCachedWidth =10028
                    LayoutCachedHeight =6103
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =5494
                    Top =5897
                    Width =4550
                    Height =223
                    FontSize =8
                    Name ="lblVersionInfo"
                    HorizontalAnchor =2
                    LayoutCachedLeft =5494
                    LayoutCachedTop =5897
                    LayoutCachedWidth =10044
                    LayoutCachedHeight =6120
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =720
                    Top =446
                    Width =3073
                    Height =300
                    TabIndex =5
                    Name ="filtWord"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =720
                    LayoutCachedTop =446
                    LayoutCachedWidth =3793
                    LayoutCachedHeight =746
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =60
                            Top =446
                            Width =660
                            Height =300
                            Name ="Label67"
                            Caption ="Filter:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =446
                            LayoutCachedWidth =720
                            LayoutCachedHeight =746
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3981
                    Top =429
                    Width =1495
                    Height =328
                    TabIndex =4
                    Name ="cmdRemoveFilter"
                    Caption ="remove filter"
                    OnClick ="[Event Procedure]"
                    BackStyle =0

                    LayoutCachedLeft =3981
                    LayoutCachedTop =429
                    LayoutCachedWidth =5476
                    LayoutCachedHeight =757
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    Gradient =0
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3754
                    Top =5803
                    Width =1697
                    Height =317
                    TabIndex =15
                    Name ="cmdSaveToFile"
                    Caption ="Export to file"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =3754
                    LayoutCachedTop =5803
                    LayoutCachedWidth =5451
                    LayoutCachedHeight =6120
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1997
                    Top =5803
                    Width =1697
                    Height =317
                    TabIndex =14
                    Name ="cmdLoadFromFile"
                    Caption ="Load from file"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =1997
                    LayoutCachedTop =5803
                    LayoutCachedWidth =3694
                    LayoutCachedHeight =6120
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =5803
                            Width =1875
                            Height =317
                            FontSize =10
                            Name ="Bezeichnungsfeld71"
                            Caption ="File: DeclarationDict"
                            GroupTable =2
                            LayoutCachedLeft =60
                            LayoutCachedTop =5803
                            LayoutCachedWidth =1935
                            LayoutCachedHeight =6120
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9008
                    Top =56
                    Width =1026
                    Height =397
                    TabIndex =16
                    Name ="cmdAPI"
                    Caption ="  API"
                    OnClick ="[Event Procedure]"
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

                    LayoutCachedLeft =9008
                    LayoutCachedTop =56
                    LayoutCachedWidth =10034
                    LayoutCachedHeight =453
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                End
            End
        End
    End
End
CodeBehindForm
' See "DeclarationDictForm.cls"
