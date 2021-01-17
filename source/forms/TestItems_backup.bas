Version =21
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17573
    DatasheetFontHeight =11
    ItemSuffix =88
    Right =25575
    Bottom =12240
    DatasheetGridlinesColor =15132391
    Filter ="[TestCode]LIKE \"*date*\""
    RecSrcDt = Begin
        0x2807c62cb795e540
    End
    RecordSource ="TestItem"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
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
            BorderColor =16777215
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
            LabelX =-1800
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
            LabelX =-1800
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
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2182
            BackColor =15064278
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =420
                    Top =180
                    Width =1680
                    Height =540
                    ForeColor =4210752
                    Name ="cmdRunTests"
                    Caption ="Run Tests..."
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =180
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =720
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2460
                    Top =180
                    Width =8880
                    TabIndex =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstMessages"
                    RowSourceType ="Value List"
                    RowSource ="18:06:58: [TestsRun] Complete. 13 test(s) run- 13 passed;18:06:58: [TestsRun] be"
                        "gin: prefix value;18:06:49: [TestsRun] Complete. 13 test(s) run- 12 passed;18:06"
                        ":49: [TestsRun] begin: prefix value;18:06:34: [TestsRun] Complete. 13 test(s) ru"
                        "n- 12 passed;18:06:34: [TestsRun] begin: prefix value;18:06:20: [TestsRun] Compl"
                        "ete. 12 test(s) run- 12 passed;18:06:20: [TestsRun] begin: prefix value"
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedLeft =2460
                    LayoutCachedTop =180
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =420
                    Top =840
                    Width =1680
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text6"
                    ControlSource ="=TestCurrentPrefix()"
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =840
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1155
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6780
                    Top =1800
                    Width =1620
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label27"
                    Caption ="Param1"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =6780
                    LayoutCachedTop =1800
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =2115
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8460
                    Top =1800
                    Width =1440
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label34"
                    Caption ="Param2"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =8460
                    LayoutCachedTop =1800
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =2115
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =9960
                    Top =1800
                    Width =1440
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label41"
                    Caption ="Param3"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =9960
                    LayoutCachedTop =1800
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =2115
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =14145
                    Top =1800
                    Width =1515
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label48"
                    Caption ="Expected"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =14145
                    LayoutCachedTop =1800
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =2115
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =15720
                    Top =1800
                    Width =1770
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label55"
                    Caption ="Actual"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =15720
                    LayoutCachedTop =1800
                    LayoutCachedWidth =17490
                    LayoutCachedHeight =2115
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =11460
                    Top =1800
                    Width =2625
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label62"
                    Caption ="ComparisonFunction"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =11460
                    LayoutCachedTop =1800
                    LayoutCachedWidth =14085
                    LayoutCachedHeight =2115
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =600
                    Top =1800
                    Width =540
                    Height =315
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label69"
                    Caption ="Pass?"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =600
                    LayoutCachedTop =1800
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =2115
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Top =1800
                    Height =315
                    TabIndex =4
                    BackColor =15983578
                    ForeColor =8355711
                    ColumnInfo ="\"\";\"\";\"10\";\"500\""
                    Name ="searchTestGroup"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT TestGroup FROM TestItem ORDER BY TestGroup; "
                    OnClick ="[Event Procedure]"
                    Format ="@;\"Test Group\""
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =1200
                    LayoutCachedTop =1800
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =2115
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1620
                    Width =426
                    Height =456
                    TabIndex =3
                    ForeColor =4210752
                    Name ="cmdClearFilters"
                    Caption ="Command84"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ab3255d65a0000000000000000000000000000000000000000 ,
                        0x3255d62d3255d693000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ae3255d6f93255d6360000000000000000000000003255d62d ,
                        0x3255d6db3255d61e000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6153255d6db3255d6f03255d630000000003255d6303255d6ea ,
                        0x3255d66300000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d61b3255d6c33255d6ed3255d66f3255d6ea3255d6ae ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x727272ff727272ff000000003255d6033255d6c63255d6ff3255d6de3255d60c ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x7272727e000000003255d6153255d6ab3255d6ff3255d6cf3255d6bd3255d696 ,
                        0x3255d609000000000000000000000000000000000000000000000000727272ff ,
                        0x000000003255d64e3255d6ed3255d6ff3255d6b73255d60c000000003255d645 ,
                        0x3255d6a53255d6420000000000000000000000000000000000000000727272ff ,
                        0x000000003255d6753255d6de3255d65a00000000000000000000000000000000 ,
                        0x000000003255d6270000000000000000000000000000000000000000727272ff ,
                        0x7272728100000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000007272726c727272ff ,
                        0x727272ff727272ff727272780000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000007272724e727272fc727272ff ,
                        0x727272ff727272ff727272ff7272725a00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000072727236727272f6727272ff727272ff ,
                        0x727272ff727272ff727272ff727272f972727242000000000000000000000000 ,
                        0x00000000000000000000000072727224727272ea727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272f07272722d0000000000000000 ,
                        0x000000000000000000000000727272d2727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272e40000000000000000 ,
                        0x000000000000000000000000727272f0727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =1620
                    LayoutCachedWidth =546
                    LayoutCachedHeight =2076
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2700
                    Top =1800
                    Width =1680
                    Height =315
                    TabIndex =5
                    BackColor =15983578
                    ForeColor =8355711
                    Name ="searchTestType"
                    RowSourceType ="Value List"
                    RowSource ="code;eval;code-array"
                    ColumnWidths ="720"
                    OnClick ="[Event Procedure]"
                    Format ="@;\"Test Type\""
                    GroupTable =2
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =2115
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4440
                    Top =1800
                    Width =2280
                    Height =315
                    TabIndex =6
                    BackColor =15983578
                    ForeColor =8355711
                    Name ="searchTestCode"
                    Format ="@;\"Test Code\""
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =1800
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =2115
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ForeTint =50.0
                    GroupTable =2
                End
            End
        End
        Begin Section
            Height =398
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1200
                    Top =30
                    Height =330
                    ColumnWidth =2880
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TestGroup"
                    ControlSource ="TestGroup"
                    DefaultValue ="\"comparisons\""
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =1200
                    LayoutCachedTop =30
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2700
                    Top =30
                    Width =1680
                    Height =330
                    ColumnWidth =2520
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TestType"
                    ControlSource ="TestType"
                    RowSourceType ="Value List"
                    RowSource ="code;eval;code-array"
                    ColumnWidths ="720"
                    GroupTable =2
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2700
                    LayoutCachedTop =30
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4440
                    Top =30
                    Width =2280
                    Height =330
                    ColumnWidth =2880
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TestCode"
                    ControlSource ="TestCode"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =30
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6780
                    Top =30
                    Width =1620
                    Height =330
                    ColumnWidth =2160
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Param1"
                    ControlSource ="Param1"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =6780
                    LayoutCachedTop =30
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8460
                    Top =30
                    Height =330
                    ColumnWidth =2160
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Param2"
                    ControlSource ="Param2"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =8460
                    LayoutCachedTop =30
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9960
                    Top =30
                    Height =330
                    ColumnWidth =2160
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Param3"
                    ControlSource ="Param3"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =9960
                    LayoutCachedTop =30
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14145
                    Top =30
                    Width =1515
                    Height =330
                    ColumnWidth =2160
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ExpectedResult"
                    ControlSource ="ExpectedResult"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =14145
                    LayoutCachedTop =30
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15720
                    Top =30
                    Width =1770
                    Height =330
                    ColumnWidth =2625
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Result"
                    ControlSource ="Result"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =15720
                    LayoutCachedTop =30
                    LayoutCachedWidth =17490
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11460
                    Top =30
                    Width =2625
                    Height =330
                    ColumnWidth =2520
                    TabIndex =7
                    BoundColumn =1
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="ComparisonFunction"
                    ControlSource ="ComparisonFunction"
                    RowSourceType ="Table/Query"
                    RowSource ="TestComparisonFunction"
                    ColumnWidths ="0;1400"
                    GroupTable =2
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =11460
                    LayoutCachedTop =30
                    LayoutCachedWidth =14085
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =2
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =600
                    Top =30
                    Width =540
                    Height =330
                    BorderColor =10921638
                    Name ="Passed"
                    ControlSource ="Passed"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =30
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
            End
        End
        Begin FormFooter
            Height =120
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function SetFilter()
    
    Dim sFilter As String
    Dim ctl As Control
    
    cmdClearFilters.SetFocus
    
    On Error Resume Next
    
    For Each ctl In Me.Controls
        If Left(ctl.Name, 6) = "search" Then
            If Not IsNull(ctl.Value) Then
                sFilter = sFilter & "[" & Mid(ctl.Name, 7) & "]"
                If TypeOf ctl Is ComboBox Then
                    ' / equals
                    sFilter = sFilter & "=""" & Replace(ctl.Value, """", """""") & """"
                ElseIf TypeOf ctl Is TextBox Then
                    ' / LIKE
                    sFilter = sFilter & "LIKE ""*" & Replace(ctl.Value, """", """""") & "*"""
                End If
            
                sFilter = sFilter & " AND "
            End If
        End If
    Next
    
    ' / strip last AND
    If Len(sFilter) > 5 Then sFilter = Left(sFilter, Len(sFilter) - 5)
    
    If Len(sFilter) > 0 Then
        If sFilter <> Me.Filter Then
            Me.Filter = sFilter
            Me.FilterOn = True
        End If
    Else
        Me.FilterOn = False
    End If
    
End Function

Public Function ClearFilterControls()
    
    Dim ctl As Control
    
    On Error Resume Next
    
    For Each ctl In Me.Controls
        If Left(ctl.Name, 6) = "search" Then
            ctl.Value = Null
        End If
    Next

End Function

Public Function TestsAddMessage(sMessage As String)
    On Error Resume Next
    lstMessages.AddItem sMessage, 0

End Function

Private Sub cmdClearFilters_Click()
    ClearFilterControls
    SetFilter
End Sub

Private Sub cmdRunTests_Click()
    
    Dim bRet As Boolean
    
    ' / save any current edits first
    DoCmd.RunCommand acCmdSaveRecord
    
    bRet = TestsRun()
    Me.Requery
    
    If Not bRet Then
        ClearFilterControls
        Me.Filter = "Passed=False"
        Me.FilterOn = True
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    
    ' / set defaults to this record's values
    On Error Resume Next
    
    SetControlDefaults True
    
End Sub

Private Sub SetControlDefaults(Optional bUseCurrentRecord As Boolean = True)
    
    Dim aControls, iControl As Long
    
    aControls = Array("TestGroup", "TestType", "TestCode", "ComparisonFunction")
    
    For iControl = LBound(aControls) To UBound(aControls)
        
        If bUseCurrentRecord Then
            Me(aControls(iControl)).DefaultValue = """" & Me(aControls(iControl)).Value & """"
        Else
            ' / TODO: use table default
            Me(aControls(iControl)).DefaultValue = Null
        End If
    Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then testItem_ContinuousUpDown KeyCode
End Sub

Private Sub Form_Load()
    On Error Resume Next
    SetControlDefaults False
    ClearFilterControls
    SetFilter
End Sub



Private Sub testItem_ContinuousUpDown(ByRef KeyCode As Integer)
    
    ' / put this in KeyUp
    
    On Error GoTo ErrHandler
    
    Select Case KeyCode
        Case vbKeyUp
            If ContinuousUpDownOk Then
                If Me.Dirty Then RunCommand acCmdSaveRecord
                RunCommand acCmdRecordsGoToPrevious
                KeyCode = 0
            End If

        Case vbKeyDown
            If ContinuousUpDownOk Then
                If Me.Dirty Then RunCommand acCmdSaveRecord
                If Not Me.NewRecord Then RunCommand acCmdRecordsGoToNext
                KeyCode = 0
            End If
    End Select

ExitHere:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 2046, 2101, 2113 'Already at first record, or save failed, or The value you entered isn't valid for this field.
            KeyCode = 0
        Case Else
            ' / TODO: log error (?)
    End Select
    Resume ExitHere

End Sub

Private Function ContinuousUpDownOk() As Boolean
    On Error GoTo Err_ContinuousUpDownOk
    'Purpose: Suppress moving up/down a record in a continuous form if:
    ' - control is not in the Detail section, or
    ' - multi-line text box (vertical scrollbar, or EnterKeyBehavior true).
    'Usage: Called by ContinuousUpDown.
    
    Dim bDontDoIt As Boolean
    Dim ctl As Control

    Set ctl = Screen.ActiveControl
    If ctl.Section = acDetail Then
        If TypeOf ctl Is TextBox Then
            bDontDoIt = ((ctl.EnterKeyBehavior) Or (ctl.ScrollBars > 1))
        End If
    Else
        bDontDoIt = True
    End If

Exit_ContinuousUpDownOk:
    ContinuousUpDownOk = Not bDontDoIt
    Set ctl = Nothing
Exit Function

Err_ContinuousUpDownOk:
    If Err.Number <> 2474 Then 'There's no active control
        ' / TODO: log error (?)
    End If
    Resume Exit_ContinuousUpDownOk
End Function

Private Sub searchTestCode_AfterUpdate()
    SetFilter
End Sub


Private Sub searchTestGroup_Click()
    SetFilter
End Sub

Private Sub searchTestType_Click()
    SetFilter
End Sub
