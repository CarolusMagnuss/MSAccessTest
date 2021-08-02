Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13833
    DatasheetFontHeight =11
    ItemSuffix =6
    Right =25575
    Bottom =12345
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xf569b345b484e540
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            ForeTint =50.0
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
        Begin Section
            Height =7540
            Name ="Detailbereich"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =623
                    Top =566
                    Width =3186
                    Height =628
                    ForeColor =4210752
                    Name ="btnFillResults"
                    Caption ="Ergebnisse füllen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =623
                    LayoutCachedTop =566
                    LayoutCachedWidth =3809
                    LayoutCachedHeight =1194
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5499
                    Top =566
                    Width =2661
                    Height =583
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnCreateTables"
                    Caption ="Tabellen Erstellen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5499
                    LayoutCachedTop =566
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =1149
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3505
                    Top =1700
                    Width =9302
                    Height =390
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtPath"
                    GridlineColor =10921638

                    LayoutCachedLeft =3505
                    LayoutCachedTop =1700
                    LayoutCachedWidth =12807
                    LayoutCachedHeight =2090
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =623
                            Top =1700
                            Width =2640
                            Height =375
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bezeichnungsfeld5"
                            Caption ="Dateipfad"
                            GridlineColor =10921638
                            LayoutCachedLeft =623
                            LayoutCachedTop =1700
                            LayoutCachedWidth =3263
                            LayoutCachedHeight =2075
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnCreateTables_Click()
Utilities.ExecuteSQLScript "C:\Users\steve\Nextcloud\Access Datenbank Jens\", "C:\Users\steve\Nextcloud\Access Datenbank Jens\Review Jens.accdb"

End Sub
Private Sub btnFillResults_Click()
    Dim Db As Object
    Dim ExportData As Recordset
    Dim dbPath As String
    Dim Studien As Recordset
    Dim Ort As Recordset
    Dim Ergebnisse As Recordset
    Dim conn As Object
    Dim RStemp As Recordset
    Dim SQLString As String
    
    Set conn = CreateObject("DAO.DBEngine.120")
    
    Set Studien = CurrentDb.OpenRecordset("Studien", dbOpenTable)
    Set Ort = CurrentDb.OpenRecordset("Ort", dbOpenTable)
    Set Ergebnisse = CurrentDb.OpenRecordset("Ergebnisse", dbOpenTable)
    txtPath.SetFocus
    
  
    
    'dbPath = "C:\Users\PillerSte\Nextcloud\Access Datenbank Jens\20200902_Extraktion_All ExportSP.xlsx"
    dbPath = txtPath.Text
    Set Db = conn.OpenDatabase(dbPath, False, True, "Excel 12.0;HDR=yes")
    'Set db = OpenDatabase(dbPath, False, True, "Excel 12.0;HDR=yes;")
    Set ExportData = Db.OpenRecordset("Select * From [Daten Export$]")
    
    ExportData.MoveFirst
    Do Until ExportData.EOF
        If Studien.RecordCount <> 0 Then
            Studien.MoveLast
        Else
            Studien.AddNew
        End If
        If ExportData![Study-Nr] <> Studien!id Then
            Studien.AddNew
            Studien!id = ExportData![Study-Nr]
            Studien!First_Autor = ExportData![First Author]
            Studien!Publication_Year = ExportData![Publication-Year]
            Studien![Study_Period] = ExportData![Study Period]
            Studien!Time_Period = ExportData![time period in years]
            Studien!study_Type = ExportData![Studie type]
            Studien.Update
        End If
        If Not IsNull(ExportData!Latitude) Then
            SQLString = "Select * From Ort Where Latitude=" & Replace(ExportData!Latitude, ",", ".") & ";"
            Set RStemp = CurrentDb.OpenRecordset(SQLString, dbOpenDynaset)
                If RStemp.RecordCount = 0 Then
                Ort.AddNew
                Ort!Region_Area = ExportData![Region/Area]
                Ort!Longitude = CDec(ExportData!Longitude)
                Ort!Latitude = CDec(ExportData!Latitude)
                Ort!country = ExportData!country
                Ort.Update
                End If
         End If
         If Not IsNull(ExportData![Study-Nr]) Then
             Ergebnisse.AddNew
             Ergebnisse!Studien_id = ExportData![Study-Nr]
                If Not IsNull(ExportData!Latitude) Then
                SQLString = "Select * From Ort Where Latitude=" & Replace(ExportData!Latitude, ",", ".") & ";"
                Set RStemp = CurrentDb.OpenRecordset(SQLString, dbOpenDynaset)
                Ergebnisse!Ort_id = RStemp!id
                End If
            Ergebnisse!Study_years = ExportData![Study-Years]
            Ergebnisse!Population = ExportData!Population
            Ergebnisse![Sample(N)] = ExportData![Sample (N)]
            Ergebnisse!AgeGroup = ExportData![Age(Groups)]
            If Not IsNull(ExportData![Age Paric Class]) Then
                Ergebnisse!AgeParicClass = CStr(ExportData![Age Paric Class])
            Else
                Ergebnisse!AgeParicClass = ExportData![Age Paric Class]
            End If
            Ergebnisse!medianAgeCD = ExportData![median Age CD]
            Ergebnisse!medianAgeUC = ExportData![median Age UC]
            Ergebnisse!medianAgeIC = ExportData![median Age IC]
            Ergebnisse!medianAgeIBD = ExportData![median Age IBD]
            Ergebnisse!Sex = ExportData!Sex
            Ergebnisse!CasesCDn = ExportData![Cases Cd n]
            Ergebnisse!CasesUCn = ExportData![Cases UCn]
            Ergebnisse!CasesICn = ExportData![Cases Ic n ]
            Ergebnisse!CasesIBD = ExportData![ Cases IBD n]
            Ergebnisse!CDIRR100k = ExportData![CDIRR100k]
            Ergebnisse!CD95CIlower = ExportData![CD95CIlower]
            Ergebnisse!CD95CIupper = ExportData![CD95CIupper]
            Ergebnisse!CDAgeStandard100k = ExportData!CDAgeStandard100k
            Ergebnisse!CD95CIlower2 = ExportData!CD95CIlower2
            Ergebnisse!CD95CIupper3 = ExportData!CD95CIupper3
            Ergebnisse!UCIRRper100k = ExportData!UCIRRper100k
            Ergebnisse!UC95CIlower = ExportData!UC95CIlower
            Ergebnisse!UC95CIupper = ExportData!UC95CIupper
            Ergebnisse!UCAgeStandard100k = ExportData!UCAgeStandard100k
            Ergebnisse!UC95CIlower4 = ExportData!UC95CIlower4
            Ergebnisse!UC95CIupper5 = ExportData!UC95CIupper5
            Ergebnisse!ICIRRper100k = ExportData!ICIRRper100k
            Ergebnisse!IC95CIlower = ExportData!IC95CIlower
            Ergebnisse!IC95CIupper = ExportData!IC95CIupper
            Ergebnisse!ICAgeStandard100k = ExportData!ICAgeStandard100k
            Ergebnisse!IC95CIlower6 = ExportData!IC95CIlower6
            Ergebnisse!IC95CIupper7 = ExportData!IC95CIupper7
            Ergebnisse!IBDIRR100k = ExportData!IBDIRR100k
            Ergebnisse!IBD95CIlower = ExportData!IBD95CIlower
            Ergebnisse!IBD95CIupper = ExportData!IBD95CIupper
            Ergebnisse!IBDAgeStandard100k = ExportData!IBDAgeStandard100k
            Ergebnisse!IBD95CIlower8 = ExportData!IBD95CIlower8
            Ergebnisse!IBD95CIupper9 = ExportData!IBD95CIupper9
            Ergebnisse!MeanAnnualIncreaseCDper100k = ExportData!MeanAnnualIncreaseCDper100k
            Ergebnisse![95CIlower] = ExportData![95CIlower]
            Ergebnisse![95CIupper] = ExportData![95CIupper]
            Ergebnisse!MeanAnnualIncreaseUCper100k = ExportData!MeanAnnualIncreaseUCper100k
            Ergebnisse![95CILower10] = ExportData![95CILower10]
            Ergebnisse![95CIupper11] = ExportData![95CIupper11]
            Ergebnisse!MeanAnnualIncreaseICper100k = ExportData!MeanAnnualIncreaseICper100k
            Ergebnisse![95CIlower12] = ExportData![95CIlower12]
            Ergebnisse![95CIupper13] = ExportData![95CIupper13]
            Ergebnisse!MeanAnnualIncreaseIBDper100k = ExportData!MeanAnnualIncreaseIBDper100k
            Ergebnisse![95CIlower14] = ExportData![95CIlower14]
            Ergebnisse![95CIupper15] = ExportData![95CIupper15]
            Ergebnisse!m_f_Ratio = ExportData!m_f_Ratio
            Ergebnisse!medianTimetoAdmissionCD = ExportData!medianTimetoAdmissionCD
            Ergebnisse!medianTimetoAdmissionUC = ExportData!medianTimetoAdmissionUC
            Ergebnisse!ParisCDL1 = ExportData!ParisCDL1
            Ergebnisse!ParisCDL2 = ExportData!ParisCDL2
            Ergebnisse!ParisCDL3 = ExportData!ParisCDL3
            Ergebnisse!ParisCDL4a = ExportData!ParisCDL4a
            Ergebnisse!ParisCDL4b = ExportData!ParisCDL4b
            Ergebnisse!ParisCDB1 = ExportData!ParisCDB1
            Ergebnisse!ParisCDB2 = ExportData!ParisCDB2
            Ergebnisse!ParisCDB3 = ExportData!ParisCDB3
            Ergebnisse!ParisCDB2B3 = ExportData![Paris CD B2B3]
            Ergebnisse!ParisCDP = ExportData![Paris CD P]
            Ergebnisse!ParisCDG0 = ExportData![Paris CD G0]
            Ergebnisse!ParisCDG1 = ExportData![Paris CD G1]
            Ergebnisse!ParisUCE1 = ExportData![Paris UC E1]
            Ergebnisse!ParisUCE2 = ExportData![Paris UC E2]
            Ergebnisse!ParisUCE3 = ExportData![Paris UC E3]
            Ergebnisse!ParisUCE4 = ExportData![Paris UC E4]
            Ergebnisse!ParisUCS0 = ExportData![Paris UC S0]
            Ergebnisse!ParisUCS1 = ExportData![Paris UC S1]
            Ergebnisse!ParisUCG0 = ExportData![Paris UC G0]
            Ergebnisse!ParisUCG1 = ExportData![Paris UC G1]
            Ergebnisse!externalManifestCD = ExportData![external Manifestations CD]
            Ergebnisse!externalManifestUC = ExportData![external Manifestations UC]
            Ergebnisse!externalManifestIBD = ExportData![external Manifestations IBD]
            Ergebnisse.Update
        End If
        ExportData.MoveNext
    Loop
    Ort.Close
    Ergebnisse.Close
    
    Set ExportData = Db.OpenRecordset("Select * From [SIGN$]")
    ExportData.MoveFirst
    Studien.MoveFirst
    
    Do Until ExportData.EOF
        Do Until Studien.EOF
            If ExportData!id = Studien!id Then
            Studien.Edit
            Studien!Quality! = ExportData!Quality
            Studien.Update
            Studien.MoveFirst
            ExportData.MoveNext
            Else:
            Studien.MoveNext
            End If
        Loop
    ExportData.MoveNext
    Loop
    
    
    
    Studien.Close
    ExportData.Close
        
    
End Sub

Private Sub Form_Load()
    txtPath.SetFocus
    txtPath.Text = "C:\Users\PillerSte\Nextcloud\Access Datenbank Jens\20200902_Extraktion_All ExportSP.xlsx"
End Sub
