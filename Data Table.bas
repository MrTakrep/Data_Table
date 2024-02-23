Dim L_KolomA                                                                                                        As MSForms.Label
Dim L_KolomB                                                                                                        As MSForms.Label
Dim L_KolomC                                                                                                       As MSForms.Label
Dim L_KolomD                                                                                                        As MSForms.Label
Dim L_KolomE                                                                                                        As MSForms.Label
Dim L_KolomF                                                                                                        As MSForms.Label
Dim L_KolomG                                                                                                       As MSForms.Label
Dim L_Edit                                                                                                                As MSForms.Label
Dim L_Delete                                                                                                          As MSForms.Label
Dim Br_Table                                                                                                         As MSForms.Label
Dim Background_Table                                                                                                 As MSForms.Label
Dim Total_Data_Table                                                                                    As MSForms.Label
Dim LS, i, TablesWidth                                                                                       As Long
Dim WS                                                                                                                      As Worksheet

Sub DataList1()
    Set WS = ThisWorkbook.Sheets("Sheet1")
    LS = WS.Range("A" & WS.Rows.Count).End(xlUp).Row

    DefaultColor = RGB(255, 255, 255)
    DataCOlor = RGB(72, 89, 112)
    EditColor = RGB(231, 139, 3)
    DeleteColor = RGB(192, 0, 0)
    BrColor = RGB(210, 215, 224)
    GenapColor = RGB(230, 233, 238)
    GanjilColor = RGB(248, 249, 250)
    AktifColor = RGB(0, 204, 0)
    NonaktifColor = RGB(192, 0, 0)
    HeaderColor = RGB(67, 94, 190)
    
    DefaultTopPos = 10
    DefaultLeftPos = 12
    
    SmallTopMargin = 0
    DefaultTopMargin = 4
    NormalTopMargin = 6
    BigTopMargin = 8
    MonsterTopMargin = 12
    
    SmallFontSize = 6
    DefaultFontSize = 8
    NormalFontSize = 10
    BigFontSize = 12
    MonsterFontSize = 14
    
    SmallHeight = 14
    DefaultHeight = 16
    NormalHeight = 18
    BigHeight = 20
    MonsterHeight = 24
    
    SmallHoverHeight = 16
    DefaultHoverHeight = 20
    NormalHoverHeight = 22
    BigHoverHeight = 24
    MonsterHoverHeight = 28
    
    WidthEdit = 26
    WidthDelete = 22
    WidthKolomA = 42
    WidthKolomB = 42
    WidthKolomC = 78
    WidthKolomD = 44
    WidthKolomE = 54
    
    WS.Range("A:Z").EntireColumn.AutoFit
    For i = 2 To LS
        'Membuat data kolom A_________________________________________
        With UserForm1.Frame_Data
            Set L_KolomA = .Controls.Add("Forms.Label.1", "L_KolomA" & i, True)
            With L_KolomA
                .Caption = WS.Range("A" & i).Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_" & WS.Range("A" & i)
                .Left = 0
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DataCOlor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("A1").Width
            End With
        End With
        'Membuat data kolom B_________________________________________
        With UserForm1.Frame_Data
            Set L_KolomB = .Controls.Add("Forms.Label.1", "L_KolomB" & i, True)
            With L_KolomB
                .Caption = WS.Range("B" & i).Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_" & WS.Range("B" & i)
                .Left = WS.Range("A1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DataCOlor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("B1").Width
            End With
        End With
        'Membuat data kolom C_________________________________________
        With UserForm1.Frame_Data
            Set L_KolomC = .Controls.Add("Forms.Label.1", "L_KolomC" & i, True)
            With L_KolomC
                .Caption = WS.Range("C" & i).Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_" & WS.Range("C" & i)
                .Left = WS.Range("A1:B1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DataCOlor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("C1").Width
            End With
        End With
        'Membuat data kolom D_________________________________________
        With UserForm1.Frame_Data
            Set L_KolomD = .Controls.Add("Forms.Label.1", "L_KolomD" & i, True)
            With L_KolomD
                .Caption = WS.Range("D" & i).Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_" & WS.Range("D" & i)
                .Left = WS.Range("A1:C1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DataCOlor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("D1").Width
            End With
        End With
        'Membuat data kolom E_________________________________________
        With UserForm1.Frame_Data
            Set L_KolomE = .Controls.Add("Forms.Label.1", "L_KolomE" & i, True)
            With L_KolomE
                .Caption = Format(WS.Range("E" & i).Value, "#,##0.0")
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_" & WS.Range("E" & i)
                .Left = WS.Range("A1:D1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DataCOlor
                .TextAlign = fmTextAlignRight
                .Width = WS.Range("E1").Width
            End With
        End With
        'Membuat tombol edit_________________________________________
        With UserForm1.Frame_Data
            Set L_Edit = .Controls.Add("Forms.Label.1", "L_Edit" & i, True)
            With L_Edit
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_edit"
                .Left = WS.Range("A1:E1").Width + 2
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 100, 60)
                .ForeColor = EditColor
                .Font.Size = NormalFontSize
                .Font.Name = "myicons"
                .Caption = ChrW("&H0033")
                .Width = WidthEdit
                .TextAlign = fmTextAlignCenter
            End With
        End With
        'Membuat tombol delete_________________________________________
        With UserForm1.Frame_Data
            Set L_Delete = .Controls.Add("Forms.Label.1", "L_Delete" & i, True)
            With L_Delete
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A" & i) & "_delete"
                .Left = WS.Range("A1:E1").Width + WidthEdit + 4
                .Height = DefaultHeight
                If i = 2 Then
                    .Top = DefaultTopMargin
                    Else
                    .Top = (DefaultTopMargin * (i - 1)) + (DefaultHeight * (i - 2))
                End If
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 100, 60)
                .ForeColor = DeleteColor
                .Font.Size = NormalFontSize
                .Font.Name = "myicons"
                .Caption = ChrW("&HE0AC")
                .Width = WidthEdit
                .TextAlign = fmTextAlignCenter
            End With
        End With
        'Membuat backrgound table_________________________________________
        With UserForm1.Frame_Data
            Set Background_Table = .Controls.Add("Forms.Label.1", "Background_Table" & i, True)
            With Background_Table
                .Tag = WS.Range("A" & i) & "_back"
                .Left = 0
                .Height = DefaultHoverHeight
                If i = 2 Then
                    .Top = 0
                    Else
                    .Top = (DefaultHoverHeight * (i - 2))
                End If
                If i Mod 2 = 0 Then
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = GenapColor
                    Else
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = GanjilColor
                End If
                .Width = WS.Range("A1:E1").Width + WidthEdit + WidthDelete + 4
                .ZOrder (1)
            End With
        End With
    Next i
        'Membuat header kolom A_________________________________________
        With UserForm1.Frame_Header
            Set L_KolomA = .Controls.Add("Forms.Label.1", "L_KolomA" & i, True)
            With L_KolomA
                .Caption = WS.Range("A1").Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("A1") & "_" & WS.Range("A1")
                .Left = 0
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                .Top = DefaultTopMargin
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DefaultColor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("A1").Width
            End With
        End With
        'Membuat header kolom B_________________________________________
        With UserForm1.Frame_Header
            Set L_KolomB = .Controls.Add("Forms.Label.1", "L_KolomB" & i, True)
            With L_KolomB
                .Caption = WS.Range("B1").Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("B1") & "_" & WS.Range("B1")
                .Left = WS.Range("A1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                .Top = DefaultTopMargin
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DefaultColor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("B1").Width
            End With
        End With
        'Membuat header kolom C_________________________________________
        With UserForm1.Frame_Header
            Set L_KolomC = .Controls.Add("Forms.Label.1", "L_KolomC" & i, True)
            With L_KolomC
                .Caption = WS.Range("C1").Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("C1") & "_" & WS.Range("C1")
                .Left = WS.Range("A1:B1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                .Top = DefaultTopMargin
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DefaultColor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("C1").Width
            End With
        End With
        'Membuat header kolom D_________________________________________
        With UserForm1.Frame_Header
            Set L_KolomD = .Controls.Add("Forms.Label.1", "L_KolomD" & i, True)
            With L_KolomD
                .Caption = WS.Range("D1").Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("D1") & "_" & WS.Range("D1")
                .Left = WS.Range("A1:C1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                .Top = DefaultTopMargin
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DefaultColor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("D1").Width
            End With
        End With
        'Membuat header kolom E_________________________________________
        With UserForm1.Frame_Header
            Set L_KolomE = .Controls.Add("Forms.Label.1", "L_KolomE" & i, True)
            With L_KolomE
                .Caption = WS.Range("E1").Value
                .Font.Size = DefaultFontSize
                .Tag = WS.Range("E1") & "_" & WS.Range("E1")
                .Left = WS.Range("A1:D1").Width
                .Font.Name = "Poppins Medium"
                .Height = DefaultHeight
                .Top = DefaultTopMargin
                .BackStyle = fmBackStyleTransparent
                .BackColor = RGB(255, 160, 0)
                .ForeColor = DefaultColor
                .TextAlign = fmTextAlignCenter
                .Width = WS.Range("E1").Width
            End With
        End With
        'Membuat backrgound table header_________________________________________
        With UserForm1.Frame_Header
            Set Background_Table = .Controls.Add("Forms.Label.1", "Background_Table" & i, True)
            With Background_Table
                .Tag = WS.Range("A1") & "_back"
                .Left = 0
                .Height = DefaultHoverHeight
                .Top = 0
                .BackStyle = fmBackStyleOpaque
                .BackColor = HeaderColor
                .Width = WS.Range("A1:E1").Width + WidthEdit + WidthDelete + 4
                .ZOrder (1)
            End With
        End With
End Sub
