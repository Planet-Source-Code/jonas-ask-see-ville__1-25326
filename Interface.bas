Attribute VB_Name = "Interface"
Public Sub MnuResClick(I)
    For a = 0 To Form1.IconResSub.Count - 1
        Form1.IconResSub(a).BorderStyle = 0
    Next a
    Form1.IconResSub(I).BorderStyle = 1
    
    Select Case I
    Case 0: SelItem = "res1"
    Case 1: SelItem = "res2"
   End Select
End Sub
Public Sub MnuPwrClick(I)
    For a = 0 To Form1.IconPwrSub.Count - 1
        Form1.IconPwrSub(a).BorderStyle = 0
    Next a
    Form1.IconPwrSub(I).BorderStyle = 1
    
    Select Case I
    Case 0: SelItem = "lines"
    Case 1: SelItem = "plant"
    Case 2: SelItem = "road"
    Case 3: SelItem = "bridge"
    End Select
End Sub


Public Sub MnuScenClick(I)
    For a = 0 To Form1.IconScenSub.Count - 1
        Form1.IconScenSub(a).BorderStyle = 0
    Next a
    Form1.IconScenSub(I).BorderStyle = 1
    
    Select Case I
    Case 0: SelItem = "trees"
    Case 1: SelItem = "park1"
    Case 2: SelItem = "park2"
    End Select
    
End Sub

Public Sub ClearMenu()
    With Form1
    For a = 0 To .PicIcon.Count - 1
        .PicIcon(a).BorderStyle = 0
    Next a
    
    For a = 0 To .IconPwrSub.Count - 1
        .IconPwrSub(a).BorderStyle = 0
        .IconPwrSub(a).Visible = False
    Next a
    For a = 0 To .IconResSub.Count - 1
        .IconResSub(a).BorderStyle = 0
        .IconResSub(a).Visible = False
    Next a
    For a = 0 To .IconScenSub.Count - 1
        .IconScenSub(a).BorderStyle = 0
        .IconScenSub(a).Visible = False
    Next a
    End With
End Sub
