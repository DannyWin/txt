Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public dicData As Object
Public dicCatagory As Object

Dim DaAN As String
Dim randomArray() As Integer
Dim Next_N As Integer
Dim IsRunning As Boolean

Dim rdmSeq As Boolean
Dim rdmTip As Boolean
Dim rdmCount As Boolean
Dim containWords As Boolean
Dim containPhrases As Boolean
Dim containSentences As Boolean
Dim containCultures As Boolean
Dim count As Integer

Dim Excel As Object
Dim Subject_Total As Integer

Private Function GetRandomArray(ByVal left As Integer, ByVal right As Integer, ByRef randomArray() As Integer)
    Dim j As Integer, i As Integer

    Randomize             ' 随机数种子

    For i = 0 To (right - left)
        randomArray(i) = Int((right - left + 1) * Rnd + left)
        For j = 0 To (i - 1)
            If randomArray(i) = randomArray(j) Then i = i - 1: Exit For
        Next j
    Next i
End Function

Sub OnSlideShowTerminate()                            ''关闭PPT时
    CommandButton3_Click
    Ex.Workbooks.Close    '关闭打开的Excel
    Set Ex = Nothing    '清空xlApp
    Ex.Quit
End Sub

Sub OnSlideShowPageChange()                           ''演示PPT时

    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 5 Then
        Set sld = Slide130
        rdmSeq = sld.chbRdmSeq.Value
        rdmTip = sld.chbRdmTip.Value
        rdmCount = sld.chbRdmCount.Value
        containWords = sld.chbWords.Value
        containPhrases = sld.chbPhrases.Value
        containSentences = sld.chbSentences.Value
        count = sld.txbCount.Text
        
        Set dic = CreateObject("Scripting.Dictionary")
         If containWords Then
            dic.Add 1, "IdiomaticWords"
         End If
         If containPhrases Then
            dic.Add 2, "NativePhrases"
         End If
         If containSentences Then
            dic.Add 3, "GoldenSentences"
         End If
         If containCultures Then
            dic.Add 4, "AmericanCultures"
         End If
                
        InitData dic
        
        
        'MsgBox Ex.Worksheets(1).Cells(1, 4).Value
        'Subject_Total = Val(Ex.Cells(1, 4).Value)             '获取总题数
        'MsgBox Subject_Total
       
        ReDim randomArray(dic.count + 1)
        Next_N = 0

        Call GetRandomArray(1, dic.count, randomArray())  '随机数组
         For i = 1 To dic.count + 1
             sld.lblQuestion.Caption = sld.lblQuestion.Caption + dicData(randomArray(i))
         Next i
        
    End If
End Sub


Sub Pause()                                     '''暂停
    
    If IsRunning = False Then
        MsgBox "请点击开始"
        Exit Sub
    End If

    IsRunning = False                          '关闭滚动数字
    Label1.Caption = G_Array(Next_N)

    If (Next_N + 1) >= Subject_Total Then
        IsRunning = False
        Label2.Caption = "题库所有的题目已经全部用完！"
    Else
        Label2.Caption = "第" & Str(G_Array(Next_N)) & " 题" & vbCrLf & Ex.Cells(G_Array(Next_N), 1).Value
    
        DaAN = Ex.Cells(G_Array(Next_N), 2).Value
        
        Next_N = Next_N + 1
    End If
End Sub

Sub Running()                             ''''运行
    Dim Num As Integer
    
    IsRunning = True

    If Next_N >= Subject_Total Then
        IsRunning = False
        Label2.Caption = "题库所有的题目已经全部用完！"
    End If
    
    'Label1.BackColor = RGB(0, 176, 240)
    'Label2.BackColor = RGB(0, 176, 240)
    Label5.Caption = ""
    Do While IsRunning = True
        Randomize
        Num = Int(Rnd * Subject_Total) + 1
        Label1.Caption = Str(Num)
        DoEvents
    Loop
End Sub

Sub Check()                             ''''检查答案
    'MsgBox "正确答案:" & DaAN
    Label5.Caption = "答案:" & DaAN
End Sub

Private Sub CommandButton3_Click()       ''''复位
    DaAN = ""
    Next_N = 0

    Label1.Caption = "GO!"
    Label2.Caption = ""
    Label5.Caption = ""
End Sub


Sub InitData(ByVal dic As Object)
    Set dicData = CreateObject("Scripting.Dictionary")
    Set Excel = CreateObject("Excel.Application")
        Excel.Workbooks.Open (ActivePresentation.Path & "\DataBase.xlsx")
        Excel.Visible = False
      n = Excel.Worksheets(1).Range("A65536").End(xlUp).Row
    For i = 1 To n
        cata = Excel.Worksheets(1).Cells(i, 3).Value
        If dic.Exists(cata) = True Then
            k = Excel.Worksheets(1).Cells(i, 1).Value
            v = Excel.Worksheets(1).Cells(i, 2).Value
            If dicData.Exists(k) = False Then
                dicData.Add k, v
            End If
        End If
    Next i
    MsgBox dicData.count
    'k = dic.keys
    'v = dic.Items
    'MsgBox dic(dic.keys(0))
End Sub

Sub InitCatagory()
    Set dicCatagory = CreateObject("Scripting.Dictionary")
    Set Excel = CreateObject("Excel.Application")
        Excel.Workbooks.Open (ActivePresentation.Path & "\DataBase.xlsx")
        Excel.Visible = False
      n = Excel.Worksheets(2).Range("A65536").End(xlUp).Row
    For i = 1 To n
        k = Excel.Worksheets(2).Cells(i, 1).Value
        v = Excel.Worksheets(2).Cells(i, 2).Value
        If dicCatagory.Exists(k) = False Then
            dicCatagory.Add k, v
        End If
    Next i
    Excel.Quit
End Sub
Private Sub CommandButton1_Click()
    'InitCatagory
     Set sld = Slide130
        rdmSeq = sld.chbRdmSeq.Value
        rdmTip = sld.chbRdmTip.Value
        rdmCount = sld.chbRdmCount.Value
        containWords = sld.chbWords.Value
        containPhrases = sld.chbPhrases.Value
        containSentences = sld.chbSentences.Value
        count = sld.txbCount.Text
        
        Set dic = CreateObject("Scripting.Dictionary")
         If containWords Then
            dic.Add 1, "Words"
         End If
         If containPhrases Then
            dic.Add 2, "Phrases"
         End If
         If containSentences Then
            dic.Add 3, "Sentences"
         End If
        InitData dic
    
        ReDim randomArray(dicData.count + 1)
        Next_N = 0

        Call GetRandomArray(1, dicData.count, randomArray())  '随机数组

      
         
         For i = 1 To dicData.count
             keys = dicData.keys
             Slide132.lblQuestion.Caption = Slide132.lblQuestion.Caption & dicData(keys(randomArray(i)))
         Next i
         MsgBox Slide132.lblQuestion.Caption
End Sub


