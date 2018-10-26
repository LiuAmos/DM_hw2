VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   18555
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox datanumber 
      Height          =   495
      Left            =   14160
      TabIndex        =   8
      Text            =   "1484"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton backward 
      Caption         =   "backward"
      Height          =   735
      Left            =   15960
      TabIndex        =   7
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton forward 
      Caption         =   "forward"
      Height          =   735
      Left            =   13560
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Entropy_based 
      Caption         =   "Entropy_based"
      Height          =   615
      Left            =   16680
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Equal_frequency 
      Caption         =   "Equal_frequency"
      Height          =   615
      Left            =   14880
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Equal_width 
      Caption         =   "Equal_width"
      Height          =   615
      Left            =   13080
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton read 
      Caption         =   "read"
      Height          =   615
      Left            =   14760
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   7440
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   12495
   End
   Begin VB.TextBox datatxt 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "yeast.txt"
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file As String
Dim datanum As Integer
Dim data2darray(9, 1483) As String
Dim ewresult(7, 8) As Double
Dim efresult(7, 8) As Double
Dim choicefb As Integer

Static Function h(ByRef result() As Double)
Dim tempresult(7, 8) As Double
Dim i1, i2, i3, i4, i5, i6, i7, i8, i9, i10 As Integer
Dim interval As String
Dim tempdata2darray As Double
i1 = 0
i2 = 0
i3 = 0
i4 = 0
i5 = 0
i6 = 0
i7 = 0
i8 = 0
i9 = 0
i10 = 0

For i = 0 To 7
    For j = 0 To 8
        tempresult(i, j) = result(i, j)
    Next j
Next i

For i = 0 To 1483
tempdata2darray = CDbl(data2darray(1, i))
If (tempdata2darray < tempresult(0, 0)) Then
    i1 = i1 + 1
    ElseIf (tempdata2darray >= tempresult(0, 0) And tempdata2darray < tempresult(0, 1)) Then
        i2 = i2 + 1
        ElseIf (tempdata2darray >= tempresult(0, 1) And tempdata2darray < tempresult(0, 2)) Then
            i3 = i3 + 1
            ElseIf (tempdata2darray >= tempresult(0, 2) And tempdata2darray < tempresult(0, 3)) Then
                i4 = i4 + 1
                ElseIf (tempdata2darray >= tempresult(0, 3) And tempdata2darray < tempresult(0, 4)) Then
                    i5 = i5 + 1
                    ElseIf (tempdata2darray >= tempresult(0, 4) And tempdata2darray < tempresult(0, 5)) Then
                        i6 = i6 + 1
                        ElseIf (tempdata2darray >= tempresult(0, 5) And tempdata2darray < tempresult(0, 6)) Then
                            i7 = i7 + 1
                            ElseIf (tempdata2darray >= tempresult(0, 6) And tempdata2darray < tempresult(0, 7)) Then
                                i8 = i8 + 1
                                ElseIf (tempdata2darray >= tempresult(0, 7) And tempdata2darray < tempresult(0, 8)) Then
                                    i9 = i9 + 1
Else
    i10 = i10 + 1
End If
Next i
interval = CStr(i1) + "," + CStr(i2) + "," + CStr(i3) + "," + CStr(i4) + "," + CStr(i5) + "," + CStr(i6) + "," + CStr(i7) + "," + CStr(i8) + "," + CStr(i9) + "," + CStr(i10)
h = interval 'return

End Function

Static Function Log02(ByVal x As Double)
   Log02 = Log(x) / Log(2#) ' return
End Function


Private Sub datanumber_Change()
datanum = CInt(datanumber.Text)
End Sub

Private Sub datatxt_Change()
file = datatxt.Text
End Sub


Private Sub Equal_frequency_Click()
choicefb = 1
Dim widtharray(1483) As Double
Dim temp As Double
Dim range As Integer

For j = 1 To 8
range = 0
List1.AddItem ""
List1.AddItem "Print attribute  " + CStr(j) + "  interval"
For i = 0 To 1483
    If data2darray(j, i) <> "" Then
        widtharray(i) = CDbl(data2darray(j, i))
        If widtharray(i) <> 0 Then
        End If
    End If
Next i

For k = 0 To 1483
    For m = k To 1483
        If widtharray(k) > widtharray(m) Then
            tmp = widtharray(k)
            widtharray(k) = widtharray(m)
            widtharray(m) = tmp
        End If
    Next m
Next k


For i = 0 To 9
List1.AddItem "Interval:" + CStr(i + 1) + "     " + CStr(widtharray(range)) + "~" + CStr(widtharray(range + 148))
range = range + 148

If range = 148 Then
range = 147
End If

If range > 1330 Then
efresult(j - 1, 8) = widtharray(range)
List1.AddItem "Interval:" + CStr(i + 2) + "     " + CStr(widtharray(range)) + "~" + CStr(widtharray(1483))
Exit For
End If
If i < 9 Then
efresult(j - 1, i) = widtharray(range)
End If
Next i

Next j


End Sub

Private Sub Equal_width_Click()
choicefb = 0
Dim widtharray(1483) As Double
Dim max As Double
Dim min As Double
Dim range As Double

For j = 1 To 8
max = -100
min = 100
List1.AddItem ""
List1.AddItem "Print attribute  " + CStr(j) + "  interval"
For i = 0 To 1483
    If data2darray(j, i) <> "" Then
        widtharray(i) = CDbl(data2darray(j, i))
        If widtharray(i) <> 0 Then
            If min > widtharray(i) Then
                min = widtharray(i)
            End If
            If max < widtharray(i) Then
                max = widtharray(i)
            End If
        End If
    End If
Next i
range = (max - min) / 10
For i = 0 To 9
List1.AddItem "Interval:" + CStr(i + 1) + "     " + CStr(min) + "~" + CStr(min + range)
min = min + range
If i < 9 Then
ewresult(j - 1, i) = min
End If

Next i

Next j

End Sub

Private Sub forward_Click()
Dim discreresult(7, 8) As Double
Dim eachinterval As String

If choicefb = 0 Then
For i = 0 To 7
For j = 0 To 8
discreresult(i, j) = ewresult(i, j)
Next j
Next i
ElseIf choicefb = 1 Then
For i = 0 To 7
For j = 0 To 8
discreresult(i, j) = efresult(i, j)
Next j
Next i
Else
'entropy
End If

For i = 0 To 7
List1.AddItem ""
For j = 0 To 8
List1.AddItem discreresult(i, j)
Next j
Next i

eachinterval = h(discreresult)
List1.AddItem eachinterval

End Sub

Private Sub read_Click()
Dim datacounter As Integer
Dim temp() As String

datacounter = 0

Open App.Path & "\" + file For Input As #1
Do While Not EOF(1) And datacounter < datanum

Line Input #1, tmpline
List1.AddItem tmpline
temp = Split(tmpline, "  ")

For i = 0 To 9
    data2darray(i, datacounter) = Trim(temp(i))
Next i
datacounter = datacounter + 1
Loop
Close #1

'List1.AddItem data2darray(3, 4)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub
