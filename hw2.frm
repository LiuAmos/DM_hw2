VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   ScaleHeight     =   15630
   ScaleWidth      =   28560
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
'操作方式
'須手動輸入檔名(yeast.txt)以及資料筆數(1484)
'看equalwidth的forward跟backward按鈕順序(read->Equal_width->forward or backward)
'看完equalwidth欲看equalfrequency請關掉再重新run一次
'看equalfrequency的forward跟backward的按鈕順序(read->Equal_frequency->forward or backward)
'entropy未實做成功,故按鈕無效用
Dim file As String
Dim datanum As Integer
Dim data2darray(9, 1483) As String
Dim ewresult(7, 10) As Double '8個屬性的11個邊界
Dim efresult(7, 10) As Double
Dim choicefb As Integer
Dim discreresult(7, 10) As Double
Dim totalh(8) As Double '每個屬性的h
Dim totalhab(8, 8) As Double
Dim attr9array(9) As String
Dim attr9counter(9) As Double
Dim totaluac(7) As Double
Dim totaluab(7, 7) As Double
Static Function subsetgoodness(ByRef subset() As Double)
Dim tempsubset() As Double
Dim totalgoodnessDenominator As Double
Dim totalgoodnessNumerator As Double
Dim totalgoodnessValue As Double
totalgoodnessDenominator = 0
totalgoodnessNumerator = 0
totalgoodnessValue = 0
tempsubset() = subset()

For i = 0 To UBound(tempsubset)
For j = 0 To UBound(tempsubset)
totalgoodnessDenominator = totalgoodnessDenominator + totaluab(tempsubset(i), tempsubset(j))
Next j
Next i
totalgoodnessDenominator = totalgoodnessDenominator ^ (1 / 2)

For i = 0 To UBound(tempsubset)
    totalgoodnessNumerator = totalgoodnessNumerator + totaluac(tempsubset(i))
Next i

totalgoodnessValue = (totalgoodnessNumerator / totalgoodnessDenominator)

subsetgoodness = totalgoodnessValue
End Function

Static Function gmax(ByRef goodarray() As Double, ByRef attrarray() As Double)
Dim tempgoodarray() As Double
Dim tempattrarray() As Double
Dim tempmax As Double
Dim attrmax As Double
attrmax = -100
tempmax = -100
tempgoodarray() = goodarray()
tempattrarray() = attrarray()

For i = 0 To UBound(tempgoodarray)
If tempmax < tempgoodarray(i) Then
tempmax = tempgoodarray(i)
attrmax = tempattrarray(i)

End If
Next i

gmax = CStr(attrmax) + "," + CStr(tempmax)
End Function

Static Function unpickAttr(ByRef resultattrs() As Double)
Dim tempresultattrs() As Double
Dim allattr(7) As Double
Dim unpickAttrs() As Double
Dim counter As Double
counter = 0
tempresultattrs() = resultattrs()

For i = 0 To 7
    If tempresultattrs(i) = -1 Then
        ReDim unpickAttrs(7 - i)
        Exit For
    End If
Next i

For i = 0 To 7
    allattr(i) = i
Next i

For i = 0 To 7
    If tempresultattrs(i) <> -1 Then
        allattr(tempresultattrs(i)) = 10
    End If
Next i

For i = 0 To 7
    If allattr(i) <> 10 Then
        unpickAttrs(counter) = allattr(i)
        counter = counter + 1
    End If
Next i

unpickAttr = unpickAttrs()

End Function

Static Function totalsetb()
Dim resultattr(7) As Double '記錄每次踢掉的那一個
Dim resultmaxvalue(7) As Double '每個set數量的最大值
Dim deleteone() As Double
Dim tempgoodness() As Double
Dim tempattr() As Double
Dim counter As Double

For i = 0 To 7
resultattr(i) = -1
resultmaxvalue(i) = -1
Next i

For i = 0 To 7
'先決定要丟入subsetgoodness的陣列長度

ReDim deleteone(6 - i)
Dim setnum() As Double '準備要扣掉一個attr前的完整版
ReDim tempgoodness(7 - i)
ReDim tempattr(7 - i) '候選被踢的attr




'接著為該陣列塞入目前已選的attr

'找出要留下來的attr
setnum() = unpickAttr(resultattr)




'每個attr都刪看看
    For j = 0 To UBound(setnum)
        If j = 0 Then
            GoTo jzero
        End If
        setnum(j - 1) = tempattr(j - 1)
jzero:
        tempattr(j) = setnum(j)
        setnum(j) = -1 '幫setnum刪掉某一個attr
        counter = 0
        For k = 0 To UBound(setnum)
            If setnum(k) <> -1 Then
                deleteone(counter) = setnum(k)
                counter = counter + 1
            End If
        Next k
         '存刪掉的那一個的attr
        tempgoodness(j) = subsetgoodness(deleteone)
    Next j
tempgmaxstr = gmax(tempgoodness, tempattr)
tempav = Split(tempgmaxstr, ",")
resultattr(i) = CDbl(tempav(0))
resultmaxvalue(i) = CDbl(tempav(1))



If i > 0 Then
If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
Exit For
End If
End If

Next i


Dim inputarr(7) As Double
For i = 0 To 5

    Dim printresult() As Double
    Dim attrstr As String
    attrstr = ""
    For j = 0 To 7
    inputarr(j) = -1
    Next j
    
    For j = 0 To i
    inputarr(j) = resultattr(j)
    Next j
    printresult() = unpickAttr(inputarr)
    For j = 0 To UBound(printresult)
    attrstr = attrstr + CStr(printresult(j) + 1)
    Next j
    List1.AddItem "attribute:" + attrstr
    List1.AddItem resultmaxvalue(i)
    If i = 0 Then
    GoTo nozero
    End If
    
    If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
    Exit For
    End If
    
nozero:
Next i

'動態陣列測試
'Dim testnum() As Double
'For i = 0 To 3
'ReDim testnum(i)
'List1.AddItem UBound(testnum)
'Next i

totalsetb = "End"
End Function

Static Function totalset()


Dim resultattr(7) As Double '每個set的最大值新選的那個attr
Dim resultmaxvalue(7) As Double '每個set數量的最大值
Dim setnum() As Double

Dim tempgoodness() As Double
Dim tempattr() As Double

For i = 0 To 7
resultattr(i) = -1
resultmaxvalue(i) = -1
Next i

For i = 0 To 7
'先決定要丟入subsetgoodness的陣列長度

ReDim setnum(i) '丟入subsetgoodness的陣列
Dim tempunpickAttr() As Double
ReDim tempgoodness(7 - i)
ReDim tempattr(7 - i)



'接著為該陣列塞入目前已選的attr

If i = 0 Then
    GoTo izero
End If
'把已經選好的resultattr丟給setnum
   For j = 0 To (UBound(setnum) - 1)
       setnum(j) = resultattr(j)
   Next j
izero:
'跑unpickAttr()回傳還未被選的attr陣列
tempunpickAttr() = unpickAttr(resultattr)


'跑回圈幫setnum(i)塞入不同還未被選的attr
    For j = 0 To UBound(tempunpickAttr)
        setnum(i) = tempunpickAttr(j) '幫setnum挑一個還沒進來的attr
        tempattr(j) = tempunpickAttr(j)
        tempgoodness(j) = subsetgoodness(setnum)
    Next j
tempgmaxstr = gmax(tempgoodness, tempattr)
tempav = Split(tempgmaxstr, ",")
resultattr(i) = CDbl(tempav(0))
resultmaxvalue(i) = CDbl(tempav(1))



If i > 0 Then
If resultmaxvalue(i) < resultmaxvalue(i - 1) Then
Exit For
End If
End If

Next i



List1.AddItem "attribute:" + CStr(resultattr(0) + 1)
List1.AddItem resultmaxvalue(0)

List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1)

List1.AddItem resultmaxvalue(1)

List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1)
List1.AddItem resultmaxvalue(2)

List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1) + CStr(resultattr(3) + 1)
List1.AddItem resultmaxvalue(3)

List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1) + CStr(resultattr(3) + 1) + CStr(resultattr(4) + 1)
List1.AddItem resultmaxvalue(4)
If choicefb = 0 Then
GoTo choicefbzero
End If
List1.AddItem "attribute:" + CStr(resultattr(0) + 1) + CStr(resultattr(1) + 1) + CStr(resultattr(2) + 1) + CStr(resultattr(3) + 1) + CStr(resultattr(4) + 1) + CStr(resultattr(5) + 1)
List1.AddItem resultmaxvalue(5)
choicefbzero:



'動態陣列測試
'Dim testnum() As Double
'For i = 0 To 3
'ReDim testnum(i)
'List1.AddItem UBound(testnum)
'Next i

totalset = "End"
End Function

Static Function uab(ByVal num1 As Integer, ByVal num2 As Integer)
Dim uabresult As Double
Dim ha As Double
Dim hb As Double
Dim hab As Double
Dim eachinterval As String
Dim HAB2darray() As Double


ha = totalh(num1)
hb = totalh(num2)
hab = totalhab(num1, num2)


If ha = 0 And hb = 0 Then
uabresult = 1
GoTo uzero
End If

uabresult = 2 * ((ha + hb - hab) / (ha + hb))

uzero:
uab = uabresult
    
End Function

Static Function h(ByVal totalinterval As String, ByVal num As Integer)

Dim totalstring() As String
Dim eachstring() As String
Dim triminterval As String
Dim log As Double
Dim prop As Double
Dim eachh As Double
Dim temp As Double
eachh = 0

triminterval = Trim(totalinterval)
totalstring() = Split(triminterval, " ")



eachstring() = Split(totalstring(num), ",")

For i = 0 To 9
    prop = CDbl(eachstring(i)) / 1484
    temp = -(prop * Log02(prop))
    eachh = eachh + temp
Next i

h = eachh

End Function

Static Function htwo(ByRef counter2darray() As Double)
Dim temp As Double
Dim totalhtwo As Double
Dim tmpcounter2darray() As Double
tmpcounter2darray() = counter2darray()


totalhtwo = 0
For i = 0 To 9
    For j = 0 To 9
        temp = -(tmpcounter2darray(i, j) * Log02(tmpcounter2darray(i, j)))
        totalhtwo = totalhtwo + temp
    Next j
Next i

htwo = totalhtwo
End Function
Static Function countac(ByRef result() As Double, ByVal num1 As Integer, ByRef attr9() As String)
Dim num1temp As Variant
Dim attr9temp As String
Dim num1result(10) As Double
Dim attr9result(9) As String 'attr9的各項名稱
Dim counter2darray(9, 9) As Double

For i = 0 To 9
    For j = 0 To 9
        counter2darray(i, j) = 0
    Next j
Next i

For i = 1 To 9
num1result(i) = result(num1, i - 1)
attr9result(i - 1) = attr9(i - 1)
Next i
num1result(10) = result(num1, 10)
num1result(0) = result(num1, 9)
attr9result(9) = attr9(9)

For i = 0 To 1483
    num1temp = Val(data2darray(num1 + 1, i))
    attr9temp = data2darray(9, i)
    For j = 0 To 8
        If num1temp >= Val(num1result(j)) And num1temp < Val(num1result(j + 1)) Then '注意邊界條件
            For k = 0 To 9
                If attr9temp = attr9result(k) Then
                    counter2darray(j, k) = counter2darray(j, k) + 1
                End If
            Next k
        End If
    Next j


    If num1temp >= Val(num1result(9)) Then
        For k = 0 To 9
            If attr9temp = attr9result(k) Then
                counter2darray(9, k) = counter2darray(9, k) + 1
            End If
        Next k
    End If
    
Next i

For i = 0 To 9
    For j = 0 To 9
        counter2darray(i, j) = counter2darray(i, j) / 1484
    Next j
Next i

countac = counter2darray
End Function
Static Function CountTwoAttr(ByRef result() As Double, ByVal num1 As Integer, ByVal num2 As Integer)

Dim num1temp As Variant
Dim num2temp As Variant
Dim num1result(10) As Double
Dim num2result(10) As Double
Dim counter2darray(9, 9) As Double
Dim testcounter As Double


For i = 0 To 9
For j = 0 To 9
counter2darray(i, j) = 0
Next j
Next i


For i = 1 To 9
num1result(i) = result(num1, i - 1)
num2result(i) = result(num2, i - 1)
Next i
num1result(10) = result(num1, 10)
num1result(0) = result(num1, 9)
num2result(10) = result(num2, 10)
num2result(0) = result(num2, 9)


For i = 0 To 1483
    num1temp = Val(data2darray(num1 + 1, i))
    num2temp = Val(data2darray(num2 + 1, i))
    For j = 0 To 9
        If num1temp >= Val(num1result(j)) And num1temp < Val(num1result(j + 1)) Then '注意邊界條件
            For k = 0 To 9
                If num2temp >= Val(num2result(k)) And num2temp < Val(num2result(k + 1)) Then
                    counter2darray(j, k) = counter2darray(j, k) + 1
                End If
            Next k
        End If
    Next j
    
    If num1temp = Val(num1result(10)) And num2temp = Val(num2result(10)) Then
        counter2darray(9, 9) = counter2darray(9, 9) + 1
    End If
    
    If num1temp = Val(num1result(10)) And num2temp <> Val(num2result(10)) Then
        For k = 0 To 9
            If num2temp >= Val(num2result(k)) And num2temp < Val(num2result(k + 1)) Then
                counter2darray(9, k) = counter2darray(9, k) + 1
            End If
        Next k
    End If
    
    If num1temp <> Val(num1result(10)) And num2temp = Val(num2result(10)) Then
        For k = 0 To 9
            If num1temp >= Val(num1result(k)) And num1temp < Val(num1result(k + 1)) Then
                counter2darray(k, 9) = counter2darray(k, 9) + 1
            End If
        Next k
    End If
    
Next i




For i = 0 To 9
    For j = 0 To 9
        counter2darray(i, j) = counter2darray(i, j) / 1484
    Next j
Next i

CountTwoAttr = counter2darray

End Function

Static Function CountEachAttr(ByRef result() As Double)
Dim tempresult(7, 10) As Double
Dim counter(10) As Double
Dim interval As String
Dim totalinterval As String
Dim temp As Variant



For i = 0 To 7
    For j = 0 To 10
        tempresult(i, j) = result(i, j)
    Next j
Next i

For k = 1 To 8
For i = 0 To 10
counter(i) = 0
Next i



For i = 0 To 1483
    temp = Val(data2darray(k, i))


    
    For j = 0 To 7

        If temp >= Val(tempresult(k - 1, j)) And temp < Val(tempresult(k - 1, j + 1)) Then
        counter(j + 2) = counter(j + 2) + 1
        Exit For
        End If
    Next j
    
    If temp >= Val(tempresult(k - 1, 8)) Then
        counter(10) = counter(10) + 1
    End If
    If temp < Val(tempresult(k - 1, 0)) And temp >= Val(tempresult(k - 1, 9)) Then
        counter(1) = counter(1) + 1
    End If
Next i


interval = CStr(counter(1)) + "," + CStr(counter(2)) + "," + CStr(counter(3)) + "," + CStr(counter(4)) + "," + CStr(counter(5)) + "," + CStr(counter(6)) + "," + CStr(counter(7)) + "," + CStr(counter(8)) + "," + CStr(counter(9)) + "," + CStr(counter(10))

totalinterval = totalinterval + interval + " "

Next k

CountEachAttr = totalinterval 'return

End Function

Static Function Log02(ByVal x As Double)
If x = 0 Then
Log02 = 0
Else
Log02 = log(x) / log(2) ' return
End If
End Function


Private Sub backward_Click()

List1.Clear
Dim eachinterval As String
Dim eachh As Double
Dim num As Integer
Dim testhtwo As Double
Dim HAB2darray() As Double
Dim testuab As Double
Dim attr9prop(9) As Double
Dim attr9temp As String
Dim attr9str As String 'attr9的各項數目統計
Dim ccountac() As Double
Dim hac As Double
Dim tempgoodness As String
Dim aaa() As String
Dim testtotalsub As String




If choicefb = 0 Then
For i = 0 To 7
For j = 0 To 10
discreresult(i, j) = ewresult(i, j)
Next j
Next i
ElseIf choicefb = 1 Then
For i = 0 To 7
For j = 0 To 10
discreresult(i, j) = efresult(i, j)
Next j
Next i
Else
'entropy
End If


attr9array(0) = "CYT"
attr9array(1) = "NUC"
attr9array(2) = "MIT"
attr9array(3) = "ME3"
attr9array(4) = "ME2"
attr9array(5) = "ME1"
attr9array(6) = "EXC"
attr9array(7) = "VAC"
attr9array(8) = "POX"
attr9array(9) = "ERL"
For j = 0 To 9
    attr9counter(j) = 0
Next j

For i = 0 To 1483
    attr9temp = data2darray(9, i)
    For j = 0 To 9
        If attr9temp = attr9array(j) Then
            attr9counter(j) = attr9counter(j) + 1
        End If
    Next j
Next i

attr9str = CStr(attr9counter(0)) + "," + CStr(attr9counter(1)) + "," + CStr(attr9counter(2)) + "," + CStr(attr9counter(3)) + "," + CStr(attr9counter(4)) + "," + CStr(attr9counter(5)) + "," + CStr(attr9counter(6)) + "," + CStr(attr9counter(7)) + "," + CStr(attr9counter(8)) + "," + CStr(attr9counter(9)) + " " + "a"
totalh(8) = h(attr9str, 0)
'List1.AddItem totalh(8)
'-------------------------
eachinterval = CountEachAttr(discreresult)


For i = 0 To 7
num = i
totalh(i) = h(eachinterval, num)
Next i


'每個attr的h值
'For i = 0 To 8
'List1.AddItem totalh(i)
'Next i

'testhtwo = htwo(eachinterval, 0, 1)
'List1.AddItem testhtwo

'For i = 0 To 9
'    For j = 0 To 9
'        List1.AddItem HAB2darray(i, j)
'    Next j
'Next i
For i = 0 To 7
For j = 0 To 7
HAB2darray = CountTwoAttr(discreresult, i, j)
totalhab(i, j) = htwo(HAB2darray)
Next j
Next i



'List1.AddItem HAB2darray(1, 4)
'List1.AddItem HAB2darray(4, 1)


For i = 0 To 7
ccountac = countac(discreresult, i, attr9array)
totalhab(i, 8) = htwo(ccountac)
totalhab(8, i) = htwo(ccountac)
'List1.AddItem ""
Next i
totalhab(8, 8) = totalh(8)


'印出H跟HAB值
'For i = 0 To 8
'List1.AddItem totalh(i)
'Next i
'List1.AddItem ""
'
'For i = 0 To 7
'For j = i To 7
'List1.AddItem totalhab(i, j)
'Next j
'List1.AddItem ""
'Next i

'看h1-8跟9的hab
'For i = 0 To 8
'List1.AddItem totalhab(8, i)
'Next i


'List1.AddItem totalh(7)
'List1.AddItem totalhab(7, 7)
'List1.AddItem totalhab(7, 4)


For i = 0 To 7
totaluac(i) = uab(i, 8)
'List1.AddItem totaluac(i)
Next i

For i = 0 To 7
    For j = 0 To 7
        totaluab(i, j) = uab(i, j)
    Next j
Next i
'List1.AddItem totaluac(0)
'List1.AddItem totaluac(2)
'List1.AddItem totaluab(1, 6)
'List1.AddItem totaluab(6, 1)
'List1.AddItem ""

'List1.AddItem totaluab(2, 5)
'List1.AddItem totaluab(5, 2)
'List1.AddItem totaluac(2)

'看uab和uac
'For i = 0 To 7
'For j = i To 7
'List1.AddItem totaluab(i, j)
'Next j
'List1.AddItem totaluac(i)
'Next i


'測試goodness值
'Dim goodnessresult As Double
'Dim testsubsetgoodness(3) As Double
'goodnessresult = 0
'testsubsetgoodness(0) = 0
'testsubsetgoodness(1) = 1
'testsubsetgoodness(2) = 2
'testsubsetgoodness(3) = 3
'goodnessresult = subsetgoodness(testsubsetgoodness)
'List1.AddItem goodnessresult

'測試gmax
'Dim testgmax(5) As Double
'Dim testgmaxnum(5) As Double
'Dim testgmaxresult As String
'testgmaxresult = ""
'testgmaxnum(0) = 1
'testgmaxnum(1) = 2
'testgmaxnum(2) = 3
'testgmaxnum(3) = 4
'testgmaxnum(4) = 5
'testgmaxnum(5) = 6
'testgmax(0) = 0.55543
'testgmax(1) = 0.44444
'testgmax(2) = 0.222222
'testgmax(3) = 0.773345
'testgmax(4) = 0.557738
'testgmax(5) = 0.61344
'testgmaxresult = gmax(testgmax, testgmaxnum)
'List1.AddItem testgmaxresult

'測試unpickAttr
'Dim result() As Double
'Dim inputattr(7) As Double
'For i = 0 To 7
'    inputattr(i) = -1
'Next i
'inputattr(0) = 5
'inputattr(1) = 2
'inputattr(2) = 0
'result() = unpickAttr(inputattr)
'List1.AddItem UBound(result)
'List1.AddItem ""
'For i = 0 To UBound(result)
'    List1.AddItem result(i)
'Next i



testtotalsub = totalsetb()
List1.AddItem testtotalsub

'GoTo endd
'endd:
End Sub

Private Sub datanumber_Change()
datanum = CInt(datanumber.Text)
End Sub

Private Sub datatxt_Change()
file = datatxt.Text
End Sub


Private Sub Equal_frequency_Click()
List1.Clear
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
efresult(j - 1, 9) = widtharray(0)
efresult(j - 1, 10) = widtharray(1483)

Next j

'測試區間
'For j = 0 To 7
'For i = 0 To 10
'List1.AddItem efresult(j, i)
'Next i
'List1.AddItem ""
'Next j


End Sub

Private Sub Equal_width_Click()
List1.Clear
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
        'If widtharray(i) <> 0 Then
            If min > widtharray(i) Then
                min = widtharray(i)
            End If
            If max < widtharray(i) Then
                max = widtharray(i)
            End If
        'End If
    End If
Next i
ewresult(j - 1, 9) = min
ewresult(j - 1, 10) = max
range = (max - min) / 10
For i = 0 To 9
List1.AddItem "Interval:" + CStr(i + 1) + "     " + CStr(min) + "~" + CStr(min + range)
min = min + range
If i < 9 Then
ewresult(j - 1, i) = min
End If

Next i

Next j

'debug
'For i = 0 To 9
'    For j = 0 To 1483
'        If data2darray(i, j) = "" Then
'        List1.AddItem i
'        List1.AddItem j
'        End If
'        'List1.AddItem data2darray(i, j)
'    Next j
'    List1.AddItem "-------------------------"
'Next i

End Sub

Private Sub forward_Click()
'Dim discreresult(7, 10) As Double
List1.Clear
Dim eachinterval As String
Dim eachh As Double
Dim num As Integer
Dim testhtwo As Double
Dim HAB2darray() As Double
Dim testuab As Double
Dim attr9prop(9) As Double
Dim attr9temp As String
Dim attr9str As String 'attr9的各項數目統計
Dim ccountac() As Double
Dim hac As Double
Dim tempgoodness As String
Dim aaa() As String
Dim testtotalsub As String




If choicefb = 0 Then
For i = 0 To 7
For j = 0 To 10
discreresult(i, j) = ewresult(i, j)
Next j
Next i
ElseIf choicefb = 1 Then
For i = 0 To 7
For j = 0 To 10
discreresult(i, j) = efresult(i, j)
Next j
Next i
Else
'entropy
End If



'-------------------------
attr9array(0) = "CYT"
attr9array(1) = "NUC"
attr9array(2) = "MIT"
attr9array(3) = "ME3"
attr9array(4) = "ME2"
attr9array(5) = "ME1"
attr9array(6) = "EXC"
attr9array(7) = "VAC"
attr9array(8) = "POX"
attr9array(9) = "ERL"
For j = 0 To 9
    attr9counter(j) = 0
Next j

For i = 0 To 1483
    attr9temp = data2darray(9, i)
    For j = 0 To 9
        If attr9temp = attr9array(j) Then
            attr9counter(j) = attr9counter(j) + 1
        End If
    Next j
Next i

attr9str = CStr(attr9counter(0)) + "," + CStr(attr9counter(1)) + "," + CStr(attr9counter(2)) + "," + CStr(attr9counter(3)) + "," + CStr(attr9counter(4)) + "," + CStr(attr9counter(5)) + "," + CStr(attr9counter(6)) + "," + CStr(attr9counter(7)) + "," + CStr(attr9counter(8)) + "," + CStr(attr9counter(9)) + " " + "a"
totalh(8) = h(attr9str, 0)
'List1.AddItem totalh(8)
'-------------------------
eachinterval = CountEachAttr(discreresult)


For i = 0 To 7
num = i
totalh(i) = h(eachinterval, num)
'List1.AddItem ""
Next i




'每個attr的h值
'For i = 0 To 8
'List1.AddItem totalh(i)
'Next i

'testhtwo = htwo(eachinterval, 0, 1)
'List1.AddItem testhtwo

'For i = 0 To 9
'    For j = 0 To 9
'        List1.AddItem HAB2darray(i, j)
'    Next j
'Next i
For i = 0 To 7
For j = 0 To 7
HAB2darray = CountTwoAttr(discreresult, i, j)
totalhab(i, j) = htwo(HAB2darray)
Next j
Next i



'List1.AddItem HAB2darray(1, 4)
'List1.AddItem HAB2darray(4, 1)


For i = 0 To 7
ccountac = countac(discreresult, i, attr9array)
totalhab(i, 8) = htwo(ccountac)
totalhab(8, i) = htwo(ccountac)
'List1.AddItem ""
Next i
totalhab(8, 8) = totalh(8)

'印出H跟HAB值
'For i = 0 To 8
'List1.AddItem totalh(i)
'Next i
'List1.AddItem ""
'
'For i = 0 To 7
'For j = i To 7
'List1.AddItem totalhab(i, j)
'Next j
'List1.AddItem ""
'Next i


'看h1-8跟9的hab
'For i = 0 To 8
'List1.AddItem totalhab(8, i)
'Next i


'List1.AddItem totalh(7)
'List1.AddItem totalhab(7, 7)
'List1.AddItem totalhab(7, 4)




'List1.AddItem totalhab(2, 7)
'List1.AddItem totalhab(7, 2)


'List1.AddItem totalhab(6, 1)
'GoTo haabb
'List1.AddItem totalhab(6, 6)
'List1.AddItem testhtwo
'List1.AddItem ""
For i = 0 To 7
totaluac(i) = uab(i, 8)
'List1.AddItem totaluac(i)
Next i


For i = 0 To 7
    For j = 0 To 7
        totaluab(i, j) = uab(i, j)
    Next j
Next i


'List1.AddItem totaluac(0)
'List1.AddItem totaluac(2)
'List1.AddItem totaluab(1, 6)
'List1.AddItem totaluab(6, 1)
'List1.AddItem ""
'List1.AddItem totaluab(4, 4)
'List1.AddItem totaluab(6, 6)
'List1.AddItem totaluab(4, 5)
'List1.AddItem totaluac(2)

'看uab和uac
'For i = 0 To 7
'For j = i To 7
'List1.AddItem totaluab(i, j)
'Next j
'List1.AddItem totaluac(i)
'List1.AddItem ""
'Next i


'測試goodness值
'Dim goodnessresult As Double
'Dim testsubsetgoodness(4) As Double
'goodnessresult = 0
'testsubsetgoodness(0) = 0
'testsubsetgoodness(1) = 1
'testsubsetgoodness(2) = 2
'testsubsetgoodness(3) = 3
'testsubsetgoodness(4) = 7
'''testsubsetgoodness(5) = 7
'''testsubsetgoodness(6) = 7
'goodnessresult = subsetgoodness(testsubsetgoodness)
'List1.AddItem goodnessresult

'測試gmax
'Dim testgmax(5) As Double
'Dim testgmaxnum(5) As Double
'Dim testgmaxresult As String
'testgmaxresult = ""
'testgmaxnum(0) = 1
'testgmaxnum(1) = 2
'testgmaxnum(2) = 3
'testgmaxnum(3) = 4
'testgmaxnum(4) = 5
'testgmaxnum(5) = 6
'testgmax(0) = 0.55543
'testgmax(1) = 0.44444
'testgmax(2) = 0.222222
'testgmax(3) = 0.773345
'testgmax(4) = 0.557738
'testgmax(5) = 0.61344
'testgmaxresult = gmax(testgmax, testgmaxnum)
'List1.AddItem testgmaxresult

'測試unpickAttr
'Dim result() As Double
'Dim inputattr(7) As Double
'For i = 0 To 7
'    inputattr(i) = -1
'Next i
'inputattr(0) = 5
'inputattr(1) = 2
'inputattr(2) = 0
'result() = unpickAttr(inputattr)
'List1.AddItem UBound(result)
'List1.AddItem ""
'For i = 0 To UBound(result)
'    List1.AddItem result(i)
'Next i



testtotalsub = totalset()
List1.AddItem testtotalsub

'GoTo endd
'endd:
End Sub

Private Sub List1_Click()

End Sub

Private Sub read_Click()
List1.Clear
Dim datacounter As Integer
Dim temp() As String

datacounter = 0

Open App.Path & "\" + file For Input As #1
Do While Not EOF(1) And datacounter < datanum

Line Input #1, tmpline

tmpline = Replace(tmpline, "  ", " ")
tmpline = Replace(tmpline, "  ", " ")
List1.AddItem tmpline
temp = Split(tmpline, " ")
'List1.AddItem temp

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
