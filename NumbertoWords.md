Yes. Use this **VBA custom Excel formula** to convert a number into **Crores / Lakhs / Thousands / Hundreds** text.

```vb
Function CroresText(ByVal Amount As Variant) As String
    Dim s As String
    Dim n As Variant, remAmt As Variant, crore As Variant
    Dim lakh As Long, thousand As Long, hundred As Long, rest As Long
    Dim parts As Collection
    Dim i As Long, txt As String, signText As String
    
    On Error GoTo InvalidInput
    
    s = Trim(CStr(Amount))
    s = Replace(s, ",", "")
    s = Replace(s, "₹", "")
    s = Replace(s, " ", "")
    
    If s = "" Then
        CroresText = ""
        Exit Function
    End If
    
    n = CDec(s)
    n = Fix(n)
    
    If n < 0 Then
        signText = "Minus "
        n = Abs(n)
    End If
    
    crore = Fix(n / CDec(10000000))
    remAmt = n - crore * CDec(10000000)
    
    lakh = CLng(Fix(remAmt / CDec(100000)))
    remAmt = remAmt - CDec(lakh) * CDec(100000)
    
    thousand = CLng(Fix(remAmt / CDec(1000)))
    remAmt = remAmt - CDec(thousand) * CDec(1000)
    
    hundred = CLng(Fix(remAmt / CDec(100)))
    rest = CLng(remAmt - CDec(hundred) * CDec(100))
    
    Set parts = New Collection
    
    If crore > 0 Then parts.Add CStr(crore) & IIf(crore = 1, " Crore", " Crores")
    If lakh > 0 Then parts.Add CStr(lakh) & IIf(lakh = 1, " Lakh", " Lakhs")
    If thousand > 0 Then parts.Add CStr(thousand) & " Thousand"
    If hundred > 0 Then parts.Add CStr(hundred) & " Hundred"
    If rest > 0 Then parts.Add CStr(rest)
    
    If parts.Count = 0 Then
        CroresText = "Zero Only"
        Exit Function
    End If
    
    For i = 1 To parts.Count
        If i = 1 Then
            txt = parts(i)
        ElseIf i = parts.Count Then
            txt = txt & " and " & parts(i)
        Else
            txt = txt & " " & parts(i)
        End If
    Next i
    
    CroresText = signText & txt & " Only"
    Exit Function
    
InvalidInput:
    CroresText = "Invalid Amount"
End Function
```

Use it in Excel like this:

```excel
=CroresText(A1)
```

Example, if `A1` contains:

```text
100,51,23,400
```

The result will be:

```text
100 Crores 51 Lakhs 23 Thousand and 4 Hundred Only
```

To add it: press **Alt + F11**, go to **Insert > Module**, paste the code, save the workbook as **.xlsm**, then use `=CroresText(A1)` in any cell.
