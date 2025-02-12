Attribute VB_Name = "JSON"
' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

Option Explicit

Const INVALID_JSON      As Long = 1
Const INVALID_OBJECT    As Long = 2
Const INVALID_ARRAY     As Long = 3
Const INVALID_BOOLEAN   As Long = 4
Const INVALID_NULL      As Long = 5
Const INVALID_KEY       As Long = 6
Const INVALID_RPC_CALL  As Long = 7

Private psErrors As String

Public Function GetParserErrors() As String
   GetParserErrors = psErrors
End Function

Public Function ClearParserErrors() As String
   psErrors = ""
End Function


'
'   parse string and create JSON object
'

Public Function parse(ByRef str As String) As Object

   Dim Index As Long
   Index = 1
   psErrors = ""
   On Error Resume Next
   Call skipChar(str, Index)
   Select Case Mid(str, Index, 1)
      Case "{"
         Set parse = parseObject(str, Index)
      Case "["
         Set parse = parseArray(str, Index)
      Case Else
         psErrors = "Invalid JSON"
   End Select


End Function

 '
 '   parse collection of key/value
 '
Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary

   Set parseObject = New Dictionary
   Dim sKey As String
   
   ' "{"
   Call skipChar(str, Index)
   If Mid(str, Index, 1) <> "{" Then
      psErrors = psErrors & "Invalid Object at position " & Index & " : " & Mid(str, Index) & vbCrLf
      Exit Function
   End If
   
   Index = Index + 1

   Do
      Call skipChar(str, Index)
      If "}" = Mid(str, Index, 1) Then
         Index = Index + 1
         Exit Do
      ElseIf "," = Mid(str, Index, 1) Then
         Index = Index + 1
         Call skipChar(str, Index)
      ElseIf Index > Len(str) Then
         psErrors = psErrors & "Missing '}': " & Right(str, 20) & vbCrLf
         Exit Do
      End If

      
      ' add key/value pair
      sKey = parseKey(str, Index)
      On Error Resume Next
      
      parseObject.Add sKey, parseValue(str, Index)
      If Err.Number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
         Exit Do
      End If
   Loop
eh:

End Function

'
'   parse list
'
Private Function parseArray(ByRef str As String, ByRef Index As Long) As Collection

   Set parseArray = New Collection

   ' "["
   Call skipChar(str, Index)
   If Mid(str, Index, 1) <> "[" Then
      psErrors = psErrors & "Invalid Array at position " & Index & " : " + Mid(str, Index, 20) & vbCrLf
      Exit Function
   End If
   
   Index = Index + 1

   Do

      Call skipChar(str, Index)
      If "]" = Mid(str, Index, 1) Then
         Index = Index + 1
         Exit Do
      ElseIf "," = Mid(str, Index, 1) Then
         Index = Index + 1
         Call skipChar(str, Index)
      ElseIf Index > Len(str) Then
         psErrors = psErrors & "Missing ']': " & Right(str, 20) & vbCrLf
         Exit Do
      End If

      ' add value
      On Error Resume Next
      parseArray.Add parseValue(str, Index)
      If Err.Number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & Mid(str, Index, 20) & vbCrLf
         Exit Do
      End If
   Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef Index As Long)

   Call skipChar(str, Index)

   Select Case Mid(str, Index, 1)
      Case "{"
         Set parseValue = parseObject(str, Index)
      Case "["
         Set parseValue = parseArray(str, Index)
      Case """", "'"
         parseValue = parseString(str, Index)
      Case "t", "f"
         parseValue = parseBoolean(str, Index)
      Case "n"
         parseValue = parseNull(str, Index)
      Case Else
         parseValue = parseNumber(str, Index)
   End Select

End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef Index As Long) As String

   Dim quote   As String
   Dim Char    As String
   Dim Code    As String

   Dim SB As New cStringBuilder

   Call skipChar(str, Index)
   quote = Mid(str, Index, 1)
   Index = Index + 1
   
   Do While Index > 0 And Index <= Len(str)
      Char = Mid(str, Index, 1)
      Select Case (Char)
         Case "\"
            Index = Index + 1
            Char = Mid(str, Index, 1)
            Select Case (Char)
               Case """", "\", "/", "'"
                  SB.Append Char
                  Index = Index + 1
               Case "b"
                  SB.Append vbBack
                  Index = Index + 1
               Case "f"
                  SB.Append vbFormFeed
                  Index = Index + 1
               Case "n"
                  SB.Append vbLf
                  Index = Index + 1
               Case "r"
                  SB.Append vbCr
                  Index = Index + 1
               Case "t"
                  SB.Append vbTab
                  Index = Index + 1
               Case "u"
                  Index = Index + 1
                  Code = Mid(str, Index, 4)
                  SB.Append ChrW(Val("&h" + Code))
                  Index = Index + 4
            End Select
         Case quote
            Index = Index + 1
            
            parseString = SB.toString
            Set SB = Nothing
            
            Exit Function
            
         Case Else
            SB.Append Char
            Index = Index + 1
      End Select
   Loop
   
   parseString = SB.toString
   Set SB = Nothing
   
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef Index As Long)

   Dim Value   As String
   Dim Char    As String

   Call skipChar(str, Index)
   Do While Index > 0 And Index <= Len(str)
      Char = Mid(str, Index, 1)
      If InStr("+-0123456789.eE", Char) Then
         Value = Value & Char
         Index = Index + 1
      '--- Evita erro de overflow no VB caso o WS retorne valores altamente precisos
      ElseIf Len(Value) > 28 Then
         parseNumber = CDec(Val(Value))
         Exit Function
      Else
         parseNumber = CDec(Value)
         Exit Function
      End If
   Loop
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean

   Call skipChar(str, Index)
   If Mid(str, Index, 4) = "true" Then
      parseBoolean = True
      Index = Index + 4
   ElseIf Mid(str, Index, 5) = "false" Then
      parseBoolean = False
      Index = Index + 5
   Else
      psErrors = psErrors & "Invalid Boolean at position " & Index & " : " & Mid(str, Index) & vbCrLf
   End If

End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef Index As Long)

   Call skipChar(str, Index)
   If Mid(str, Index, 4) = "null" Then
      parseNull = Null
      Index = Index + 4
   Else
      psErrors = psErrors & "Invalid null value at position " & Index & " : " & Mid(str, Index) & vbCrLf
   End If

End Function

Private Function parseKey(ByRef str As String, ByRef Index As Long) As String

   Dim dquote  As Boolean
   Dim squote  As Boolean
   Dim Char    As String

   Call skipChar(str, Index)
   Do While Index > 0 And Index <= Len(str)
      Char = Mid(str, Index, 1)
      Select Case (Char)
         Case """"
            dquote = Not dquote
            Index = Index + 1
            If Not dquote Then
               Call skipChar(str, Index)
               If Mid(str, Index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                  Exit Do
               End If
            End If
         Case "'"
            squote = Not squote
            Index = Index + 1
            If Not squote Then
               Call skipChar(str, Index)
               If Mid(str, Index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                  Exit Do
               End If
            End If
         Case ":"
            Index = Index + 1
            If Not dquote And Not squote Then
               Exit Do
            Else
               parseKey = parseKey & Char
            End If
         Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
            Else
               parseKey = parseKey & Char
            End If
            Index = Index + 1
      End Select
   Loop

End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef Index As Long)
   Dim bComment As Boolean
   Dim bStartComment As Boolean
   Dim bLongComment As Boolean
   Do While Index > 0 And Index <= Len(str)
      Select Case Mid(str, Index, 1)
      Case vbCr, vbLf
         If Not bLongComment Then
            bStartComment = False
            bComment = False
         End If
         
      Case vbTab, " ", "(", ")"
         
      Case "/"
         If Not bLongComment Then
            If bStartComment Then
               bStartComment = False
               bComment = True
            Else
               bStartComment = True
               bComment = False
               bLongComment = False
            End If
         Else
            If bStartComment Then
               bLongComment = False
               bStartComment = False
               bComment = False
            End If
         End If
         
      Case "*"
         If bStartComment Then
            bStartComment = False
            bComment = True
            bLongComment = True
         Else
            bStartComment = True
         End If
         
      Case Else
         If Not bComment Then
            Exit Do
         End If
      End Select
      
      Index = Index + 1
   Loop

End Sub

Public Function toString(ByRef obj As Variant) As String
   Dim SB As New cStringBuilder
   Select Case VarType(obj)
      Case vbNull
         SB.Append "null"
      Case vbDate
         SB.Append """" & CStr(obj) & """"
      Case vbString
         SB.Append """" & Encode(obj) & """"
      Case vbObject
         
         Dim bFI As Boolean
         Dim i As Long
         
         bFI = True
         If TypeName(obj) = "Dictionary" Then

            SB.Append "{"
            Dim keys
            keys = obj.keys
            For i = 0 To obj.Count - 1
               If bFI Then bFI = False Else SB.Append ","
               Dim key
               key = keys(i)
               SB.Append """" & key & """:" & toString(obj.Item(key))
            Next i
            SB.Append "}"

         ElseIf TypeName(obj) = "Collection" Then

            SB.Append "["
            Dim Value
            For Each Value In obj
               If bFI Then bFI = False Else SB.Append ","
               SB.Append toString(Value)
            Next Value
            SB.Append "]"

         End If
      Case vbBoolean
         If obj Then SB.Append "true" Else SB.Append "false"
      Case vbVariant, vbArray, vbArray + vbVariant
         Dim sEB
         SB.Append multiArray(obj, 1, "", sEB)
      Case Else
         SB.Append Replace(obj, ",", ".")
   End Select

   toString = SB.toString
   Set SB = Nothing
   
End Function

Private Function Encode(str) As String

   Dim SB As New cStringBuilder
   Dim i As Long
   Dim j As Long
   Dim aL1 As Variant
   Dim aL2 As Variant
   Dim c As String
   Dim P As Boolean

   aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
   aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
   For i = 1 To Len(str)
      P = True
      c = Mid(str, i, 1)
      For j = 0 To 7
         If c = Chr(aL1(j)) Then
            SB.Append "\" & Chr(aL2(j))
            P = False
            Exit For
         End If
      Next

      If P Then
         Dim a
         a = AscW(c)
         If a > 31 And a < 127 Then
            SB.Append c
         ElseIf a > -1 Or a < 65535 Then
            SB.Append "\u" & VBA.String(4 - Len(Hex(a)), "0") & Hex(a)
         End If
      End If
   Next
   
   Encode = SB.toString
   Set SB = Nothing
   
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
   
   Dim iDU As Long
   Dim iDL As Long
   Dim i As Long
   
   On Error Resume Next
   iDL = LBound(aBD, iBC)
   iDU = UBound(aBD, iBC)

   Dim SB As New cStringBuilder

   Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
   If Err.Number = 9 Then
      sPB1 = sPT & sPS
      For i = 1 To Len(sPB1)
         If i <> 1 Then sPB2 = sPB2 & ","
         sPB2 = sPB2 & Mid(sPB1, i, 1)
      Next
      '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
      SB.Append toString(aBD(sPB2))
   Else
      sPT = sPT & sPS
      SB.Append "["
      For i = iDL To iDU
         SB.Append multiArray(aBD, iBC + 1, i, sPT)
         If i < iDU Then SB.Append ","
      Next
      SB.Append "]"
      sPT = Left(sPT, iBC - 2)
   End If
   Err.Clear
   multiArray = SB.toString
   
   Set SB = Nothing
End Function

' Miscellaneous JSON functions

Public Function StringToJSON(st As String) As String
   
   Const FIELD_SEP = "~"
   Const RECORD_SEP = "|"

   Dim sFlds As String
   Dim sRecs As New cStringBuilder
   Dim lRecCnt As Long
   Dim lFld As Long
   Dim fld As Variant
   Dim rows As Variant

   lRecCnt = 0
   If st = "" Then
      StringToJSON = "null"
   Else
      rows = Split(st, RECORD_SEP)
      For lRecCnt = LBound(rows) To UBound(rows)
         sFlds = ""
         fld = Split(rows(lRecCnt), FIELD_SEP)
         For lFld = LBound(fld) To UBound(fld) Step 2
            sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
         Next 'fld
         sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
      Next 'rec
      StringToJSON = IIf(lRecCnt > 1, ("[" & vbCrLf & sRecs.toString & vbCrLf & "]"), (vbCrLf & sRecs.toString & vbCrLf))
   End If
End Function


Public Function RStoJSON(RS As ADODB.Recordset) As String
   On Error GoTo errHandler
   Dim sFlds As String
   Dim sRecs As New cStringBuilder
   Dim lRecCnt As Long
   Dim fld As ADODB.Field

   lRecCnt = 0
   If RS.State = adStateClosed Then
      RStoJSON = "null"
   Else
      If RS.EOF Or RS.BOF Then
         RStoJSON = "null"
      Else
         Do While Not RS.EOF And Not RS.BOF
            lRecCnt = lRecCnt + 1
            sFlds = ""
            For Each fld In RS.Fields
               sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
            Next 'fld
            sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
            RS.MoveNext
         Loop
         '--- IDT15987 Murilo 22/02/2018 --- Caso seja mais de um registro transforma em vetor com []
         'RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
         RStoJSON = IIf(lRecCnt > 1, "[" & vbCrLf & sRecs.toString & vbCrLf & "]", vbCrLf & sRecs.toString & vbCrLf)
      End If
   End If

   Exit Function
errHandler:

End Function


Public Function toUnicode(str As String) As String

   Dim X As Long
   Dim uStr As New cStringBuilder
   Dim uChrCode As Integer

   For X = 1 To Len(str)
      uChrCode = Asc(Mid(str, X, 1))
      Select Case uChrCode
         Case 8:   ' backspace
            uStr.Append "\b"
         Case 9: ' tab
            uStr.Append "\t"
         Case 10:  ' line feed
            uStr.Append "\n"
         Case 12:  ' formfeed
            uStr.Append "\f"
         Case 13: ' carriage return
            uStr.Append "\r"
         Case 34: ' quote
            uStr.Append "\"""
         Case 39:  ' apostrophe
            uStr.Append "\'"
         Case 92: ' backslash
            uStr.Append "\\"
         Case 123, 125:  ' "{" and "}"
            uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
         Case Is < 32, Is > 127: ' non-ascii characters
            uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
         Case Else
            uStr.Append Chr$(uChrCode)
      End Select
   Next
   toUnicode = uStr.toString
   Exit Function

End Function

Private Sub Class_Initialize()
   psErrors = ""
End Sub

Public Function FormataCampoJSON(Campo1 As String, Valor1 As Variant) As String

Dim Campo As String
Dim Valor As Variant

Campo = Campo1
Valor = Valor1

If Not IsNumeric(Valor) Then Valor = """" & Valor & """"

FormataCampoJSON = """" & Campo & """" & ": " & Valor

End Function

Public Function FormataCampoJSON2(Campo1 As String, Valor1 As Variant) As String

Dim Campo As String
Dim Valor As Variant

Campo = Campo1
Valor = Valor1

FormataCampoJSON2 = """" & Campo & """" & ": " & """" & Valor & """"

End Function

Public Function ExecutaRequestHttpJson(Endereco As String, DadosJson As String, Optional Token As String) As String

Dim HTTP As Object
Dim ObjJson As Object
Const HTTP_BAD_REQUEST As Long = 500

Set HTTP = CreateObject("Microsoft.XmlHttp")
HTTP.Open "POST", Endereco, False
HTTP.setRequestHeader "Content-Type", "application/json"

If Token <> "" Then
    HTTP.setRequestHeader "Authorization", "Bearer " & Token
End If

Set ObjJson = JSON.parse(DadosJson)

HTTP.send JSON.toString(ObjJson)

If HTTP.Status >= HTTP_BAD_REQUEST Then
    Err.Raise HTTP.Status, "ExecutaRequestHttpJson", HTTP.responseText
Else
    ExecutaRequestHttpJson = HTTP.responseText
End If

End Function


Public Function ExecutaRequestHttpJsonRioVerde(Endereco As String, DadosJson As String, Optional Uf As String, Optional Tenant As String, Optional Token As String) As String

Dim HTTP As Object
Dim ObjJson As Object
Const HTTP_BAD_REQUEST As Long = 500

Set HTTP = CreateObject("Microsoft.XmlHttp")
HTTP.Open "POST", Endereco, False
HTTP.setRequestHeader "Content-Type", "application/json"

Set ObjJson = JSON.parse(DadosJson)

HTTP.send JSON.toString(ObjJson)

If HTTP.Status >= HTTP_BAD_REQUEST Then
    Err.Raise HTTP.Status, "ExecutaRequestHttpJsonRioVerde", HTTP.responseText
Else
    ExecutaRequestHttpJsonRioVerde = HTTP.responseText
End If


End Function

Public Function ExecutaRequestHttpJsonSISGENFEWS(Endereco As String, DadosJson As String, Optional Token As String) As String

Dim HTTP As Object
Dim ObjJson As Object
Const HTTP_BAD_REQUEST As Long = 500

Set HTTP = CreateObject("Microsoft.XmlHttp")
HTTP.Open "POST", Endereco, False
HTTP.setRequestHeader "Content-Type", "application/json"

If Token <> "" Then
    HTTP.setRequestHeader "Authorization", "Bearer " & Token
End If

Set ObjJson = JSON.parse(DadosJson)

HTTP.send JSON.toString(ObjJson)

ExecutaRequestHttpJsonSISGENFEWS = HTTP.responseText

End Function
