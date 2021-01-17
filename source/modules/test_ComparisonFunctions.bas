Option Compare Database
Option Explicit

Public Function compare_CaseSensitive(s1 As String, s2 As String) As Boolean

    compare_CaseSensitive = (StrComp(s1, s2, vbBinaryCompare) = 0)
    
End Function


Public Function compare_date(sDate1 As String, sDate2 As String) As Boolean

    Dim vDate1 As Date, vDate2 As Date
    
    vDate1 = parseDate(sDate1)
    vDate2 = parseDate(sDate2)
    
    If Not IsNullDate(vDate1) And Not IsNullDate(vDate2) Then
        compare_date = (vDate1 = vDate2)
    End If
    
End Function

Private Function IsNullDate(vInputDate As Date) As Boolean
    Const ISNULLDATE_NULLDATE = #12:00:00 AM#
    
    IsNullDate = (vInputDate = ISNULLDATE_NULLDATE)
    
End Function

Private Function parseDate(ByVal sDateString As String) As Date
     
     On Error Resume Next
     
    If Len(sDateString) > 3 And Left(sDateString, 1) = "#" And Right(sDateString, 1) = "#" Then
        sDateString = Mid(sDateString, 2, Len(sDateString) - 2)
    End If
    
    parseDate = CVDate(sDateString)
     
     
End Function
Public Function compare_IsInArray(sResult As String, sConcatenatedArray As String) As Boolean

    Dim aArray, iArray As Long
    
    aArray = Split(sConcatenatedArray, ",")
    
    If IsArray(aArray) Then
        For iArray = LBound(aArray) To UBound(aArray)
            If sResult = Trim("" & aArray(iArray)) Then
                compare_IsInArray = True
                Exit For
            End If
        Next
    End If

End Function


Public Function compare_Ubound(sResult As String, sExpectedUbound As String) As Boolean
        
    Dim aArray
        
    aArray = Split(sResult, "|")
    compare_Ubound = (UBound(aArray) = Val(sExpectedUbound))

End Function