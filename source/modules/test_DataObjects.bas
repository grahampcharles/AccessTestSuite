Option Compare Database
Option Explicit

Private Const TEST_PREFIX_PROPERTY_NAME = "test_prefix"
Private Const PROPERTY_NULLSTRING = "{NULLSTRING}"

' // TODO:
' // Test Harness table
' contain the text of the module that should be created for the tests; the mod gets created before and deleted after
' optionally contain _beforeTests code (to create test tables, etc.)
' // TestItems form
' support using a single TestItems form to view multiple TestPrefixes

Public Function TestsInstall(Optional ByVal sPrefix As String = "") As Boolean

    Dim bRet As Boolean
    
    ' / if there was an existing prefix, default to that
    If Len(sPrefix) = 0 Then sPrefix = test_GetDatabaseProperty(TEST_PREFIX_PROPERTY_NAME)
     
    ' / save the prefix
    test_SetDatabaseProperty TEST_PREFIX_PROPERTY_NAME, sPrefix
    
    bRet = test_InstallTables(sPrefix)
    bRet = bRet And test_InstallQueries(sPrefix)
    bRet = bRet And test_InstallForms(sPrefix)
    
    TestsInstall = bRet
    
End Function


Public Function TestsGetCurrentPrefix() As String
    TestsGetCurrentPrefix = test_GetDatabaseProperty(TEST_PREFIX_PROPERTY_NAME)
End Function

Public Function TestsSetCurrentPrefix(vNV As String)
    test_SetDatabaseProperty TEST_PREFIX_PROPERTY_NAME, vNV
End Function

Private Function test_SetDatabaseProperty(sPropName As String, sPropValue As String)
    
    ' / twist: this only functions on an attached database
    Dim db As Database
    Dim prp As Property
        
    Set db = CurrentDb
        
    If sPropValue = "" Then sPropValue = PROPERTY_NULLSTRING
        
    On Error Resume Next
    
    Set prp = db.Properties(sPropName)
    If Err.Number = 3270 Then       '  property not found
        Set prp = db.CreateProperty(sPropName, dbText, sPropValue)
        db.Properties.Append prp
        db.Properties.Refresh
    Else
        prp.Value = sPropValue
    End If
    
End Function


Private Function test_GetDatabaseProperty(sPropName As String) As String
    
    Dim db As Database
    Set db = CurrentDb
    Dim sValue As String
        
    On Error Resume Next
    sValue = db.Properties(sPropName).Value
    If sValue = PROPERTY_NULLSTRING Then sValue = ""
    test_GetDatabaseProperty = sValue

End Function

Public Function TestsRun(Optional sPrefix As String, Optional vTestGroup As String = "", Optional bFailedOnly As Boolean = False) As Boolean

    Dim rsTest As Recordset
    Dim sExpectedResult As String, sResult As String
    Dim sCode As String
    Dim cTest As Long, cSuccess As Long
    
    ' TODO: some more reporting
    
    On Error Resume Next
    
    ' / switch to last-installed prefix
    If Len(sPrefix) = 0 Then
        sPrefix = test_GetDatabaseProperty(TEST_PREFIX_PROPERTY_NAME)
    End If
    
    test_DebugPrint sPrefix, "begin: prefix value " & sPrefix, "TestsRun"
    
    Set rsTest = CurrentDb.OpenRecordset(test_QualifiedObjectName("TestItem", sPrefix))
    If Err.Number <> 0 Then
        test_DebugPrint sPrefix, "aborted: could not open test table; are you sure it has been installed?", "TestsRun"
        Exit Function
    End If
    
    If Not (rsTest.EOF Or rsTest.BOF) Then
        
        rsTest.MoveFirst
        Do Until rsTest.EOF
            If (Len(vTestGroup) = 0) Or (rsTest!TestGroup & "" = vTestGroup) Then
                If (Not bFailedOnly) Or (Not rsTest!Passed) Then
                    
                    ' / set it to failed in case running it crashes this routine
                    cTest = cTest + 1
                    rsTest.Edit
                    rsTest!Passed = False
                    rsTest!DatePassed = Null
                    rsTest!Result = Null
                    rsTest.Update
                    sExpectedResult = rsTest!ExpectedResult
                
                    ' / try to run it
                    Err.Clear
                    
                    ' / get the code
                    sCode = Trim("" & rsTest!TestCode)
                    If Len(sCode) = 0 Then
                        sResult = "** NO CODE PROVIDED **"
                    Else
                        ' / get the result
                        Select Case rsTest!TestType
                            Case "eval"
                                sResult = Eval(sCode)
                            Case "code-array"
                                ' / the result will be an array
                                sResult = Join(test_RunIt(rsTest!TestCode, _
                                    test_GetValue(rsTest!Param1.Value), _
                                    test_GetValue(rsTest!Param2.Value), _
                                    test_GetValue(rsTest!Param3.Value)), _
                                    "|")
                            Case Else   ' / default to "code"
                                sResult = test_RunIt(rsTest!TestCode, _
                                    test_GetValue(rsTest!Param1.Value), _
                                    test_GetValue(rsTest!Param2.Value), _
                                    test_GetValue(rsTest!Param3.Value))

                        End Select
                    End If
                    
                    ' / did we succeed?
                    rsTest.Edit
    
                    If Err.Number <> 0 Then
                        rsTest!Result = Err.Description
                    Else
                        rsTest!Result = test_TokenizeValue(sResult)
                        If test_CheckResult(sResult, sExpectedResult, "" & rsTest!ComparisonFunction) Then
                            If Err.Number = 0 Then
                                rsTest!Passed = True
                                rsTest!DatePassed = Now
                                cSuccess = cSuccess + 1
                            Else
                                rsTest!Result = Err.Description
                            End If
                        End If
                    End If
                    rsTest.Update
                    Err.Clear        ' / e.g. catch
                End If
            End If
            rsTest.MoveNext
            If Err.Number <> 0 Then
                test_DebugPrint sPrefix, "aborted: error during run", "TestsRun"
            End If
        Loop
    End If
        
    test_DebugPrint sPrefix, "Complete. " & cTest & " test(s) run; " & cSuccess & " passed", "TestsRun"
        
    TestsRun = (cSuccess = cTest)
        
End Function

Private Function test_GetValue(vInput)

    Dim aTokens: aTokens = Array("{SPACE}", "{CRLF}")
    Dim aReplacements:    aReplacements = Array(" ", vbCrLf)

    ' / process tokens
    If "" & vInput = "{NULLSTRING}" Then
        test_GetValue = vbNullString
    ElseIf "" & vInput = "{TRUE}" Then
        test_GetValue = True
    ElseIf "" & vInput = "{FALSE}" Then
        test_GetValue = False
    ElseIf "" & vInput = "{NULL}" Then
        test_GetValue = Null
    ElseIf Len("" & vInput) >= 5 And Left("" & vInput, 1) = "#" And Right("" & vInput, 1) = "#" Then
        ' / it's  a date?
        ' / TODO: deal with parse errors
        test_GetValue = CVDate(Mid("" & vInput, 2, Len("" & vInput) - 2))
    ElseIf VarType(vInput) = vbString Then
        test_GetValue = test_ReplaceTokens("" & vInput, aTokens, aReplacements)
    Else
        test_GetValue = vInput
    End If

End Function


Private Function test_TokenizeValue(vInput) As String

    Dim aTokens: aTokens = Array(vbCrLf)
    Dim aReplacements: aReplacements = Array("{CRLF}")

    ' / turns result back into tokens for viewing results
    If IsNull(vInput) Then
        test_TokenizeValue = "{NULL}"
    ElseIf VarType(vInput) = vbBoolean Then
        If CBool(vInput) Then
            test_TokenizeValue = "{TRUE}"
        Else
            test_TokenizeValue = "{FALSE}"
        End If
    ElseIf VarType(vInput) = vbString Then
        If Len("" & vInput) = 0 Then
            test_TokenizeValue = "{NULLSTRING}"
        Else
            test_TokenizeValue = test_ReplaceTokens("" & vInput, aTokens, aReplacements)
        End If
    Else
        test_TokenizeValue = vInput
    End If

End Function


Private Function test_ReplaceTokens(sInput As String, aTokens, aReplacements) As String
    On Error Resume Next
    Dim i As Long, sReturn As String
    
    sReturn = sInput
    
    If IsArray(aTokens) And IsArray(aReplacements) Then
        For i = LBound(aTokens) To UBound(aTokens)
            sReturn = Replace(sReturn, "" & aTokens(i), "" & aReplacements(i))
        Next
    End If
    
    test_ReplaceTokens = sReturn
        
End Function

Private Function test_RunIt(sCode As String, vParam1, vParam2, vParam3)

    ' / runs Application.Run with 0, 1, 2, or 3 parameters
    If Not (IsNull(vParam3)) Then
        test_RunIt = Application.Run(sCode, test_ParseValue(vParam1), test_ParseValue(vParam2), test_ParseValue(vParam3))
    ElseIf Not (IsNull(vParam2)) Then
        test_RunIt = Application.Run(sCode, test_ParseValue(vParam1), test_ParseValue(vParam2))
    ElseIf Not (IsNull(vParam1)) Then
        test_RunIt = Application.Run(sCode, test_ParseValue(vParam1))
    Else
        test_RunIt = Application.Run(sCode)
    End If
        
End Function

Public Function test_ParseValue(vInput)

    ' / parses special case values
    If Left(vInput & "", 6) = "Array(" Then
       test_ParseValue = test_ParseArray(vInput)
    ElseIf Left(vInput, 1) = "=" Then
        test_ParseValue = Eval(Mid("" & vInput, 2))
    Else
        test_ParseValue = vInput
    End If

End Function

Public Function test_ParseArray(vInput)
    Dim sInput As String
    Dim aResult, iResult As Long
    
    sInput = Trim("" & vInput)
    
    If Left(sInput, 6) = "Array(" Then
        sInput = Trim(Mid(sInput, 7, Len(sInput) - 7))  '/ strip Array() and whitespace
    End If
        
    aResult = Split(sInput, ",")
    
    For iResult = LBound(aResult) To UBound(aResult)
        aResult(iResult) = Eval(aResult(iResult) & "")        ' / get rid of spaces and quotes
    Next
    
    
    test_ParseArray = aResult
    
End Function


Private Function test_CheckResult(ByVal s1 As String, ByVal s2 As String, Optional ByRef sComparisonFunction As String = "=") As Boolean
    
    ' / return true on successful test
    Dim varArray
    
    ' / special case: result has date delimiters
    If Len(s2) > 3 And Left(s2, 1) = "#" And Right(s2, 1) = "#" And (sComparisonFunction = "=" Or sComparisonFunction = "") Then
        sComparisonFunction = "compare_date"
    End If
    
    ' / special case: the comparitee needs evaluation because it's in the form "=vbCrLf" or whatever
    If Left(s2, 1) = "=" And Len(s2) > 1 Then
        s2 = Eval(Mid(s2, 2))
    End If
    
    ' / run the standard or custom comparison function
    Select Case sComparisonFunction
        Case "=", ""
            test_CheckResult = (s1 = test_GetValue(s2))
        Case Else
            test_CheckResult = Application.Run(sComparisonFunction, s1, s2)
    End Select

End Function

Private Sub test_DebugPrint(ByVal sPrefix As String, ByVal sMessage As String, Optional vModule As String = "")
    
    Dim fMessage As Form
    
    If Len(vModule) > 0 Then sMessage = "[" & vModule & "] " & sMessage
    sMessage = Format(Now, "hh:nn:ss") & ": " & sMessage
    Debug.Print sMessage
    
    On Error Resume Next
    Set fMessage = Forms(test_QualifiedObjectName("TestItems", sPrefix))
    If fMessage Is Nothing Then Set fMessage = test_GetOpenTestItemsForm
    If Not (fMessage Is Nothing) Then
        fMessage.TestsAddMessage Replace(Replace(sMessage, ",", "-"), ";", "-")
    End If

End Sub

Private Function test_GetOpenTestItemsForm() As Form
    Dim iForm As Long
    
    For iForm = 0 To Forms.Count
        If Right(Forms(iForm).Name, 10) = "TestItems" Then
            Set test_GetOpenTestItemsForm = Forms(iForm)
            Exit Function
        End If
    Next
    
End Function

Private Function test_DataTables()
    test_DataTables = Array("TestComparisonFunction", "TestItem")
End Function

Private Function test_InstallTables(Optional ByVal sPrefix As String = "") As Boolean

    ' / install test tables
    Dim sTableName As String
    Dim aTables, iTable As Long, iTerator As Long, aComparisonFunctions
    Dim fld As Field, tbl As TableDef, idx As Index, rs As Recordset, rel As Relation
    Dim db As Database
    
    Set db = CurrentDb   ' / install into the parent database, not the code database
    
    aTables = test_DataTables
    
    ' / TODO: get this from the test_ComparisonFunctions module
    aComparisonFunctions = Array("=", "compare_date", "compare_IsInArray", "compare_CaseSensitive", "compare_UBound")
    
    For iTable = LBound(aTables) To UBound(aTables)
        
        sTableName = test_QualifiedObjectName(aTables(iTable), sPrefix)
        
        ' / create the table and its primary key
        If test_TableExists(db, sTableName) Then
            Set tbl = db.TableDefs(sTableName)
        Else
            ' / create the table
            Set tbl = db.CreateTableDef(sTableName)
            
            With tbl
                ' / primary key
                Set fld = .CreateField("id" & aTables(iTable), dbLong)
                fld.Attributes = dbAutoIncrField
                .Fields.Append fld
                
                ' / PK index
                Set idx = .CreateIndex(aTables(iTable) & "_PK")
                idx.Primary = True
                idx.Fields.Append idx.CreateField("id" & aTables(iTable))
                .Indexes.Append idx
            End With
            
            db.TableDefs.Append tbl
            
            Set fld = Nothing
            Set idx = Nothing
            db.TableDefs.Refresh
            
        End If
        
        ' / table-specific fields
        ' Set tbl = db.TableDefs(sTableName)

        Select Case aTables(iTable)
            Case "TestComparisonFunction"
                Set fld = test_AddFieldIfMissing(tbl, "ComparisonFunction", dbText, 50)
                
                If Not test_IndexExists(tbl, aTables(iTable) & "_Unique") Then
                    With tbl
                        ' / unique index
                        Set idx = .CreateIndex(aTables(iTable) & "_Unique")
                        idx.Unique = True
                        idx.Fields.Append idx.CreateField("ComparisonFunction")
                        .Indexes.Append idx
                    End With
                End If
                
                ' / add records (ignore errors if they already exist)
                On Error Resume Next
                Set rs = tbl.OpenRecordset
                
                For iTerator = LBound(aComparisonFunctions) To UBound(aComparisonFunctions)
                    rs.AddNew: rs!ComparisonFunction = aComparisonFunctions(iTerator) & "": rs.Update
                Next
                rs.Close
                Set rs = Nothing
                On Error GoTo 0
        
        
            Case "TestItem"
                test_AddFieldIfMissing tbl, "TestType", dbText, 50
                test_AddFieldIfMissing tbl, "TestGroup", dbText, 250, 2
                test_AddFieldIfMissing tbl, "TestCode", dbText, 250, 2
                test_AddFieldIfMissing tbl, "Param1", dbText, 250, 1.5
                test_AddFieldIfMissing tbl, "Param2", dbText, 250, 1.5
                test_AddFieldIfMissing tbl, "Param3", dbText, 250, 1.5
                test_AddFieldIfMissing tbl, "ExpectedResult", dbText, 250, 1.5
                test_AddFieldIfMissing tbl, "Result", dbText, 250, 1.5
                test_AddFieldIfMissing tbl, "ComparisonFunction", dbText, 50, 1.75
                test_AddFieldIfMissing tbl, "Passed", dbBoolean
                test_AddFieldIfMissing tbl, "DatePassed", dbDate
                
                ' / default values and combo boxes
                tbl.Fields("ComparisonFunction").DefaultValue = """="""
                tbl.Fields("TestType").DefaultValue = """code"""
                test_AddRowSource tbl.Fields("TestType"), "code;eval;code-array", bIsValueList:=True
                test_AddRowSource tbl.Fields("ComparisonFunction"), test_QualifiedObjectName("TestComparisonFunction", sPrefix), 2, 2, "0;1400"
                
        End Select
        
        tbl.Fields.Refresh
    
    Next
    
    
    Application.RefreshDatabaseWindow
    

End Function


Private Function test_InstallQueries(Optional ByVal sPrefix As String = "") As Boolean

    ' / install
    Dim sQualifiedPrefix As String, aQueries, iQuery As Integer, sQueryName As String
    Dim qdf As QueryDef
    Dim db As Database
    
    Set db = CurrentDb   ' / install into the parent database, not the code database
    
    sQualifiedPrefix = test_QualifiedPrefix(sPrefix)
'    aQueries = Array("qTestItems", "SELECT [{0}TestGroup].*, [{0}TestItem].* FROM [{0}TestGroup] INNER JOIN [{0}TestItem] ON [{0}TestGroup].idTestGroup = [{0}TestItem].idTestGroup;")
    aQueries = Array()
    
    For iQuery = LBound(aQueries) To UBound(aQueries) Step 2
        
        sQueryName = sQualifiedPrefix & aQueries(iQuery)
            
        ' / create the querydef
        If test_QueryExists(db, sQueryName) Then db.QueryDefs.Delete (sQueryName)
        
        Set qdf = db.CreateQueryDef(sQueryName, Replace(aQueries(iQuery + 1) & "", "{0}", sQualifiedPrefix))
    
    Next
    
End Function

Private Function test_QueryExists(theDb As Database, sQueryName As String) As Boolean
    
    Dim bRet As Boolean, qdf As QueryDef
    
    For Each qdf In theDb.QueryDefs
        If qdf.Name = sQueryName Then
            bRet = True
            Exit For
        End If
    Next
    
    test_QueryExists = bRet
    
End Function


Private Sub test_AddRowSource(fld As Field, sRowSource As String, Optional wColumnCount As Integer = 1, Optional wBoundColumn As Integer = 1, Optional sColumnWidths As String = "720", Optional bIsValueList As Boolean = False)
    
    test_FieldCreateOrSetProperty fld, "DisplayControl", dbInteger, AcControlType.acComboBox   '  111
    
    test_FieldCreateOrSetProperty fld, "RowSourceType", dbText, IIf(bIsValueList, "Value List", "Table/Query")
    test_FieldCreateOrSetProperty fld, "RowSource", dbText, sRowSource
    test_FieldCreateOrSetProperty fld, "BoundColumn", dbInteger, wBoundColumn
    test_FieldCreateOrSetProperty fld, "ColumnCount", dbInteger, wColumnCount
    test_FieldCreateOrSetProperty fld, "ColumnWidths", dbText, sColumnWidths
    
    fld.Properties.Refresh
    
End Sub

Private Sub test_FieldCreateOrSetProperty(fld As Field, sPropertyName As String, vType, vValue)

    On Error Resume Next
    Dim prp As Property

    Set prp = fld.Properties(sPropertyName)
    If prp Is Nothing Then
        Set prp = fld.CreateProperty(sPropertyName, vType, vValue)
        fld.Properties.Append prp
    Else
        prp.Value = vValue
    End If
    
End Sub

Private Function test_AddFieldIfMissing(ByRef tdf As TableDef, sFieldName As String, fieldType As DataTypeEnum, Optional vFieldLength, Optional sDefaultWidthInches As Single = 0) As Field

    Dim iField As Long
    Dim fld As Field, bFound As Boolean, prp As Property
    
    If IsMissing(vFieldLength) Or IsNull(vFieldLength) Then vFieldLength = 0  ' / 0 works for non-variable fields, like dates
    
    For Each fld In tdf.Fields
        If fld.Name = sFieldName Then
            bFound = True
            Exit For
        End If
    Next
    
    If Not bFound Then
        Set fld = tdf.CreateField(sFieldName, fieldType, vFieldLength)
        tdf.Fields.Append fld
        tdf.Fields.Refresh
        Set fld = tdf.Fields(sFieldName)
        
        ' set default datasheet width
        If sDefaultWidthInches > 0 Then
            fld.Properties.Append fld.CreateProperty("ColumnWidth", dbInteger, Int(sDefaultWidthInches * 1440))
            fld.Properties.Refresh
        End If
        
        ' set default control
        If fieldType = dbBoolean Then
            Set prp = fld.CreateProperty("DisplayControl", dbInteger, AcControlType.acCheckBox)
            fld.Properties.Append prp
        End If
        
    End If
    
    Set test_AddFieldIfMissing = fld

End Function

Private Function test_QualifiedObjectName(vObjectName, ByVal sObjectPrefix As String) As String
    
    sObjectPrefix = test_QualifiedPrefix(sObjectPrefix)
    test_QualifiedObjectName = sObjectPrefix & vObjectName
    
End Function

Private Function test_QualifiedPrefix(ByVal sPrefix As String) As String
    
    ' / clean prefix
    If sPrefix = PROPERTY_NULLSTRING Then sPrefix = ""
    If Len(sPrefix) > 0 Then
        sPrefix = Trim(Replace(sPrefix, " ", ""))
        If Right(sPrefix, 1) <> "_" Then sPrefix = sPrefix & "_"
    End If
    test_QualifiedPrefix = sPrefix

End Function

Private Function test_DataObjectsExist(Optional ByVal sPrefix As String = "") As Boolean

    Dim aTables, iTables As Long, bRet As Boolean
    
    aTables = test_DataTables
    bRet = True
    
    For iTables = LBound(aTables) To UBound(aTables)
        If Not test_TableExists(CurrentDb, test_QualifiedObjectName(aTables(iTables), sPrefix)) Then
            bRet = False
            Exit For
        End If
    Next
    
    test_DataObjectsExist = bRet

End Function

Private Function test_TableExists(db As DAO.Database, tableName As String) As Boolean

    ' Define iterator to query the object model.
    Dim iTerator As Integer

    For iTerator = 0 To db.TableDefs.Count - 1
        If db.TableDefs(iTerator).Name = tableName Then
            test_TableExists = True
            Exit Function
        End If
    Next

End Function

Private Function test_RelationExists(db As DAO.Database, relationName As String) As Boolean

    ' Define iterator to query the object model.
    Dim iTerator As Integer

    For iTerator = 0 To db.Relations.Count - 1
        If db.Relations(iTerator).Name = relationName Then
            test_RelationExists = True
            Exit Function
        End If
    Next

End Function

Private Function test_IndexExists(tbl As TableDef, indexName As String) As Boolean

    Dim iTerator As Integer

    For iTerator = 0 To tbl.Indexes.Count - 1
        If tbl.Indexes(iTerator).Name = indexName Then
            test_IndexExists = True
            Exit Function
        End If
    Next

End Function


Private Function test_InstallForms(sPrefix As String) As Boolean

    Dim dbCode As Database, dbData As Database
    Dim sDestinationName As String
    Dim aForms, iForm As Long
    Dim bRet As Boolean
    
    Set dbCode = codeDB
    Set dbData = CurrentDb
    bRet = True
    
    aForms = Array("TestItems")
    If sPrefix = PROPERTY_NULLSTRING Then sPrefix = ""
    
    If dbCode.Name = dbData.Name And Len(sPrefix) = 0 Then
        test_DebugPrint sPrefix, "Cannot copy a form over itself; forms not copied.", "test_CopyForms"
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    For iForm = LBound(aForms) To UBound(aForms)
        sDestinationName = test_QualifiedObjectName("" & aForms(iForm), sPrefix)
    
        If test_FormExists(sDestinationName) Then
            ' / delete the form in the destination
            DoCmd.DeleteObject acForm, sDestinationName
        End If
   
        ' / save the form to the active database
        DoCmd.TransferDatabase acImport, "Microsoft Access", dbCode.Name, acForm, aForms(iForm), sDestinationName

        ' / change the form and control datasources to match the new tablename
        If Len(sPrefix) > 0 Then test_UpdateForm "" & aForms(iForm), sPrefix

NextForm:
    Next
    
ExitHere:
    test_InstallForms = bRet
    Exit Function
    
ErrHandler:
    bRet = False
    If Err.Number = 2501 Then
        test_DebugPrint sPrefix, "Microsoft Access has protected this project with a password. Forms cannot be installed while password protection is on.", "test_InstallForms"
    Else
        test_DebugPrint sPrefix, Err.Description, "test_InstallForms"
    End If
    Resume NextForm
    Resume
    
    
End Function

Private Function test_UpdateForm(sFormName As String, sPrefix As String)

    Dim sQualifiedFormName As String
    Dim f As Form
    Dim ctl As Control
    
    sQualifiedFormName = test_QualifiedObjectName(sFormName, sPrefix)
    
    DoCmd.OpenForm sQualifiedFormName, acDesign, , , , acHidden
    Set f = Forms(sQualifiedFormName)
    
    f.RecordSource = test_QualifiedRecordSource(f.RecordSource, sPrefix)
    
    ' / do the same to control datasources
    For Each ctl In f.Controls
        If TypeOf ctl Is ComboBox Or TypeOf ctl Is ListBox Then
            If ctl.RowSourceType = "Table/Query" Then
                ctl.RowSource = test_QualifiedRecordSource(ctl.RowSource, sPrefix)
            End If
        End If
    Next
    
    DoCmd.Close acForm, sQualifiedFormName, acSaveYes
    
End Function

Public Function test_QualifiedRecordSource(sRecordsource As String, sPrefix As String)

    Dim sQualifiedPrefix As String
    
    sQualifiedPrefix = test_QualifiedPrefix(sPrefix)

    If Left(sRecordsource, 6) = "SELECT" Then
        test_QualifiedRecordSource = Replace(sRecordsource, "FROM ", "FROM " & sQualifiedPrefix)
    Else
        test_QualifiedRecordSource = sQualifiedPrefix & sRecordsource
    End If
End Function

Private Function test_FormExists(formName As String) As Boolean
    
    ' / TODO: check how this works when it's in a code database
    
    Dim frm As AccessObject
    
    For Each frm In Application.CurrentProject.AllForms
        If frm.Name = formName Then
            test_FormExists = True
            Exit For
        End If
    Next

End Function