
Sub FromPairsToSel(Model, Cube, Pairs(), Sel())
    Dim N, I, J, D
    N = CallMDS(MdsT.TableDimensionsCount(Model, Cube))
    For I = 1 To N
        D = UCase(CallMDS(MdsT.TableDimensionsName(Model, Cube, I)))
        For J = 0 To UBound(Pairs, 2)
            If Pairs(0, J) = D Then Sel(I - 1) = Pairs(1, J)
        Next
    Next
End Sub

Sub PrepareCopyDataArea(Model, Cube, Sel())
    Dim Matrix(), N, I, D
    N = CallMDS(MdsT.TableDimensionsCount(Model, Cube))
    ReDim Matrix(2, N - 1)
    For I = 1 To N
        D = UCase(CallMDS(MdsT.TableDimensionsName(Model, Cube, I)))
        If Sel(I - 1) <> "" Then
            Matrix(1, I - 1) = Sel(I - 1)
        Else
            Matrix(0, I - 1) = "*"
        End If
    Next
    Dim Val1, Op1, Val2, Op2, DW
    Val1 = 0
    Op1 = mdsOperatorNone
    Val2 = 0
    Op2 = mdsOperatorNone
    DW = ""
    CallMDS MdsT.DataareaDefine_VBS(Model, Cube, DW, Matrix, Val1, Op1, Val2, Op2, True, True)
End Sub

Sub ExportRows(FN, N, Sel())
    Dim I, Cols(), Continue
    If TypeName(MdsT.RecordLoopFromDataarea(".")) <> "Error" Then
        Do
            Continue = MdsT.RecordLoopGetNext_VBS(False, Cols)
            If Cols(N) <> "#N/A" Then
                For I = 1 To N
                    FN.Write IIF(Sel(I - 1) = "", Cols(I - 1), Sel(I - 1)) & vbTab
                Next
                FN.WriteLine Cols(N)
            End If
        Loop While Continue
    End If
End Sub

Sub ExportCopyData(Model, Cube, N, SelFrom(), SelTo(), FileName)
    GetFileSystemObject()
    FileName = FS.GetTempName()
    FileName = FS.BuildPath(FS.GetSpecialFolder(2), FileName)
    Dim FN, R, Continue
    PrepareCopyDataArea Model, Cube, SelFrom
    Set FN = FS.CreateTextFile(FileName)
    ExportRows FN, N, SelTo
    Fn.Close
    CallMDS MdsT.DataareaDestroy
End Sub

Sub ImportCopyData(Model, Cube, N, FileName)
    Dim Columns(), I, FI, FE, Line, L2, Skipped, Written, R
    Skipped = 0
    Written = 0
    ReDim Columns(N - 1)
    Set FE = FS.CreateTextFile(FS.BuildPath(FS.GetSpecialFolder(2),"Err.Log"))
    Set FI = FS.OpenTextFile(FileName, 1)
    Do While Not(FI.AtEndOfStream)
        Line = FI.ReadLine
		L2 = Line
		For I = 1 To N
            Columns(I - 1) = GetField(Line, vbTab)
		Next
        R = MdsDc.DataPutValue2Ex_VBS(Model, Cube, 1, Cstr(Line), Columns)
        If TypeName(R) = "Error" Then
		    Skipped = Skipped + 1
			FE.WriteLine L2
        Else
		    Written = Written + 1
        End If
    Loop
	FE.Close
    MsgBox "Written: " & Written & ", Skipped: " & Skipped
End Sub

Sub CopyCubeData(Model, Cube, PairsFrom(), PairsTo())
    Dim SelFrom(), SelTo(), I, N, FileName, ColDef()
    N = CallMDS(MdsT.TableDimensionsCount(Model, Cube))
    ReDim SelFrom(N - 1)
    ReDim SelTo(N - 1)
    ReDim ColDef(N)
    For I = 0 To N
 	   ColDef(I) = I
    Next
    FromPairsToSel Model, Cube, PairsFrom , SelFrom
    FromPairsToSel Model, Cube, PairsFrom , SelTo
    FromPairsToSel Model, Cube, PairsTo   , SelTo
    ExportCopyData Model, Cube, N, SelFrom, SelTo, FileName
    PrepareCopyDataArea Model, Cube, SelTo
    CallMDS MdsT.DataareaSetValue(Null)
    CallMDS MdsT.DataareaDestroy
    ImportCopyData Model, Cube, N, FileName
    'CallMDS MdsT.TableImport(Model, Cube, FileName, "err.log", vbTab)
    FS.DeleteFile FileName
End Sub
