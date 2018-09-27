Option Explicit
@Include "SaveSelections"
@Include "Routines"


Const dimTCalc    = "TCALC"
Const dimDataType = "DATATYPE"
Const FreezeCube  = "TFREEZE"

Sub Spreadsheet_StartRecalc(IsRecalc)
	Dim Action
	Action = Application.GetReportVariable("FreezeAction")
	If Action <> "" Then
		Application.SetReportVariable "FreezeAction", ""
		FreezeAction Action, Application.GetReportVariable("FreezeDims"), Application.GetReportVariable("FreezeElements")
	End If
End Sub

Sub FreezeAction(Action, Dims, Elems)
	Dim Model
    Dim Sel(20)
    Dim I, D, E, Calc, Alias, cube
	InitAlea
	Alias = Application.GetReportVariable("AliasUN")
	GetField Alias, "].["
	Alias = GetField(Alias, "]")
	Model = Connect(Alias)
	Do While Dims <> ""
        D = UCASE(GetField(Dims, ","))
		E = GetField(Elems, ",")
		If IsCalculation(Model, D, E) Then Calc = E
		I = HasDim(Model, FreezeCube, D)
		If I > 0 Then Sel(I) = E
	Loop
	If Calc = "" Then Err.Raise -1,, "Cannot " + Action + ", no Calculation present."
    For I = 1 To CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube)) - 2
        D = CallMDS(MdsT.TableDimensionsName(Model, FreezeCube, I))
		If Sel(I) = "" Then Err.Raise -1, , "Cannot " + Action + " " + Calc + ", dimension not present: " + D
    Next
    If Action = "Freeze" Then
        FreezeCalcUpTo Model, Calc, Sel
    Else
        UnFreezeCalcFrom Model, Calc, Sel
    End If
End Sub

Sub UnFreezeCalcFrom(Model, Calc, Sel() )
    Dim I, E 
    For I = CallMDS(MdsE.ElementIndex(Model, dimTCalc, Calc)) To CallMDS(MdsE.ElementsCount(Model, dimTCalc))
        E = CallMDS(MdsE.ElementsName(Model, dimTCalc, I))
        Sel(CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube)) - 1) = E
        SetFrozen Model, Sel, False
    Next
End Sub

Sub FreezeCalcUpTo(Model, Calc, Sel())
	Set FS = GetFileSystemObject()
    Dim I, E, FileNames(), N
	N = CallMDS(MdsE.ElementIndex(Model, dimTCalc, Calc))
    ReDim FileNames(N)
    For I = 1 To N
        E = CallMDS(MdsE.ElementsName(Model, dimTCalc, I))
        ExportCalculation Model, E, Sel, FileNames(I)
    Next
'	CallMds MdsDC.BulkTransferBegin(1)
    For I = 1 To N
        E = CallMDS(MdsE.ElementsName(Model, dimTCalc, I))
        If FileNames(I) <> "" Then
            ImportCalculation Model, E, Sel, FileNames(I)
        End If
    Next
	Dim ErrorLog
	ErrorLog = ""
'	CallMds MdsDC.BulkTransferCommit(1, true, ErrorLog)
End Sub

Sub ExportCalculation(Model, Calc, Sel(), FileName )
    Dim I, Cube
    Cube = CubeFromCalculation(Model, Calc)
    If HasDim(Model, Cube, dimDataType) > 0 Then 'is freezable
        Sel(CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube)) - 1) = Calc
        If Not IsCompletelyFrozen(Model, Sel) Then
            ExportFreezeSelection Model, Cube, Sel, FileName, "Value", "Frozen Value"
        End If
    End If
End Sub

Sub ImportCalculation(Model, Calc, Sel(), FileName )
    Sel(CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube)) - 1) = Calc
    If FileName <> "" Then
        Dim I, Cube 
        Cube = CubeFromCalculation(Model, Calc)
        ImportFreezeSelection Model, Cube, Sel, FileName, "Frozen Value"
        FS.DeleteFile FileName
    End If
    SetFrozen Model, Sel, True
End Sub

Function SetFrozen(Model, Sel(), Frozen)
    Sel(CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube))) = "Is Frozen"
    PrepareFreezeDataArea Model, FreezeCube, Sel, "", mdsOperatorNone
    If Frozen Then
        CallMDS MdsT.DataareaSetValue(1)
    Else
        CallMDS MdsT.DataareaSetValue(Null)
    End If
    CallMDS MdsT.DataareaDestroy
    Sel(CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube))) = ""
End Function

Function IsCompletelyFrozen(Model, Sel() )
	Dim N
	N = CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube))
    Sel(N) = "Is Frozen"
    Dim R1, R2
    R1 = CallMDS(MdsDC.DataGetValue(Model, FreezeCube, Sel(1), Sel(2), Sel(3), Sel(4), Sel(5), Sel(6), Sel(7), Sel(8), Sel(9), Sel(10), Sel(11), Sel(12), Sel(13), Sel(14), Sel(15), Sel(16), Sel(17), Sel(18), Sel(19), Sel(20)))
'    Sel(N) = "Count"
'    R2 = CallMDS(MdsDC.DataGetValue(Model, FreezeCube, Sel(1), Sel(2), Sel(3), Sel(4), Sel(5), Sel(6), Sel(7), Sel(8), Sel(9), Sel(10), Sel(11), Sel(12), Sel(13), Sel(14), Sel(15), Sel(16), Sel(17), Sel(18), Sel(19), Sel(20)))
'    Sel(N) = ""
	IsCompletelyFrozen = False
    On Error Resume Next
    If R1 = 1 Then IsCompletelyFrozen = True
End Function

Function HasDim(Model, Cube, D ) 
    Dim I 
    For I = 1 To CallMDS(MdsT.TableDimensionsCount(Model, Cube))
        If UCase(D) = UCase(CallMDS(MdsT.TableDimensionsName(Model, Cube, I))) Then
            HasDim = I
            Exit For
        End If
    Next
End Function

Sub FillWithTimeRange(Model, Matrix(), D, I )
End Sub

Sub FillWithInputs(Model, Matrix(), D, I )
    Dim J, K, E 
    For J = 1 To CallMDS(MdsE.ElementsCount(Model, D))
        E = CallMDS(MdsE.ElementsName(Model, D, J))
        If CallMDS(MdsE.ElementChildrenCount(Model, D, E)) = 0 Then
            K = K + 1
            Matrix(K, I - 1) = E
        End If
    Next
End Sub

Sub PrepareFreezeDataArea(Model, Cube, Sel(), Datatype, Operator)
    Dim Matrix(), N, L, M, I, D, J 
    N = CallMDS(MdsT.TableDimensionsCount(Model, Cube))
    For I = 1 To N - 1
        D = CallMDS(MdsT.TableDimensionsName(Model, Cube, I))
        L = CallMDS(MdsE.ElementsCount(Model, D))
        If L > M Then M = L
    Next
    ReDim Matrix(M + 1, N - 1)
    For I = 1 To N
        D = UCase(CallMDS(MdsT.TableDimensionsName(Model, Cube, I)))
        If D = dimDataType Then
            Matrix(1, I - 1) = Datatype
        Else
            J = HasDim(Model, FreezeCube, D)
            If J <> 0 Then
                If Sel(J) = "*" Then
                    FillWithInputs Model, Matrix, D, I
                ElseIf CallMDS(MdsE.ElementChildrenCount(Model, D, Sel(J))) = 0 Then
                    Matrix(1, I - 1) = Sel(J)
                Else
                    FillWithChildren Matrix, I - 1, Model, D, Sel(J), 1
                End If
            ElseIf I = N Then
                Matrix(1, I - 1) = Sel(CallMDS(MdsT.TableDimensionsCount(Model, FreezeCube)) - 1)
            Else
                FillWithInputs Model, Matrix, D, I
            End If
        End If
    Next
    Dim Val1, Op1, Val2, Op2, DW
    Val1 = 0
    Op1 = Operator
    Val2 = 0
    Op2 = mdsOperatorNone
    
    DW = ""
    CallMDS MdsT.DataareaDefine_VBS(Model, Cube, DW, Matrix, Val1, Op1, Val2, Op2, False, False)
End Sub

Function ListDims(Dims()) 
    Dim D, I 
    For I = 1 To UBound(Dims)
        If Dims(I) = "" Then Exit For
        D = D + ", " + Dims(I)
    Next
    ListDims = UCase(Mid(D, 2))
End Function

Sub ExportFreezeSelection(Model, Cube, Sel(), FileName, Source, Target )
    Dim N, DI 
    N = CallMDS(MdsT.TableDimensionsCount(Model, Cube))
    DI = HasDim(Model, Cube, dimDataType)
    FileName = FS.GetTempName
    FileName = FS.BuildPath(FS.GetSpecialFolder(2), FileName)
    Dim FN, R, Continue
    PrepareFreezeDataArea Model, Cube, Sel, Source, mdsOperatorNotEqual
    Set FN = FS.CreateTextFile(FileName)
    ExportRows FN, N, DI, Target
    Fn.Close
    CallMDS MdsT.DataareaDestroy
End Sub

Sub ExportRows(FN, N, DI, Target)
    Dim I, Cols(), Continue
    If TypeName(MdsT.RecordLoopFromDataarea(".")) <> "Error" Then
        Do
            Continue = MdsT.RecordLoopGetNext_VBS(False, Cols)
            If Cols(N) <> "#N/A" Then
                For I = 1 To N
                    If I = DI Then
                        FN.Write Target & vbTab
                    Else
                        FN.Write Cols(I - 1) & vbTab
                    End If
                Next
                FN.WriteLine Cols(N)
            End If
        Loop While Continue
    End If
End Sub

Sub ImportFreezeSelection(Model, Cube, Sel(), FileName, Target )
    PrepareFreezeDataArea Model, Cube, Sel, Target, mdsOperatorNone
    CallMDS MdsT.DataareaSetValue(Null)
    CallMDS MdsT.DataareaDestroy
    Dim Columns(), I, N, FI, Line
    N = CallMDS(MdsT.TableDimensionsCount(Model, Cube))
    ReDim Columns(N - 1)
	Set FI = FS.OpenTextFile(FileName, 1)
	Do While Not(FI.AtEndOfStream)
		Line = FI.ReadLine
	    For I = 1 To N
	        Columns(I - 1) = GetField(Line, vbTab)
	    Next
		CallMDS MdsDc.DataPutValueEx_VBS(Model, Cube, Line, Columns)
	Loop
    On Error Resume Next
End Sub

Sub FillWithChildren(Matrix(), I, Model, Dimension, Parent, K)
    Dim J, E
    For J = 1 To CallMDS(MdsE.ElementChildrenCount(Model, Dimension, Parent))
        E = CallMDS(MdsE.ElementChildrenName(Model, Dimension, Parent, J))
        If CallMDS(MdsE.ElementChildrenCount(Model, Dimension, E)) = 0 Then
            Matrix(K, I) = E
            K = K + 1
        Else
            FillWithChildren Matrix, I, Model, Dimension, E, K
        End If
    Next
End Sub

Function CubeFromCalculation(Model, Calculation)
    CubeFromCalculation = CallMds(MdsA.ATableFieldGetValue(Model, dimTCalc, 1, Calculation, "Cube"))
End Function

Function IsCalculation(Model, dimName, elem)
	Dim Cube
	Cube = MdsA.ATableFieldGetValue(Model, dimTCalc, 1, elem, "Cube")
	if TypeName(Cube) = "String" Then
		If dimName = dimTCalc then
			IsCalculation = True
		Else
			IsCalculation = UCase(Mid(Cube, 2)) = UCase(Mid(dimName, 2)) 
		End If
	End If
End Function
