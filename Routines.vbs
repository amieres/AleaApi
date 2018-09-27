
Dim Mds, MdsS, MdsDC, MdsD, MdsE, MdsT, MdsA
Dim FS

Function GetFileSystemObject()
	If IsEmpty(FS) Then Set FS = CreateObject("Scripting.FileSystemObject")
	Set GetFileSystemObject = FS
End Function

Function Connect(Alias)
	InitAlea
	Dim V, Info, DB, Server
	Info = Application.GetDatasourceInfo(Alias)
	Server = ExtractElement(Info, "server")
	If UCASE(Server) = "LOCALHOST" THEN Server = "LOCAL"
	DB = Server + "/" + ExtractElement(Info, "database")
	Connect = DB
	V = MDSS.ServerConnectTicket(CStr(Connect), CStr(ExtractElement(Info, "ticket")))
	'V = MDSS.ServerConnectWin(CStr(Connect))
'MsgBox DB + " " + TypeName(V)
	If TypeName(V) = "Error" Then
		If MDS.MdsGetLastError <> 2019 Then
			CallMds V
		End If
	End If
End Function

Sub InitAlea
	if IsEmpty(MDS) Then
		Set Mds   = Application.GetAleaAPI
		Set MdsS  = Application.GetAleaObject(aoServer)
		Set MdsDC = Application.GetAleaObject(aoDataCell)
		Set MdsD  = Application.GetAleaObject(aoDimension)
		Set MdsE  = Application.GetAleaObject(aoElement)
		Set MdsT  = Application.GetAleaObject(aoCube)
		Set MdsA  = Application.GetAleaObject(aoAttribute)
	End If
End Sub

Function CallMDS(V)
    If TypeName(V) = "Error" Then
        Dim ErrNum, ErrDes
        ErrNum = MDS.MdsGetLastError
        If ErrNum <> 0 Then
            ErrDes = MDS.MdsError(ErrNum)
MsgBox ErrDes
            Err.Raise vbObjectError + ErrNum, "Alea", CStr(ErrNum) + ":" + ErrDes
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    CallMDS = V
End Function

Function GetCubeName(Model, Calculation)
	GetCubeName = CallMDS(MdsA.ATableFieldGetValue(Model, "TCalc", 1, Calculation, "Cube"))
End Function

Function GetField(L, Sep)
    Dim I
    I = InStr(L, Sep)
    If I = 0 Then
        GetField = L
        L = ""
    Else
        GetField = Left(L, I - 1)
        L = Mid(L, I + Len(Sep))
    End If
End Function

Function GetElementName(Elm)
    Dim e, v
	e = Elm
    if Right(e,4) = ".[1]" then e = left(e,len(e)-4)
    Do While e <> ""
		v = GetField(e, "].[") 
	Loop
	GetElementName = GetField(v, "]")
End Function

Function ExtractElement(Xml, Element)
	Dim S, E
	S = InStr(Xml, "<" + Element + ">") + Len(Element) + 2
	E = InStr(S, Xml, "</" + Element + ">")
	ExtractElement = Mid(Xml, S, E - S)
End Function

Function IIf(Expression, TruePart, FalsePart)
	If Expression = True Then
		If IsObject(TruePart) Then
			Set IIf = TruePart
		Else
			IIf = TruePart
		End If
	Else
		If IsObject(FalsePart) Then
			Set IIf = FalsePart
		Else
			IIf = FalsePart
		End If
	End If
End Function
