'If Not IsObject(Application) Then
'   Set SapGuiAuto = GetObject("SAPGUI")
'   Set Application = SapGuiAuto.GetScriptingEngine
'End If
'If Not IsObject(Connection) Then
'   Set Connection = Application.Children(0)
'End If
'If Not IsObject(session) Then
'   Set session = Connection.Children(0)
'End If
'If IsObject(WScript) Then
'   WScript.ConnectObject session, "on"
'   WScript.ConnectObject Application, "on"
'End If

Dim WsShell
Set WsShell = CreateObject("WScript.shell")
Dim SessionNumber  

strComputer = "."

Set objNetwork = CreateObject("Wscript.Network")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

 ''' Processo que será verificado '''''''
 Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'wscript.exe'")
cdr = 0
 ''' elimina o processo definido '''
 For each Processo in ColProcesses
	cdr = cdr + 1
	If cdr > 1 then
		result = MsgBox ("O programa ja esta rodando, deseja fechar?", vbYesNo, "Yes No Example")
		Select Case result
		Case vbYes
			For each ProcessoWS in ColProcesses
				ProcessoWs.Terminate()
			Next
		Case vbNo
			WScript.Quit
		End Select
	end if
 Next

If Not IsObject(application) Then
   
	on error resume next
	Set SapGuiAuto  = GetObject("SAPGUI")

		if Err.Number <> 0 then
			WsShell.Run "https://www.myweg.net/irj/portal?NavigationTarget=pcd:portal_content/net.weg.folder.weg/net.weg.folder.core/net.weg.folder.roles/net.weg.role.ecc/net.weg.iview.ecc"
			WScript.Sleep 5000
			FreakoutAndFixTheError
			err.clear
			WScript.Sleep 5000
			If Not IsObject(application) Then
				Set SapGuiAuto  = GetObject("SAPGUI")
				Set application = SapGuiAuto.GetScriptingEngine
			End If
			
			If Not IsObject(connection) Then
			
				on error resume next
				Set connection = application.Children(0)
				
				if Err.Number <> 0 then
				msgbox("Favor iniciar o SAP antes de execultar.")
				Wscript.quit
				end if
			
			end if
		end if
	Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
	
	on error resume next
    Set connection = application.Children(0)

	if Err.Number <> 0 then
		WsShell.Run "https://www.myweg.net/irj/portal?NavigationTarget=pcd:portal_content/net.weg.folder.weg/net.weg.folder.core/net.weg.folder.roles/net.weg.role.ecc/net.weg.iview.ecc"
		WScript.Sleep 5000
		FreakoutAndFixTheError
		err.clear
		WScript.Sleep 5000
		If Not IsObject(application) Then
			Set SapGuiAuto  = GetObject("SAPGUI")
			Set application = SapGuiAuto.GetScriptingEngine
		End If
		
		If Not IsObject(connection) Then
		
			on error resume next
			Set connection = application.Children(0)
			
			if Err.Number <> 0 then
			msgbox("Favor iniciar o SAP antes de execultar.")
			Wscript.quit
			end if
		
		end if
	end if
	
End If

If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

	on error resume next
	session.findById("wnd[0]").sendVKey 0
	FreakoutAndFixTheError
	err.clear
	
SessionNumber = connection.Children.Count-1

for i = 0 to SessionNumber
	Set session = connection.sessions.Item(Cint(i))
	Wscript.Sleep 100	
	if session.Info.Transaction = "SESSION_MANAGER" then
		sessionok = 1
		exit for
	end if
next

if sessionok <> 1 then

	do While connection.Children.Count > 5
		Response = MsgBox("Favor fechar uma Janela do SAP para Continuar!", vbRetryCancel, "Aviso")
		If Response = vbCancel then
			WScript.Quit
		end if
	loop
	
	session.createSession
	Wscript.Sleep 1000
	SessionNumber = connection.Children.Count-1
	Wscript.Sleep 1000
	for i = 0 to SessionNumber
		Set session = connection.sessions.Item(Cint(i))
		Wscript.Sleep 100	
		if session.Info.Transaction = "SESSION_MANAGER" then
			sessionok = 1
			exit for
		end if
	next
	
end if

'msgbox(session.Children.Count)
'do While connection.Children.Count > 5
'Response = MsgBox("Favor fechar uma Janela do SAP para Continuar!", vbRetryCancel, "Aviso")
'If Response = vbCancel then
'WScript.Quit
'end if

'loop
'session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
'set GRID = session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
'GRID.setCurrentCell 0,""
'ROW = grid.currentCellRow
'Datte = GRID.getcellvalue (ROW, "DAT00")
										


Dim objDesktop
Dim objFolder
Dim FileSys
Dim colFiles
Dim objFile
Dim sinput
Dim ii
Dim i
Dim Item
Dim Item2
Dim ib
Dim bb
Dim ic
Dim nulin
Dim ulinha
Dim message, sapi


UserName = WsShell.ExpandEnvironmentStrings("%UserName%")
Set FileSys = CreateObject("Scripting.FileSystemObject")
Set FileShell = CreateObject("Shell.Application")
'Set objDesktop = WsShell.SpecialFolders(desk)
'Path2 = Path & "nomeantigo.xlsx"
'FileSys.GetFile(Path2).Name = "nomenovo.xlsx"
'FileSys.FileExists(ender)
'WsShell.Run "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\Scripts\speak.vbs"
nVar = InputBox(" Bem Vindo Sr(a). " & UserName  _
& vbNewLine & " __________________________________________ "_
& vbNewLine & " Insira um numero de opeção para execultar: " _
& vbNewLine & "  " _
& vbNewLine & " ( 1 ) - Remesas " _
& vbNewLine & " ( 2 ) - Ordens " _
& vbNewLine & " ( 3 ) - Diagramas " _
& vbNewLine & " ( 4 ) - Remesa unico " _
& vbNewLine & " ( 5 ) - Ordem unico " _
& vbNewLine & " ( 6 ) - Diagrama unico " _
& vbNewLine & " ( 7 ) - Faltante Ordens " _
& vbNewLine & " ( 8 ) - Faltante Diagrama " _
& vbNewLine & " ( 9 ) - Relatório Inspeção 1306 " _
& vbNewLine & " ( 10 ) - Relatório Inspeção 1321 " _
& vbNewLine & " ( 11 ) - Mapa de Processo IA " _
& vbNewLine & " ( 12 ) - Mapa de Processo IB" _
& vbNewLine & " ( 13 ) - Ordens Serralheria" _
& vbNewLine & " ( 14 ) - Materiais de Entrada Direta" _
& vbNewLine & " ( 15 ) - Contagem Materiais" _
& vbNewLine & " ( 16 ) - MB51 - Extração" _
& vbNewLine & " ( 17 ) - Solicitante N.F" _
& vbNewLine & " ( 18 ) - Em teste Guedes (nao usar)" _
& vbNewLine & " ( 19 ) - Materiais 180 dias sem movimento" _
& vbNewLine & " ( 20 ) - Diagramas Serralheria" _
& vbNewLine & " ( 22 ) - Faltante Ordens c/ Remessa" _
& vbNewLine & " ( 23 ) - Faltante Diagrama c/ Remessa" _
& vbNewLine & " __________________________________________ "_
& vbNewLine & " "_
  , "Execultaveis", "")
'WsShell.Run "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\Scripts\speak2.vbs"

'###########################################################################################################

If nVar = 1 Then

sinput = InputBox("Insira a data Limit da Busca", "Data", "" & Date & "")

Datainicial = Date - 15
DataFinal = Date

If sinput = "" Then
Else


Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Remessas\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Remessas"




If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZTSD092"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "HELIOCD"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 1, "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = Day(Datainicial) & "." & Month(Datainicial) & "." & Year(Datainicial)
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").Text = Day(DataFinal) & "." & Month(DataFinal) & "." & Year(DataFinal)


session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate

With objExcel
.Columns("A").Delete
.Columns("C").Delete
.Range("1:8").Delete
.Range("2:2").Delete
.Columns("C").Replace ".", "/"
.Range("A:E").AutoFilter 3, "=" & sinput
'testelinha = .Range("A1").End(xlDown)).Row
'.Range("A2:A" & testelinha).copy
.Range("A2:A999999").Copy
End With

session.findById("wnd[0]/tbar[0]/okcd").Text = "VL06O"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btnBUTTON6").press
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press

session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/WIJ REMESSA"
session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 12
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[14,2]").setFocus
session.findById("wnd[3]/usr/lbl[14,2]").caretPosition = 9
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").SetFocus
'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").caretPosition = 15
'session.findById("wnd[1]").sendVKey 2

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "remessas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "remessas.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("remessas").Activate
objExcel.Range("1:1").Delete
objExcel.Range("2:2").Delete

coluna = objExcel.Range("E1")
If coluna = Empty Or coluna = "" Then
objExcel.Range("F2:F999999").Copy
Else
objExcel.Range("E2:E999999").Copy
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "mb25"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

session.findById("wnd[0]/usr/ctxtBDTER-LOW").Text = Day(Datainicial) & "." & Month(Datainicial) & "." & Year(Datainicial)
session.findById("wnd[0]/usr/ctxtBDTER-HIGH").Text = Day(DataFinal) & "." & Month(DataFinal) & "." & Year(DataFinal)

session.findById("wnd[0]/usr/ctxtBDTER-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtBDTER-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").expandNode "          1"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         13"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "          6"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         13"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").Text = "RS01"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").SetFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "faltantes.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

If coluna = Empty Or coluna = "" Then
objExcel.Range("F2:F999999").Copy
Else
objExcel.Range("E2:E999999").Copy
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "MB52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/VBSWIJALMOX"
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1306"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "RS01"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "estoque.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

If coluna = Empty Or coluna = "" Then
objExcel.Range("F2:F999").Copy
Else
objExcel.Range("E2:E999").Copy
End If

session.findById("wnd[0]/tbar[0]/okcd").text = "ZTMM091"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "WELLINGTONFC"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 12
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "locais.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionRomaneio.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True


objExcel.Run "estoque"

'Dim dtewait
'dtewait = DateAdd("M",2,Now())
'do Until (Now() > dtewait)
'Loop

On Error Resume Next
objExcel.Range("A1").Copy

objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

ender = "C:\Users\" + UserName + "\Desktop\Remessas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
End If

End If
'###########################################################################################################
If nVar = 2 Then

sinput = InputBox("Insira a data Limit da Busca", "Data", "" & Date+1 & "")
If sinput = "" Then
Else


Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Ordens\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Ordens"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "coois"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOO000"
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "heliocd"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate

With objExcel
.Columns("A:B").Delete
.Range("1:3").Delete
.Range("2:2").Delete
.Columns("C").Replace ".", "/"
.Range("A:E").AutoFilter 3, "<=" & Month(sinput) & "/" & Day(sinput) & "/" & Year(sinput)
'testelinha = .Range("A1").End(xlDown)).Row
'.Range("A2:A" & testelinha).copy
.Range("A2:A999").Copy
End With

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]").resizeWorkingPane 267, 38, False
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOM000"

session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/WAU-PCP"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").SetFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 8
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

FileSys.DeleteFile (endereco)


session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ops.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]").resizeWorkingPane 267, 38, False
session.findById("wnd[0]/tbar[0]/btn[12]").press

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionOrdens.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Gerar_OPs"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

LC = Path + "ops.xls"
FileSys.DeleteFile (LC)

ender = "C:\Users\" + UserName + "\Desktop\Ordens"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing

End If
End If
'###########################################################################################################
If nVar = 3 Then



Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Diagramas\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Diagramas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").Maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "cm03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txt[35,3]").Text = "04030349"
session.findById("wnd[0]/usr/txt[35,5]").Text = ""
session.findById("wnd[0]/usr/txt[35,3]").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/chk[1,7]").selected = true
session.findById("wnd[0]/usr/chk[1,8]").selected = true
session.findById("wnd[0]/usr/chk[1,33]").selected = true
session.findById("wnd[0]/usr/chk[1,34]").selected = true
session.findById("wnd[0]/usr/chk[1,34]").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[49,1]").SetFocus
session.findById("wnd[0]/usr/lbl[49,1]").caretPosition = 38
session.findById("wnd[0]").sendVKey 20

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate

With objExcel
.Columns("A").Delete
.Columns("B").Delete
.Columns("C:F").Delete
.Range("1:5").Delete
.Range("2:3").Delete
 
.Columns("C").Replace ".", "/"
.Range("A:k").AutoFilter 9, "=LIB"
'testelinha = .Range("A1").End(xlDown)).Row
'.Range("A2:A" & testelinha).copy
.Range("B2:B999").Copy
End With

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

session.findById("wnd[0]/tbar[0]/okcd").Text = "mb26"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = ""
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

FileSys.DeleteFile (endereco)

session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 0
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "diagramas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionDiagramas.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Gerar_OPs"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

LC = Path + "diagramas.xls"
FileSys.DeleteFile (LC)

ender = "C:\Users\" + UserName + "\Desktop\Diagramas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder(ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
End If
'###########################################################################################################
If nVar = 4 Then


Datainicial = Date - 15
DataFinal = Date

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Remessas\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Remessas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "VL06O"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btnBUTTON6").press
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press

For i = 0 To 50

sinput = InputBox("Insira a Remessa", "Remessa")

If sinput = "" Then Exit For
If i >= 8 Then
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = sinput
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 2
Else
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").Text = sinput
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").caretPosition = 2
End If

If i >= 7 Then
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = (i - 6)
End If

Next

session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press

session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/WIJ REMESSA"
session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 12
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[14,2]").setFocus
session.findById("wnd[3]/usr/lbl[14,2]").caretPosition = 9
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").SetFocus
'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").caretPosition = 15
'session.findById("wnd[1]").sendVKey 2


session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "remessas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "remessas.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("remessas").Activate
objExcel.Range("1:1").Delete
objExcel.Range("2:2").Delete

If objExcel.Range("E2") = "" or objExcel.Range("E1") = "" Then
objExcel.Range("F2:F999999").Copy
Else
objExcel.Range("E2:E999999").Copy
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "mb25"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtBDTER-LOW").Text = Day(Datainicial) & "." & Month(Datainicial) & "." & Year(Datainicial)
session.findById("wnd[0]/usr/ctxtBDTER-HIGH").Text = Day(DataFinal) & "." & Month(DataFinal) & "." & Year(DataFinal)
session.findById("wnd[0]/usr/ctxtBDTER-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtBDTER-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").expandNode "          1"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         13"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "          6"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         13"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").Text = "RS01"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").SetFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "faltantes.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

If objExcel.Range("E2") = "" or objExcel.Range("E1") = "" Then
objExcel.Range("F2:F999999").Copy
Else
objExcel.Range("E2:E999999").Copy
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "MB52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/VBSWIJALMOX"
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1306"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "RS01"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "estoque.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


If objExcel.Range("E2") = "" or objExcel.Range("E1") = "" Then
objExcel.Range("F2:F999999").Copy
Else
objExcel.Range("E2:E999999").Copy
End If

session.findById("wnd[0]/tbar[0]/okcd").text = "ZTMM091"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "WELLINGTONFC"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 12
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "locais.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionRomaneio.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True


objExcel.Run "estoque"

'Dim dtewait
'dtewait = DateAdd("M",2,Now())
'do Until (Now() > dtewait)
'Loop

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

ender = "C:\Users\" + UserName + "\Desktop\Remessas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing

End If

'###########################################################################################################
If nVar = 5 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Ordens\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Ordens"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If


'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "coois"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOM000"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/WAU-PCP"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").SetFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press

For i = 0 To 50

sinput = InputBox("Insira a Ordem", "Ordem")

If sinput = "" Then Exit For
If i >= 8 Then
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = sinput
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 2
Else
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").Text = sinput
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").caretPosition = 2
End If

If i >= 7 Then
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = (i - 6)
End If

Next

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ops.xls"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionOrdens.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True


objExcel.Run "Gerar_OPs"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

LC = Path + "ops.xls"
FileSys.DeleteFile (LC)

ender = "C:\Users\" + UserName + "\Desktop\Ordens"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
End If
'###########################################################################################################
If nVar = 6 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Diagramas\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Diagramas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "MB26"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = ""
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press

For i = 0 To 50

sinput = InputBox("Insira o Diagrama", "Diagrama")

If sinput = "" Then Exit For
If i >= 8 Then
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = sinput
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 2
Else
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").Text = sinput
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").caretPosition = 2
End If

If i >= 7 Then
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = (i - 6)
End If

Next

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 0
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "diagramas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionDiagramas.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Gerar_OPs"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

LC = Path + "diagramas.xls"
FileSys.DeleteFile (LC)

ender = "C:\Users\" + UserName + "\Desktop\Diagramas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
End If
'###########################################################################################################
If nVar = 7 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\OrdensFalt\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\OrdensFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "coois"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOM000"
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Falt_Ordens.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]").resizeWorkingPane 272, 40, False
session.findById("wnd[0]/tbar[0]/btn[12]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "Falt_Ordens.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("Falt_Ordens").Activate

With objExcel
.Columns("A:A").Delete
.Range("1:3").Delete
.Range("2:2").Delete
.Columns("I").Replace ".", "/"
.Range("B2:B999").Copy
End With

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "mb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1306"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "RS01"
session.findById("wnd[0]/usr/chkNOZERO").Selected = False
session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/VBSWIJALMOX"
session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Est_Falt_Ordens.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionFaltOrdens1.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True


objExcel.Run "Deposito"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

ender = "C:\Users\" + UserName + "\Desktop\OrdensFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
End If
'###########################################################################################################
If nVar = "Old08" Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\DiagramasFalt\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "cm03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txt[35,3]").Text = "04030349"
session.findById("wnd[0]/usr/txt[35,5]").Text = ""
session.findById("wnd[0]/usr/txt[35,3]").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/chk[1,7]").Selected = True
'session.findById("wnd[0]/usr/chk[1,8]").selected = true
session.findById("wnd[0]/usr/chk[1,33]").Selected = True
'session.findById("wnd[0]/usr/chk[1,34]").selected = true
session.findById("wnd[0]/usr/chk[1,34]").SetFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[49,1]").SetFocus
session.findById("wnd[0]/usr/lbl[49,1]").caretPosition = 38
session.findById("wnd[0]").sendVKey 20

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate

With objExcel
.Range("1:5").Delete
.Range("2:3").Delete
.Columns("A").Delete
.Columns("B").Delete
.Columns("C:F").Delete
.Range("A:k").AutoFilter 9, "<>LIB"
.Range("B2:B999").Copy
End With

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

session.findById("wnd[0]/tbar[0]/okcd").Text = "MB26"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "Lguedes"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

FileSys.DeleteFile (endereco)

session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select

'############################################## INICIO Seleção layout #################################
'session.findById("wnd[0]/tbar[1]/btn[33]").press
'session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 90
'session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "90"
'session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DiagramasFalt.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionFaltDiagrama.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Ordens"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit


ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
Else

End If
'###########################################################################################################
If nVar = 9 Then

Datainput = InputBox("Insira a data Limit da Busca", "Data", Day(Date) & "." & Month(Date) & "." & Year(Date))
Data = "\" & Datainput & "\"
Path = "C:\Users\" + UserName + "\Desktop\Relatorios\"
ender = "C:\Users\" + UserName + "\Desktop\Relatorios"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZTQM018"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkP_SKIP").selected = true
session.findById("wnd[0]/usr/chkP_REIMP").selected = false
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtSO_HERK-LOW").text = "01"
session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = Datainput
session.findById("wnd[0]/usr/ctxtSO_ENST-HIGH").text = Datainput
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CRM-WIJ"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"LIFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LIFNR"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "10000"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"EBELN"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELN"
session.findById("wnd[0]/tbar[1]/btn[28]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "STATUS"
session.findById("wnd[0]/tbar[1]/btn[40]").press

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1306.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


Set objExcel = CreateObject("Excel.Application")
endereco = Path + "1306.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("1306").Activate


objExcel.Columns("A:B").Delete
objExcel.Range("1:3").Delete
objExcel.Range("2:2").Delete
ulinha = objExcel.Range("K65536").End(-4162).Row
'objExcel.Range("A:V").AutoFilter 9, ">9999"

for i = 2 to ulinha

if objExcel.Range("P" & i) = "LIB AMOS RNEC QLRE" or objExcel.Range("P" & i) = "LIB AMOS RNEC" or objExcel.Range("P" & i) = "LIB AMOS" or objExcel.Range("P" & i) = "LIB AMOS QLRE" then

Item = objExcel.Range("K" & i)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZTQM018"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkP_SKIP").selected = true
session.findById("wnd[0]/usr/chkP_REIMP").selected = false
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtSO_HERK-LOW").text = "01"
session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = Datainput
session.findById("wnd[0]/usr/ctxtSO_ENST-HIGH").text = Datainput
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CRM-WIJ"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"LIFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LIFNR"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "10000"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"EBELN"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELN"
session.findById("wnd[0]/tbar[1]/btn[28]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "STATUS"
session.findById("wnd[0]/tbar[1]/btn[40]").press

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


if objExcel.Range("K" & i) <> objExcel.Range("K" & i+1) then
session.findById("wnd[0]/tbar[0]/okcd").text = "MB90"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRG_KSCHL-LOW").text = "WE01"
'session.findById("wnd[0]/usr/ctxtPM_VERMO").text = "2"
session.findById("wnd[0]/usr/txtRG_MBLNR-LOW").text = Item
session.findById("wnd[0]/usr/txtRG_MBLNR-LOW").setFocus
session.findById("wnd[0]/usr/txtRG_MBLNR-LOW").caretPosition = 10
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[5]").press
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[14]").press
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[86]").press
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press




end if

end if

next

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

objExcel.application.Application.DisplayAlerts = False
objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Quit

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

end if

'###########################################################################################################
If nVar = 10 Then

Datainput = InputBox("Insira a data Limit da Busca", "Data", Day(Date) & "." & Month(Date) & "." & Year(Date))
Data = "\" & Datainput & "\"
Path = "C:\Users\" + UserName + "\Desktop\Relatorios\"
ender = "C:\Users\" + UserName + "\Desktop\Relatorios"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If


'session.createSession

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZTQM018"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkP_SKIP").selected = true
session.findById("wnd[0]/usr/chkP_REIMP").selected = false
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").text = "1321"
session.findById("wnd[0]/usr/ctxtSO_HERK-LOW").text = "01"
session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = Datainput
session.findById("wnd[0]/usr/ctxtSO_ENST-HIGH").text = Datainput
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CRM-WIJ"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"LIFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LIFNR"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "10000"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"EBELN"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELN"
session.findById("wnd[0]/tbar[1]/btn[28]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "STATUS"
session.findById("wnd[0]/tbar[1]/btn[40]").press

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1321.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


Set objExcel = CreateObject("Excel.Application")
endereco = Path + "1321.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("1321").Activate


objExcel.Columns("A:B").Delete
objExcel.Range("1:3").Delete
objExcel.Range("2:2").Delete
ulinha = objExcel.Range("K65536").End(-4162).Row
'objExcel.Range("A:V").AutoFilter 9, ">9999"

for i = 2 to ulinha

if objExcel.Range("P" & i) = "LIB AMOS RNEC QLRE" or objExcel.Range("P" & i) = "LIB AMOS RNEC" or objExcel.Range("P" & i) = "LIB AMOS" or objExcel.Range("P" & i) = "LIB AMOS QLRE" then

Item = objExcel.Range("K" & i)


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZTQM018"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkP_SKIP").selected = true
session.findById("wnd[0]/usr/chkP_REIMP").selected = false
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").text = "1321"
session.findById("wnd[0]/usr/ctxtSO_HERK-LOW").text = "01"
session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = Datainput
session.findById("wnd[0]/usr/ctxtSO_ENST-HIGH").text = Datainput

session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CRM-WIJ"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"LIFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LIFNR"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "10000"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"EBELN"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELN"
session.findById("wnd[0]/tbar[1]/btn[28]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "STATUS"
session.findById("wnd[0]/tbar[1]/btn[40]").press

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


if objExcel.Range("K" & i) <> objExcel.Range("K" & i+1) then
session.findById("wnd[0]/tbar[0]/okcd").text = "MB90"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtRG_KSCHL-LOW").text = "WE01"
'session.findById("wnd[0]/usr/ctxtPM_VERMO").text = "2"
session.findById("wnd[0]/usr/txtRG_MBLNR-LOW").text = Item
session.findById("wnd[0]/usr/txtRG_MBLNR-LOW").setFocus
session.findById("wnd[0]/usr/txtRG_MBLNR-LOW").caretPosition = 10

On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[5]").press

On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[14]").press
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[86]").press
On Error Resume Next

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press





end if

end if

next

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

objExcel.application.Application.DisplayAlerts = False
objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Quit

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

end if

'###########################################################################################################
If nVar = 11 Then

Path = "C:\Users\" + UserName + "\Desktop\Mapas IA"

If FileSys.FolderExists(Path) Then
    FileSys.DeleteFolder (Path)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "mf60"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell


session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "ia01"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MAPA-IA01-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IA02"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MAPA-IA02-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IA03"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MAPA-IA03-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/usr/ctxtPA_PLANT").text = "1306"
session.findById("wnd[0]/usr/ctxtPA_PLANT").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IA01"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MAPA-IA01-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IA02"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MAPA-IA02-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IA03"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MAPA-IA03-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


end if

'###########################################################################################################
If nVar = 12 Then

Path = "C:\Users\" + UserName + "\Desktop\Mapas IB"

If FileSys.FolderExists(Path) Then
    FileSys.DeleteFolder (Path)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "mf60"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtPA_PLANT").text = "1321"

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IB05"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4

session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press



'===============================================================Novo====================================================================================
session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1-MAPA-IB05-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press

on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/tbar[0]/btn[16]").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 18
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
'===============================================================Novo====================================================================================



session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "2-MAPA-IB05-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

session.findById("wnd[0]/usr/ctxtPA_PLANT").text = "1306"
session.findById("wnd[0]/usr/ctxtPA_PLANT").caretPosition = 4
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IB05"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4

session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press


'===============================================================Novo====================================================================================
session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1-MAPA-IB05-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press

on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/tbar[0]/btn[16]").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 18
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
'===============================================================Novo====================================================================================




session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "2-MAPA-IB05-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/tbar[0]/btn[12]").press

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------




session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "mf60"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtPA_PLANT").text = "1321"

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IB01"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4

session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press



'===============================================================Novo====================================================================================
session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1-MAPA-IB01-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press

on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/tbar[0]/btn[16]").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 18
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
'===============================================================Novo====================================================================================


session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "2-MAPA-IB01-1321.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press


on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

session.findById("wnd[0]/usr/ctxtPA_PLANT").text = "1306"
session.findById("wnd[0]/usr/ctxtPA_PLANT").caretPosition = 4
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").text = "IB01"
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ENJOY_PARA/tabpTAB_FERT/ssub%_SUBSCREEN_ENJOY_PARA:RMPU_SEL_SCREEN:0200/ctxtSO_LGORT-LOW").caretPosition = 4

session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select


session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = false
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/MAPA_ALMOX"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[1,2]").setFocus
session.findById("wnd[3]/usr/lbl[1,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

'===============================================================Novo====================================================================================

session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press



session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1-MAPA-IB01-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press

on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/usr/lbl[6,9]").setFocus
session.findById("wnd[0]/usr/lbl[6,9]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/tbar[0]/btn[16]").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "13797431"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "13055118"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "13797765"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14122063"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "10192443"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "13644089"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "14122064"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 1
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 3
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 4
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 5
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 6
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "14410878"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "12051824"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "10453138"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "14410103"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "12056499"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "10192355"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 7
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 9
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 10
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 11
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 12
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 18
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "10412098"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "14610864"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "14610866"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "11156164"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "11156171"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

'===============================================================Novo====================================================================================

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "2-MAPA-IB01-1306.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press



on error resume next
session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
on error resume next
session.findById("wnd[1]/tbar[0]/btn[13]").press

'on error resume next
'session.findById("wnd[0]/tbar[0]/btn[86]").press
'on error resume next
'session.findById("wnd[1]/tbar[0]/btn[13]").press
on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


session.findById("wnd[0]/tbar[0]/btn[12]").press

'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


end if

'###########################################################################################################
If nVar = 13 Then

Path = "C:\Users\" + UserName + "\Desktop\MapaOrdenSerralheria\"
ender = "C:\Users\" + UserName + "\Desktop\MapaOrdenSerralheria"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If


sinput = InputBox("Insira a data Limit da Busca", "Data", Day(Date)+1 & "." & Month(Date) & "." & Year(Date))

'session.createSession
session.findById("wnd[0]").maximize

session.findById("wnd[0]/tbar[0]/okcd").text = "Coois"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "lguedes"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 3
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"FSSBD"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "FSSBD"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&MB_FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").text = sinput
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").setFocus
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"VSTTXT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "VSTTXT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&MB_FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN004_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "LIB"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "IMPR LIB"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 3
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

'msgbox("oi")
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = -1
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "AUFNR"
session.findById("wnd[0]/tbar[0]/btn[12]").press


'########################################################### E X C E L L ##########################################
Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate


objExcel.Columns("A").Delete
objExcel.Columns("B").Delete
objExcel.Range("1:5").Delete
objExcel.Range("2:2").Delete
objExcel.Range("A2:A999999").Copy
'########################################################### E X C E L L ##########################################



session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/SEP-SERRAL"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").setFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 11

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"LGORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "LGORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&MB_FILTER"

session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = ""
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 0
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


'########################################################### E X C E L L ##########################################
objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit
'########################################################### E X C E L L ##########################################


session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"AUFNR"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "AUFNR"
session.findById("wnd[0]/tbar[0]/btn[12]").press

'########################################################### E X C E L L ##########################################
Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate


objExcel.Columns("A").Delete
objExcel.Range("1:3").Delete
objExcel.Range("2:2").Delete
objExcel.Range("A2:A999999").Copy
'########################################################### E X C E L L ##########################################

session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOO000"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/SEP-SERRAL"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

'########################################################### E X C E L L ##########################################
objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit
'########################################################### E X C E L L ##########################################


session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press


'########################################################### E X C E L L ##########################################
Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate


objExcel.Columns("A").Delete
objExcel.Columns("B").Delete
objExcel.Range("1:3").Delete
objExcel.Range("2:2").Delete
objExcel.Range("$B$2:$B$999999").copy
objExcel.Range("$E$1:$E$999999").PasteSpecial -4163

objExcel.Columns("E").RemoveDuplicates 1,2


linha = objExcel.Range("E65536").End(-4162).Row
ulinha = objExcel.Range("A65536").End(-4162).Row

'########################################################### E X C E L L ##########################################

for i = 1 to linha


session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"ARBPL"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "ARBPL"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&MB_FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "0" & objExcel.Range("E" & i)
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&PRINT_BACK"
session.findById("wnd[1]/tbar[0]/btn[13]").press

on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press

next

session.findById("wnd[0]/tbar[0]/btn[12]").press


for ib = 2 to ulinha

if objExcel.Range("B"& ib) <> objExcel.Range("B" & ib - 1) then 

nulin = 0

for bb = 1 to 99

if objExcel.Range("B" & ib) = objExcel.Range("B" & ib + bb) and objExcel.Range("B" & ib + bb) <> "" then
 nulin = nulin + 1
end if

next

objExcel.Range("A" & ib & ":A" & ib + nulin).copy



session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/SEP-SERRAL"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").setFocus
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 11
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"PLPLA"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "PLPLA"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&SORT_ASC"

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&PRINT_BACK"
session.findById("wnd[1]/tbar[0]/btn[13]").press

on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press

for ic = ib to (ib + nulin)

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"AUFNR"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "AUFNR"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&MB_FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = objExcel.Range("A" & ic)
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").setCurrentCell -1,"PLPLA"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn "PLPLA"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&SORT_ASC"


session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&PRINT_BACK"
session.findById("wnd[1]/tbar[0]/btn[13]").press

on error resume next
session.findById("wnd[2]/tbar[0]/btn[0]").press

next

session.findById("wnd[0]/tbar[0]/btn[12]").press

end if

next

session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------



'########################################################### E X C E L L ##########################################
objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit
'########################################################### E X C E L L ##########################################

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

end if
'###########################################################################################################
If nVar = 14 Then

nVar = InputBox("Insira a data Limit da Busca", "Data", Day(Date) & "." & Month(Date) & "." & Year(Date))

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\EntradaMaterial\"
Path2 = "C:\Users\" + UserName + "\Desktop\EntradaMaterialDireto\"
ender = "C:\Users\" + UserName + "\Desktop\EntradaMaterial"


If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZTQM018"
session.findById("wnd[0]").sendVKey 0

On error resume next
session.findById("wnd[0]/usr/chkP_SKIP").selected = true
On error resume next
session.findById("wnd[0]/usr/chkP_REIMP").selected = true

session.findById("wnd[0]/usr/ctxtSO_HERK-LOW").text = "01"
session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = nVar
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CRM-WIJ"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
session.findById("wnd[0]/usr/btn%_SO_WERK_%_APP_%-VALU_PUSH").press
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SO_WERK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 4
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "PAULOANTONIO"
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"

session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_LIFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_KUNNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_MAT_KDAU_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "101"
session.findById("wnd[0]/usr/ctxtBWART-LOW").setFocus
session.findById("wnd[0]/usr/ctxtBWART-LOW").caretPosition = 3
session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "109"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[8]").press

'session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RS01"
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RF01"
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
'session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = nVar
'session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = nVar
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export2.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]").resizeWorkingPane 267, 38, False

Set obj = CreateObject("Excel.Application")
obj.Application.Visible = True

endereco = Path + "export.xls"
endereco2 = Path + "export2.xls"

Set objExcel = obj.Workbooks.Open(endereco)
Set objExcelb = obj.Workbooks.Open(endereco2)

objExcel.Sheets("export").Activate

objExcel.Sheets("export").Columns("A:B").Delete
objExcel.Sheets("export").Range("1:3").Delete
objExcel.Sheets("export").Range("2:2").Delete



objExcelb.Sheets("export2").Activate

objExcelb.Sheets("export2").Columns("A:B").Delete
objExcelb.Sheets("export2").Columns("B").Delete
objExcelb.Sheets("export2").Range("1:5").Delete
objExcelb.Sheets("export2").Range("2:2").Delete
ulinha = objExcelb.Sheets("export2").Range("A65536").End(-4162).Row

objExcelb.Sheets("export2").Range("K1") = "Analise"
'objExcelb.Sheets("export2").Range("K2").Formula = "=IFERROR(VLOOKUP(RC[-10],'" & Path & "[export.xls]export'!C5,1,0),"""")"
objExcelb.Sheets("export2").Range("K2").Formula = "=IFERROR(VLOOKUP(RC[-10],export.xls!C5,1,0),""Direto"")"

objExcelb.Sheets("export2").Range("K2").copy
objExcelb.Sheets("export2").Range("K3:K" & ulinha).PasteSpecial -4123
objExcelb.Sheets("export2").Range("K:K").Copy
objExcelb.Sheets("export2").Range("K:K").PasteSpecial -4163


objExcelb.Sheets("export2").Range("A:O").AutoFilter 11, "=Direto"
objExcelb.Sheets("export2").Range("A:O").AutoFilter 4, "<>IB03"
objExcelb.Sheets("export2").Range("A:O").AutoFilter 10, "=PINFEUSER"
objExcelb.Sheets("export2").Range("A:O").Copy
objExcelb.Sheets.add()
objExcelb.Sheets("Planilha1").Range("A:O").PasteSpecial -4163
'objExcelb.Sheets("Planilha1").Range("A:O").RemoveDuplicates 13,2
ulinhab = objExcelb.Sheets("Planilha1").Range("A65536").End(-4162).Row

objExcelb.Sheets("Planilha1").Columns("A:O").EntireColumn.AutoFit
objExcelb.Sheets("Planilha1").Columns("E:F").EntireColumn.Hidden = True
objExcelb.Sheets("Planilha1").Columns("J:L").EntireColumn.Hidden = True
'objExcelb.Sheets("Planilha1").Columns("N:N").EntireColumn.Hidden = True
objExcelb.Sheets("Planilha1").Columns("H:H").NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
''objExcelb.Sheets("Planilha1").PageSetup.PrintTitleColumns = "A:O"

With objExcelb.Sheets("Planilha1").PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
End With


If not FileSys.FolderExists(Path2) Then
    FileSys.CreateFolder (Path2)
End If


objExcelb.Sheets("Planilha1").ExportAsFixedFormat xlTypePDF, Path2 & " " & nVar & ".pdf", 1, 1, 0, 1


''objExcelb.Sheets("Planilha1").PrintOut 1, 2, 1

objExcelb.Sheets("Planilha1").Range("A1").Copy
objExcelb.Saved = True
objExcelb.Close


objExcel.Sheets("export").Range("A1").Copy
objExcel.Saved = True
objExcel.Close
obj.Quit

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

session.findById("wnd[0]").maximize

FileShell.ShellExecute (Path2 & " " & nVar & ".pdf")

end if
'###########################################################################################################

If nVar = 100 Then

nVar = InputBox("Insira a data Limit da Busca", "Data", Day(Date)-1 & "." & Month(Date) & "." & Year(Date))

if nVar = Cancel then

else
Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "Q:\GROUPS\BR_SC_ITJ_WAU_LOGISTICA_ENGENHEIRADOS\02_ALMOXARIFADO\05_Job\Nao Mexer\Banco de dados\DB\"
enderecob = "Q:\GROUPS\BR_SC_ITJ_WAU_LOGISTICA_ENGENHEIRADOS\03_GESTÃO_ALMOXARIFADO\003 - Grafico Picking Diário3.xlsx"

Path2 = "Q:\GROUPS\BR_SC_ITJ_WAU_LOGISTICA_ENGENHEIRADOS\02_ALMOXARIFADO\05_Job\Nao Mexer\Banco de dados\"
FileSys.DeleteFile (Path2 + "Qualidade.xls")

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_LIFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_KUNNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_MAT_KDAU_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_USNAM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_VGART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_XBLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_BUDAT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

'session.findById("wnd[0]/usr/radRHIER_L").setFocus
'session.findById("wnd[0]/usr/radRHIER_L").select

session.findById("wnd[0]/usr/radRFLAT_L").setFocus
session.findById("wnd[0]/usr/radRFLAT_L").select

session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/WIJ_BAIXAS"

session.findById("wnd[0]/usr/ctxtPA_AISTR").text = "ZDOC_MATERIAL"
session.findById("wnd[0]/usr/chkDATABASE").selected = true
session.findById("wnd[0]/usr/chkSHORTDOC").selected = false
session.findById("wnd[0]/usr/chkARCHIVE").selected = false

session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").setFocus
session.findById("wnd[0]/usr/ctxtWERKS-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
'session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "281"
'session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "281"
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "261"
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 3
'session.findById("wnd[1]/tbar[0]/btn[8]").press

'session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "RS01"

session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RS01"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RF01"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = nVar
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = nVar
session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/PAULOA"
session.findById("wnd[0]/usr/ctxtPA_AISTR").text = "ZDOC_MATERIAL"
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press

'lines = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount

'session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select
'session.findById("wnd[1]/tbar[0]/btn[0]").press
'session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
'session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
'session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 29

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8

On error resume next
session.findById("wnd[1]/tbar[0]/btn[11]").press

On error resume next
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/tbar[0]/okcd").text = "mb26"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1306"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_S_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RS01"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RF01"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_BDTER-HIGH").text = nVar
session.findById("wnd[0]/usr/ctxtS_BDTER-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_BDTER-HIGH").caretPosition = 6
session.findById("wnd[0]/usr/btn%_S_VORNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "0001"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").text = "0003"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,2]").text = "0006"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,3]").text = "0065"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,4]").text = "0067"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,5]").text = "0240"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,6]").text = "0241"

session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,5]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
'session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]/mbar/menu[0]/menu[2]").select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

'session.findById("wnd[0]/tbar[1]/btn[38]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"VORNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "VORNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0006"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "0067"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0241"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"MEINS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "MEINS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "M"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 1
session.findById("wnd[1]").sendVKey 0

'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"ERFME"
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "ERFME"
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
'session.findById("wnd[0]/tbar[1]/btn[29]").press
'session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").text = "M"
'session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").caretPosition = 1
'session.findById("wnd[1]").sendVKey 0



caball = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"COWB_VMENG"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "COWB_VMENG"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN004-LOW").setFocus
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN004-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 3,"TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "3"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN004-LOW").text = "              0"
session.findById("wnd[1]/tbar[0]/btn[0]").press

cab1 = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount

session.findById("wnd[0]/tbar[1]/btn[38]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"VORNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "VORNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0065"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "0240"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0003"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "0001"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").caretPosition = 4
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"MEINS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "MEINS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "UN"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 1
session.findById("wnd[1]").sendVKey 0

'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"ERFME"
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "ERFME"
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
'session.findById("wnd[0]/tbar[1]/btn[29]").press
'session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").text = "UN"
'session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").caretPosition = 1
'session.findById("wnd[1]").sendVKey 0

linall = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"COWB_VMENG"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "COWB_VMENG"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN004-LOW").setFocus
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN004-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 3,"TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "3"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN004-LOW").text = "              0"
session.findById("wnd[1]/tbar[0]/btn[0]").press


lin = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount



session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/tbar[0]/okcd").text = "ZTQM018"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_SO_WERK_%_APP_%-VALU_PUSH").press
session.findById("wnd[0]/usr/chkP_REIMP").selected = true
session.findById("wnd[0]/usr/chkP_SKIP").selected = true
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtSO_HERK-LOW").text = "01"
session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = nVar
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CRM-WIJ"
session.findById("wnd[0]/usr/ctxtSO_WERK-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SO_WERK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

lin2 = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount

session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtSO_ENST-LOW").text = nVar
session.findById("wnd[0]/tbar[1]/btn[8]").press

lin3 = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


session.findById("wnd[0]/tbar[0]/okcd").text = "qa32"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "PAULOANTONIO"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 12
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtQL_ENSTD-LOW").text = nVar
session.findById("wnd[0]/usr/ctxtQL_ENSTD-HIGH").text = nVar
session.findById("wnd[0]/usr/ctxtQL_ENSTD-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtQL_ENSTD-HIGH").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "Q:\GROUPS\BR_SC_ITJ_WAU_LOGISTICA_ENGENHEIRADOS\02_ALMOXARIFADO\05_Job\Nao Mexer\Banco de dados"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Qualidade.xls"
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------

'msgbox("Foram separados " & lines & " Itens de " & lin & " Ordens/Diagramas")

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
'objExcel.Sheets("Sheet1").Activate


objExcel.Columns("A").Delete
objExcel.Range("1:3").Delete
objExcel.Range("2:2").Delete

objExcel.Range("P1") = "Total Picking"
objExcel.Range("P2") = lin
objExcel.Range("Q1") = "Picking"
objExcel.Range("Q2").Formula = "=COUNTIFS(C[-4],""281"",C[-6],""MARCOMORAES"")+COUNTIFS(C[-4],""281"",C[-6],""DAIANAALINE"")+COUNTIFS(C[-4],""281"",C[-6],""LUCCASF"")+COUNTIFS(C[-4],""281"",C[-6],""IRENE"")+COUNTIFS(C[-4],""281"",C[-6],""SUELENK"")+COUNTIFS(C[-4],""281"",C[-6],""LGUEDES"")+COUNTIFS(C[-4],""261"",C[-6],""MARCOMORAES"")+COUNTIFS(C[-4],""261"",C[-6],""DAIANAALINE"")+COUNTIFS(C[-4],""261"",C[-6],""LUCCASF"")+COUNTIFS(C[-4],""261"",C[-6],""IRENE"")+COUNTIFS(C[-4],""261"",C[-6],""SUELENK"")+COUNTIFS(C[-4],""261"",C[-6],""LGUEDES"")"

objExcel.Range("R1") = "Total Quality"
objExcel.Range("R2") = lin2 + lin3
objExcel.Range("S1") = "Quality"
objExcel.Range("S2").Formula = "=COUNTIF(C[-6],""321"")/2"

objExcel.Range("T1") = "Total"
objExcel.Range("T2").Formula = "=RC[-3]+RC[-4]+RC[5]"

objExcel.Range("U1") = "Total Picking cabos"
objExcel.Range("U2") = cab1
objExcel.Range("V1") = "Picking Cabos"
objExcel.Range("V2").Formula = "=COUNTIFS(C[-9],""281"",C[-11],""DEBORAP"")+COUNTIFS(C[-9],""281"",C[-11],""ELIASPC"")+COUNTIFS(C[-9],""281"",C[-11],""outro"")+COUNTIFS(C[-9],""261"",C[-11],""DEBORAP"")+COUNTIFS(C[-9],""261"",C[-11],""ELIASPC"")+COUNTIFS(C[-9],""261"",C[-11],""outro"")"
objExcel.Range("W1") = "Totalb"
objExcel.Range("W2").Formula = "=RC[-2]+RC[-1]+RC[1]"

objExcel.Range("X1") = "Falt. Cabos"
objExcel.Range("X2") = caball - cab1
objExcel.Range("Y1") = "Falt. Compo."
objExcel.Range("Y2") = linall - lin

'objExcel.ActiveWorkbook.Save
'const xlOpenXMLWorkbookMacroEnabled = 52
'const xlOpenXMLWorkbook = 51
FileSys.DeleteFile (Path + "export.xlsx")

objExcel.ActiveWorkbook.RefreshAll
objExcel.ActiveWorkbook.SaveAs Path + "export.xlsx", 51
objExcel.ActiveWorkbook.Close
objExcel.Quit

''result = MsgBox ("Deseja Enviar o Email?", vbYesNo, "Send Email")

''Select Case result
''Case vbYes

' xOutMsg = 	"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><span id=""j_id0:emailTemplate:j_id3:j_id4:j_id6"">" & _
			' "<html>" & _
			' "<body>" & _
			' "Bom dia<br/><br/>" & _
			' "Segue em anexo Graficos de Controle Almox "& nVar &".<br/><br/>" & _
			' "</body>" & _
			' "</html></span>" & _
			' "<span style=""color:#104E8B""><b>Almoxarifado RS01</b></span style=""color:#104E8B""><br/>" & _
			' "<span style=""color:#104E8B; font-size:12px; font-family:'Arial'"">Coordenacao Almoxarifado e Expedicao (Itajai)</span><br/>" & _
			' "<span style=""color:#104E8B; font-size:12px; font-family:'Arial'"">Fone:+55 (47) 3276 7350</span><br/>" & _
			' "<span style=""color:#104E8B; font-size:12px; font-family:'Arial'"">WEG Drives & Controls</span><br/>" & _
			' "<span style=""color:#0000CD; font-size:12px; font-family:'Arial'"">www.weg.net</span><br/>"
			
	' Set emailObj = CreateObject("CDO.Message")

	' Set emailConfig = emailObj.Configuration
	' emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.weg.net"
	' emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	' emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2 
	' emailConfig.Fields.Update

	' emailObj.From = "Almoxarifado Coordenação <almoxrs01@weg.net>" '"no-reply@weg.net"
	' emailObj.To = "heliocd@weg.net;lguedes@weg.net"
	' emailObj.Cc = "pauloantonio@weg.net;herold@weg.net"
	' emailObj.Subject = "Graficos de Controle Almox " & nVar
	' emailObj.HTMLBody = xOutMsg
	
	' emailObj.AddAttachment enderecob

	' emailObj.Send
	' Set emailObj = Nothing

''End Select

Set objExcelb = CreateObject("Excel.Application")
Set objWorkbookb = objExcelb.Workbooks.Open(enderecob)
objExcelb.Application.Visible = True

end if

end if
'###########################################################################################################

If nVar = 8 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\DiagramasFalt\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "cm03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txt[35,3]").Text = "04030349"
session.findById("wnd[0]/usr/txt[35,5]").Text = ""
session.findById("wnd[0]/usr/txt[35,3]").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/chk[1,7]").Selected = True
'session.findById("wnd[0]/usr/chk[1,8]").selected = true
session.findById("wnd[0]/usr/chk[1,33]").Selected = True
'session.findById("wnd[0]/usr/chk[1,34]").selected = true
session.findById("wnd[0]/usr/chk[1,34]").SetFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[49,1]").SetFocus
session.findById("wnd[0]/usr/lbl[49,1]").caretPosition = 38
session.findById("wnd[0]").sendVKey 20

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate

With objExcel
.Range("1:5").Delete
.Range("2:3").Delete
.Columns("A").Delete
.Columns("B").Delete
.Columns("C:F").Delete
.Range("A:k").AutoFilter 9, "=LIB"
.Range("B2:B999").Copy
End With

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

session.findById("wnd[0]/tbar[0]/okcd").Text = "MB26"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "Lguedes"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

'session.findById("wnd[0]").sendVKey 12

objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

FileSys.DeleteFile (endereco)

session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DiagramasFalt.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


'session.findById("wnd[0]/tbar[0]/btn[15]").press '------------------------Fecha Janela --------------------


Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionFaltDiagrama.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Ordens"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit


ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"


If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
Else

End If

'###########################################################################################################

If nVar = 15 Then

sinput = InputBox("Insira a data Limit da Busca", "Data", "01." & Month(Date) & "." & Year(Date))
sinputf = InputBox("Insira a data Limit da Busca", "Data", "31." & Month(Date) & "." & Year(Date))

Path = "Q:\GROUPS\BR_SC_ITJ_WAU_LOGISTICA_ENGENHEIRADOS\03_GESTÃO_ALMOXARIFADO\"


On Error Resume Next
FileSys.DeleteFile (Path + "Contagem Estoque.xls")

On Error Resume Next
FileSys.DeleteFile (Path + "Contagem Estoque.xlsx")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MI20"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "Herold"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/chkIM_SELKB").selected = true
session.findById("wnd[0]/usr/chkIM_SELKP").selected = true
session.findById("wnd[0]/usr/ctxtGIDAT-LOW").text = sinput
session.findById("wnd[0]/usr/ctxtGIDAT-HIGH").text = sinputf
session.findById("wnd[0]/usr/chkIM_SELKP").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Contagem Estoque.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'########################################################### E X C E L L ##########################################
Set objExcel = CreateObject("Excel.Application")
endereco = Path + "Contagem Estoque.xls"

Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("Contagem Estoque").Activate


objExcel.Columns("A").Delete
objExcel.Columns("B").Delete
objExcel.Range("1:1").Delete
objExcel.Range("2:2").Delete
objExcel.Range("S1") = "Semana"
ulinha = objExcel.Range("A65536").End(-4162).Row
objExcel.Columns("R").Replace "", "1"
objExcel.Columns("R").Replace "100", "2"
objExcel.Columns("P").Replace ".", "/"
objExcel.Columns("P").NumberFormat = "dd/mm/yyyy"

objExcel.Range("S2").Formula = "=CONCATENATE(""Semana "",WEEKNUM(RC[-3]))"
objExcel.Range("S2").copy
objExcel.Range("S3:S" & ulinha).PasteSpecial -4123
objExcel.Range("S:S").Copy
objExcel.Range("S:S").PasteSpecial -4163
'objExcel.Range(ulinha -1 & ":" & ulinha -1).Delete
'objExcel.Range(ulinha -1 & ":" & ulinha -1).Delete

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.SaveAs Path + "Contagem Estoque.xlsx", 51
objExcel.ActiveWorkbook.Close
objExcel.Quit

On Error Resume Next
FileSys.DeleteFile (Path + "Contagem Estoque.xls")

End if

'###########################################################################################################

If nVar = 16 Then

nVar = InputBox("Insira a data Limit da Busca", ,"01." & Month(Date) & "." & Year(Date))
nVarf = InputBox("Insira a data Limit da Busca", ,"31." & Month(Date) & "." & Year(Date))

if nVar = Cancel then

else

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "Q:\GROUPS\BR_SC_ITJ_WAU_LOGISTICA_ENGENHEIRADOS\03_GESTÃO_ALMOXARIFADO\"


On Error Resume Next
FileSys.DeleteFile (Path + "MB51.xls")

On Error Resume Next
FileSys.DeleteFile (Path + "MB51.xlsx")

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_LIFNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_KUNNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_MAT_KDAU_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_USNAM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_VGART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_XBLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/btn%_BUDAT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

'session.findById("wnd[0]/usr/radRHIER_L").setFocus
'session.findById("wnd[0]/usr/radRHIER_L").select

session.findById("wnd[0]/usr/radRFLAT_L").setFocus
session.findById("wnd[0]/usr/radRFLAT_L").select

session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/WIJ_BAIXAS"

session.findById("wnd[0]/usr/ctxtPA_AISTR").text = "ZDOC_MATERIAL"
session.findById("wnd[0]/usr/chkDATABASE").selected = true
session.findById("wnd[0]/usr/chkSHORTDOC").selected = false
session.findById("wnd[0]/usr/chkARCHIVE").selected = false

session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").setFocus
session.findById("wnd[0]/usr/ctxtWERKS-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
'session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "281"
'session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "281"
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "261"
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 3
'session.findById("wnd[1]/tbar[0]/btn[8]").press

'session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "RS01"

session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RS01"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RF01"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = nVar
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = nVarf
session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/PAULOA"
session.findById("wnd[0]/usr/ctxtPA_AISTR").text = "ZDOC_MATERIAL"
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB51.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8

On error resume next
session.findById("wnd[1]/tbar[0]/btn[11]").press

On error resume next
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


'########################################################### E X C E L L ##########################################
Set objExcel = CreateObject("Excel.Application")
endereco = Path + "MB51.xls"

Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("MB51").Activate


objExcel.Columns("A").Delete
objExcel.Range("1:3").Delete
objExcel.Range("2:2").Delete

objExcel.Range("P1") = "Semana"
ulinha = objExcel.Range("A999999").End(-4162).Row
objExcel.Columns("E").Replace ".", "/"
objExcel.Columns("E").NumberFormat = "dd/mm/yyyy"

objExcel.Range("P2").Formula = "=CONCATENATE(""Semana "",WEEKNUM(RC[-11]))"
objExcel.Range("P2").copy
objExcel.Range("P3:P" & ulinha).PasteSpecial -4123
objExcel.Range("P:P").Copy
objExcel.Range("P:P").PasteSpecial -4163
objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.SaveAs Path + "MB51.xlsx", 51
objExcel.ActiveWorkbook.Close
objExcel.Quit


On Error Resume Next
FileSys.DeleteFile (Path + "MB51.xls")

end if


End if

'###########################################################################################################

If nVar = 17 Then

sinput = InputBox("Insira Numero da N.F de Referencia", "Consulta Solicitante N.F", "")
sinput2 = InputBox("Insira Data da N.F de Referencia", "Consulta Solicitante N.F", Day(date) & "." & Month(date) & "." & Year(date))

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZJ1B3N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 4

session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").text = sinput
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").setFocus
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").caretPosition = 5
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = sinput2
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[37,2]").setFocus
session.findById("wnd[3]/usr/lbl[37,2]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATNR[2,0]").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATNR[2,0]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsITEM_TAB/tabpITEM/ssubITEM_TABS:SAPLJ1BB2:3100/btnSOURCE_BUTTON").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/txtVBAK-ERNAM").setFocus
resp = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/txtVBAK-ERNAM").text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/txtVBAK-ERNAM").caretPosition = 0
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
MsgBox(resp)
End if

'###########################################################################################################

If nVar = 18 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\DiagramasFalt\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "MB26"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "Lguedes"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

'session.findById("wnd[0]").sendVKey 12

session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select

'############################################## INICIO Seleção layout #################################
'session.findById("wnd[0]/tbar[1]/btn[33]").press
'session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 90
'session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "90"
'session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DiagramasFalt.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"


If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
Else

End If

'########################################################################################################### 19

If nVar = 19 Then

Path = "C:\Users\" + UserName + "\Desktop\180 dias\"
ender = "C:\Users\" + UserName + "\Desktop\180 dias"

' If FileSys.FolderExists(ender) Then
' FileSys.DeleteFolder (ender)
' End If

session.findById("wnd[0]").maximize
' session.findById("wnd[0]/tbar[0]/okcd").text = "ZTMM402"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/usr/ctxtP_VARI").text = "/VBSWIJALMOX"
' session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1306"
' session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").caretPosition = 4
' session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
' session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
' session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
' session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = "RS01"
' session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").setFocus
' session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").caretPosition = 4
' session.findById("wnd[0]/usr/btn%_S_LGORT_%_APP_%-VALU_PUSH").press
' session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RF01"
' session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
' session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/tbar[1]/btn[8]").press

' session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"DAYS_WITHOUT_CONSUME"
' session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DAYS_WITHOUT_CONSUME"
' session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = ""
' session.findById("wnd[0]/tbar[1]/btn[40]").press
' session.findById("wnd[0]/tbar[1]/btn[45]").press
' session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
' session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
' session.findById("wnd[1]/tbar[0]/btn[0]").press
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "exportH.xls"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
' session.findById("wnd[1]/tbar[0]/btn[0]").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press


Set objExcel = CreateObject("Excel.Application")
endereco = Path + "exportH.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("exportH").Activate
' objExcel.Range("1:3").Delete
' objExcel.Range("2:2").Delete
' objExcel.Columns("A").Delete
objExcel.Columns("H:BA").Delete
objExcel.Range("H1") = "Dep 1"
objExcel.Range("I1") = "Dep 2"
objExcel.Range("J1") = "Dep 3"
objExcel.Range("K1") = "Chaves"
objExcel.Range("L1") = "Usado?"
objExcel.Range("M1") = "Quando?"
objExcel.Range("N1") = "Comprador"

objExcel.Range("O1") = "Usado?"
objExcel.Range("P1") = "Quando?"
objExcel.Range("Q1") = "Comprador"
objExcel.Range("R1") = "Usado?"
objExcel.Range("S1") = "Quando?"
objExcel.Range("T1") = "Comprador"
objExcel.Range("U1") = "Usado?"
objExcel.Range("V1") = "Quando?"
objExcel.Range("W1") = "Comprador"
objExcel.Range("X1") = "DocTest"

ulinha = objExcel.Range("A999999").End(-4162).Row

For i = 2 to 6425 'ulinha

	SessionNumber = connection.Children.Count-1
	Wscript.Sleep 1000
	Set session = connection.sessions.Item(Cint(SessionNumber))
	Wscript.Sleep 100

	mat = objExcel.Range("A" & i)

	session.findById("wnd[0]").maximize
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "MC.9"
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/ctxtSL_WERKS-LOW").text = ""
	session.findById("wnd[0]/usr/ctxtSL_MATNR-LOW").text = mat
	session.findById("wnd[0]/usr/ctxtSL_MATNR-LOW").setFocus
	session.findById("wnd[0]/usr/ctxtSL_MATNR-LOW").caretPosition = 8
	session.findById("wnd[0]/tbar[1]/btn[8]").press

	session.findById("wnd[0]/tbar[1]/btn[7]").press
	session.findById("wnd[1]/usr/sub:SAPLMCS2:0201/radLMCS2-MRKKZ[2,0]").select
	session.findById("wnd[1]/usr/sub:SAPLMCS2:0201/radLMCS2-MRKKZ[2,0]").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press

	session.findById("wnd[0]/usr/lbl[20,4]").setFocus
	session.findById("wnd[0]/usr/lbl[20,4]").caretPosition = 12
	session.findById("wnd[0]/tbar[1]/btn[16]").press

	Dep = ""
	Dep2 = ""
	Dep3 = ""

	On Error Resume Next
	if session.findById("wnd[0]/usr/lbl[40,5]").text > 0 then
	  Dep = session.findById("wnd[0]/usr/lbl[1,5]").text
	end if 

	If Dep = "1321RS01" or Dep = "1306RS01" or Dep = "1321RF01" or Dep = "1306RF01" then

	On Error Resume Next
	if session.findById("wnd[0]/usr/lbl[40,8]").text > 0 then
		Dep = session.findById("wnd[0]/usr/lbl[1,8]").text
	else
		Dep = ""
	end if 

	end if

	On Error Resume Next
	if session.findById("wnd[0]/usr/lbl[40,6]").text > 0 then
		Dep2 = session.findById("wnd[0]/usr/lbl[1,6]").text
	end if

	If Dep2 = "1321RS01" or Dep2 = "1306RS01" or Dep2 = "1321RF01" or Dep2 = "1306RF01" then

	On Error Resume Next
	if session.findById("wnd[0]/usr/lbl[40,9]").text > 0 then
		Dep2 = session.findById("wnd[0]/usr/lbl[1,9]").text
	else
	 Dep2 = ""
	end if 

	end if

	On Error Resume Next
	if session.findById("wnd[0]/usr/lbl[40,7]").text > 0 then
	  Dep3 = session.findById("wnd[0]/usr/lbl[1,7]").text
	end if

	If Dep3 = "1321RS01" or Dep3 = "1306RS01" or Dep3 = "1321RF01" or Dep3 = "1306RF01" then

	On Error Resume Next
	if session.findById("wnd[0]/usr/lbl[40,10]").text > 0 then
		Dep3 = session.findById("wnd[0]/usr/lbl[1,10]").text
	else
		Dep3 = ""
	end if 

	end if


	If Dep = "1321RS01" or Dep = "1306RS01" or Dep = "1321RF01" or Dep = "1306RF01" then
		Dep = ""
	end if

	If Dep2 = "1321RS01" or Dep2 = "1306RS01" or Dep2 = "1321RF01" or Dep2 = "1306RF01" then
		Dep2 = ""
	end if

	If Dep3 = "1321RS01" or Dep3 = "1306RS01" or Dep3 = "1321RF01" or Dep3 = "1306RF01" then
		Dep3 = ""
	end if


	session.findById("wnd[0]/tbar[0]/btn[12]").press
	session.findById("wnd[0]/tbar[0]/btn[12]").press


	session.findById("wnd[0]/tbar[0]/okcd").text = "CS15"
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/chkRC29L-DIRKT").selected = true
	session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = mat
	session.findById("wnd[0]/usr/ctxtRC29L-DATUV").text = ""
	session.findById("wnd[0]/usr/chkRC29L-DIRKT").setFocus
	session.findById("wnd[0]/tbar[1]/btn[5]").press


	session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = objExcel.Range("C" & i)

	session.findById("wnd[0]/usr/ctxtRC29L-WERKS").setFocus
	session.findById("wnd[0]/usr/ctxtRC29L-WERKS").caretPosition = 4
	session.findById("wnd[0]/tbar[1]/btn[8]").press

'================================= NOVO =======================================================
DocTest = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
set GRID = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
GRID.setCurrentCell 0,""

DocTest = GRID.getcellvalue (0, "OJTXB")
objExcel.Range("X" & i) = DocTest
'================================= NOVO =======================================================



	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"OJTXB"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "OJTXB"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
	session.findById("wnd[0]/tbar[1]/btn[29]").press
	
	session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "CHAV*"
	session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 6
	session.findById("wnd[1]/tbar[0]/btn[0]").press

	Ch = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount
	

	Set MyGrid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
	Set MyGrid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
	util = MyGrid.GetCellValue(2,"OJTXB") '&", "& MyGrid.GetCellValue(3,"OJTXB") &", "& MyGrid.GetCellValue(4,"OJTXB")

	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press

	opg = 0
	tpo = ""
	comprador = ""
	usse = "Não"
	datauso = ""
	comprador1 = ""
	usse1 = "Não"
	datauso1 = ""
	comprador2 = ""
	usse2 = "Não"
	datauso2 = ""
	comprador3 = ""
	usse3 = "Não"
	datauso3 = ""
	 tt = 0
	 tt2 = 0
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "md04"
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = mat
	session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").text = "1306"
	session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "1306"
	session.findById("wnd[0]").sendVKey 0

	If session.findById("wnd[0]/sbar").messagetype = "E" then
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").text = "1321"
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "1321"
		session.findById("wnd[0]").sendVKey 0
	end if
	
	If session.findById("wnd[0]/sbar").messagetype = "E" then
		opg = 1
	end if
	
' '================================================== Inicio Verificação DEP 01 =====================================================
	if opg = 0 then

		do until tt = 5
			session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt &"]").setFocus
			tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt &"]").text

			if tpo = "ResOrd" or tpo = "NecDep"then
				session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt &"]").setFocus
				datauso = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt &"]").text
				usse = "Sim"

				session.findById("wnd[0]/usr/btnBUTTON_GROKO").press
				session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").setFocus
				comprador = session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").text
				exit do
			end if
			tt = tt + 1
		loop

		session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-BERID").text = "1321"
		session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-WERKS").text = "1321"
		session.findById("wnd[0]").sendVKey 0
		If session.findById("wnd[0]/sbar").messagetype = "I"  then
			opg = 1
			session.findById("wnd[0]").sendVKey 0
		end if
		
		if opg = 0 then
			do until tt2 = 5

				session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt2 &"]").setFocus
				tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt2 &"]").text

				if tpo = "ResOrd" or tpo = "NecDep"then
					session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt2 &"]").setFocus
					datauso = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt2 &"]").text
					usse = "Sim"

					session.findById("wnd[0]/usr/btnBUTTON_GROKO").press
					session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").setFocus
					comprador = session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").text
					exit do
				end if
				tt2 = tt2 + 1
			loop
		end if
	end if
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
'================================================== Fim Verificação DEP 01 =====================================================

	

'================================================== Inicio Verificação DEP 02 =====================================================
	if Dep <> "" then
		session.findById("wnd[0]/tbar[0]/okcd").text = "md04"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = mat
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").text = Mid(Trim(Dep), 1, 4)
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = Mid(Trim(Dep), 1, 4)

		session.findById("wnd[0]").sendVKey 0
		tt1 = 0
		do until tt1 = 5
			session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt1 &"]").setFocus
			tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt1 &"]").text

			if tpo = "ResOrd" or tpo = "NecDep"then
				session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt1 &"]").setFocus
				datauso1 = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt1 &"]").text
				usse1 = "Sim"

				session.findById("wnd[0]/usr/btnBUTTON_GROKO").press
				session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").setFocus
				comprador1 = session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").text
				exit do
			end if
			tt1 = tt1 + 1
		loop
		session.findById("wnd[0]/tbar[0]/btn[15]").press
	end if
'================================================== Fim Verificação DEP 01 =====================================================


'================================================== Inicio Verificação DEP 02 =====================================================

	if Dep2 <> "" then
		session.findById("wnd[0]/tbar[0]/okcd").text = "md04"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = mat
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").text = Mid(Trim(Dep2), 1, 4)
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = Mid(Trim(Dep2), 1, 4)

		session.findById("wnd[0]").sendVKey 0
		tt2 = 0
		do until tt2 = 5

			session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt2 &"]").setFocus
			tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt2 &"]").text

			if tpo = "ResOrd" or tpo = "NecDep"then
				session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt2 &"]").setFocus
				datauso2 = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt2 &"]").text
				usse2 = "Sim"

				session.findById("wnd[0]/usr/btnBUTTON_GROKO").press
				session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").setFocus
				comprador2 = session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").text
				exit do
			end if
			tt2 = tt2 + 1
		loop
		session.findById("wnd[0]/tbar[0]/btn[15]").press
	end if
'================================================== Fim Verificação DEP 02 =====================================================

'================================================== Inicio Verificação DEP 03 =====================================================

	if Dep3 <> "" then
		session.findById("wnd[0]/tbar[0]/okcd").text = "md04"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = mat
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").text = Mid(Trim(Dep3), 1, 4)
		session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = Mid(Trim(Dep3), 1, 4)

		session.findById("wnd[0]").sendVKey 0
		tt3 = 0
		do until tt3 = 5
			session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt3 &"]").setFocus
			tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"& tt3 &"]").text

			if tpo = "ResOrd" or tpo = "NecDep"then
				session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt3 &"]").setFocus
				datauso3 = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& tt3 &"]").text
				usse3 = "Sim"
				session.findById("wnd[0]/usr/btnBUTTON_GROKO").press
				session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").setFocus
				comprador3 = session.findById("wnd[0]/usr/tabsTABTC/tabpTB01/ssubINCLUDE7XX:SAPLM61K:0101/txtT024-EKNAM").text
				exit do
			end if
			tt3 = tt3 + 1
		loop
		session.findById("wnd[0]/tbar[0]/btn[15]").press
	end if

'================================================== Fim Verificação DEP 03 =====================================================
	
	objExcel.Range("H" & i) = Dep 'Mid(Trim(Dep), 1, 4)
	objExcel.Range("I" & i) = Dep2 'Mid(Trim(Dep2), 1, 4)
	objExcel.Range("J" & i) = Dep3 'Mid(Trim(Dep3), 1, 4)


	If Ch > 0 then
		objExcel.Range("K" & i) = "Chave"
	else
		objExcel.Range("K" & i) = "Outros"
	end if
	
	objExcel.Range("L" & i) = usse
	objExcel.Range("M" & i) = datauso
	objExcel.Range("N" & i) = comprador

	objExcel.Range("O" & i) = usse1
	objExcel.Range("P" & i) = datauso1
	objExcel.Range("Q" & i) = comprador1
	
	objExcel.Range("R" & i) = usse2
	objExcel.Range("S" & i) = datauso2
	objExcel.Range("T" & i) = comprador2
	
	objExcel.Range("U" & i) = usse3
	objExcel.Range("V" & i) = datauso3
	objExcel.Range("W" & i) = comprador3
	
Next

Msgbox("Termino")

Else

End If
'########################################################################################################### 20

If nVar = 20 Then

nVara = InputBox("Insira a data Limit da Busca", , Day(Date)+2 &"."& Month(Date) & "." & Year(Date))

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\DiagramasSerralheria\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\DiagramasSerralheria"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MB26"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "LGUEDES"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_BDTER-LOW").text = nVara
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 4
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "4"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/mbar/menu[0]/menu[2]").select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 9
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 7
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "9"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"LGPBE"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "LGPBE"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/tbar[0]/btn[12]").press

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 11,"RSNUM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clearSelection
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "11"
session.findById("wnd[0]/tbar[1]/btn[45]").press

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export2.xls"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

Set obj = CreateObject("Excel.Application")
endereco = Path + "export2.xls"
Set objExcel = obj.Workbooks.Open(endereco)

objExcel.Application.Visible = True
objExcel.Sheets("export2").Activate

objExcel.Sheets("export2").Range("1:4").Delete
objExcel.Sheets("export2").Range("2:2").Delete
objExcel.Sheets("export2").Columns("A").Delete
ulinha = objExcel.Sheets("export2").Range("A999999").End(-4162).Row

objExcel.Sheets("export2").Range("A:J").AutoFilter 6, "=0"
objExcel.Sheets("export2").Range("C:C").Copy
objExcel.Sheets.add()
objExcel.ActiveSheet.Name = "Plan1"
objExcel.Sheets("Plan1").Activate
objExcel.Sheets("Plan1").Range("A1").PasteSpecial -4163
objExcel.Sheets("Plan1").Range("A:A").RemoveDuplicates 1,2
ulinhaa = objExcel.Sheets("Plan1").Range("A999999").End(-4162).Row
objExcel.Sheets("Plan1").Range("A2:A" & ulinhaa).Copy

session.findById("wnd[0]/tbar[0]/okcd").text = "MB52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1306"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").setFocus
session.findById("wnd[0]/usr/ctxtWERKS-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/chkPA_SOND").selected = false
session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "RS01"
session.findById("wnd[0]/usr/ctxtP_VARI").text = "/VBSWIJALMOX"
session.findById("wnd[0]/usr/ctxtP_VARI").setFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 12
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "saldo.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

endereco2 = Path + "saldo.xls"
Set objExcelb = obj.Workbooks.Open(endereco2)
objExcelb.Application.Visible = True
objExcelb.Sheets("saldo").Activate

objExcelb.Sheets("saldo").Range("1:3").Delete
objExcelb.Sheets("saldo").Range("2:2").Delete
objExcelb.Sheets("saldo").Columns("A").Delete
ulinhab = objExcelb.Sheets("saldo").Range("A999999").End(-4162).Row

objExcelb.Sheets("saldo").Range("E2").FormulaR1C1 = "=CONCATENATE(RC[-4],RC[-3])"
objExcelb.Sheets("saldo").Range("E2").Copy
objExcelb.Sheets("saldo").Range("E2:E" & ulinhab).PasteSpecial -4123 

objExcel.Sheets("export2").Activate
objExcel.Sheets("export2").Range("F2").FormulaR1C1 = "=IFERROR(VLOOKUP(CONCATENATE(RC[-3],RC[-4]),saldo.xls!C5:C6,2,0),0)"
objExcel.Sheets("export2").Range("F2").Copy
objExcel.Sheets("export2").Range("F2").Select
objExcel.Sheets("export2").Range(objExcel.Sheets("export2").Range("F2"), objExcel.Sheets("export2").Range("F2").End(-4121)).Select
objExcel.ActiveSheet.Paste 'Special -4123
objExcel.Sheets("export2").Range("A:J").AutoFilter 6, "=0"
objExcel.Sheets("export2").Range("A:A").Copy

objExcel.Sheets("Plan1").Activate
objExcel.Sheets("Plan1").Range("B1").PasteSpecial -4163
objExcel.Sheets("Plan1").Range("B:B").RemoveDuplicates 1,2
ulinhabb = objExcel.Sheets("Plan1").Range("B999999").End(-4162).Row

for u = 2 to ulinhabb

filtro = objExcel.Sheets("Plan1").Range("B" & u) 

objExcel.Sheets("export2").Activate
objExcel.ActiveSheet.ShowAllData
objExcel.Sheets("export2").Range("A:J").AutoFilter 1, "=" & filtro

objExcel.Sheets("export2").Range("F2").Select
objExcel.Sheets("export2").Range(objExcel.Sheets("export2").Range("F2"), objExcel.Sheets("export2").Range("F2").End(-4121)).EntireRow.Delete
next

objExcel.ActiveSheet.ShowAllData
objExcel.Sheets("export2").Range("A:A").Copy

objExcel.Sheets("Plan1").Activate
objExcel.Sheets("Plan1").Range("C1").PasteSpecial -4163
objExcel.Sheets("Plan1").Range("C:C").RemoveDuplicates 1,2
ulinhac = objExcel.Sheets("Plan1").Range("C999999").End(-4162).Row
objExcel.Sheets("Plan1").Range("C2:C" & ulinhac).Copy


session.findById("wnd[0]/tbar[0]/okcd").text = "MB26"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "LGUEDES"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_BDTER-LOW").text = nVar
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 4
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "4"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[2]").select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

'session.findById("wnd[0]/tbar[1]/btn[32]").press
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 11
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "11"
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 9
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "9"
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 11
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "11"
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 5
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "5"
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 8
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "8"
'session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
'session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[0]/btn[86]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press



for ui = 2 to ulinhac

diag = objExcel.Sheets("Plan1").Range("C" & ui)

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"AUFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "AUFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
session.findById("wnd[0]/tbar[1]/btn[29]").press

session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = diag
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press

if ui <> ulinhac then
session.findById("wnd[0]/tbar[0]/btn[86]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
else
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
end if

next


objExcelb.Sheets("saldo").Range("A1").Copy
objExcelb.Saved = True
objExcelb.Close


objExcel.Sheets("Plan1").Range("A1").Copy
objExcel.Saved = True
objExcel.Close


ender = "C:\Users\" + UserName + "\Desktop\DiagramasSerralheria"


If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

Set objWorkbook = Nothing
Set objExcel = Nothing
Set objExcelb = Nothing

Else

End If
'########################################################################################################### 21

If nVar = 21 Then

Dim tipo, Deposito, Vendedora, Responsavel, local, Conta, CC1, CC2, codigo, Qtd, UM, centro

tipo = "K"
Deposito = "RS01"
Vendedora = "167"
Responsavel = "PAULOANTONIO"

local = "CRM - AUTOMAÇÂO"
Conta1 = "411075061" 'Consumiveis
Conta2 = "411060145" 'Uniformes
CC1 = "31067056"
CC2 = "31067686"

conta00 = InputBox("Favor escolher a conta: " _
& vbNewLine & "  " _
& vbNewLine & " ( 1 ) - Consumivel " _
& vbNewLine & " ( 2 ) - Uniformes " _
, "Código Conta", "")

If conta00 = 1 then
Conta = Conta1
elseIf conta00 = 2 then
Conta = Conta2
elseIf conta00 = "" then
 WScript.Quit
end if

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ME51n"
session.findById("wnd[0]").sendVKey 0

codigo1 = InputBox("Favor colocar Código do Material", "Código Material", "")
If codigo1 = "" then
 WScript.Quit
end if

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"KNTTP",tipo
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"MATNR",codigo1
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"MENGE",InputBox("Favor colocar a quantidade a ser comprada", "Qtd Material", "1")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"MEINS",InputBox("Favor colocar a Unidade de Medida do Material", "U.M Material", "UN")

centro1 = InputBox("Favor colocar o Centro do Material", "Centro Material", "1306")

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"NAME1",centro1
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"LGOBE",Deposito
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"EKGRP",Vendedora
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"AFNAM",Responsavel
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/txtMEACCT1100-ABLAD").text = local
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/txtMEACCT1100-WEMPF").text = Responsavel
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = Conta

If centro1 = "1306" then
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = CC1
elseif centro1 = "1321" then
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = CC2
end if

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").caretPosition = 8
session.findById("wnd[0]").sendVKey 0


For i = 1 to 100

codigo = InputBox("Favor colocar Código do Material", "Código Material", "")
If codigo = "" then exit for end if

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"KNTTP",tipo
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"MATNR",codigo
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"MENGE",InputBox("Favor colocar a quantidade a ser comprada", "Qtd Material", "1")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"MEINS",InputBox("Favor colocar a Unidade de Medida do Material", "U.M Material", "UN")

centro = InputBox("Favor colocar o Centro do Material", "Centro Material", "1306")

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"NAME1",centro
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"LGOBE",Deposito
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"EKGRP",Vendedora
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell i,"AFNAM",Responsavel
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/txtMEACCT1100-ABLAD").text = local
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/txtMEACCT1100-WEMPF").text = Responsavel
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = Conta

If centro = "1306" then
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = CC1
elseif centro = "1321" then
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = CC2
end if

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
next

'session.findById("wnd[0]/tbar[0]/btn[11]").press
'session.findById("wnd[0]/tbar[0]/btn[15]").press

Else

End If

'###########################################################################################################
If nVar = 22 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\OrdensFalt\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\OrdensFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

ender2 = "C:\Users\" + UserName + "\Desktop\Remessas"

If FileSys.FolderExists(ender2) Then
    FileSys.DeleteFolder (ender2)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "coois"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOM000"
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "herold"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Falt_Ordens.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]").resizeWorkingPane 272, 40, False
session.findById("wnd[0]/tbar[0]/btn[12]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "Falt_Ordens.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("Falt_Ordens").Activate

With objExcel
.Columns("A:A").Delete
.Range("1:3").Delete
.Range("2:2").Delete
.Columns("I").Replace ".", "/"
.Range("B2:B999").Copy
End With

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "mb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1306"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "RS01"
session.findById("wnd[0]/usr/chkNOZERO").Selected = False
session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/VBSWIJALMOX"
session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Est_Falt_Ordens.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit


Datainicial = Date - 15
DataFinal = Date

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Remessas\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Remessas"

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZTSD092"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "HELIOCD"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 1, "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = Day(Datainicial) & "." & Month(Datainicial) & "." & Year(Datainicial)
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").Text = Day(DataFinal) & "." & Month(DataFinal) & "." & Year(DataFinal)


session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"

Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate
 
objExcel.Columns("A").Delete
objExcel.Columns("C").Delete
objExcel.Range("1:7").Delete
objExcel.Range("2:2").Delete
objExcel.Columns("C").Replace ".", "/"

If Weekday(date) = 2 then
datea = date - 4
else
datea  = date - 2
end if

dateb =  "" & Month(datea) & "/" & Day(datea) & "/" & Year(datea) & ""

objExcel.Range("A:E").AutoFilter 3, ">=" & dateb & ""
'msgbox("Se ja ajusto o filtro de Data click 'OK'")
objExcel.Range("A2:A999999").Copy


session.findById("wnd[0]/tbar[0]/okcd").Text = "VL06O"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btnBUTTON6").press
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press

session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/WIJ REMESSA"
session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 12
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[14,2]").setFocus
session.findById("wnd[3]/usr/lbl[14,2]").caretPosition = 9
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").SetFocus
'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").caretPosition = 15
'session.findById("wnd[1]").sendVKey 2

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "remessas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionFaltOrdens2.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Deposito"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

ender = "C:\Users\" + UserName + "\Desktop\OrdensFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

ender = "C:\Users\" + UserName + "\Desktop\Remessas"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

End If

'###########################################################################################################

If nVar = 23 Then

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\DiagramasFalt\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"

If FileSys.FolderExists(ender) Then
    FileSys.DeleteFolder (ender)
End If

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "cm03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txt[35,3]").Text = "04030349"
session.findById("wnd[0]/usr/txt[35,5]").Text = ""
session.findById("wnd[0]/usr/txt[35,3]").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/chk[1,7]").Selected = True
'session.findById("wnd[0]/usr/chk[1,8]").selected = true
session.findById("wnd[0]/usr/chk[1,33]").Selected = True
'session.findById("wnd[0]/usr/chk[1,34]").selected = true
session.findById("wnd[0]/usr/chk[1,34]").SetFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[49,1]").SetFocus
session.findById("wnd[0]/usr/lbl[49,1]").caretPosition = 38
session.findById("wnd[0]").sendVKey 20

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate

With objExcel
.Range("1:5").Delete
.Range("2:3").Delete
.Columns("A").Delete
.Columns("B").Delete
.Columns("C:F").Delete
.Range("A:k").AutoFilter 9, "=LIB"
.Range("B2:B999").Copy
End With

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

session.findById("wnd[0]/tbar[0]/okcd").Text = "MB26"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "Lguedes"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[1]/tbar[0]/btn[24]").press

session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

'session.findById("wnd[0]").sendVKey 12

objExcel.Range("N1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

FileSys.DeleteFile (endereco)

session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select

'############################################## INICIO Seleção layout #################################
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
'############################################## FIM Seleção layout #################################

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DiagramasFalt.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press


Set obj = CreateObject("Excel.Application")
endereco = Path + "DiagramasFalt.xls"
Set objExcel = obj.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("DiagramasFalt").Activate

objExcel.Sheets("DiagramasFalt").Columns("A:A").Delete
objExcel.Sheets("DiagramasFalt").Range("1:4").Delete
objExcel.Sheets("DiagramasFalt").Range("2:2").Delete
objExcel.Sheets("DiagramasFalt").Columns("I").Replace ".", "/"
objExcel.Sheets("DiagramasFalt").Range("C2:C999").Copy
ulinha = objExcel.Sheets("DiagramasFalt").Range("A999999").End(-4162).Row

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "mb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1306"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1321"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "RS01"
session.findById("wnd[0]/usr/chkNOZERO").Selected = False
session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/VBSWIJALMOX"
session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Est_Falt_Diagramas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press




endereco2 = Path + "Est_Falt_Diagramas.xls"
Set objExcelb = obj.Workbooks.Open(endereco2)
objExcelb.Application.Visible = True
objExcelb.Sheets("Est_Falt_Diagramas").Activate

objExcelb.Sheets("Est_Falt_Diagramas").Range("1:3").Delete
objExcelb.Sheets("Est_Falt_Diagramas").Range("2:2").Delete
objExcelb.Sheets("Est_Falt_Diagramas").Columns("A").Delete


objExcel.Sheets("DiagramasFalt").Range("N2") = "est."
objExcel.Sheets("DiagramasFalt").Range("N2").Formula = "=IF(RC[-7]=0,IF(SUMIFS(Est_Falt_Diagramas.xls!C6,Est_Falt_Diagramas.xls!C1,RC[-11],Est_Falt_Diagramas.xls!C2,RC[-12])=0,""0"",SUMIFS(Est_Falt_Diagramas.xls!C6,Est_Falt_Diagramas.xls!C1,RC[-11],Est_Falt_Diagramas.xls!C2,RC[-12])),RC[-7])"
objExcel.Sheets("DiagramasFalt").Range("N2").copy
objExcel.Sheets("DiagramasFalt").Range("N2:N" & ulinha).PasteSpecial -4123
objExcel.Sheets("DiagramasFalt").Range("N:N").Copy
objExcel.Sheets("DiagramasFalt").Range("N:N").PasteSpecial -4163
objExcel.Sheets("DiagramasFalt").Range("N2:N" & ulinha).Copy
objExcel.Sheets("DiagramasFalt").Range("G2:G" & ulinha).PasteSpecial -4163

objExcel.Save
objExcel.Close
objExcel.Quit

objExcelb.Sheets("Est_Falt_Diagramas").Range("N1").Copy
objExcelb.Saved = True
objExcelb.Close
objExcelb.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing
Set objExcelb = Nothing
Set obj = Nothing

Datainicial = Date - 9
DataFinal = Date

Data = "\" & Day(Date) & "." & Month(Date) & "." & Year(Date) & "\"
Path = "C:\Users\" + UserName + "\Desktop\Remessas\" + Data + "\"
ender = "C:\Users\" + UserName + "\Desktop\Remessas"

'session.createSession
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZTSD092"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "HELIOCD"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 1, "TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = Day(Datainicial) & "." & Month(Datainicial) & "." & Year(Datainicial)
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").Text = Day(DataFinal) & "." & Month(DataFinal) & "." & Year(DataFinal)


session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

Set objExcel = CreateObject("Excel.Application")
endereco = Path + "export.xls"
Set objWorkbook = objExcel.Workbooks.Open(endereco)
objExcel.Application.Visible = True
objExcel.Sheets("export").Activate
 
objExcel.Columns("A").Delete
objExcel.Columns("C").Delete
objExcel.Range("1:7").Delete
objExcel.Range("2:2").Delete
objExcel.Columns("C").Replace ".", "/"

If Weekday(date) = 2 then
datea = date - 4
else
datea  = date - 2
end if

dateb =  "" & Month(datea) & "/" & Day(datea) & "/" & Year(datea) & ""

objExcel.Range("A:E").AutoFilter 3, ">=" & dateb  & ""
'msgbox("Se ja ajusto o filtro de Data click 'OK'")
objExcel.Range("A2:A999999").Copy


session.findById("wnd[0]/tbar[0]/okcd").Text = "VL06O"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btnBUTTON6").press
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press

session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/WIJ REMESSA"
session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 12
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[3]/usr/lbl[14,2]").setFocus
session.findById("wnd[3]/usr/lbl[14,2]").caretPosition = 9
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").SetFocus
'session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/11[0,11]/lbl[14,11]").caretPosition = 15
'session.findById("wnd[1]").sendVKey 2

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "remessas.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objExcel = CreateObject("Excel.Application")
ender = "Q:\PUBLIC\BR_SC_ITJ_WDC_CRM_ALMOXARIFADO\ConnectionFaltDiagrama2.xlsm"
Set objWorkbook = objExcel.Workbooks.Open(ender)
objExcel.Application.Visible = True

objExcel.Run "Ordens"

objExcel.Range("A1").Copy
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

ender = "C:\Users\" + UserName + "\Desktop\DiagramasFalt"

If FileSys.FolderExists(ender) Then
    'FileSys.DeleteFolder (ender)
End If

ender = "C:\Users\" + UserName + "\Desktop\Remessas"

If FileSys.FolderExists(ender) Then
    'FileSys.DeleteFolder (ender)
End If


Else

End If

do until session.Info.Transaction = "SESSION_MANAGER"
	session.findById("wnd[0]/tbar[0]/btn[3]").press
loop
session.findById("wnd[0]/tbar[0]/btn[15]").press
Set session = Nothing	
'###########################################################################################################

