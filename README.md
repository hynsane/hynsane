Public Plan As Worksheet
Function UsuarioRede() As String
Dim GetUserN
Dim ObjNetwork
Set ObjNetwork = CreateObject("WScript.Network")
GetUserN = ObjNetwork.UserName
UsuarioRede = GetUserN
End Function
Sub Exp_Bases_HSE()

Range("I6:I7").Select
If IsEmpty(ActiveCell) Then
MsgBox "Você não preencheu todos os campos de LOGIN e SENHA", vbCritical
Range("E10:E11").Select
Else

'Desabilita atualização de tela

Application.ScreenUpdating = False

'Desabilita alertas na tela

Application.DisplayAlerts = False

'Apaga os arquivos das planilhas antigas nas pastas

'Kill ("N:\DADOS\PSM\Interno\TMM\CENTRAL DE TELEMETRIA\TORRE LOGISTICA\Indicador apontamento grua\Base")

'Apaga as informações antigas da abas das planilhas

'Call Apagar

'Abre o SAP e faz logon

Dim sKillSapGui1 As String

sKillSapGui1 = "TASKKILL /F /IM Saplogon.exe"
Shell sKillSapGui1, vbHide
Application.Wait Now + TimeValue("0:00:04")


Set Plan = Sheets("home")
NomeUsuario = Application.UserName
inicio = Plan.Range("D4")
fim = Plan.Range("H4")
Mes = Plan.Range("B4")
login = Plan.Range("I6")
senha = Plan.Range("I7")
Application.DisplayAlerts = False

Dim SapGui
Dim Applic
Dim connection
Dim session
Dim WSHShell
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
Set WSHShell = CreateObject("WScript.Shell")
Do Until WSHShell.AppActivate("SAP Logon ")
Application.Wait Now + TimeValue("0:00:01")

Loop

Set WSHShell = Nothing
Set SapGui = GetObject("SAPGUI")
Set Applic = SapGui.GetScriptingEngine
Set connection = Applic.OpenConnection("SBP - ERP ECC - Produção", True)
Set session = connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "500"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = login
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha
session.findById("wnd[0]").sendVKey 0

End If


'baixa bases ZFLRCM150

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
End If

If Not IsObject(connection) Then
   Set connection = Application.Children(0)
End If

If Not IsObject(session) Then
   Set session = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZFLRCM100"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/lbl[1,4]").SetFocus
session.findById("wnd[2]/usr/lbl[1,4]").caretPosition = 2
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[1]/usr/btnCMD_OK").press
session.findById("wnd[0]/tbar[1]/btn[9]").press
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/tbar[0]/btn[12]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/tbar[0]/btn[12]").press
session.findById("wnd[2]/tbar[0]/btn[12]").press


'Copiar os dados para a abas das Planilha

Dim caminho100 As Variant

Application.DisplayAlerts = False

Dim Este As Workbook, Outro As Workbook

caminho100 = "N:\DADOS\PSM\Interno\TMM\CENTRAL DE TELEMETRIA\TORRE LOGISTICA\Pasta Equipe\PEDRO\BASE\CAMINHO PARA BASE\Base_zflrcm100.xlsx"

'base de viagens

Workbooks.Open caminho100, , True
Set Este = ThisWorkbook
Set Outro = ActiveWorkbook

Outro.Sheets(1).Range("A1:CS15000").CurrentRegion.Copy
Planalha1.Range("A1").PasteSpecial
Outro.Close False


'Abilita atualização de tela

Application.ScreenUpdating = True


Workbooks("BASE DE MACRO.XLSM1").Activate
Planilha3.Activate

MsgBox "Dados importados com Sucesso!"


End Sub

Sub Apagar()

' Apagar informações antigas

    Sheets("ZFLRCM10").Select
    Range("A2:CS15000").Select
    Selection.ClearContents
    Range("A2").Select
    
End Sub
