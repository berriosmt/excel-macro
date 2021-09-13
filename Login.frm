VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Iniciar Secci�n"
   ClientHeight    =   8424
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8760
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnIniciar_Click()
'Verificar que los campos no est�n vacios
If Trim(txtUsuario.Text) = "" Or Trim(txtPassword) = "" Then
MsgBox "Ingrese el usuario y/o contrase�a.", vbInformacion, "Informaci�n incompleta" 't�tulo de la ventana
txtUsuario.SetFocus 'el enfoque estar� en este textbox
Else
'Usuario sici y contrase�a aplicada
If Me.txtUsuario.Value = "sici" And Me.txtPassword.Value = "aplicada" Then
'Registrar el login en la hoja Logfile
Dim wb As Workbook 'Variable que hace referencia al documento completo
Dim ws As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb = ActiveWorkbook
Set ws = wb.Sheets("LogFile")
ws.Select
'Colocar la informaci�n en las filas
cRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = txtUsuario.Text
Cells(cRow, 2) = Date
Cells(cRow, 3) = Time
Cells(cRow, 4) = "Inici� Secci�n"

'Abir la pantalla del men�
Menu.Show
'Usuario profe y contrase�a programaci�n
ElseIf Me.txtUsuario.Value = "profe" And Me.txtPassword.Value = "programacion" Then

'Registrar el login en la hoja Logfile
Dim wb2 As Workbook 'Variable que hace referencia al documento completo
Dim ws2 As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb2 = ActiveWorkbook
Set ws2 = wb2.Sheets("LogFile")
ws2.Select
'Colocar la informaci�n en las filas
cRow = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = txtUsuario.Text
Cells(cRow, 2) = Date
Cells(cRow, 3) = Time
Cells(cRow, 4) = "Inici� Secci�n"

'Abrir la pantalla del men�
Menu.Show
Else
'Si el usuario es incorrecto, se muestra un mensaje de error
If Me.txtUsuario <> "sici" And Me.txtUsuario <> "profe" Then
MsgBox "Usuario incorrecto", vbInformation, "Informaci�n incorrecta"

Else
'Si la contrase�a es incorrecta, se muestra un mensaje de error
If Me.txtPassword <> "aplicada" And Me.txtPassword <> "programacion" Then
MsgBox " Contrase�a incorrecta", vbInformation, "Informaci�n incorrecta"
End If
End If
End If
End If
End Sub
'Bot�n Salir
Private Sub btnSalir_Click()
'Preguntar al usuario si desea salir
If MsgBox("�Desea salir del sistema?", vbQuestion + vbYesNo) = vbYes Then
'Salir del sistema
Application.Quit
End If
End Sub
