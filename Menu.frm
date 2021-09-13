VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Menú"
   ClientHeight    =   8268
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7284
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEventos_Click()
'Abrir la pantalla de Registro Eventos
RegistroEventos.Show
End Sub

Private Sub btnSalir_Click()
'Preguntar al usuario si desea salir
If MsgBox("¿Desea salir del sistema?", vbQuestion + vbYesNo) = vbYes Then
'Registrar nuevo evento en la hoja LogFile
Dim wb2 As Workbook 'Variable que hace referencia al documento completo
Dim ws2 As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb2 = ActiveWorkbook
Set ws2 = wb2.Sheets("LogFile")
ws2.Select

'Colocar la información en las celdas
cRow = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = Login.txtUsuario 'Usuario ingresado en la forma Login
Cells(cRow, 2) = Date 'Fecha
Cells(cRow, 3) = Time 'Hora
Cells(cRow, 4) = "Cerró Sección"
'Salir del sistema
Application.Quit
End If
End Sub

Private Sub btnVideos_Click()
'Abrir la pantalla de Registro videos
RegistroVideos.Show
End Sub
