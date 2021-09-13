VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistroProgramas 
   Caption         =   "Registrar Programas"
   ClientHeight    =   6300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9300
   OleObjectBlob   =   "RegistroProgramas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegistroProgramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLimpiar_Click()
'Limpiar los campos de los textbox
txtPrograma = Empty
txtNombre = Empty
cboDept.Value = Empty
End Sub

Private Sub btnRegistrar_Click()
'Verificar que los campos no estén vacios
If Trim(txtPrograma.Text) = "" Or Trim(txtNombre) = "" Then
MsgBox "Ingrese la información del video.", vbInformacion, "Información incompleta" 'título de la ventana
txtPrograma.SetFocus 'el enfoque va a estar ahí

Else
'Verificar que el usuario solo escribe 4 letras para el textbox Programa
If Len(txtPrograma.Text) > 4 Then
MsgBox "Debe escribir solo 4 letras.", vbInformacion, "Información incorrecta"
Else
'Registrar el programa en la hoja Programas
Dim wb As Workbook 'Variable que hace referencia al documento completo
Dim ws As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Programas")
ws.Select

'Colocar la información de los textbox en las celdas
cRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = txtPrograma.Text
Cells(cRow, 2) = txtNombre.Text
Cells(cRow, 3) = cboDept.Value

'Registrar nuevo programa en la hoja LogFile
Dim wb2 As Workbook 'Variable que hace referencia al documento completo
Dim ws2 As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb2 = ActiveWorkbook
Set ws2 = wb2.Sheets("LogFile")
ws2.Select

'Colocar la información en las celdas
cRow = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = Login.txtUsuario 'Usuario ingresado en la forma login
Cells(cRow, 2) = Date 'Fecha
Cells(cRow, 3) = Time 'Hora
Cells(cRow, 4) = "Nuevo Programa"
End If
End If
End Sub

Private Sub buscarDepartamentos()
'Buscar los pueblos ingresados en la hoja Pueblos
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Dept")
'Añadir los valores de la hoja Pueblo al combobox
cRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
For d = 2 To cRow
 With Me.cboDept
  Set curCell = ws.Cells(d, 1)
   .AddItem curCell.Value
   End With
Next d
End Sub

Private Sub btnRegresar_Click()
'Cerrar la pantalla para regresar al menú
Unload Me
End Sub

Private Sub UserForm_Initialize()
'Llamando la función para llenar el combobox de departamentos
buscarDepartamentos
End Sub
