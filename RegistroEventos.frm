VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistroEventos 
   Caption         =   "Registrar Eventos"
   ClientHeight    =   10032
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7920
   OleObjectBlob   =   "RegistroEventos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegistroEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLimpiar_Click()
'Limpiar los campos de los textbox
txtID = Empty
txtFecha = Empty
txtNombre = Empty
txtCurso = Empty
txtURL = Empty
txtTema = Empty
txtDescripcion = Empty
txtMinutos = Empty
txtSegundos = Empty
txtDuracion = Empty
cbCanal.Value = False
lblURL.Visible = False
txtURL.Visible = False
End Sub

Private Sub btnRegistrar_Click()
'Verificar que los campos no estén vacios
If Trim(txtID.Text) = "" Or Trim(txtFecha.Text) = "" Or Trim(txtNombre.Text) = "" Or Trim(txtCurso.Text) = "" Or Trim(txtTema.Text) = "" Or Trim(txtMinutos.Text) = "" Or Trim(txtSegundos.Text) = "" Then
MsgBox "Ingrese la información del video.", vbInformacion, "Información incompleta" 'título de la ventana
txtID.SetFocus 'el enfoque va a estar ahí

Else

Dim wb As Workbook 'Variable que hace referencia al documento completo
Dim ws As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Eventos")
ws.Select

'Colocar la información de los textbox en las celdas
cRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = txtID.Text
Cells(cRow, 2) = txtFecha.Text
Cells(cRow, 3) = txtNombre.Text
Cells(cRow, 4) = txtCurso.Text
'Si se marca el checkbox, se coloca "Sí" en la celda de Canal
If cbCanal.Value = True Then
Cells(cRow, 5) = "Sí"
Else
'Si no se marca el checkbox, se coloca "No" en la celda de Canal
Cells(cRow, 5) = "No"
End If
Cells(cRow, 6) = txtURL.Text
Cells(cRow, 7) = txtTema.Text
Cells(cRow, 8) = txtDescripcion.Text
Cells(cRow, 9) = txtMinutos.Text
Cells(cRow, 10) = txtSegundos.Text
Cells(cRow, 11) = txtDuracion.Text

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
Cells(cRow, 4) = "Nuevo Evento"

End If
End Sub
Private Sub cbCanal_Click()
'Si se marca el checkbox, se muestra el label y textbox de URL
If cbCanal.Value = True Then
lblURL.Visible = True
txtURL.Visible = True
Else
'Si no se marca el checkbox, no se muestra el label ni el textbox de URL
lblURL.Visible = False
txtURL.Visible = False
End If
End Sub
Private Sub btnRegresar_Click()
'Cerrar la pantalla para regresar al menú
Unload Me
End Sub

'Calcular duración final
Private Sub calcularDuracion()
'Verificar que los números son numéricos
If IsNumeric(txtMinutos.Value) And IsNumeric(txtSegundos.Value) Then
'Varibles enteras
Dim X, Y, duracion, minutos As Integer
'Asignar a la variables el valor de los textbox
X = Val(txtMinutos.Value)
Y = Val(txtSegundos.Value)
'Convertir los minutos en segundos
minutos = X * 60
'Calcular duración sumando los segundos y minutos
duracion = minutos + Y
'Mostrar la duración en el textbox
Me.txtDuracion.Value = duracion
Else
'Si los valores no son numéricos, el textbox se queda en blanco
Me.txtDuracion.Value = Empty
End If
End Sub
Private Sub txtMinutos_AfterUpdate()
'Si los valores no son numéricos, se muestra un mensaje de error.
If Not IsNumeric(txtMinutos.Value) And (txtMinutos.Value <> vbNullString) Then
MsgBox "El valor debe ser numerico", vbInformation, "Error"
txtMinutos.Value = vbNullString 'Vaciar el textbox
Else
'Verificar que el número no sea menor que cero.
If txtMinutos.Value < 0 Then
'Si es menor que cero, se muestra un mensaje de error
MsgBox "El valor debe ser mayor que cero", vbInformation, "Error"
txtMinutos.Value = vbNullString 'Vaciar el textbox
Else
'Llamando la función para calcular la duración
calcularDuracion
End If
End If
End Sub
Private Sub txtSegundos_AfterUpdate()
'Si los valores no son numéricos, se muestra un mensaje de error.
If Not IsNumeric(txtSegundos.Value) And (txtSegundos.Value <> vbNullString) Then
MsgBox "El valor debe ser numerico", vbInformation, "Error"
txtSegundos.Value = vbNullString 'Vaciar el textbox
Else
'Verificar que el número no sea menor que cero.
If txtSegundos.Value < 0 Then
'Si es menor que cero, se muestra un mensaje de error
MsgBox "El valor debe ser mayor que cero", vbInformation, "Error"
txtMinutos.Value = vbNullString 'Vaciar el textbox
Else
'Llamando la función para calcular la duración
calcularDuracion
End If
End If
End Sub
