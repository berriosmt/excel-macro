VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistroVideos 
   Caption         =   "Registrar Videos"
   ClientHeight    =   10152
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9432
   OleObjectBlob   =   "RegistroVideos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegistroVideos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
'Abrir la pantalla de Registrar Programas
RegistroProgramas.Show
End Sub

Private Sub btnLimpiar_Click()
'Limpiar los campos de los textbox
txtID = Empty
txtAutor = Empty
txtCola = Empty
txtFecha = Empty
cboPrograma.Value = Empty
txtDescripcion = Empty
txtEmail = Empty
txtMinutos = Empty
txtSegundos = Empty
txtDuracion = Empty
End Sub

Private Sub btnRegistrar_Click()
'Verificar que los campos no estén vacios
If Trim(txtID.Text) = "" Or Trim(txtAutor) = "" Or Trim(txtDescripcion.Text) = "" Or Trim(txtEmail.Text) = "" Or Trim(txtFecha.Text) = "" Or Trim(txtMinutos.Text) = "" Or Trim(txtSegundos.Text) = "" Then
MsgBox "Ingrese la información del video.", vbInformacion, "Información incompleta" 'título de la ventana
txtID.SetFocus 'el enfoque va a estar ahí

Else
'Registrar el video en la hoja Videos
Dim wb As Workbook 'Variable que hace referencia al documento completo
Dim ws As Worksheet 'Variable que hace referencia a la hoja de calculo
'Hacer referencia al documento abierto
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Videos")
ws.Select

'Colocar la información de los textbox en las celdas
cRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
Cells(cRow, 1) = txtID.Text
Cells(cRow, 2) = txtAutor.Text
Cells(cRow, 3) = txtCola.Text
Cells(cRow, 4) = txtFecha.Text
Cells(cRow, 5) = cboPrograma.Value
Cells(cRow, 6) = txtDescripcion.Text
Cells(cRow, 7) = txtEmail.Text
Cells(cRow, 8) = txtMinutos.Text
Cells(cRow, 9) = txtSegundos.Text
Cells(cRow, 10) = txtDuracion.Text

'Registrar nuevo video en la hoja LogFile
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
Cells(cRow, 4) = "Nuevo Video"
End If
End Sub
Private Sub btnRegresar_Click()
'Cerrar la pantalla para regresar al menú
Unload Me
End Sub

Private Sub UserForm_Initialize()
'Llamando la función para llenar el combobox de Programas
buscarProgramas
End Sub

'Función para buscar pueblos
Private Sub buscarProgramas()
'Buscar los pueblos ingresados en la hoja Pueblos
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Programas")
'Añadir los valores de la hoja Pueblo al combobox
cRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
For d = 2 To cRow
 With Me.cboPrograma
  Set curCell = ws.Cells(d, 1)
   .AddItem curCell.Value
   End With
Next d
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

