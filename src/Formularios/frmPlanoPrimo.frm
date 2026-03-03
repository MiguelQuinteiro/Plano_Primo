VERSION 5.00
Begin VB.Form frmPlanoPrimo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   13.309
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   13.891
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuEtiquetas 
      Caption         =   "Etiquetas"
   End
   Begin VB.Menu mnuColumnas 
      Caption         =   "Columnas"
   End
End
Attribute VB_Name = "frmPlanoPrimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miEtiqueta As Boolean
Dim miColumnas As Long
Dim miNueva As Long
Dim miDivisor As Long

' AL CARGAR EL FORMULARIO
Private Sub Form_Load()
  miColumnas = 1
  miEtiqueta = True
  miNueva = 1
  miDivisor = 4
End Sub
' CANTIDAD DE COLUMNAS
Private Sub mnuColumnas_Click()
  miColumnas = InputBox("Cantidad de Columnas", "Columnas")
  Call Form_DblClick
End Sub

' ACTIVAR LAS ETIQUETAS
Private Sub mnuEtiquetas_Click()
  If miEtiqueta = False Then
    miEtiqueta = True
  Else
    miEtiqueta = False
  End If
  Call Form_DblClick
End Sub

' AL HACER DOBLE CLICK
Private Sub Form_DblClick()
  Dim x As Long
  Dim y As Long
  Dim p As Long
  Dim r As Integer

  y = 1
  x = 0
  Cls
  'For p = 1 To 3402
  For p = 1 To 100000

    ' Pinta primo y no primo de distinto color
    If Primo(p) Then
      If Primo(p + 2) Then
        PSet (miNueva + (x / miDivisor), y / miDivisor), QBColor(12)  ' Dibuja confetti
        If miEtiqueta = True Then
          Print p
        End If
        Circle (miNueva + (x / miDivisor), y / miDivisor), 1 / 10, QBColor(12)  ' Dibuja confetti
      Else
        'For r = 0 To 100
        PSet (miNueva + (x / miDivisor), y / miDivisor), QBColor(12)  ' Dibuja confetti
        Circle (miNueva + (x / miDivisor), y / miDivisor), 1 / 10, QBColor(12)  ' Dibuja confetti
        'Next
      End If
    Else
      'For r = 0 To 100
      PSet (miNueva + (x / miDivisor), y / miDivisor), QBColor(7)  ' Dibuja confetti
      Circle (miNueva + (x / miDivisor), y / miDivisor), 1 / 10, QBColor(7)  ' Dibuja confetti
      'Next r
    End If

    x = x + 1
    'If x > 20250 Then
    If x > miColumnas Then
      miColumnas = miColumnas + 1
      y = y + 1
      x = 1
    End If
    'DoEvents
  Next p
End Sub


' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function


' AL HACER DOBLE CLICK
'Private Sub Form_DblClick()
'    Dim x As Long
'    Dim y As Long
'    Dim p As Long
'    Dim r As Integer
'
'    y = 200
'    x = 200
'
'    'For p = 1 To 3402
'    For p = 1 To 10000
'        ' Pinta primo y no primo de distinto color
'        If Primo(p) Then
'            For r = 0 To 100
'                Circle (x, y), r, QBColor(12)  ' Dibuja confetti
'            Next
'        Else
'            'For r = 0 To 100
'                Circle (x, y), 100, QBColor(2)  ' Dibuja confetti
'            'Next r
'        End If
'
'        x = x + 250
'        'If x > 20250 Then
'        If x > 52500 Then
'            y = y + 250
'            x = 200
'        End If
'        DoEvents
'    Next p
'End Sub


'' AL HACER DOBLE CLICK
'Private Sub Form_DblClick()
'    Dim x As Long
'    Dim y As Long
'    Dim p As Long
'    Dim r As Integer
'
'    y = 1
'    x = 1
'    Cls
'    'For p = 1 To 3402
'    For p = 1 To 100000
'        ' Pinta primo y no primo de distinto color
'        If Primo(p) Then
'            If Primo(p + 2) Then
'                PSet (x / 4, y / 4), QBColor(12) ' Dibuja confetti
'                If miEtiqueta = True Then
'                    Print p
'                End If
'                Circle (x / 4, y / 4), 1 / 10, QBColor(12) ' Dibuja confetti
'            Else
'            'For r = 0 To 100
'                PSet (x / 4, y / 4), QBColor(6) ' Dibuja confetti
'                Circle (x / 4, y / 4), 1 / 10, QBColor(6) ' Dibuja confetti
'            'Next
'            End If
'        Else
'            'For r = 0 To 100
'                PSet (x / 4, y / 4), QBColor(13) ' Dibuja confetti
'            'Circle (x / 2, y / 2), 1 / 5, QBColor(14) ' Dibuja confetti
'            'Next r
'        End If
'
'        x = x + 1
'        'If x > 20250 Then
'        If x >= miColumnas Then
'            y = y + 1
'            x = 1
'        End If
'        'DoEvents
'    Next p
'End Sub


'' AL HACER DOBLE CLICK
'Private Sub Form_DblClick()
'    Dim x As Long
'    Dim y As Long
'    Dim p As Long
'    Dim r As Integer
'
'    y = 1
'    x = 1
'    Cls
'    'For p = 1 To 3402
'    For p = 0 To 100000
'
'        ' Pinta primo y no primo de distinto color
'        If Primo(p) Then
'            If Primo(p + 2) Then
'                PSet (miNueva + (x / miDivisor), y / miDivisor), QBColor(0) ' Dibuja confetti
'                If miEtiqueta = True Then
'                    Print p
'                End If
'                Circle (miNueva + (x / miDivisor), y / miDivisor), 1 / 10, QBColor(0)  ' Dibuja confetti
'            Else
'            'For r = 0 To 100
'                PSet (miNueva + (x / miDivisor), y / miDivisor), QBColor(12)  ' Dibuja confetti
'                Circle (miNueva + (x / miDivisor), y / miDivisor), 1 / 10, QBColor(12) ' Dibuja confetti
'            'Next
'            End If
'        Else
'            'For r = 0 To 100
'                PSet (miNueva + (x / miDivisor), y / miDivisor), QBColor(7) ' Dibuja confetti
'                Circle (miNueva + (x / miDivisor), y / miDivisor), 1 / 10, QBColor(7) ' Dibuja confetti
'            'Next r
'        End If
'
'        x = x + 1
'        'If x > 20250 Then
'        If x >= miColumnas Then
'            miColumnas = miColumnas + 1
'            y = y + 1
'            x = 1
'        End If
'        'DoEvents
'    Next p
'End Sub
'

