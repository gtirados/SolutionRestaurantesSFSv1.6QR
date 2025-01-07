VERSION 5.00
Begin VB.Form frmComandaDescuentos 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elija Descuentos"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optsubfamilia 
      BackColor       =   &H8000000C&
      Caption         =   "Por Subfamilia"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   720
      Left            =   7440
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   720
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtDescuento 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.OptionButton optFamilia 
      BackColor       =   &H8000000C&
      Caption         =   "Por Familia"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton opttotal 
      BackColor       =   &H8000000C&
      Caption         =   "Total"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton optPorcentual 
      BackColor       =   &H8000000C&
      Caption         =   "Porcentual"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
   End
End
Attribute VB_Name = "frmComandaDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gTIPO As String
Public gDSCTO As Double
Private vPUNTO As Boolean

Private Sub cmdAceptar_Click()

    If Me.optPorcentual.Value And Len(Trim(Me.txtDescuento.Text)) = 0 Then
        MsgBox "Debe ingresar el Descuento.", vbInformation, Pub_Titulo
    ElseIf Me.opttotal.Value And Len(Trim(Me.txtDescuento.Text)) = 0 Then
        MsgBox "Debe ingresar el Descuento.", vbInformation, Pub_Titulo
    Else
        gAcepta = True

        If Me.optFamilia.Value Then
            gTIPO = 3 '"F"
        ElseIf Me.optPorcentual.Value Then
            gTIPO = 2 '"P"
        ElseIf Me.opttotal Then
            gTIPO = 1 '"T"
        ElseIf Me.optsubfamilia Then
            gTIPO = 4  '"S"
        End If

        gDSCTO = IIf(Len(Trim(Me.txtDescuento.Text)) = 0, 0, Me.txtDescuento.Text)
        Unload Me
    End If
  
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
gAcepta = False
End Sub

Private Sub optFamilia_Click()
 If Me.optFamilia.Value Then
        Me.txtDescuento.Text = ""
        Me.txtDescuento.Visible = False
    End If
End Sub

Private Sub optPorcentual_Click()

    If Me.optPorcentual.Value Then
        Me.txtDescuento.Text = ""
        Me.txtDescuento.Visible = True
        Me.txtDescuento.SetFocus
    End If

End Sub

Private Sub optsubfamilia_Click()
If Me.optsubfamilia.Value Then
        Me.txtDescuento.Text = ""
        Me.txtDescuento.Visible = False
    End If
End Sub

Private Sub opttotal_Click()
 If Me.opttotal.Value Then
        Me.txtDescuento.Text = ""
        Me.txtDescuento.Visible = True
        Me.txtDescuento.SetFocus
    End If
End Sub

Private Sub txtDescuento_Change()
If InStr(Me.txtDescuento.Text, ".") Then
    vPUNTO = True
Else
    vPUNTO = False
End If
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
If NumerosyPunto(KeyAscii) Then KeyAscii = 0
 If KeyAscii = 46 Then
    If vPUNTO Or Len(Trim(Me.txtDescuento.Text)) = 0 Then
        KeyAscii = 0
    End If
    End If
End Sub
