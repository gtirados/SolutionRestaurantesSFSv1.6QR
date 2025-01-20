VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form frmFacComandaFP2credito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Días de Crédito"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
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
   ScaleHeight     =   2520
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroDiasCredito 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      _Version        =   786433
      _ExtentX        =   5318
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "DiasCredito"
      Appearance      =   6
      Begin XtremeSuiteControls.PushButton pbAumentar 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   375
         _Version        =   786433
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "+"
         Appearance      =   4
         DrawFocusRect   =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiasCredito 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   495
         _Version        =   786433
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   2
         MaxLength       =   2
      End
      Begin XtremeSuiteControls.PushButton pbDisminuir 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   375
         _Version        =   786433
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "-"
         Appearance      =   4
         DrawFocusRect   =   0   'False
      End
      Begin XtremeSuiteControls.Label lblFechaVenc 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2775
         _Version        =   786433
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Días de Crédito:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   1410
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   855
      _Version        =   786433
      _ExtentX        =   1508
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "PushButton1"
      Appearance      =   5
      DrawFocusRect   =   0   'False
      Picture         =   "frmFacComandaFP2credito.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdCancelar 
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   855
      _Version        =   786433
      _ExtentX        =   1508
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "PushButton1"
      Appearance      =   5
      DrawFocusRect   =   0   'False
      Picture         =   "frmFacComandaFP2credito.frx":1CDA
   End
End
Attribute VB_Name = "frmFacComandaFP2credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gDiasCredito As Integer
Public gAcepta As Boolean

Private Sub cmdAceptar_Click()
If Len(Trim(Me.txtDiasCredito.Text)) = 0 Then
    MsgBox "Debe ingresar los días de crédito.", vbInformation, Pub_Titulo
    Exit Sub
End If
If Not IsNumeric(Me.txtDiasCredito.Text) Then
    MsgBox "Días de crédito incorrecto.", vbCritical, Pub_Titulo
    Me.txtDiasCredito.SetFocus
    Exit Sub
End If
If val(Me.txtDiasCredito.Text) = 0 Then
     MsgBox "Días de crédito incorrecto.", vbCritical, Pub_Titulo
    Me.txtDiasCredito.SetFocus
    Exit Sub
End If
gAcepta = True
gDiasCredito = Me.txtDiasCredito.Text
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.txtDiasCredito.Text = 1
gAcepta = False
End Sub

Private Sub pbAumentar_Click()
If Me.txtDiasCredito.Text = 99 Then Exit Sub
Me.txtDiasCredito.Text = val(Me.txtDiasCredito.Text) + 1
End Sub

Private Sub pbDisminuir_Click()
If Me.txtDiasCredito.Text = 1 Then Exit Sub
Me.txtDiasCredito.Text = val(Me.txtDiasCredito.Text) - 1
End Sub

Private Sub txtDiasCredito_Change()
If IsNumeric(Me.txtDiasCredito.Text) Then
Me.lblFechaVenc.Caption = DateAdd("d", Me.txtDiasCredito.Text, LK_FECHA_DIA)
End If
End Sub

Private Sub txtDiasCredito_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub

Private Sub txtDiasCredito_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' click derecho
If Button = 2 Then MsgBox "Acción no permitida", vbCritical, Pub_Titulo

End Sub
