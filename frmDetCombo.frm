VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmDetCombo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8640
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvListado 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblNumSec 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblNumFac 
      Caption         =   "Label2"
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblNumSer 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
   End
End
Attribute VB_Name = "frmDetCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub lvListado_DblClick()
 If Me.lvListado.ListItems.count = 0 Then Exit Sub
        frmComandaProdCaracteristicas.gIDproducto = Me.lvListado.SelectedItem.Tag
        frmComandaProdCaracteristicas.gNUMFAC = Me.lblNumFac.Caption
        frmComandaProdCaracteristicas.gNUMSER = Me.lblNumSer.Caption
        frmComandaProdCaracteristicas.gNUMSEC = Me.lblNumSec.Caption
        frmComandaProdCaracteristicas.Show vbModal
End Sub
