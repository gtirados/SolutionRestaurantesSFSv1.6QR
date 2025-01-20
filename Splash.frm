VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "con"
   ClientHeight    =   6030
   ClientLeft      =   825
   ClientTop       =   1155
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin XtremeSuiteControls.ProgressBar pbSplash 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   5775
      _Version        =   786433
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblMarca 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   11
      Top             =   4560
      Width           =   3075
   End
   Begin XtremeSuiteControls.Label Empresa 
      Height          =   885
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   5955
      _Version        =   786433
      _ExtentX        =   10504
      _ExtentY        =   1561
      _StockProps     =   79
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion de Restaurantes y Bares"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5790
   End
   Begin VB.Label LblMensa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label lblporcentaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Autorizado a:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   90
      TabIndex        =   4
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EAC793&
      Height          =   105
      Left            =   0
      TabIndex        =   3
      Top             =   1170
      Width           =   6135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Solution"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   720
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "for Business"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2130
   End
   Begin VB.Label lbl_Top 
      BackColor       =   &H8000000D&
      Height          =   1170
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   0
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3405
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Height          =   810
      Left            =   0
      TabIndex        =   9
      Top             =   5280
      Width           =   6135
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iStatusBarWidth As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE
If KeyCode = 120 Then
    Splash.WindowState = 1
    Splash.Caption = "Proceso Abortado!!! . . ."
    Screen.MousePointer = 0
    EN.Close
    CN.Close
    Pub_ConnAdo.Close
    Screen.MousePointer = 0
    End
End If
Exit Sub
SALE:
End
End Sub
Private Sub Form_Load()
Me.lblMarca.Caption = "Copyright" & " © " & Year(Now) & vbCrLf & Leer_Ini(App.Path & "\config.ini", "TITULAR", "GT SOFTWARE") & "." & vbCrLf & "All rights reserved."
Dim wflag_bloq As String * 1
Dim success%
Dim PB
PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
'On Error GoTo SALE
Screen.MousePointer = 11
If App.PrevInstance Then
  pub_mensaje = App.Path & " " & "Software"
  pub_mensaje = pub_mensaje & PB & "Posiblemente la Aplicación este cargada o no ha sido cerrada Correctamente "
  pub_mensaje = pub_mensaje & PB & "Debe Cerrar todos los Programas e Iniciar la seccion como Usuario Distinto ..."
  MsgBox pub_mensaje, vbCritical, "Software"
  Screen.MousePointer = 0
  End
End If

Pub_Titulo = "UniSoft S.A.C. - Solution"
LK_CODCIA = ""
LK_CODUSU = ""
If Nulo_Valor0(PUB_FLAG) = 0 Then
  wflag_bloq = ""
  If dir("C:\WINDOWS\Sisgts", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If dir("C:\Winnt\Sisgts", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If dir("C:\Win98\Sisgts", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If wflag_bloq <> "A" Then
    MsgBox "Equipo: MicroProcesador No Identificado..." & Chr(13) & "- Esta copia del Ejecutable no procede - No tiene licencia de uso", vbCritical, "Proveedor del Software - Celular: #990905152"
    End
  End If

  Splash.Show
  'success% = SetWindowPos(Splash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  DoEvents
  CONEXION_GEN
End If
PUB_FLAG = 0
DoEvents
'success% = SetWindowPos(Splash.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
Load MDIForm1
Screen.MousePointer = 0
Exit Sub
SALE:
 Screen.MousePointer = 0
 MsgBox Err.Description, 48, "pub_titulo"
End

End Sub


