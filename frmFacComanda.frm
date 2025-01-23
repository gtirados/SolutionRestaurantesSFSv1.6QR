VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form frmFacComanda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturar Comanda"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacComanda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   855
      Left            =   30
      TabIndex        =   53
      Top             =   9600
      Width           =   1935
      _Version        =   786433
      _ExtentX        =   3413
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "&Aceptar"
      Appearance      =   4
      DrawFocusRect   =   0   'False
      Picture         =   "frmFacComanda.frx":1CCA
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   1680
      TabIndex        =   24
      Top             =   3360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeSuiteControls.GroupBox gbEmpresa 
      Height          =   1140
      Left            =   30
      TabIndex        =   14
      Top             =   0
      Width           =   9405
      _Version        =   786433
      _ExtentX        =   16589
      _ExtentY        =   2011
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.PushButton pbEmpresa 
         Height          =   975
         Index           =   0
         Left            =   6600
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _Version        =   786433
         _ExtentX        =   2355
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "PushButton3"
         Appearance      =   6
         DrawFocusRect   =   0   'False
      End
      Begin XtremeSuiteControls.PushButton pbEmpresaAnt 
         Height          =   975
         Left            =   30
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _Version        =   786433
         _ExtentX        =   2355
         _ExtentY        =   1720
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   6
         DrawFocusRect   =   0   'False
         Picture         =   "frmFacComanda.frx":39A4
      End
      Begin XtremeSuiteControls.PushButton pbEmpresaSig 
         Height          =   975
         Left            =   8040
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _Version        =   786433
         _ExtentX        =   2355
         _ExtentY        =   1720
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   6
         DrawFocusRect   =   0   'False
         Picture         =   "frmFacComanda.frx":567E
      End
   End
   Begin XtremeSuiteControls.GroupBox gbTipoDoc 
      Height          =   1140
      Left            =   30
      TabIndex        =   15
      Top             =   1080
      Width           =   6300
      _Version        =   786433
      _ExtentX        =   11112
      _ExtentY        =   2011
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.PushButton pbDocAnt 
         Height          =   975
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   1245
         _Version        =   786433
         _ExtentX        =   2196
         _ExtentY        =   1720
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   6
         DrawFocusRect   =   0   'False
         Picture         =   "frmFacComanda.frx":7358
      End
      Begin XtremeSuiteControls.PushButton pbDocSig 
         Height          =   975
         Left            =   5010
         TabIndex        =   4
         Top             =   120
         Width           =   1245
         _Version        =   786433
         _ExtentX        =   2196
         _ExtentY        =   1720
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   6
         DrawFocusRect   =   0   'False
         Picture         =   "frmFacComanda.frx":9032
      End
      Begin XtremeSuiteControls.PushButton pbDoc 
         Height          =   975
         Index           =   0
         Left            =   2040
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
         _Version        =   786433
         _ExtentX        =   2196
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Documento"
         Appearance      =   6
         DrawFocusRect   =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1140
      Left            =   6360
      TabIndex        =   16
      Top             =   1080
      Width           =   3075
      _Version        =   786433
      _ExtentX        =   5424
      _ExtentY        =   2011
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.CheckBox chkEdit 
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Editar Número"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtNro 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   1575
         _Version        =   786433
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         MaxLength       =   10
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblSerie 
         Height          =   375
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   975
         _Version        =   786433
         _ExtentX        =   1720
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   615
      Left            =   30
      TabIndex        =   18
      Top             =   2160
      Width           =   9405
      _Version        =   786433
      _ExtentX        =   16589
      _ExtentY        =   1085
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.CheckBox chkprom 
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
         _Version        =   786433
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Imprime Promoción"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkConsumo 
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
         _Version        =   786433
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Facturar x Consumo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkGratuita 
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   3015
         _Version        =   786433
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Transferencias Gratuita"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
   End
   Begin XtremeSuiteControls.GroupBox gbCliente 
      Height          =   1695
      Left            =   30
      TabIndex        =   19
      Top             =   2760
      Width           =   9405
      _Version        =   786433
      _ExtentX        =   16589
      _ExtentY        =   2990
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.PushButton cmdSunat 
         Height          =   375
         Left            =   8400
         TabIndex        =   26
         Top             =   240
         Width           =   855
         _Version        =   786433
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Sunat"
         Appearance      =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRS 
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   4095
         _Version        =   786433
         _ExtentX        =   7223
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   7575
         _Version        =   786433
         _ExtentX        =   13361
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDni 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   4095
         _Version        =   786433
         _ExtentX        =   7223
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRuc 
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         Top             =   240
         Width           =   1695
         _Version        =   786433
         _ExtentX        =   2990
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RAZÓN SOCIAL"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº RUC"
         Height          =   195
         Left            =   5880
         TabIndex        =   22
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCIÓN:"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   810
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI:"
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   1290
         Width           =   405
      End
   End
   Begin XtremeSuiteControls.GroupBox GroDetalle 
      Height          =   2415
      Left            =   30
      TabIndex        =   27
      Top             =   4440
      Width           =   9405
      _Version        =   786433
      _ExtentX        =   16589
      _ExtentY        =   4260
      _StockProps     =   79
      Appearance      =   6
      Begin MSComctlLib.ListView lvDetalle 
         Height          =   2265
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin XtremeSuiteControls.GroupBox GroMontos 
      Height          =   2655
      Left            =   30
      TabIndex        =   29
      Top             =   6840
      Width           =   9405
      _Version        =   786433
      _ExtentX        =   16589
      _ExtentY        =   4683
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.PushButton cmdFormasPago 
         Height          =   720
         Left            =   240
         TabIndex        =   55
         Top             =   840
         Width           =   2775
         _Version        =   786433
         _ExtentX        =   4895
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "Formas de Pago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Picture         =   "frmFacComanda.frx":AD0C
      End
      Begin VB.CommandButton cmdCobrar 
         Caption         =   "&Cobrar"
         Height          =   360
         Left            =   2160
         TabIndex        =   52
         Top             =   1080
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdDscto 
         Caption         =   "Descuento"
         Height          =   375
         Left            =   480
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin XtremeSuiteControls.PushButton pbAumentar 
         Height          =   375
         Left            =   2400
         TabIndex        =   46
         Top             =   2040
         Width           =   375
         _Version        =   786433
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "+"
         Appearance      =   4
         DrawFocusRect   =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox CheckBox3 
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   260
         Width           =   1215
         _Version        =   786433
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "SERVICIO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   6
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCopias 
         Height          =   375
         Left            =   1920
         TabIndex        =   45
         Top             =   2040
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
         Left            =   1560
         TabIndex        =   47
         Top             =   2040
         Width           =   375
         _Version        =   786433
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "-"
         Appearance      =   4
         DrawFocusRect   =   0   'False
      End
      Begin VB.Label lblVUELTO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label labelv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VUELTO:"
         Height          =   195
         Left            =   1800
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblDscto 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblNroDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Copias:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label lblpICBPER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         Height          =   195
         Left            =   4920
         TabIndex        =   43
         Top             =   1800
         Visible         =   0   'False
         Width           =   660
      End
      Begin XtremeSuiteControls.Label Label23 
         Height          =   240
         Left            =   5655
         TabIndex        =   42
         Top             =   2267
         Width           =   1680
         _Version        =   786433
         _ExtentX        =   2963
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "TOTAL A PAGAR:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label22 
         Height          =   240
         Left            =   6555
         TabIndex        =   41
         Top             =   1867
         Width           =   780
         _Version        =   786433
         _ExtentX        =   1376
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "ICBPER:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblporcigv 
         Height          =   240
         Left            =   6900
         TabIndex        =   40
         Top             =   1467
         Width           =   435
         _Version        =   786433
         _ExtentX        =   767
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "IGV:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label20 
         Height          =   240
         Left            =   6360
         TabIndex        =   39
         Top             =   1467
         Width           =   435
         _Version        =   786433
         _ExtentX        =   767
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "IGV:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label19 
         Height          =   240
         Left            =   5865
         TabIndex        =   38
         Top             =   1067
         Width           =   1470
         _Version        =   786433
         _ExtentX        =   2593
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "VALOR VENTA:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label18 
         Height          =   240
         Left            =   5445
         TabIndex        =   37
         Top             =   667
         Width           =   1890
         _Version        =   786433
         _ExtentX        =   3334
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "OPER. GRATUITAS:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblImporte 
         Height          =   375
         Left            =   7440
         TabIndex        =   35
         Top             =   2200
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblICBPER 
         Height          =   375
         Left            =   7440
         TabIndex        =   34
         Top             =   1800
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblIGV 
         Height          =   375
         Left            =   7440
         TabIndex        =   33
         Top             =   1400
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblvvta 
         Height          =   375
         Left            =   7440
         TabIndex        =   32
         Top             =   1000
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblgratuita 
         Height          =   375
         Left            =   7440
         TabIndex        =   31
         Top             =   600
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblServicio 
         Height          =   375
         Left            =   7440
         TabIndex        =   30
         Top             =   200
         Width           =   1815
         _Version        =   786433
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.PushButton cmdCancelar 
      Height          =   855
      Left            =   2040
      TabIndex        =   54
      Top             =   9600
      Width           =   1935
      _Version        =   786433
      _ExtentX        =   3413
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "&Cancelar"
      Appearance      =   4
      DrawFocusRect   =   0   'False
      Picture         =   "frmFacComanda.frx":C9E6
   End
End
Attribute VB_Name = "frmFacComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vSerCom As String
Private buscars As Boolean
Public vNroCom, vCodMoz As Integer
Public vMesa As String
Public vTOTAL As Currency
Dim loc_key  As Integer
Private oRsPago As ADODB.Recordset
Private oRsTipPag As ADODB.Recordset
Public vAcepta As Boolean
Public vOper As Integer 'variable para capturar el allnumoper de allog para imprimir del facart
Private vBuscar As Boolean 'variable para la busqueda de clientes
Public xMostrador As Boolean
Public gDESCUENTO As Double 'VARIABLE PARA ALMACENAR EL DESCUENTO PARA LAS COMANDAS
Public gPAGO As Double 'VARIABLE PARA ALMACENAR EL PAGO PARA LAS COMANDAS
Private ORStd As ADODB.Recordset 'VARIABLE PARA SABER SI EL TIPO DE DOCUMENTO ES EDITABLE
Private vPagActEmp, vPagTotEmp As Integer
Private vPagActDoc, vPagTotDoc As Integer
Private pCodEmp As String
Private pCodTipDoc As String
Private pDesTipDoc As String

Private Function VerificaPassPrecios(vUSUARIO As String, vClave As String, ByRef vMSN As String) As Boolean
Dim orsPass As ADODB.Recordset
Dim vtpass As String, vPasa As Boolean
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpDevuelveClaveprecios"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
oCmdEjec.Execute

'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
vtpass = oCmdEjec.Parameters("@MSN").Value
vPasa = oCmdEjec.Parameters("@PASA").Value
vMSN = vtpass

    VerificaPassPrecios = vPasa
End Function
Private Sub ConfigurarLVDetalle()
With Me.lvDetalle
    .ColumnHeaders.Add , , "NumFac", 0
    .ColumnHeaders.Add , , "Codigo", 800
    .ColumnHeaders.Add , , "Plato", 4000
    .ColumnHeaders.Add , , "Cta", 500
    .ColumnHeaders.Add , , "Precio", 800
    .ColumnHeaders.Add , , "Can. Tot.", 1000
    .ColumnHeaders.Add , , "Faltante", 1000
    .ColumnHeaders.Add , , "Importe", 1200
    .ColumnHeaders.Add , , "Sec", 0
    .ColumnHeaders.Add , , "apro", 0
    .ColumnHeaders.Add , , "aten", 0
    .ColumnHeaders.Add , , "unidad", 0
    .ColumnHeaders.Add , , "pednumsec", 0
    .ColumnHeaders.Add , , "cambio", 800
    .ColumnHeaders.Add , , "ICBPER", 300
    .ColumnHeaders.Add , , "comboICBPER", 300
    .HideColumnHeaders = False
    .View = lvwReport
    .FullRowSelect = True
    .LabelEdit = lvwManual
    .Gridlines = True
    .HideSelection = False
End With
End Sub

Private Function VerificaPass(vUSUARIO As String, vClave As String, ByRef vMSN As String) As Boolean
Dim orsPass As ADODB.Recordset
Dim vtpass As String, vPasa As Boolean
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpDevuelveClaveCaja"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
oCmdEjec.Execute

'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
vtpass = oCmdEjec.Parameters("@MSN").Value
vPasa = oCmdEjec.Parameters("@PASA").Value
vMSN = vtpass

    VerificaPass = vPasa
End Function
Private Function VerificaPassprecio(vUSUARIO As String, vClave As String, ByRef vMSN As String) As Boolean
Dim orsPass As ADODB.Recordset
Dim vtpass As String, vPasa As Boolean
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpDevuelveClaveprecios"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
oCmdEjec.Execute

'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
vtpass = oCmdEjec.Parameters("@MSN").Value
vPasa = oCmdEjec.Parameters("@PASA").Value
vMSN = vtpass

    VerificaPassprecio = vPasa
End Function




Private Sub ConfiguraLV()
With Me.ListView1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "Ruc", 0
    .ColumnHeaders.Add , , "Direcion", 0
    .MultiSelect = False
    .HideSelection = False
    .HideColumnHeaders = False
End With
End Sub



Private Sub cboMoneda_KeyPress(index As Integer, KeyAscii As Integer)
If LK_TIPO_CAMBIO = 0 Then
    MsgBox "ingresar tipo de cambio"
End If
End Sub


Private Sub chkEdit_Click()
Me.txtNro.Enabled = Me.chkEdit.Value
If Me.chkEdit.Value Then
Me.txtNro.SetFocus
Me.txtNro.SelStart = 0
Me.txtNro.SelLength = Len(Me.txtNro.Text)
End If
End Sub



Private Sub chkGratuita_Click()

    If Me.chkGratuita.Value Then

        frmClaveCaja.Show vbModal
    
        If frmClaveCaja.vAceptar Then
    
            Dim vS As String
    
            ' If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
            If VerificaPassPrecios(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                'Me.lblgratuita.Caption = Me.lblImporte
                Me.lblgratuita.Caption = Format(val(Me.lblImporte.Caption), "########0.#0")
                Me.lblvvta.Caption = "0.00"
                Me.lblIgv.Caption = "0.00"
                Me.lblImporte.Caption = "0.00"
               
                Me.cmdformaspago.Enabled = False
            Else
                MsgBox "Clave incorrecta", vbCritical, NombreProyecto

            End If

        End If

    Else
        Me.cmdformaspago.Enabled = True
        CalcularImporte

    End If
End Sub

Private Sub cmdAceptar_Click()

If Len(Trim(Me.txtcopias.Text)) = 0 Then
        MsgBox "Debe ingresar el nro de copias a imprimir.", vbInformation, Pub_Titulo
        Me.txtcopias.SetFocus
        Exit Sub
    End If
    
    If val(Me.txtcopias.Text) <= 0 Then
    ' MsgBox "Nro de copias incorrecto.", vbInformation, Pub_Titulo
     '   Me.txtcopias.SetFocus
     '   Exit Sub
    End If

If pCodTipDoc = "" Then
   MsgBox "Debe elegir el Tipo de documento.", vbCritical, Pub_Titulo

    Exit Sub
End If

If Me.cmdformaspago.Enabled And oRSfp.RecordCount = 0 Then
    MsgBox "Debe ingresar pagos", vbCritical, Pub_Titulo

    Exit Sub

End If
    
'    If oRSfp.RecordCount <> 0 Then
'        oRSfp.MoveFirst
'        oRSfp.Filter = "IDFORMAPAGO=4"
'        If oRSfp.RecordCount <> 0 And Len(Trim(Me.txtRuc.Tag)) = 0 Then
'            MsgBox "Debe elegir el cliente.", vbCritical, Pub_Titulo
'            oRSfp.Filter = ""
'        oRSfp.MoveFirst
'        Exit Sub
'        End If
'    End If

Dim xPAGOS         As Double

Dim xCONTINUA      As Boolean

Dim xARCENCONTRADO As Boolean

xCONTINUA = False
xARCENCONTRADO = False
xPAGOS = 0

oRSfp.MoveFirst

Do While Not oRSfp.EOF
    xPAGOS = xPAGOS + oRSfp!monto
    oRSfp.MoveNext
Loop
    
If Me.cmdformaspago.Enabled And xPAGOS < val(Me.lblImporte.Caption) Then
    MsgBox "Falta importe por pagar", vbCritical, Pub_Titulo

    Exit Sub

End If

On Error GoTo Graba

'If Me.cboTipoDocto.ListIndex = 0 Then 'F

If pCodTipDoc = "01" Then
    If Me.lvDetalle.ListItems.count > par_llave!par_fac_lines And Me.chkConsumo.Value = 0 Then
        MsgBox "Numero Máximo de Lineas alcanzado"

        Exit Sub

    End If

'ElseIf Me.cboTipoDocto.ListIndex = 1 Then 'B
ElseIf pCodTipDoc = "03" Then 'B

    If Me.lvDetalle.ListItems.count > par_llave!par_BOL_lines And Me.chkConsumo.Value = 0 Then
        MsgBox "Numero Máximo de Lineas alcanzado"

        Exit Sub

    End If
End If

Dim f As Integer

Dim sOri, sMod As Integer

If Me.lvDetalle.ListItems.count = 0 Then
    MsgBox "No hay ningun plato para procesar"

    Exit Sub

End If

If pCodTipDoc = "01" Then
    If Len(Trim(Me.txtRuc.Text)) = 0 Then
        MsgBox "Debe ingresar el Ruc para poder generar la Factura", vbInformation, "Error"

        Exit Sub

    End If
End If

'valida la uit
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_VALIDA_UIT"

Dim ORSuit As ADODB.Recordset

Set ORSuit = oCmdEjec.Execute(, Me.lblImporte.Caption)

If Not ORSuit.EOF Then
    If ORSuit!Dato = 1 Then
        If Len(Trim(Me.txtRS.Text)) = 0 Or (Len(Trim(Me.txtDni.Text)) = 0 And Len(Trim(Me.txtRuc.Text)) = 0) Then
            MsgBox "El Importe sobrepasa media UIT, debe ingresar el cliente.", vbCritical, Pub_Titulo

            Exit Sub

        End If
    End If
End If
    
If Len(Trim(Me.txtDni.Text)) <> 0 Then
    If Len(Trim(Me.txtDni.Text)) < 8 Then
        MsgBox "El DNI debe tener 8 dígitos.", vbInformation, Pub_Titulo

        Exit Sub

    End If
End If

'Armando Xml
Dim CP As Double

Dim pr, st As Currency

Dim vXml, un, cam As String

vXml = "<r>"

For f = 1 To Me.lvDetalle.ListItems.count

    CP = Me.lvDetalle.ListItems(f).SubItems(1)
'    Me.dgrdData.Row = f
'    'Codigo de Plato
'    Me.dgrdData.COL = 1
'    CP = CInt(Trim(Me.dgrdData.Text))
'Cantidad a facturar
    st = Me.lvDetalle.ListItems(f).SubItems(6)
'    Me.dgrdData.COL = 5
'    st = CInt(Trim(Me.dgrdData.Text))
'Precio
'Me.dgrdData.COL = 6
'pr = CDec(Trim(Me.dgrdData.Text))
    pr = Me.lvDetalle.ListItems(f).SubItems(4)
'Unidad de medida
'Me.dgrdData.COL = 10
'un = Trim(Me.dgrdData.Text)
    un = Trim(Me.lvDetalle.ListItems(f).SubItems(11))
'SECUENCIA
'Me.dgrdData.COL = 7
'    sc = Trim(Me.dgrdData.Text)
    sc = Me.lvDetalle.ListItems(f).SubItems(8)
    cam = Me.lvDetalle.ListItems(f).SubItems(13)
    
    vXml = vXml & "<d "
    vXml = vXml & "cp=""" & Trim(str(CP)) & """ "
    vXml = vXml & "st=""" & Trim(str(st)) & """ "
    vXml = vXml & "pr=""" & Trim(str(pr)) & """ "
    vXml = vXml & "un=""" & Trim(un) & """ "
    vXml = vXml & "sc=""" & Trim(sc) & """ "
    vXml = vXml & "cam=""" & Trim(cam) & """ "
    vXml = vXml & "/>"
Next

vXml = vXml & "</r>"
 
'obteniendo datos de tipo de pago tabla sub_transa
'Dim alltipdoc As String
'Dim allcp As String
'oRsTipPag.Filter = "sut_secuencia=" & Me.dcboPago.BoundText
'alltipdoc = oRsTipPag!sut_tipdoc
'allcp = oRsTipPag!sut_cp
    
'recorriendo las formas de pago
Dim xFP      As String

Dim xPAGACON As Double, xVUELTO As Double
    
If oRSfp.RecordCount <> 0 Then
    oRSfp.MoveFirst
    xPAGACON = oRSfp!pagacon
    xVUELTO = oRSfp!VUELTO
    xFP = "<r>"

    Do While Not oRSfp.EOF
        xFP = xFP & "<d "
        xFP = xFP & "idfp=""" & Trim(str(oRSfp!idformapago)) & """ "
        xFP = xFP & "fp=""" & Trim(oRSfp!formapago) & """ "
        xFP = xFP & "mon=""" & "S" & """ "
        xFP = xFP & "monto=""" & Trim(str(oRSfp!monto)) & """ "
        xFP = xFP & "ref=""" & Trim(oRSfp!referencia) & """ "
        xFP = xFP & "dcre=""" & Trim(oRSfp!diascredito) & """ "
        xFP = xFP & "/>"
        oRSfp.MoveNext
    Loop

    xFP = xFP & "</r>"
End If

With oCmdEjec
        
'PARCHE - SE TENDRIA QUE HACER MEJOR EN EL SP PARA UNA MEJOR CONSISTENCIA DE DATOS
    If vAcepta Then
        MsgBox "Ya se facturo.", vbInformation, Pub_Titulo

        Exit Sub

    Else
'VALIDANDO SI EL ARCHIVO DEL REPORTE EXISTE
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_ARCHIVO_PRINT"
        oCmdEjec.CommandType = adCmdStoredProc
    
        Dim ORSd        As ADODB.Recordset

        Dim RutaReporte As String

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, pCodTipDoc)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMSUMO", adBoolean, adParamInput, , Me.chkConsumo.Value)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, pCodEmp)
    
        Set ORSd = oCmdEjec.Execute
        RutaReporte = PUB_RUTA_REPORTE & ORSd!ReportE
    
        If ORSd!ReportE = "" Then
            If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                xCONTINUA = True
            End If

        Else
            FileName = dir(RutaReporte)

            If FileName = "" Then
                If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                    xCONTINUA = True
                Else

                    Exit Sub

                End If

            Else
                xCONTINUA = True
                xARCENCONTRADO = True
            End If
        End If

        If xCONTINUA Then
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpFacturarComanda"
            .Parameters.Append .CreateParameter("@codcia", adChar, adParamInput, 2, pCodEmp)
            .Parameters.Append .CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            .Parameters.Append .CreateParameter("@usuario", adVarChar, adParamInput, 20, LK_CODUSU)
            .Parameters.Append .CreateParameter("@SerCom", adChar, adParamInput, 3, vSerCom)
            .Parameters.Append .CreateParameter("@nroCom", adInteger, adParamInput, , vNroCom)
            .Parameters.Append .CreateParameter("@SerDoc", adChar, adParamInput, 3, Me.lblSerie.Caption)
            .Parameters.Append .CreateParameter("@NroDoc", adDouble, adParamInput, , CDbl(Me.txtNro.Text))
            .Parameters.Append .CreateParameter("@Fbg", adChar, adParamInput, 1, Left(pDesTipDoc, 1)) ' Me.cboTipoDocto.Text)
            .Parameters.Append .CreateParameter("@XmlDet", adVarChar, adParamInput, 4000, Trim(vXml))

            If pCodTipDoc = "01" Then
                .Parameters.Append .CreateParameter("@codcli", adVarChar, adParamInput, 10, Me.txtRuc.Tag)
            Else
                .Parameters.Append .CreateParameter("@codcli", adVarChar, adParamInput, 10, IIf(Len(Trim(Me.txtRuc.Tag)) = 0, 1, Me.txtRuc.Tag))
            End If

            .Parameters.Append .CreateParameter("@codMozo", adInteger, adParamInput, , vCodMoz)
            .Parameters.Append .CreateParameter("@totalfac", adDouble, adParamInput, , Me.lblImporte.Caption)
            If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
            '.Parameters.Append .CreateParameter("@sec", adInteger, adParamInput, , Me.dcboPago.BoundText)
            .Parameters.Append .CreateParameter("@sec", adInteger, adParamInput, , oRSfp!idformapago)
            .Parameters.Append .CreateParameter("@moneda", adChar, adParamInput, 1, "S")
            .Parameters.Append .CreateParameter("@diascre", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("@farjabas", adTinyInt, adParamInput, , IIf(Me.chkConsumo.Value = 1, 1, 0))
'.Parameters.Append .CreateParameter("@dscto", adDouble, adParamInput, , IIf(Len(Trim(Me.lblDscto.Caption)) = 0, 0, Me.lblDscto.Caption))
            .Parameters.Append .CreateParameter("@dscto", adDouble, adParamInput, , gDESCUENTO)
            .Parameters.Append .CreateParameter("@CODIGODOCTO", adChar, adParamInput, 2, pCodTipDoc)
            .Parameters.Append .CreateParameter("@Xmlpag", adVarChar, adParamInput, 4000, xFP)
            .Parameters.Append .CreateParameter("@PAGACON", adDouble, adParamInput, , xPAGACON)
            .Parameters.Append .CreateParameter("@VUELTO", adDouble, adParamInput, , xVUELTO)
                
            .Parameters.Append .CreateParameter("@VALORVTA", adDouble, adParamInput, , Me.lblvvta.Caption)
            .Parameters.Append .CreateParameter("@VIGV", adDouble, adParamInput, , Me.lblIgv.Caption)
            .Parameters.Append .CreateParameter("@GRATUITO", adBoolean, adParamInput, , Me.chkGratuita.Value)
            .Parameters.Append .CreateParameter("@CIAPEDIDO", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@ALL_ICBPER", adDouble, adParamInput, , IIf(Len(Trim(Me.lblicbper.Caption)) = 0, 0, Me.lblicbper.Caption))
            .Parameters.Append .CreateParameter("@SERVICIO", adDouble, adParamInput, , Me.lblServicio.Caption)
            .Parameters.Append .CreateParameter("@ALL_GRATUITO", adBoolean, adParamInput, , Me.chkGratuita.Value)
            .Parameters.Append .CreateParameter("@MaxNumOper", adInteger, adParamOutput, , 0)
            .Parameters.Append .CreateParameter("@AUTONUMFAC", adInteger, adParamOutput, , 0)
        
            .Execute

            If Not IsNull(oCmdEjec.Parameters("@MaxNumOper").Value) Then
                vOper = oCmdEjec.Parameters("@MaxNumOper").Value
            End If

            Me.txtNro.Text = oCmdEjec.Parameters("@AUTONUMFAC").Value
                
            If vAcepta = False Then

                vAcepta = True
'MsgBox "Datos Almacenados correctamente", vbInformation, Pub_Titulo

                CreaCodigoQR "6", pCodTipDoc, Me.lblSerie.Caption, Me.txtNro.Text, LK_FECHA_DIA, CStr(Me.lblIgv.Caption), Me.lblImporte.Caption, Me.txtRuc.Text, Me.txtDni.Text
                If xARCENCONTRADO Then
                    ImprimirDocumentoVenta pCodTipDoc, pDesTipDoc, Me.chkConsumo.Value, Me.lblSerie.Caption, Me.txtNro.Text, Me.lblImporte.Caption, Me.lblvvta.Caption, Me.lblIgv.Caption, Me.txtDireccion.Text, Me.txtRuc.Text, Me.txtRS.Text, Me.txtDni.Text, pCodEmp, IIf(Len(Trim(Me.lblicbper.Caption)) = 0, 0, Me.lblicbper.Caption), Me.chkprom.Value, Me.chkGratuita.Value, Me.txtcopias.Text
                End If

                    If LK_PASA_BOLETAS = "A" And (pCodTipDoc = "01" Or pCodTipDoc = "03") Then
                        CrearArchivoPlano Left(pDesTipDoc, 1), Me.lblSerie.Caption, Me.txtNro.Text
                    ElseIf pCodTipDoc = "01" Then
                        CrearArchivoPlano Left(pDesTipDoc, 1), Me.lblSerie.Caption, Me.txtNro.Text
                    End If
               ' End If

                Unload Me
            Else
                MsgBox "Ya se facturo"
            End If

        End If
    End If

'FIN DEL PARCHE
End With
   
Exit Sub

Graba:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me

  

End Sub

Private Sub cmdCobrar_Click()
If oRSfp.RecordCount = 0 Then
        MsgBox "Debe ingresar pagos", vbCritical, Pub_Titulo
        Exit Sub
    End If

  Dim xPAGOS As Double

    xPAGOS = 0

    'If Not oRSfp.EOF Then oRSfp.MoveFirst
    Do While Not oRSfp.EOF
        xPAGOS = xPAGOS + oRSfp!monto
        oRSfp.MoveNext
    Loop
    
    If xPAGOS <= 0 Then
        MsgBox "Falta importe por pagar", vbCritical, Pub_Titulo

        Exit Sub

    End If
    

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DOCUMENTO_PAGAR"

    Dim xFP As String
    
    If oRSfp.RecordCount <> 0 Then
        oRSfp.MoveFirst
        xFP = "<r>"

        Do While Not oRSfp.EOF
            xFP = xFP & "<d "
            xFP = xFP & "idfp=""" & Trim(str(oRSfp!idformapago)) & """ "
            xFP = xFP & "fp=""" & Trim(oRSfp!formapago) & """ "
            xFP = xFP & "mon=""" & Trim(oRSfp!moneda) & """ "
            xFP = xFP & "monto=""" & Trim(str(oRSfp!monto)) & """ "
            xFP = xFP & "/>"
            oRSfp.MoveNext
        Loop

        xFP = xFP & "</r>"
    End If
    
    Dim oMSN As String

    oMSN = ""

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@SERIECOMANDA", adChar, adParamInput, 3, frmComanda.lblSerie.Caption)
        .Parameters.Append .CreateParameter("@NUMEROCOMANDA", adBigInt, adParamInput, , frmComanda.lblNumero.Caption)
        .Parameters.Append .CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        .Parameters.Append .CreateParameter("@Xmlpag", adVarChar, adParamInput, 4000, xFP)
        .Parameters.Append .CreateParameter("@totalfac", adDouble, adParamInput, , val(frmComanda.lblTot.Caption))
        .Parameters.Append .CreateParameter("@usuario", adVarChar, adParamInput, 10, LK_CODUSU)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@exito", adVarChar, adParamOutput, 300, oMSN)
        oCmdEjec.Execute

        oMSN = oCmdEjec.Parameters("@exito").Value
        
    End With
    
     If Len(Trim(oMSN)) <> 0 Then
        MsgBox oMSN, vbCritical, Pub_Titulo
    Else
        MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
        vAcepta = True
        Unload Me
    End If
    
End Sub

Private Sub cmdDscto_Click()

    frmClaveCaja.Show vbModal
    If frmClaveCaja.vAceptar Then
     Dim vS As String
     If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
        frmAsigCantFac.txtCantidad.Text = Me.lblDscto.Caption
        frmAsigCantFac.Show vbModal
        If frmAsigCantFac.vAcepta Then
            Me.lblDscto.Caption = frmAsigCantFac.vCANTIDAD
            CalcularImporte
        End If
      Else
       MsgBox "Clave incorrecta", vbCritical, NombreProyecto
      End If
    End If
    
  

   
    
    

   

End Sub

Private Sub cmdFormasPago_Click()
'frmFacComandaFP.gMostrador = xMostrador
frmFacComandaFP2.gDELIVERY = False
'frmFacComandaFP.Show vbModal
frmFacComandaFP2.lblTotalPagar.Caption = FormatCurrency(Me.lblImporte.Caption, 2)
frmFacComandaFP2.Show vbModal

End Sub

Private Sub cmdSunat_Click()

    On Error GoTo cCruc

    Dim p          As Object

    Dim TEXTO      As String, xTOk As String

    Dim CADENA     As String, xvRUC As String

    Dim sInputJson As String, xEsRuc As Boolean

    xEsRuc = True

    MousePointer = vbHourglass
    Set httpURL = New WinHttp.WinHttpRequest
    
    If IsNumeric(Me.txtRS.Text) Then
        If Len(Trim(Me.txtRS.Text)) = 8 Then
            xEsRuc = False
        End If

        xvRUC = Me.txtRS.Text
    Else

        If Len(Trim(Me.txtRS.Text)) = 8 Then
            xEsRuc = False
        End If

        xvRUC = Me.txtRuc.Text
    End If

    xTOk = Leer_Ini(App.Path & "\config.ini", "TOKEN", "")
    
    If xEsRuc Then
        CADENA = "http://dniruc.apisperu.com/api/v1/ruc/" & xvRUC & "?token=" & xTOk
    Else
        CADENA = "http://dniruc.apisperu.com/api/v1/dni/" & xvRUC & "?token=" & xTOk
    End If
    
    httpURL.Open "GET", CADENA
    httpURL.Send
    
    TEXTO = httpURL.ResponseText

    'sInputJson = "{items:" & Texto & "}"

    Set p = JSON.parse(TEXTO)

    '    Me.lblRUC.Caption = p.Item("ruc")
    '    Me.lblRazonSocial.Caption = p.Item("razonSocial")
    '    Me.lblDireccion.Caption = p.Item("direccion")
    '    Me.lblTipo.Caption = p.Item("tipo")
    '    Me.lblEstado.Caption = p.Item("estado")
    '    Me.lblcondicion.Caption = p.Item("condicion")
    
    If Len(Trim(Me.txtRuc.Text)) = 0 Then
        If IsNumeric(Me.txtRS.Text) Then
            If Len(Trim(Me.txtRS.Text)) = 11 Or Len(Trim(Me.txtRS.Text)) = 8 Then
                If TEXTO = "[]" Then
                    MousePointer = vbDefault
                    MsgBox ("No se obtuvo resultados")
                    Me.txtRuc.Text = ""
                    Me.txtRS.Text = ""
                    Me.txtDireccion.Text = ""

                    Exit Sub

                End If

                If Len(Trim(TEXTO)) = 0 Then
                    MousePointer = vbDefault
                    MsgBox ("No se obtuvo resultados")
                    Me.txtRuc.Text = ""
                    Me.txtRS.Text = ""
                    Me.txtDireccion.Text = ""

                    Exit Sub

                End If

                If xEsRuc Then
                    Me.txtDireccion.Text = IIf(IsNull(p.Item("direccion")), "", p.Item("direccion"))
                    Me.txtRS.Text = p.Item("razonSocial")
                    Me.txtRuc.Text = p.Item("ruc")
                    Me.txtDni.Text = ""
                Else
                    Me.txtRuc.Text = ""
                    Me.txtDireccion.Text = ""
                    Me.txtDni.Text = p.Item("dni")
                    Me.txtRS.Text = p.Item("nombres") & " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
                End If
    
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
                oCmdEjec.CommandType = adCmdStoredProc
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, Left(Trim(Me.txtRS.Text), 200))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Left(Trim(Me.txtDireccion.Text), 200))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Me.txtRuc.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , 0)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@salida", adBigInt, adParamOutput, , 0)
                oCmdEjec.Execute
                Me.txtRuc.Tag = oCmdEjec.Parameters("@salida").Value
            
            Else
                MsgBox "El ruc debe tener 11 caracteres", vbCritical, Pub_Titulo
            End If

        Else
            MsgBox "El ruc debe ser Numeros", vbCritical, Pub_Titulo
        End If

    Else

        If TEXTO = "[]" Then
            MousePointer = vbDefault
            MsgBox ("No se obtuvo resultados")
            Me.txtRuc.Text = ""
            Me.txtRS.Text = ""
            Me.txtDireccion.Text = ""

            Exit Sub

        End If

        If Len(Trim(TEXTO)) = 0 Then
            MousePointer = vbDefault
            MsgBox ("No se obtuvo resultados")
            Me.txtRuc.Text = ""
            Me.txtRS.Text = ""
            Me.txtDireccion.Text = ""

            Exit Sub

        End If
        
        If xEsRuc Then
            Me.txtDni.Text = ""
            'Me.txtDireccion.Text = p.Item("direccion")
            Me.txtDireccion.Text = IIf(IsNull(p.Item("direccion")), "", p.Item("direccion"))
            Me.txtRS.Text = p.Item("razonSocial")
            Me.txtRuc.Text = p.Item("ruc")
        Else
            Me.txtRuc.Text = ""
            Me.txtDireccion.Text = ""
            Me.txtDni.Text = p.Item("dni")
            Me.txtRS.Text = p.Item("nombres") & " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
        End If
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 60, Trim(Me.txtRS.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, Left(Trim(Me.txtRS.Text), 200))
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 50, Left(Trim(Me.txtDireccion.Text), 50))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Left(Trim(Me.txtDireccion.Text), 200))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Me.txtRuc.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, 10, Me.txtRuc.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@salida", adBigInt, adParamOutput, , 0)
        oCmdEjec.Execute
    End If
       
    MousePointer = vbDefault

    Exit Sub

cCruc:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo


End Sub

Sub CargarDocumentos(xCODCIA As String)

    For i = 1 To Me.pbDoc.count - 1
        Unload Me.pbDoc(i)
    Next
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, xCODCIA)
    Set ORStd = oCmdEjec.Execute
    
    Dim vIniLeft, cantTotaDoc, j As Integer

    Dim valor As Double

    j = 1
    vIniLeft = 30
    
    If Not ORStd.EOF Then
        cantTotaDoc = ORStd.RecordCount
        valor = cantTotaDoc / 3
        pos = InStr(Trim(str(valor)), ".")
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)

        If pos = 0 Then
            ent = ""
        Else
            ent = Left(Trim(str(valor)), pos - 1)

        End If
        
        If ent = "" Then: ent = 0
        If pos2 > 0 Then: vPagTotDoc = ent + 1
        
        If cantTotaDoc >= 1 Then: vPagActDoc = 1
        If cantTotaDoc > 3 Then: Me.pbDocSig.Enabled = True
        
        For i = 1 To ORStd.RecordCount
            Load pbDoc(i)
            
            If j = 1 Then
                vIniLeft = vIniLeft + Me.pbDocAnt.Width
            Else
                vIniLeft = vIniLeft + Me.pbDoc(i - 1).Width

            End If
           
            Me.pbDoc(i).Tag = ORStd!Codigo
            Me.pbDoc(i).Left = vIniLeft
            Me.pbDoc(i).Visible = True
            Me.pbDoc(i).Caption = ORStd!Nombre

            'MsgBox "2"
            If j = 3 Then
                j = 1
                vIniLeft = 30
            Else
                j = j + 1

            End If
            
            If i = 1 Then pbDoc_Click (i)

            ORStd.MoveNext
        Next

    End If

End Sub

Sub cargarSeries(cTipoDocto As String)

      If pCodEmp = "" Then Exit Sub

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, pCodEmp)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(cTipoDocto, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNro.Text = CStr(oCmdEjec.Parameters("@MAXIMO").Value)

    ORStd.Filter = "CODIGO='" & pCodTipDoc & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNro.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub DatEmpresas_Change()
sumatoria
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
'If KeyCode = vbKeyF6 Then CargarEmpresas
End Sub

Private Sub Form_Load()
    Set oRSfp = Nothing
    vAcepta = False
    vBuscar = False
    Me.ListView1.Visible = False
    buscars = True
    ConfiguraLV
    ConfigurarLVDetalle


    CargarEmpresas
    'DatEmpresas_Click 1
    
    'CargarDocumentos LK_CODCIA
    'cargarSeries
    
   

    LimpiaParametros oCmdEjec
   
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)

    If xMostrador Then
        oCmdEjec.CommandText = "SpCargarComanda2"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)
    Else
    
        oCmdEjec.CommandText = "SpCargarComanda"
    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fac", adBoolean, adParamInput, , 1)

    Set oRsPago = oCmdEjec.Execute '(, Array(LK_CODCIA, vMesa, LK_FECHA_DIA))


    Do While Not oRsPago.EOF

        Set itemX = Me.lvDetalle.ListItems.Add(, , oRsPago!NumFac)
        itemX.SubItems(1) = oRsPago!CODPLATO
        itemX.SubItems(2) = Trim(oRsPago!plato)
        'itemX.SubItems(3) = Trim(oRsPago!cuenta)
        itemX.SubItems(3) = IIf(IsNull(oRsPago!cuenta), "", Trim(oRsPago!cuenta))
        itemX.SubItems(4) = oRsPago!PRECIO
        itemX.SubItems(5) = oRsPago!CantTotal
        itemX.SubItems(6) = oRsPago!faltante
        itemX.SubItems(7) = FormatNumber(oRsPago!Importe, 2)
        itemX.SubItems(8) = oRsPago!Sec
        itemX.SubItems(9) = oRsPago!aPRO
        itemX.SubItems(10) = oRsPago!aten
        itemX.SubItems(11) = oRsPago!uni
        itemX.SubItems(12) = oRsPago!PED_numsec
        itemX.SubItems(14) = CStr(oRsPago!icbper)
        itemX.SubItems(15) = oRsPago!combo_icbper
                Me.lblpICBPER.Caption = oRsPago!gen_icbper
       
        oRsPago.MoveNext
   
    Loop

  

    Me.lblImporte.Caption = Format(vTOTAL, "########0.#0") 'FormatNumber(vTOTAL, 2)
   
    

   
    Me.labelv.Caption = "0.00"
    Me.chkConsumo.Value = 0
    'Me.lblgratuita.Caption = Format("########0.#0")
    Me.lblServicio.Caption = "0.00"
     Me.lblgratuita.Caption = "0.00"
     Me.lblicbper.Caption = "0.00"

    If LK_CODUSU = "MOZOB" Then
        pCodTipDoc = "01"
        Me.gbTipoDoc.Enabled = False
    End If

    vBuscar = True
    sumatoria
    'RECORDSET PARA LAS FORMAS DE PAGO

    Set oRSfp = New ADODB.Recordset
    oRSfp.CursorType = adOpenDynamic ' setting cursor type
    oRSfp.Fields.Append "idformapago", adBigInt
    oRSfp.Fields.Append "formapago", adVarChar, 120
    oRSfp.Fields.Append "referencia", adVarChar, 100
    oRSfp.Fields.Append "monto", adDouble
    oRSfp.Fields.Append "tipo", adChar, 1
    oRSfp.Fields.Append "pagacon", adDouble
    oRSfp.Fields.Append "vuelto", adDouble
    oRSfp.Fields.Append "diascredito", adInteger
    
    oRSfp.Fields.Refresh
    oRSfp.Open
    
    oRSfp.AddNew
    oRSfp!idformapago = 1
    oRSfp!formapago = "CONTADO"
    oRSfp!referencia = ""
    oRSfp!monto = Me.lblImporte.Caption
    oRSfp!tipo = "E"
    oRSfp!pagacon = 0
    oRSfp!VUELTO = 0
    oRSfp!diascredito = 0
    oRSfp.Update
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_USUARIO_VERIFICACOBRO"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)

    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute

    If Not ORSd.EOF Then
      
        ' Me.cmdFormasPago.Enabled = IIf(ORSd!Fact = "A", True, False)
        Me.cmdCobrar.Enabled = CBool(ORSd!cobra)
       
    End If


End Sub

Private Sub CrearArchivoPlano(cTipoDocto As String, cSerie As String, cNumero As Double)
    Dim oRS As ADODB.Recordset

    LimpiaParametros oCmdEjec

    If cTipoDocto = "F" Then
           oCmdEjec.CommandText = "SP_VENTA_FACTURA_SFS"
    ElseIf cTipoDocto = "B" Then
           oCmdEjec.CommandText = "SP_VENTA_BOLETA_SFS"
    ElseIf LK_CODTRA = 1111 Then
    
    End If
    
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, pCodEmp)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adVarChar, adParamInput, 3, IIf(LK_CODTRA = 1111, PUB_NUMSER_C, cSerie))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adDouble, adParamInput, , IIf(LK_CODTRA = 1111, PUB_NUMFAC_C, cNumero))
    If LK_CODTRA = 1111 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, PUB_FBG)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TRANSACCION", adBigInt, adParamInput, , LK_CODTRA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    End If
    
    Set oRS = oCmdEjec.Execute
    
    Dim sCadena As String

    sCadena = ""
    
    Dim obj_FSO     As Object

    Dim ArchivoCab  As Object
    Dim ArchivoTri As Object
    Dim ArchivoDet  As Object
    Dim ArchivoLey As Object
    Dim ArchivoAca As Object
    
    Dim sARCHIVOcab As String
    Dim sARCHIVOdet As String
    Dim sARCHIVOtri As String
    Dim sARCHIVOley As String
    Dim sARCHIVOaca As String
    
    Dim sRUC        As String
    
    If pCodEmp = "01" Then
        sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "C:\")
    ElseIf pCodEmp = "02" Then
        sRUC = Leer_Ini(App.Path & "\config2.ini", "RUC", "C:\")
    ElseIf pCodEmp = "03" Then
        sRUC = Leer_Ini(App.Path & "\config3.ini", "RUC", "C:\")
    Else
        sRUC = Leer_Ini(App.Path & "\config4.ini", "RUC", "C:\")
    End If
     
    sARCHIVOcab = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".cba", ".cab"))
        
    If LK_CODTRA <> 1111 Then
        sARCHIVOdet = sRUC & "-" & oRS!Nombre + ".det"
        sARCHIVOtri = sRUC & "-" & oRS!Nombre + ".tri" 'IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".tri", ".tri"))
        sARCHIVOley = sRUC & "-" & oRS!Nombre + ".ley" 'IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".ley", ".ley"))
        sARCHIVOaca = sRUC & "-" & oRS!Nombre + ".aca"
        If cTipoDocto = "F" Then 'es factura
            sARCHIVOpag = sRUC & "-" & oRS!Nombre + ".pag"
            sARCHIVOdpa = sRUC & "-" & oRS!Nombre + ".dpa"
            sARCHIVOrtn = sRUC & "-" & oRS!Nombre + ".rtn"
        End If
    End If
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")

    'Creamos un archivo con el método CreateTextFile
    If pCodEmp = "01" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
            Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    ElseIf pCodEmp = "02" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
           Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    ElseIf pCodEmp = "03" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
           Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    ElseIf pCodEmp = "04" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
           Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    End If
    If LK_CODTRA <> 1111 Then
        If pCodEmp = "01" Then
            Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
         ElseIf pCodEmp = "02" Then
            Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
         ElseIf pCodEmp = "03" Then
            Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
         Else
            Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
         End If
    End If
    
    
    If LK_CODTRA = 2412 Then

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!fecemision & "|" & oRS!CODMOTIVO & "|" & oRS!DESCMOTIVO & "|" & oRS!TIPODOCAFECTADO & "|" & oRS!NUMDOCAFECTADO & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!CLI1 & "|" & oRS!TIPMONEDA & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOOPERINAFECTAS & "|" & oRS!MTOOPEREXONERADAS & "|" & oRS!MTOIGV & "|" & oRS!MTOISC & "|" & oRS!MTOOTROSTRIBUTOS & "|" & oRS!MTOIMPVENTA & "|"
            oRS.MoveNext
        Loop
    
    ElseIf LK_CODTRA = 1111 Then
         Do While Not oRS.EOF
            sCadena = sCadena & oRS!FEC_GENERACcION & "|" & oRS!FEC_COMUNICACION & "|" & oRS!TIPDOCBAJA & "|" & oRS!NUMDOCBAJA & "|" & oRS!DESMOTIVOBAJA & "|"
            oRS.MoveNext
        Loop
    Else

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!TIPOPERACION & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!FECHAVENC & "|" & oRS!codlocalemisor & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!rznsocialusuario & "|" & oRS!TIPMONEDA & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!SUMDSCTOGLOBAL & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!TOTANTICIPOS & "|" & oRS!IMPTOTALVENTA & "|" & oRS!UBL & "|" & oRS!CUSTOMDOC & "|"
         oRS.MoveNext
        Loop

    End If
   
    'Escribimos lineas
    ArchivoCab.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoCab.Close
    Set ArchivoCab = Nothing
    
    If LK_CODTRA <> 1111 Then
         'DIRECCION
    oRS.MoveFirst
    sCadena = ""
    Do While Not oRS.EOF
        sCadena = sCadena & oRS!ACA1 & "|" & oRS!ACA2 & "|" & oRS!ACA3 & "|" & oRS!ACA4 & "|" & oRS!ACA5 & "|" & oRS!PAIS & "|" & oRS!UBIGEO & "|" & oRS!dir & "|" & oRS!PAIS1 & "|" & oRS!UBIGEO1 & "|" & oRS!dir1 & "|"
        oRS.MoveNext
    Loop
    
    'Escribimos LINEAS
    ArchivoAca.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoAca.Close
    Set ArchivoAca = Nothing
    Else
    End If
    
   
    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oRS.NextRecordset
   
    sCadena = ""
    Dim c As Integer
    c = 1

    If LK_CODTRA = 2412 Then

        Do While Not oRSdet.EOF
         
            sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!CODOTROITEM & "|" & oRSdet!GRATUITO & "|"
            
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
            
        Loop

    ElseIf LK_CODTRA <> 1111 Then
    

        Do While Not oRSdet.EOF
       
           ' sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTODSCTOITEM & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!TIPAFEIGV & "|" & oRSdet!MTOISCITEM & "|" & oRSdet!TIPSISISC & "|" & oRSdet!MTOPRECIOVENTAITEM & "|" & oRSdet!MTOVALORVENTAITEM & "|"
           sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(oRSdet!DESITEM) & _
           "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & _
           oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & _
           "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & _
           "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!CODIGOICBPER & _
           "|" & oRSdet!IMPORTEICBPER & "|" & oRSdet!CANTIDADICBPER & "|" & oRSdet!TITULOICBPER & "|" & oRSdet!IDEICBPER & "|" & oRSdet!MONTOICBPER & "|" & _
           oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
             
        Loop

    End If

    'Escribimos lineas
    If LK_CODTRA <> 1111 Then
    ArchivoDet.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoDet.Close
    Set ArchivoDet = Nothing
    
    Dim orsTri As ADODB.Recordset
    Set orsTri = oRS.NextRecordset
    
    sCadena = ""
    c = 1
    'ARCIVO .TRI
    Do While Not orsTri.EOF
    sCadena = sCadena & orsTri!Codigo & "|" & orsTri!Nombre & "|" & orsTri!cod & "|" & orsTri!BASEIMPONIBLE & "|" & orsTri!TRIBUTO & "|"
    If c < orsTri.RecordCount Then
        sCadena = sCadena & vbCrLf
    End If
    c = c + 1
        orsTri.MoveNext
    Loop
    
    
     ArchivoTri.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoTri.Close
    Set ArchivoTri = Nothing
    
    Dim orsLey As ADODB.Recordset
    Set orsLey = oRS.NextRecordset
    
    c = 1
    sCadena = ""
    Do While Not orsLey.EOF
        sCadena = sCadena & orsLey!cod & "|" & Trim(CONVER_LETRAS(Me.lblImporte.Caption, "S")) & "|"
        If c < orsLey.RecordCount Then
            sCadena = sCadena & vbCrLf
        End If
        c = c + 1
        orsLey.MoveNext
    Loop
    
    ArchivoLey.WriteLine sCadena
    ArchivoLey.Close
    Set ArchivoLey = Nothing
    
    Dim xFormaPago As String
    If cTipoDocto = "F" Then
            'PAG
            Dim orsPAG As ADODB.Recordset
            Set orsPAG = oRS.NextRecordset
            
            c = 1
            sCadena = ""
            Do While Not orsPAG.EOF
                xFormaPago = orsPAG!formapago
                sCadena = sCadena & orsPAG!formapago & "|" & orsPAG!pendientepago & "|" & orsPAG!TIPMONEDA & "|"
                If c < orsPAG.RecordCount Then
                    sCadena = sCadena & vbCrLf
                End If
                c = c + 1
                orsPAG.MoveNext
            Loop
            
            ArchivoPAG.WriteLine sCadena
            ArchivoPAG.Close
            Set ArchivoPAG = Nothing
            
            'DPA
            Dim orsDPA As ADODB.Recordset
            Set orsDPA = oRS.NextRecordset
            If UCase(xFormaPago) = "CREDITO" Or UCase(xFormaPago) = "CRÉDITO" Then
                Set ArchivoDPA = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdpa, True)
               
                
                c = 1
                sCadena = ""
                Do While Not orsDPA.EOF
                    sCadena = sCadena & orsDPA!cuotapago & "|" & orsDPA!fechavcto & "|" & orsDPA!TIPMONEDA & "|"
                    If c < orsDPA.RecordCount Then
                        sCadena = sCadena & vbCrLf
                    End If
                    c = c + 1
                    orsDPA.MoveNext
                Loop
                
                ArchivoDPA.WriteLine sCadena
                ArchivoDPA.Close
                Set ArchivoDPA = Nothing
            End If
             'RTN
'            Dim orsRTN As ADODB.Recordset
'            Set orsRTN = oRS.NextRecordset
'
'            c = 1
'            sCadena = ""
'            Do While Not orsRTN.EOF
'                sCadena = sCadena & orsRTN!impoperacion & "|" & orsRTN!porretencion & "|" & orsRTN!impretencion & "|"
'                If c < orsRTN.RecordCount Then
'                    sCadena = sCadena & vbCrLf
'                End If
'                c = c + 1
'                orsRTN.MoveNext
'            Loop
'
'            ArchivoRTN.WriteLine sCadena
'            ArchivoRTN.Close
'            Set ArchivoRTN = Nothing
        End If
    
    End If
    
   
    
    Set obj_FSO = Nothing
    
End Sub

Private Sub CalcularImporte()
'1combo
'2 combo
'3 combo

Dim vp1  As Currency, vp2 As Currency, vp3 As Currency

Dim vDol As Currency

'If Me.cboMoneda1.ListIndex = 0 Then 'soles
'    vp1 = val(Me.txtMoney1.Text)
'ElseIf Me.cboMoneda1.ListIndex = 1 Then 'dolares
'    vp1 = LK_TIPO_CAMBIO * val(Me.txtMoney1.Text)
'    'vp1 = vDol - val(Me.lblImporte.Caption)
'End If

Dim Item As Object
Dim icbper As Double
vp1 = 0
icbper = 0

For Each Item In Me.lvDetalle.ListItems

    vp1 = vp1 + Item.SubItems(7)
    If Item.SubItems(14) = 1 Then
        icbper = icbper + (Item.SubItems(6) * Me.lblpICBPER.Caption)
    End If
    
    If Item.SubItems(15) > 0 Then
         icbper = icbper + Item.SubItems(15)
    End If
    
Next

'If Me.cboMoneda1.ListIndex = 1 Then 'dolares
If LK_MONEDA = "D" Then
    vp1 = LK_TIPO_CAMBIO * vp1
'vp1 = vDol - val(Me.lblImporte.Caption)
End If

vp1 = vp1 - val(Me.lblDscto.Caption)
Me.lblicbper.Caption = FormatNumber(icbper, 2)

'If Me.txtMoney1.Text <> 0 Then
    Me.lblImporte.Caption = vp1
    Me.labelv.Caption = val(Me.lblImporte.Caption) - vp1
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
gDESCUENTO = 0
gPAGO = 0

End Sub


Private Sub lvDetalle_DblClick()
frmFaccomandaOtroPlato.Show vbModal
End Sub

Private Sub lvDetalle_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If Not Me.lvDetalle.SelectedItem Is Nothing Then
            Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.index
            CalcularImporte
            sumatoria
      
           ' Me.txtMoney1.Text = Me.lblImporte.Caption
            
            If oRSfp.RecordCount <> 0 Then
                oRSfp.MoveFirst
                
                Do While Not oRSfp.EOF
                    oRSfp.Delete
                    oRSfp.MoveNext
                Loop
                
            End If
            
            oRSfp.AddNew
            oRSfp!idformapago = 1
            oRSfp!formapago = "CONTADO"
            oRSfp!referencia = ""
            oRSfp!tipo = "E"
            oRSfp!monto = Me.lblImporte.Caption
            oRSfp!diascredito = 0
            oRSfp.Update
      
        End If
    End If

End Sub

Private Sub sumatoria()
Dim vIgv As Integer
 LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "USP_EMPRESA_IGV"
            Dim orsIGV As ADODB.Recordset
            Set orsIGV = oCmdEjec.Execute(, pCodEmp)

            Dim Item As Object
        
            If Not orsIGV.EOF Then
            vIgv = orsIGV.Fields(0).Value
            Me.lblporcigv.Caption = vIgv & "%"
            End If

    Dim vimp As Double

    For i = 1 To Me.lvDetalle.ListItems.count
        vimp = vimp + Me.lvDetalle.ListItems(i).SubItems(7)
       
    Next

    'Me.lblImporte.Caption = Format(vimp, "########0.#0") 'FormatNumber(vimp, 2)
     Me.lblServicio.Caption = "0.00"
     Me.lblgratuita.Caption = "0.00"
     Me.lblvvta.Caption = Round(vimp / ((vIgv / 100) + 1), 2)
     Me.lblIgv.Caption = vimp - Me.lblvvta.Caption
     Me.lblImporte.Caption = Format(val(vimp) + val(Me.lblicbper.Caption), "########0.#0")
        

    
End Sub

Private Sub lvDetalle_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        frmAsigCantFac.txtCantidad.Text = Me.lvDetalle.SelectedItem.SubItems(5)
        frmAsigCantFac.Show vbModal

        If frmAsigCantFac.vAcepta Then
            If frmAsigCantFac.vCANTIDAD > Me.lvDetalle.SelectedItem.SubItems(5) Then
                MsgBox "La cantidad supera la actual", vbInformation, "Error"

                Exit Sub

            End If

            Me.lvDetalle.SelectedItem.SubItems(6) = frmAsigCantFac.vCANTIDAD
            Me.lvDetalle.SelectedItem.SubItems(7) = Me.lvDetalle.SelectedItem.SubItems(6) * Me.lvDetalle.SelectedItem.SubItems(4)
            CalcularImporte
            sumatoria
      
            If oRSfp.RecordCount <> 0 Then
                oRSfp.MoveFirst
                
                Do While Not oRSfp.EOF
                    oRSfp.Delete
                    oRSfp.MoveNext
                Loop
                
            End If
            
    oRSfp.AddNew
    oRSfp!idformapago = 1
    oRSfp!formapago = "CONTADO"
    oRSfp!referencia = ""
    oRSfp!monto = Me.lblImporte.Caption
    oRSfp!tipo = "E"
    oRSfp.Update
            
        End If

    End If

End Sub

Private Sub pbAumentar_Click()
If Me.txtcopias.Text = 99 Then Exit Sub
Me.txtcopias.Text = val(Me.txtcopias.Text) + 1
End Sub

Private Sub pbDisminuir_Click()
If Me.txtcopias.Text = 1 Then Exit Sub
Me.txtcopias.Text = val(Me.txtcopias.Text) - 1
End Sub

Private Sub pbDoc_Click(index As Integer)
pCodTipDoc = Me.pbDoc(index).Tag
pDesTipDoc = Me.pbDoc(index).Caption
cargarSeries pDesTipDoc

 For i = 1 To Me.pbDoc.count - 1

        If index = i Then
            Me.pbDoc(i).Checked = True
        Else
            Me.pbDoc(i).Checked = False

        End If

    Next
End Sub

Private Sub pbDocAnt_Click()
Dim ini, fin, f, FF As Integer
If vPagActDoc = 2 Then
    ini = 1
    fin = ini * 3
ElseIf vPagActDoc = 1 Then
    Exit Sub
Else
    FF = vPagActDoc - 1
    ini = (3 * FF) - 2
    fin = 3 * FF
End If

For f = ini To fin
    Me.pbDoc(f).Visible = True
Next
If vPagActDoc > 1 Then
    vPagActDoc = vPagActDoc - 1
    If vPagActDoc = 1 Then: Me.pbDocAnt.Enabled = False
    
    Me.pbDocSig.Enabled = True
End If
End Sub

Private Sub pbDocSig_Click()
Dim ini, fin, f As Integer
If vPagActDoc = 1 Then
    ini = 1
    fin = ini * 3
ElseIf vPagActDoc = vPagTotDoc Then
    Exit Sub
Else
    ini = (3 * vPagActDoc) - 2
    fin = 3 * vPagActDoc
End If

For f = ini To fin
    Me.pbDoc(f).Visible = False
Next
If vPagActDoc < vPagTotDoc Then
    vPagActDoc = vPagActDoc + 1
    If vPagActDoc = vPagTotDoc Then: Me.pbDocSig.Enabled = False
    
    Me.pbDocAnt.Enabled = True
End If
End Sub

Private Sub pbEmpresa_Click(index As Integer)

pCodTipDoc = ""
pCodEmp = Me.pbEmpresa(index).Tag
CargarDocumentos Me.pbEmpresa(index).Tag


 For i = 1 To Me.pbEmpresa.count - 1

        If index = i Then
            Me.pbEmpresa(i).Checked = True
        Else
            Me.pbEmpresa(i).Checked = False

        End If

    Next
End Sub

Private Sub pbEmpresaAnt_Click()
Dim ini, fin, f, FF As Integer
If vPagActEmp = 2 Then
    ini = 1
    fin = ini * 5
ElseIf vPagActEmp = 1 Then
    Exit Sub
Else
    FF = vPagActEmp - 1
    ini = (5 * FF) - 4
    fin = 5 * FF
End If

For f = ini To fin
    Me.pbEmpresa(f).Visible = True
Next
If vPagActEmp > 1 Then
    vPagActEmp = vPagActEmp - 1
    If vPagActEmp = 1 Then: Me.pbEmpresaAnt.Enabled = False
    
    Me.pbEmpresaSig.Enabled = True
End If
End Sub

Private Sub pbEmpresaSig_Click()
Dim ini, fin, f As Integer
If vPagActFam = 1 Then
    ini = 1
    fin = ini * 5
ElseIf vPagActEmp = vPagTotEmp Then
    Exit Sub
Else
    ini = (5 * vPagActEmp) - 4
    fin = 5 * vPagActEmp
End If

For f = ini To fin
    Me.pbEmpresa(f).Visible = False
Next
If vPagActEmp < vPagTotEmp Then
    vPagActEmp = vPagActEmp + 1
    If vPagActEmp = vPagTotEmp Then: Me.pbEmpresaSig.Enabled = False
    
    Me.pbEmpresaAnt.Enabled = True
End If
End Sub

Private Sub txtCopias_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtMoney1_Change()
CalcularImporte
End Sub

Private Sub txtMoney2_Change()
CalcularImporte
End Sub

Private Sub txtMoney3_Change()
CalcularImporte
End Sub

Private Sub txtnro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtRS.SetFocus
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtRS_Change()
vBuscar = True
End Sub

Private Sub txtRS_GotFocus()
buscars = True
End Sub

Private Sub txtRS_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo SALE

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then 'flecha arriba
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then    'avanzar pagina
        loc_key = loc_key + 17

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then        'regresar pagina
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then        'tecla escape
        Me.ListView1.Visible = False
        Me.txtRS.Text = ""
        Me.txtRuc.Text = ""
        Me.txtDireccion.Text = ""
    End If

    GoTo fin
posicion:
If Me.ListView1.ListItems.count = 0 Then Exit Sub
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    'txtRS.SelStart = Len(txtRS.Text)

fin:

    Exit Sub

SALE:
MsgBox Err.Description
End Sub

Private Sub txtRS_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            If pCodEmp = "" Then
                Me.ListView1.ListItems.Clear
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SpListarCliProv"
                Set oRsPago = oCmdEjec.Execute(, Array(pCodEmp, "C", Me.txtRS.Text))

                Dim Item As Object
        
                If Not oRsPago.EOF Then

                    Do While Not oRsPago.EOF
                        Set Item = Me.ListView1.ListItems.Add(, , oRsPago!CodClie)
                        Item.SubItems(1) = Trim(oRsPago!Nombre)
                        Item.SubItems(2) = IIf(IsNull(oRsPago!RUC), "", oRsPago!RUC)
                        Item.SubItems(3) = Trim(oRsPago!dir)
                        Item.Tag = oRsPago!DNI
                        oRsPago.MoveNext
                    Loop

                    Me.ListView1.Visible = True
                    Me.ListView1.ListItems(1).Selected = True
                    loc_key = 1
                    Me.ListView1.ListItems(1).EnsureVisible
                    vBuscar = False
                Else

                    If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                        frmCLI.Show vbModal

                    End If

                End If

            Else
                MsgBox "Debe elegir la Empresa para continuar con la facturación.", vbInformation, Pub_Titulo

            End If
        
        Else
            
            Me.txtRuc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
            Me.txtDireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
            Me.txtRS.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
            Me.ListView1.Visible = False
            Me.txtDni.Text = Me.ListView1.ListItems(loc_key).Tag
            Me.txtRuc.Tag = Me.ListView1.ListItems(loc_key)
            Me.lvDetalle.SetFocus

        End If

    End If

End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'Me.ListView1.Visible = True
If Len(Trim(Me.txtRuc.Text)) <> 0 Then
For i = 1 To Me.ListView1.ListItems.count
    If Trim(Me.txtRuc.Text) = Trim(Me.ListView1.ListItems(i).SubItems(2)) Then
        'Me.ListView1.ListItems(i).Selected = True
        buscars = False
        loc_key = i
        Me.ListView1.ListItems(i).EnsureVisible
        Me.txtRS.Text = Me.ListView1.ListItems(i).SubItems(1)
        Me.txtDireccion.Text = Me.ListView1.ListItems(i).SubItems(3)
        Me.txtRuc.Tag = Me.ListView1.ListItems(i)
        Exit For
    Else
       ' Me.ListView1.ListItems(i).Selected = False
        loc_key = -1
        
    End If
Next
Else
Me.ListView1.Visible = False
Me.txtRuc.Text = ""
Me.txtDireccion.Text = ""
Me.txtRS.Text = ""
End If
End If
End Sub

Private Sub CargarEmpresas()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CIAS_FACTURACION"
   ' oCmdEjec.Prepared = True
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
    
    Dim ORSe As ADODB.Recordset
    Set ORSe = oCmdEjec.Execute
    
    Dim vIniLeft, cantTotaEmp, j As Integer
    Dim valor As Double
    j = 1
    vIniLeft = 30
    
    If Not ORSe.EOF Then
        cantTotaEmp = ORSe.RecordCount
        valor = cantTotaEmp / 5
        pos = InStr(Trim(str(valor)), ".")
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
        ent = Left(Trim(str(valor)), pos - 1)
        
        If ent = "" Then: ent = 0
        If pos2 > 0 Then: vPagTotEmp = ent + 1
        
        If cantTotaEmp >= 1 Then: vPagActEmp = 1
        If cantTotaEmp > 5 Then: Me.pbEmpresaSig.Enabled = True
        
        For i = 1 To ORSe.RecordCount
            Load pbEmpresa(i)
            
            If j = 1 Then
                vIniLeft = vIniLeft + Me.pbEmpresaAnt.Width
            Else
                vIniLeft = vIniLeft + Me.pbEmpresa(i - 1).Width
            End If
           
            Me.pbEmpresa(i).Tag = ORSe!CodCia
            Me.pbEmpresa(i).Left = vIniLeft
            Me.pbEmpresa(i).Visible = True
            Me.pbEmpresa(i).Caption = ORSe!PAR_NOMBRE
            'MsgBox "2"
             If j = 5 Then
            j = 1
            vIniLeft = 30
            Else
            j = j + 1
            End If
            
            If i = 1 Then pbEmpresa_Click (i)

            ORSe.MoveNext
        Next
    End If
    
End Sub
