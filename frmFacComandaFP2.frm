VERSION 5.00
Object = "{BB35AEF3-E525-4F8B-81F2-511FF805ABB1}#2.1#0"; "ScrollerII.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form frmFacComandaFP2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacComandaFP2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   840
      Left            =   9960
      TabIndex        =   24
      Top             =   6360
      Width           =   1935
      _Version        =   786433
      _ExtentX        =   3413
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "&Aceptar"
      Appearance      =   5
      DrawFocusRect   =   0   'False
      Picture         =   "frmFacComandaFP2.frx":1CCA
   End
   Begin XtremeSuiteControls.GroupBox GroFpListado 
      Height          =   4455
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   9255
      _Version        =   786433
      _ExtentX        =   16325
      _ExtentY        =   7858
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.PushButton pbQuitar 
         Height          =   855
         Left            =   8160
         TabIndex        =   4
         Top             =   600
         Width           =   975
         _Version        =   786433
         _ExtentX        =   1720
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "&Quitar"
         Appearance      =   4
         DrawFocusRect   =   0   'False
         Picture         =   "frmFacComandaFP2.frx":39A4
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.ListView lvFP 
         Height          =   4095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7935
         _Version        =   786433
         _ExtentX        =   13996
         _ExtentY        =   7223
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
   End
   Begin ScrollerII.ScrollableContainer ScrollableContainer1 
      Height          =   4365
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7699
      SmallChange     =   1
      ScaleMode       =   0
      BorderStyle     =   3
      Begin XtremeSuiteControls.PushButton cmdFP 
         Height          =   960
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   3855
         _Version        =   786433
         _ExtentX        =   6800
         _ExtentY        =   1693
         _StockProps     =   79
         Caption         =   "PushButton1"
         Appearance      =   5
         DrawFocusRect   =   0   'False
         Picture         =   "frmFacComandaFP2.frx":411E
      End
   End
   Begin XtremeSuiteControls.GroupBox GroTotales 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   13815
      _Version        =   786433
      _ExtentX        =   24368
      _ExtentY        =   3201
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.FlatEdit txtPagaCon 
         Height          =   435
         Left            =   6480
         TabIndex        =   20
         Top             =   1200
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
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
         Alignment       =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Contado:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   23
         Top             =   322
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movilidad:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1185
         TabIndex        =   22
         Top             =   802
         Width           =   1260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICBPER:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   21
         Top             =   1282
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tarjeta:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4725
         TabIndex        =   19
         Top             =   322
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4680
         TabIndex        =   18
         Top             =   802
         Width           =   1755
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paga con:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5220
         TabIndex        =   17
         Top             =   1282
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Credito:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8760
         TabIndex        =   16
         Top             =   322
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pagos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8910
         TabIndex        =   15
         Top             =   802
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vuelto:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9570
         TabIndex        =   14
         Top             =   1282
         Width           =   900
      End
      Begin XtremeSuiteControls.Label lblVuelto 
         Height          =   435
         Left            =   10560
         TabIndex        =   13
         Top             =   1200
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblTotalPagos 
         Height          =   435
         Left            =   10560
         TabIndex        =   12
         Top             =   720
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblTC 
         Height          =   435
         Left            =   10560
         TabIndex        =   11
         Top             =   240
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblTotalPagar 
         Height          =   435
         Left            =   6480
         TabIndex        =   10
         Top             =   720
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblTT 
         Height          =   435
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblicbper 
         Height          =   435
         Left            =   2520
         TabIndex        =   8
         Top             =   1200
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblMovilidad 
         Height          =   435
         Left            =   2520
         TabIndex        =   7
         Top             =   720
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Begin XtremeSuiteControls.Label lblTE 
         Height          =   435
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   1755
         _Version        =   786433
         _ExtentX        =   3096
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "0.00"
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
      Height          =   840
      Left            =   12000
      TabIndex        =   25
      Top             =   6360
      Width           =   1935
      _Version        =   786433
      _ExtentX        =   3413
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "&Cancelar"
      Appearance      =   5
      DrawFocusRect   =   0   'False
      Picture         =   "frmFacComandaFP2.frx":5DF8
   End
End
Attribute VB_Name = "frmFacComandaFP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gDELIVERY As Boolean


Private Sub cmdAceptar_Click()

    If Me.lvFP.ListItems.count = 0 Then
        MsgBox "Debe ingresar las formas de pago.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    If CDec(Me.lblTotalPagar.Caption) <> CDec(Me.lblTotalPagos.Caption) Then
        MsgBox "Importe no coincide", vbInformation, Pub_Titulo

        Exit Sub

    End If

    If Not oRSfp Is Nothing Then

        'If Not oRSfp.EOF Then oRSfp.Delete
        If oRSfp.RecordCount <> 0 Then
            oRSfp.MoveFirst

            Do While Not oRSfp.EOF
                oRSfp.Delete adAffectCurrent
                oRSfp.MoveNext
            Loop

        End If
    End If

    Dim DD As Object

    If gDELIVERY Then

        Dim sSERIE As String, sNUMERO As Double

        Dim xEXITO As String

        sSERIE = frmDeliveryApp.lblSerie.Caption
        sNUMERO = frmDeliveryApp.lblNumero.Caption
            
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_DELIVERY_ASIGNAR_FORMAPAGO"
        oCmdEjec.CommandType = adCmdStoredProc
            
        xEXITO = ""
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, sSERIE)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , sNUMERO)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PAGO", adDouble, adParamInput, , IIf(Len(Trim(Me.txtPagaCon.Text)) = 0, 0, Me.txtPagaCon.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCUENTO", adDouble, adParamInput, , 0) 'Me.txtDescuento.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VUELTO", adDouble, adParamInput, , Me.lblVuelto.Caption)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 300, xEXITO)
        oCmdEjec.Execute
        frmFacComanda.gDESCUENTO = 0 ' Me.txtDescuento.Text
    
        xEXITO = oCmdEjec.Parameters("@EXITO").Value

        If Len(Trim(xEXITO)) = 0 Then
            'GRABA EN TABLA
            'borra LOS APGOS
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_DELIVERY_ELIMINARFORMASPAGO"
            oCmdEjec.Execute , Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption)
       
            oCmdEjec.CommandText = "SP_DELIVERY_REGISTRAFORMASPAGO"
            
            For Each DD In Me.lvFP.ListItems

                oRSfp.AddNew
                oRSfp!idformapago = DD.Tag
                oRSfp!formapago = DD.SubItems(1) ' Me.lblFormaPago(DD).Caption
                oRSfp!referencia = DD.SubItems(2) 'IIf(Me.cbomoneda(DD).ListIndex = 0, "S", "D")
                oRSfp!monto = DD.SubItems(3)
                oRSfp!tipo = DD.SubItems(4)
                oRSfp!pagacon = IIf(Len(Trim(Me.txtPagaCon.Text)) = 0, 0, Me.txtPagaCon.Text)
                oRSfp!VUELTO = Me.lblVuelto.Caption
                'oRSfp!diascredito = DD.SubItems(5)
                
                oRSfp.Update
            
                If oRSfp.RecordCount <> 0 Then

                    Do While Not oRSfp.EOF
                        LimpiaParametros oCmdEjec
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, sSERIE)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , sNUMERO)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFORMAPAGO", adInteger, adParamInput, , oRSfp!idformapago) ' Me.lblFormaPago(i).Tag) ' Me.lblFormaPago(oCTRL.Tag).Caption)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , oRSfp!monto) ' Me.txtImporte(i).Text) '  Me.txtImporte(oCTRL.Tag).Text)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PAGACON", adDouble, adParamInput, , oRSfp!pagacon)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VUELTO", adDouble, adParamInput, , oRSfp!VUELTO)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, oRSfp!tipo)
                        oCmdEjec.Execute
                        oRSfp.MoveNext
                    Loop

                End If

            Next

            '            For i = 1 To Me.txtImporte.count - 1
            '
            '                If Len(Trim(Me.txtImporte(i).Text)) <> 0 And IsNumeric(Me.txtImporte(i).Text) Then
            '                    LimpiaParametros oCmdEjec
            '                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            '                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, sSERIE)
            '                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , sNUMERO)
            '                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFORMAPAGO", adInteger, adParamInput, , Me.lblFormaPago(i).Tag) ' Me.lblFormaPago(oCTRL.Tag).Caption)
            '                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , Me.txtImporte(i).Text) '  Me.txtImporte(oCTRL.Tag).Text)
            '                    oCmdEjec.Execute
            '                End If
            '
            '            Next

        Else
            MsgBox xEXITO, vbCritical, Pub_Titulo
        End If

    Else
        '   vAcepta = True

        For Each DD In Me.lvFP.ListItems

            oRSfp.AddNew
            oRSfp!idformapago = DD.Tag
            oRSfp!formapago = DD.SubItems(1) ' Me.lblFormaPago(DD).Caption
            oRSfp!referencia = DD.SubItems(2) 'IIf(Me.cbomoneda(DD).ListIndex = 0, "S", "D")
            oRSfp!monto = DD.SubItems(3)
            oRSfp!tipo = DD.SubItems(4)
            oRSfp!pagacon = IIf(Len(Trim(Me.txtPagaCon.Text)) = 0, 0, Me.txtPagaCon.Text)
            oRSfp!VUELTO = Me.lblVuelto.Caption
            oRSfp!diascredito = IIf(DD.SubItems(5) = "", 0, DD.SubItems(5))
            oRSfp.Update
        Next

    End If
           
    Unload Me

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdFP_Click(index As Integer)

    '    frmFacComandaFP3.MuestraReferencia = True
    '
    '    frmFacComandaFP3.Show vbModal
    '
    '    If frmFacComandaFP3.gACEPTA Then
   

    Dim strDato As String

    '
    If (CDec(Me.lblTotalPagar.Caption) - CDec(Me.lblTotalPagos.Caption)) <> 0 Then

        Dim itemX As Object
        
         esCredito = Split(Me.cmdFP(index).Tag, "|")(2)
        If esCredito Then
            frmFacComandaFP2credito.Show vbModal
            If frmFacComandaFP2credito.gAcepta Then
                strDato = frmFacComandaFP2credito.gDiasCredito
            Else
                Exit Sub
            End If
           
        End If

        Set itemX = Me.lvFP.ListItems.Add(, , Me.lvFP.ListItems.count + 1)
        itemX.SubItems(1) = Me.cmdFP(index).Caption
        itemX.Tag = Split(Me.cmdFP(index).Tag, "|")(0)
        itemX.SubItems(2) = Trim(frmFacComandaFP3.gREF)
        itemX.SubItems(3) = CStr(CDec(Me.lblTotalPagar.Caption) - CDec(Me.lblTotalPagos.Caption))
        itemX.SubItems(4) = Split(Me.cmdFP(index).Tag, "|")(1)
        itemX.SubItems(5) = strDato
       

        CalcularTotales
    End If
        
    '    End If

End Sub

Private Sub CalcularTotales()

    Dim itemX As Object

    Dim TE, tt, tC As Double, TP As Double

    TE = 0
    tt = 0
    tC = 0

    For Each itemX In Me.lvFP.ListItems

        If itemX.SubItems(4) = "E" Then
            TE = CDec(TE) + CDec(itemX.SubItems(3))
        ElseIf itemX.SubItems(4) = "T" Then
            tt = CDec(tt) + CDec(itemX.SubItems(3))
        ElseIf itemX.SubItems(4) = "C" Then
            tC = CDec(tC) + CDec(itemX.SubItems(3))
        End If

    Next

    Me.lblTE.Caption = FormatCurrency(TE, 2)
    Me.lblTT.Caption = FormatCurrency(tt, 2)
    Me.lblTC.Caption = FormatCurrency(tC, 2)
        
    TP = TE + tt + tC
    Me.lblTotalPagos.Caption = FormatCurrency(TP, 2)

    If gDELIVERY Then
        Me.lblTotalPagar.Caption = FormatCurrency(CDec(CDec(frmDeliveryApp.lblTot.Caption) + IIf(pINCMOV, CDec(Me.lblMovilidad.Caption), 0)) + val(frmDeliveryApp.lblicbper.Caption), 2)
    Else
        Me.lblTotalPagar.Caption = FormatCurrency(CDec(CDec(frmFacComanda.lblImporte.Caption) + IIf(pINCMOV, CDec(Me.lblMovilidad.Caption), 0)) + val(frmDeliveryApp.lblicbper.Caption), 2)
    End If

    If TE <> 0 Then
        If IsNumeric(Me.txtPagaCon.Text) Then
            If CDec(IIf(Len(Trim(Me.txtPagaCon.Text)) = 0, 0, Me.txtPagaCon.Text)) > 0 Then
                If val(Me.txtPagaCon.Text) > TE Then
        
                    Me.lblVuelto.Caption = FormatCurrency(CDec(Me.txtPagaCon.Text) - TE, 2)
                Else
                    Me.lblVuelto.Caption = FormatCurrency(CDec(0))
                End If

            Else
                Me.lblVuelto.Caption = FormatCurrency(CDec(0))
            End If

        Else
            Me.lblVuelto.Caption = FormatCurrency(CDec(0))
        End If

    Else
        Me.lblVuelto.Caption = FormatCurrency(CDec(0))
    End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    Dim vTop  As Integer

    Dim vLeft As Integer

    vTop = 120
    vLeft = 120
    
   
    
    Me.lblicbper.Caption = frmDeliveryApp.lblicbper.Caption
    
     LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_INCLUYEMOVILIDAD_GET"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IM", adBoolean, adParamOutput, , 0)
    
    oCmdEjec.Execute
    
    pINCMOV = oCmdEjec.Parameters("@IM").Value
    
    
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpCargarFormasPago"
  
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodTran", adBigInt, adParamInput, 2, 2401)
        
    Set oRStp = oCmdEjec.Execute

    Dim f As Integer

    For f = 1 To oRStp.RecordCount
        Load Me.cmdFP(f)
      
        Me.cmdFP(f).Top = vTop
        Me.cmdFP(f).Left = vLeft
    
        vTop = vTop + Me.cmdFP(f).Height + 50
        
        Me.cmdFP(f).Visible = True
        Me.cmdFP(f).Caption = oRStp!formapago
        Me.cmdFP(f).Tag = oRStp!Codigo & "|" & oRStp!tipo & "|" & oRStp!CRE
     
        oRStp.MoveNext
    Next

    ConfiguraLV
    
     If Not oRSfp Is Nothing Then

            'If Not oRSfp.EOF Then oRSfp.Delete
            If oRSfp.RecordCount <> 0 Then
                oRSfp.MoveFirst

                Do While Not oRSfp.EOF
                    oRSfp.Delete adAffectCurrent
                    oRSfp.MoveNext
                Loop

            End If
        End If

    'CARGANDO LAS FORMAS DE PAGO
    If gDELIVERY Then

 'DATOS
        oCmdEjec.CommandText = "SP_DELIVERY_CARGATARIFAPEDIDO"
        LimpiaParametros oCmdEjec

        Dim orsT As ADODB.Recordset

        Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption))
      
        If Not orsT.EOF Then
        
            'cMontoTarifa = orsT!tarifa
            Me.lblMovilidad.Caption = FormatCurrency(orsT!tarifa, 2)
            'Me.txtDescuento.Text = orsT!descuento
            'Me.txtPago.Text = IIf(IsNull(orsT!PAGO), 0, orsT!PAGO)
            'Me.lblVuelto.Caption = orsT!VUELTO
            
        End If
        
        Dim ORSpr As ADODB.Recordset
        
       

        oCmdEjec.CommandText = "SP_DELIVERY_CARGAFORMASPAGO"
        LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, frmDeliveryApp.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , frmDeliveryApp.lblNumero.Caption)
                
        Set ORSpr = oCmdEjec.Execute
        
        If ORSpr.RecordCount = 0 Then
            oRSfp.AddNew
            oRSfp!idformapago = 1
            oRSfp!formapago = "CONTADO"
            oRSfp!referencia = ""
            oRSfp!monto = CDec(frmDeliveryApp.lblTot.Caption) + IIf(pINCMOV, CDec(Me.lblMovilidad.Caption), 0) + frmDeliveryApp.lblicbper.Caption
            oRSfp!tipo = "E"
            oRSfp!pagacon = 0
            oRSfp!VUELTO = 0
            oRSfp!diascredito = 0
            oRSfp.Update
        Else
        
            Do While Not ORSpr.EOF
                oRSfp.AddNew
                oRSfp!idformapago = ORSpr!idfp
                oRSfp!formapago = Trim(ORSpr!fp) ' Me.lblFormaPago(DD).Caption
                oRSfp!referencia = "" 'DD.SubItems(2) 'IIf(Me.cbomoneda(DD).ListIndex = 0, "S", "D")
                oRSfp!monto = ORSpr!monto + val(frmDeliveryApp.lblicbper.Caption)
                oRSfp!tipo = ORSpr!tipo ' "E" 'DD.SubItems(4)
                oRSfp!pagacon = ORSpr!pagacon + val(frmDeliveryApp.lblicbper.Caption) 'IIf(Len(Trim(Me.txtPagaCon.Text)) = 0, 0, Me.txtPagaCon.Text)
                oRSfp!VUELTO = ORSpr!VUELTO ' Me.lblVuelto.Caption
                oRSfp!diascredito = 0
                oRSfp.Update
                ORSpr.MoveNext
            Loop

        End If

       

    End If
    
    If Not oRSfp Is Nothing Then

        Dim itemX As Object

        If Not oRSfp.EOF Then oRSfp.MoveFirst
        If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
        
        Me.txtPagaCon.Text = 0 'oRSfp!pagacon
        Me.lblVuelto.Caption = 0 ' oRSfp!VUELTO
        
        
         oRSfp.AddNew
            oRSfp!idformapago = 1
            oRSfp!formapago = "CONTADO"
            oRSfp!referencia = ""
            If gDELIVERY Then
            oRSfp!monto = CDec(frmDeliveryApp.lblTot.Caption) + frmDeliveryApp.lblicbper.Caption
            Else
            oRSfp!monto = CDec(frmComanda.lblTot.Caption)
            End If
            'oRSfp!monto = IIf(gDELIVERY, CDec(frmDeliveryApp.lblTot.Caption) + frmDeliveryApp.lblicbper.Caption, CDec(frmComanda.lblTot.Caption))    ' + IIf(pINCMOV, CDec(Me.lblMovilidad.Caption), 0) + frmDeliveryApp.lblicbper.Caption
            oRSfp!tipo = "E"
            oRSfp!pagacon = 0
            oRSfp!VUELTO = 0
            oRSfp!diascredito = 0
            oRSfp.Update
     
        Do While Not oRSfp.EOF
           
            Set itemX = Me.lvFP.ListItems.Add(, , Me.lvFP.ListItems.count + 1)
            itemX.SubItems(1) = oRSfp!formapago
            itemX.SubItems(2) = oRSfp!referencia
            itemX.SubItems(3) = oRSfp!monto
            itemX.SubItems(4) = oRSfp!tipo
            itemX.SubItems(5) = oRSfp!diascredito
            itemX.Tag = oRSfp!idformapago
            'MsgBox "demo"
           
            '            If TypeOf oCTRL Is label And oCTRL.Tag <> "" Then
            '
            '                If CInt(oCTRL.Tag) = oRSfp!idformapago Then
            '                    Me.txtImporte(oCTRL.Tag).Text = oRSfp!monto
            '                    xpago = xpago + oRSfp!monto
            '
            '                    Exit For
            '
            '                End If
            '            End If

            oRSfp.MoveNext
        Loop

    End If

    CalcularTotales
    
End Sub

'Private Sub LlenarFP()
'
'    If oRSfp Is Not Nothing Then
'        If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
'
'        Dim itemFP As Object
'
'        Do While Not oRSfp.EOF
'            'set itemFP=me.lvFP.ListItems.Add(,,oRSfp!
'            oRSfp.MoveNext
'        Loop
'
'    End If
'
'End Sub

Private Sub ConfiguraLV()

    With Me.lvFP
        .HideColumnHeaders = False
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = xtpListViewLabelManual
        .View = xtpListViewReport
        .ColumnHeaders.Add , , "Nro", 700
        .ColumnHeaders.Add , , "Forma Pago", 3000
        .ColumnHeaders.Add , , "Nro Referencia", 2100
        .ColumnHeaders.Add , , "Monto", 1400
        .ColumnHeaders.Add , , "Tipo", 0
        .ColumnHeaders.Add , , "FechaVenc", 0
        .MultiSelect = False
        
    End With

End Sub

Private Sub lvFP_DblClick()

    If Me.lvFP.SelectedItem.SubItems(4) = "T" Then
        frmFacComandaFP3.MuestraReferencia = True
    Else
        frmFacComandaFP3.MuestraReferencia = False
    End If

    frmFacComandaFP3.txtMonto.Text = Me.lvFP.SelectedItem.SubItems(3)
    frmFacComandaFP3.txtReferencia.Text = Me.lvFP.SelectedItem.SubItems(2)
    frmFacComandaFP3.Show vbModal

    If frmFacComandaFP3.gAcepta Then
        Me.lvFP.SelectedItem.SubItems(3) = frmFacComandaFP3.gMONTO
        Me.lvFP.SelectedItem.SubItems(2) = frmFacComandaFP3.gREF
        CalcularTotales
    End If

End Sub


Private Sub pbQuitar_Click()
If Me.lvFP.ListItems.count = 0 Then Exit Sub
If Me.lvFP.SelectedItem Is Nothing Then Exit Sub
 Me.lvFP.ListItems.Remove Me.lvFP.SelectedItem.index
 CalcularTotales
End Sub

Private Sub txtPagaCon_Change()
CalcularTotales
End Sub
