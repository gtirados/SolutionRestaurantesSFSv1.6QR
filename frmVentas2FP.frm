VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{BB35AEF3-E525-4F8B-81F2-511FF805ABB1}#2.1#0"; "scrollerii.ocx"
Begin VB.Form frmVentas2FP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14055
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
   ScaleHeight     =   7275
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12240
      TabIndex        =   16
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10440
      TabIndex        =   15
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton Command1 
         Caption         =   "Quitar"
         Height          =   360
         Left            =   8040
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvFP 
         Height          =   4095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin ScrollerII.ScrollableContainer ScrollableContainer1 
      Height          =   4360
      Left            =   120
      TabIndex        =   0
      Top             =   95
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7699
      SmallChange     =   1
      ScaleMode       =   0
      BorderStyle     =   3
      Begin VB.CommandButton cmdFP 
         Caption         =   "Command1"
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   13815
      Begin VB.TextBox txtPagaCon 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6480
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lblVuelto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   10560
         TabIndex        =   20
         Top             =   1440
         Width           =   1755
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
         Left            =   9480
         TabIndex        =   19
         Top             =   1522
         Width           =   900
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
         Left            =   5160
         TabIndex        =   18
         Top             =   1522
         Width           =   1215
      End
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6480
         TabIndex        =   14
         Top             =   878
         Width           =   1755
      End
      Begin VB.Label lblTotalPagos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   10560
         TabIndex        =   13
         Top             =   878
         Width           =   1755
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
         Left            =   4635
         TabIndex        =   12
         Top             =   960
         Width           =   1755
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
         TabIndex        =   11
         Top             =   960
         Width           =   1560
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
         TabIndex        =   10
         Top             =   360
         Width           =   1710
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
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   1710
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
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1845
      End
      Begin VB.Label lblTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   10560
         TabIndex        =   7
         Top             =   278
         Width           =   1755
      End
      Begin VB.Label lblTT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6480
         TabIndex        =   6
         Top             =   278
         Width           =   1755
      End
      Begin VB.Label lblTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2520
         TabIndex        =   5
         Top             =   278
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmVentas2FP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()

    If Me.lvFP.ListItems.count = 0 Then
        MsgBox "Debe ingresar las formas de pago.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    If CDec(Me.lblTotalPagar.Caption) <> CDec(Me.lblTotalPagos.Caption) Then
        MsgBox "Importe no coincide", vbInformation, Pub_Titulo

        Exit Sub

    End If

'    If Not oRSfp Is Nothing Then
'
'        'If Not oRSfp.EOF Then oRSfp.Delete
'        If oRSfp.RecordCount <> 0 Then
'            oRSfp.MoveFirst
'
'            Do While Not oRSfp.EOF
'                oRSfp.Delete adAffectCurrent
'                oRSfp.MoveNext
'            Loop
'
'        End If
'    End If

    Dim DD As Object
  Set oRSfp = New ADODB.Recordset
        oRSfp.CursorType = adOpenDynamic ' setting cursor type
        oRSfp.Fields.Append "IDFORMAPAGO", adInteger
        oRSfp.Fields.Append "FORMAPAGO", adVarChar, 100
        oRSfp.Fields.Append "REFERENCIA", adVarChar, 200
        oRSfp.Fields.Append "MONTO", adDouble
        oRSfp.Fields.Append "PAGACON", adDouble
        oRSfp.Fields.Append "VUELTO", adDouble
        oRSfp.Fields.Append "TIPO", adVarChar, 2
        'oRSfp.Fields.Append "formapago", adVarChar, 120
    
        oRSfp.Fields.Refresh
        oRSfp.Open
   
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
            oRSfp.Update
        Next

If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
    If Not oRSfp.EOF Then oRSfp.MoveFirst
           
    Unload Me

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdFP_Click(Index As Integer)

    '    frmFacComandaFP3.MuestraReferencia = True
    '
    '    frmFacComandaFP3.Show vbModal
    '
    '    If frmFacComandaFP3.gACEPTA Then
    '
    If (CDec(Me.lblTotalPagar.Caption) - CDec(Me.lblTotalPagos.Caption)) <> 0 Then

        Dim itemX As Object

        Set itemX = Me.lvFP.ListItems.Add(, , Me.lvFP.ListItems.count + 1)
        itemX.SubItems(1) = Me.cmdFP(Index).Caption
        itemX.Tag = Split(Me.cmdFP(Index).Tag, "|")(0)
        itemX.SubItems(2) = Trim(frmFacComandaFP3.gREF)
        itemX.SubItems(3) = CDec(Me.lblTotalPagar.Caption) - CDec(Me.lblTotalPagos.Caption)
        itemX.SubItems(4) = Split(Me.cmdFP(Index).Tag, "|")(1)

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

    
       ' Me.lblTotalPagar.Caption = FormatCurrency(CDec(frmFacComanda.lblImporte.Caption), 2)
 

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
If Me.lvFP.ListItems.count = 0 Then Exit Sub
 Me.lvFP.ListItems.Remove Me.lvFP.SelectedItem.Index
 CalcularTotales
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    Dim vTop  As Integer

    Dim vLeft As Integer

    vTop = 120
    vLeft = 120
    
    
    
    
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
        Me.cmdFP(f).Tag = oRStp!Codigo & "|" & oRStp!tipo
     
        oRStp.MoveNext
    Next

    ConfiguraLV

Dim xITEM As Object
    Dim CI As Integer
    CI = 1
Do While Not oRSfp.EOF
    Set xITEM = Me.lvFP.ListItems.Add(, , CI) ')
    xITEM.SubItems(1) = oRSfp!formapago
    xITEM.Tag = oRSfp!idformapago
    xITEM.SubItems(2) = oRSfp!referencia
    xITEM.SubItems(3) = oRSfp!monto
    xITEM.SubItems(4) = oRSfp!tipo
    CI = CI + 1
    oRSfp.MoveNext
Loop

If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst

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
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .ColumnHeaders.Add , , "Nro", 700
        .ColumnHeaders.Add , , "Forma Pago", 3000
        .ColumnHeaders.Add , , "Nro Referencia", 2100
        .ColumnHeaders.Add , , "Monto", 1400
        .ColumnHeaders.Add , , "Tipo", 500
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


Private Sub txtPagaCon_Change()
CalcularTotales
End Sub
