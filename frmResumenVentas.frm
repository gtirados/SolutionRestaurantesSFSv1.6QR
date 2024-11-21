VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmResumenVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Ventas por Usuario"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdVer 
      Caption         =   "Imprimir"
      Height          =   600
      Left            =   7200
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   600
      Left            =   5640
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSDataListLib.DataCombo DatUsuario 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSMask.MaskEdBox mbHasta 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   495
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mbDesde 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   495
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmResumenVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMostrar_Click()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_VENTAS_CONSULTA_RESUMEN"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@desde", adDBTimeStamp, adParamInput, , Me.mbDesde.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@hasta", adDBTimeStamp, adParamInput, , Me.mbHasta.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODUSU", adVarChar, adParamInput, 20, Me.DatUsuario.BoundText)
    
    Dim orsD As ADODB.Recordset
    Set orsD = oCmdEjec.Execute
    
    Me.lvDatos.ListItems.Clear
    
    Dim xItem As Object
    Do While Not orsD.EOF
        Set xItem = Me.lvDatos.ListItems.Add(, , orsD!PRODUCTO)
        xItem.SubItems(1) = orsD!Cantidad
        orsD.MoveNext
    
    Loop
    
End Sub

Private Sub cmdVer_Click()

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim RutaReporte As String

    RutaReporte = PUB_RUTA_REPORTE & "resumenVentas.rpt"  'LO CAMBIA CUANDO QUIERA

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "pFECHAS"
                crParamDef.AddCurrentValue "DESDE: " & Me.mbDesde.Text & Space(10) & "DESDE: " & Me.mbHasta.Text ' str(vPlato)
                   Case "pUSUARIO"
                crParamDef.AddCurrentValue Me.DatUsuario.Text ' str(vPlato)
        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_VENTAS_CONSULTA_RESUMEN"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset



 oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@desde", adDBTimeStamp, adParamInput, , Me.mbDesde.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@hasta", adDBTimeStamp, adParamInput, , Me.mbHasta.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODUSU", adVarChar, adParamInput, 20, Me.DatUsuario.BoundText)
    
   Set rsd = oCmdEjec.Execute

  

    VReporte.Database.SetDataSource rsd, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    
      
        frmVisor.cr.ReportSource = VReporte
          frmVisor.cr.ViewReport
        frmVisor.Show vbModal
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub Form_Load()
ConfiguraLV
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "splistarusuarios"
Dim orsD As ADODB.Recordset
Set orsD = oCmdEjec.Execute
Set Me.DatUsuario.RowSource = orsD
Me.DatUsuario.ListField = orsD.Fields(1).Name
Me.DatUsuario.BoundColumn = orsD.Fields(0).Name
Me.mbDesde.Text = LK_FECHA_DIA
Me.mbHasta.Text = LK_FECHA_DIA
Me.DatUsuario.BoundText = orsD!Codigo
End Sub

Private Sub ConfiguraLV()
With Me.lvDatos
.Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Producto", 4000
    .ColumnHeaders.Add , , "Cantidad", 700
    

End With
End Sub
