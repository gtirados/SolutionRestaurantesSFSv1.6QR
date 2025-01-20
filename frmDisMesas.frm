VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.1#0"; "Codejock.Controls.v12.1.1.ocx"
Begin VB.Form frmDisMesas 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Distribuci�n de Mesas"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDisMesas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton pbLibre 
      Height          =   1095
      Left            =   8040
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
      _Version        =   786433
      _ExtentX        =   2143
      _ExtentY        =   1931
      _StockProps     =   79
      Appearance      =   6
      DrawFocusRect   =   0   'False
      Picture         =   "frmDisMesas.frx":06EA
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   9240
      Top             =   3240
   End
   Begin VB.Timer tmrMensaje 
      Interval        =   500
      Left            =   5520
      Top             =   2640
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar Distribuci�n de Mesas"
      Height          =   735
      Left            =   120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   975
      Index           =   0
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin XtremeSuiteControls.PushButton pbMesa 
      Height          =   1095
      Index           =   0
      Left            =   10200
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
      _Version        =   786433
      _ExtentX        =   2143
      _ExtentY        =   1931
      _StockProps     =   79
      Appearance      =   6
      DrawFocusRect   =   0   'False
   End
   Begin XtremeSuiteControls.PushButton pbOcupada 
      Height          =   1095
      Left            =   9240
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
      _Version        =   786433
      _ExtentX        =   2143
      _ExtentY        =   1931
      _StockProps     =   79
      Appearance      =   6
      DrawFocusRect   =   0   'False
      Picture         =   "frmDisMesas.frx":3944
   End
   Begin XtremeSuiteControls.PushButton pbCuenta 
      Height          =   1095
      Left            =   8040
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
      _Version        =   786433
      _ExtentX        =   2143
      _ExtentY        =   1931
      _StockProps     =   79
      Appearance      =   6
      DrawFocusRect   =   0   'False
      Picture         =   "frmDisMesas.frx":6B9E
   End
   Begin XtremeSuiteControls.PushButton pbReservada 
      Height          =   1095
      Left            =   9240
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
      _Version        =   786433
      _ExtentX        =   2143
      _ExtentY        =   1931
      _StockProps     =   79
      Appearance      =   6
      DrawFocusRect   =   0   'False
      Picture         =   "frmDisMesas.frx":9DF8
   End
   Begin XtremeSuiteControls.PushButton cmdSalir 
      Height          =   1095
      Left            =   10680
      TabIndex        =   16
      Top             =   5520
      Width           =   1095
      _Version        =   786433
      _ExtentX        =   1931
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "&Salir"
      Appearance      =   6
      DrawFocusRect   =   0   'False
      Picture         =   "frmDisMesas.frx":D052
      TextImageRelation=   1
   End
   Begin VB.Image imgR 
      Height          =   495
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblsolution 
      BackStyle       =   0  'Transparent
      Caption         =   "GT SOFTWARE SAC Tel.044-250522 Cel.949433704  RPM: *366258"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1800
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblmE 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "MESA EN CUENTA"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7200
      TabIndex        =   9
      Top             =   6720
      Width           =   1530
   End
   Begin VB.Image imgE 
      Height          =   495
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblF4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "F4 - Mantenimiento Mesas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   2625
   End
   Begin VB.Label lblF5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "F5 - Actualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   840
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblmR 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "MESA RESERVADA"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   6750
      Width           =   1605
   End
   Begin VB.Label lblmO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "MESA OCUPADA"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5040
      TabIndex        =   5
      Top             =   6750
      Width           =   1410
   End
   Begin VB.Label lblmL 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "MESA LIBRE"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   6750
      Width           =   1050
   End
   Begin VB.Image imgO 
      Height          =   495
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   615
   End
   Begin VB.Image imgL 
      Height          =   495
      Left            =   120
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblmensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
   Begin VB.Label lblNomMesa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   10050
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmDisMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private p As Integer
Dim oMesa() As String
Dim oPosLeft() As String
Dim oPosTop() As String
Private vModifica As Boolean
Private vEstadoFlag As Boolean
Public vMesa As String
Public VZONA As Integer

Private Sub LlamarxTecla()
Dim ve As Boolean
ve = False
For i = 1 To Me.pbMesa.count - 1
If vMesa = Me.lblNomMesa(i).Tag Then
    ve = True
    Exit For
End If
Next

If ve Then
   ' MsgBox vMesa
    vMesa = ""
     If Me.pbMesa(i).Tag = "L" Or Me.pbMesa(i).Tag = "R" Then 'Mesa Libre
        frmComanda.VNuevo = True
    Else
        frmComanda.VNuevo = False
    End If
    frmComanda.vEstado = Me.pbMesa(i).Tag
    frmComanda.vMesa = Me.lblNomMesa(i).Tag
    frmComanda.vCodZona = VZONA
    'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
    frmComanda.Caption = "Comanda : " & Me.lblNomMesa(i).Caption
    frmComanda.lblmesa.Caption = Me.lblNomMesa(i).Caption  'agregado gts
    
    frmComanda.Show vbModal
Else
vMesa = ""
End If


'If vModifica = False Then
'    If Me.pbMESA(Index).Tag = "L" Or Me.pbMESA(Index).Tag = "R" Then 'Mesa Libre
'        frmComanda.VNuevo = True
'    Else
'        frmComanda.VNuevo = False
'    End If
'    frmComanda.vEstado = Me.pbMESA(Index).Tag
'    frmComanda.vMesa = Me.lblNomMesa(Index).Tag
'    frmComanda.vCodZona = vZona
'    'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
'    frmComanda.Caption = "Comanda : " & Me.lblNomMesa(Index).Caption
'    frmComanda.Show vbModal
'End If
End Sub

Private Sub CreaEstructuraXML(vCadena As String)

Dim i As Integer
vCadena = "<r>"
For i = 1 To Me.pbMesa.count - 1
    vCadena = vCadena & "<d "
    vCadena = vCadena & "codmesa=""" & Trim(Me.lblNomMesa(i).Tag) & """ "
    vCadena = vCadena & "posleft=""" & Trim(Me.pbMesa(i).Left) & """ "
    vCadena = vCadena & "postop=""" & Trim(Me.pbMesa(i).Top) & """ "
    vCadena = vCadena & "/>"
Next
vCadena = vCadena & "</r>"
End Sub

Private Sub CmdModificar_Click()
If Left(Me.cmdModificar.Caption, 1) = "M" Then 'Modifica
    vModifica = True
    Me.Timer1.Enabled = False
    Me.cmdModificar.Caption = "Graba Distribuci�n de las Mesas"
    Me.lblMensaje.Visible = True
    Me.lblMensaje.Caption = "Arrastre la Mesa hasta la nueva posici�n."
Else 'Graba

On Error GoTo Graba
    Me.lblMensaje.Visible = False
    vModifica = False
    Me.cmdModificar.Caption = "Modifica Distribuci�n de las Mesas"
    Dim vStrMesas As String
    CreaEstructuraXML vStrMesas
    oCmdEjec.CommandText = "SpModificarUbicacionMesas"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xmldata", adVarChar, adParamInput, 8000, vStrMesas)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZona", adInteger, adParamInput, , VZONA) 'JULIO 09/02/2011
    oCmdEjec.Execute
     Me.Timer1.Enabled = True
    Exit Sub
Graba:
    LimpiaParametros oCmdEjec
    MsgBox Err.Description, vbInformation, TituloMsgBox
End If
End Sub

Public Sub CargarMesas(xZona As Integer)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpCargarMesas"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , xZona)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    Set ORSmESAS = oCmdEjec.Execute

    If ORSmESAS.RecordCount = 0 Then Exit Sub

    Dim loopIndice As Integer

    'limpiando las mesas
    For loopIndice = 1 To Me.pbMesa.count - 1
        Unload Me.lblNomMesa(loopIndice)
        Unload Me.pbMesa(loopIndice)
    Next

    For loopIndice = 1 To ORSmESAS.RecordCount
        Load pbMesa(loopIndice)
        
        Me.pbMesa(loopIndice).Left = ORSmESAS!posleft
        Me.pbMesa(loopIndice).Top = ORSmESAS!postop
    
        If Trim(ORSmESAS!ESTADO) = "L" Then
            Me.pbMesa(loopIndice).Picture = Me.pbLibre.Picture
        ElseIf Trim(ORSmESAS!ESTADO) = "E" Then
            Me.pbMesa(loopIndice).Picture = Me.pbCuenta.Picture
        ElseIf Trim(ORSmESAS!ESTADO) = "R" Then
            Me.pbMesa(loopIndice).Picture = Me.pbReservada.Picture
        ElseIf Trim(ORSmESAS!ESTADO) = "O" Then
            Me.pbMesa(loopIndice).Picture = Me.pbOcupada.Picture
        End If

        Me.pbMesa(loopIndice).Tag = Trim(ORSmESAS!ESTADO)

        Select Case ORSmESAS!ESTADO

            Case "L": Me.pbMesa(loopIndice).ToolTipText = "Mesa Libre"

            Case "E": Me.pbMesa(loopIndice).ToolTipText = "Mesa En Cuenta"

            Case "O": Me.pbMesa(loopIndice).ToolTipText = "Mesa Ocupada"

            Case Else: Me.pbMesa(loopIndice).ToolTipText = "Mesa Reservada"
        End Select
    
        Load lblNomMesa(loopIndice)
       
        Me.lblNomMesa(loopIndice).Caption = Trim(ORSmESAS!mesa)
      ' If Trim(ORSmESAS!cliente) <> "" And Trim(ORSmESAS!ESTADO) <> "L" Then Me.lblNomMesa(loopIndice).Caption = Me.lblNomMesa(loopIndice).Caption & vbCrLf & Trim(ORSmESAS!cliente)  ' gts aca se muestra nombre cliente
        Me.lblNomMesa(loopIndice).Tag = Trim(ORSmESAS!codmesa)
        Me.lblNomMesa(loopIndice).Move ORSmESAS!posleft, ORSmESAS!postop + pbMesa(loopIndice).Height
        'picTable(loopIndice).Caption = Trim(oRsMesas!Mesa)
        
         Me.pbMesa(loopIndice).Visible = True
          Me.lblNomMesa(loopIndice).Visible = True
        ORSmESAS.MoveNext
    Next

End Sub
Private Sub btnObj_Click()
MsgBox btnObj.Name
End Sub

Private Sub cmdSalir_Click()
frmZonas.Show
Unload Me
End Sub

Private Sub Command3_Click(index As Integer)
CargarMesas VZONA
End Sub

Private Sub Form_Activate()
Me.WindowState = 2
End Sub

Sub Form_DragDrop(Source As Control, x As Single, y As Single)
If vModifica Then
Source.Move (x - DragX), (y - DragY)
lblNomMesa(p).Move Me.pbMesa(p).Left, Me.pbMesa(p).Top + Me.pbMesa(p).Height
Else
CargarMesas VZONA
Me.Timer1.Enabled = True
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    frmZonas.Show
    Unload Me
End If
If KeyCode = vbKeyF4 Then frmMesas.Show vbModal
If KeyCode = vbKeyF5 Then CargarMesas VZONA
Select Case Shift
      Case 1
        If KeyCode <> 16 Then
            vMesa = vMesa & Chr(KeyCode)
        End If
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
      Case 16 '-- TECLA shift
        LlamarxTecla

   End Select
End Sub

Private Sub Form_Load()
    InhabilitarCerrar Me
    vEstadoFlag = True

    vModifica = False

    If LK_CODUSU = "MOZO" Then
        cmdModificar.Enabled = False
    End If

    'Leer_Mesas App.Path & "\Mesas.txt", vbTab
    p = 0
    CargarMesas VZONA
    'Cargar Imagenes de Leyenda
    Me.imgL.Picture = Me.pbLibre.Picture
    Me.imgE.Picture = Me.pbCuenta.Picture
    Me.imgR.Picture = Me.pbReservada.Picture
    Me.imgO.Picture = Me.pbOcupada.Picture
    'Me.imgU.Picture = Me.ilMesas.ListImages(5).Picture
End Sub

Private Sub Form_Resize()
Me.lblF5.Move 30, Me.ScaleHeight - Me.imgL.Height - 200
Me.lblF4.Move 30, Me.ScaleHeight - Me.imgL.Height - 400
Me.imgL.Top = Me.lblF4.Top + 300
Me.imgO.Top = Me.imgL.Top
Me.imgR.Top = Me.imgL.Top
Me.imgE.Top = Me.imgL.Top
'Me.imgU.Top = Me.imgL.Top

Me.lblmL.Top = Me.lblF4.Top + 400
Me.lblmO.Top = Me.lblmL.Top
Me.lblmR.Top = Me.lblmL.Top
Me.lblmE.Top = Me.lblmL.Top
'Me.lblmU.Top = Me.lblmL.Top

'Me.imgL.Move 30, Me.ScaleHeight - Me.imgL.Height  ' - (Me.imgO.Height + Me.imgR.Height)
'Me.lblmL.Move Me.imgL.Width + 170, Me.imgL.Top + 150
'
'Me.imgO.Move Me.lblmL.Width + Me.imgL.Width + 300, Me.imgL.Top
'Me.lblmO.Move Me.imgL.Width + Me.lblmL.Width + Me.imgO.Width + 300, Me.imgL.Top + 150
'
'Me.imgR.Move Me.lblmL.Width + Me.imgL.Width + Me.lblmO.Width + Me.imgO.Width + 350, Me.imgL.Top
'Me.lblmR.Move Me.imgL.Width + Me.lblmL.Width + Me.imgO.Width + Me.lblmO.Width + Me.imgR.Width + 300, Me.imgL.Top + 150
'
'Me.imgE.Move Me.lblmR.Width + Me.imgR.Width + Me.lblmL.Width + Me.lblmO.Width + Me.imgO.Width + 1000, Me.imgL.Top
'Me.lblmE.Move Me.imgR.Width + Me.lblmR.Width + Me.imgL.Width + Me.lblmO.Width + Me.imgO.Width + Me.imgR.Width + 1500, Me.imgL.Top + 150
'
'Me.imgU.Move Me.lblmE.Width + Me.lblmR.Width + Me.imgR.Width + Me.lblmL.Width + Me.lblmO.Width + Me.imgO.Width + 1000, Me.imgL.Top
'Me.lblmU.Move Me.imgE.Width + Me.imgR.Width + Me.lblmR.Width + Me.imgL.Width + Me.lblmO.Width + Me.imgO.Width + Me.imgR.Width + 1500, Me.imgL.Top + 150

Me.cmdSalir.Move (Me.ScaleWidth - Me.cmdSalir.Width), (Me.ScaleHeight - Me.cmdSalir.Height)

'Me.lblsolution.Move 8800, Me.ScaleHeight - Me.imgE.Height - 150
'Me.lblsolution.Move 30, Me.ScaleLeft - Me.imgE.Left - 18080

End Sub

Private Sub imgMesa_Click(index As Integer)

  

End Sub

Private Sub imgMesa_DragDrop(index As Integer, _
                             Source As Control, _
                             x As Single, _
                             y As Single)

 
End Sub

Private Sub imgMesa_MouseDown(index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)

  
End Sub

Private Sub imgMesa_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub pbMESA_Click(index As Integer)
  If vModifica = False Then
        '        If Me.pbMESA(Index).Tag = "U" Then
        '            MsgBox "La Mesa se encuentra en uso.", vbInformation, Pub_Titulo
        '
        '            Exit Sub
        '
        '        End If

        ' AGREGADO GTS PARA VERIFICAR FECHA DEL DIA=========================================================
        SQ_OPER = 1
        PUB_CODCIA = LK_CODCIA
        LEER_PAR_LLAVE

        If par_llave!par_flag_cierre = 9 Then
            MsgBox "!!! Compa�ia ... Cerr� Operaciones ... Llamar al Administrador ", 48, Pub_Titulo
            Unload Me

            'GoTo salirf
            Exit Sub

        Else
        End If

        If LK_FLAG_GRIFO <> "A" Then
            If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
                MsgBox "!!!FECHA YA NO COINCIDE CON LA ACTUAL , OTRO USUARIO HA CERRADO EL DIA!!! SALGA Y REINICIE SU SISTEMA...", 48, Pub_Titulo
                End
                'GoTo salirf
            End If
        End If

        Dim RSmesaszonas As ADODB.Recordset

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPVERIRICA_MESAMOZO"
        Set RSmesaszonas = oCmdEjec.Execute(, LK_CODCIA)

        Dim vPasa As Boolean

        Dim rMOZO As Boolean

        vPasa = False
        
        ' AGREGADO GTS PARA C�VERIFICAR FECHA DEL DIA=======================================================
        Dim orsM  As ADODB.Recordset

        Dim xmozo As Integer

        If Me.pbMesa(index).Tag = "L" Or Me.pbMesa(index).Tag = "R" Then 'Mesa Libre
        
            If LK_USU_PASSWORD = "A" Then   'requiere password
                
                If Not RSmesaszonas.EOF Then 'no hay mozos asignados a la mesa
                    ' If CBool(RSmesaszonas!Dato) Then 'aqui entra a pedir la clave del mozo de la mesa asignada
                    'Me.lblNomMesa(Index).Tag 'codido de mesa
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SPVERIFICA_MOZOMESA"

                    Set orsM = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblNomMesa(index).Tag))

                    If orsM.EOF Then
                        vPasa = True
                    Else
                        'frmcomandamozomesa.gCodMesa = Me.lblNomMesa(Index).Tag
                        Set frmcomandamozomesa.DatMozos.RowSource = orsM
                        frmcomandamozomesa.DatMozos.ListField = "MOZO"
                        frmcomandamozomesa.DatMozos.BoundColumn = "CODMOZO"
                        frmcomandamozomesa.DatMozos.BoundText = orsM!CODMOZO
                        frmcomandamozomesa.gModificaMozo = False
                        frmcomandamozomesa.Show vbModal
                        xmozo = frmcomandamozomesa.gCodMozo
                        vPasa = frmcomandamozomesa.gEntro
                    End If

                    ' End If
                End If

                If vPasa Then
                    frmComanda.vPrimero = True
                    frmComanda.VNuevo = True
                    frmComanda.gMozo = xmozo
                End If
            End If

            '            End If
           
        Else 'mesa ocupada

            If LK_USU_PASSWORD = "A" Then   'requiere password
                
                If Not RSmesaszonas.EOF Then 'verifica si hay mozos asignados a la mesa
                    ' If CBool(RSmesaszonas!Dato) Then 'aqui entra a pedir la clave del mozo de la mesa asignada
                    'Me.lblNomMesa(Index).Tag 'codido de mesa
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SPVERIFICA_MOZOMESA"

                    Set orsM = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblNomMesa(index).Tag))

                    If orsM.EOF Then
                        vPasa = True
                    Else
                        Set frmcomandamozomesa.DatMozos.RowSource = orsM
                        frmcomandamozomesa.DatMozos.ListField = "MOZO"
                        frmcomandamozomesa.DatMozos.BoundColumn = "CODMOZO"
                        frmcomandamozomesa.DatMozos.BoundText = orsM!CODMOZO
                        frmcomandamozomesa.gModificaMozo = False
                        frmcomandamozomesa.Show vbModal
                                
                        vPasa = frmcomandamozomesa.gEntro
                        xmozo = frmcomandamozomesa.gCodMozo
                    End If

                    ' End If
                    
                End If
  
                '            Else 'no requiere password
                '                LimpiaParametros oCmdEjec
                '                oCmdEjec.CommandText = "spCargarMozosBYmesa"
                '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , xZona)
                '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                '                Set oRsMesas = oCmdEjec.Execute
            End If
        
            frmComanda.VNuevo = False
            frmComanda.vPrimero = False
            frmComanda.gMozo = xmozo

        End If

        If LK_USU_PASSWORD = "A" Then
            If Not vPasa Then Exit Sub
            If Me.pbMesa(index).Tag = "L" Or Me.pbMesa(index).Tag = "R" Then 'Mesa Libre
        
                frmComanda.VNuevo = True
                frmComanda.vPrimero = True
                frmComanda.vEstado = Me.pbMesa(index).Tag
                frmComanda.vMesa = Me.lblNomMesa(index).Tag
                frmComanda.vCodZona = VZONA
                'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
                frmComanda.lblmesa.Caption = Me.lblNomMesa(index).Caption
    
                Me.Timer1.Enabled = False
                frmComanda.Show vbModal
                Me.Timer1.Enabled = True

            Else
                frmComanda.VNuevo = False
                frmComanda.vPrimero = False
                frmComanda.vEstado = Me.pbMesa(index).Tag
                frmComanda.vMesa = Me.lblNomMesa(index).Tag
                frmComanda.vCodZona = VZONA
                'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
                frmComanda.lblmesa.Caption = Me.lblNomMesa(index).Caption
    
                Me.Timer1.Enabled = False
                frmComanda.Show vbModal
                Me.Timer1.Enabled = True
            End If

        Else

            If Me.pbMesa(index).Tag = "L" Or Me.pbMesa(index).Tag = "R" Then 'Mesa Libre
        
                frmComanda.VNuevo = True
                frmComanda.vPrimero = True
                frmComanda.vEstado = Me.pbMesa(index).Tag
                frmComanda.vMesa = Me.lblNomMesa(index).Tag
                frmComanda.vCodZona = VZONA
                'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
                frmComanda.lblmesa.Caption = Me.lblNomMesa(index).Caption
                Me.Timer1.Enabled = False
                frmComanda.Show vbModal
                Me.Timer1.Enabled = True
            Else
                frmComanda.VNuevo = False
                frmComanda.vPrimero = False
                frmComanda.vEstado = Me.pbMesa(index).Tag
                frmComanda.vMesa = Me.lblNomMesa(index).Tag
                frmComanda.vCodZona = VZONA
                'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
                frmComanda.lblmesa.Caption = Me.lblNomMesa(index).Caption
    
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "spcargamozosBYmesa"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adVarChar, adParamInput, 10, Me.lblNomMesa(index).Tag) 'JULIO 09/02/2011
               
                Dim ORSx As ADODB.Recordset

                Set ORSx = oCmdEjec.Execute
                
                If Not ORSx.EOF Then
                    frmComanda.gMozo = ORSx!CODMOZO
                End If
    
                Me.Timer1.Enabled = False
                frmComanda.Show vbModal
                Me.Timer1.Enabled = True
            End If
        End If
   
    End If
End Sub

Private Sub pbMESA_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
   'MsgBox "Mesa origen: " & Me.lblNomMesa(Source.Index).Tag & " => " & Me.lblNomMesa(Source.Index).Caption
    'MsgBox "Mesa destino: " & Me.lblNomMesa(Index).Tag & " => " & Me.lblNomMesa(Index).Caption

    On Error GoTo UneMesas

    '    LimpiaParametros oCmdEjec
    '    oCmdEjec.CommandText = "CP_VERIFICA_UNION_MESAS"
    '
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@usuario", adVarChar, adParamInput, 10, LK_CODUSU)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@origen", adVarChar, adParamInput, 10, Me.lblNomMesa(Source.Index).Tag)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@destino", adVarChar, adParamInput, 10, Me.lblNomMesa(Index).Tag)
    '    Set oRSfp = oCmdEjec.Execute
    '
    '    Dim xMENSAJE As String
    '
    '    If Not oRSfp.EOF Then
    '        xMENSAJE = oRSfp!Mensaje
    '
    '        If Split(xMENSAJE, "=")(0) <> 0 Then
    '            MsgBox Split(xMENSAJE, "=")(1), vbInformation, Pub_Titulo
    '        Else
        
    If MsgBox("�Desea mover el Pedido de la mesa " & Me.lblNomMesa(Source.index).Caption & vbCrLf & "a la mesa " & Me.lblNomMesa(index).Tag & "?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_PEDIDOS_UNIR_MESAS"
       
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@usuario", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@origen", adVarChar, adParamInput, 10, Me.lblNomMesa(Source.index).Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@destino", adVarChar, adParamInput, 10, Me.lblNomMesa(index).Tag)
    
    Set oRSfp = oCmdEjec.Execute
            
    If Split(oRSfp!Exito, "=")(0) <> 0 Then
        MsgBox Split(oRSfp!Exito, "=")(1), vbCritical, Pub_Titulo
    Else
        CargarMesas VZONA
        Timer1.Enabled = True
        MsgBox Split(oRSfp!Exito, "=")(1), vbInformation, Pub_Titulo
    End If
            
    '        End If
    '
    '    Else
    '        MsgBox "Error en la aplicaci�n." & vbCrLf & "Contacte con el administrador del sistema", vbCritical, Pub_Titulo
    '    End If
    
    Exit Sub
    
UneMesas:
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub pbMESA_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  'Si el boton del raton es el derecho, no hacemos nada
    If Button = 1 Then Exit Sub
    'If vModifica = False Then Exit Sub
    Me.pbMesa(index).Drag 1
    Timer1.Enabled = False
    'DragX = X
    'DragY = Y
    p = index
End Sub

Private Sub pbMESA_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Me.pbMesa(index).Drag 2
End Sub

Private Sub Timer1_Timer()
CargarMesas VZONA
End Sub

Private Sub tmrMensaje_Timer()
If vModifica Then
    If Me.lblMensaje.Visible = False Then
        Me.lblMensaje.Visible = True
    Else
        Me.lblMensaje.Visible = False
    End If
End If
End Sub


