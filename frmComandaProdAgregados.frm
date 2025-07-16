VERSION 5.00
Begin VB.Form frmComandaProdAgregados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregados"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8745
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
   ScaleHeight     =   6090
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   8535
      Begin VB.CommandButton cmdAgregadoAnt 
         Height          =   1200
         Left            =   10
         Picture         =   "frmComandaProdAgregados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdAgregadoSig 
         Height          =   1200
         Left            =   7200
         Picture         =   "frmComandaProdAgregados.frx":70CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton cmdAgregado 
         Caption         =   "Command1"
         Height          =   1200
         Index           =   0
         Left            =   10
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Label lblProducto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8295
      End
   End
End
Attribute VB_Name = "frmComandaProdAgregados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPagActAgr, vPagTotAgr As Integer
Public gIDpadre As Integer
Private vAgregado As Integer
Public gIDfamilia As Integer
Public gSerie As String
Public gNumero As Double
Public gCliente As String
Public gComensales As Integer
Public gMozo As Integer
Public gMesa As String
Private vMaxFac As Double

Public Function adicionarAgregado(vcp As Double, _
                                  vc As Double, _
                                  vpre As Double, _
                                  vimp As Double, _
                                  vd As String, _
                                  vnumser As String, _
                                  vNumFac As Double, _
                                  VcLIENTE As String, _
                                  VcOMENSALES As Integer, _
                                  vPadre As Integer, _
                                  Optional ByRef vnumsec As Integer) As Boolean
    LimpiaParametros oCmdEjec

    Dim xPedido     As String

    Dim NumSer      As String

    Dim NumFac      As Double

    Dim vMaxNumoper As String

    On Error GoTo ErrorGraba

    Pub_ConnAdo.BeginTrans
    oCmdEjec.CommandType = adCmdStoredProc

    With oCmdEjec
        .CommandText = "SpModificarComanda1"
        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, LK_CODUSU)
        .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, gMesa)
        .Parameters.Append .CreateParameter("@cp", adDouble, adParamInput, , vcp) 'julio 11-01-2011
        .Parameters.Append .CreateParameter("@cant", adDouble, adParamInput, , vc)
        .Parameters.Append .CreateParameter("@pre", adDouble, adParamInput, , vpre)
        .Parameters.Append .CreateParameter("@imp", adDouble, adParamInput, , vimp)
        .Parameters.Append .CreateParameter("@d", adVarChar, adParamInput, 50, vd)
            
        '.Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , CInt(Me.lblMozo.Tag))
        .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , gMozo)
        
        .Parameters.Append .CreateParameter("@NumSer", adChar, adParamInput, 3, vnumser)
        .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamInput, , vNumFac)
        .Parameters.Append .CreateParameter("@NUMSEC", adInteger, adParamOutput)
        .Parameters.Append .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        .Parameters.Append .CreateParameter("@CodFam", adInteger, adParamInput, , gIDfamilia)  'linea nueva
        .Parameters.Append .CreateParameter("@CLIENTE", adVarChar, adParamInput, 120, VcLIENTE) 'Linea nueva

        .Parameters.Append .CreateParameter("@COMENSALES", adDouble, adParamInput, , VcOMENSALES)  'Linea nueva
        .Parameters.Append .CreateParameter("@PADRE", adDouble, adParamInput, , vPadre)
        '.Parameters.Append .CreateParameter("@ZONA", adInteger, adParamInput, , vCodZona)
        .Execute
        vnumsec = .Parameters("@NUMSEC").Value

    End With

    'oCmdEjec.Execute
    
    'actualiza stock
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpActualizaStock"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adDouble, adParamInput, , vcp)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adInteger, adParamInput, , 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adInteger, adParamInput, , vc)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ser", adChar, adParamInput, 3, vnumser)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , vNumFac)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adVarChar, adParamInput, 10, gMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 1) '0 cuando es extorno
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , vnumsec)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumFac", adDouble, adParamOutput, , vMaxFac)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumOper", adDouble, adParamOutput, , 0)

    oCmdEjec.Execute
    vMaxFac = oCmdEjec.Parameters("@MaxNumFac").Value
    vMaxNumoper = oCmdEjec.Parameters("@MaxNumOper").Value

    'LimpiaParametros oCmdEjec
    'oCmdEjec.CommandText = "SpActualizarPedTrans"
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adInteger, adParamInput, , Me.lblNumero.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , vnumsec)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumOper", adInteger, adParamInput, , vMaxNumoper)
    'oCmdEjec.Execute
    adicionarAgregado = True
    Pub_ConnAdo.CommitTrans

    Exit Function

ErrorGraba:
    Pub_ConnAdo.RollbackTrans
    adicionarAgregado = False
    MsgBox Err.Description

End Function

Private Sub cargarAgregados()
    Me.cmdAgregadoSig.Enabled = False
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_COMANDA_LISTAGREGADOSxFAMILIA]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idfamilia", adInteger, adParamInput, , gIDfamilia)
            
    Dim orsA As ADODB.Recordset

    Set orsA = oCmdEjec.Execute

    For i = 1 To Me.cmdAgregado.count - 1
        Unload Me.cmdAgregado(i)
    Next

    FiltrarAgregados orsA.RecordCount, orsA
   

End Sub

Private Sub FiltrarAgregados(cant As Integer, oRS As ADODB.Recordset)
    vAgregado = cant

    'Dim vPri As Boolean
    'vPri = True
    Dim f, c As Integer

    c = 1

    Dim valor As Double

    valor = vAgregado / 26
    ' agregado julio 06/08/2012==============
    pos = InStr(Trim(str(valor)), ".")

    If pos <> 0 Then
        If pos = 1 Then
            ent = Left(CStr(valor), pos)
        Else
            ent = Left(CStr(valor), pos - 1)
        End If

    Else
        ent = Int(valor)
    End If
    
    If pos <> 0 Then
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
    Else
        pos2 = 0
    End If

    If pos2 > 0 Then
        ent = ent + 1
    End If

    vPagTotAgr = ent

    If valor <> 0 Then vPagActAgr = 1

    If vAgregado > 26 Then: Me.cmdAgregadoSig.Enabled = True

    'descargar los objetos primero
    If Me.cmdAgregado.count > 1 Then

        For i = 1 To cmdAgregado.count - 1
            Unload cmdAgregado.Item(i)
        Next

    End If

    '============================
    vIniLeft = 10
    vIniTop = 120

    For i = 1 To vAgregado
        Load Me.cmdAgregado(i)

        If c <= 6 Then '1 fila
            If c = 1 Then
                vIniLeft = vIniLeft + Me.cmdAgregadoAnt.Width
            Else
                vIniLeft = vIniLeft + Me.cmdAgregado(i - i).Width
            End If

            '        Else: viniLeft = viniLeft + 970
            '        End If
        ElseIf c <= 13 Then '2º Fila

            'viniLeft = 30
            If c = 7 Then
                vIniLeft = 10
                vIniTop = vIniTop + Me.cmdAgregadoAnt.Height
                Else: vIniLeft = vIniLeft + Me.cmdAgregado(i - 1).Width
            End If

        ElseIf c <= 20 Then '3º Fila

            If c = 14 Then
                vIniTop = vIniTop + Me.cmdAgregado(13).Height
                vIniLeft = 10
                Else: vIniLeft = vIniLeft + Me.cmdAgregado(i - 1).Width
            End If

        Else '4º y ultima fila

            If c = 21 Then
                vIniTop = vIniTop + Me.cmdAgregado(20).Height
                vIniLeft = 30
                Else: vIniLeft = vIniLeft + Me.cmdAgregado(i - 1).Width
            End If
        End If

        Me.cmdAgregado(i).Left = vIniLeft
        Me.cmdAgregado(i).Top = vIniTop
        Me.cmdAgregado(i).Visible = True
        'Me.cmdAgregado(i).BackColor = Me.cmdSubFam(vColor).BackColor   'gts para mostrar el color de la familia
    
        'Me.cmdAgregado(i).Style = 1
        'Me.cmdFam(i).Visible = vPri
        Me.cmdAgregado(i).Visible = True
        Me.cmdAgregado(i).Caption = Trim(oRS!plato)
        Me.cmdAgregado(i).ToolTipText = Trim(oRS!alt)
        Me.cmdAgregado(i).Tag = oRS!Codigo & "|" & oRS!PRECIO
        '  Me.cmdAgregado(i).BackColor = Trim(ORS!alt)
        '    If c <= 14 Then
        '        Me.cmdFam(i).Visible = True
        '    Else
        '        Me.cmdFam(i).Visible = False
        '    End If
        oRS.MoveNext

        If c = 26 Then
            '        vPri = False
            c = 1
            'vuelve a empezar
            vIniLeft = 30
            vIniTop = 120
        Else
            c = c + 1
        End If
   
    Next

End Sub

Private Sub cmdAgregado_Click(Index As Integer)

    If Not VNuevo Then
        If VerificaMesa Then
            MsgBox "No se pueden agregar mas agregados.  Mesa ya Facturada.", vbInformation, Pub_Titulo
            Unload Me

            If gDefecto Then Unload frmMainMesas
            Exit Sub

        End If

    End If

    Dim oRStemp As ADODB.Recordset

    Dim DD      As Integer
    Dim vPrecio As Double
    Dim vIdPlato As Double
    
    vPrecio = Split(Me.cmdAgregado(Index).Tag, "|")(1)
    vIdPlato = Split(Me.cmdAgregado(Index).Tag, "|")(0)

'    oRsPlatos.Filter = "Codigo = '" & Me.cmdAgregado(Index).Tag & "'"

    If adicionarAgregado(vIdPlato, 1, FormatNumber(vPrecio, 2), vPrecio, "", gSerie, gNumero, gCliente, gComensales, gIDpadre, DD) Then
    
'        With Me.lvPlatos.ListItems.Add(, , Me.cmdPlato(Index).Caption, Me.ilComanda.ListImages.Item(1).key, Me.ilComanda.ListImages.Item(1).key)
'            .Tag = Me.cmdPlato(Index).Tag
'            .Checked = True
'            .SubItems(3) = FormatNumber(1, 2)
'            'obteniendo precio
'            oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"
'
'            If Not oRsPlatos.EOF Then: .SubItems(4) = FormatNumber(oRsPlatos!PRECIO, 2)
'            .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
'            .SubItems(6) = DD
'            .SubItems(7) = 0   'linea nueva
'            .SubItems(8) = vMaxFac
'            .SubItems(9) = 0
'
'        End With
'
'        oRsPlatos.Filter = ""
'        oRsPlatos.MoveFirst
'
'        If Me.lvPlatos.ListItems.count <> 0 Then
'            Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
'            Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
'            Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
'        End If

       frmComanda.CargarComanda LK_CODCIA, frmComanda.vMesa
        
'        For c = 1 To Me.lvPlatos.ListItems.count
'            Me.lvPlatos.ListItems(c).Selected = False
'        Next
'
'        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
    
    End If

End Sub

Private Sub cmdAgregadoAnt_Click()

    Dim ini, fin, f As Integer

    If vPagActAgr = 2 Then
        ini = 1
        fin = 26
    ElseIf vPagActAgr = 1 Then
        Exit Sub
    Else
        FF = vPagActAgr - 1
        ini = (26 * FF) - 25
        fin = 26 * FF

    End If

    For f = ini To fin
        Me.cmdAgregado(f).Visible = True
    Next

    If vPagActAgr > 1 Then
        vPagActAgr = vPagActAgr - 1

        If vPagActAgr = 1 Then: Me.cmdAgregadoAnt.Enabled = False
    
        Me.cmdAgregadoSig.Enabled = True

    End If

End Sub

Private Sub cmdAgregadoSig_Click()

    Dim ini, fin, f As Integer

    If vPagActAgr = 1 Then
        ini = 1
        fin = 26
    ElseIf vPagActAgr = vPagTotAgr Then
        Exit Sub
    Else
        ini = (26 * vPagActAgr) - 25
        fin = 26 * vPagActAgr

    End If

    For f = ini To fin
        Me.cmdAgregado(f).Visible = False
    Next

    If vPagActAgr < vPagTotAgr Then
        vPagActAgr = vPagActAgr + 1

        If vPagActAgr = vPagTotAgr Then: Me.cmdAgregadoSig.Enabled = False
    
        Me.cmdAgregadoAnt.Enabled = True

    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
'If KeyCode = vbKeyF3 Then cargarAgregados
End Sub

Private Sub Form_Load()
Me.cmdAgregadoAnt.Enabled = False
Me.cmdAgregadoSig.Enabled = False
cargarAgregados
End Sub
