VERSION 5.00
Begin VB.Form frmCriaAjuste 
   BackColor       =   &H00393939&
   BorderStyle     =   0  'None
   Caption         =   "Cria Ajuste"
   ClientHeight    =   3945
   ClientLeft      =   1725
   ClientTop       =   3960
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraInfo 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   1575
      Left            =   0
      TabIndex        =   11
      Top             =   1545
      Width           =   9495
      Begin VB.ComboBox cmbMotivo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1920
         TabIndex        =   21
         ToolTipText     =   "Motivo"
         Top             =   960
         Width           =   7575
      End
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   150
         TabIndex        =   20
         ToolTipText     =   "Ajuste"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtEstoque 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8400
         TabIndex        =   19
         ToolTipText     =   "Estoque"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         ToolTipText     =   "Descrição"
         Top             =   360
         Width           =   6255
      End
      Begin VB.TextBox txtReferencia 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   150
         MaxLength       =   7
         TabIndex        =   17
         ToolTipText     =   "Referência"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lbEstoque 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   120
         Width           =   585
      End
      Begin VB.Label lbDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbReferencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Frame fraAjuste 
      BackColor       =   &H002E2E2E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   6345
      TabIndex        =   7
      Top             =   150
      Width           =   3135
      Begin VB.OptionButton lbInventario 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         Caption         =   "Inventário"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Inventário"
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton lbContagem 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         Caption         =   "Contagem"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Contagem"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAjuste 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         Caption         =   "Ajuste"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Ajuste"
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraLoja 
      BackColor       =   &H002E2E2E&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   6015
      Begin VB.OptionButton optAjusteSistemaLoja 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         Caption         =   "Ajustar Estoque do Sistema e Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "Ajustar Estoque do Sistema e Loja"
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox cmbLoja 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "Loja"
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optAjusteSistema 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         Caption         =   "Ajustar Estoque Somente do Sistema"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Ajustar Estoque Somente do Sistema"
         Top             =   240
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optAjusteLoja 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         Caption         =   "Ajustar Estoque Somente da Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Ajustar Estoque Apenas da Loja"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lbLoja 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   300
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   9345
      TabIndex        =   0
      Top             =   3135
      Width           =   9350
   End
   Begin CentroDeDistribuicao.chameleonButton cmdGrava 
      Height          =   510
      Left            =   6645
      TabIndex        =   22
      ToolTipText     =   "Grava"
      Top             =   3285
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Gravar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   5263440
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorno 
      Height          =   510
      Left            =   8085
      TabIndex        =   23
      ToolTipText     =   "Grava"
      Top             =   3285
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Retornar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   5263440
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmCriaAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdRetorno_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    Me.Width = 9640
    Me.Height = 3920
    carregarPosicaoTela Me
End Sub

Private Sub Form_Load()

    carregaComboloja
    CarregaComboMotivo
    
'    carregarPosicaoFrame frmCriaAjuste
    
End Sub

Private Sub Gravar()

    Dim sql As String
    Dim alteracao As String
    Dim Tipo As String
    
    
    If optAjusteSistema Then
        alteracao = "'S'"
    ElseIf optAjusteLoja Then
        alteracao = "'L'"
    Else
        alteracao = "'A'"
    End If
    
    If optAjuste Then
        Tipo = "'A'"
    ElseIf optContagem Then
        Tipo = "'C'"
    Else
        Tipo = "'I'"
    End If
    
    
    sql = "exec SP_CDM_Cria_ajuste '" & Trim(cmbLoja.Text) & " ','" & txtReferencia.Text & "'," & txtAjuste.Text & ",'" & Trim(Mid(cmbMotivo.Text, 1, 2)) & "','13'," & alteracao & ",'A'," & Tipo
    ADO_Cn_CDLocal.Execute (sql)
    
    txtReferencia.Text = ""
    txtAjuste.Text = ""
    txtDescricao.Text = ""
    txtEstoque.Text = ""
    
    

End Sub


Private Sub cmdGrava_Click()
    
    Gravar
    
End Sub

Private Sub CarregaProduto()
    
    Dim sql As String
    Dim rsProduto As New ADODB.Recordset
    
    txtReferencia.Enabled = False
    cmbLoja.Enabled = False
    
    
    sql = "Exec sp_cdm_Busca_Ajuste 3, '" & Trim(cmbLoja.Text) & "',0,0,0,0,0, '" & txtReferencia.Text & "'"
    
    rsProduto.CursorLocation = adUseClient
    rsProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsProduto.EOF Then
    
        txtDescricao.Text = Trim(rsProduto("pr_descricao"))
        txtEstoque.Text = Trim(rsProduto("es_estoque"))
        txtReferencia.Enabled = True
        txtAjuste.SetFocus
    
    Else
    
         MsgBox "Referência não cadastrada.", vbExclamation, "Atenção!"
         txtReferencia.Text = ""
         txtReferencia.Enabled = True
         txtReferencia.SetFocus
    
    End If
    
    rsProduto.Close
    cmbLoja.Enabled = True
    

End Sub


Private Sub carregaComboloja()
    
    Dim rsLoja As New ADODB.Recordset
    Dim sql As String
    
    sql = "Exec sp_cdm_busca_ajuste 1, 0,0,0,0,0,0,0"
    
    rsLoja.CursorLocation = adUseClient
    rsLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    cmbLoja.Text = "181"
    If Not rsLoja.EOF Then
        Do While Not rsLoja.EOF
            
            cmbLoja.AddItem rsLoja("lo_loja")
            rsLoja.MoveNext
        
        Loop
    End If
    
    rsLoja.Close
    
End Sub


Private Sub txtReferencia_lostfocus()
    CarregaProduto
End Sub


Private Sub CarregaComboMotivo()

    Dim rsMotivo As New ADODB.Recordset
    Dim sql As String
    
    sql = "Exec sp_cdm_busca_ajuste 4,0,0,0,0,0,0,0 "
    
    rsMotivo.CursorLocation = adUseClient
    rsMotivo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsMotivo.EOF Then
        cmbMotivo.Text = rsMotivo("ma_codigoMotivo") & " - " & rsMotivo("ma_descricao")
        
        Do While Not rsMotivo.EOF
            cmbMotivo.AddItem rsMotivo("ma_codigoMotivo") & " - " & rsMotivo("ma_descricao")
            rsMotivo.MoveNext
        Loop
        
    End If
    rsMotivo.Close

End Sub
