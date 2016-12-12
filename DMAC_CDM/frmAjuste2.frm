VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmAjuste 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Ajuste"
   ClientHeight    =   7605
   ClientLeft      =   1710
   ClientTop       =   2715
   ClientWidth     =   17895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   17895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPeriodo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   150
      TabIndex        =   17
      Top             =   150
      Width           =   4065
      Begin VB.ComboBox cmbLojas 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   150
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2685
         TabIndex        =   19
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataInicial 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2445
         TabIndex        =   23
         Top             =   705
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   705
         Width           =   570
      End
      Begin VB.Label lbLoja 
         AutoSize        =   -1  'True
         BackColor       =   &H00505050&
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   195
         Width           =   300
      End
   End
   Begin VB.Frame fraTipoAjuste 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   11640
      TabIndex        =   13
      Top             =   150
      Width           =   3390
      Begin VB.OptionButton optInventario 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Inventário"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   16
         Top             =   750
         Width           =   2175
      End
      Begin VB.OptionButton optContagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Contagem"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   15
         Top             =   450
         Width           =   2175
      End
      Begin VB.OptionButton optAjuste 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Ajuste"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   14
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   8025
      TabIndex        =   10
      Top             =   150
      Width           =   3450
      Begin VB.OptionButton optProcessados 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Processados"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   12
         Top             =   450
         Width           =   2175
      End
      Begin VB.OptionButton optAberto 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Em Aberto"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   11
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Frame fraTipoAlteracao 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   4395
      TabIndex        =   6
      Top             =   150
      Width           =   3450
      Begin VB.OptionButton optAmbos 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Ambos"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   2175
      End
      Begin VB.OptionButton optSistema 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Alteração no Sistema"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   8
         Top             =   450
         Width           =   2175
      End
      Begin VB.OptionButton optloja 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Alteração na Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   215
         Left            =   150
         TabIndex        =   7
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   10740
      TabIndex        =   2
      ToolTipText     =   "Grava"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Pesquisa"
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
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste2.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   0
      Top             =   6810
      Width           =   14880
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdAjuste 
      Height          =   5220
      Left            =   150
      TabIndex        =   1
      Top             =   1425
      Width           =   14880
      _cx             =   26247
      _cy             =   9208
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   3158064
      ForeColor       =   12632256
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16423203
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   5263440
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAjuste2.frx":001C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdNovo 
      Height          =   510
      Left            =   9300
      TabIndex        =   3
      ToolTipText     =   "Grava"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Criar Ajuste"
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
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste2.frx":0148
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdLimpa 
      Height          =   510
      Left            =   12180
      TabIndex        =   4
      ToolTipText     =   "Grava"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Limpa"
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
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste2.frx":0164
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13620
      TabIndex        =   5
      ToolTipText     =   "Grava"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Retorna"
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
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste2.frx":0180
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEnviaLoja 
      Height          =   510
      Left            =   7800
      TabIndex        =   24
      ToolTipText     =   "Grava"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Enviar p/ Loja"
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
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmAjuste2.frx":019C
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
Attribute VB_Name = "frmAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviaLoja_Click()
    Dim sql As String
    
    If Not GLB_modoOffline Then
        sql = "SP_Atualiza_Tarefas 1"
        
        ADO_Cn_CDLocal.Execute (sql)
    End If
End Sub

Private Sub cmdLimpa_Click()

    mskDataInicial.Text = Format(DateAdd("d", -1, Date), "dd/mm/yyyy")
    mskDataFinal.Text = Format(Date, "dd/mm/yyyy")
    cmbLojas.Text = "Todas"
    grdAjuste.Rows = 1
    grdAjuste.AddItem ""
    grdAjuste.RemoveItem (1)
    
End Sub

Private Sub cmdNovo_Click()

    frmCriaAjuste.Show 1

End Sub

Private Sub cmdPesquisa_Click()

    Pesquisa

End Sub

Private Sub cmdRetorna_Click()
    
    Unload Me
 
End Sub

Private Sub Form_Load()
    
    carregarPosicaoTamanhoTela Me
    carregaComboloja
    mskDataInicial.Text = Format(DateAdd("d", -1, Date), "dd/mm/yyyy")
    mskDataFinal.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub carregaComboloja()
    
    Dim rsLoja As New ADODB.Recordset
    Dim sql As String
    
    sql = "Exec sp_cdm_busca_ajuste 1, 0,0,0,0,0,0,0"
    
    rsLoja.CursorLocation = adUseClient
    rsLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    cmbLojas.AddItem "Todas"
    cmbLojas.Text = "Todas"
    
    If Not rsLoja.EOF Then
        Do While Not rsLoja.EOF
            
            cmbLojas.AddItem rsLoja("lo_loja")
            rsLoja.MoveNext
        
        Loop
        
    End If
    
    rsLoja.Close
    
End Sub

Private Sub Pesquisa()

    Dim sql As String
    Dim rsAjuste As New ADODB.Recordset
    Dim alteracao As String
    Dim status As String
    Dim Tipo As String
    
    'Limpa Grid Pesquisa
    grdAjuste.Rows = 1
    grdAjuste.AddItem ""
    grdAjuste.RemoveItem (1)
    
    If optSistema Then
        alteracao = "'S'"
    ElseIf optloja Then
        alteracao = "'L'"
    Else
        alteracao = "'A'"
    End If
    
    
    If optAberto Then
        status = "'A'"
    Else
        status = "'P'"
    End If
    
    
    If optAjuste Then
        Tipo = "'A'"
    ElseIf optContagem Then
        Tipo = "'C'"
    Else
        Tipo = "'I'"
    End If
    
    sql = "Exec sp_cdm_busca_ajuste 2,'" & Trim(cmbLojas.Text) & "', '" & Format(mskDataInicial.Text, "yyyy/mm/dd") & "', '" & Format(mskDataFinal.Text, "yyyy/mm/dd") & "'," & alteracao & "," & status & "," & Tipo & ", 0"
    rsAjuste.CursorLocation = adUseClient
    rsAjuste.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsAjuste.EOF Then
        Do While Not rsAjuste.EOF
        
            grdAjuste.AddItem rsAjuste("aj_codigoAjuste") & Chr(9) _
            & rsAjuste("aj_loja") & Chr(9) _
            & rsAjuste("aj_referencia") & Chr(9) _
            & rsAjuste("pr_descricao") & Chr(9) _
            & rsAjuste("aj_quantidade") & Chr(9) _
            & rsAjuste("aj_precoVenda") & Chr(9) _
            & rsAjuste("aj_descricaoMotivo") & Chr(9) _
            & Format(rsAjuste("aj_data"), "dd/mm/yyyy") & Chr(9) _
            & rsAjuste("aj_codigoUsuario") & Chr(9) _
            & rsAjuste("aj_Situacao")
        
            rsAjuste.MoveNext
        
        Loop
    Else
         MsgBox "Não foi localizado Registros para esta Consulta.", vbExclamation, "Atenção!"
    End If
    
    rsAjuste.Close
    
     
End Sub


Private Sub grdAjuste_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
    Dim numAjuste As String
    Dim Loja As String
    Dim referencia As String
    
    If KeyCode = 46 Then
        
        If MsgBox("Confirma a Exclusão deste Ajuste?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão de Produto") = vbYes Then
            
            numAjuste = Trim(grdAjuste.TextMatrix(grdAjuste.row, 0))
            Loja = Trim(grdAjuste.TextMatrix(grdAjuste.row, 1))
            referencia = Trim(grdAjuste.TextMatrix(grdAjuste.row, 2))
    
            sql = "Exec sp_cdm_deleta_Ajuste " & numAjuste & ", '" & Loja & "', '" & referencia & "'"
    
            ADO_Cn_CDLocal.Execute (sql)
            
            Pesquisa
            
        End If
        
    End If
End Sub

