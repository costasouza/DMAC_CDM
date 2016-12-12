VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmModelarMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Modelar Menu Principal"
   ClientHeight    =   7890
   ClientLeft      =   2190
   ClientTop       =   2205
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   15585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6540
      Left            =   4560
      TabIndex        =   12
      Top             =   150
      Width           =   10470
      Begin VSFlex7DAOCtl.VSFlexGrid grdReferenciaSemBarras 
         Height          =   5805
         Left            =   150
         TabIndex        =   16
         Top             =   555
         Width           =   8700
         _cx             =   15346
         _cy             =   10239
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         BackColorBkg    =   4210752
         BackColorAlternate=   3947580
         GridColor       =   5263440
         GridColorFixed  =   8421504
         TreeColor       =   3947580
         FloodColor      =   5263440
         SheetBorder     =   3947580
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmModelarMenu.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         Editable        =   1
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
      Begin CentroDeDistribuicao.chameleonButton cmdDesce 
         Height          =   510
         Left            =   8955
         TabIndex        =   18
         Top             =   3330
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   900
         BTYPE           =   14
         TX              =   "\/"
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
         MICON           =   "frmModelarMenu.frx":004A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdSobe 
         Height          =   510
         Left            =   8955
         TabIndex        =   19
         Top             =   2730
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   900
         BTYPE           =   14
         TX              =   "/\"
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
         MICON           =   "frmModelarMenu.frx":0066
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Editar Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   200
         Left            =   150
         TabIndex        =   13
         Top             =   150
         Width           =   2655
      End
   End
   Begin VB.Frame frameOpcoesUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6540
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   4245
      Begin VB.TextBox txtFormulario 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   150
         LinkTimeout     =   0
         MaxLength       =   50
         TabIndex        =   17
         ToolTipText     =   "Formulario"
         Top             =   2340
         Width           =   3930
      End
      Begin VB.ComboBox cmbMenu 
         BackColor       =   &H00A3A3A3&
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Text            =   "000000 - Raiz"
         Top             =   3240
         Width           =   3930
      End
      Begin VB.TextBox txtNomeMenu 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   135
         MaxLength       =   100
         TabIndex        =   9
         ToolTipText     =   "Descrição Menu"
         Top             =   765
         Width           =   3930
      End
      Begin VB.OptionButton optAplicativo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Aplicativo"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1620
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optMenuAplicativo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Menu de Aplicativo"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   1620
         Width           =   2055
      End
      Begin CentroDeDistribuicao.chameleonButton cmdGravar 
         Height          =   420
         Left            =   135
         TabIndex        =   5
         Top             =   5910
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   741
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
         BCOLO           =   0
         FCOL            =   16423203
         FCOLO           =   16423203
         MCOL            =   5263440
         MPTR            =   1
         MICON           =   "frmModelarMenu.frx":0082
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdLimparCamposNovoMenu 
         Height          =   420
         Left            =   2175
         TabIndex        =   6
         Top             =   5910
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   741
         BTYPE           =   14
         TX              =   "Limpar campos"
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
         MICON           =   "frmModelarMenu.frx":009E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   2985
         Width           =   2235
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   150
         TabIndex        =   14
         Top             =   2085
         Width           =   2235
      End
      Begin VB.Label lblDescricaoFornecedor 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição Menu / Aplicativo"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   525
         Width           =   2235
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Adicionar Novo Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   200
         Left            =   150
         TabIndex        =   7
         Top             =   150
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   1
      Top             =   6810
      Width           =   14880
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13620
      TabIndex        =   0
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
      MICON           =   "frmModelarMenu.frx":00BA
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
Attribute VB_Name = "frmModelarMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim menuatual As Integer
Dim Descricao As String
Dim desc As String
Dim descAnte As String

Private Sub cmdRetorna_Click()

    Unload Me
    
End Sub


Private Sub Form_Load()

    carregarPosicaoTamanhoTela Me
    carregaMenu
    
End Sub


Private Function carregaMenu()

    Dim adoMenuGrid As New ADODB.Recordset
    Dim sql As String
    Dim codigo As String
    
    sql = "select * from GLB_MenuSistema where MSI_nomeForm like 'MENU%' order by MSI_Codigo"
    
    adoMenuGrid.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoMenuGrid.EOF Then
        limpaGrid
        Do While Not adoMenuGrid.EOF
                
            grdReferenciaSemBarras.AddItem adoMenuGrid("msi_codigo") & Chr(9) & adoMenuGrid("msi_descricao")
            adoMenuGrid.MoveNext
            
        Loop
        
    End If
    
    adoMenuGrid.Close
    menuatual = 1

End Function

Private Function carregasubMenu(codPai As String)
    
    Dim adoMenuGrid As New ADODB.Recordset
    Dim sql As String
    Dim codigo As String
    
    sql = "select * from GLB_MenuSistema where MSI_nomeForm not like 'MENU%' and MSI_Codigo like '" & codPai & "%' order by MSI_Codigo "
    adoMenuGrid.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoMenuGrid.EOF Then
        limpaGrid
        Do While Not adoMenuGrid.EOF
        
            grdReferenciaSemBarras.AddItem adoMenuGrid("msi_codigo") & Chr(9) & adoMenuGrid("msi_descricao")
            adoMenuGrid.MoveNext
            
        Loop
        
    End If
    
    adoMenuGrid.Close
    menuatual = 2

End Function


Private Sub grdReferenciaSemBarras_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
    
    descAnte = grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 1)
    
End Sub

Private Sub grdReferenciaSemBarras_DblClick()
    Dim codPai As String
    
    If menuatual = 1 Then
        codPai = Mid(grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 0), 1, 2)
        carregasubMenu codPai
    ElseIf menuatual = 2 Then
        carregaMenu
    End If
    
End Sub

Private Function limpaGrid()
    
    grdReferenciaSemBarras.Rows = 1
    grdReferenciaSemBarras.AddItem ""
    grdReferenciaSemBarras.RemoveItem (1)
    
End Function



Private Function MoveMenu(menu As Integer, move As Integer)

    ' Menu = 1  Carrega lista de menú raiz
    ' Menu = 2  Carrega lista de subMenus
    ' move = -1 Move referência para baixo
    ' move = 1  Move referência para cima
    
    Dim sql As String
    Dim ref1 As String
    Dim ref2 As String
    Dim codPai As String
    Dim Linha As Integer
    
    Linha = grdReferenciaSemBarras.row
    ref1 = grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 0)
    ref2 = grdReferenciaSemBarras.TextMatrix((grdReferenciaSemBarras.row + move), 0)
    
    If ref1 <> "000000" And ref2 <> "000000" Then
    
        sql = "update glb_menusistema set msi_codigo = 'Troca' where msi_codigo = '" & ref1 & " '"
        ADO_Cn_CDLocal.Execute (sql)
    
        sql = "update glb_menusistema set msi_codigo = '" & ref1 & "' where msi_codigo = '" & ref2 & " '"
        ADO_Cn_CDLocal.Execute (sql)
    
        sql = "update glb_menusistema set msi_codigo = '" & ref2 & "' where msi_codigo = 'Troca'"
        ADO_Cn_CDLocal.Execute (sql)
    
        limpaGrid
    
        If menu = 1 Then
            carregaMenu
        ElseIf menu = 2 Then
            codPai = Mid(ref1, 1, 2)
            carregasubMenu codPai
        End If
    
        ' Seleciona linha De novo
        grdReferenciaSemBarras.row = Linha + move
        
    End If
End Function

Private Sub cmdSobe_Click()
    MoveMenu menuatual, -1
End Sub

Private Sub cmdDesce_Click()
    MoveMenu menuatual, 1
End Sub

Private Sub grdReferenciaSemBarras_AfterEdit(ByVal row As Long, ByVal col As Long)
    
    If col = 1 Then
    
        desc = Trim(grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 1))
        
        If MsgBox("Confirma a Alteração da Descrição do Menu?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão de Produto") = vbYes Then
            gravaDesc desc
        Else
            grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 1) = descAnte
        End If
        
    End If
    
End Sub
Private Sub gravaDesc(desc As String)
    
    Dim sql As String
    Dim codigo As String
    
    codigo = Trim(grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 0))
    sql = "update glb_menusistema set msi_descricao = '" & desc & "' where msi_codigo = '" & codigo & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
End Sub



