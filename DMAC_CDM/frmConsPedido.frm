VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmConsProduto 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Consulta Vendas por Itens"
   ClientHeight    =   7470
   ClientLeft      =   2550
   ClientTop       =   1950
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   14895
      TabIndex        =   25
      Top             =   6600
      Width           =   14900
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3765
      TabIndex        =   12
      Top             =   765
      Width           =   1005
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItemNf 
      Height          =   5160
      Left            =   105
      TabIndex        =   16
      Top             =   1230
      Width           =   11685
      _cx             =   20611
      _cy             =   9102
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
      BackColorBkg    =   1347440720
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   -2147483632
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      BackColorFrozen =   0
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.TextBox txtVendedor 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4845
      MaxLength       =   3
      TabIndex        =   13
      Top             =   765
      Width           =   1305
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Moto Serra"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   10215
      TabIndex        =   8
      Top             =   270
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Substituição Tributária"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   8220
      TabIndex        =   7
      Top             =   270
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   90
      TabIndex        =   15
      Top             =   45
      Width           =   11685
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6135
         TabIndex        =   17
         Top             =   615
         Visible         =   0   'False
         Width           =   2760
         Begin VB.OptionButton Option3 
            BackColor       =   &H00404040&
            Caption         =   "Que Contenha"
            Enabled         =   0   'False
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   1
            Left            =   1335
            TabIndex        =   19
            Top             =   165
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Começe Por"
            Enabled         =   0   'False
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   18
            Top             =   165
            Visible         =   0   'False
            Width           =   1350
         End
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8955
         TabIndex        =   14
         Top             =   705
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Descrição"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   2535
         TabIndex        =   2
         Top             =   255
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Informações Contábeis"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   6150
         TabIndex        =   6
         Top             =   255
         Width           =   1920
      End
      Begin VB.ComboBox cmbLoja 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Refêrencia"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   1335
         TabIndex        =   1
         Top             =   255
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Classe"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   3645
         TabIndex        =   3
         Top             =   255
         Width           =   795
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Linha"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   5370
         TabIndex        =   5
         Top             =   255
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Bloq"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   4590
         TabIndex        =   4
         Top             =   255
         Width           =   705
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   10
         Top             =   720
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Index           =   1
         Left            =   2535
         TabIndex        =   11
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   165
         TabIndex        =   24
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1395
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2535
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Vendedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   4780
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   3690
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   11880
      TabIndex        =   26
      Top             =   6840
      Width           =   1530
      _ExtentX        =   2699
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
      MICON           =   "frmConsPedido.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmRetorna 
      Height          =   510
      Left            =   13440
      TabIndex        =   27
      Top             =   6840
      Width           =   1530
      _ExtentX        =   2699
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
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmConsPedido.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImpressao 
      Height          =   510
      Left            =   10200
      TabIndex        =   28
      Top             =   6840
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprime"
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
      MICON           =   "frmConsPedido.frx":0038
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
Attribute VB_Name = "frmConsProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private WithEvents clsItemNF As ControlaGrid
Attribute clsItemNF.VB_VarHelpID = -1


Dim rdoCombo As New ADODB.Recordset
Dim rdoDados As New ADODB.Recordset

Dim wVarLoja As String
Dim wVariavel As String
Dim LinhaTotal As String
Dim loja As String
Dim wVarVendedorTemp As String
Dim wTotalMer As Double
Dim wTotalGeral As Double
Dim wVarVendedorAnterior As String
Dim Serie As Long
Dim Linha As Long
Dim Pagina As Long
Dim NumItens As Long
Dim CodVendedor As Long
Dim I As Long
Dim NumItensGeral As Long
Dim NumItensNF As Long
Dim wVarVendedor As String
Dim wVarFornecedor As String
Dim wVarVende As String
Dim wTotalMerVende As Double
Dim NumItensVende As Long
Dim VlrMercadoriaMultiplicado As Long
Dim VlrMercadoriaMultiplicadoVende As Long
Dim wDesconto As Double
Dim wValorLiquido As Double
Dim wDescontoGeral As Double
Dim wValorLiquidoGeral As Double
Dim wDescontoLoja As Double
Dim wValorLiquidoLoja As Double
Dim NumItensLoja As Long
Dim wVlrMercadoria As Double
Dim wVlrMercadoriaLoja As Double
Dim wVarLojaTemp As String
Dim wVarLojaAnterior As String

Dim Wrow As Integer
Dim wRowTemp As Integer
Dim CountCor As Integer
Dim Cor As String


Private Sub cmbloja_GotFocus()
    
    mskData(0).Enabled = True
    mskData(1).Enabled = True

End Sub

Private Sub cmbLoja_LostFocus()
   
   If cmbLoja.Text = "" Then
      cmbLoja.ListIndex = 1
   End If

End Sub

Private Sub cmdImpressao_Click()
   Dim rdoVendaItens As rdoResultset
   Dim x As Integer
   Dim DescTotal As String
   
   Printer.Orientation = 2
   Printer.ScaleMode = vbMillimeters
   
   If Printer.PaperSize <> vbPRPSA4 Then
       Printer.ScaleLeft = -23
   End If

   If UCase(left(Printer.DeviceName, 7)) = "LEXMARK" Then
      Printer.Duplex = Duplex
   End If

   Linha = 55
   Pagina = 0
      
   For x = 1 To grdItemNf.Rows - 1
      If Linha >= 55 Then
         Cabecalho
         Linha = 7
      End If
      
      If Trim(left(grdItemNf.TextMatrix(x, 0) & Space(5), 5)) = "Total" Then
         
         If cmbLoja.Text = "CONSO" Then
            If Linha + 4 >= 55 Then
               Cabecalho
               Linha = 7
            End If
         End If
         If Linha + 2 >= 55 Then
            Cabecalho
            Linha = 7
         End If
         
         If grdItemNf.TextMatrix(x, 0) = "Total Vendedor" Then
            DescTotal = grdItemNf.TextMatrix(x - 1, 5)
         ElseIf grdItemNf.TextMatrix(x, 0) = "Total Loja" Then
                DescTotal = Trim(left(grdItemNf.TextMatrix(x - 3, 0) & Space(3), 3))
         ElseIf grdItemNf.TextMatrix(x, 0) = "Total Geral" Then
                DescTotal = ""
         End If
         
         Printer.Print ""
         Printer.FontBold = True
         Printer.Print Tab(6); grdItemNf.TextMatrix(x, 0); Tab(22); DescTotal; Tab(36); grdItemNf.TextMatrix(x, 2); Tab(48); grdItemNf.TextMatrix(x, 3); Tab(90); right(Space(6) & grdItemNf.TextMatrix(x, 4), 6); Tab(103); right(Space(9) & Format(grdItemNf.TextMatrix(x, 6), "##,###0.00"), 9); Tab(113); right(Space(9) & Format(grdItemNf.TextMatrix(x, 7), "##,###0.00"), 9); Tab(126); right(Space(9) & Format(grdItemNf.TextMatrix(x, 8), "##,###0.00"), 9); Tab(138); right(Space(3) & grdItemNf.TextMatrix(x, 5), 3); Tab(143); grdItemNf.TextMatrix(x, 9); Tab(146); grdItemNf.TextMatrix(x, 10); Tab(149); grdItemNf.TextMatrix(x, 11)
         Printer.FontBold = False
         
         loja = left(grdItemNf.TextMatrix(1, 0) & Space(3), 3)
         
         If cmbLoja.Text = "CONSO" Then
            Linha = Linha + 4
         Else
            Linha = Linha + 2
         End If
         
      Else
         If Linha + 2 >= 55 Then
            Cabecalho
            Linha = 7
         End If
         
         Printer.Print Tab(6); grdItemNf.TextMatrix(x, 0); Tab(23); grdItemNf.TextMatrix(x, 1); Tab(36); grdItemNf.TextMatrix(x, 2); Tab(48); grdItemNf.TextMatrix(x, 3); Tab(90); right(Space(6) & grdItemNf.TextMatrix(x, 4), 6); Tab(103); right(Space(9) & Format(grdItemNf.TextMatrix(x, 6), "##,###0.00"), 9); Tab(113); right(Space(9) & Format(grdItemNf.TextMatrix(x, 7), "##,###0.00"), 9); Tab(126); right(Space(9) & Format(grdItemNf.TextMatrix(x, 8), "##,###0.00"), 9); Tab(138); right(Space(3) & grdItemNf.TextMatrix(x, 5), 3); Tab(143); grdItemNf.TextMatrix(x, 9); Tab(146); grdItemNf.TextMatrix(x, 10); Tab(149); grdItemNf.TextMatrix(x, 11)
         Linha = Linha + 1
         
         
      End If
      
   Next x
   
   Printer.EndDoc
End Sub

Sub Cabecalho()
   
   Pagina = Pagina + 1
   
   If Pagina > 1 Then
      Printer.NewPage
   End If
   
   Printer.FontName = "ARIAL"
   Printer.FontBold = True
   Printer.FontSize = 9
   
   Printer.DrawWidth = 5
   Printer.CurrentX = 0
   Printer.CurrentY = 0
   
   Printer.Line (5, 10)-(307, 10)
   Printer.Line (5, 24)-(307, 24)
   Printer.Line (5, 192)-(307, 192)
   Printer.Line (5, 192)-(210, 192)
   
   Printer.CurrentY = 10
   Printer.Print Tab(8); NomeEmpresa; Tab(175); "PÁGINA: "; Pagina
   Printer.CurrentY = Printer.CurrentY + 2
   If Option1(6).Value = True Then
      If Option2(0).Value = True Then
         Printer.Print Tab(51); "CONSULTA  DE  VENDAS  DE  MOTO SERRA  NO  PERÍODO  DE  :  "; Format(mskData(0).Text, "DD/MM/YYYY"); "  À   "; Format(mskData(1).Text, "DD/MM/YYYY"); Tab(175); Format(Date, "DD/MM/YYYY")
      Else
         Printer.Print Tab(51); "CONSULTA  DE  VENDAS  POR  SUBST.TRIBUTÁRIA  NO  PERÍODO  DE  :  "; Format(mskData(0).Text, "DD/MM/YYYY"); "  À   "; Format(mskData(1).Text, "DD/MM/YYYY"); Tab(175); Format(Date, "DD/MM/YYYY")
      End If
   Else
      Printer.Print Tab(51); "CONSULTA  DE  VENDAS  POR  ITENS  NO  PERÍODO  DE  :  "; Format(mskData(0).Text, "DD/MM/YYYY"); "  À   "; Format(mskData(1).Text, "DD/MM/YYYY"); Tab(175); Format(Date, "DD/MM/YYYY")
   End If
   Printer.FontSize = 9
   Printer.FontName = "COURIER NEW"
   Printer.CurrentY = 28
   Printer.Print Tab(6); "NOTA FISCAL      DATA EMISSÃO REFERÊNCIA  DESCRIÇÃO                                   QTDE  VLR MERCADORIA  DESCONTO  VLR LÍQUIDO  VEND LN  CL  BL  "
   Printer.FontBold = False
   
   Printer.CurrentY = 33

End Sub

Private Sub cmdPesquisa_Click()
   
   Dim I As Long
   
   On Error GoTo Error
   
   wTotalMer = 0
   NumItens = 0
   wTotalGeral = 0
   NumItensGeral = 0
   wVarVendedor = ""
   wVarLoja = ""
   wVariavel = ""

   For I = 0 To 1
        mskData(2).Text = mskData(I).Text
       If InStr(mskData(I).Text, "_") = 0 Then
          If Not IsDate(mskData(I).Text) Then
             status "Pronto."
             mskData(I).SelStart = 0
             mskData(I).SelLength = Len(mskData(I).Text)
             mskData(I).SetFocus
             Exit Sub
          End If
       Else
          status "Pronto."
          mskData(I).SelStart = 0
          mskData(I).SelLength = Len(mskData(I).Text)
          mskData(I).SetFocus
          Exit Sub
       End If
   Next I
   
   If Format(mskData(0).Text, "mm/dd/yyyy") > Format(mskData(1).Text, "mm/dd/yyyy") Then
          mskData(0).SelStart = 0
          mskData(0).SelLength = Len(mskData(0).Text)
          mskData(0).SetFocus
          Exit Sub
   End If
   
   If cmbLoja.Text <> "CONSO" Then
      wVarLoja = "VC_LOJAVENDA = '" & cmbLoja.Text & "' and "
   Else
      wVarLoja = "VC_LOJAVENDA NOT IN ('CONSO') And "
   End If
   
   If txtVendedor.Visible = True Then
      If txtVendedor.Text = "999" Or Trim(txtVendedor.Text) = "" Then
         wVarVendedor = ""
      Else
         wVarVendedor = "VC_VendedorLojaVenda = " & txtVendedor.Text & " and "
      End If
   Else
      wVarVendedor = ""
   End If
   
   If Option1(0).Value = True Then
      wVariavel = "VI_Referencia = '" & Text1.Text & "' and "
   ElseIf Option1(1).Value = True Then
      If Text1.Text = "999" Then
         wVariavel = ""
      Else
         wVariavel = "VI_Referencia LIKE '" & Text1.Text & "%' and "
      End If
   ElseIf Option1(2).Value = True Then
      If Option3(0).Value = True Then
         wVariavel = "PR_Descricao LIKE '" & txtDesc.Text & "%' and "
      ElseIf Option3(1).Value = True Then
         wVariavel = "PR_Descricao LIKE '%" & txtDesc.Text & "%' and "
      End If
   ElseIf Option1(3).Value = True Then
      wVariavel = "PR_Classe = '" & txtDesc.Text & "' and "
   ElseIf Option1(4).Value = True Then
      wVariavel = "PR_Bloqueio = " & Text1.Text & " and "
   ElseIf Option1(5).Value = True Then
      wVariavel = "PR_Linha = " & Text1.Text & " and "
   ElseIf Option1(6).Value = True Then
      If Option2(0).Value = True Then
         wVariavel = "PR_Referencia IN (Select PR_Referencia From Produto " _
                   & "Where (PR_Descricao like 'moto serra%' or PR_Descricao like 'motosserra%') and " _
                   & "PR_Situacao='A' and PR_CodigoFornecedor <> 731) and "
      Else
         wVariavel = "(PR_ClasseFiscal Like '2710%' or " _
                   & "PR_ClasseFiscal Like '3810%') and "
      End If
   End If
    
    Screen.MousePointer = 11
    grdItemNf.Rows = 1

    loja = ""
    
    sql = "Select (VI_VALORMERCADORIA - VI_Desconto) as ValorLiquido,(VC_LOJAVENDA + ' ' + convert(varchar(10),VC_NOTAFISCAL) + ' ' + VC_SERIE) as wdocumento, " _
                  & "VC_DATAEMISSAO,VC_TotalNota,VI_Desconto, VI_Referencia, PR_Descricao, " _
                  & "VI_Quantidade, VI_PrecoUnitario,  VC_VENDEDORLOJAVENDA,VE_Nome, " _
                  & "PR_LINHA, PR_CLASSE, PR_BLOQUEIO, VI_VALORMERCADORIA, " _
                  & "(Convert(Char(6),VC_Cliente) + ' - ' + VC_NomeCliente) as VC_NomeCliente,VC_TelefoneCliente, " _
                  & "VC_LojaVenda from ItemNFVenda, Produto, CAPANFVENDA,Vendedor " _
                  & "where PR_Referencia = VI_Referencia and " _
                  & "VI_NOTAFISCAL = VC_NOTAFISCAL AND " _
                  & "VI_SERIE = VC_SERIE AND " _
                  & "VI_LOJAORIGEM = VC_LOJAORIGEM AND " _
                  & "VC_TIPONOTA = 'V' and " _
                  & "vc_tiponota <> 'C' and " _
                  & wVarLoja _
                  & wVariavel _
                  & wVarVendedor _
                  & "VC_DATAEMISSAO Between '" & Format(mskData(0).Text, "mm/dd/yyyy") & "' and " _
                  & "'" & Format(mskData(1).Text, "mm/dd/yyyy") & "' and VC_VendedorLojaVenda = VE_CodigoVendedor and VE_Loja = VC_LojaVenda " _
                  & "ORDER BY VC_LOJAVENDA, VC_VENDEDORLOJAVENDA, VC_DATAEMISSAO, VC_NOTAFISCAL, VC_SERIE"
    
    
        rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
        wVarVendedorTemp = "000"
        wVarLojaTemp = "000"
        CountCor = 0
        wRowTemp = 1
    If Not rdoDados.EOF Then
        wVarVendedorAnterior = rdoDados("VC_VendedorLojaVenda")
        wVarLojaAnterior = rdoDados("VC_LojaVenda")
        Do While Not rdoDados.EOF
            If wVarVendedorAnterior <> rdoDados("VC_VendedorLojaVenda") Then
                If CountCor = 0 Then
                    CountCor = 1
                Else
                    CountCor = 0
                End If
                wRowTemp = grdItemNf.Rows + 2
                grdItemNf.AddItem "Total Vendedor" & vbTab & vbTab & vbTab & vbTab & NumItensVende & vbTab & vbTab & wTotalMerVende & vbTab & wDesconto & vbTab & wValorLiquido
                grdItemNf.AddItem ""
                NumItensVende = "0"
                wTotalMerVende = "0"
                wDesconto = "0"
                wValorLiquido = "0"
            End If
            If wVarLojaAnterior <> rdoDados("VC_LojaVenda") Then
                '----------------------------------------------------------
                If wVarVendedorAnterior = rdoDados("VC_VendedorLojaVenda") Then
                   If CountCor = 0 Then
                      CountCor = 1
                   Else
                      CountCor = 0
                   End If
                   wRowTemp = grdItemNf.Rows + 2
                   grdItemNf.AddItem "Total Vendedor" & vbTab & vbTab & vbTab & vbTab & NumItensVende & vbTab & vbTab & wTotalMerVende & vbTab & wDesconto & vbTab & wValorLiquido
                   grdItemNf.AddItem ""
                   NumItensVende = "0"
                   wTotalMerVende = "0"
                   wDesconto = "0"
                   wValorLiquido = "0"
                End If
                '----------------------------------------------------------
                wVarLojaTemp = "000"
                wRowTemp = grdItemNf.Rows + 2
                grdItemNf.AddItem "Total Loja" & vbTab & vbTab & vbTab & vbTab & NumItensLoja & vbTab & vbTab & wVlrMercadoriaLoja & vbTab & wDescontoLoja & vbTab & wValorLiquidoLoja
                grdItemNf.AddItem ""
                wDescontoLoja = "0"
                wValorLiquidoLoja = "0"
                NumItensLoja = "0"
                wVlrMercadoriaLoja = "0"
            End If
                
                grdItemNf.AddItem rdoDados("wDocumento") & vbTab & rdoDados("VC_DATAEMISSAO") & vbTab & rdoDados("VI_REFERENCIA") & vbTab & rdoDados("PR_DESCRICAO") & vbTab & rdoDados("VI_QUANTIDADE") _
                                  & vbTab & rdoDados("VC_VENDEDORLOJAVENDA") & vbTab & rdoDados("VI_VALORMERCADORIA") & vbTab & rdoDados("VI_DESCONTO") & vbTab & rdoDados("ValorLiquido") & vbTab & rdoDados("PR_LINHA") _
                                  & vbTab & rdoDados("PR_CLASSE") & vbTab & rdoDados("PR_Bloqueio") & vbTab & rdoDados("VE_Nome") & vbTab & rdoDados("VC_NomeCliente") & vbTab & rdoDados("VC_TelefoneCliente")
                Call PintaGrid
    
    '----------------------------------Soma Geral-----------------------------------------'
                wTotalMer = wTotalMer + rdoDados("VI_VALORMERCADORIA")
                NumItensNF = NumItensNF + rdoDados("VI_Quantidade")
    '-------------------------------------------------------------------------------------'
    
    '----------------------------------Soma VLRMERCADORIA Por Vendedor--------------------'
                wTotalMerVende = wTotalMerVende + rdoDados("VI_VALORMERCADORIA")
                NumItensVende = NumItensVende + rdoDados("VI_Quantidade")
                wVlrMercadoriaLoja = wVlrMercadoriaLoja + rdoDados("VI_VALORMERCADORIA")
                NumItensLoja = NumItensLoja + rdoDados("VI_Quantidade")
    '-------------------------------------------------------------------------------------'
    
    '----------------------------------Soma DESCONTO--------------------------------------'
                wDesconto = wDesconto + rdoDados("VI_Desconto")
                wDescontoGeral = wDescontoGeral + rdoDados("VI_Desconto")
                wDescontoLoja = wDescontoLoja + rdoDados("VI_Desconto")
    '-------------------------------------------------------------------------------------'
    
    '----------------------------------Soma VALORLIQUIDO----------------------------------'
                wValorLiquido = wValorLiquido + rdoDados("ValorLiquido")
                wValorLiquidoGeral = wValorLiquidoGeral + rdoDados("ValorLiquido")
                wValorLiquidoLoja = wValorLiquidoLoja + rdoDados("ValorLiquido")
    '-------------------------------------------------------------------------------------'
                
                wVarVendedorAnterior = rdoDados("VC_VENDEDORLOJAVENDA")
                wVarLojaAnterior = rdoDados("VC_LojaVenda")
    
                rdoDados.MoveNext
        Loop
    Else
        MsgBox "Nenhum registro encontrado para este filtro.", vbInformation, "Consulta Produto"
        Screen.MousePointer = 0
        'status "Pronto"
        Exit Sub
    End If
        
    If wVarVendedorAnterior <> "000" Then
        Call PintaGrid
        grdItemNf.AddItem "Total Vendedor" & vbTab & vbTab & vbTab & vbTab & NumItensVende & vbTab & vbTab & wTotalMerVende & vbTab & wDesconto & vbTab & wValorLiquido
        grdItemNf.AddItem ""
        grdItemNf.AddItem "Total Loja" & vbTab & vbTab & vbTab & vbTab & NumItensLoja & vbTab & vbTab & wVlrMercadoriaLoja & vbTab & wDescontoLoja & vbTab & wValorLiquidoLoja
        grdItemNf.AddItem ""
        grdItemNf.AddItem "Total Geral" & vbTab & vbTab & vbTab & vbTab & NumItensNF & vbTab & vbTab & wTotalMer & vbTab & wDescontoGeral & vbTab & wValorLiquidoGeral
    End If
    wVarVendedorTemp = "000"
    wTotalMerVende = 0
    NumItensVende = 0
    wTotalMer = 0
    NumItensNF = 0
    wDesconto = 0
    wValorLiquido = 0
    wDescontoGeral = 0
    wValorLiquidoGeral = 0
    NumItensLoja = 0
    wVlrMercadoriaLoja = 0
    wDescontoLoja = 0
    wValorLiquidoLoja = 0
    
   Screen.MousePointer = 0
   
   'status "Pronto"
   
   grdItemNf.SetFocus

Exit Sub

Error:
    
    Screen.MousePointer = 0
    
   ' status "Pronto"

End Sub

Private Sub cmRetorna_Click()
   
   Unload Me

End Sub

Private Sub Form_Load()

  ' Skin1.LoadSkin "C:\windows\system\skin.skn"
  ' Skin1.ApplySkin Me.hwnd
   ' frmEtiquetaPreco.Top = ((Screen.Height - frmEtiquetaPreco.Height) / 2 + 350)
  '  frmEtiquetaPreco.Left = (Screen.Width - frmEtiquetaPreco.Width) / 2

    
    frmTransferenciaEntradaCD.top = (Screen.Height - frmTransferenciaEntradaCD.Height) / 2
   frmTransferenciaEntradaCD.left = (Screen.Width - frmTransferenciaEntradaCD.Width) / 2
   
    carregarPosicaoTamanhoTela Me
 ' ConectaODBC Conexao, NomeUsuario, SenhaUsuario

   
   
   
    sql = "Select LO_Loja from Loja Where LO_Loja not in('ALM01', 'CMC') and LO_Situacao='A'"
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
   PreencheCombo cmbLoja, rdoCombo, "LO_Loja", ""
      
   Set clsItemNF = New ControlaGrid
   Set clsItemNF.ConexaoGrid = ADO_Cn_CDLocal
   clsItemNF.NomeGrid = "grdItemNF"
'----------------------------------------------------Grid Antigo---------------------'
'   clsItemNF.NomeFormulario = Me.Name
'
'   clsItemNF.Colunas = 12
'   clsItemNF.LinhasVisiveis = 5

'   clsItemNF.cabecalho = "<Nota Fiscal; <Dt. Emissão; <Referência; " _
                       & "<Descrição; >Qtd.; >Pço.Unitário; >Vend; " _
                       & ">Desconto; >Vr.Mercadoria; >Linha; >Classe; >Bloqueio"

'   clsItemNF.Campos = "wdocumento; VC_DATAEMISSAO; VI_Referencia; PR_Descricao; " _
                    & "VI_Quantidade; VI_PrecoUnitario; VC_VENDEDORLOJAVENDA; " _
                    & "VC_Desconto;VI_VALORMERCADORIA;PR_LINHA; PR_CLASSE; PR_BLOQUEIO;"
   
'   clsItemNF.Alinhamento = "Esquerda; Esquerda; Esquerda; Esquerda; Direita; Esquerda; Direita; Direita; Esquerda; Esquerda; Direita"
   
'   clsItemNF.Formato = "Caractere; Data; Caractere; Caractere; Numero; Decimal; Caractere; Decimal; Caractere; Caractere; Numero"
   
'   clsItemNF.Tamanho = "1350; 1050; 1100; 3600; 500; 1100; 500; 800; 1100; 500; 550; 700"

'   clsItemNF.MontaCabecalho
'----------------------------------------------------Grid Antigo---------------------'
   
   rdoCombo.Close
   
   'wVarVendedorTemp = "000"
   
   Text1.Visible = True
   
   For I = 0 To 1
      Option2(I).Visible = False
   Next I
   
End Sub

Private Sub mskData_GotFocus(Index As Integer)
      
      mskData(Index).SelStart = 0
      mskData(Index).SelLength = Len(mskData(Index).Text)

End Sub

Private Sub mskData_LostFocus(Index As Integer)
   
    If IsDate(mskData(Index)) = True Then
        If Index = 0 Then
           mskData(1).Text = mskData(0).Text
        Else
           If Option1(2).Value = True Then
              txtVendedor.Enabled = True
              txtVendedor.SetFocus
           ElseIf Option1(3).Value = True Or Option1(4).Value = True Then
              txtDesc.Enabled = True
              txtDesc.SetFocus
           ElseIf Option1(6).Value = False Then
              Text1.Enabled = True
              Text1.SetFocus
           End If
        End If
    Else
        MsgBox "Data Inválida."
        mskData(Index).SelStart = 0
        mskData(Index).SelLength = Len(mskData(Index).Text)
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
     
     If Option1(0).Value = True Then
        Text1.Width = 1000
        txtVendedor.Visible = True
        Label1.Caption = "Referência"
        Label2.Visible = True
        Label1.Visible = True
        Option3(0).Visible = False
        Option3(1).Visible = False
        Frame2.Visible = False
        Label1.left = "3685"
        txtDesc.Visible = False
        Text1.Enabled = True
        Text1.Visible = True
       ' Label1.ForeColor = &HFF0000
     ElseIf Option1(1).Value = True Then
        Text1.Width = 1000
        txtVendedor.Visible = True
        Label1.Caption = "Fornecedor"
        Option3(0).Visible = False
        Option3(1).Visible = False
       Frame2.Visible = False
        Label2.Visible = True
        Label1.Visible = True
       ' Label1.ForeColor = &HFF0000
        Label1.left = "3685"
        txtDesc.Visible = False
        Text1.Enabled = True
        Text1.Visible = True
     ElseIf Option1(2).Value = True Then
        Text1.Width = 1000
        txtVendedor.Visible = True
      '  Label1.ForeColor = &HFF0000
        Text1.Enabled = False
        Label2.Visible = True
        Label1.Visible = True
        Option3(0).Visible = True
        Option3(1).Visible = True
       Frame2.Visible = True
        Label1.left = "8955"
        Label1.Caption = "Descrição"
        txtDesc.Visible = True
        txtDesc.left = 8970
        txtDesc.Width = 2295
     ElseIf Option1(3).Value = True Then
        txtDesc.Width = 2100
        txtDesc.Visible = True
        txtDesc.Enabled = True
        txtDesc.left = 3670
        txtVendedor.Visible = False
        Label1.left = "3685"
        Label1.Caption = "Classe"
       ' Label1.ForeColor = &HFF0000
        Label2.Visible = False
        Label1.Visible = True
        Option3(0).Visible = False
        Option3(1).Visible = False
        Frame2.Visible = False
        Text1.Enabled = True
'        Label1.Left = "3685"
        Text1.Visible = False
     ElseIf Option1(4).Value = True Then
        Text1.Width = 2000
        txtVendedor.Visible = False
        'Label1.ForeColor = &HFF0000
        Label1.left = "3685"
        Label1.Caption = "Bloqueio"
        Label2.Visible = False
        Option3(0).Visible = False
        Option3(1).Visible = False
        Frame2.Visible = False
        Label1.Visible = True
        txtDesc.Visible = False
'        Label1.Left = "3685"
        Text1.Visible = True
     ElseIf Option1(5).Value = True Then
        Option3(0).Visible = False
        Text1.Enabled = True
        Option3(1).Visible = False
        Frame2.Visible = False
        Label1.left = "3685"
        Label1.Caption = "Linha"
        Label2.Visible = False
        Label1.Visible = True
        Option3(0).Visible = False
        Option3(1).Visible = False
        Frame2.Visible = False
      '  Label1.ForeColor = &HFF0000
'        Label1.Left = "3685"
        Text1.Visible = True
        txtVendedor.Visible = False
        txtDesc.Visible = False
     ElseIf Option1(6).Value = True Then
        Text1.Width = 2000
        txtVendedor.Visible = False
        Label1.Visible = False
        Option3(0).Visible = False
       ' Label1.ForeColor = &HFF0000
        Option3(1).Visible = False
        Frame2.Visible = False
        Text1.Enabled = True
        Label2.Visible = False
        txtDesc.Visible = False
        Option3(0).Visible = False
        Option3(1).Visible = False
        Frame2.Visible = False
        Text1.Visible = False
        For I = 0 To 1
            Option2(I).Visible = True
        Next I
     End If
     If Option1(6).Value = False Then
        For I = 0 To 1
            Option2(I).Visible = False
        Next I
     End If
     cmbLoja.Enabled = True
     txtVendedor.Text = ""
     Text1.Text = ""
     txtDesc.Text = ""
     
End Sub

Private Sub Option3_Click(Index As Integer)

    txtDesc.SetFocus

End Sub

Private Sub Option3_GotFocus(Index As Integer)

    txtDesc.Enabled = True

End Sub

Private Sub Option3_LostFocus(Index As Integer)

    txtDesc.SetFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    VerTecla KeyAscii

End Sub

Private Sub Text1_GotFocus()
      
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1.Text)
      
      txtVendedor.Enabled = True
      
      If Option1(0).Value = True Then
         Text1.MaxLength = 7
      ElseIf Option1(1).Value = True Then
         Text1.MaxLength = 3
      ElseIf Option1(2).Value = True Then
         Text1.MaxLength = 15
      ElseIf Option1(3).Value = True Then
         Text1.MaxLength = 2
      ElseIf Option1(4).Value = True Then
         Text1.MaxLength = 1
      ElseIf Option1(5).Value = True Then
         Text1.MaxLength = 1
      End If

End Sub

Private Sub Text1_LostFocus()
   
'   If Len(Text1.Text) >= "4" Then
        
   
   If Len(Text1.Text) <> 0 Then
      If Option1(0).Value = True Then
    sql = "Select PR_REFERENCIA from PRODUTO WHERE PR_REFERENCIA = '" & Text1.Text & "'"
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      ElseIf Option1(1).Value = True Then
            sql = "Select FO_CODIGOFORNECEDOR from FORNECEDOR WHERE FO_CODIGOFORNECEDOR = " & Text1.Text
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      
      ElseIf Option1(2).Value = True Then
                    sql = "Select PR_Descricao from Produto WHERE PR_Descricao like '%" & txtDesc.Text & "%'"
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      ElseIf Option1(3).Value = True Then
                    sql = "Select CL_CODIGOCLASSE from CLASSE WHERE CL_CODIGOCLASSE = '" & Text1.Text & "'"
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      ElseIf Option1(4).Value = True Then
                    sql = "Select BL_CODIGOBLOQUEIO from BLOQUEIO WHERE BL_CODIGOBLOQUEIO = '" & Text1.Text & "'"
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      ElseIf Option1(5).Value = True Then
                    sql = "Select LI_CODIGOLINHA from LINHA WHERE LI_CODIGOLINHA = " & Text1.Text
    rdoCombo.CursorLocation = adUseServer
    rdoCombo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      End If
      
      If rdoCombo.EOF Then
         MsgBox "Chave de Pesquisa Informada Não Existe.", vbCritical, "Atenção"
         Text1.SetFocus
      ElseIf txtVendedor.Visible = True Then
         txtVendedor.Enabled = True
         txtVendedor.SetFocus
      End If
      rdoCombo.Close
   ElseIf Text1.Text = "" Then
      MsgBox "Preenchimento Obrigátorio.", vbCritical, "Atenção"
      Text1.SetFocus
   End If

End Sub

Sub PintaGrid()

    If CountCor = 0 Then
        Cor = &HB9CFFF
    Else
        Cor = &HDDFFFF
    End If
    
    Wrow = grdItemNf.Rows - 1
    grdItemNf.Select wRowTemp, 0, Wrow, 14
    grdItemNf.FillStyle = flexFillRepeat
    grdItemNf.CellBackColor = Cor
    grdItemNf.FillStyle = flexFillSingle
    grdItemNf.Select 0, 14
    
End Sub

Private Sub txtDesc_LostFocus()

    If txtDesc.Text = "" Then
        MsgBox "Campo obrigatório."
        txtDesc.SetFocus
    End If
    
End Sub

Private Sub txtVendedor_GotFocus()

    If Option1(2).Value = True Then
        Option3(0).Visible = True
        Option3(1).Visible = True
        Frame2.Visible = True
    End If

End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)

    VerTecla KeyAscii

End Sub

Private Sub txtVendedor_LostFocus()
    
    If Option1(2).Value = True And txtVendedor.Text <> "" Then
        Option3(0).Enabled = True
        Option3(1).Enabled = True
        Frame2.Visible = True
    End If

End Sub
