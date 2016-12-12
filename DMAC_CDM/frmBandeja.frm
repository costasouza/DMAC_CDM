VERSION 5.00
Begin VB.Form frmBandeja 
   BackColor       =   &H00000000&
   Caption         =   "DMAC CDM"
   ClientHeight    =   10125
   ClientLeft      =   2640
   ClientTop       =   555
   ClientWidth     =   15090
   Icon            =   "frmBandeja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   18493.15
   ScaleMode       =   0  'User
   ScaleWidth      =   15090
   Begin VB.Image imgTarefas 
      Height          =   11520
      Left            =   -465
      Picture         =   "frmBandeja.frx":23FA
      Top             =   240
      Width           =   15360
   End
End
Attribute VB_Name = "frmBandeja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    frmControleCD.Show 1
End Sub

Private Sub Form_Load()
    imgTarefas.top = 0
    imgTarefas.left = 0
    Me.Height = (imgTarefas.Height) - 100
    Me.Width = (imgTarefas.Width)
    Me.top = -500
    Me.left = -100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

