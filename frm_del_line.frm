VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{096AE495-34E9-4DD8-B744-B72D16ADFB8C}#1.0#0"; "button.ocx"
Begin VB.Form frm_del_line 
   Caption         =   "Produtos"
   ClientHeight    =   8340
   ClientLeft      =   6720
   ClientTop       =   1620
   ClientWidth     =   12930
   Icon            =   "frm_del_line.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   12930
   Begin Dacara_dcButton.dcButton cmd_remove_deselecionado 
      Height          =   585
      Left            =   240
      TabIndex        =   1
      Top             =   7380
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   1032
      BackColor       =   12648447
      ButtonStyle     =   7
      Caption         =   "Apagar os não selecionados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicNormal       =   "frm_del_line.frx":0442
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6945
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   12250
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frm_del_line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_remove_deselecionado_Click()
        Dim linha As ListItem
        Dim nIni As Integer
        Dim nFim As Integer
        Dim nPos As Integer
        
        nIni = 1
        nFim = ListView1.ListItems.Count
        nPos = 1
        
        Do While nIni <= nFim And nPos <= nFim
        
           Set linha = ListView1.ListItems(nPos)
           
           If linha.Checked = False Then
              ListView1.ListItems.Remove linha.Index
              
              nFim = nFim - 1
              
           Else
                nPos = nPos + 1
           End If
        
        Loop
End Sub

Private Sub Form_Load()
        Dim linha As ListItem
        Dim n As Integer
        With ListView1
        
             .Checkboxes = True
             .View = lvwReport
             
             .ColumnHeaders.Clear
             .ListItems.Clear
             
             .ColumnHeaders.Add , , "Código", 1800
             .ColumnHeaders.Add , , "Descrição", 2500, lvwColumnLeft
        
        
             For n = 1 To 15
                 Set linha = .ListItems.Add(, , Right("0000" & n, 7))
                 
                 linha.SubItems(1) = "Descrição " & Right("0000" & n, 7)
             
             Next
        End With
End Sub
