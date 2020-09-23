VERSION 5.00
Begin VB.Form fTesteCbo 
   Caption         =   "JZ XP Combo Demo/Tutor/Test"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt5 
      Caption         =   "Copy"
      Height          =   255
      Left            =   3435
      TabIndex        =   34
      ToolTipText     =   "Test it Set/Reset TagCode Option above"
      Top             =   1620
      Width           =   1095
   End
   Begin VB.ComboBox NCB 
      Height          =   315
      Left            =   2235
      TabIndex        =   33
      Text            =   "Normal Combo Box"
      Top             =   1905
      Width           =   2340
   End
   Begin VB.CommandButton Bt6 
      Caption         =   "Append"
      Height          =   240
      Left            =   2640
      TabIndex        =   32
      ToolTipText     =   "Add Work Text as last item in Combo List"
      Top             =   3390
      Width           =   900
   End
   Begin VB.CommandButton Bt7 
      Caption         =   "Upd"
      Height          =   240
      Left            =   3090
      TabIndex        =   31
      ToolTipText     =   "Updates actual list =  Work Text"
      Top             =   3135
      Width           =   450
   End
   Begin VB.CommandButton Bt8 
      Caption         =   "Remove actual item"
      Height          =   495
      Left            =   3525
      TabIndex        =   30
      ToolTipText     =   "Eliminates the actual list item"
      Top             =   3135
      Width           =   1005
   End
   Begin VB.CommandButton Bt9 
      Caption         =   "Ins"
      Height          =   240
      Left            =   2625
      TabIndex        =   29
      ToolTipText     =   "Inserts Work Text BEFORE actual list"
      Top             =   3135
      Width           =   450
   End
   Begin VB.TextBox Tx9 
      Height          =   315
      Left            =   135
      TabIndex        =   27
      ToolTipText     =   "Simulate a Combo Entry string"
      Top             =   3270
      Width           =   2460
   End
   Begin VB.TextBox TxI 
      Height          =   315
      Left            =   1845
      TabIndex        =   25
      Text            =   "00"
      ToolTipText     =   "When Code has duplicities to Search (e.g.,Celebrities)"
      Top             =   5100
      Width           =   330
   End
   Begin VB.Frame Frame 
      Caption         =   "Load these Sample Texts"
      Height          =   840
      Left            =   2475
      TabIndex        =   20
      Top             =   4590
      Width           =   2070
      Begin VB.OptionButton Op3 
         Caption         =   "Colors"
         Height          =   195
         Left            =   1140
         TabIndex        =   24
         Top             =   255
         Width           =   750
      End
      Begin VB.OptionButton Op4 
         Caption         =   "My Test"
         Height          =   195
         Left            =   1140
         TabIndex        =   23
         Top             =   510
         Width           =   870
      End
      Begin VB.OptionButton Op2 
         Caption         =   "Countries"
         Height          =   195
         Left            =   75
         TabIndex        =   22
         Top             =   495
         Width           =   960
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Celebrities"
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.CommandButton Cmd5 
      Caption         =   "5 - Save Combo to  File"
      Height          =   315
      Left            =   2490
      TabIndex        =   19
      ToolTipText     =   "Requires a File Name expressed Below "
      Top             =   3765
      Width           =   2055
   End
   Begin VB.CommandButton Cmd6 
      Caption         =   "6 - Reload Combo from File"
      Height          =   315
      Left            =   2490
      TabIndex        =   18
      ToolTipText     =   "Requires a File Name expressed Below previos saved"
      Top             =   4155
      Width           =   2055
   End
   Begin VB.TextBox Tx4 
      Height          =   315
      Left            =   135
      TabIndex        =   16
      Top             =   5715
      Width           =   4410
   End
   Begin VB.CommandButton Cmd4 
      Caption         =   "4 - Set Combo List to a TagCode"
      Height          =   450
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Enter a Data TagCode (See Msg)"
      Top             =   4965
      Width           =   1710
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "3 - Set Combo List to Index"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Try modif Actual ListIndex Box before"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CheckBox Ck2 
      Caption         =   "Locked Option"
      Height          =   255
      Left            =   150
      TabIndex        =   13
      ToolTipText     =   "Try Type any Text on Combo"
      Top             =   1830
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "2 - Blank Combo Text"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Cleaning only Combo Text"
      Top             =   4155
      Width           =   2055
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "1 - Load On Board Itens"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "After a TagCode Option"
      Top             =   3765
      Width           =   2055
   End
   Begin VB.TextBox Tx3 
      Height          =   285
      Left            =   135
      TabIndex        =   9
      Text            =   "00"
      Top             =   1470
      Width           =   405
   End
   Begin VB.TextBox Tx2 
      Height          =   285
      Left            =   135
      TabIndex        =   7
      Text            =   "00"
      Top             =   1110
      Width           =   330
   End
   Begin VB.CheckBox Ck1 
      Caption         =   "TagCode Option"
      Height          =   255
      Left            =   135
      TabIndex        =   6
      ToolTipText     =   "After charge, try press a Load Command below"
      Top             =   780
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.TextBox Tx1 
      Height          =   315
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "Control has no '.List(i)"" - uses .Text instead"
      Top             =   2595
      Width           =   1770
   End
   Begin VB.TextBox TxCod 
      Height          =   300
      Left            =   3285
      TabIndex        =   1
      Top             =   105
      Width           =   1260
   End
   Begin JZXPCb.JZXPCbo CBO 
      Height          =   315
      Left            =   2235
      TabIndex        =   0
      Top             =   450
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   556
      Text            =   "THIS is the control"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      TagCode         =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Work Text to Ins, Upd  or Append"
      Height          =   195
      Left            =   150
      TabIndex        =   28
      Top             =   3075
      Width           =   2400
   End
   Begin VB.Line Line2 
      X1              =   105
      X2              =   4530
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   75
      Picture         =   "fTesteCbo.frx":0000
      Top             =   450
      Width           =   240
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Next"
      Height          =   195
      Left            =   1845
      TabIndex        =   26
      Top             =   4920
      Width           =   330
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "FileName: Path were Combo Contents will be Saved/Loaded"
      Height          =   195
      Left            =   150
      TabIndex        =   17
      Top             =   5475
      Width           =   4290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Actual ListCount"
      Height          =   195
      Left            =   615
      TabIndex        =   10
      Top             =   1515
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Actual ListIndex"
      Height          =   195
      Left            =   525
      TabIndex        =   8
      ToolTipText     =   "If you modf it, Try Cmd 3"
      Top             =   1155
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Actual List Retrieve"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   2370
      Width           =   1380
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   4530
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "App Elegant Form Combo"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   495
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "App Internal Simultaneos Data TagCode"
      Height          =   195
      Left            =   375
      TabIndex        =   2
      Top             =   195
      Width           =   2850
   End
End
Attribute VB_Name = "fTesteCbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Copying CBO contents to a normal combo
Private Sub Bt5_Click()
   Dim i As Long
   NCB.Clear
   i = 0
   Do While i < CBO.ListCount
      NCB.AddItem CBO.GetItem(i)
      i = i + 1
   Loop
End Sub

'Removeing actual item
Private Sub Bt8_Click()
  Dim i As Integer
  i = CBO.ListIndex
  If i >= 0 Then
     CBO.RemoveItem i
  End If
End Sub

'Appending a new item to Combo
Private Sub Bt6_Click()
  Dim s As String
  s = Trim(Tx9.Text)
  If Len(s) <> 0 Then
     CBO.AddItem s ' will be appended
  End If
End Sub

'Updating a item to Combo actual list
Private Sub Bt7_Click()
  Dim s As String
  s = Trim(Tx9.Text)
  If Len(s) <> 0 Then
     CBO.UpdateItem s, CBO.ListIndex   ' will be updated to actual listindex
  End If
End Sub

'Adding a new item to Combo BEFORE actual list
Private Sub Bt9_Click()
  Dim s As String
  s = Trim(Tx9.Text)
  If Len(s) <> 0 Then
     CBO.AddItem s, CBO.ListIndex  ' will be added before actual listindex
  End If
End Sub

'Getting simultaneos results
Private Sub CBO_Change()
  Tx1.Text = CBO.Text
  TxCod.Text = CBO.GetTagCode
  Tx2.Text = CStr(CBO.ListIndex)
  Tx3.Text = Format(CBO.ListCount, "#00")
  If CBO.ListIndex < 0 Then
     Cmd2_Click 'clean some boxes
  Else
     Tx1.Text = CBO.Text
     TxCod.Text = CBO.GetTagCode
  End If
End Sub

'App testing values on TagCod option
Private Sub Ck1_Click()
   If Ck1.Value = 0 Then
      CBO.TagCode = False
   Else
      CBO.TagCode = True
   End If
End Sub

'A normal Combo additem using TagCodes
Private Sub Combo_Load()
  Cmd2_Click ' clear boxes on App Screen
  Op1.Value = False
  Op2.Value = False
  Op3.Value = False
  Op4.Value = False
  TxI.Text = "00"
  
'default on-board example
  CBO.Clear
  CBO.AddItem "One=First"
  CBO.AddItem "Two=Second"
  CBO.AddItem "Three=Third"
  CBO.AddItem "Four=Fourth"
  CBO.AddItem "Five=Fifth"
  CBO.AddItem "Six=Sixth"
  CBO.AddItem "Seven=Seventh"
  CBO.AddItem "Eight=Eighth"
  CBO_Change
End Sub

'App testing Locked - user will not be able digit into Combo Box
Private Sub Ck2_Click()
   If Ck2.Value = 0 Then
      CBO.Locked = False
   Else
      CBO.Locked = True
   End If
End Sub

'A App reset to initial status of program
Private Sub Cmd1_Click()
  Combo_Load
End Sub

'App is cleanning screen fields
Private Sub Cmd2_Click()
   CBO.Text = ""
   TxCod.Text = ""
   Tx1.Text = ""
End Sub

'App is positioning Combo at a determinated index
'(index may be digited in TextBox)
Private Sub Cmd3_Click()
   Dim i As Long
   i = CLng(Tx2.Text)
   If Not i < CBO.ListCount Then
      i = CBO.ListCount - 1
   End If
   If i < 0 Then
      i = 0
   End If
   CBO.ListIndex = i
End Sub

'Simules you reading a Data Code from file and
'setting appropriate Combo value
'the StartFromIndex is used only when there is
'repetitions of same TagCode on List (e.g, Celebrities example)
'The most cases StartFromIndex is Zero
Private Sub Cmd4_Click()
   Dim i As Long
   i = CLng(TxI.Text)
   If Not CBO.SetTagCode(TxCod.Text, i) Then
      MsgBox TxCod.Text & " is Invalid or Not Existent"
      TxI.Text = "00"
   Else
'next tagcode to search init returned
      TxI.Text = Format(i, "00")
   End If
End Sub

'App will Save the actual Combo Box (and TagCode if it uses it)
'The complete filename is expressed on TextBox at Form botton
Private Sub Cmd5_Click()
   If Len(Trim(Tx4.Text)) = 0 Then
      MsgBox "Needs a valid FileName first"
      Exit Sub
   End If
   CBO.FileName = Trim(Tx4.Text)
   CBO.SaveToFile
   MsgBox "Saved on " & CBO.FileName
End Sub

'Now App will reload the last saved version of Combo
Private Sub Cmd6_Click()
   If Len(Trim(Tx4.Text)) = 0 Then
      MsgBox "Needs a valid FileName first"
      Exit Sub
   End If
   FromFile Trim(Tx4.Text)
End Sub

'This is a App generic routine to Load a Combo from File
Private Sub FromFile(fNam As String)
   Cmd2_Click 'clear fields
   CBO.FileName = fNam
   If CBO.LoadedFromFile Then
      MsgBox fNam & " Loaded."
   Else
      MsgBox fNam & " Error."
   End If
   If CBO.TagCode = True Then
      Ck1.Value = 1
   Else
      Ck1.Value = 0
   End If
   TxI.Text = "00"
   Tx2.Text = CStr(CBO.ListIndex)
   Tx3.Text = Format(CBO.ListCount, "00")
   CBO.FileName = Trim(Tx4.Text)
End Sub

'At App starting, some defaults
Private Sub Form_Load()
  Tx4.Text = App.Path & "\" & CBO.Name & ".TxT"
  Combo_Load
'loading a normal combo box to compare visual
  NCB.Clear
  NCB.AddItem "One=First"
  NCB.AddItem "Two=Second"
  NCB.AddItem "Three=Third"
  NCB.AddItem "Four=Fourth"
  NCB.AddItem "Five=Fifth"
  NCB.AddItem "Six=Sixth"
  NCB.AddItem "Seven=Seventh"
  NCB.AddItem "Eight=Eighth"
End Sub

'Celebrities - has repetitions in TagCode designs
'Use this as a litle Data Table or Search Engine
Private Sub Op1_Click()
   If Op1.Value = True Then
      FromFile App.Path & "\Celebs.TxT"
   End If
End Sub

'This is a "same size" formated example
'Some countries and corresponding automobilistic siglas
Private Sub Op2_Click()
   If Op2.Value = True Then
      FromFile App.Path & "\Countrs.TxT"
   End If
End Sub

'Well, App uses it as a "so-so" HTML colors menu
Private Sub Op3_Click()
   If Op3.Value = True Then
      FromFile App.Path & "\Colors.TxT"
   End If
End Sub

'App uses a simple test
'You may edit yourself "MyTest.TxT" to see effects you want
Private Sub Op4_Click()
   If Op4.Value = True Then
      FromFile App.Path & "\MyTest.TxT"
   End If
End Sub

'This is only a App logic to prepare the
'"Start" of Next search-and-set operation
'after modified teor of a TagCode
'The control does not it automaticaly mode not
'restringe versatily of search solutions from
'programmers
Private Sub TxCod_GotFocus()
   TxI.Text = "00"
End Sub
Private Sub CBO_Click()
   TxI.Text = "00"
End Sub


