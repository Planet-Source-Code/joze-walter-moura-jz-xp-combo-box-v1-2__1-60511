VERSION 5.00
Begin VB.UserControl JZXPCbo 
   BackColor       =   &H00D8E9EC&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   3090
   ToolboxBitmap   =   "JZXPCbo.ctx":0000
   Begin VB.PictureBox BackMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Width           =   2445
      Begin VB.PictureBox JImgCbo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         Picture         =   "JZXPCbo.ctx":0312
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox JTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Text            =   "0"
         Top             =   60
         Width           =   720
      End
      Begin VB.Shape ShapeBorder 
         BorderColor     =   &H00B99D7F&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.ComboBox JCombo 
      Height          =   315
      ItemData        =   "JZXPCbo.ctx":06C8
      Left            =   0
      List            =   "JZXPCbo.ctx":06CA
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   510
      Width           =   2295
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   150
      Picture         =   "JZXPCbo.ctx":06CC
      Top             =   900
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   1
      Left            =   150
      Picture         =   "JZXPCbo.ctx":0A82
      Top             =   1170
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   2
      Left            =   150
      Picture         =   "JZXPCbo.ctx":0E38
      Top             =   1470
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   3
      Left            =   150
      Picture         =   "JZXPCbo.ctx":11EE
      Top             =   1800
      Width           =   255
   End
End
Attribute VB_Name = "JZXPCbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'.------------------------------------------------------------------
' Control   : JZ XP Combo Box 1.2
' Edition   : 13-May-2005
' Author    : JOZE Walter de Moura - RIO DE JANEIRO, BRASIL.
'           : me: www.joze.kit.net   or   qualyum@globo.com
'           :
'           : Well, I've used basics on Combo Box "as XP"
'           : codes from several authors at Internet whose
'           : credits I acknowledge e appreciate.
'           :
'           : I've made a "brush up" certain routines and
'           : wrote all engines for TagCodes and Save/Load.
'           :
'Application: Another XP style Combo Box, but
'           : - as easy as TextBox programming: .Text, .Locked, etc.
'           : - a TagCode option allowing AddItens with leading
'           :   correlative data code, e.g., "US=United States" so
'           :   Combo lists is only "United States" and you can
'           :   retrieve "US" when selected by user.
'           : - A Save/Load engine using Text Files, notepadding
'           :   editable, so many applications as:
'           :   - Language and regional terminology supports;
'           :   - Small tables without DB;
'           :   - The same combo can get many text files - may be
'           :     a refined tree navigation.
'           :   - etc.
'           : - Maintenance functions to Append, Insert, Update and
'           :   Remove already loaded list items.
'           :
' License   : Freeware - you may distribute, alter, sold, anything
'           : as you want. This code is for you, don't it?
'           : I'm sure you will apply maximum of honesty and ethics
'           : concerning it.
'           :
' PS.       : TagCode treatment is limited to 100. If you need more,
'           : only to do is alter de MaxTCods constant value.
'           : Also, I've not improved other functions, as sort, auto-
'           : completes, etc., due avoiding "strong code".
'           :
'           : Joze.
'           :
' --------- :
' vers 1.2  : 16-May-2005
' --------- :
'           : Thanks a million to Territop (Paul) who have help to depure
'           : some bugs and suggest enhancements.
'           :
'  Enhances : 1. Function GetItem([index]) As String
'           :    Returns a string reflecting pointed Combo.List
'           :    If using TagCode then returns a string in format
'           :    "xxx=yyyy". i.e., TagCode & "=" & List.
'           :
'     Fixed : 1. MouseMove, MouseDown, KeyPress, KeyDown, KeyUp
'           :    for proper functions.
'           :
'           : 2. Bright effect on Combo Box Pick Botton now works ok.
'           :
'Know Errors: 1. When combo scrolling, in non-XP Windows, the
'           :    ScrollBar is not a Stylized "as XP".
'           :
'           :    Community Attention: I'd like your feedback if it
'           :    is a fundamental design adjust or not, and if
'           :    who had a nice and light suplemental code to do
'           :    this, ok?
'           :
'           : Thanks, Joze.
'           :
'`------------------------------------------------------------------'
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const CB_SHOWDROPDOWN = &H14F
Const CB_GETDROPPEDSTATE = &H157

'Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when te Text area of Combo is mouse clicked."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the Combo List is single modified (selected, add, upd, ins, removed, etc)."
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Const m_def_Enabled = 0
Const MaxTCods = 99

Dim m_Enabled As Boolean
Dim m_TagCode As Boolean
Dim m_FileName As String

Dim TCods(0 To MaxTCods) As String
Dim TLim As Long 'actual limit of array
Dim Tix As Long 'work pointer to array
Dim Titem As String
Dim TCod As String
Dim TTex As String
Dim m_Buf As String


Private Sub OpenCombo(chwnd As Long)
    Dim rc As Long
    rc = SendMessage(chwnd, CB_GETDROPPEDSTATE, 0, 0)
    If rc = 0 Then
        SendMessage chwnd, CB_SHOWDROPDOWN, True, 0
    Else
        SendMessage chwnd, CB_SHOWDROPDOWN, False, 0
    End If
    DoEvents
    ResetPic
End Sub

Private Sub RePos()
Dim i As Integer
    If Width < 400 Then Width = 400
    ShapeBorder.Width = Width
    JImgCbo.Left = Width - 285
    BackMain.Width = Width
    
    With JCombo
        .Top = 30
        .Left = 0
        .Width = Width
    End With
    
    Height = 315
    
    With JTexto
        .Width = Width - 375
        .Top = 60
        If .FontSize > 8 Then
            i = .FontSize - 8
            i = i * 15
            .Top = .Top - i
        End If
    End With
    
End Sub

Private Sub JCombo_Click()
    JTexto = JCombo.Text
End Sub

Private Sub JImgCbo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    JImgCbo.Picture = Img(2).Picture
    OpenCombo JCombo.hWnd
End Sub

Private Sub JImgCbo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If JImgCbo.Picture <> Img(1).Picture Then JImgCbo.Picture = Img(1).Picture
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub JTexto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub JTexto_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub JTexto_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub JTexto_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
    RePos
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = JTexto.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    JTexto.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Complete Path where Combo txt files will be on (App.Path &  .Name & '.TxT')"
   FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
   m_FileName = New_FileName
   PropertyChanged "FileName"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    JTexto.Text = PropBag.ReadProperty("Text", "0")
    JTexto.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Set JTexto.Font = PropBag.ReadProperty("Font", Ambient.Font)
    JTexto.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    JCombo.ListIndex = PropBag.ReadProperty("ListIndex", -1)
    JTexto.Locked = PropBag.ReadProperty("Locked", False)
    JTexto.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    JTexto.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    JTexto.DataField = PropBag.ReadProperty("FieldName", "")
    
    RePos
    
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_TagCode = PropBag.ReadProperty("TagCode", False)
    m_FileName = PropBag.ReadProperty("FileName", "")

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Text", JTexto.Text, "0")
    Call PropBag.WriteProperty("BackColor", JTexto.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Font", JTexto.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", JTexto.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ListIndex", JCombo.ListIndex, -1)
    Call PropBag.WriteProperty("Locked", JTexto.Locked, False)
    Call PropBag.WriteProperty("MaxLength", JTexto.MaxLength, 0)
    Call PropBag.WriteProperty("ToolTipText", JTexto.ToolTipText, "")
    Call PropBag.WriteProperty("FieldName", JTexto.DataField, "")
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("TagCode", m_TagCode, False)
    Call PropBag.WriteProperty("SelStart", JTexto.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", JTexto.SelLength, 0)
    Call PropBag.WriteProperty("SelText", JTexto.SelText, "")
    Call PropBag.WriteProperty("FileName", m_FileName, "")

End Sub

Sub ResetPic()
    If JImgCbo.Picture <> Img(0).Picture Then
        JImgCbo.Picture = Img(0).Picture
    End If
End Sub

Private Sub JTexto_Click()
    RaiseEvent Click
End Sub

Private Sub JTexto_Change()
   RaiseEvent Change
End Sub

Public Sub UpdateItem(Item As String, ByVal Index As Variant)
Attribute UpdateItem.VB_Description = "Replaces the pointed list (and TagCode) by new Item value."
   If Index < 0 Or Index > JCombo.ListCount - 1 Then ' test bounds
      Exit Sub
   End If
   If m_TagCode = False Then 'normal
      JCombo.List(Index) = Item 'updates it as is
   Else
'TagCode is on
      If Index > MaxTCods Or Index > TLim Then 'tcodes limits
         JCombo.List(Index) = Item 'updates it as is
      Else
         Call IsTagCode(Item)
         JCombo.List(Index) = TTex 'updates combo segment
         Tix = CLng(Index) ' for coherence proposes only
         TCods(Tix) = TCod
      End If
   End If
   JTexto.Text = JCombo.Text
   RaiseEvent Change
End Sub

Public Sub AddItem(Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to ComboBox. If no index parameter, the item will be apended else item will be inserted BEFORE corresponding list."
   If IsMissing(Index) Then
      PutItem Item
   Else
      If Index < 0 Then
         Index = 0
      End If
      If Index > JCombo.ListCount - 1 Then
         PutItem Item
      Else
         PutItem Item, Index
      End If
   End If
   JCombo.ListIndex = Tix
   JTexto.Text = JCombo.Text
   RaiseEvent Change
End Sub

Public Function GetItem(Optional Index As Variant) As String
Attribute GetItem.VB_Description = "Returns a String refleting Control List image, ponted by index. Empty if out of boundaries. Format ""xxx=yyy"" if TagCode on. "
   Dim i As Long
   Dim s As String
   If IsMissing(Index) Then
      If JCombo.ListCount > 0 Then
         i = JCombo.ListCount - 1
      Else
         GetItem = ""
         Exit Function
      End If
   Else
      i = CLng(Index)
      If i < 0 Then
         i = 0
      End If
   End If
   s = ""
   If m_TagCode = True Then
      s = TCods(i) & "="
   End If
   GetItem = s & JCombo.List(CInt(i)) 'maybe adjust code to future VB version
End Function

Public Sub RemoveItem(Optional Index As Variant)
Attribute RemoveItem.VB_Description = "Removes a item from list. If no index removes the last of combo list item else removes the pointed one."
   Dim i As Long
   If IsMissing(Index) Then
      If JCombo.ListCount > 0 Then
         i = JCombo.ListCount - 1
      Else
         Exit Sub
      End If
   Else
      i = CLng(Index)
      If i < 0 Then
         i = 0
      End If
   End If
   If m_TagCode = True Then
      RemoveTagCode i
   End If
   JCombo.RemoveItem CInt(i) 'maybe adjust code to future VB version
   RaiseEvent Change
End Sub

'before removeing a combo item
Private Sub RemoveTagCode(Index As Long)
  Dim j As Long
  If TLim > 0 Then
     For Tix = Index To TLim
         j = Tix + 1
         TCods(Tix) = TCods(j)
     Next Tix
  End If
End Sub

'returns True/False and TCod/TTex vars
Private Function IsTagCode(Item As String) As Boolean
  Dim s As String
  Dim len_s As Long
  Dim C As Long ' resultant column
  s = Trim(Item) ' cut blanks
  len_s = Len(s)
  If len_s <> 0 Then 'sure significant value
     C = InStr(1, s, "=", vbTextCompare)
     If C = 1 Then ' "equal" sign at 1st column?
        s = Mid(s, 2, len_s - 1) ' cut it
        C = 0 ' as not found
     End If
     If C = len_s Then '"equal" sign is last character
        s = Mid(s, 1, len_s - 1) 'cut it
        C = 0 ' as not found
     End If
  Else
     C = 0 'as not found
  End If
  Item = s ' returns item as interpreted
  If C = 0 Then ' "equal" sign not found
     TCod = Item
     TTex = Item
     IsTagCode = False
  Else
     TCod = Trim(Mid(s, 1, C - 1))
     TTex = Trim(Mid(s, C + 1, len_s - C))
     IsTagCode = True
  End If
End Function

Private Sub PutItem(Item As String, Optional ByVal Index As Variant)
  Dim Col As Long
  Dim j As Long
    If m_TagCode = False Then
       If Len(Item) = 0 Then ' no null itens
          Exit Sub
       Else
          JCombo.AddItem Item, Index  ' normal additem
       End If
    Else
       Call IsTagCode(Item)
       If Len(Item) = 0 Then ' resulted in null itens
          Exit Sub
       Else
          If IsMissing(Index) Then
             JCombo.AddItem TTex 'appending
             Tix = JCombo.ListCount - 1
             If Not Tix > MaxTCods Then ' limited to 100 TagCode entries
                TCods(Tix) = TCod
                TLim = Tix
             End If
          Else
             JCombo.AddItem TTex, Index
             Tix = CLng(Index)
             If Tix < 0 Then
                Tix = 0
             End If
             If Tix > TLim Then
                Tix = TLim
             End If
             Col = Tix
             Tix = TLim
             Do While Tix > Col - 1
                j = Tix + 1
'if inserting TagCode above TagCode array limits then
'last TagCode will be lost to prevent overflow
                If Not j > MaxTCods Then
                   TCods(j) = TCods(Tix)
                End If
                Tix = Tix - 1
             Loop
             TCods(Col) = TCod
             TLim = TLim + 1
          End If
       End If
       If TLim > MaxTCods Then
          TLim = MaxTCods
       End If
    End If
    If IsMissing(Index) Then
       Tix = JCombo.ListCount - 1
    Else
       Tix = Val(Index)
    End If
End Sub

'Saves Combo contents to a Txt File as App.Path & "\" & .Name & ".TxT"
Public Sub SaveToFile()
Attribute SaveToFile.VB_Description = "Saves all Combo elements, including TagCodes, to a new file as .Filename path."
  Dim FN As Integer
  Dim fNam As String
  Dim i As Long
  Dim s As String
  Dim t As String
  Dim u As String

' The filename would be initialized
  fNam = Trim(m_FileName)
  If Len(fNam) = 0 Then ' assumes anything
     fNam = App.Path & "\" & UserControl.Name & ".Txt"
  End If
  On Error GoTo Clo
  FN = FreeFile
  Open fNam For Output As #FN
  For i = 0 To JCombo.ListCount - 1
      If m_TagCode = False Then
         s = JCombo.List(i)
      Else
         If i > MaxTCods Then ' maximum TagCode array
            s = JCombo.List(i)
         Else
            t = TCods(i)
            u = JCombo.List(i)
            s = t & "=" & u
         End If
      End If
      Print #FN, s
  Next i
Clo:
  Close FN
End Sub

'Load Combo Itens from a Txt File as App.Path & "\" & .Name & ".TxT"
'Initializes the Combo and Sets/Resets TagCode flag by looking for a "=" anywere on text
'First line char = ";" this is a remark line
'
Public Function LoadedFromFile() As Boolean
Attribute LoadedFromFile.VB_Description = "Reloads Combo from actual .Filename  path (automatic TagCode option if finds a (=) at 1st line) (On error, False)."
  Dim FN As Integer
  Dim fNam As String
  Dim Cnt As Long
  Dim Col As Long
  Dim s As String
  
  m_Buf = ""
' The filename would be initialized
  fNam = Trim(m_FileName)
  If Len(fNam) = 0 Then ' assumes anything
     fNam = App.Path & "\" & UserControl.Name & ".Txt"
  End If
  On Error GoTo R_Error
  FN = FreeFile
  Open fNam For Input As #FN
  m_Buf = Input(LOF(FN), #FN)
  Close FN
  If Len(m_Buf) = 0 Then ' not empty file
     LoadedFromFile = False
     Exit Function
  End If
  
  Cnt = 0 'first line
  s = ""
  Do While Len(s) < 2 ' Ignores non-significant first lines
     s = A_line_of_Buf(Cnt)
     If Cnt = 0 Then
        LoadedFromFile = False
        Exit Function
     End If
  Loop
'Ok, we have at least 1 line : we empty the combo and prepare to load
  JCombo.Clear
  If IsTagCode(s) Then ' if at least a "=" character in 1st line
     m_TagCode = True
  Else
     m_TagCode = False
  End If
  PutItem s
' continue loading
  s = ""
  Do While Cnt > 0
     s = A_line_of_Buf(Cnt)
     If Cnt = 0 Then
        LoadedFromFile = True
        Exit Function
     End If
     If s <> ";" Then
        If Len(Trim(s)) <> 0 Then
           PutItem s
        End If
     End If
  Loop
  LoadedFromFile = True 'only for bug security
  Exit Function
R_Error:
  LoadedFromFile = False
End Function

'its a preference against use Split function (maybe a very strong one)
'initializes with LineNumber = 0
'no more lines returns LineNumber = 0
'maximum line lenght = 80 characteres
'first char ";" returns only the character ";"
Private Function A_line_of_Buf(LineNumber As Long) As String
  Dim j As Long
  Dim Cnt As Long
  Dim lia As String
  Cnt = LineNumber
  If Cnt = 0 Then 'eliminates char(10) 0AH
     lia = Replace(m_Buf, vbLf, "")
     m_Buf = lia
     If Not Mid(m_Buf, Len(m_Buf), 1) = vbCr Then
        m_Buf = m_Buf & vbCr 'asseveres last caracter
     End If
  End If
  lia = "" 'empty
  j = InStr(m_Buf, vbCr)
  If j = 0 Then
     If Len(m_Buf) = 0 Then
        Cnt = 0
     End If
  Else
     Cnt = Cnt + 1
     lia = Mid(m_Buf, 1, j - 1)
     m_Buf = Mid(m_Buf, 2, Len(m_Buf) - 1) 'eliminates first char
     If Len(m_Buf) = Len(lia) Then
        m_Buf = "" 'all empty
     Else
        m_Buf = Mid(m_Buf, Len(lia) + 1, Len(m_Buf) - Len(lia))
     End If
     If Len(lia) > 0 Then
        If Mid(lia, 1, 1) = ";" Then
           lia = ";"
        End If
     End If
  End If
  If Len(lia) > 80 Then 'security line limit
     lia = Mid(lia, 1, 80)
  End If
  LineNumber = j
  A_line_of_Buf = lia
End Function

'Returns actual index TagCode
Public Function GetTagCode() As String
Attribute GetTagCode.VB_Description = "Returns a string containing the TagCode corresponding to actual ListIndex at Combo (Empty on error)."
    TCod = ""
    If JCombo.ListIndex > -1 Then
       If m_TagCode = True Then
          If Not JCombo.ListIndex > MaxTCods Then
             TCod = TCods(JCombo.ListIndex)
          End If
       End If
    End If
    GetTagCode = TCod
End Function

'Shows item corresponding TagCode received and returns True
'or Shows blanc item and returns False
Public Function SetTagCode(TgCode As String, StartingFromIndex As Long) As Boolean
Attribute SetTagCode.VB_Description = "Search received string (case ignored) at TagCode, starting received pointer, and positiones the Combo to respectives list, text, etc. (On error, False)."
  Dim b As Boolean
  Dim i As Long
  Dim lim As Long
  Dim Col As Long
  i = StartingFromIndex
  
  If JCombo.ListCount - 1 > MaxTCods Then ' if exceeds arrays limit
     lim = MaxTCods
  Else
     lim = JCombo.ListCount - 1
  End If
  
  If i > lim Then 'accidental overflow prevent
     i = 0
  End If
    
  TCod = UCase(Trim(TgCode))
  b = False
  If Not Len(TCod) = 0 Then
     For Col = i To lim
         If UCase(TCods(Col)) = TCod Then
            b = True
            Exit For
         End If
     Next Col
  End If
  If b = True Then
     JCombo.ListIndex() = Col
     JTexto.Text = JCombo.Text
     Col = Col + 1
     If Col > lim Then
        Col = 0
     End If
     StartingFromIndex = Col ' returns sucess index
  Else
     JCombo.ListIndex() = -1
     JTexto.Text = ""
  End If
  SetTagCode = b
  RaiseEvent Change
End Function

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = JTexto.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    JTexto.BackColor() = New_BackColor
    BackMain.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of Combo and resets pointers, TagCode array, etc."
    JCombo.Clear
    TLim = 0 ' TagCode array will be empty
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = JTexto.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set JTexto.Font = New_Font
    RePos
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = JTexto.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    JTexto.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = JCombo.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    JCombo.ListIndex() = New_ListIndex
    If JCombo.ListIndex > -1 Then
       JTexto.Text = JCombo.Text
       If m_TagCode = True Then
          If Not JCombo.ListIndex > MaxTCods Then
             TCod = TCods(JCombo.ListIndex)
          End If
       Else
          TCod = ""
       End If
    End If
    PropertyChanged "ListIndex"
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = JCombo.ListCount
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = JTexto.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    JTexto.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = JTexto.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    JTexto.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Private Sub JTexto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = JTexto.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    JTexto.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get FieldName() As String
Attribute FieldName.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    FieldName = JTexto.DataField
End Property

Public Property Let FieldName(ByVal New_FieldName As String)
    JTexto.DataField() = New_FieldName
    PropertyChanged "FieldName"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/Sets the starting point of the selected text.  "
    SelStart = JTexto.SelStart
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/Set the number of characteres selected."
    SelLength = JTexto.SelLength
End Property

Public Property Get SelText() As Long
Attribute SelText.VB_Description = "Returns the string containing the currently selected text."
    SelText = JTexto.SelText
End Property

Public Property Get TagCode() As Boolean
Attribute TagCode.VB_Description = "Set/Reset if Combo uses or not the parallel TagCode resource, i.e., itens in format ""xxx=yyyyy"" auto separated."
    TagCode = m_TagCode
End Property

Public Property Let TagCode(ByVal New_TagCode As Boolean)
    m_TagCode = New_TagCode
    PropertyChanged "TagCode"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    JTexto.Enabled = New_Enabled
    If New_Enabled = False Then
        JImgCbo.Picture = Img(3).Picture
        ShapeBorder.BorderColor = &HC0C0C0
    Else
        ResetPic
        ShapeBorder.BorderColor = &HB99D7F
    End If
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_TagCode = False
    m_FileName = ""
End Sub

