VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Load/Save TreeView from/to Text File"
   ClientHeight    =   6600
   ClientLeft      =   1785
   ClientTop       =   1455
   ClientWidth     =   10725
   LinkTopic       =   "Form10"
   ScaleHeight     =   6600
   ScaleWidth      =   10725
   Begin VB.TextBox txtText 
      Height          =   5535
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   3  '¾ç¹æÇâ
      TabIndex        =   4
      Top             =   960
      Width           =   5055
   End
   Begin VB.ComboBox cboSelect 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run!"
      Height          =   330
      Left            =   6840
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox cboOpenSave 
      Height          =   300
      ItemData        =   "Form1.frx":0004
      Left            =   120
      List            =   "Form1.frx":0006
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
   Begin MSComctlLib.ImageList imgNode 
      Left            =   8760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0008
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":035C
            Key             =   "TextFile"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TView 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9763
      _Version        =   393217
      Indentation     =   441
      Style           =   7
      ImageList       =   "imgNode"
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CoCreateGuid Lib "ole32.dll" (pGUID As Any) As Long
Private iFreeFile As Integer



Public Function CreateGUID() As String
    Dim i As Long, b(0 To 15) As Byte
    If CoCreateGuid(b(0)) = 0 Then
        For i = 0 To 15
            CreateGUID = CreateGUID & Right$("00" & Hex$(b(i)), 2)
        Next i
    Else
        MsgBox "Error While creating GUID!"
    End If
End Function

Public Sub OpenTreeViewFromFileWithTab()
    On Error GoTo Err_Handle
    txtText.Text = GetFileText(cboOpenSave.Text)
    iFreeFile = FreeFile
    Open cboOpenSave.Text For Input As iFreeFile
    LoadNodesFromFileWithTab
    Close iFreeFile
    MsgBox "Complete! Loading nodes from text file, structured with tabs", vbInformation, "Load TreeView"
    Exit Sub
Err_Handle:
    Close iFreeFile
    MsgBox Err.Number & vbCr & Err.Description
End Sub
Private Sub LoadNodesFromFileWithTab()
    Dim text_line As String
    Dim level As Integer
    Dim tree_nodes() As Node
    Dim num_nodes As Integer
    
    TView.Nodes.Clear
    Do While Not EOF(iFreeFile)
        Line Input #iFreeFile, text_line
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If
        If level = 1 Then
            Set tree_nodes(level) = TView.Nodes.Add(, , CreateGUID, text_line, 1)
        Else
            Set tree_nodes(level) = TView.Nodes.Add(tree_nodes(level - 1), tvwChild, CreateGUID, text_line, 1)
            tree_nodes(level).EnsureVisible
        End If
    Loop
    TView.Nodes.Item(1).EnsureVisible
End Sub

Public Sub OpenTreeViewFromFileWithFullPath()
    On Error GoTo Err_Handle
    txtText.Text = GetFileText(cboOpenSave.Text)
    iFreeFile = FreeFile
    Open cboOpenSave.Text For Input As iFreeFile
    LoadNodesFromFileWithFullPath
    Close iFreeFile
    MsgBox "Complete! Loading nodes from text file containing full paths", vbInformation, "Load TreeView"
    Exit Sub
Err_Handle:
    Close iFreeFile
    MsgBox Err.Number & vbCr & Err.Description
End Sub
Private Sub LoadNodesFromFileWithFullPath()
    Dim text_line As String
    Dim level As Integer
    Dim tree_nodes() As Node
    Dim num_nodes As Integer
    Dim pos As Long
    
    TView.Nodes.Clear
    Do While Not EOF(iFreeFile)
        Line Input #iFreeFile, text_line
        level = UBound(Split(text_line, "\")) + 1
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If
        
        pos = InStrRev(text_line, "\")
        If pos Then text_line = Mid$(text_line, pos + 1)

        If level = 1 Then
            Set tree_nodes(level) = TView.Nodes.Add(, , CreateGUID, text_line, 1)
        Else
            Set tree_nodes(level) = TView.Nodes.Add(tree_nodes(level - 1), tvwChild, CreateGUID, text_line, 1)
            tree_nodes(level).EnsureVisible
        End If
    Loop
    TView.Nodes.Item(1).EnsureVisible
End Sub

Sub SaveTreeViewToFileWithTab()
    On Error GoTo Err_Handle
    iFreeFile = FreeFile()
    Open cboOpenSave.Text For Output As #iFreeFile
    SaveTreeWithTab TView.Nodes.Item(1)
    Close #iFreeFile
    MsgBox "Complete! Saving nodes to text file, structured with tabs", vbInformation, "Save TreeView"
    Exit Sub
Err_Handle:
    Close iFreeFile
    MsgBox Err.Number & vbCr & Err.Description
End Sub
Private Sub SaveTreeWithTab(oNode As Node)
    Dim oSibNode As Node
    Set oSibNode = oNode
    Do
        Print #iFreeFile, String(UBound(Split(oSibNode.FullPath, "\")), vbTab) & oSibNode.Text
        If Not oSibNode.Child Is Nothing Then
            SaveTreeWithTab oSibNode.Child
        End If
        Set oSibNode = oSibNode.Next
   Loop While Not oSibNode Is Nothing
End Sub

Sub SaveTreeViewToFileWithFullPath()
    On Error GoTo Err_Handle
    iFreeFile = FreeFile()
    Open cboOpenSave.Text For Output As #iFreeFile
    SaveTreeWithFullPath TView.Nodes.Item(1)
    Close #iFreeFile
    MsgBox "Complete! Saving nodes to text file containing full paths", vbInformation, "Save TreeView"
    Exit Sub
Err_Handle:
    Close iFreeFile
    MsgBox Err.Number & vbCr & Err.Description
End Sub

Private Sub SaveTreeWithFullPath(oNode As Node)
    Dim oSibNode As Node
    Set oSibNode = oNode
    Do
         Print #iFreeFile, oSibNode.FullPath
         If Not oSibNode.Child Is Nothing Then
             SaveTreeWithFullPath oSibNode.Child
         End If
         Set oSibNode = oSibNode.Next
    Loop While Not oSibNode Is Nothing
End Sub

Private Sub cmdRun_Click()
    Select Case cboSelect.ListIndex
    Case 0: OpenTreeViewFromFileWithTab
    Case 1: OpenTreeViewFromFileWithFullPath
    Case 2: SaveTreeViewToFileWithTab
    Case 3: SaveTreeViewToFileWithFullPath
    End Select
End Sub

Private Sub Form_Load()
    With cboOpenSave
        .AddItem App.Path & "\tree-tab.txt"
        .AddItem App.Path & "\tree-fullpath.txt"
        .ListIndex = 0
    End With
    With cboSelect
        .AddItem "Load nodes from text file, structured with tabs"
        .AddItem "Load nodes from text file containing full paths"
        .AddItem "Save nodes to text file, structured with tabs"
        .AddItem "Save nodes to text file containing full paths"
        .ListIndex = 0
    End With
End Sub

Function GetFileText(ByVal strFilePathName As String) As String
    If FileExists(strFilePathName) Then
        Dim Buffer() As Byte
        ReDim Buffer(FileLen(strFilePathName))
        Open strFilePathName For Binary As #1  'Source
        Get #1, , Buffer
        Close
        GetFileText = Replace(BytesToStr(Buffer, False, False), Chr(0), "")
    End If
End Function
Private Function BytesToStr(Buffer() As Byte, _
                           Optional IsAnsi As Boolean = True, _
                           Optional IsUnicode As Boolean = False) As String
    Dim Unspecified     As Boolean
    Unspecified = (Abs(IsAnsi) + Abs(IsUnicode)) = 0
    If IsAnsi Or Unspecified Then
        BytesToStr = StrConv(Buffer, vbUnicode)
    Else
        BytesToStr = Buffer
    End If
End Function
Private Function FileExists(ByVal sFullPath As String) As Boolean
    FileExists = (Len(Dir(sFullPath)) > 0)
End Function

