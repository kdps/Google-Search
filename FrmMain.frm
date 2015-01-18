VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8055
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ListView ListBackup 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListFunction 
      Height          =   2415
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton FunctionAdd 
      Caption         =   "▶"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton FunctionDel 
      Caption         =   "◀"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton FunctionDown 
      Caption         =   "▼"
      Height          =   1215
      Left            =   7560
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton FunctionUp 
      Caption         =   "▲"
      Height          =   1215
      Left            =   7560
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type Engine
    Tip As String
    Function As String
    Pos As String
End Type

Private Sub Form_Load()
Dim arrEngine() As Engine
Dim strTmp As String
Dim arrTmp() As String
Dim a As Integer
ReDim arrEngine(0)
Open App.Path & "\Engine.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, strTmp
    If Trim(strTmp) <> "" Then
        a = a + 1
        arrTmp = Split(strTmp, ",")
        ReDim Preserve arrEngine(a)
        arrEngine(a).Tip = Trim(Replace(arrTmp(0), """", ""))
        arrEngine(a).Function = Trim(Replace(arrTmp(1), """", ""))
        arrEngine(a).Pos = Trim(Replace(arrTmp(2), """", ""))
        ListFunction.ListItems.Add , arrEngine(a).Function, arrEngine(a).Tip
        'ListFunction.ListItems.Item(ListFunction.ListItems.Count - 1).Text = arrEngine(a).Pos
Loop
Close #1
End Sub

Private Sub FunctionAdd_Click()
Dim i As Long
For i = 0 To ListFunction.ListCount - 1
If ListFunction.ListItems(i) = ListBackup.ListItems(ListBackup.ListIndex) Then
Exit Sub
End If
Next i
ListFunction.AddItem ListBackup.ListItems(ListBackup.ListIndex)
ListBackup.ListItems.Remove (ListBackup.ListIndex)
End Sub

Private Sub FunctionDel_Click()
ListBackup.AddItem ListFunction.List(ListFunction.ListIndex)
ListFunction.RemoveItem (ListFunction.ListIndex)
End Sub

Private Sub FunctionDown_Click()
Dim Backup As String
Backup = ListFunction.List(ListFunction.ListIndex)
If Not ListFunction.ListIndex = ListFunction.ListCount - 1 Then
ListFunction.List(ListFunction.ListIndex) = ListFunction.List(ListFunction.ListIndex + 1)
ListFunction.List(ListFunction.ListIndex + 1) = Backup
ListFunction.ListIndex = ListFunction.ListIndex + 1
End If
End Sub

Private Sub FunctionUp_Click()
Dim Backup As String
Backup = ListFunction.List(ListFunction.ListIndex)
If Not ListFunction.ListIndex = 0 Then
ListFunction.List(ListFunction.ListIndex) = ListFunction.List(ListFunction.ListIndex - 1)
ListFunction.List(ListFunction.ListIndex - 1) = Backup
ListFunction.ListIndex = ListFunction.ListIndex - 1
End If
End Sub

