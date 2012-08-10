VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmManageLetters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Letters"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   15270
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7080
      Top             =   4080
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   13920
      TabIndex        =   11
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   2775
      Begin VB.OptionButton optToAssign 
         Caption         =   "To Assign"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optToComplete 
         Caption         =   "To Complete"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   760
         Width           =   2535
      End
      Begin VB.OptionButton optToReply 
         Caption         =   "To Reply"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1160
         Width           =   2535
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sort Order"
      Height          =   1935
      Left            =   3000
      TabIndex        =   0
      Top             =   6600
      Width           =   2775
      Begin VB.OptionButton optReplyDue 
         Caption         =   "Reply before Ascending"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
      End
      Begin VB.OptionButton optReceivedDec 
         Caption         =   "Received Date Desending"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1160
         Width           =   2535
      End
      Begin VB.OptionButton optReceivedAsc 
         Caption         =   "Received Date Ascending"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   760
         Width           =   2535
      End
      Begin VB.OptionButton optPriority 
         Caption         =   "Priority"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridList 
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11245
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmManageLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit
    Dim frmResize As New clsResizer

    
    Dim formActive As Boolean
    Dim activeCount As Integer

Private Sub btnSave_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call prepareResize
    SetColours Me
    Me.BorderStyle = 2
    GetCommonSettings Me
    Call process
End Sub

Private Sub Form_Activate()
    formActive = True
    activeCount = 0
End Sub

Private Sub Form_Deactivate()
    formActive = False
    activeCount = 0
End Sub

Private Sub Timer1_Timer()
    If formActive = True Then
        If activeCount > 10 Then
            activeCount = 0
            Call process
        Else
            activeCount = activeCount + 1
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub Form_Resize()
  Call frmResize.FormResized(Me)
End Sub

Private Sub prepareResize()
  frmResize.KeepRatio = False
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
End Sub


Private Sub gridList_DblClick()
    Dim myForm As New frmLetter
    myForm.Show
    myForm.ZOrder 0
    myForm.txtLetterID = Val(gridList.TextMatrix(gridList.row, 0))
End Sub

Private Sub optAll_Click()
    Call process
End Sub

Private Sub optPriority_Click()
    Call process
End Sub

Private Sub process()
    Dim temTopic As String
    Dim temSubTopic As String
    Dim temSelect As String
    Dim temFrom As String
    Dim temWhere As String
    Dim temGroupBy As String
    Dim temOrderBy As String
    Dim temSQL As String
    
    
    Dim D(0) As Integer
    Dim p(0) As Integer
    
    temTopic = "Letter List"
    
    If optAll.Value = True Then
        temTopic = temTopic & " - All"
    ElseIf optToAssign.Value = True Then
        temTopic = temTopic & " - Yet Assign"
    ElseIf optToComplete.Value = True Then
        temTopic = temTopic & " - To Complete"
    ElseIf optToReply.Value = True Then
        temTopic = temTopic & " - To Reply"
    End If
    
    
    
    
    
    temSelect = "SELECT tblLetter.LetterID,tblLetter.LetterDate as [Letter Date], Sender.PersonName as [From], tblLetter.LetterNumber as [Ref No], tblLetter.LetterTopic as [Topic], tblPriority.PriorityName as [Priority], Assigner.PersonName as [Assigned to] , format$(tblLetter.Completed, 'yes/no') as  [Completed], format$(tblLetter.ReplyDueDate, '" & ProgramVariable.ShortDateFormat & "' ) as [Reply Before] "
    temFrom = "FROM ((tblLetter LEFT JOIN tblPerson AS Sender ON tblLetter.SenderPersonID = Sender.PersonID) LEFT JOIN tblPerson AS Assigner ON tblLetter.AssignedPersonID = Assigner.PersonID) LEFT JOIN tblPriority ON tblLetter.PriorityID = tblPriority.PriorityID "
    
    If optAll.Value = True Then
        temWhere = "WHERE (((tblLetter.Deleted)=False)) "
    ElseIf optToAssign.Value = True Then
        temWhere = "WHERE (((tblLetter.Deleted)=False) AND ((tblLetter.Assigned)=False))    "
    ElseIf optToReply.Value = True Then
        temWhere = "WHERE (((tblLetter.Completed)=True) AND ((tblLetter.Deleted)=False)  AND ((tblLetter.Replied)=False) )    "
    ElseIf optToComplete.Value = True Then
        temWhere = "WHERE (((tblLetter.Completed)=False) AND ((tblLetter.Deleted)=False) AND ((tblLetter.Assigned)=True))    "
    End If
    
    If optPriority.Value = True Then
        temOrderBy = temOrderBy & " ORDER BY tblPriority.OrderNo "
    ElseIf optReceivedAsc.Value = True Then
        temOrderBy = temOrderBy & " ORDER BY tblLetter.ReceivedDate "
    ElseIf optReceivedDec.Value = True Then
        temOrderBy = temOrderBy & " ORDER BY tblLetter.ReceivedDate DESC "
    ElseIf optReplyDue.Value = True Then
        temOrderBy = temOrderBy & " ORDER BY tblLetter.ReplyDueDate"
    End If
    
    
    temGroupBy = ""
    
    temSQL = temSelect & temFrom & temWhere & temGroupBy & temOrderBy
    FillAnyGrid temSQL, gridList, 0, D, p
    gridList.ColWidth(0) = 0
    
End Sub


Private Sub optReceivedAsc_Click()
    Call process
End Sub

Private Sub optReceivedDec_Click()
    Call process
End Sub

Private Sub optToAssign_Click()
    Call process
End Sub

Private Sub optToComplete_Click()
    Call process
End Sub

Private Sub optToReply_Click()
    Call process
End Sub
