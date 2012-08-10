VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Letter Manager"
   ClientHeight    =   4770
   ClientLeft      =   -150
   ClientTop       =   540
   ClientWidth     =   7515
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileBackUp 
         Caption         =   "Back up"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditPersons 
         Caption         =   "Persons"
      End
      Begin VB.Menu mnuEditMyDetails 
         Caption         =   "My Details"
      End
      Begin VB.Menu mnuEditUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu mnuEditPrevilages 
         Caption         =   "Previlages"
      End
      Begin VB.Menu mnuEditPriorities 
         Caption         =   "Priorities"
      End
   End
   Begin VB.Menu mnuLetters 
      Caption         =   "Letters"
      Begin VB.Menu mnuLettersReceive 
         Caption         =   "Receive"
      End
      Begin VB.Menu mnuLettersManageLetters 
         Caption         =   "Manage Letters"
      End
   End
End
Attribute VB_Name = "MDIMain"
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

Private Sub mnuEditMyDetails_Click()
    frmMyDetails.Show
    frmMyDetails.ZOrder 0
End Sub

Private Sub mnuEditPersons_Click()
    frmPerson.Show
    frmPerson.ZOrder 0
End Sub

Private Sub mnuEditPriorities_Click()
    frmPriority.Show
    frmPriority.ZOrder 0
End Sub

Private Sub mnuFileBackUp_Click()
    frmBackUp.Show
    frmBackUp.ZOrder 0
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileRestore_Click()
    frmRestore.Show
    frmRestore.ZOrder 0
End Sub

Private Sub mnuLettersManageLetters_Click()
    frmManageLetters.Show
    frmManageLetters.ZOrder 0
End Sub

Private Sub mnuLettersReceive_Click()
    Dim myForm As New frmLetter
    myForm.Show
    myForm.ZOrder 0
End Sub
