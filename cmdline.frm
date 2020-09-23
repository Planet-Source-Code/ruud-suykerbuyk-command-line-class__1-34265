VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' WHEN YOUR APP. IS READY FOR SYSTEM WIDE FILE-ASSOCIATIONS ,
' THIS SIMPLE CLASS IS ALL YOU WILL NEED TO EXTRACT ANY COMMAND LINE ARGUMENTS DURING RUN-TIME.


' 1 : COMPILE THIS FORM INTO A .EXE FILE.

' 2 : DRAG SOME FILES ON THE .EXE FILE WITHIN YOUR MS-EXPLORER WINDOW

' 3 : CREATE A NEW SHOTCUT TO THE .EXE, ADD SOME ARGUMENTS. ( "C:\..\..blabla.exe" -Play /p /r )


Private AcCommandLine As New AcCommandLine

Private Sub Form_Initialize()

 Dim a, b
    
    MsgBox AcCommandLine.FileNames.Count & " filename(s) in the command line."
    
    MsgBox AcCommandLine.Arguments.Count & " Argument(s) in the command line."
    
    
    For Each a In AcCommandLine.FileNames

           Call List1.AddItem(a)
    Next

    For Each b In AcCommandLine.Arguments

            Call List2.AddItem(b)
    Next

End Sub


