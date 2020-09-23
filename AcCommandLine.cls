VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcCommandLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'==================================================================================================================
' Copyright ©1999-2001 Ruud Suykerbuyk | http://boz.Myweb.nl                                last updated : 28.03.00
'
' Permission to use, copy, modify, and distribute this software for any purpose is hereby granted,
' provided that the above copyright notice appears in all copies and all supporting documentation.
'
' This class extracts the command line arguments during run-time.
'
'==================================================================================================================
'

Option Compare Text


Private AcArguments  As New Collection
Private AcFileNames  As New Collection


Public Property Get Arguments() As Collection
 Set Arguments = AcArguments
End Property

Public Property Get FileNames() As Collection
 Set FileNames = AcFileNames
End Property


Private Static Sub Class_Initialize()

 Const Mask As String = "\/:*?<>|"

 ' Setup default properties when the class Initializes.

 Dim Cmd As String, Char As String, Sz As String, a As Integer, b As Integer, c As Integer

        Cmd = LTrim(VBA.Command) ' LTrim(" /a /b /c/d  d:\Boot.ini d:\Boot.txt ")

 If Len(Cmd) > 0 Then

    ' Flag to tell if we're currently processing a command line argument.
    a = False

    For b = 1 To LenB(Cmd) Step 1

        Char = Mid(Cmd, b, 1)

        If InStr(" / ", Char) = 0 And InStr(Chr(9), Char) = 0 And InStr(Chr(34), Char) = 0 Then

                If Not (a = True) Then

                ' a new command line argument.

                For c = b To LenB(Cmd) Step 1

                    Char = Mid(Cmd, c, 1)

                    If InStr(" / ", Char) = 0 And InStr(Chr(9), Char) = 0 And InStr(Chr(34), Char) = 0 Then

                    Else
                            Sz = Format(Mid(Cmd, b, c - b), "<")

                            ' Got it , now add to the proper collection ..

                            If InStr(Sz, ".") And Not InStr(Sz, "\/:*?<>|") Then

                                ' This argument is probely a filename.

                                If IsFileName(Sz) Then Call AcFileNames.Add(Item:=Sz) Else _
                                                       Call AcArguments.Add(Item:=Sz)
                            Else
                                    Call AcArguments.Add(Item:=Sz)
                            End If

                            Exit For
                    End If

                Next

                a = True

                End If
        Else
                a = False
        End If

    Next

 End If

End Sub


Private Function IsFileName(ByVal Sz As String) As Boolean
 On Error Resume Next
 Call FileLen(Sz)
 IsFileName = (err = 0)
End Function
