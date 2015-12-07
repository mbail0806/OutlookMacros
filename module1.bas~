Attribute VB_Name = "Module1"

''''''''''''''''''''''''''''''''''''''''''''''''''
' Main: Subroutine ExtractAttachments()
' Description: Extract the attachment 
' Author: Mike Bail <bail@infionline.net>
' Version: 1.0
' Build: 1
' Date: 2015-12-06
' Contains: 
'   f: ea_CleanSubj
' Dependancy: 
'   f: ExtractDTGwSender
' ToDo: 
'   Still have to get attachment object and see how to save it to a particular directory
''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Sub ExtractAttachments()

' Declaring variables
Dim ea_Msg As String
Dim ea_Title As String
Dim ea_MsgRtn As String
Dim ea_MsgBttn As Integer
Dim ea_CurrMsg As Object
Dim ea_Subject As String
Dim ea_MsgDTG As String

' Get information out of the current message
Set ea_CurrMsg = Application.ActiveInspector.CurrentItem
ea_MsgDTG = ExtractDTGwSender(ea_CurrMsg, False)
ea_Subject = ea_CleanSubj(ea_CurrMsg.Subject)

' Set parameters for the message box
ea_MsgBttn = vbYesNoCancel + vbQuestion + vbDefaultButton3 + vbSystemModal
ea_Msg = "Getting Date-Time Group (" & ea_MsgDTG & "); do you want the subject also?"
ea_Title = "Extracting DTG w/Sender"
ea_MsgRtn = MsgBox(ea_Msg, ea_MsgBttn, ea_Title)

' Copy to clipboard using Windows API
Select Case ea_MsgRtn
    Case vbYes
        ClipBoard_SetData (ea_MsgDTG & "_" & ea_Subject)
    Case vbNo
        ClipBoard_SetData (ea_MsgDTG)
    Case vbCancel
        ClipBoard_SetData ("")
End Select
        
End Sub

Private Function ea_CleanSubj(eacs_String As String) As String
Dim eacs_PreLen As Integer
Dim eacs_RtnString As String

eacs_PreLen = InStr(eacs_String, ": ")

If (eacs_PreLen = 4) Or (eacs_PreLen = 5) Then
    Select Case LCase(Left(eacs_String, eacs_PreLen))
        Case "re: "
            eacs_RtnString = Right(eacs_String, (Len(eacs_String) - 4))
        Case "fwd: "
            eacs_RtnString = Right(eacs_String, (Len(eacs_String) - 5))
        Case Else
            eacs_RtnString = eacs_String
    End Select
Else
    eacs_RtnString = eacs_String
End If

ea_CleanSubj = eacs_RtnString

End Function

