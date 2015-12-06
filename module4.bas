Attribute VB_Name = "Module4"
Option Explicit

Function ExtractDTGwSender(eds_CurrMsg As Object, eds_Clip As Boolean) As String

' Declaring variables
Dim eds_Msg As String
Dim eds_Title As String
Dim eds_MsgRtn As String
Dim eds_MsgBttn As Integer
Dim eds_SentDTG As Date
Dim eds_Sender As String
Dim eds_SenderIntl As String
' Dim eds_TZ As Object
' Dim eds_ZoneNum As Integer


' Get information out of the current message
eds_SentDTG = eds_CurrMsg.SentOn
eds_Sender = eds_CurrMsg.SenderEmailAddress
' eds_TZ = TimeZone.CurrentTimeZone
' eds_ZoneNum = eds_TZ.Bias


' Get the senders initials from a case statement
Select Case LCase(eds_Sender)
    Case "val.kozak@iecbiz.com"
        eds_SenderIntl = "vk"
    Case "gklott@yarcom.com"
        eds_SenderIntl = "gkl"
    Case "iecsouth@aol.com", "bill.stark@iecbiz.com"
        eds_SenderIntl = "wrs"
    Case "ieceast@mindspring.com", "bail@infionline.net", "mike.bail@iecbiz.com"
        eds_SenderIntl = "mtb"
    Case "iecnorth@aol.com", "dave.bailey@iecbiz.com"
        eds_SenderIntl = "dcb"
    Case Else
        eds_SenderIntl = "unk"
End Select

eds_Msg = eds_SenderIntl & "_" & Format(eds_SentDTG, "yyyymmddThhmmss") & get_TZLtr()

If eds_Clip Then ClipBoard_SetData (eds_Msg)

ExtractDTGwSender = eds_Msg

End Function

