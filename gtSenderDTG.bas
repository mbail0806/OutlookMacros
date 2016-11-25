Attribute VB_Name = "gtSenderDTG"

''''''''''''''''''''''''''''''''''''''''''''''''''
' Main: F= ExtractDTGwSender(Object, Boolean) String
' Description: Extract the date and time the message was sent.  The time is extracted as local time.  If Boolean is true, copy data to the system clipboard.  
' Author: Mike Bail <bail@infionline.net>
' Version: 1.0
' Build: 1
' Date: 2015-12-07
' Contains: None
' Dependancy: 
'   get_TZLtr()
'   Clipboard_SetData(String)
' Notes: 
' ToDo: 
''''''''''''''''''''''''''''''''''''''''''''''''''

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
    Case "lbail@infionline.net"
        eds_SenderIntl = "lmb"
    Case "thomasrichardbail@gmail.com"
        eds_SenderIntl = "trb"
    Case "emiman@msn.com"
        eds_SenderIntl = "hag"
    Case "gliposchak@othr.net"
        eds_SenderIntl = "gml"
    Case "dktaylor@wrsystems.com", "davetaylor76@yahoo.com"
        eds_SenderIntl = "dkt"
    Case "hruehle@wrsystems.com"
        eds_SenderIntl = "hjr"
    Case "driffle@wrsystems.com"
        eds_SenderIntl = "dlr"
    Case "ieceast@mindspring.com", "bail@infionline.net"
        eds_SenderIntl = "mtb"
    Case "n26_sanchez@hotmail.com"
        eds_SenderIntl = "jns"
    Case "roy_d_esquibel@raytheon.com"
        eds_SenderIntl = "rde"
    Case "iecsouth@aol.com", "bill.stark@iecbiz.com"
        eds_SenderIntl = "wrs"
    Case "gklott@yarcom.com"
        eds_SenderIntl = "gkl"
    Case "iecnorth@aol.com", "dave.bailey@iecbiz.com"
        eds_SenderIntl = "dcb"
    Case Else
        eds_SenderIntl = "unk"
End Select

eds_Msg = eds_SenderIntl & "_" & Format(eds_SentDTG, "yyyymmddThhmmss") & get_TZLtr()

If eds_Clip Then ClipBoard_SetData (eds_Msg)

ExtractDTGwSender = eds_Msg

End Function
