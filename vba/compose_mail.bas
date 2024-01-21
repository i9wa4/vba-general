Attribute VB_Name = "compose_mail"
' compose_mail.bas
Option Explicit


Sub Compose_Active_Sheet_Mail()
    Call Compose_Mail(ThisWorkbook.ActiveSheet)
End Sub


Function Compose_Mail(ws As Worksheet)
    Dim obj_Outlook As Outlook.Application
    Dim obj_Mail_Item As Outlook.MailItem

    Set obj_Outlook = CreateObject("Outlook.Application")
    Set obj_Mail_Item = obj_Outlook.CreateItem(olMailItem)

    With ws
        obj_Mail_Item.BodyFormat = olFormatPlain
        obj_Mail_Item.To = .Range("B2").Value
        obj_Mail_Item.CC = .Range("B3").Value
        obj_Mail_Item.BCC = .Range("B4").Value
        obj_Mail_Item.Subject = .Range("B5").Value
        obj_Mail_Item.Body = .Range("B6").Value & vbCrLf & vbCrLf & .Range("B7").Value

        If .Range("B8").Value <> "" Then
            Dim obj_Attatchments As Outlook.Attachments
            Set obj_Attatchments = obj_Mail_Item.Attachments
            obj_Attatchments.Add .Range("B8").Value
        End If
    End With

    obj_Mail_Item.Display

    Set obj_Outlook = Nothing
    Set obj_Mail_Item = Nothing
End Function
