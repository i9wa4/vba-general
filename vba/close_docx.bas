Attribute VB_Name = "Module1"
' close docx.bas
Option Explicit


Sub Close_Document()
    Speed_Up = True

    Dim toc As Word.TableOfContents
    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next

    Speed_Up = False

    If Documents.Count > 1 Then
        Application.DisplayAlerts = False
        ActiveDocument.Close SaveChanges:=True
        Application.DisplayAlerts = True
    Else
        Application.DisplayAlerts = False
        ActiveDocument.Close SaveChanges:=True
        Application.DisplayAlerts = True
        Application.Quit
    End If
End Sub


Property Let Speed_Up(ByVal b_shouold_Speed_Up As Boolean)
    Application.ScreenUpdating = Not b_shouold_Speed_Up
End Property
