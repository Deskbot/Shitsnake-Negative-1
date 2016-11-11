Attribute VB_Name = "Game"
Option Explicit

Private background As Long

Public Sub Start()
    'display settings
    With Data.settings
        If .Exists("bg-col") Then
            GameSheet.Select Data.settings("bg-col")
        End If
    End With
    
    GameSheet.Activate
End Sub

Public Property Let bgImg(i As String)
    GameSheet.SetBackgroundPicture (Data.filepath & "\bg.jpg")
End Property
