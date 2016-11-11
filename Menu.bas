Attribute VB_Name = "Menu"
Option Explicit

Public thisGame As Game

Public Sub Start()
    Data.CompileData
    Set thisGame = New Game
    thisGame.Start
    Control.SetControls
End Sub

Public Sub StopGame()
    Dim n As Integer
    n = 1 / 0
End Sub

