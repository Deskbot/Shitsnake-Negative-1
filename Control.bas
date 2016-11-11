Attribute VB_Name = "Control"
Option Explicit

Private hiddenCell As Range

Public Sub SetControls()
    Set hiddenCell = GameSheet.cells(1, 1) 'GameSheet.rows.count, GameSheet.columns.count
    Application.OnKey "{LEFT}", "Left"
    Application.OnKey "{RIGHT}", "Right"
    Application.OnKey "{UP}", "Up"
    Application.OnKey "{DOWN}", "Down"
End Sub

Public Sub Left()
    thisGame.player.moveLeft
    hiddenCell.Select
End Sub
Public Sub Right()
    thisGame.player.moveRight
    hiddenCell.Select
End Sub
Public Sub Up()
    thisGame.player.moveUp
    hiddenCell.Select
End Sub
Public Sub Down()
    thisGame.player.moveDown
    hiddenCell.Select
End Sub

