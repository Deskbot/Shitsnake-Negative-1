Attribute VB_Name = "Factory"
Option Explicit

Function NewSprite(name As String) As Sprite
    Set NewSprite = New Sprite
    NewSprite.Init name
End Function

Function NewElement(s As Sprite) As Element
    Set NewElement = New Element
    NewElement.Init s
End Function

Function NewSnake(g As Game, r As Range) As Snake
    Set NewSnake = New Snake
    NewSnake.Init g, r
End Function
