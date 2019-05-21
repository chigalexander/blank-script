Imports Illustrator
Public Class Form_main
    Dim ill As Illustrator.Application
    Dim doc As Illustrator.Document
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ill = CreateObject("Illustrator.Application.CC.2019")
        draw()
        End
    End Sub
    Private Sub draw()
        Dim comStr As String
        'Try
        comStr = Command$()
        ill.SendScriptMessage("Alliance Blank", "blank", comStr)
    End Sub
End Class
