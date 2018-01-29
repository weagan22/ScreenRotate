Public Class MainForm

    Private Sub RotateScreen(ByVal pRotateDegrees As Rotate.RotateDegress)

        Dim oRotate As Rotate
        Dim rotRet As Rotate.RotateResult

        oRotate = New Rotate()
        rotRet = oRotate.RotateScreen(pRotateDegrees)
        oRotate = Nothing

        Select Case rotRet
            Case Rotate.RotateResult.RotateSuccessful
                ' do nothing, the new screen rotation indicates the positive result
            Case Rotate.RotateResult.RotateRequiresRestart
                MsgBox("Rotate requires restart.")
            Case Else
                MsgBox("Rotate failed.")
        End Select

    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        RotateScreen(CInt(Rnd() * 3))

        'Select Case e.KeyCode
        '    Case Keys.Up
        '        RotateScreen(Rotate.RotateDegress.Degrees_0)
        '    Case Keys.Right
        '        RotateScreen(Rotate.RotateDegress.Degrees_90)
        '    Case Keys.Down
        '        RotateScreen(Rotate.RotateDegress.Degrees_180)
        '    Case Keys.Left
        '        RotateScreen(Rotate.RotateDegress.Degrees_270)
        'End Select

    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Randomize()
    End Sub
End Class
