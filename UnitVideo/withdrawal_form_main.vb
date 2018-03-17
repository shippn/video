Public Class withdrawal_form_main

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        withdrawal_form_01.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        withdrawal_form_02.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        main_form.Show()
        Me.Hide()
    End Sub
End Class