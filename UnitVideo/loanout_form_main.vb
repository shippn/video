Public Class loanout_form_main

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        loanout_form_01.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        main_form.Show()
        Me.Hide()
    End Sub

End Class