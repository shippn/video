Public Class membership_form

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        membership_registration.Show()
        Me.Hide()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        withdrawal_form_main.Show()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        main_form.Show()
        Me.Hide()
    End Sub
End Class