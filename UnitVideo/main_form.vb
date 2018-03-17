Public Class main_form
    Dim MyDate As Date

    Dim dbEg As DBEngine
    Public mDb As Database

    Private Sub main_form_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("RentalShop.accdb")
    End Sub



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        loanout_form_main.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        membership_form.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        loanout_form_02.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        vtr_registration.Show()
        Me.Hide()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick

        MyDate = Date.Now
        Label1.Text = String.Format("{0:yyyy年MM月dd日(dddd)HH:mm:ss}", MyDate)
    End Sub



End Class
