Public Class loanout_form_03
    Dim MyDate As Date
    Dim day As Date
    Dim mDb As Database
    Dim dbEg As DBEngine


    Private Sub loanout_form_03_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim mRs As Recordset
        Dim mRs1 As Recordset 'なまえ
        Dim str As String
        Dim str1 As String
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("RentalShop.accdb")

        str = "select * from 貸出管理 where ビデオ通番 ='" & loanout_form_02.TextBox1.Text & "'"
        mRs = mDb.OpenRecordset(str)

        str1 = "select * from 会員管理 where 会員番号='" & mRs.Fields("会員番号").Value & "'"
        mRs1 = mDb.OpenRecordset(str1)

        day = mRs.Fields("返却予定日").Value


        GroupBox1.Text = loanout_form_02.TextBox1.Text
        Label5.Text = mRs.Fields("会員番号").Value
        Label6.Text = mRs1.Fields("氏名").Value
        Label7.Text = String.Format("{0:yyyy-MM-dd}", day)

        mRs.Close()
        mDb.Close()

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim str As String '正常時の文
        Dim str1 As String '延滞時の文
        Dim str2 As String '返却完了時の文
        Dim rs As DialogResult

        Dim strnum As String
        Dim mRs As Recordset
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("RentalShop.accdb")

        strnum = "select * from 貸出管理 where ビデオ通番 ='" & loanout_form_02.TextBox1.Text & "'"
        mRs = mDb.OpenRecordset(strnum)


        str = "返却を行います" & vbCrLf & "以下の内容でよろしいですか?" & vbCrLf & "----" & vbCrLf & _
                "ビデオID:" & loanout_form_02.TextBox1.Text

        str1 = "以下の内容を確認後,延滞料を徴収してください" & vbCrLf & "----" & vbCrLf & _
                "'" & mRs.Fields("返却予定日").Value & "'" & "延滞料:300円"

        str2 = "返却完了" & vbCrLf & "----" & vbCrLf & "ビデオID:" & loanout_form_02.TextBox1.Text

        If (mRs.Fields("返却予定日").Value >= Date.Today) Then
            rs = MessageBox.Show(str, "期限内返却", MessageBoxButtons.YesNo)
            If rs = DialogResult.Yes Then
                MessageBox.Show(str2, "返却処理完了", MessageBoxButtons.OK)
                mRs.Delete()
                main_form.Show()
                Me.Hide()
            End If
        Else
            rs = MessageBox.Show(str1, "遅延返却", MessageBoxButtons.YesNo)
            If rs = DialogResult.Yes Then
                MessageBox.Show(str2, "返却処理完了", MessageBoxButtons.OK)
                mRs.Delete()
                main_form.Show()
                Me.Hide()
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        loanout_form_02.Show()
        Me.Hide()
    End Sub


End Class