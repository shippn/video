Public Class vtr_registration
    Dim bi As Date
    Private Sub vtr_registration_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        bi = Date.Today

        TextBox3.Text = String.Format("{0:yyyy-MM-dd}", bi)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim id As Integer
        If (ComboBox1.Text = "アクション") Then
            id = "01"
        ElseIf (ComboBox1.Text = "アニメ") Then
            id = "02"
        ElseIf (ComboBox1.Text = "SF") Then
            id = "03"
        ElseIf (ComboBox1.Text = "戦争") Then
            id = "04"
        ElseIf (ComboBox1.Text = "テレビドラマ") Then
            id = "05"
        ElseIf (ComboBox1.Text = "名作") Then
            id = "06"
        ElseIf (ComboBox1.Text = "ラブロマンス") Then
            id = "07"
        ElseIf (ComboBox1.Text = "ラブストーリー") Then
            id = "08"
        ElseIf (ComboBox1.Text = "ホラー") Then
            id = "09"
        ElseIf (ComboBox1.Text = "アクション(R)") Then
            id = "21"
        ElseIf (ComboBox1.Text = "戦争(R)") Then
            id = "22"
        ElseIf (ComboBox1.Text = "ホラー(R)") Then
            id = "23"
        ElseIf (ComboBox1.Text = "その他") Then
            id = "99"
        End If
        DataGridView1.Rows.Add(id & 10, ComboBox1.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text)

    End Sub


    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        MessageBox.Show("新規タイトルを追加しました", "商品登録完了", MessageBoxButtons.OK)
        main_form.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        main_form.Show()
        Me.Hide()
    End Sub



End Class