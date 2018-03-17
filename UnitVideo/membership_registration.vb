Public Class membership_registration
    Dim MyDate As Date
    Dim dbEg As DBEngine
    Dim mDb As Database
    Dim mRs As Recordset

    Private Sub membership_registration_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("RentalShop.accdb")

        MyDate = Date.Today

        TextBox6.Text = String.Format("{0:yyyy-MM-dd}", MyDate)
    End Sub

    '登録ボタン
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim str As String 'メッセージボックス文字列
        Dim str1 As String
        Dim rndm As Long
        Dim MF As String '性別表示用
        Dim Gender As String '性別DB登録用
        Dim Id As String '身分証種別
        Dim rc As DialogResult

        Dim mRs As Recordset
        Dim OldID As Integer
        Dim NewID As String

        mRs = mDb.OpenRecordset("会員管理")

        mRs.MoveLast()                                      '新規会員番号追加のやつ
        OldID = CInt(mRs.Fields("会員番号").Value) + 1
        NewID = String.Format("{0:D6}", OldID)

        rndm = Rnd() * 1000000

        MF = "未登録"
        If (RadioButton1.Checked = True) Then
            MF = "男"
            Gender = "M"
        ElseIf (RadioButton2.Checked = True) Then
            MF = "女"
            Gender = "F"
        End If

        If (ComboBox1.Text = "免許証") Then   '身分証識別文
            Id = "D"
        Else
            Id = "I"
        End If

        str = "登録日:" & TextBox6.Text & vbCrLf & "----" & vbCrLf & "身分証番号:" & TextBox1.Text & _
                vbCrLf & "身分証種別:" & ComboBox1.Text & vbCrLf & "氏名:" & TextBox2.Text & vbCrLf & _
                "性別:" & MF & vbCrLf & "生年月日:" & ComboBox2.Text & "年" & ComboBox3.Text & _
                "月" & ComboBox4.Text & "日" & vbCrLf & "電話番号:" & TextBox3.Text & vbCrLf & "郵便番号:" & _
                TextBox4.Text & vbCrLf & "住所:" & TextBox5.Text & vbCrLf & "----" & vbCrLf & "以上の情報で登録してもよろしいですか?"

        str1 = "会員番号:" & rndm & vbCrLf & "会員を追加しました。"


        rc = MessageBox.Show(str, "登録内容確認", MessageBoxButtons.YesNo)

        If rc = DialogResult.Yes Then
            MessageBox.Show(str1, "会員登録完了", MessageBoxButtons.OK)
            '会員追加処理このへん
            mRs.AddNew()
            mRs.Fields("会員番号").Value = NewID
            mRs.Fields("氏名").Value = TextBox2.Text
            mRs.Fields("性別").Value = Gender
            mRs.Fields("生年月日").Value = ComboBox2.Text & "/" & ComboBox3.Text & "/" & ComboBox4.Text
            mRs.Fields("入会日").Value = TextBox6.Text
            mRs.Fields("電話番号").Value = TextBox3.Text
            mRs.Fields("郵便番号").Value = TextBox4.Text
            mRs.Fields("住所").Value = TextBox5.Text
            mRs.Fields("身分証明書種別").Value = ComboBox1.Text
            mRs.Fields("身分証明書番号").Value = TextBox1.Text
            mRs.Update()

            Me.Hide()
            main_form.Show()

        End If
        mRs.Close()
        mDb.Close()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        membership_form.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs)
        membership_form.Show()
        Me.Hide()
    End Sub


 
End Class