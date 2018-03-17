Public Class withdrawal_form_02
    Dim dbEg As DBEngine
    Dim mDb As Database

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim str As String
        Dim rc As DialogResult
        Dim mRs As Recordset
        Dim strSQL As String
        Dim Gender As String '性別
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("RentalShop.accdb")



        mRs = mDb.OpenRecordset("会員管理")

        strSQL = "select * from 会員管理 where 身分証明書番号 ='" & TextBox1.Text & "'"
        mRs = mDb.OpenRecordset(strSQL)


        If (mRs.Fields("性別").Value = "M") Then  '性別文
            Gender = "男"
        Else
            Gender = "女"
        End If

        str = "会員番号:" & mRs.Fields("会員番号").Value & vbCrLf & "----" & vbCrLf & "身分証番号:" & TextBox1.Text & _
                vbCrLf & "身分証種別" & ComboBox1.Text & vbCrLf & "氏名:" & TextBox2.Text & vbCrLf & _
                "性別" & Gender & vbCrLf & "生年月日:" & mRs.Fields("生年月日").Value & vbCrLf & _
                "電話番号" & mRs.Fields("電話番号").Value & vbCrLf & "郵便番号:" & mRs.Fields("郵便番号").Value & _
                 vbCrLf & "住所:" & mRs.Fields("住所").Value & vbCrLf & "----" & vbCrLf & "以上の情報を解約してもよろしいですか?"

        rc = MessageBox.Show(str, "会員情報確認", MessageBoxButtons.YesNo)

        If rc = DialogResult.Yes Then
            MessageBox.Show("退会処理を行いました", "会員登録完了", MessageBoxButtons.OK)

            mRs.Delete()

            Me.Hide()
            main_form.Show()
        End If
        mRs.Close()
        mDb.Close()
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        withdrawal_form_main.Show()
        Me.Hide()
    End Sub


End Class