Public Class loanout_form_01
    Dim MyDate As Date
    Dim dbEg As DBEngine
    Dim mDb As Database
    'Dim mDb As Database
    Dim age As Integer  '年齢用
    Dim sum1 As Integer '料金合計各行ごと
    Dim sum2 As Integer
    Dim sum3 As Integer
    Dim count As Integer '本数カウント


    Public Sub loanout_form_01_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("RentalShop.accdb")

        sum1 = 0 : sum2 = 0 : sum3 = 0 : count = 0 '料金初期化
        Dim mRs As Recordset
        Dim strSQL As String
        strSQL = "select * from 会員管理 where 会員番号 = '" & loanout_form_main.TextBox1.Text & "'"

        mRs = mDb.OpenRecordset(strSQL)

        '年齢計算
        Dim birthday As Date = mRs.Fields("生年月日").Value
        Dim birthY, thisY As Integer
        Dim birthMD, thisMD As Date


        birthY = birthday.Year
        thisY = Today.Year
        birthMD = CDate(Format(birthday, "MM/dd"))
        thisMD = CDate(Format(Today, "MM/dd"))
        If (birthMD > thisMD) Then
            age = thisY - birthY - 1
        Else
            age = thisY - birthY
        End If


        Label15.Text = loanout_form_main.TextBox1.Text
        Label16.Text = mRs.Fields("氏名").Value
        Label17.Text = age

        mRs.Close()


    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")

        MessageBox.Show("貸出処理を完了しました", "貸出処理終了", MessageBoxButtons.OK)
        'ここに書く
        Dim mRs As Recordset    '貸出管理
        Dim mRs1 As Recordset   'ビデオ貸出管理
        Dim str As String
        Dim Inc As String   '貸出数増加用
        Dim day As Integer

        mRs = mDb.OpenRecordset("貸出管理")
        mRs1 = mDb.OpenRecordset("ビデオ貸出管理")

        
        If (TextBox1.Text <> "") Then
            If (ComboBox1.Text = "当日") Then
                MyDate = Date.Today
                day = 0
            ElseIf (ComboBox2.Text = "2泊3日") Then
                MyDate = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 2, MyDate))
                day = 2
            Else
                MyDate = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 6, MyDate))
                day = 6
            End If

            mRs.AddNew()
            mRs.Fields("会員番号").Value = Label15.Text
            mRs.Fields("ビデオ通番").Value = TextBox1.Text
            mRs.Fields("貸出日").Value = Label25.Text
            mRs.Fields("貸出日数").Value = day
            mRs.Fields("返却予定日").Value = MyDate
            mRs.Update()
            'ここまで貸出管理
            str = "select * from ビデオ貸出管理 where ビデオ通番 = '" & TextBox1.Text & "'"
            mRs1 = mDb.OpenRecordset(str)
            Inc = mRs.Fields("貸出件数").Value
            Inc = Inc + 1
            mRs1.Edit()
            mRs1.Fields("貸出件数").Value = Inc
            mRs1.Update()
        End If
        If (TextBox2.Text <> "") Then
            If (ComboBox1.Text = "当日") Then
                MyDate = Date.Today
                day = 0
            ElseIf (ComboBox2.Text = "2泊3日") Then
                MyDate = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 2, MyDate))
                day = 2
            Else
                MyDate = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 6, MyDate))
                day = 6
            End If
            mRs.AddNew()
            mRs.Fields("会員番号").Value = Label15.Text
            mRs.Fields("ビデオ通番").Value = TextBox2.Text
            mRs.Fields("貸出日").Value = Label26.Text
            mRs.Fields("貸出日数").Value = day
            mRs.Fields("返却予定日").Value = MyDate
            mRs.Update()
            'ここまで貸出管理
            str = "select * from ビデオ貸出管理 where ビデオ通番 = '" & TextBox2.Text & "'"
            mRs1 = mDb.OpenRecordset(str)
            Inc = mRs.Fields("貸出件数").Value
            Inc = Inc + 1
            mRs1.Edit()
            mRs1.Fields("貸出件数").Value = Inc
            mRs1.Update()
        End If
        If (TextBox3.Text <> "") Then
            If (ComboBox1.Text = "当日") Then
                MyDate = Date.Today
                day = 0
            ElseIf (ComboBox2.Text = "2泊3日") Then
                MyDate = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 2, MyDate))
                day = 2
            Else
                MyDate = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 6, MyDate))
                day = 6
            End If
            mRs.AddNew()
            mRs.Fields("会員番号").Value = Label15.Text
            mRs.Fields("ビデオ通番").Value = TextBox3.Text
            mRs.Fields("貸出日").Value = Label27.Text
            mRs.Fields("貸出日数").Value = day
            mRs.Fields("返却予定日").Value = MyDate
            mRs.Update()
            'ここまで貸出管理
            str = "select * from ビデオ貸出管理 where ビデオ通番 = '" & TextBox3.Text & "'"
            mRs1 = mDb.OpenRecordset(str)
            Inc = mRs.Fields("貸出件数").Value
            Inc = Inc + 1
            mRs1.Edit()
            mRs1.Fields("貸出件数").Value = Inc
            mRs1.Update()
        End If


        mRs.Close()
        mRs1.Close()

        main_form.Show()
        Me.Hide()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        loanout_form_main.Show()
        Me.Hide()
    End Sub

    Public Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs)
        '変更する(2017
        'MyDate = Date.Today

        'Label25.Text = String.Format("{0:yyyy-MM-dd}", MyDate)

        'Label19.Text = "aaa"
        'Label22.Text = "なし"
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")


        Dim mRs As Recordset    '貸出管理
        Dim mRs1 As Recordset   'ビデオ貸出管理
        Dim vidnum As Integer  'ビデオ番号用
        Dim not18 As Integer
        Dim str As String

        vidnum = TextBox1.Text / 10000 'ビデオ番号
        not18 = TextBox1.Text / 100  '年齢制限用

        mRs = mDb.OpenRecordset("貸出管理")
        str = "select * from ビデオ管理 where ビデオ番号 = '" & vidnum & "'"
        mRs1 = mDb.OpenRecordset(str)

        Label19.Text = mRs1.Fields("タイトル").Value

        If (not18 >= 21) Then  '年齢制限用
            If (not18 = 23) Then
                Label22.Text = "18禁"
                Label22.ForeColor = Color.Red
            Else
                Label22.Text = "15禁"
                Label22.ForeColor = Color.Red
            End If
        Else
            Label23.Text = ""
        End If

        If (not18 > age) Then
            Label18.Text = "年齢制限のため、レンタルできません。" & vbCrLf & _
                           "ビデオID:" & TextBox1.Text
        End If

        Label25.Text = String.Format("{0:yyyy-MM-dd}", Now)

        mRs.Close()
        mRs1.Close()
        mDb.Close()

    End Sub

    Public Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs)
        '変更する(2017
        'MyDate = Date.Today

        'Label26.Text = String.Format("{0:yyyy-MM-dd}", MyDate)

        'Label20.Text = "bbb"
        'Label23.Text = "15禁"
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")


        Dim mRs As Recordset    '貸出管理
        Dim mRs1 As Recordset   'ビデオ貸出管理
        Dim vidnum As Integer  'ビデオ番号用
        Dim not18 As Integer
        Dim str As String

        vidnum = TextBox1.Text / 10000 'ビデオ番号
        not18 = TextBox1.Text / 100  '年齢制限用

        mRs = mDb.OpenRecordset("貸出管理")
        str = "select * from ビデオ管理 where ビデオ番号 = '" & vidnum & "'"
        mRs1 = mDb.OpenRecordset(str)

        Label20.Text = mRs1.Fields("タイトル").Value

        If (not18 >= 21) Then  '年齢制限用
            If (not18 = 23) Then
                Label23.Text = "18禁"
                Label23.ForeColor = Color.Red
            Else
                Label23.Text = "15禁"
                Label23.ForeColor = Color.Red
            End If
        Else
            Label23.Text = ""
        End If

        Label26.Text = String.Format("{0:yyyy-MM-dd}", MyDate)

        If (not18 > age) Then
            Label18.Text = "年齢制限のため、レンタルできません。" & vbCrLf & _
                           "ビデオID:" & TextBox2.Text
        End If

        mRs.Close()
        mRs1.Close()

    End Sub

    Public Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs)
        '変更する(2017
        'MyDate = Date.Today

        'Label27.Text = String.Format("{0:yyyy-MM-dd}", MyDate)

        'Label21.Text = "ccc"
        'Label24.Text = "18禁"
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")

        Dim mRs As Recordset    '貸出管理
        Dim mRs1 As Recordset   'ビデオ貸出管理
        Dim vidnum As Integer  'ビデオ番号用
        Dim not18 As Integer
        Dim str As String

        vidnum = TextBox1.Text / 10000 'ビデオ番号
        not18 = TextBox1.Text / 100  '年齢制限用

        mRs = mDb.OpenRecordset("貸出管理")
        str = "select * from ビデオ管理 where ビデオ番号 = '" & vidnum & "'"
        mRs1 = mDb.OpenRecordset(str)

        Label21.Text = mRs1.Fields("タイトル").Value

        If (not18 >= 21) Then  '年齢制限用
            If (not18 = 23) Then
                Label24.Text = "18禁"
                Label24.ForeColor = Color.Red
            Else
                Label24.Text = "15禁"
                Label24.ForeColor = Color.Red
            End If

        Else
            Label24.Text = ""
        End If

        If (not18 > age) Then
            Label18.Text = "年齢制限のため、レンタルできません。" & vbCrLf & _
                           "ビデオID:" & TextBox3.Text
        End If

        Label27.Text = String.Format("{0:yyyy-MM-dd}", MyDate)

        mRs.Close()
        mRs1.Close()
    End Sub


    ' Private Sub TextBox1_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    '  Label19.Text = "映画1"
    '  Label22.Text = "なし"
    ' Label25.Text = ""
    'End Sub

    'Private Sub TextBox2_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged
    '   Label20.Text = "映画2"
    '  Label23.Text = "なし"
    ' Label26.Text = ""
    'End Sub

    'Private Sub TextBox3_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged
    '    Label21.Text = "映画3"
    '    Label24.Text = "18禁"
    '    Label24.ForeColor = Color.Red
    '    Label27.Text = ""
    'End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Label25.Text = "2017-12-18"
        ' Label28.Text = "2017-12-18"
        'Label31.Text = "200"
        'Label34.Text = "X"
        'Label35.Text = "XXX"

        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")


        Dim mRs As Recordset    '貸出管理
        Dim vidnum As Integer  'ビデオ番号用
        Dim not18 As Integer
        Dim str As String


        vidnum = TextBox1.Text / 100 'ビデオ番号
        vidnum = vidnum.ToString("0000")
        not18 = TextBox1.Text / 100  '年齢制限用
        not18 = not18 Mod 100


        Label19.Text = vidnum

        str = "select * from ビデオ管理 where ビデオ番号 = '" & vidnum.ToString("0000") & "'"
        mRs = mDb.OpenRecordset(str)

        Label19.Text = mRs.Fields("タイトル").Value

        If (not18 >= 21) And (not18 <> 99) Then  '年齢制限用
            If (not18 = 23) Then
                Label22.Text = "18禁"
                Label22.ForeColor = Color.Red
            Else
                Label22.Text = "15禁"
                Label22.ForeColor = Color.Red
            End If
        Else
            Label22.Text = ""
        End If

        If (not18 > age) Then
            Label18.Text = "年齢制限のため、レンタルできません。" & vbCrLf & _
                           "ビデオID:" & TextBox1.Text
        End If

        Label25.Text = String.Format("{0:yyyy-MM-dd}", Now)

        If (ComboBox1.Text = "当日") Then
            Label28.Text = String.Format("{0:yyyy-MM-dd}", MyDate)
            Label31.Text = "200"
            sum1 = 200
            Label35.Text = sum1 + sum2 + sum3

        ElseIf (ComboBox1.Text = "2泊3日") Then
            Label28.Text = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 2, MyDate))
            Label31.Text = "300"
            sum1 = 300
            Label35.Text = sum1 + sum2 + sum3

        ElseIf (ComboBox1.Text = "6泊7日") Then
            Label28.Text = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 6, MyDate))
            Label31.Text = "400"
            sum1 = 400
            Label35.Text = sum1 + sum2 + sum3

        End If


        mRs.Close()
        mDb.Close()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '  Label26.Text = "2017-12-18"
        ' Label29.Text = "2017-12-19"
        'Label32.Text = "300"
        'Label34.Text = "X"
        'Label35.Text = "XXX"

        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")


        Dim mRs As Recordset    '貸出管理
        Dim vidnum As Integer  'ビデオ番号用
        Dim not18 As Integer
        Dim str As String


        vidnum = TextBox2.Text / 100 'ビデオ番号
        vidnum = vidnum.ToString("0000")
        not18 = TextBox2.Text / 10000  '年齢制限用
        not18 = not18 Mod 100

        str = "select * from ビデオ管理 where ビデオ番号 = '" & vidnum.ToString("0000") & "'"
        mRs = mDb.OpenRecordset(str)

        Label20.Text = mRs.Fields("タイトル").Value

        If (not18 >= 21) And (not18 <> 99) Then  '年齢制限用
            If (not18 = 23) Then
                Label23.Text = "18禁"
                Label23.ForeColor = Color.Red
            Else
                Label23.Text = "15禁"
                Label23.ForeColor = Color.Red
            End If
        Else
            Label23.Text = ""
        End If

        If (not18 > age) Then
            Label18.Text = "年齢制限のため、レンタルできません。" & vbCrLf & _
                           "ビデオID:" & TextBox1.Text
        End If

        Label26.Text = String.Format("{0:yyyy-MM-dd}", Now)

        If (ComboBox2.Text = "当日") Then
            Label29.Text = String.Format("{0:yyyy-MM-dd}", MyDate)
            Label32.Text = "200"
            sum1 = 200
            Label35.Text = sum1 + sum2 + sum3

        ElseIf (ComboBox2.Text = "2泊3日") Then
            Label29.Text = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 2, MyDate))
            Label32.Text = "300"
            sum1 = 300
            Label35.Text = sum1 + sum2 + sum3

        ElseIf (ComboBox2.Text = "6泊7日") Then
            Label29.Text = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 6, MyDate))
            Label32.Text = "400"
            sum1 = 400
            Label35.Text = sum1 + sum2 + sum3

        End If

        mRs.Close()
        mDb.Close()


    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        'Label27.Text = "2017-12-18"
        'Label30.Text = "2017-12-25"
        'Label33.Text = "400"
        'Label34.Text = "X"
        'Label35.Text = "XXX"
        dbEg = CreateObject("DAO.DBEngine.120")
        mDb = dbEg.OpenDatabase("C:\Users\S1421181\Documents\RentalShop.accdb")


        Dim mRs As Recordset    '貸出管理
        Dim vidnum As Integer  'ビデオ番号用
        Dim not18 As Integer
        Dim str As String


        vidnum = TextBox3.Text / 100 'ビデオ番号
        vidnum = vidnum.ToString("0000")
        not18 = TextBox3.Text / 10000  '年齢制限用
        not18 = not18 Mod 100


        Label19.Text = vidnum

        str = "select * from ビデオ管理 where ビデオ番号 = '" & vidnum.ToString("0000") & "'"
        mRs = mDb.OpenRecordset(str)

        Label19.Text = mRs.Fields("タイトル").Value

        If (not18 >= 21) And (not18 <> 99) Then  '年齢制限用
            If (not18 = 23) Then
                Label24.Text = "18禁"
                Label24.ForeColor = Color.Red
            Else
                Label24.Text = "15禁"
                Label24.ForeColor = Color.Red
            End If
        Else
            Label24.Text = ""
        End If

        If (not18 > age) Then
            Label18.Text = "年齢制限のため、レンタルできません。" & vbCrLf & _
                           "ビデオID:" & TextBox1.Text
        End If

        Label27.Text = String.Format("{0:yyyy-MM-dd}", Now)

        If (ComboBox3.Text = "当日") Then
            Label30.Text = String.Format("{0:yyyy-MM-dd}", MyDate)
            Label33.Text = "200"
            sum1 = 200
            Label35.Text = sum1 + sum2 + sum3

        ElseIf (ComboBox3.Text = "2泊3日") Then
            Label30.Text = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 2, MyDate))
            Label33.Text = "300"
            sum1 = 300
            Label35.Text = sum1 + sum2 + sum3

        ElseIf (ComboBox3.Text = "6泊7日") Then
            Label30.Text = String.Format("{0:yyyy-MM-dd}", DateAdd(DateInterval.Day, 6, MyDate))
            Label33.Text = "400"
            sum1 = 400
            Label35.Text = sum1 + sum2 + sum3

        End If

        mRs.Close()
        mDb.Close()
    End Sub



End Class