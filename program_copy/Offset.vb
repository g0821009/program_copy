Imports System.Drawing

Public Class Offset
    Private textFile As String
    Private deleteflag As Boolean = True

    Public Sub loadFile(ByVal arg As String, ByVal fontsize As String, ByVal L2KFlag As String)

        TextBox1.Font = New Font("MS UI Gothic", fontsize)

        '読み込むテキストファイル
        textFile = arg
        '文字コード(ここでは、Shift JIS)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

        'テキストファイルの中身をすべて読み込む
        'Console.WriteLine(System.IO.File.ReadAllText(textFile, enc))
        TextBox1.Text = System.IO.File.ReadAllText(textFile, enc)
        '2016-03-08追加 ファナックのシステムの違いにより、LをKに置換する場合の処理,TOOLの文字列はそのまま
        If (L2KFlag = "true") Then
            TextBox1.Text = TextBox1.Text.Replace("L0", "K0")
            TextBox1.Text = TextBox1.Text.Replace("L1", "K1")
            TextBox1.Text = TextBox1.Text.Replace("L2", "K2")
            TextBox1.Text = TextBox1.Text.Replace("L3", "K3")
            TextBox1.Text = TextBox1.Text.Replace("L4", "K4")
            TextBox1.Text = TextBox1.Text.Replace("L5", "K5")
            TextBox1.Text = TextBox1.Text.Replace("L6", "K6")
            TextBox1.Text = TextBox1.Text.Replace("L7", "K7")
            TextBox1.Text = TextBox1.Text.Replace("L8", "K8")
            TextBox1.Text = TextBox1.Text.Replace("L9", "K9")

        End If
        'windowタイトルのテキストのこと
        Me.Text = System.IO.Path.GetFileName(arg)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '文字コード(ここでは、Shift JIS) 
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

        'ファイルが存在しているときは、上書きする
        System.IO.File.WriteAllText(textFile, TextBox1.Text, enc)
        deleteflag = False

        Me.Close()
    End Sub

    Private Sub Offset_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If deleteflag Then
            System.IO.File.Delete(textFile)
        End If
        Module1.fontsize = TextBox1.Font.Size
    End Sub

    Private Sub Offset_Load(sender As Object, e As EventArgs) Handles Me.Load
        'フォームの境界線スタイルを「None」にする
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = Windows.Forms.FormWindowState.Maximized
        Me.TopMost = True
        Me.Activate()
    End Sub

    '文字サイズダウン
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Font.Size > 10 Then
            TextBox1.Font = New Font("MS UI Gothic", TextBox1.Font.Size - 4)
        End If
    End Sub

    '文字サイズアップ
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Font.Size < 70 Then
            TextBox1.Font = New Font("MS UI Gothic", TextBox1.Font.Size + 4)
        End If
    End Sub

End Class