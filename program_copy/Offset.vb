Imports System.Drawing

Public Class Offset
    Private textFile As String
    Private deleteflag As Boolean = True

    Public Sub loadFile(ByVal arg As String, ByVal fontsize As String)

        TextBox1.Font = New Font("MS UI Gothic", fontsize)

        '読み込むテキストファイル
        textFile = arg
        '文字コード(ここでは、Shift JIS)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

        'テキストファイルの中身をすべて読み込む
        'Console.WriteLine(System.IO.File.ReadAllText(textFile, enc))
        TextBox1.Text = System.IO.File.ReadAllText(textFile, enc)
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