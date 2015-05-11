Public Class Offset
    Private textFile As String
    Public Sub loadFile(ByVal arg As String)

        '読み込むテキストファイル
        textFile = arg
        '文字コード(ここでは、Shift JIS)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

        'テキストファイルの中身をすべて読み込む
        Console.WriteLine(System.IO.File.ReadAllText(textFile, enc))
        TextBox1.Text = System.IO.File.ReadAllText(textFile, enc)
        Me.Text = System.IO.Path.GetFileName(arg)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        
        '文字コード(ここでは、Shift JIS) 
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

        'ファイルが存在しているときは、上書きする
        System.IO.File.WriteAllText(textFile, TextBox1.Text, enc)

        Me.Close()
    End Sub

    Private Sub Offset_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles Me.KeyDown

    End Sub

    Private Sub Offset_Load(sender As Object, e As EventArgs) Handles Me.Load
        'フォームの境界線スタイルを「None」にする
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.TopMost = True
        Me.Activate()
    End Sub
End Class