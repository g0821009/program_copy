'2014/10/21
'部品加工　プログラムコピー　選択Form
'

Imports System.IO.Path

Public Class Form1

    Public Sub setFiles(ByVal files)
        Module1.files = Nothing

        For Each f As String In files
            ListBox1.Items.Add(New FileList(f))
        Next

        With ListBox1
            'ListBox1を複数選択可に設定
            .SelectionMode = Windows.Forms.SelectionMode.MultiSimple
            'DisplayMemberプロパティを項目の表示に利用する
            .DisplayMember = "FileName"
        End With
    End Sub

    '
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Module1.files = ArrayList.Adapter(ListBox1.SelectedValue).ToArray(GetType(String))

        Dim arr As New List(Of String)
        For Each fl As FileList In ListBox1.SelectedItems
            arr.Add(fl.FullPath)
        Next
        Module1.files = ArrayList.Adapter(arr).ToArray(GetType(String))

        Me.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'フォームの境界線スタイルを「None」にする
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.TopMost = True
        Me.Activate()
    End Sub
End Class

Public Class FileList
    Private myfullpath As String

    Public Sub New(ByVal fullpath As String)
        Me.myfullpath = fullpath
    End Sub

    Public ReadOnly Property FileName() As String
        Get
            Return GetFileName(Me.myfullpath) & vbTab & " (フォルダ名:" & GetFileName(GetDirectoryName(Me.myfullpath)) & ")"
        End Get
    End Property

    Public ReadOnly Property FullPath() As String
        Get
            Return Me.myfullpath
        End Get
    End Property

End Class