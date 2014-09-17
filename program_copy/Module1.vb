'version 1.000 2014.09.17 release

Imports System.IO

Module Module1
    Private fromPath As String
    Private toPath As String
    Private deleteFlag As String = "false"

    Sub Main(args As String())
        ' System.Environment.GetCommandLineArgs() でも可

        If args.Length = 0 Then
            importSetting()
            fileCopy(fromPath, toPath, "*00000-A4*")
            Console.WriteLine("コマンドライン引数はありません。")
        Else
            importSetting()
            fileCopy(fromPath, toPath, args(0))
        End If
        '        Windows.Forms.MessageBox.Show("Hello, world!", "キャプション")

    End Sub

    Sub fileCopy(ByVal sourceDirName As String, ByVal destDirName As String, ByVal searchFileName As String)
        'コピー先のディレクトリ名の末尾に"\"をつける
        If destDirName(destDirName.Length - 1) <> Path.DirectorySeparatorChar Then
            destDirName = destDirName + Path.DirectorySeparatorChar
        End If

        'todo ワイルドカードはサポートされていません
        ' 削除にはfor eachでも使ってください.
        If deleteFlag = "true" Then
            For Each tempFile As String In System.IO.Directory.GetFiles(destDirName)
                System.IO.File.Delete(tempFile)
            Next
        End If

        Dim files As String()
        'コピー元のディレクトリにあるファイルをコピー
        Console.WriteLine("searching filenames...")
        Try
            files = Directory.GetFiles(sourceDirName, searchFileName, SearchOption.AllDirectories)
        Catch ex As Exception
            MsgBox(ex.Message)
            Console.WriteLine("ネットに繋がっていません")
            Environment.Exit(0)
            files = Nothing
        End Try
        Console.WriteLine("search filenames finished")

        Dim f As String
        Dim maxDim As Long = UBound(files)
        '書込むファイルを指定する
        '（2番目の引数をfalseにすることで新規ファイルを作成する）
        Dim sw As StreamWriter = New StreamWriter(System.Windows.Forms.Application.StartupPath & Path.DirectorySeparatorChar & "copy_log.txt", False, System.Text.Encoding.Default)

        For Each f In files
            Dim destFileName As String = destDirName + Path.GetFileName(f)
            'コピー先にファイルが存在しない、
            '存在してもコピー元より更新日時が古い時はコピーする
            Try
                If Not File.Exists(destFileName) OrElse File.GetLastWriteTime(destFileName) < File.GetLastWriteTime(f) Then
                    File.Copy(f, destFileName, True)
                    Console.WriteLine("    copy:" & Path.GetFileName(f))

                    '書込むファイルを指定する
                    '（2番目の引数をfalseにすることで新規ファイルを作成する）
                    Console.WriteLine(System.Windows.Forms.Application.StartupPath)
                    'ファイルに書込む
                    sw.Write(Path.GetFileName(f) & "をコピーしました" & Environment.NewLine)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
                Console.WriteLine("コピーに失敗しました")
                'Environment.Exit(0)
                Exit For
            End Try
        Next
        If files.Count = 0 Then
            'ファイルに書込む
            sw.Write("")

        End If
        sw.Flush()
        sw.Close()
    End Sub

    Private Sub importSetting()
        Dim settingFile As IO.StreamReader
        Dim sa As New ArrayList
        Dim stBuffer As String

        Debug.WriteLine("import Setting start.")
        Try
            settingFile = New IO.StreamReader(".\copy_config.txt", System.Text.Encoding.GetEncoding("Shift-JIS"), False)
            While (settingFile.Peek() >= 0)
                stBuffer = settingFile.ReadLine()
                Debug.WriteLine(stBuffer)
                sa.Add(stBuffer)
            End While
            settingFile.Close()
        Catch ex As Exception
            Debug.WriteLine("file open error:")
            Debug.WriteLine(ex)
        End Try

        If sa.Count > 0 Then
            Dim setting_item() As String
            For Each setting_line As String In sa
                Debug.WriteLine(setting_line)
                ' 空行、行頭2文字が//の場合は次の行へ
                If setting_line.Length = 0 OrElse setting_line.Substring(0, 2) = "//" Then
                    Continue For
                End If
                setting_item = Split(setting_line, "::", 2, CompareMethod.Text)
                Select Case setting_item(0)
                    Case "from"
                        fromPath = setting_item(1)
                    Case "to"
                        toPath = setting_item(1)
                    Case "delete"
                        deleteFlag = setting_item(1)
                    Case Else
                        ' それ以外
                End Select
            Next
        End If
    End Sub

End Module
