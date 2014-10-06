'version 1.000 2014.09.17 release
'version 2.000 2014.10.03 ファイル転送方法をFTPに変更
'version 2.010 2014.10.06 FTPログイン設定をパラメータ化

Imports System.IO

Module Module1
    Private fromPath As String              ' コピー元フォルダパス
    Private toPath As String                ' コピー先フォルダパス
    Private deleteFlag As String = "false"  ' コピー先フォルダ初期化フラグ
    Private myUri As String = "ftp://localhost/"    ' FTP転送先
    Private ftp_name As String = ""         ' FTP ログイン名
    Private ftp_pass As String = ""         ' FTP ログインパス

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
        Dim myCredential As New System.Net.NetworkCredential(ftp_name, ftp_pass)

        'コピー先のディレクトリ名の末尾に"\"をつける
        If destDirName(destDirName.Length - 1) <> Path.DirectorySeparatorChar Then
            destDirName = destDirName + Path.DirectorySeparatorChar
        End If

        'コピー先フォルダ初期化処理
        If deleteFlag = "true" Then
            '    For Each tmpFile As String In Directory.GetFiles(destDirName)
            '        System.IO.File.Delete(tmpFile)
            '    Next

            'FTPでファイル一覧を取得
            'ファイル一覧を取得するディレクトリのURI
            Dim u As New Uri(myUri)

            'FtpWebRequestの作成
            Dim ftpReq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
            'ログインユーザー名とパスワードを設定
            ftpReq.Credentials = myCredential
            'MethodにWebRequestMethods.Ftp.ListDirectoryDetails("LIST")を設定
            ftpReq.Method = System.Net.WebRequestMethods.Ftp.ListDirectory
            '要求の完了後に接続を閉じない
            ftpReq.KeepAlive = True
            'PASSIVEモードを無効にする
            ftpReq.UsePassive = False

            'FtpWebResponseを取得
            Dim ftpRes As System.Net.FtpWebResponse = CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
            'FTPサーバーから送信されたデータを取得
            Dim sr As New System.IO.StreamReader(ftpRes.GetResponseStream())
            Dim dirList As New System.Collections.Generic.List(Of String)()
            While True
                Dim line As String = sr.ReadLine()
                If line Is Nothing Then
                    Exit While
                End If
                dirList.Add(line)
            End While
            sr.Close()

            'FTPサーバーから送信されたステータスを表示
            Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
            '閉じる
            ftpRes.Close()

            'ファイル一覧にもとづきファイル削除
            For Each fn In dirList
                If fn.Contains(".") Then
                    'ファイル一覧を表示
                    Debug.WriteLine(fn)
                    '削除するファイルのURI
                    u = New Uri(myUri & fn)
                    'FtpWebRequestの作成
                    ftpReq = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
                    'ログインユーザー名とパスワードを設定
                    ftpReq.Credentials = myCredential
                    '要求の完了後に接続を閉じる
                    ftpReq.KeepAlive = False
                    'PASSIVEモードを無効にする
                    ftpReq.UsePassive = False
                    'MethodにWebRequestMethods.Ftp.DeleteFile(DELE)を設定
                    ftpReq.Method = System.Net.WebRequestMethods.Ftp.DeleteFile
                    'FtpWebResponseを取得
                    ftpRes = CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
                    'FTPサーバーから送信されたステータスを表示
                    Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
                End If
            Next fn
            '閉じる
            ftpRes.Close()

            MsgBox("end")
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

                    'ファイル転送部分
                    'File.Copy(f, destFileName, True)

                    '---FTPでのファイル転送部分-------------------
                    'WebClientオブジェクトを作成
                    Dim wc As New System.Net.WebClient()
                    'ログインユーザー名とパスワードを指定
                    wc.Credentials = myCredential
                    'FTPサーバーにアップロード
                    wc.UploadFile(myUri & Path.GetFileName(f), f)
                    '解放する
                    wc.Dispose()
                    '---------------------------------------------

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
                    Case "uri"
                        myUri = setting_item(1)
                    Case "ftpName"
                        ftp_name = setting_item(1)
                    Case "ftpPass"
                        ftp_pass = setting_item(1)
                    Case Else
                        ' それ以外
                End Select
            Next
        End If
    End Sub

End Module
