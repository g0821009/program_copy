'version 1.000 2014.09.17 release
'version 2.000 2014.10.03 ファイル転送方法をFTPに変更
'version 2.010 2014.10.06 FTPログイン設定をパラメータ化
'version 2.011 2014.10.06 動作確認用msgBox削除
'version 2.012 2014.10.06 引数なしの場合の動作確認処理をコメントアウト
'version 2.013 2014.10.06 コピー仕様をコメントアウトしFTP転送のみとした
'version 2.020 2014.10.21 複数ファイルが検索された場合、選択できるようにした
'version 2.100 2014.12.12 ハードコーティング等の削除・細かい修正
'version 2.200 2015.02.24 試験用からの変更前fix
'version 3.200 2015.10.06 安定版
'version 3.300 2016.03.08 ファナックの仕様が異なるため、繰り返し文字"L"を"K"に置換するかフラグで管理する
'version 3.400 2017.02.08 FTP送信時にデフォルトだとバイナリ形式で送信するため必ずアスキー（テキスト）で送信するように変更

Imports System.IO
Imports System.Data.SQLite

Module Module1

    Private fromPath As String                      ' コピー元フォルダパス
    Private toPath As String                        ' コピー先フォルダパス
    Private deleteFlag As String = "false"          ' コピー先フォルダ初期化フラグ
    Private myUri As String = "ftp://localhost/"    ' FTP転送先
    Private ftp_name As String = ""                 ' FTP ログイン名
    Private ftp_pass As String = ""                 ' FTP ログインパス
    Public fontsize As String = "14"                ' 編集画面フォントサイズ
    Private L2KFlag As String = "false"               ' Offsetクラスで文字LをKに置換するフラグ
    Private ini_comment As String = ""              ' コンフィグコメント
    Public files As String() = Nothing              ' 検索結果変数
    Private db_path As String = ".\filelist.db"     ' 検索用sqliteDBファイルパス

    Private Connection As SQLiteConnection          ' SQLiteコネクション
    Private Command As SQLiteCommand                ' SQLiteクエリ

    Sub Main(args As String())

        Connection = New SQLiteConnection
        Command = Connection.CreateCommand

        '接続文字列を設定
        'Connection.ConnectionString = "Data Source=" & db_path & ";New=False;Compress=false;"
        Connection.ConnectionString = "Data Source=" & My.Application.Info.DirectoryPath & "\filelist.db;New=False;Compress=false;"


        'パスワードをセット
        'Connection.SetPassword("password")
        Try
            'オープン
            Connection.Open()
        Catch sqlex As SQLiteException
        Catch ex As Exception
        End Try

        '引数は System.Environment.GetCommandLineArgs() でも可
        'Proxyを使わない設定
        System.Net.WebRequest.DefaultWebProxy = Nothing

        importSetting()
        If args.Length = 0 Then
            '自分自身のAssemblyを取得
            Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            'バージョンの取得
            Dim ver As System.Version = asm.GetName().Version
            'デバック用
            'fileCopy(fromPath, "%444%")
            Console.WriteLine("バージョン：" & ver.ToString)
            Console.WriteLine("備考：L2Kフラグ追加 2016-03-08")
            Console.WriteLine("コマンドライン引数はありません。")
            Console.WriteLine("-i         :DB初期化")
            Console.WriteLine("-p [図番]  :[図番]PDFがあればadobe Readerで表示")
            Console.WriteLine("[%文字列%] :文字列を検索 %はワイルドカード")
            MsgBox("終了します。", MsgBoxStyle.MsgBoxSetForeground)
        ElseIf args(0) = "-i" Then
            Console.WriteLine("Delete DB table start")
            Command = Connection.CreateCommand
            Command.CommandText = "DELETE from sawairi_files"
            Command.ExecuteNonQuery()
            Console.WriteLine("Delete DB table finish")
            ReloadDB(fromPath)
        ElseIf args(0) = "-p" And args.Length = 2 Then
            showPDF(args(1))
        ElseIf args.Length = 2 Then
            showPDF(args(1))
            fileCopy(fromPath, args(0))
        ElseIf args.Length = 1 Then
            fileCopy(fromPath, args(0))
        Else
            Console.WriteLine("不明なオプション")
        End If

        Command.Dispose()
        Try
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
                Connection.Dispose()
            End If
        Catch sqlex As SQLiteException
            Debug.WriteLine(sqlex.Message)
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try

        exportSetting()

    End Sub

    '2014/10/06 destDirNameを引数から消した、しかし今後コピーを復活させるならoverloadすべきなのかもしれない. 設定ファイルの仕様はそのままにしておく
    Sub fileCopy(ByVal sourceDirName As String, ByVal searchFileName As String)
        ' FTP接続用使い回しCredential
        Dim myCredential As New System.Net.NetworkCredential(ftp_name, ftp_pass)

        'コピー先のディレクトリ名の末尾に"\"をつける
        'If destDirName(destDirName.Length - 1) <> Path.DirectorySeparatorChar Then
        '    destDirName = destDirName + Path.DirectorySeparatorChar
        'End If

        '------------------------------------------------------------------------
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
            'FtpWebResponseのCloseメソッドが呼び出された時に、ログアウトしてサーバーとの接続を切断
            ftpReq.KeepAlive = False
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

        End If

        '------------------------------------------------------------------------
        'Dim files As String() = Nothing
        ''コピー元のディレクトリにあるファイルをコピー
        'Console.WriteLine("searching filenames...")
        'Try
        '    files = System.IO.Directory.GetFiles(sourceDirName, searchFileName, SearchOption.AllDirectories)
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        '    Console.WriteLine("ネットに繋がっていません")
        '    Environment.Exit(0)
        '    files = Nothing
        'End Try
        'Console.WriteLine("search filenames finished")

        'If files.Length = 0 Then
        '    MsgBox("ファイルが見つかりません")
        '    Exit Sub
        'End If

        Dim DataReader As SQLiteDataReader

        'コマンド作成
        Command = Connection.CreateCommand
        'SQL作成
        Command.CommandText = "SELECT * FROM sawairi_files where FileName like '" & searchFileName & ".NC'"

        'データリーダーにデータ取得
        DataReader = Command.ExecuteReader

        Dim filearray As New ArrayList
        'データを全件出力
        Do Until Not DataReader.Read
            filearray.Add(DataReader.Item("FullPath").ToString)
        Loop
        files = filearray.ToArray(GetType(String))

        '破棄
        DataReader.Close()
        Command.Dispose()

        '2015/05/11 PDF表示追加に伴いコメントアウト
        'Try
        '    If Connection.State = ConnectionState.Open Then
        '        Connection.Close()
        '        Connection.Dispose()
        '        End If
        'Catch sqlex As SQLiteException
        '    Debug.WriteLine(sqlex.Message)
        'Catch ex As Exception
        '    Debug.WriteLine(ex.Message)
        'End Try

        '------------------------------------------------------------------------

        ' ファイル名チェック、変換
        fileCheck(files)

        Dim f As String
        Dim maxDim As Long = UBound(files)

        'logファイルを指定する
        '（2番目の引数をfalseにすることで新規ファイルを作成する）
        Dim sw As StreamWriter = New StreamWriter(System.Windows.Forms.Application.StartupPath & Path.DirectorySeparatorChar & "copy_log.txt", False, System.Text.Encoding.Default)

        For Each f In files

            Try
                'Dim destFileName As String = destDirName + Path.GetFileName(f)
                'コピー先にファイルが存在しない、
                '存在してもコピー元より更新日時が古い時はコピーする
                'If Not File.Exists(destFileName) OrElse File.GetLastWriteTime(destFileName) < File.GetLastWriteTime(f) Then

                'ファイル転送部分
                'File.Copy(f, destFileName, True)

                '---FTPでのファイル転送部分-------------------
                'WebClientオブジェクトを作成
                Dim wc As New System.Net.WebClient()
                'ログインユーザー名とパスワードを指定
                wc.Credentials = myCredential
                'FTPサーバーにアップロード()
                'getFileNameはファイル名を変える関数
                Dim sendFilePath As String = mygetFileName(f)
                If System.IO.File.Exists(sendFilePath) Then
                    wc.UploadFile(myUri & Path.GetFileName(sendFilePath), sendFilePath)
                    '一時ファイルを消す
                    System.IO.File.Delete(sendFilePath)
                    'ファイルに書込む
                    sw.Write(Path.GetFileName(sendFilePath) & "をコピーしました" & Environment.NewLine)
                    Console.WriteLine("    copy:" & Path.GetFileName(sendFilePath))
                End If
                '解放する
                wc.Dispose()
                '---------------------------------------------

            Catch ex As Exception
                MsgBox(ex.Message)
                Console.WriteLine("コピーに失敗しました")
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

    Private Sub exportSetting()
        Dim settingFile As IO.StreamWriter
        Try
            '既に存在するテキストに追加する場合は第２引数をTrueにする。
            settingFile = New IO.StreamWriter(".\copy_config.txt", False, System.Text.Encoding.GetEncoding("Shift-JIS"))
            settingFile.Write(ini_comment)
            settingFile.WriteLine("from::" & fromPath)
            settingFile.WriteLine("delete::" & deleteFlag)
            settingFile.WriteLine("uri::" & myUri)
            settingFile.WriteLine("ftpName::" & ftp_name)
            settingFile.WriteLine("ftpPass::" & ftp_pass)
            settingFile.WriteLine("DBPath::" & db_path)
            settingFile.WriteLine("fontsize::" & fontsize)
            settingFile.WriteLine("L2KFlag::" & L2KFlag)
            settingFile.Close()
        Catch ex As Exception
            Debug.WriteLine("Error:")
            Debug.WriteLine(ex)
        End Try
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
                    ini_comment = ini_comment & setting_line.ToString & System.Environment.NewLine
                    Continue For
                End If
                setting_item = Split(setting_line, "::", 2, CompareMethod.Text)
                Select Case setting_item(0)
                    Case "from"         ' コピー元フォルダパス
                        fromPath = setting_item(1)
                    Case "to"           ' コピー先フォルダパス
                        toPath = setting_item(1)
                    Case "delete"       ' コピー・転送先フォルダ初期化フラグ
                        deleteFlag = setting_item(1)
                    Case "uri"          ' FTP接続先
                        myUri = setting_item(1)
                    Case "ftpName"      ' FTPログイン名
                        ftp_name = setting_item(1)
                    Case "ftpPass"      ' FTPログインパスワード
                        ftp_pass = setting_item(1)
                    Case "DBPath"      ' DB path
                        db_path = setting_item(1)
                    Case "fontsize"      ' 編集画面フォントサイズ
                        fontsize = setting_item(1)
                    Case "L2KFlag"
                        L2KFlag = setting_item(1)
                    Case Else
                        ' それ以外
                End Select
            Next
        End If
    End Sub

    ' コピーファイル名の重複チェックと処理
    Private Sub fileCheck(ByRef files As String())

        ' todo 重複ファイルチェック
        '未実装
        For Each f As String In files
            Debug.Write("fileCheck:")
            Debug.WriteLine(f)
        Next

        If files.Length > 1 Then
            Dim hoge As New Form1
            hoge.setFiles(files)
            System.Windows.Forms.Application.Run(hoge)
            files = Module1.files
        ElseIf files.Length = 0 Then
            MsgBox(".NC file not found.")
        End If

    End Sub

    '転送するファイル名変更
    '転送前プログラム編集　2015/05/08
    Private Function mygetFileName(ByRef f As String) As String
        Dim filePath As String = My.Application.Info.DirectoryPath.ToString()
        filePath = filePath + Path.DirectorySeparatorChar + Path.GetFileName(f)
        'ファイルをローカルにコピー(上書き)
        System.IO.File.Copy(f, filePath, True)

        Dim ff As String = Path.GetFileName(f)

        Debug.WriteLine(ff.Substring(0, ff.IndexOf("-")))
        'getFileName = Path.GetFileName(ff.Substring(0, ff.IndexOf("-")) & Path.GetExtension(f))
        'ファイル名を変更しない 2015/02/24

        '転送前プログラム編集画面表示
        Dim hoge As New Offset
        hoge.loadFile(filePath, fontsize, L2KFlag)
        System.Windows.Forms.Application.Run(hoge)

        mygetFileName = filePath
    End Function

    Private Sub ReloadDB(ByVal mypath As String)

        'コマンド作成
        Command = Connection.CreateCommand

        Dim files As DirectoryInfo = New DirectoryInfo(mypath)
        Dim SQLtransaction As SQLite.SQLiteTransaction
        Dim index As Integer = 1

        Console.WriteLine(System.DateTime.Now & ": start get file list in " + mypath)
        ' トランザクション開始
        SQLtransaction = Connection.BeginTransaction()

        Try
            For Each f As FileInfo In files.EnumerateFiles("*-*-*.*", SearchOption.AllDirectories)
                Console.WriteLine(f.Name)
                'Dim zuban As String = f.Name.Substring(f.Name.IndexOf("-",0)+1,
                'If System.Text.RegularExpressions.Regex.IsMatch(f.Name, "") Then
                'End If

                'リビジョン抜き出し
                Dim rev As String = ""
                If f.Name.Length - f.Name.LastIndexOf("-") > 3 Then
                    '2015/06/25 追加した指示書のファイル名が仕様でなかったため、リビジョンの取り方を変更
                    'rev = f.Name.Substring(f.Name.LastIndexOf("-") + 1, 3)
                    Dim tmpname As String = f.Name.Replace("_", "-")
                    rev = tmpname.Substring(tmpname.LastIndexOf("-") + 1, 3)
                    If System.Text.RegularExpressions.Regex.IsMatch(rev, "^..(\.|-|\s)") Then
                        rev = rev.Substring(0, 2)
                    Else
                        rev = "00"
                    End If
                Else
                    rev = f.Name.Substring(f.Name.LastIndexOf("-") + 1, 2)
                End If

                '図番抜き出し
                Dim zuban As String
                Dim tmp_zuban As String()
                tmp_zuban = Split(f.Name, "-", -1, CompareMethod.Binary)
                If tmp_zuban.Length >= 2 Then
                    zuban = tmp_zuban(1) & "-" & tmp_zuban(2)
                    Console.WriteLine(zuban)
                Else
                    zuban = f.Name
                End If
                If f.Extension.Length > 0 Then
                    zuban = zuban.Replace(f.Extension, "")
                End If

                Command.CommandText = "INSERT into sawairi_files(FullPath,FileName,Directory,Extension,Zuban,Revision) values (""" & f.FullName & """, """ & f.Name & """, """ & f.DirectoryName & """, """ & f.Extension & """, """ & zuban & """, """ & rev & """)"
                'Console.WriteLine(Command.CommandText)
                Command.ExecuteNonQuery()
                index = index + 1

            Next f
            'トランザクションコミット
            SQLtransaction.Commit()
        Catch ex As UnauthorizedAccessException
            Console.WriteLine(ex.Message)
        Catch ex As Exception
            Console.WriteLine("---errer msg----------")
            Console.WriteLine(ex)
            Console.WriteLine("---errer msg end------")
        End Try

        Console.WriteLine(System.DateTime.Now & ": finish get file list.")
        System.Threading.Thread.Sleep(1000)
    End Sub

    Private Sub showPDF(searchFileName As String)
        Dim DataReader As SQLiteDataReader

        'コマンド作成
        Command = Connection.CreateCommand
        'SQL作成
        '2015/06/25 ZUBANの検索を完全一意から前方一致のlikeに変更、追加した指示書の命名規則が仕様でなかったため
        'Command.CommandText = "SELECT FullPath, Zuban, Extension, OrderNum FROM sawairi_files INNER JOIN revision_order ON sawairi_files.Revision = revision_order.Revision WHERE Extension = '.pdf' AND Zuban = '" & searchFileName & "' ORDER BY OrderNum DESC"
        Command.CommandText = "SELECT FullPath, Zuban, Extension, OrderNum FROM sawairi_files INNER JOIN revision_order ON sawairi_files.Revision = revision_order.Revision WHERE Extension = '.pdf' AND Zuban LIKE '" & searchFileName & "%' ORDER BY OrderNum DESC"

        'データリーダーにデータ取得
        DataReader = Command.ExecuteReader

        Dim filearray As New ArrayList
        'データを全件出力
        Do Until Not DataReader.Read
            filearray.Add(DataReader.Item("FullPath").ToString)
        Loop
        files = filearray.ToArray(GetType(String))

        'PDFファイルが見つかれば
        If files.Length > 0 Then
            'PDFを表示する
            System.Diagnostics.Process.Start("AcroRd32.exe", """" & files(0).ToString & """")
        End If

        '破棄
        DataReader.Close()
        Command.Dispose()

        '2015/05/11 PDF表示追加に伴いコメントアウト
        'Try
        '    If Connection.State = ConnectionState.Open Then
        '        Connection.Close()
        '        Connection.Dispose()
        '    End If
        'Catch sqlex As SQLiteException
        '    Debug.WriteLine(sqlex.Message)
        'Catch ex As Exception
        '    Debug.WriteLine(ex.Message)
        'End Try

    End Sub

End Module
