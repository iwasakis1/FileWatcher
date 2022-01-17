Public Class Form1









    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs) Handles TextBox1.DragEnter
        'ファイル形式の場合のみ、ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextBox1_DragDrop(sender As Object, e As DragEventArgs) Handles TextBox1.DragDrop
        'ドラッグされたファイル・フォルダのパスを格納します。
        Dim strFileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'ファイルの存在確認を行い、ある場合にのみ、
        'テキストボックスにパスを表示します。
        '（この処理でフォルダを対象外にしています。）
        If System.IO.File.Exists(strFileName(0).ToString) = True Then
            Me.TextBox1.Text = strFileName(0).ToString
        ElseIf System.IO.Directory.Exists(strFileName(0).ToString) = True Then
            Me.TextBox1.Text = strFileName(0).ToString
        End If
    End Sub

    Private watcher As System.IO.FileSystemWatcher = Nothing
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()


        'ListBox1.Items.Add("aaaaa")

        If Not (watcher Is Nothing) Then
            Return
        End If

        watcher = New System.IO.FileSystemWatcher
        Try
            '監視するディレクトリを指定
            watcher.Path = TextBox1.Text
            '最終アクセス日時、最終更新日時、ファイル、フォルダ名の変更を監視する
            watcher.NotifyFilter = System.IO.NotifyFilters.LastAccess Or
                System.IO.NotifyFilters.LastWrite Or
                System.IO.NotifyFilters.FileName Or
                System.IO.NotifyFilters.DirectoryName
            'すべてのファイルを監視
            watcher.Filter = ""
            watcher.IncludeSubdirectories = True
            'UIのスレッドにマーシャリングする
            'コンソールアプリケーションでの使用では必要ない
            watcher.SynchronizingObject = Me

            'イベントハンドラの追加
            AddHandler watcher.Changed, AddressOf watcher_Changed
            AddHandler watcher.Created, AddressOf watcher_Changed
            AddHandler watcher.Deleted, AddressOf watcher_Changed
            AddHandler watcher.Renamed, AddressOf watcher_Renamed

            '監視を開始する
            watcher.EnableRaisingEvents = True
        Catch ex As Exception
            MsgBox(ex.ToString)
            ListBox1.Items.Add("監視を開始出来ませんでした。")
            If Not (watcher Is Nothing) Then
                watcher.Dispose()
                watcher = Nothing
            End If
            Exit Sub
        End Try
        ListBox1.Items.Add("監視を開始しました。")
        ListBox1.SelectedIndex = ListBox1.Items.Count() - 1

    End Sub

    'イベントハンドラ
    Private Sub watcher_Changed(ByVal source As System.Object,
        ByVal e As System.IO.FileSystemEventArgs)
        Select Case e.ChangeType
            Case System.IO.WatcherChangeTypes.Changed
                ListBox1.Items.Add(("ファイル 「" + e.FullPath +
                    "」が変更されました。"))
            Case System.IO.WatcherChangeTypes.Created
                ListBox1.Items.Add(("ファイル 「" + e.FullPath +
                    "」が作成されました。"))
            Case System.IO.WatcherChangeTypes.Deleted
                ListBox1.Items.Add(("ファイル 「" + e.FullPath +
                    "」が削除されました。"))
        End Select
        ListBox1.SelectedIndex = ListBox1.Items.Count() - 1

    End Sub

    Private Sub watcher_Renamed(ByVal source As System.Object,
        ByVal e As System.IO.RenamedEventArgs)
        ListBox1.Items.Add(("ファイル 「" + e.FullPath +
            "」の名前が変更されました。"))
        ListBox1.SelectedIndex = ListBox1.Items.Count() - 1
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If watcher Is Nothing OrElse Not watcher.EnableRaisingEvents Then
            ListBox1.Items.Add("監視していません。")
            ListBox1.SelectedIndex = ListBox1.Items.Count() - 1
            Exit Sub
        End If
        '監視を終了
        If Not (watcher Is Nothing) Then
            watcher.EnableRaisingEvents = False
            watcher.Dispose()
            watcher = Nothing
        End If
        ListBox1.Items.Add("監視を終了しました。")
        ListBox1.SelectedIndex = ListBox1.Items.Count() - 1
    End Sub
End Class
