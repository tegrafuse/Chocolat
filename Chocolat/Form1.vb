Imports System.Security.Cryptography
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Text
Imports System.IO
Imports System.Reflection.Emit

Public Class Form1

    Dim app As String = My.Application.Info.DirectoryPath

    '出力フォルダ取得ボタン
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New FolderBrowserDialog With {
            .Description = "出力場所を指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .ShowNewFolderButton = True
        }
        If ofd.ShowDialog(Me) = DialogResult.OK Then
            TextBox2.Text = ofd.SelectedPath
        End If
    End Sub
    'コマンドコピーボタン
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Clipboard.SetText(TextBox7.Text)
    End Sub

    'URLペーストボタン
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Clipboard.ContainsText() Then
            TextBox1.Text = Clipboard.GetText()
        End If
    End Sub
    'URL消去ボタン
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
    End Sub
    'ダウンロードボタン
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox7.Text = ""
        'サムネイル保存分岐
        If (ComboBox1.SelectedIndex = 2) Then
            TextBox7.Text = "--write-thumbnail --skip-download --convert-thumbnails png "
        Else
            '切り抜き機能
            If (TextBox3.Text <> "" And TextBox4.Text <> "") Then
                TextBox7.Text = "--download-sections *" + TextBox3.Text + "-" + TextBox4.Text + " "
            ElseIf (TextBox3.Text = "" Or TextBox4.Text = "") Then
                If (TextBox3.Text = "" And TextBox4.Text = "") Then
                Else
                    Return
                End If
            End If
            '動画ダウンロード分岐
            If (ComboBox1.SelectedIndex = 0) Then

                '拡張子設定されているか確認
                If (ComboBox5.SelectedIndex = -1) Then
                    Return
                End If

                '動画画質設定処理
                '-1 設定されていない
                '0 設定しない
                '1 最高設定
                '2以上 任意
                If (ComboBox2.SelectedIndex = -1) Then
                    Return
                Else
                    '音なしかどうか
                    If (CheckBox8.Checked = True) Then
                        TextBox7.Text = TextBox7.Text + "-f " + Chr(34) + ComboBox6.Text + Chr(34) + " --merge-output-format " + ComboBox5.Text + " "
                    Else
                        TextBox7.Text = TextBox7.Text + "-f " + Chr(34) + ComboBox6.Text + "+ba" + Chr(34) + " --merge-output-format " + ComboBox5.Text + " "
                    End If
                End If

                'SSL認証をするかどうか
                If (CheckBox1.Checked = True) Then
                    TextBox7.Text = TextBox7.Text + " --no-check-certificate "
                End If

                'クッキーを用いるかどうか
                If (CheckBox3.Checked = True And ComboBox4.SelectedIndex <> -1) Then
                    'クッキーがファイル指定の場合
                    If (ComboBox4.SelectedIndex = 4) Then
                        TextBox7.Text = TextBox7.Text + " --cookies " + TextBox6.Text + " "

                    Else 'クッキーをブラウザから取得する場合
                        If (TextBox6.Text = "") Then
                            TextBox7.Text = TextBox7.Text + " --cookies-from-browser " + ComboBox4.Text + " "
                        Else
                            TextBox7.Text = TextBox7.Text + " --cookies-from-browser " + ComboBox4.Text + ":" + TextBox6.Text + " "
                        End If
                    End If
                End If

                'プロキシー使用するかどうか
                If (CheckBox4.Checked = True And TextBox5.Text <> "") Then
                    TextBox7.Text = TextBox7.Text + " --proxy " + TextBox5.Text + " "
                End If

            End If

            '音声ダウンロード分岐
            If (ComboBox1.SelectedIndex = 1) Then
                TextBox7.Text = TextBox7.Text + "-x --extract-audio --audio-format " + ComboBox5.Text + " "
            End If


            '動画URLがあるかどうか
            If (TextBox1.Text = "") Then
                MsgBox("URLを指定してください")
                Return
            End If

            'プレイリストを無視するかどうか
            If (CheckBox6.Checked = True) Then
                TextBox7.Text = TextBox7.Text + "--yes-playlist "
            Else
                TextBox7.Text = TextBox7.Text + "--no-playlist "
            End If

            '保存場所を指定
            If (TextBox2.Text = "") Then
            Else
                TextBox7.Text = TextBox7.Text + "-o " + TextBox2.Text + "\%(title)s.%(ext)s "
            End If

            'コマンドプロンプトを表示するかどうか
            Dim t As Integer = ProcessWindowStyle.Normal
            If (CheckBox2.Checked = True) Then
                t = ProcessWindowStyle.Hidden
            Else
                t = ProcessWindowStyle.Normal
            End If

            TextBox7.Text = "yt-dlp --console-title --progress " + TextBox7.Text + TextBox1.Text

            If (CheckBox7.Checked = False) Then
                Dim Download As New ProcessStartInfo With {
                .FileName = Environment.GetEnvironmentVariable("ComSpec"),
                .WindowStyle = t,
                .Arguments = "/c " + TextBox7.Text
            }
                Dim d As Process = Process.Start(Download)
            End If
        End If
    End Sub
    '動画情報を取得する
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'プレイリストは無理
        If (TextBox1.Text.Contains("?list") = True) Then
            MsgBox("プレイリストのものは使わないでください")
            Return
        End If

        Dim p As New Process()
        p.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardInput = False
        p.StartInfo.CreateNoWindow = True
        If (TextBox5.Text = "") Then
        End If
        p.StartInfo.Arguments = "/c yt-dlp -F " + TextBox1.Text
        p.Start()
        Dim results As String = p.StandardOutput.ReadToEnd()
        p.WaitForExit()
        'p.Close()

        ComboBox1.SelectedIndex = 0

        Dim lines = results.Replace(vbCrLf, vbLf).Split({vbLf(0), vbCr(0)})
        For i As Integer = 0 To lines.Count - 1
            If lines(i).Contains("[info] Available formats for") Then

                Dim check As New ArrayList()

                For j As Integer = i + 3 To lines.Count - 1
                    '空白なら抜ける
                    If lines(j) = "" Then
                        Continue For
                    End If
                    'ダブリがなかったら、もしくは映像がないものでなければ追加
                    Dim c As Boolean = check.Contains(lines(j).Substring(lines(i + 1).IndexOf("RESOLUTION"), lines(i + 1).IndexOf("FPS") - lines(i + 1).IndexOf("RESOLUTION")).Replace(" ", ""))
                    If (c = False And lines(j).Substring(lines(i + 1).IndexOf("RESOLUTION"), lines(i + 1).IndexOf("FPS") - lines(i + 1).IndexOf("RESOLUTION")).Replace(" ", "") <> "audioonly") Then
                        'コンボボックスに追加
                        ComboBox6.Items.Add(lines(j).Substring(0, lines(i + 1).IndexOf("EXT") - 1).Replace(" ", ""))
                        ComboBox2.Items.Add(lines(j).Substring(lines(i + 1).IndexOf("RESOLUTION"), lines(i + 1).IndexOf("FPS") - lines(i + 1).IndexOf("RESOLUTION")).Replace(" ", ""))
                        'ダブリチェックに追加
                        check.Add(lines(j).Substring(lines(i + 1).IndexOf("RESOLUTION"), lines(i + 1).IndexOf("FPS") - lines(i + 1).IndexOf("RESOLUTION")).Replace(" ", ""))
                    End If
                Next
            End If
        Next
        Console.WriteLine()



    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'コンボボックスリセット
        ComboBox5.Items.Clear()
        '動画の場合
        If (ComboBox1.SelectedIndex = 0) Then
            ComboBox2.Enabled = True
            ComboBox5.Enabled = True

            ComboBox5.Items.Add("mp4")
            ComboBox5.Items.Add("flv")
            ComboBox5.Items.Add("3gp")
            ComboBox5.Items.Add("webm")

            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("")
            ComboBox6.Items.Add("vb")
            ComboBox6.SelectedIndex = 1

            ComboBox2.Items.Clear()
            ComboBox2.Items.Add("設定しない")
            ComboBox2.Items.Add("最高設定")
            ComboBox2.SelectedIndex = 1
        ElseIf (ComboBox1.SelectedIndex = 1) Then
            ComboBox2.Enabled = False
            ComboBox5.Enabled = True

            ComboBox5.Items.Add("mp3")
            ComboBox5.Items.Add("wav")
            ComboBox5.Items.Add("ogg")
            ComboBox5.Items.Add("m4a")
            ComboBox5.Items.Add("acc")

            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("")
            ComboBox6.Items.Add("vb/va")
            ComboBox6.SelectedIndex = 1

            ComboBox2.Items.Clear()
            ComboBox2.Items.Add("設定しない")
            ComboBox2.Items.Add("最高設定")
            ComboBox2.SelectedIndex = 1
        ElseIf (ComboBox1.SelectedIndex = 2) Then
            ComboBox2.Enabled = False
            ComboBox5.Enabled = False
            ComboBox5.Items.Add("png")
        End If
        ComboBox5.SelectedIndex = 0
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If (ComboBox4.SelectedIndex = 4) Then
            Label9.Text = "Cookieのファイルパス"
        Else
            Label9.Text = "Cookieのプロファイル"
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If (CheckBox7.Checked = True) Then
            Button5.Text = "コマンドを生成"
        Else
            Button5.Text = "ダウンロード"
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox6.SelectedIndex = ComboBox2.SelectedIndex
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'コンボボックスリセット
        ComboBox5.Items.Clear()
        '動画の場合
        If (ComboBox1.SelectedIndex = 0) Then
            ComboBox2.Enabled = True
            ComboBox5.Enabled = True

            ComboBox5.Items.Add("mp4")
            ComboBox5.Items.Add("flv")
            ComboBox5.Items.Add("3gp")
            ComboBox5.Items.Add("webm")
            ComboBox5.SelectedIndex = 0

            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("")
            ComboBox6.Items.Add("vb")
            ComboBox6.SelectedIndex = 1

            ComboBox2.Items.Clear()
            ComboBox2.Items.Add("設定しない")
            ComboBox2.Items.Add("最高設定")
            ComboBox2.SelectedIndex = 1
        ElseIf (ComboBox1.SelectedIndex = 1) Then
            ComboBox2.Enabled = False
            ComboBox5.Enabled = True

            ComboBox5.Items.Add("mp3")
            ComboBox5.Items.Add("wav")
            ComboBox5.Items.Add("ogg")
            ComboBox5.Items.Add("m4a")
            ComboBox5.Items.Add("acc")
            ComboBox5.SelectedIndex = 0

            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("")
            ComboBox6.Items.Add("vb/va")
            ComboBox6.SelectedIndex = 1

            ComboBox2.Items.Clear()
            ComboBox2.Items.Add("設定しない")
            ComboBox2.Items.Add("最高設定")
            ComboBox2.SelectedIndex = 1
        ElseIf (ComboBox1.SelectedIndex = 2) Then
            ComboBox2.Enabled = False
            ComboBox5.Enabled = False
            ComboBox5.Items.Add("png")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim meta As New ArrayList()
        meta.Clear()
        Dim sr As New StreamReader(app + "/meta.dontopen", Encoding.GetEncoding("utf-8"))
        Do While sr.EndOfStream = False
            meta.Add(sr.ReadLine)
        Loop
        sr.Close()
        '設定どおりに変更
        CheckBox1.Checked = Boolean.Parse(meta(0))
        CheckBox2.Checked = Boolean.Parse(meta(1))
        CheckBox3.Checked = Boolean.Parse(meta(2))
        CheckBox4.Checked = Boolean.Parse(meta(3))
        CheckBox6.Checked = Boolean.Parse(meta(4))
        CheckBox7.Checked = Boolean.Parse(meta(5))
        CheckBox8.Checked = Boolean.Parse(meta(6))
        ComboBox3.SelectedIndex = Integer.Parse(meta(7))
        ComboBox4.SelectedIndex = Integer.Parse(meta(8))
        TextBox2.Text = meta(9)
        TextBox5.Text = meta(10)
        TextBox6.Text = meta(11)
        Label15.Text = meta(12)
    End Sub
    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim sw As New StreamWriter(app + "/meta.dontopen", False, Encoding.GetEncoding("utf-8"))
        sw.Write(CheckBox1.Checked & vbCrLf & CheckBox2.Checked & vbCrLf & CheckBox3.Checked & vbCrLf & CheckBox4.Checked & vbCrLf & CheckBox6.Checked & vbCrLf & CheckBox7.Checked & vbCrLf & CheckBox8.Checked & vbCrLf & ComboBox3.SelectedIndex & vbCrLf & ComboBox4.SelectedIndex & vbCrLf & TextBox2.Text & vbCrLf & TextBox5.Text & vbCrLf & TextBox6.Text & vbCrLf & Label15.Text & vbCrLf)
        sw.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim up As New Process()
        up.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
        up.StartInfo.UseShellExecute = False
        up.StartInfo.RedirectStandardOutput = True
        up.StartInfo.RedirectStandardInput = False
        up.StartInfo.CreateNoWindow = True
        up.StartInfo.Arguments = "/c yt-dlp -U"
        up.Start()
        Dim ve As String = up.StandardOutput.ReadToEnd()
        up.WaitForExit()
        up.Close()
        If (ve.Contains("Current version:")) Then
            MsgBox("Successfully updated")
        Else
            MsgBox("No update")
        End If

        Dim vr As New Process()
        vr.StartInfo.FileName = Environment.GetEnvironmentVariable("ComSpec")
        vr.StartInfo.UseShellExecute = False
        vr.StartInfo.RedirectStandardOutput = True
        vr.StartInfo.RedirectStandardInput = False
        vr.StartInfo.CreateNoWindow = True
        vr.StartInfo.Arguments = "/c yt-dlp --version "
        vr.Start()
        Dim rs As String = vr.StandardOutput.ReadToEnd()
        vr.WaitForExit()
        vr.Close()
        Label15.Text = rs
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MsgBox("Twitter : @forda368906" & vbCrLf & "Discord : xheirloom")
    End Sub
End Class