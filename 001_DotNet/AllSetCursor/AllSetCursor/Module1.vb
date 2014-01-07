Imports System.Windows.Forms
Imports System.Drawing
Imports System.Runtime.InteropServices

Module Module1

    Public Const SPI_SETCURSORS = 87
    Public Const SPIF_SENDWININICHANGE = &H2

    <DllImport("user32", CharSet:=CharSet.Auto)> _
    Public Function SystemParametersInfo( _
            ByVal intAction As Integer, _
            ByVal intParam As Integer, _
            ByVal strParam As String, _
            ByVal intWinIniFlag As Integer) As Integer
    End Function

    Sub Main()

        Dim targetPath As String = Nothing
        Dim FolderBrowserDialog1 As New FolderBrowserDialog()
        FolderBrowserDialog1.Description = "フォルダを選択してください"

        ' Default -> SpecialFolder.Desktop
        FolderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.SelectedPath = "C:\"

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            targetPath = FolderBrowserDialog1.SelectedPath
        End If

        If targetPath Is Nothing Then
            Exit Sub
        End If

        FolderBrowserDialog1.Dispose()

        Dim files As String() = System.IO.Directory.GetFiles(targetPath, "*", System.IO.SearchOption.TopDirectoryOnly)

        For Each file As String In files
            setCursor(file)
        Next

        ' 設定変更メッセージ送信
        SystemParametersInfo(SPI_SETCURSORS, 0, Nothing, SPIF_SENDWININICHANGE)

        MsgBox("設定完了！")

    End Sub

    Sub setCursor(ByVal curFile As String)

        If InStr(curFile, "通常の選択") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("Arrow", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "ヘルプの選択") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("Help", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "バックグラウンドで作業中") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("AppStarting", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "待ち状態") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("Wait", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "領域選択") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("Crosshair", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "テキスト選択") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("Ibeam", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "手書き") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("NWPen", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "利用不可") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("No", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "上下に拡大縮小") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("SizeNS", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "左右に拡大縮小") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("SizeWE", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "斜めに拡大縮小1") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("SizeNWSE", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "斜めに拡大縮小2") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("SizeNESW", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "移動") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("SizeAll", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "代替選択") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("UpArrow", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        ElseIf InStr(curFile, "リンクの選択") > 0 Then
            Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\Cursors", True)
            regkey.SetValue("Hand", curFile, Microsoft.Win32.RegistryValueKind.ExpandString)
            regkey.Close()
        End If

    End Sub

End Module
