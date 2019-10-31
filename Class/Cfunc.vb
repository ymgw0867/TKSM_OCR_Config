Public Class Cfunc
    '----------------------------------------
    '　DSNファイルの展開
    '----------------------------------------
    Public Function fncGetConnect(ByVal sDsnPath As String) As String
        Dim sReadBuf As String
        Dim sCnnStr As String              '接続文字列格納用

        sCnnStr = ""

        'DSNファイルを開く
        Dim fileReader = My.Computer.FileSystem.OpenTextFileReader(sDsnPath)
        If fileReader.EndOfStream = True Then
            fileReader.Close()
            fncGetConnect = ""
            Exit Function
        End If

        '１行読み飛ばし
        sReadBuf = fileReader.ReadLine()
        Do While (fileReader.EndOfStream <> True)
            sReadBuf = fileReader.ReadLine()
            sCnnStr = sCnnStr + sReadBuf & ";"
        Loop
        fileReader.Close()

        If frmMain.txtPassWord.Text <> "" Then
            'パスワードが設定されている場合のみ、パスワードを追加
            sCnnStr = sCnnStr & "PWD=" & frmMain.txtPassWord.Text & ";"
        End If


        '接続文字列を返す
        fncGetConnect = sCnnStr

    End Function

    Public Function GetPath()
        '----------------------------------------------------------------
        '   パス情報取得
        '----------------------------------------------------------------
        Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\FKDL", False)
        GetPath = ""
        If regkey Is Nothing Then
            Exit Function
        End If

        '展開して取得する
        GetPath = CStr(regkey.GetValue("InstDir"))

    End Function


End Class
