Imports wstr = Microsoft.VisualBasic.Strings
'Imports config.ｆErr
'Imports config.pfunc

Public Class frmMain
    Const MDBCONNECT As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Const MSGTITLE_CF = "ＯＣＲ変換プログラム"
    Const DIR_HENKAN As String = "henkan\"
    Const CONFIGFILE As String = "Kanjo2kconfig.mdb"         '設定データベース
    Dim pblDsnPath As String
    Dim pblDsnFlg As String
    Dim pblDsnPassWord As String
    Dim pblBfDsnPath As String
    Dim pblBfDsnFlg As String
    Dim pblBfDsnPassWord As String
    Dim pblInstPath As String

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        Dim wrkDsnPath As String

        FDialog.ShowDialog()
        wrkDsnPath = FDialog.FileName

        If wrkDsnPath <> "" Then
            If ChkDsn(wrkDsnPath) = True Then
                MsgBox("接続に成功しました。", , "接続テスト")
                pblDsnPath = wrkDsnPath
                txtDsn.Text = GetDsn(wrkDsnPath)
            Else
                MsgBox("データソースが不正です。ファイルを再度選択してください。", , "接続テスト")
            End If
        End If

    End Sub

    Private Function ChkDsn(ByVal DsnPath As String) As Boolean
        Dim Ret As Integer
        Dim RetValue As Boolean
        Dim AdoData As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim Rs As ADODB.Recordset             'Recordsetオブジェクトの生成
        Dim wrkConnectInfo As String

        On Error GoTo ErrProc

        RetValue = False

        ' 勘定奉行用のDSNかチェック
        Dim fileReader = My.Computer.FileSystem.OpenTextFileReader(DsnPath)
        Do While (fileReader.EndOfStream <> True)
            Dim stringReader = fileReader.ReadLine()
            Ret = InStr(1, stringReader, "info", CompareMethod.Binary)
            If Ret <> 0 Then
                RetValue = True
                Exit Do
            End If
        Loop
        fileReader.Close()

        ' 接続のチェック
        Dim cf As New config.Cfunc
        wrkConnectInfo = cf.fncGetConnect(DsnPath)

        AdoData.Open(wrkConnectInfo)
        ' 会社情報テーブルをオープン
        Rs = New ADODB.Recordset
        Rs.Open("SELECT * FROM wcompany", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)
        RetValue = True
        '接続を切断
        Rs.Close()
        AdoData.Close()
        AdoData = Nothing

        On Error GoTo 0

        ChkDsn = RetValue

        Exit Function

        'エラー時
ErrProc:
        ChkDsn = False

    End Function


    Function GetDsn(ByVal DsnPath As String) As String
        Dim Cnt As Integer          '文字列展開用カウンタ
        Dim hcnt As Integer         '文字列展開用カウンタ
        Dim PathLen As Integer      'ドライブ名取得用バッファ

        'パス取得
        Cnt = 0
        Do While (1)
            Cnt = InStr(Cnt + 1, DsnPath, "\", CompareMethod.Binary)
            If (Cnt < 1) Then
                Exit Do
            End If
            hcnt = Cnt
        Loop

        PathLen = Len(DsnPath) - hcnt
        GetDsn = wstr.Right(DsnPath, PathLen)

    End Function

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sSql As String
        Dim WorkDir As String
        Dim wrkConnectInfo As String
        Dim db As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim drs As New ADODB.Recordset

        On Error GoTo ErrPrc

        Dim cf As New config.Cfunc
        pblInstPath = cf.GetPath()

        ' データベースを開く
        db.Open(MDBCONNECT & pblInstPath & DIR_HENKAN & CONFIGFILE & ";")

        sSql = "SELECT DsnPath,DsnFlg,sub1,DsnPassWord FROM Config"
        drs = Db.Execute(sSql)

        If Information.IsDBNull(drs.Fields("DsnPath").Value) Then
            'If Information.IsDBNull(oRed.Item("DsnPAth")) Then
            pblDsnPath = ""
        Else
            pblDsnPath = drs.Fields("DsnPath").Value
        End If

        pblDsnFlg = wstr.Trim(drs.Fields("DsnFlg").Value)

        If Information.IsDBNull(drs.Fields("DsnPassWord").Value) Then
            pblDsnPassWord = ""
        Else
            pblDsnPassWord = wstr.Trim(drs.Fields("DsnPassWord").Value)
        End If
        pblBfDsnPassWord = pblDsnPassWord

        pblBfDsnPath = pblDsnPath
        pblBfDsnFlg = pblDsnFlg

        drs.Close()
        drs = Nothing
        db.Close()
        db = Nothing

        txtPassWord.Text = pblBfDsnPassWord

        txtDsn.Text = GetDsn(pblDsnPath)
        txtDsn.Visible = True
        btnSelect.Enabled = True

        Exit Sub

        On Error GoTo 0

        ' エラー処理
ErrPrc:

        Call ErrMessage("設定データ取得中")

    End Sub


    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Dim Ssql As String
        Dim db As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        'Dim drs As New ADODB.Recordset

        pblDsnPassWord = Trim(txtPassWord.Text)

        ' データソースの選択があるかチェック
        If (pblDsnPath = "") Then
            MsgBox("データソースが選択されていません。", , "注意")
            Exit Sub
        End If
        If (pblDsnPath <> "") And (ChkDsn(pblDsnPath) = False) Then
            MsgBox("データソースが不正です。ファイルを再度選択してください。", , "接続テスト")
            Exit Sub
        End If

        ' 変更がない場合はそのまま閉じる
        If pblDsnPath = pblBfDsnPath And pblDsnFlg = pblBfDsnFlg And _
           pblBfDsnPassWord = pblDsnPassWord Then

            End
        Else
            ' 変更があった場合
            ' データベースへの書込み
            ' データベースを開く
            db.Open(MDBCONNECT & pblInstPath & DIR_HENKAN & CONFIGFILE & ";")

            Ssql = "UPDATE Config SET DsnPath = '" & pblDsnPath & _
                                       "', DsnFlg = '" & pblDsnFlg
            If pblDsnPassWord = "" Then
                Ssql = Ssql & "', DsnPassWord = NULL"
            Else
                Ssql = Ssql & "', DsnPassWord = '" & pblDsnPassWord & "'"
            End If
            db.Execute(Ssql)
            'drs.Close()
            'drs = Nothing
            db.Close()
            db = Nothing

            MsgBox("接続設定の内容を変更しました。", , "変更終了")
            End
        End If
    End Sub

    Private Sub ErrMessage(ByVal Msg As String)

        MsgBox(Msg & "にエラーが発生したため、処理を終了します。", , "エラー")
        End

    End Sub
End Class

