Imports System.Security.Principal.WindowsIdentity
Imports System.NullReferenceException
Imports Tulpep.NotificationWindow
Imports System.Globalization
Imports System.Configuration
Imports System.Data.OleDb
Imports System.Threading
Imports System.Security
Imports System.Data
Imports System.Windows.Forms


Module Rotinas
    Public userLogin As String
    Public idUser As Byte
    Public provider As String
    Public arquivoDados As String
    Public connString As String
    Public conexaoAccess As OleDbConnection = New OleDbConnection
    Public DR As OleDbDataReader
    Public SQL As String
    Public TEXTO As String
    Public SW As Boolean
    Public Intervalo As Integer
    Public whoscall As String

    Public Sub abreConexao()
        If conexaoAccess.State = ConnectionState.Open Then conexaoAccess.Close()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        arquivoDados = "\\grupo.qualicorp\qualidocs\ORGANIZACIONAL\GER CADASTRO\GERENCIAMENTO DE DADOS\" &
        "REPOSITÓRIO - FLÁVIO\BASES_APLICAÇÕES\INTERFACE\GERENCIADOR DE REATIVAÇÃO\BASE DE DADOS\BD_REATIVACAO.accdb"
        connString = provider & arquivoDados
        conexaoAccess.ConnectionString = connString
        'End If
    End Sub
    Public Sub popUp()

        Dim notificacao = New PopupNotifier()
        notificacao.TitleText = "       *** NOTIFICAÇÃO DE PENDÊNCIA ***" & vbNewLine
        notificacao.TitleColor = Color.Black
        notificacao.ContentText = vbNewLine & TEXTO
        notificacao.AnimationInterval = 30000
        notificacao.IsRightToLeft = False
        notificacao.ShowCloseButton = True
        notificacao.ContentColor = Color.Black
        notificacao.ShowGrip = True
        notificacao.Image = Image.FromFile("\\grupo.qualicorp\qualidocs\ORGANIZACIONAL\GER CADASTRO\INTERFACE\" & _
                                           "CALENDARIO_INTERFACE\IMG\relogio alerta.jpg", True)
        notificacao.Popup()
    End Sub
    Public Function userLOGADO() As String

        userLogin = GetCurrent.Name.ToString
        Dim texto As String = userLogin
        Dim palavras As String() = texto.Split(New Char() {"\"c})
        userLogin = palavras(1)
        Return userLogin

    End Function
    Public Sub pegaIdUser()

        SQL = ""
        SQL = SQL & "SELECT COD_ANALISTA FROM TB_ANALISTAS WHERE LOGIN=" & "'" & userLogin & "'"

        whoscall = "pegaIdUser"
        Try
            abreConexao()
            conexaoAccess.Open()
            Dim CMD As OleDbCommand = New OleDbCommand(SQL, conexaoAccess)
            DR = CMD.ExecuteReader
            DR.Read()
            idUser = DR(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        conexaoAccess.Close()

    End Sub
    Public Sub procuraTarefas()
        Dim dt As String
        Dim partNome As String
        Dim comprimento As Byte

        whoscall = "procuraTarefas"

        dt = CDate(Now).ToString("MM/dd/yyyy")

        SQL = ""
        SQL = SQL & "SELECT * FROM CS_TAREFAS "
        SQL = SQL & "WHERE INPUTDATE=#" & dt & "# AND USERRESP='" & userLogin & "' AND STATUS=0"

        abreConexao()
        conexaoAccess.Open()
        Dim CMD As OleDbCommand = New OleDbCommand(SQL, conexaoAccess)
        DR = CMD.ExecuteReader
        TEXTO = ""
        SW = False
        If DR.HasRows Then
            SW = True
            While DR.Read()
                partNome = DR("INPUTTEXT")
                comprimento = Len(partNome)
                partNome = Trim(Mid(partNome, 1, comprimento))

                TEXTO = TEXTO & partNome & " | "
            End While
        End If
    End Sub
    Public Sub consultaIntervalo()

        whoscall = "consultaIntervalo"

        SQL = ""
        SQL = SQL & "SELECT INTERVALO_ALERTA FROM TB_INTERVALO_ALERTA WHERE LIGADO=TRUE"

        abreConexao()
        conexaoAccess.Open()

        Dim CMD As OleDbCommand = New OleDbCommand(SQL, conexaoAccess)
        DR = CMD.ExecuteReader
        If DR.HasRows Then
            DR.Read()
            Intervalo = DR(0) * 1000 * 60
            conexaoAccess.Close()
        End If
    End Sub
End Module
