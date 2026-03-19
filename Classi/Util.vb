Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail

Public Module Util
    Private bError As Boolean = False
    Public sError As String = ""

    Public Property IsAppInError As Boolean
        Get
            IsAppInError = bError
        End Get

        Set(value As Boolean)
            bError = value
        End Set
    End Property

    Public Sub ThrowException(ByVal Message As String, Optional ByVal isAlreadyException As Boolean = False)
        Dim Log As New FileLog(GetValueFromConfig("LogPath"))

        IsAppInError = True

        Log.NewLine()
      Log.WriteLine(Now.ToString & " - " & Message)

      Log = Nothing

        If CInt(GetValueFromConfig("enable_err_mail")) = 1 Then
            sError &= Message
        End If

        If Not isAlreadyException Then
            Throw New Exception(Message)
        End If
    End Sub

    Public Function SendEmail(ByVal toaddress As String, ByVal cc As String, ByVal ccn As String, ByVal subject As String, ByVal message As String, ByVal isHtml As Boolean, Optional ByVal Priority As MailPriority = MailPriority.Normal) As Boolean
        Dim lReturn As Boolean = False

        Try
            'Se non c'è il destinatario, esce
            If IsEmpty(toaddress.ToString) Then
                Exit Try
            End If

            Dim mailFrom As String
            Dim mailFrom_displayName As String

            mailFrom = GetValueFromConfig("mail_From_address").ToString
            mailFrom_displayName = GetValueFromConfig("mail_From_displayName").ToString

            Dim oMail As New MailMessage
            oMail.From = New MailAddress(mailFrom, mailFrom_displayName)
            oMail.To.Add(Replace(toaddress.ToString, ";", ","))

            If (String.IsNullOrEmpty(cc) = False) AndAlso (cc.Trim(" "c) <> "") Then
                oMail.CC.Add(Replace(cc.ToString, ";", ","))
            End If

            If (String.IsNullOrEmpty(ccn) = False) AndAlso (ccn.Trim(" "c) <> "") Then
                oMail.Bcc.Add(Replace(ccn.ToString, ";", ","))
            End If

            oMail.Priority = Priority
            oMail.Subject = subject
            oMail.Body = message

            oMail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess
            oMail.IsBodyHtml = isHtml

            Dim oSmtp As SmtpClient = New SmtpClient(GetValueFromConfig("mail_host").ToString, CInt(GetValueFromConfig("mail_port")))

            If (GetValueFromConfig("defaultCredentials").ToString = "1") Then
                oSmtp.UseDefaultCredentials = True
                oSmtp.Credentials = CredentialCache.DefaultNetworkCredentials
            Else
                oSmtp.UseDefaultCredentials = False
                If Len(GetValueFromConfig("mail_userName").ToString) > 0 Then
                    If (Len(GetValueFromConfig("mail_domain").ToString) > 0) Then
                        oSmtp.Credentials = New NetworkCredential(GetValueFromConfig("mail_userName").ToString, GetValueFromConfig("mail_password").ToString, GetValueFromConfig("mail_domain").ToString)
                    Else
                        oSmtp.Credentials = New NetworkCredential(GetValueFromConfig("mail_userName").ToString, GetValueFromConfig("mail_password").ToString)
                    End If
                End If
            End If

            If (GetValueFromConfig("enableSSL").ToString = "1") Then
                oSmtp.EnableSsl = True
            End If

            oSmtp.DeliveryMethod = SmtpDeliveryMethod.Network
            oSmtp.Send(oMail)
            oSmtp = Nothing

            oMail.Dispose()

            lReturn = True
        Catch ex As SmtpException


        Catch ex As Exception

        End Try

        Return lReturn
    End Function

    Public Function FromHashTableToQueryString(ByVal HT As Hashtable) As String
        Dim sReturn As String = ""

        For Each element As DictionaryEntry In HT
            sReturn &= element.Key & "=" & element.Value & "&"
        Next

        Return Left(sReturn, Len(sReturn) - 1)
    End Function

    Public Function GetValueFromConfig(ByVal Campo As String) As String
        Dim _ret As String = ""
        _ret = ConfigurationManager.AppSettings.Get(Campo.ToUpper)
        If Not IsNothing(_ret) Then
            Return _ret
        Else
            Return ""
        End If
    End Function

    Public Function getConnectionString() As String
        Return ConfigurationManager.ConnectionStrings("DEFAULT").ConnectionString
    End Function

    Public Function TerminatePath(ByVal path As String) As String
        Return IIf(Right(path, 1) = "\", path, path & "\")
    End Function

    Public Sub addToJSON(ByRef sJson As String, ByVal sChiaveDaAggiungere As String, ByVal sValoreDaAggiungere As String, ByVal sTipoValore As String, ByVal lVirgola As Boolean)
        Const Doppie As String = """"

        Select Case sTipoValore
            Case "S" 'String
                sJson &= Doppie & sChiaveDaAggiungere & Doppie & ": " & Doppie & escapeValueJSON(sValoreDaAggiungere) & Doppie & IIf(lVirgola, ",", "")

            Case "I" 'Integer
                sJson &= Doppie & sChiaveDaAggiungere & Doppie & ": " & Replace(sValoreDaAggiungere, ",", ".") & IIf(lVirgola, ",", "")

            Case "B" 'Boolean
                sJson &= Doppie & sChiaveDaAggiungere & Doppie & ":" & LCase(sValoreDaAggiungere) & IIf(lVirgola, ",", "")

            Case "D" 'Date
                sJson &= Doppie & sChiaveDaAggiungere & Doppie & ": " & Doppie & sValoreDaAggiungere & IIf(Not IsEmpty(sValoreDaAggiungere), " 00:00:00", "") & Doppie & IIf(lVirgola, ",", "")

        End Select
    End Sub

    Private Function escapeValueJSON(ByVal Value As String) As String
        'Backslash is replaced with \\
        Value = Value.Replace(Chr(92), "\\")
        'Backspace is replaced with \b
        Value = Value.Replace(Chr(8), "\b")
        'Form feed is replaced with \f
        Value = Value.Replace(Chr(12), "\f")
        'Newline (Line Feed) is replaced with \n
        Value = Value.Replace(Chr(10), "\n")
        'Carriage return is replaced with \r
        Value = Value.Replace(Chr(13), "\r")
        'Tab is replaced with \t
        Value = Value.Replace(Chr(9), "\t")
        'Double quote is replaced with \"
        Value = Value.Replace(Chr(34), "\""")

        Return Value
    End Function

    Public Function RboApice(ByVal value As String) As String
        Dim _ret As String
        _ret = value.Replace("'", "''")
        Return _ret
    End Function

    Public Function RboData(ByVal value As String) As String
        Dim _ret As String
        Dim str As String
        Dim strArr() As String
        str = value.ToString
        strArr = str.Split(CChar("/"))
        _ret = Mid(strArr(2), 1, 4) & "-" & strArr(1) & "-" & strArr(0)
        Return _ret 'formato yyyy-mm-dd
    End Function

    Public Class FileLog
        Private m_FileName As String
        Private m_DateTimeInfoMessage As Boolean

        Public Sub New()
            m_FileName = ""
        End Sub

        Public Sub New(ByVal filename As String)
            Me.New()
            m_FileName = filename
        End Sub

        Public Sub New(ByVal filename As String, ByVal dateTimeInfo As Boolean)
            Me.New(filename)
            m_DateTimeInfoMessage = dateTimeInfo
        End Sub

        Public Sub NewLine()
            Me.Write(System.Environment.NewLine)
        End Sub
        Public Sub NewLine(ByVal row As Int32)
            For ii As Int32 = 1 To row
                Me.NewLine()
            Next
        End Sub

        Public Sub WriteLine(ByVal message As String)
            If m_DateTimeInfoMessage Then
                message = GetDateTimeMessage(message)  'System.DateTime.Now.ToString() + " " + message
            End If
            message += System.Environment.NewLine
            Me.Write(message)
        End Sub

        Public Sub Write(ByVal message As String, ByVal dateTimeInfoMessage As Boolean)
            If dateTimeInfoMessage Then message = GetDateTimeMessage(message)
            Me.Write(message)
        End Sub

        Private Function GetDateTimeMessage(ByVal message As String) As String

            message = System.DateTime.Now.ToString() + " " + message

            Return (message)
        End Function

        Public Sub Write(ByVal message As String)
            If m_FileName.Length > 0 Then

                Dim sw As System.IO.StreamWriter = System.IO.File.AppendText(m_FileName)

                Try
                    sw.Write(message)
                    sw.Flush()
                Catch ex As Exception
                    Throw ex
                Finally
                    sw.Close()
                End Try

            End If
        End Sub
    End Class

    Public Class SQL
        Implements IDisposable

        Private _cnn As SqlConnection = Nothing
        Private _cmd As SqlCommand = Nothing
        Private _connectionString As String = ""
        Private _SQLCommandTimeout As Integer = 0
        Private _err As New CProprieta

        Sub New(ByVal ConnectionString As String, Optional ByVal SQLCommandTimeout As Integer = 0) '0 = no Timeout
            _cnn = New SqlConnection
            _cmd = _cnn.CreateCommand
            _connectionString = ConnectionString
            _SQLCommandTimeout = SQLCommandTimeout
            _err.Clear()
            If Not OpenCnn() Then
                _err.Scrivi("ERRORE", "Non riesce ad aprire la connessione!")
            End If
        End Sub

        Private Function OpenCnn() As Boolean
            Dim _ret As Boolean = False
            Try
                _cnn.ConnectionString = _connectionString
                _cnn.Open()
                _ret = True
            Catch ex As Exception
                _ret = False
            End Try
            Return _ret
        End Function

        Private Function CloseCnn() As Boolean
            Dim _ret As Boolean = False
            Try
                _cnn.Close()
                _ret = True
            Catch ex As Exception
                _ret = False
            End Try
            Return _ret
        End Function

        Private Sub CreaParametri(ByVal nome As String, ByVal valore As Object)
            Dim P As SqlParameter = _cmd.CreateParameter
            P.ParameterName = nome
            P.SqlDbType = TipoDiDato(valore)
            P.Value = valore
            _cmd.Parameters.Add(P)
        End Sub

        Private Function TipoDiDato(ByVal valore As Object) As SqlDbType
            Dim _ret As SqlDbType
            Select Case VarType(valore)
                Case VariantType.String : _ret = SqlDbType.VarChar
                Case VariantType.Byte : _ret = SqlDbType.TinyInt
                Case VariantType.Short : _ret = SqlDbType.SmallInt
                Case VariantType.Integer : _ret = SqlDbType.Int
                Case VariantType.Long : _ret = SqlDbType.BigInt
                Case VariantType.Single : _ret = SqlDbType.Real
                Case VariantType.Double : _ret = SqlDbType.Float
                Case VariantType.Decimal : _ret = SqlDbType.Decimal
                Case VariantType.Date : _ret = SqlDbType.DateTime
                Case VariantType.Boolean : _ret = SqlDbType.Bit
                Case VariantType.Empty : _ret = SqlDbType.VarChar
                Case VariantType.Null : _ret = Nothing
            End Select
            Return _ret
        End Function

        Public Function EseguiSQL(ByVal sSQL As String, Optional ByRef RowsAffected As Integer = 0, Optional ByVal pPram As CProprieta = Nothing) As Boolean
            Dim _ret As Boolean = False

            If Not CiSonoErrori() Then
                Try
                    'Definiamo un comando
                    RefreshCMD(_cmd)

                    'Definiamo i parametri
                    If Not IsNothing(pPram) Then
                        Dim entry As DictionaryEntry
                        Dim ht As New Hashtable
                        pPram.GetHashTableFromCProprieta(pPram, ht)
                        For Each entry In ht
                            Dim chiave As String = CStr(entry.Key)
                            Dim valore As String = CType(entry.Value, String)
                            CreaParametri("@" & chiave, valore)
                        Next
                    End If

                    _cmd.CommandText = sSQL
                    'Se RowsAffected > 0    ==>    query andata a buon fine 
                    RowsAffected = _cmd.ExecuteNonQuery()
                    _cmd.Dispose()
                    _ret = True
                Catch ex As Exception
                    _err.Scrivi("ERRORE", "Source: " & ex.Source & " Message: " & ex.Message & " StackTrace: " & ex.StackTrace)
                    _ret = False
                End Try
            Else
                _ret = False
            End If

            Return _ret
        End Function

        ''' <summary>
        ''' Corretto utilizzo: If hSQL.OttieniValori OK allora --> If DataTable.Rows.Count>0 OK allora --> dati estratti correttamente
        ''' </summary>
        ''' <param name="sSQL"></param>
        ''' <param name="DataTable"></param>
        ''' <param name="pPram"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function OttieniValori(ByVal sSQL As String, ByRef DataTable As DataTable, Optional ByVal pPram As CProprieta = Nothing) As Boolean
            Dim _ret As Boolean = False

            If Not CiSonoErrori() Then
                Try
                    'Definiamo un comando
                    RefreshCMD(_cmd)

                    'Definiamo i parametri
                    If Not IsNothing(pPram) Then
                        Dim entry As DictionaryEntry
                        Dim ht As New Hashtable
                        pPram.GetHashTableFromCProprieta(pPram, ht)
                        For Each entry In ht
                            Dim chiave As String = CStr(entry.Key)
                            Dim valore As String = CType(entry.Value, String)
                            CreaParametri("@" & chiave, valore)
                        Next
                    End If

                    _cmd.CommandText = sSQL
                    'Adattatore dati
                    Dim da As SqlDataAdapter = New SqlDataAdapter
                    da.SelectCommand = _cmd
                    'Oggetto tabella dati
                    Dim dt As New DataTable
                    da.Fill(dt)
                    dt.Dispose()
                    'NB. Se la query da errore "Divide by zero error encountered."
                    'la procedura, non si ferma, si limita a non mettere la riga in errore

                    If dt.Rows.Count > 0 Then
                        DataTable = dt.Clone
                        For Each dr As DataRow In dt.Rows
                            DataTable.ImportRow(dr)
                        Next
                    End If

                    _cmd.Dispose()
                    _ret = True
                Catch ex As Exception
                    _err.Scrivi("ERRORE", "Source: " & ex.Source & " Message: " & ex.Message & " StackTrace: " & ex.StackTrace)
                    _ret = False
                End Try
            Else
                _ret = False
            End If

            Return _ret
        End Function

        Public Function LeggiValore(ByVal Tabella As String, ByVal Attributo As String, ByVal Condizione As String, ByRef Value As String, Optional ByVal pPram As CProprieta = Nothing) As Boolean
            Dim _ret As Boolean = False

            If Not CiSonoErrori() Then
                Try
                    'Definiamo un comando
                    RefreshCMD(_cmd)

                    'Definiamo i parametri
                    If Not IsNothing(pPram) Then
                        Dim entry As DictionaryEntry
                        Dim ht As New Hashtable
                        pPram.GetHashTableFromCProprieta(pPram, ht)
                        For Each entry In ht
                            Dim chiave As String = CStr(entry.Key)
                            Dim valore As String = CType(entry.Value, String)
                            CreaParametri("@" & chiave, valore)
                        Next
                    End If

                    _cmd.CommandText = "SELECT TOP 1 " & Attributo & " AS AT FROM " & Tabella & " WHERE " & Condizione

                    Value = ""

                    Dim _reader As SqlDataReader
                    _reader = _cmd.ExecuteReader
                    If _reader.HasRows Then
                        Do While _reader.Read()
                            '-
                            'OLD
                            'Value = _reader.GetString(0) 'AT
                            '-
                            'Gianluca 30/09/2015
                            If IsNull(_reader.Item(0)) Then
                                Value = ""
                            Else
                                Value = _reader.Item(0).ToString
                            End If
                            Exit Do
                        Loop
                    End If
                    _reader.Close()

                    _cmd.Dispose()
                    _ret = True
                Catch ex As Exception
                    _err.Scrivi("ERRORE", "Source: " & ex.Source & " Message: " & ex.Message & " StackTrace: " & ex.StackTrace)
                    _ret = False
                End Try
            Else
                _ret = False
            End If

            Return _ret
        End Function

        Public Sub LeggiErrore(ByRef err As CProprieta)
            err = _err
        End Sub

        Public Sub LeggiErrore(ByRef err As String)
            err = _err.Leggi("ERRORE")
        End Sub

        Public Function LeggiErrore() As String
            Return _err.Leggi("ERRORE")
        End Function

        Public Sub PulisciErrore()
            _err.Clear()
        End Sub

        Private Function CiSonoErrori() As Boolean
            If _err.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Sub RefreshCMD(ByRef cmd As SqlCommand)
            cmd = Nothing
            cmd = _cnn.CreateCommand
            cmd.CommandTimeout = _SQLCommandTimeout
            cmd.CommandType = CommandType.Text
        End Sub

        ''' <summary>
        ''' Crea stringa (con parametri in SqlCommand) per operatore SQL IN
        ''' </summary>
        ''' <param name="cmd"></param>
        ''' <param name="value">es. val1#val2#val3</param>
        ''' <param name="prefissoParam">es. T</param>
        ''' <param name="divisore">es. #</param>
        ''' <returns>@T1,@T2,@T3</returns>
        ''' <remarks></remarks>
        Public Function CreaStringaIn(ByRef cmd As SqlCommand, ByVal value As String, ByVal prefissoParam As String, ByVal divisore As String) As String
            Dim Arr() As String = Nothing
            value = value & divisore
            Arr = value.Split(CChar(divisore))
            Dim ciN As String = ""
            For i As Integer = 0 To Arr.Length - 1
                If (Arr(i) <> "") Then
                    If cmd.Parameters.Contains("@" & prefissoParam & i) Then
                        cmd.Parameters.Remove(cmd.Parameters("@" & prefissoParam & i))
                    End If
                    cmd.Parameters.AddWithValue("@" & prefissoParam & i, Arr(i))
                    ciN = ciN & "@" & prefissoParam & i & ","
                End If
            Next
            If Right(ciN, 1) = "," Then
                ciN = Left(ciN, ciN.Length - 1)
            End If
            Return ciN
        End Function

        ''' <summary>
        ''' Crea stringa (con parametri in CProprieta) per operatore SQL IN
        ''' </summary>
        ''' <param name="cmd"></param>
        ''' <param name="value">es. val1#val2#val3</param>
        ''' <param name="prefissoParam">es. T</param>
        ''' <param name="divisore">es. #</param>
        ''' <returns>@T1,@T2,@T3</returns>
        ''' <remarks></remarks>
        Public Function CreaStringaIn(ByRef cmd As CProprieta, ByVal value As String, ByVal prefissoParam As String, ByVal divisore As String) As String
            Dim Arr() As String = Nothing
            value = value & divisore
            Arr = value.Split(CChar(divisore))
            Dim ciN As String = ""
            For i As Integer = 0 To Arr.Length - 1
                If (Arr(i) <> "") Then
                    If cmd.Contains(prefissoParam & i) Then
                        cmd.Scrivi(prefissoParam & i, "")
                    End If
                    cmd.Scrivi(prefissoParam & i, Arr(i))
                    ciN = ciN & "@" & prefissoParam & i & ","
                End If
            Next
            If Right(ciN, 1) = "," Then
                ciN = Left(ciN, ciN.Length - 1)
            End If
            Return ciN
        End Function

        ''' <summary>
        ''' Crea stringa per operatore SQL IN
        ''' </summary>
        ''' <param name="value">es. val1|val2|val3</param>
        ''' <param name="divisore">es. |</param>
        ''' <returns>'val1','val2','val3'</returns>
        ''' <remarks></remarks>
        Public Function CreaStringaIn(ByVal value As String, ByVal divisore As String) As String
            Dim Arr() As String = Nothing
            value = value & divisore
            Arr = value.Split(CChar(divisore))
            Dim ciN As String = ""
            For i As Integer = 0 To Arr.Length - 1
                If (Arr(i) <> "") Then
                    ciN = ciN & "'" & Arr(i) & "',"
                End If
            Next
            If Right(ciN, 1) = "," Then
                ciN = Left(ciN, ciN.Length - 1)
            End If
            Return ciN
        End Function

        ''' <summary>
        ''' Esegue un blocco di query in transaction, se da errore va in rollback
        ''' </summary>
        ''' <param name="myConnString"></param>
        ''' <param name="ArrQuery"></param>
        ''' <param name="err"></param>
        ''' <param name="mySQLCommandTimeout"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RunSQLTransaction(ByVal myConnString As String, ByVal ArrQuery As String(), Optional ByRef err As String = "", Optional ByVal mySQLCommandTimeout As Integer = 0) As Boolean '0 = no Timeout
            Dim _ret As Boolean = False

            Dim myConnection As New SqlConnection(myConnString)
            myConnection.Open()

            Dim myCommand As SqlCommand = myConnection.CreateCommand()
            Dim myTrans As SqlTransaction
            Dim tName As String = ""
            tName = "TSQL" & Date.Now.ToShortDateString.Replace("/", "") & Date.Now.ToShortTimeString.Replace(".", "").Replace(":", "")

            ' Start a local transaction
            myTrans = myConnection.BeginTransaction(IsolationLevel.ReadCommitted, tName)
            ' Must assign both transaction object and connection
            ' to Command object for a pending local transaction
            myCommand.Connection = myConnection
            myCommand.CommandTimeout = mySQLCommandTimeout
            myCommand.Transaction = myTrans
            Try
                For i As Integer = 0 To ArrQuery.Length - 1
                    If Not IsEmpty(ArrQuery(i)) Then
                        If Not IsNull(ArrQuery(i)) Then
                            myCommand.CommandText = ArrQuery(i)
                            myCommand.ExecuteNonQuery()
                        End If
                    End If
                Next
                myTrans.Commit()
                _ret = True
            Catch e As Exception
                err = err & "Source: " & e.Source & vbNewLine
                err = err & "Message: " & e.Message & vbNewLine
                err = err & "StackTrace: " & e.StackTrace & vbNewLine
                err = err & "GetType: " & e.GetType.ToString()
                Try
                    myTrans.Rollback(tName)
                Catch ex As SqlException
                    If Not myTrans.Connection Is Nothing Then
                        err = err & "Source: " & ex.Source & vbNewLine
                        err = err & "Message: " & ex.Message & vbNewLine
                        err = err & "StackTrace: " & ex.StackTrace & vbNewLine
                        err = err & "GetType: " & ex.GetType.ToString()
                    End If
                End Try
            Finally
                myConnection.Close()
            End Try

            Return _ret
        End Function

#Region "IDisposable Support"
        Private disposedValue As Boolean ' Per rilevare chiamate ridondanti

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then

                End If
            End If
            Me.disposedValue = True
        End Sub

        ' Questo codice è aggiunto da Visual Basic per implementare in modo corretto il modello Disposable.
        Public Sub Dispose() Implements IDisposable.Dispose
            CloseCnn()
            PulisciErrore()
            ' Non modificare questo codice. Inserire il codice di pulizia in Dispose(ByVal disposing As Boolean).
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class

#Region " IsNull "
    Public Function IsNull(ByVal value As String, ByVal defaultValue As String) As String
        Dim returnValue As String

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Decimal, ByVal defaultValue As Decimal) As Decimal
        Dim returnValue As Decimal

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Integer, ByVal defaultValue As Integer) As Integer
        Dim returnValue As Integer

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Single, ByVal defaultValue As Single) As Single
        Dim returnValue As Single

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Double, ByVal defaultValue As Double) As Double
        Dim returnValue As Double

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Object, ByVal defaultValue As Object) As Object
        Dim returnValue As Object

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Long, ByVal defaultValue As Long) As Long
        Dim returnValue As Long

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Short, ByVal defaultValue As Short) As Short
        Dim returnValue As Short

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As DateTime, ByVal defaultValue As DateTime) As DateTime
        Dim returnValue As DateTime

        If IsNull(value) Then
            returnValue = defaultValue
        Else
            returnValue = value
        End If

        Return (returnValue)
    End Function

    Public Function IsNull(ByVal value As Object) As Boolean
        Dim lEmpty As Boolean = False

        If value Is Nothing Then
            lEmpty = True
        ElseIf IsDBNull(value) Then
            lEmpty = True
        End If

        Return (lEmpty)
    End Function

    Public Function IsNull(ByVal value As DateTime) As Boolean
        Dim lEmpty As Boolean = False

        If value = Nothing Then
            lEmpty = True
        ElseIf IsDBNull(value) Then
            lEmpty = True
        End If

        Return (lEmpty)
    End Function

#End Region

#Region " IsEmpty "
    Public Function IsEmpty(ByVal value As String) As Boolean
        Return (Trim(IsNull(value, "")).Length = 0)
    End Function
    Public Function IsEmpty(ByVal value As Decimal) As Boolean
        Return (IsNull(value, 0) = 0)
    End Function
    Public Function IsEmpty(ByVal value As Integer) As Boolean
        Return (IsNull(value, 0) = 0)
    End Function
    Public Function IsEmpty(ByVal value As Single) As Boolean
        Return (IsNull(value, 0) = 0)
    End Function
    Public Function IsEmpty(ByVal value As Double) As Boolean
        Return (IsNull(value, 0) = 0)
    End Function
    Public Function IsEmpty(ByVal value As Short) As Boolean
        Return (IsNull(value, 0) = 0)
    End Function
    Public Function IsEmpty(ByVal value As Long) As Boolean
        Return (IsNull(value, 0) = 0)
    End Function
    Public Function IsEmpty(ByVal value As DateTime) As Boolean
        Return (IsNull(value, Nothing) = Nothing)
    End Function
    Public Function IsEmpty(ByVal value As Object) As Boolean
        Return (IsNull(value, Nothing) Is Nothing)
    End Function
    Public Function IsEmpty(ByVal value As String, ByVal defaultValue As String) As String
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Decimal, ByVal defaultValue As Decimal) As Decimal
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Integer, ByVal defaultValue As Integer) As Integer
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Single, ByVal defaultValue As Single) As Single
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Double, ByVal defaultValue As Double) As Double
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Short, ByVal defaultValue As Short) As Short
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Long, ByVal defaultValue As Long) As Long
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As DateTime, ByVal defaultValue As DateTime) As DateTime
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function
    Public Function IsEmpty(ByVal value As Object, ByVal defaultValue As Object) As Object
        If IsEmpty(value) Then value = defaultValue
        Return (value)
    End Function

#End Region

End Module