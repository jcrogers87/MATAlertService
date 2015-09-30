Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Text
Public Class MATAlert_Service
    'ReadConfig Globals
    Dim intervalTime As Single, expMin As Single, emailList() As String, emailFileName As String, checkreaders As Boolean, assetAlertHeader As String, readerAlertHeader As String
    'smtp settings
    Dim smtpUser As String, smtpPass As String, smtpPort As String, smtpSSL As Boolean, smtpServer As String
    'sql settings
    Dim query As String, sqlHost As String, sqlDBName As String, sqlUser As String, sqlPass As String, connectString As String
    Dim dbException As String, remainingTime As Integer
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        Me.MatLog1 = New System.Diagnostics.EventLog
        If Not System.Diagnostics.EventLog.SourceExists("MATAlert") Then
            System.Diagnostics.EventLog.CreateEventSource("MATAlert", "MATAlertServiceLog")
        End If
        MatLog1.Source = "MATAlert"
        MatLog1.Log = "MATAlertServiceLog"
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        MatLog1.WriteEntry("Starting service")
        ReadConfig()
        DoAlertsQuery(1)
        If DBTest() = True Then
            Dim timer As System.Timers.Timer = New System.Timers.Timer()
            timer.Interval = 60000 * intervalTime
            AddHandler timer.Elapsed, AddressOf Me.OnTimer
            timer.Start()
        Else
            MatLog1.WriteEntry("Could not connect to DB")
        End If
    End Sub
    Protected Overrides Sub OnStop()
        MatLog1.WriteEntry("Stopping")
    End Sub
    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)
        ' TODO: Insert monitoring activities here.
        Dim eventID As Integer = 0
        eventID = eventID + 1
        Try
            MatLog1.WriteEntry("Going Active. Starting DB poll every " & intervalTime & " minutes")
            DoAlertsQuery(eventID)
        Catch ex As Exception
            MatLog1.WriteEntry("An exception has occurred:  " & ex.ToString)
        End Try
    End Sub
    Public Function DBTest() As Boolean
        Try
            Dim objConn As SqlConnection = New SqlConnection(connectString)
            objConn.Open()
            objConn.Close()
            MatLog1.WriteEntry("Connected to " & sqlHost)
            Return True
        Catch ex As Exception
            dbException = ex.ToString
            MatLog1.WriteEntry("Cannot connect to SQL DB at: " & sqlHost & Chr(13) & dbException)
            Return False
        End Try
    End Function
    Public Sub ReadConfig()
        If Not File.Exists("C:\ProgramData\Plataine\MatAlert.config") Then
            BuildConfig()
        End If
        'default values 
        checkreaders = False
        assetAlertHeader = "<!DOCTYPE html><html><head><h2>MAT Expiration Alert</h2><style>table, th, td {border: 1px solid black;border-collapse: collapse;}th, td {padding: 3px;</style></head><body><table style=""width:100%"">"
        readerAlertHeader = "<!DOCTYPE html><html><head><h2>Offline Readers</h2><style>table, th, td {border: 1px solid black;border-collapse: collapse;}th, td {padding: 3px;</style></head><body><table style=""width:100%"">"
        Try
            Dim configFile() As String = File.ReadAllLines("C:\ProgramData\Plataine\MatAlert.config")
            For Each line As String In configFile
                Dim setting() As String = Split(line, "=")
                If UCase(setting(0)) = "EMAILLIST" Then
                    If File.Exists(setting(1)) Then
                        If Path.GetExtension(setting(1)).ToString = ".csv" Or Path.GetExtension(setting(1)).ToString = ".txt" Then
                            emailList = File.ReadAllLines(setting(1).ToString)
                            emailFileName = setting(1).ToString
                        End If
                    Else
                        BuildEmails()
                        emailList = File.ReadAllLines("C:\ProgramData\Plataine\recipients.csv")
                        emailFileName = "C:\ProgramData\Plataine\recipients.csv"
                    End If
                ElseIf UCase(setting(0)) = "SMTP.USER" Then
                    smtpUser = setting(1).ToString
                ElseIf UCase(setting(0)) = "SMTP.PASS" Then
                    smtpPass = setting(1).ToString
                ElseIf UCase(setting(0)) = "SMTP.PORT" Then
                    smtpPort = setting(1).ToString
                ElseIf UCase(setting(0)) = "SMTP.SSL" Then
                    If UCase(setting(1).ToString) = "TRUE" Then smtpSSL = True Else smtpSSL = False
                ElseIf UCase(setting(0)) = "SMTP.SERVER" Then
                    smtpServer = setting(1).ToString
                ElseIf UCase(setting(0)) = "DEFAULTINTERVAL" Then
                    intervalTime = setting(1)
                ElseIf UCase(setting(0)) = "DEFUALTEXPMINIMUM" Then
                    expMin = setting(1)
                ElseIf UCase(setting(0)) = "QUERY" Then
                    query = setting(1).ToString
                    If setting.Length > 2 Then
                        For i = 2 To setting.Length - 1
                            query = query & "=" & setting(i).ToString
                        Next
                    End If
                ElseIf UCase(setting(0)) = "SQL.USER" Then
                    sqlUser = setting(1).ToString
                ElseIf UCase(setting(0)) = "SQL.PASS" Then
                    sqlPass = setting(1).ToString
                ElseIf UCase(setting(0)) = "SQL.HOST" Then
                    sqlHost = setting(1).ToString
                ElseIf UCase(setting(0)) = "SQL.DBNAME" Then
                    sqlDBName = setting(1).ToString
                ElseIf UCase(setting(0)) = "CHECKREADERS" Then
                    checkreaders = setting(1).ToString
                ElseIf UCase(setting(0)) = "ASSETALERTHEADER" Then
                    assetAlertHeader = setting(1).ToString
                ElseIf UCase(setting(0)) = "READERALERTHEADER" Then
                    readerAlertHeader = setting(1).ToString
                End If
            Next
            If IsNothing(emailList) Or IsNothing(intervalTime) Or IsNothing(expMin) Or _
                IsNothing(smtpUser) Or IsNothing(smtpPass) Or IsNothing(smtpPort) Or IsNothing(smtpServer) Then
                Dim buildNew As MsgBoxResult = MsgBox("There were missing fields in the config file. Would you like to build a new one?" & Chr(13) _
                                                      & "(This will clear existing)", vbYesNo, "Missing Config")
                If buildNew = vbYes Then
                    BuildConfig()
                    ReadConfig()
                Else
                    End
                End If
            End If
        Catch ex As Exception
            'Call MsgBox("Your config file is missing." _
            '& Chr(13) & "Config location must be: C:\ProgramData\Plataine\MatAlert.config")
            MatLog1.WriteEntry("Could not locate config file: building one")
            BuildConfig()
            ReadConfig()
        End Try
        If query = "" Then
            query = "select ta.[Key], ta.Name, round(ta.MaxExposureTimeMinutes - (ta.ExposureTimeMinutes + DATEDIFF(mi,te.CheckInDate,getdate())*s.ExposureFactor),1) as RemainingTime, ta.currentstationkey as Station, ta.Discriminator as Type, ta.Material, ta.MaterialDescription from Stations as s, TrackingEvents as te inner join TrackedAssets as ta on ta.[key]=te.AssetKey where ta.CurrentStationKey <> 'Freezer' and te.CheckOutDate is null and ta.ExpirationDate > GETDATE() and s.[key]=ta.CurrentStationKey and ta.archived=0 order by te.AssetKey"
        End If
        connectString = "Data Source=" & sqlHost & "\SQLEXPRESS;Initial Catalog=" & sqlDBName & ";User ID=" & sqlUser & ";Password=" & sqlPass & ";"
        MatLog1.WriteEntry("Read config data" & Chr(13) _
                           & "Interval Time = " & intervalTime & " minutes" & Chr(13) _
                           & "Minimum Time = " & expMin & " minutes" & Chr(13) _
                           & "Database = " & sqlHost)
    End Sub
    Public Sub BuildConfig()
        If (Not Directory.Exists("C:\ProgramData\Plataine\")) Then
            Directory.CreateDirectory("C:\ProgramData\Plataine\")
        End If
        Dim sw As New StreamWriter("C:\ProgramData\Plataine\MatAlert.config", False)
        sw.WriteLine("smtp.user=platainetestingusps@gmail.com" & Chr(13) _
        & "smtp.pass=Plataine123" & Chr(13) _
        & "smtp.port=587" & Chr(13) _
        & "smtp.ssl=True" & Chr(13) _
        & "smtp.server=smtp.gmail.com" & Chr(13) _
        & "sql.user=web" & Chr(13) _
        & "sql.pass=web" & Chr(13) _
        & "sql.host=172.20.0.105" & Chr(13) _
        & "sql.dbname=ManualMAT" & Chr(13) _
        & "defaultInterval=1" & Chr(13) _
        & "defualtExpMinimum=500" & Chr(13) _
        & "emailList=C:\ProgramData\Plataine\recipients.csv" & Chr(13) _
        & "checkreaders=false" & Chr(13) _
        & "assetalertheader=<!DOCTYPE html><html><head><h2>MAT Expiration Alert</h2><style>table, th, td {border: 1px solid black;border-collapse: collapse;}th, td {padding: 3px;</style></head><body><table style=""width:100%"">" & Chr(13) _
        & "readeralertheader=<!DOCTYPE html><html><head><h2>Offline Readers</h2><style>table, th, td {border: 1px solid black;border-collapse: collapse;}th, td {padding: 3px;</style></head><body><table style=""width:100%"">")
        sw.Close()
    End Sub
    Public Sub BuildEmails()
        If (Not Directory.Exists("C:\ProgramData\Plataine\")) Then
            Directory.CreateDirectory("C:\ProgramData\Plataine\")
        End If
        Dim sw As New StreamWriter("C:\ProgramData\Plataine\recipients.csv", False)
        sw.WriteLine("nesternet@plataine.com")
        sw.Close()
    End Sub
    Function IsValidEmailFormat(ByVal s As String) As Boolean
        Try
            Dim a As New System.Net.Mail.MailAddress(s)
        Catch
            Return False
        End Try
        Return True
    End Function
    Public Sub DoAlertsQuery(interval As Integer)
        If emailList.Count = 0 Then
            MatLog1.WriteEntry("There is nobody to alert. Recipient list must be populated.")
            Exit Sub
        End If
        MatLog1.WriteEntry("Reading from database at interval " & interval)
        Dim schemaTable As DataTable, myReader As SqlDataReader, columnNames() As String
        Dim myField As DataRow, myProperty As DataColumn
        Dim j As Integer = 0
        Dim queryOutput As New DataTable
        'Open a connection to the SQL Server
        Using cn As New SqlConnection(connectString)
            cn.Open()
            Using cmd As New SqlCommand(query, cn)
                myReader = cmd.ExecuteReader()
                'Retrieve column schema into a DataTable
                schemaTable = myReader.GetSchemaTable()
                Dim foundRemainingTime As Boolean
                'For each field in the table
                For Each myField In schemaTable.Rows
                    'For each property of the field
                    For Each myProperty In schemaTable.Columns
                        'Display the field name and value
                        If UCase(myProperty.ColumnName.ToString) = "COLUMNNAME" Then
                            If UCase(myField(myProperty).ToString) = "REMAININGTIME" Then
                                remainingTime = j
                                foundRemainingTime = True
                                MatLog1.WriteEntry("Found column remaining Time at index: " & remainingTime)
                            End If
                            If myField(myProperty).ToString <> "RowVersion" Then
                                ReDim Preserve columnNames(j)
                                columnNames(j) = myField(myProperty).ToString
                            End If
                            j = j + 1
                        End If
                    Next
                Next
                If foundRemainingTime = False Then
                    MatLog1.WriteEntry("Could not find RemainingTime field")
                    Exit Sub
                End If
                myReader.Close()
                Using myAdapter As New SqlDataAdapter(cmd)
                    myAdapter.Fill(queryOutput)
                End Using
            End Using
            cn.Close()
        End Using

        'loop through sql query output search for any values below expMin. Add these to a text array called alertItems()
        Dim alertItems() As String, k As Integer, columns As Integer
        k = 0
        For i As Integer = 0 To queryOutput.Rows.Count - 1
            If queryOutput.Rows(i).Item(CInt(remainingTime)) < expMin And queryOutput.Rows(i).Item(CInt(remainingTime)) > 0 Then
                MatLog1.WriteEntry("Found asset expiration time below " & expMin)
                ReDim Preserve alertItems(k)
                Dim outputRow As String = ""
                For columns = 0 To queryOutput.Columns.Count - 1
                    outputRow = outputRow & queryOutput.Rows(i)(columns) & ","
                Next
                outputRow = Microsoft.VisualBasic.Left(outputRow, outputRow.Length - 1)
                alertItems(k) = outputRow
                k = k + 1
            End If
        Next
        If Not IsNothing(alertItems) And Not IsNothing(columnNames) Then
            SendAlerts(alertItems, columnNames)
        Else
            MatLog1.WriteEntry("There were no assets close to expiration found")
        End If
    End Sub
    Public Sub SendAlerts(ByVal alerts() As String, ByVal columnnames() As String)
        Dim columnHeaders As String = "<tr>"
        For i = 0 To columnnames.Count - 1
            columnHeaders = columnHeaders & "<th>" & columnnames(i) & "</th>"
        Next
        columnHeaders = columnHeaders & "</tr>"
        Dim strFooter As String = "</table></body></html>"
        Dim sbContent As New StringBuilder()
        For index = 0 To alerts.Length - 1
            Dim messageSplit() As String = Split(alerts(index), ",")
            sbContent.Append("<tr>")
            For j = 0 To messageSplit.Count - 1
                If j = remainingTime Then
                    sbContent.Append(String.Format("<td>{0}</td>", Math.Round(messageSplit(j) / 60, 1)))
                Else
                    sbContent.Append(String.Format("<td>{0}</td>", messageSplit(j)))
                End If
            Next j
            sbContent.Append("</tr>")
        Next

        Dim emailTemplate As String = assetAlertHeader & columnHeaders & sbContent.ToString() & strFooter
        Dim addresses As String = ""
        For i = 0 To emailList.Count - 1
            addresses = addresses & emailList(i) & ","
        Next
        If Microsoft.VisualBasic.Right(addresses, 1) = "," Then addresses = Microsoft.VisualBasic.Left(addresses, addresses.Length - 1)

        Send("Plataine MAT Alert - Assets have exposure times below " & Math.Round(expMin / 60, 1) & " hours.", emailTemplate, addresses)
        MatLog1.WriteEntry("Sent mail to " & addresses & " with " & alerts.Length & " items.")

    End Sub
    Public Sub Send(ByVal subject As String, ByVal body As String, addresses As String)
        Dim Mail As New MailMessage
        Mail.From = New MailAddress("MATAlert@MAT.server")
        Mail.To.Add(addresses)
        Mail.IsBodyHtml = True
        Mail.Subject = subject
        Mail.Body = body
        Try
            Dim smtp As New SmtpClient(smtpServer)
            smtp.EnableSsl = smtpSSL
            smtp.Credentials = New Net.NetworkCredential(smtpUser, smtpPass)
            smtp.Port = smtpPort
            smtp.SendMailAsync(Mail)
            MatLog1.WriteEntry("Sent mail to: " & addresses)
        Catch ex As Exception
            MatLog1.WriteEntry("Could not send mail" & Chr(13) & ex.ToString)
        End Try
    End Sub
End Class
