Imports System.IO
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Net
Imports System.Convert
Imports FirebirdSql.Data.FirebirdClient

Module REACTSMS

    Dim strEmail_server As String = SaveSets.Load("Email", "Server")
    Dim strEmail_address As String = SaveSets.Load("Email", "Address")
    Dim strSystemlogs_path As String = SaveSets.Load("TestOfCure", "FileNameREACT")
    Dim strCPMSServer As String = SaveSets.Load("CPMS", "Server")
    Dim strCPMSPath As String = SaveSets.Load("CPMS", "Location")
    Dim gDBConnectionString As String = SaveSets.Load("SQL", "ConnectionStringMSHC")

    Dim gSMSServer As String = SaveSets.Load("SMS", "Server")
    Dim gSMSUsername As String = SaveSets.Load("TestOfCure", "SMSUsername")
    Dim gSMSPassword As String = SaveSets.Load("TestOfCure", "SMSPassword")
    Dim gSMSType As String = SaveSets.Load("TestOfCure", "SMSType")
    Dim gSMSSenderID As String = SaveSets.Load("TestOfCure", "SMSSenderID")
    Dim gSMSMessage As String = SaveSets.Load("TestOfCure", "MessageREACT")
    Dim gEmailSubjectREACT As String = SaveSets.Load("TestOfCure", "EmailSubjectREACT")
    Dim gDaysREACT As String = SaveSets.Load("TestOfCure", "DaysREACT")

    Sub Main()
        Console.WriteLine("Preparing Chlamydia SMS Reminder (CPMSv3 database)")
        Console.WriteLine("- Querying and processing CPMSv3 database...")
        ProcessDatabase()
        ProcessDatabase_Rec()
        Console.WriteLine("Chlamydia SMS Reminder successfully executed")
    End Sub

    Private Sub ProcessDatabase()
        Dim dsCX As DataSet
        Dim dtReminderPeriod As String
        Dim strSQL As String
        Dim intDaysREACT As Integer = Convert.ToInt32(gDaysREACT) * -1
        Dim strSQL2 As String = ""
        Dim strMessage2 As String = ""
        Dim dsTemp2 As DataSet

        dtReminderPeriod = Format(Date.Today.AddDays(intDaysREACT), "yyyy-MM-dd")   '90 days

        'Get list of clients who have positive chlamydia
        strSQL = "SELECT C.URNO, C.MOBILE, MT.LABNO, C.SMSRESULTCONSENTCODE " &
                 "FROM CLIENT C " &
                 "JOIN MDUSERVICE MS ON (MS.OWNER_OID = C.OID) " &
                 "JOIN MDUTEST MT ON (MT.OWNER_OID = MS.OID) " &
                 "LEFT JOIN MDUTEST_HVS MT_HVS ON (MT_HVS.OID = MT.OID) " &
                 "LEFT JOIN MDUTEST_CX MT_CX ON (MT_CX.OID = MT.OID) " &
                 "LEFT JOIN MDUTEST_UR MT_UR ON (MT_UR.OID = MT.OID) " &
                 "LEFT JOIN MDUTEST_TH MT_TH ON (MT_TH.OID = MT.OID) " &
                 "LEFT JOIN MDUTEST_FPU MT_FPU ON (MT_FPU.OID = MT.OID) " &
                 "WHERE CAST (MS.CREATEDBYDATETIME AS DATE) >= '" & dtReminderPeriod & "' " &
                 "AND CAST (MS.CREATEDBYDATETIME AS DATE) <= '" & dtReminderPeriod & " ' " &
                 "AND (C.URNO NOT IN (0,99999,888885,888886,888887,888888,999997,999998,999999)) " &
                 "AND (MT_TH.CHLAMYDIANAATPANTHERCODE IN (1,11) OR MT_HVS.CHLAMYDIANAATPANTHERCODE IN (1,11) OR MT_CX.CHLAMYDIANAATPANTHERCODE IN (1,11) OR MT_UR.CHLAMYDIANAATPANTHERCODE IN (1,11) OR MT_FPU.CHLAMYDIANAATPANTHERCODE IN (1,11) ) " &
                 "AND C.SMSRESULTCONSENTCODE <> 2 " &
                 "GROUP BY C.URNO, C.MOBILE, MT.LABNO, C.SMSRESULTCONSENTCODE " &
                 "ORDER BY C.URNO, C.MOBILE, MT.LABNO, C.SMSRESULTCONSENTCODE"

        dsCX = returnFirebirdDataSet(strSQL, "Chlamydia_Oth")

        If dsCX.Tables(0).Rows.Count <> 0 Then
            ProcessChlamydia(dsCX)
        Else
            strMessage2 = "No records to be processed!"
            strMessage2 = ReplaceQuote(strMessage2)
            strSQL2 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage2 & "') "
            dsTemp2 = returnSQLDataSet(strSQL2)
            Console.WriteLine("No records to be processed!")
        End If
    End Sub

    Private Sub ProcessDatabase_Rec()
        Dim dsCS_Rec As DataSet
        Dim dtReminderPeriod_Rec As String
        Dim strSQL_Rec As String
        Dim intDaysREACT As Integer = Convert.ToInt32(gDaysREACT) * -1
        Dim strSQL3 As String = ""
        Dim strMessage3 As String = ""
        Dim dsTemp3 As DataSet

        dtReminderPeriod_Rec = Format(Date.Today.AddDays(intDaysREACT), "yyyy-MM-dd")   '42 days

        'Get list of clients who have positive REC chlamydia
        strSQL_Rec = "SELECT C.URNO, C.MOBILE, MT.LABNO, C.SMSRESULTCONSENTCODE " &
                 "FROM CLIENT C " &
                 "JOIN MDUSERVICE MS ON (MS.OWNER_OID = C.OID) " &
                 "JOIN MDUTEST MT ON (MT.OWNER_OID = MS.OID) " &
                 "LEFT JOIN MDUTEST_REC MT_REC ON (MT_REC.OID = MT.OID) " &
                 "WHERE CAST (MS.CREATEDBYDATETIME AS DATE) >= '" & dtReminderPeriod_Rec & "' " &
                 "AND CAST (MS.CREATEDBYDATETIME AS DATE) <= '" & dtReminderPeriod_Rec & " ' " &
                 "AND (C.URNO NOT IN (0,99999,888885,888886,888887,888888,999997,999998,999999)) " &
                 "AND MT_REC.CHLAMYDIANAATPANTHERCODE IN (1,11) " &
                 "AND C.SMSRESULTCONSENTCODE <> 2 " &
                 "GROUP BY C.URNO, C.MOBILE, MT.LABNO, C.SMSRESULTCONSENTCODE " &
                 "ORDER BY C.URNO, C.MOBILE, MT.LABNO, C.SMSRESULTCONSENTCODE"

        dsCS_Rec = returnFirebirdDataSet(strSQL_Rec, "Chlamydia_Rec")

        strMessage3 = "Preparing Chlamydia Rec SMS Reminder for " & Format(Date.Today.AddDays(intDaysREACT), "dd/MM/yyyy") & " on " & Format(Date.Today)
        strMessage3 = ReplaceQuote(strMessage3)
        strSQL3 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage3 & "') "
        dsTemp3 = returnSQLDataSet(strSQL3)

        If dsCS_Rec.Tables(0).Rows.Count <> 0 Then
            ProcessChlamydia(dsCS_Rec)
        Else
            strMessage3 = "No records to be processed!"
            strMessage3 = ReplaceQuote(strMessage3)
            strSQL3 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage3 & "') "
            dsTemp3 = returnSQLDataSet(strSQL3)
            Console.WriteLine("No records to be processed!")
        End If
    End Sub

    Private Sub ProcessChlamydia(ByVal dsData As DataSet)
        Dim x As Integer
        Dim strURNO As String
        Dim strMobile As String
        Dim strLabNo As String
        Dim intSMSConsentCode As Integer

        Dim strSQL4 As String = ""
        Dim strMessage4 As String = ""
        Dim dsTemp4 As DataSet

        '******* Database Items
        ' URNO               0
        ' Mobile             1
        ' Lab No             2
        ' SMS Consent Code   3

        For x = 0 To dsData.Tables(0).Rows.Count - 1
            strURNO = dsData.Tables(0).Rows(x).Item(0).ToString
            strLabNo = dsData.Tables(0).Rows(x).Item(2).ToString
            intSMSConsentCode = dsData.Tables(0).Rows(0).Item(3)

            If IsDBNull(dsData.Tables(0).Rows(x).Item(1)) = True Then
                strMessage4 = "Error sending SMS to " & strURNO & " - dbnull. Lab No: " & strLabNo
                strMessage4 = ReplaceQuote(strMessage4)
                strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                dsTemp4 = returnSQLDataSet(strSQL4)
            ElseIf dsData.Tables(0).Rows(x).Item(1).ToString = "DUPLICATE" Then
                strMessage4 = "SMS sent previously for Rec CX - " & strURNO & ". Lab No: " & strLabNo
                strMessage4 = ReplaceQuote(strMessage4)
                strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                dsTemp4 = returnSQLDataSet(strSQL4)
            ElseIf intSMSConsentCode = 2 Then
                strMessage4 = "Decline sending SMS to " & strURNO & ". Lab No: " & strLabNo
                strMessage4 = ReplaceQuote(strMessage4)
                strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                dsTemp4 = returnSQLDataSet(strSQL4)
            Else
                strMobile = Replace(dsData.Tables(0).Rows(x).Item(1).ToString, " ", "")

                If strMobile = "" Then
                    strMessage4 = "Error sending SMS to " & strURNO & " - no mobile number. Lab No: " & strLabNo
                    strMessage4 = ReplaceQuote(strMessage4)
                    strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                    dsTemp4 = returnSQLDataSet(strSQL4)
                Else
                    If checkMobileNoRegex(strMobile) Then
                        'Send SMS
                        If Client_SendSMS_New(strMobile, strURNO) = True Then
                            strMessage4 = "SMS successfully sent to " & strURNO & " Lab No: " & strLabNo
                            strMessage4 = ReplaceQuote(strMessage4)
                            strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                            dsTemp4 = returnSQLDataSet(strSQL4)
                        Else
                            strMessage4 = "Error sending SMS to " & strURNO & " (" & strMobile & ") - SMS error. Lab No: " & strLabNo
                            strMessage4 = ReplaceQuote(strMessage4)
                            strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                            dsTemp4 = returnSQLDataSet(strSQL4)
                        End If
                    Else
                        strMessage4 = "Error sending SMS to " & strURNO & " (" & strMobile & ") - invalid number. Lab No: " & strLabNo
                        strMessage4 = ReplaceQuote(strMessage4)
                        strSQL4 = "INSERT INTO [REACT-SMS-Log] VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & "', '" & strMessage4 & "') "
                        dsTemp4 = returnSQLDataSet(strSQL4)
                    End If
                End If
            End If
        Next x
    End Sub

    'Matches only Australian mobile numbers
    Private Function checkMobileNoRegex(ByRef strNumber As String) As Boolean
        If Not Regex.IsMatch(strNumber, "^04\d{8}$") Then
            checkMobileNoRegex = False
        Else
            checkMobileNoRegex = True
        End If
    End Function

    'Connect to CPMS database and execute SQL syntax
    Private Function returnFirebirdDataSet(ByVal strSQL As String, ByVal strDataSet As String) As DataSet
        Dim dsTemp As New DataSet()
        Dim conn As New FbConnection("User=SYSDBA;Password=masterkey;Database=" & strCPMSPath & ";DataSource=" & strCPMSServer & ";Charset=NONE;ServerType=0")
        Dim adapter As New FbDataAdapter()
        Dim cmd As New FbCommand(strSQL, conn)

        adapter.SelectCommand = cmd
        adapter.Fill(dsTemp, strDataSet)
        conn.Close()

        Return dsTemp
    End Function

    Public Function returnSQLDataSet(ByVal strSQL As String) As Data.DataSet
        Dim dsTemp As New DataSet()
        Dim conn As New SqlClient.SqlConnection(gDBConnectionString)
        Dim adapter As New SqlClient.SqlDataAdapter()
        Dim cmd As New SqlClient.SqlCommand(strSQL, conn)

        cmd.CommandTimeout = 2000
        adapter.SelectCommand = cmd
        adapter.Fill(dsTemp, "iMSHC")
        conn.Close()

        returnSQLDataSet = dsTemp
    End Function

    'Replace single quote with double quote
    Private Function ReplaceQuote(ByVal strToBeReplaced As String) As String
        Return Replace(strToBeReplaced, "'", "''")
    End Function

    'Send email
    Private Sub SendEmail(ByVal strTo As String, ByVal strHTMLBody As String, ByVal strSubject As String)
        Dim objMM As New MailMessage
        Dim objSMTP As New SmtpClient
        Dim strMailSvr As String = strEmail_server

        Try
            '20201025
            'objMM.From = New MailAddress(strEmail_address, "IT Services - MSHC")
            objMM.From = New MailAddress(strEmail_address)

            objMM.To.Add(strTo)
            'objMM.To.Add(strEmail_address)

            objMM.Subject = strSubject
            objMM.Body = strHTMLBody
            objMM.IsBodyHtml = True

            objSMTP.Host = strMailSvr
            objSMTP.UseDefaultCredentials = True

            objSMTP.DeliveryMethod = SmtpDeliveryMethod.Network
            objSMTP.Send(objMM)
        Catch ex As SmtpException
            'do nothing
        Catch ex As Exception
            'do nothing
        End Try
    End Sub

    'Send SMS
    Private Function Client_SendSMS_New(ByVal strMobile As String, ByVal strURNO As String) As Boolean
        Try
            Dim client As WebClient = New WebClient

            client.QueryString.Add("username", gSMSUsername)
            client.QueryString.Add("password", gSMSPassword)
            client.QueryString.Add("message", gSMSMessage)
            client.QueryString.Add("type", gSMSType)
            client.QueryString.Add("senderid", gSMSSenderID)
            client.QueryString.Add("to", strMobile)

            Dim url As String = gSMSServer
            Dim stream As Stream = client.OpenRead(url)
            Dim reader As StreamReader = New StreamReader(stream)
            Dim response As String = reader.ReadToEnd() 'Trap response

            stream.Close()
            reader.Close()

            If Left(response, 3) = "err" Then
                SendEmail(strEmail_address, "Error sending Chlamydia SMS Reminder to " & strURNO & " (" & strMobile & "): Error connecting to SMS gateway.", gEmailSubjectREACT)
                Return False
            ElseIf Left(response, 2) = "id" Then
                Return True
            Else
                SendEmail(strEmail_address, "Error sending Chlamydia SMS Reminder to " & strURNO & " (" & strMobile & "): Error connecting to SMS gateway.", gEmailSubjectREACT)
                Return False
            End If
        Catch ex As Exception
            SendEmail(strEmail_address, "Error sending Chlamydia SMS Reminder to " & strURNO & " (" & strMobile & "): " & ex.Message, gEmailSubjectREACT)
            Return False
        End Try
    End Function

End Module
