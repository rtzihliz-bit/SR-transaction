Imports MySql.Data.MySqlClient

Public Class ServiceRequestTransaction

    Public EffDate As String = ""

    'Combobox Name
    Public Sub cmbName()

        Try
            openMySQL()

            Dim strsqlcmb As String = "SELECT CONCAT(LNAME, ', ', FNAME, ' ', MIDDLE) as FULLNAME FROM concfile ORDER by FULLNAME ASC"
            Dim cmd1 As New MySqlCommand
            cmd1.Connection = conn
            cmd1.CommandText = strsqlcmb

            Dim cmbds As New DataSet
            Dim cmbda As New MySqlDataAdapter
            cmbda.SelectCommand = cmd1
            cmbda.Fill(cmbds, "concfile")

            ComboBox_ConName.DataSource = cmbds.DefaultViewManager
            ComboBox_ConName.DisplayMember = "concfile.FULLNAME"
            'ComboBox_ConName.ValueMember = "concfile.ID"

            conn.Close()
        Catch ex As Exception

            MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try

    End Sub

    'Public Sub ORNumber()
    '    openMySQL()
    '    Dim sql2 As String = "SELECT SRNUM FROM refe2"
    '    Dim cmd2 As New MySqlCommand(sql2, conn)
    '    da = New MySqlDataAdapter(cmd2)
    '    ds = New DataSet
    '    da.Fill(ds)
    '    dt = ds.Tables(0)
    '    Dim find2 As String = cmd2.ExecuteScalar
    '    TextBox_Num2.Text = find2

    '    Do While TextBox_Checker.Text = ""

    '        Dim strsqlchanged As String = "SELECT SRNO FROM servrequ WHERE SRNO = '" & TextBox_Num2.Text & "'"
    '        Dim txcda As MySqlDataAdapter = New MySqlDataAdapter(strsqlchanged, conn)
    '        Dim txcds As DataSet = New DataSet
    '        txcda.Fill(txcds, "servrequ")
    '        Dim txcdt As DataTable = txcds.Tables("servrequ")

    '        Dim row As DataRow

    '        For Each row In txcdt.Rows

    '            TextBox_Checker.Text = row("SRNO")

    '        Next row

    '        If TextBox_Checker.Text <> "" Then
    '            TextBox_Num2.Text = Val(TextBox_Checker.Text) + 1
    '            TextBox_Checker.Clear()
    '        Else
    '            TextBox_Checker.Text = "X"
    '        End If

    '    Loop

    '    Do While TextBox_Checker.Text = "X"
    '        TextBox_SRNo.Text = TextBox_Num2.Text
    '        TextBox_Checker.Text = "Y"
    '    Loop

    '    conn.Close()
    'End Sub

    Public Sub ORNumber()

        TextBox_Checker.Clear()

        openMySQL()
        Dim sql2 As String = "SELECT SRNUM FROM refe2"
        Dim cmd2 As New MySqlCommand(sql2, conn)
        da = New MySqlDataAdapter(cmd2)
        ds = New DataSet
        da.Fill(ds)
        dt = ds.Tables(0)
        Dim find2 As String = cmd2.ExecuteScalar
        Dim s As String
        s = find2.PadLeft(4, "0")
        find2 = s

        TextBox_Num2.Text = Format(DateTimePicker_TranDate.Value, "yy") + "-" + find2

        Do While TextBox_Checker.Text = ""

            Dim strsqlchanged As String = "SELECT SRNO FROM servrequ WHERE SRNO = '" & TextBox_Num2.Text & "'"
            Dim txcda As MySqlDataAdapter = New MySqlDataAdapter(strsqlchanged, conn)
            Dim txcds As DataSet = New DataSet
            txcda.Fill(txcds, "servrequ")
            Dim txcdt As DataTable = txcds.Tables("servrequ")

            Dim row As DataRow

            For Each row In txcdt.Rows

                TextBox_Checker.Text = row("SRNO")

            Next row

            If TextBox_Checker.Text <> "" Then

                '
                Dim sql2X = "UPDATE refe2 SET SRNUM = SRNUM + 1"
                Dim cmd2X As New MySqlCommand(sql2X, conn)
                Dim reader2X As MySqlDataReader
                reader2X = cmd2X.ExecuteReader
                reader2X.Close()

                Dim xsql2 As String = "SELECT SRNUM FROM refe2"
                Dim xcmd2 As New MySqlCommand(xsql2, conn)
                da = New MySqlDataAdapter(xcmd2)
                ds = New DataSet
                da.Fill(ds)
                dt = ds.Tables(0)
                Dim xfind2 As String = xcmd2.ExecuteScalar
                Dim xs As String
                xs = xfind2.PadLeft(4, "0")
                xfind2 = xs

                TextBox_Num2.Text = Format(DateTimePicker_TranDate.Value, "yy") + "-" + xfind2
                '

                TextBox_Checker.Clear()
            Else
                TextBox_Checker.Text = "X"
            End If

        Loop

        Do While TextBox_Checker.Text = "X"
            TextBox_SRNo.Text = TextBox_Num2.Text
            TextBox_Checker.Text = "Y"
        Loop

        conn.Close()

    End Sub

    Public Sub Clear()
        TextBox_Account.Clear()
        TextBox_ConName.Clear()
        ComboBox_ConName.SelectedIndex = -1
        TextBox_Address.Clear()
        DateTimePicker_TranDate.ResetText()
        TextBox_MeterNo.Clear()
        TextBox_DateInst.Clear()
        CB_Dirty.Checked = False
        CB_BallValve.Checked = False
        CB_StandPipe.Checked = False
        CB_LowPress.Checked = False
        CB_MeterLeak.Checked = False
        CB_NoWater.Checked = False
        CB_ReRead.Checked = False
        CB_TempoDisco.Checked = False

        CheckBox_CheckLeaks.Checked = False
        CheckBox_SLLeak.Checked = False
        CheckBox_TailPiece.Checked = False
        CheckBox_TappLeak.Checked = False
        CheckBox_SurvReco.Checked = False
        CheckBox_SurvRelo.Checked = False
        CheckBox_Cutoff.Checked = False

        ComboBox_SizeLine.SelectedIndex = -1
        ComboBox_ClassLine.SelectedIndex = -1

        TextBox_Others.Clear()
        'TextBox_Received.Clear()
        TextBox_Reviewed.Clear()
        ComboBox_Plumber.SelectedIndex = -1
        'TextBox_Approved.Clear()
        ComboBox_ServiceStatus.SelectedIndex = -1
        ComboBox_ActionTaken.SelectedIndex = -1
        DateTimePicker_Completed.ResetText()
        TextBox_ServiceReq.Clear()
        TextBox_Reading.Text = "0"
        Label_Read.Hide()
        TextBox_Reading.Hide()

        Button1.Hide()
        DateTimePicker_Finished.Hide()
        DateTimePicker_Finished.ResetText()
        TextBox1.Show()
        Button2.Show()
        DateTimePicker_Start.Show()
        DateTimePicker_Start.ResetText()
        TextBox2.Hide()
        Button3.Hide()
        DateTimePicker_Completed.Hide()
        TextBox3.Show()
        Button4.Hide()
        DateTimePicker_ActStarted.Hide()
        TextBox4.Show()

        TextBox_Zone.Text = "0"

        Button_Print.Enabled = False
        Button_Save.Enabled = True

        TextBox_Checker.Clear()

        ComboBox_WOC1.SelectedIndex = -1
        ComboBox_WOC2.SelectedIndex = -1

        TextBox_TelNo.Clear()
        TextBox_CountSkilled.Text = 0
        TextBox_CountUnskilled.Text = 0

        ORNumber()
    End Sub

    Public Sub Clear1()
        TextBox_Address.Clear()
        DateTimePicker_TranDate.ResetText()
        TextBox_MeterNo.Clear()
        TextBox_DateInst.Clear()
        CB_Dirty.Checked = False
        CB_BallValve.Checked = False
        CB_StandPipe.Checked = False
        CB_LowPress.Checked = False
        CB_MeterLeak.Checked = False
        CB_NoWater.Checked = False
        CB_ReRead.Checked = False
        CB_TempoDisco.Checked = False

        CheckBox_CheckLeaks.Checked = False
        CheckBox_SLLeak.Checked = False
        CheckBox_TailPiece.Checked = False
        CheckBox_TappLeak.Checked = False
        CheckBox_SurvReco.Checked = False
        CheckBox_SurvRelo.Checked = False
        CheckBox_Cutoff.Checked = False

        ComboBox_SizeLine.SelectedIndex = -1
        ComboBox_ClassLine.SelectedIndex = -1

        TextBox_Others.Clear()
        'TextBox_Received.Clear()
        TextBox_Reviewed.Clear()
        ComboBox_Plumber.SelectedIndex = -1
        'TextBox_Approved.Clear()
        ComboBox_ServiceStatus.SelectedIndex = -1
        ComboBox_ActionTaken.SelectedIndex = -1
        DateTimePicker_Completed.ResetText()
        TextBox_ServiceReq.Clear()
        TextBox_Reading.Text = "0"
        Label_Read.Hide()
        TextBox_Reading.Hide()

        Button1.Hide()
        DateTimePicker_Finished.Hide()
        DateTimePicker_Finished.ResetText()
        TextBox1.Show()
        Button2.Show()
        DateTimePicker_Start.Show()
        DateTimePicker_Start.ResetText()
        TextBox2.Hide()
        Button3.Hide()
        DateTimePicker_Completed.Hide()
        TextBox3.Show()
        Button4.Hide()
        DateTimePicker_ActStarted.Hide()
        TextBox4.Show()


        TextBox_Zone.Text = "0"

        Button_Print.Enabled = False
        Button_Save.Enabled = True

        TextBox_Checker.Clear()

        ComboBox_WOC1.SelectedIndex = -1
        ComboBox_WOC2.SelectedIndex = -1

        TextBox_TelNo.Clear()
        TextBox_CountSkilled.Text = 0
        TextBox_CountUnskilled.Text = 0

        ORNumber()
    End Sub

    Private Sub Button_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Clear.Click
        Clear()
    End Sub

    Private Sub Account_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox_Account.KeyPress, TextBox_Reading.KeyPress

        If Char.IsPunctuation(e.KeyChar) = False And Char.IsNumber(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If

    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox_ConName.KeyPress, TextBox_Others.KeyPress, ComboBox_ActionTaken.KeyPress, ComboBox_Plumber.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ServiceRequestTransaction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            openMySQL()
            Dim strsqlcmb As String = "SELECT NAME FROM plumber ORDER by NAME ASC"

            Dim cmd1 As New MySqlCommand
            cmd1.Connection = conn
            cmd1.CommandText = strsqlcmb

            Dim cmbds As New DataSet
            Dim cmbda As New MySqlDataAdapter
            cmbda.SelectCommand = cmd1
            cmbda.Fill(cmbds, "plumber")

            ComboBox_Plumber.DataSource = cmbds.DefaultViewManager
            ComboBox_Plumber.DisplayMember = "plumber.NAME"
            'ComboBox_Zone.ValueMember = "code.bgyname"

            ComboBox_Plumber.SelectedIndex = -1

            Dim strsqlcmbV As String = "SELECT WORK FROM workorder ORDER by WORK ASC"
            Dim cmd1V As New MySqlCommand
            cmd1V.Connection = conn
            cmd1V.CommandText = strsqlcmbV

            Dim cmbdsV As New DataSet
            Dim cmbdaV As New MySqlDataAdapter
            cmbdaV.SelectCommand = cmd1V
            cmbdaV.Fill(cmbdsV, "workorder")

            ComboBox_ActionTaken.DataSource = cmbdsV.DefaultViewManager
            ComboBox_ActionTaken.DisplayMember = "workorder.WORK"

            ComboBox_ActionTaken.SelectedIndex = -1

            conn.Close()
        Catch ex As Exception

            MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try

        RadioButton1.Checked = True
        ComboBox_ConName.Hide()

        cmbName()
        Clear()

        Button1.Hide()
        Button2.Show()
        Button3.Hide()
        DateTimePicker_Start.Show()
        DateTimePicker_Finished.Hide()
        DateTimePicker_Completed.Hide()
        TextBox1.Show()
        TextBox2.Hide()
        TextBox3.Show()

        ORNumber()

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False

            Clear()

            TextBox_ConName.Show()
            TextBox_Account.ReadOnly = False
            ComboBox_ConName.Hide()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton1.Checked = False Then
            RadioButton2.Checked = True

            Clear()

            ComboBox_ConName.Show()
            TextBox_Account.ReadOnly = True
            TextBox_ConName.Hide()
        End If
    End Sub

    Private Sub TextBox_acctno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Account.TextChanged

        Clear1()
        TextBox_ConName.Clear()

        If TextBox_Account.Text = "" Then
            Clear()
        Else
            Try
                openMySQL()
                Dim strsqlchanged As String = "SELECT CONCAT(LNAME, ', ', FNAME, ' ', MIDDLE) as FULLNAME, ACCTNUM, ADDRESS, TELNO, MTRSERIAL, ZN, DTINSTLD FROM concfile WHERE ACCTNUM = '" & TextBox_Account.Text & "'"
                Dim txcda As MySqlDataAdapter = New MySqlDataAdapter(strsqlchanged, conn)
                Dim txcds As DataSet = New DataSet
                txcda.Fill(txcds, "concfile")
                Dim txcdt As DataTable = txcds.Tables("concfile")

                Dim row As DataRow

                For Each row In txcdt.Rows

                    If row("DTINSTLD") Is System.DBNull.Value Then 'NO ACTION IF OUTPUT FROM DATABASE IS NULL

                    Else
                        TextBox_DateInst.Text = row("DTINSTLD")
                        DateTimePicker_DateInst.Text = row("DTINSTLD")
                    End If

                    TextBox_ConName.Text = row("FULLNAME")
                    TextBox_Account.Text = row("ACCTNUM")
                    TextBox_Address.Text = row("ADDRESS")
                    TextBox_Zone.Text = row("ZN")
                    TextBox_MeterNo.Text = row("MTRSERIAL")
                    TextBox_TelNo.Text = row("TELNO")

                Next row


                conn.Close()
            Catch ex As Exception

                MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End Try

            ORNumber()

            'Try
            '    openMySQL()
            '    strsql = "SELECT GENMGR FROM signfile"

            '    conn = New MySqlConnection(connStr)
            '    conn.Open()

            '    Dim cmd As New MySqlCommand(strsql, conn)
            '    da = New MySqlDataAdapter(cmd)
            '    ds = New DataSet
            '    da.Fill(ds)
            '    dt = ds.Tables(0)

            '    Dim find As String = cmd.ExecuteScalar

            '    TextBox_Approved.Text = find

            '    conn.Close()
            'Catch ex As Exception

            '    MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            'End Try
        End If

    End Sub

    Private Sub ComboBox_Name_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox_ConName.SelectedIndexChanged, ComboBox_ConName.TextChanged

        If ComboBox_ConName.Text = "" Then
            Clear()
        End If

        Clear1()

        Try
            openMySQL()
            Dim strsqlchanged As String = "SELECT ACCTNUM, ADDRESS, TELNO, MTRSERIAL, ZN, DTINSTLD FROM concfile WHERE (CONCAT(LNAME, ', ', FNAME, ' ', MIDDLE)) = '" & ComboBox_ConName.Text & "'"

            Dim txcda As MySqlDataAdapter = New MySqlDataAdapter(strsqlchanged, conn)
            Dim txcds As DataSet = New DataSet
            txcda.Fill(txcds, "concfile")
            Dim txcdt As DataTable = txcds.Tables("concfile")

            Dim row As DataRow

            For Each row In txcdt.Rows

                If row("DTINSTLD") Is System.DBNull.Value Then 'NO ACTION IF OUTPUT FROM DATABASE IS NULL

                Else
                    TextBox_DateInst.Text = row("DTINSTLD")
                    DateTimePicker_DateInst.Text = row("DTINSTLD")
                End If

                TextBox_Account.Text = row("ACCTNUM")
                TextBox_Address.Text = row("ADDRESS")

                If row("ZN") Is System.DBNull.Value Then 'NO ACTION IF OUTPUT FROM DATABASE IS NULL

                Else
                    TextBox_Zone.Text = row("ZN")
                End If

                TextBox_MeterNo.Text = row("MTRSERIAL")
                TextBox_TelNo.Text = row("TELNO")

            Next row

            conn.Close()
        Catch ex As Exception

            MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try


        ORNumber()

        'Try
        '    openMySQL()
        '    strsql = "SELECT GENMGR FROM signfile"

        '    conn = New MySqlConnection(connStr)
        '    conn.Open()

        '    Dim cmd As New MySqlCommand(strsql, conn)
        '    da = New MySqlDataAdapter(cmd)
        '    ds = New DataSet
        '    da.Fill(ds)
        '    dt = ds.Tables(0)

        '    Dim find As String = cmd.ExecuteScalar

        '    TextBox_Approved.Text = find

        '    conn.Close()
        'Catch ex As Exception

        '    MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        'End Try

    End Sub

    Private Sub Button_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Save.Click

     
        If RadioButton1.Checked = True And TextBox_ServiceReq.Text <> "" Then

            Button_Save.Enabled = False

            ORNumber()

            Try
                openMySQL()
                Dim sql As String = "INSERT into servrequ VALUES('" & TextBox_Account.Text & "', '" & TextBox_ConName.Text & "', '" & TextBox_Address.Text & "', '" & Format(DateTimePicker_TranDate.Value, "yyyy-MM-dd HH:mm:ss") & "', '" & TextBox_SRNo.Text & "', '" & TextBox_ServiceReq.Text & "', '" & TextBox_Received.Text & "', '" & TextBox_Approved.Text & "', '" & TextBox_Reviewed.Text & "', '" & ComboBox_ActionTaken.Text & "', '" & ComboBox_Plumber.Text & "', '" & ComboBox_ServiceStatus.Text & "', NULL, '" & TextBox_Zone.Text & "', '" & Format(DateTimePicker_TranDate.Value, "yyyy-MM-dd HH:mm:ss") & "', NULL, NULL, '" & ComboBox_SizeLine.Text & "', '" & ComboBox_ClassLine.Text & "', '" & TextBox_MeterNo.Text & "', NULL, NULL, '', '', '" & Format(DateTimePicker_Start.Value, "HH:mm") & "', '" & ComboBox_WOC1.Text & "', '" & ComboBox_WOC2.Text & "', '', '" & TextBox_TelNo.Text & "', '" & TextBox_CountSkilled.Text & "', '" & TextBox_CountUnskilled.Text & "', '', NULL, NULL, NULL, NULL, '', '', '0', '', '')"
                Dim cmd As New MySqlCommand(sql, conn)
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader
                reader.Close()

                If TextBox_DateInst.Text <> "" Then
                    Dim sql1 = "UPDATE servrequ SET DTINSTLD = '" & Format(DateTimePicker_DateInst.Value, "yyyy-MM-dd") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_Completed.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET COMPLETED = '" & Format(DateTimePicker_Completed.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_Start.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET DATESTART = '" & Format(DateTimePicker_Start.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_ActStarted.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET ACTSTART = '" & Format(DateTimePicker_ActStarted.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_Finished.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET DATEFINISHED = '" & Format(DateTimePicker_Finished.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                ''
                If TextBox_ServiceReq.Text = "TEMPO. DISCONNECTION" Then
                    Dim sql1 = "UPDATE concfile SET REMARKS = 'DISCONNECTED', DISCODATE = '" & Format(DateTimePicker_Completed.Value, "yyyy-MM-dd") & "', STATUS = 'INACTIVE' WHERE ACCTNUM = '" & TextBox_Account.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If TextBox_ServiceReq.Text = "FOR RECONNECTION" Then
                    'Dim sql2x As String = "INSERT into discreco VALUES('" & TextBox_Account.Text & "', '" & TextBox_ConName.Text & "', '" & TextBox_Address.Text & "', '" & Format(DateTimePicker_TranDate.Value, "yyyy-MM-dd") & "', '" & TextBox_MeterNo.Text & "', '" & TextBox_Reading.Text & "', 'RECONNECTED', '" & TextBox_Zone.Text & "', NULL)"
                    'Dim cmd2x As New MySqlCommand(sql2x, conn)
                    'Dim reader2x As MySqlDataReader
                    'reader2x = cmd2x.ExecuteReader
                    'reader2x.Close()

                    Dim sql1 = "UPDATE concfile SET REMARKS = 'RECONNECTED', STATUS = 'ACTIVE' WHERE ACCTNUM = '" & TextBox_Account.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If
                ''

                'Dim sql2 = "UPDATE refe2 SET SRNUM = '" & TextBox_SRNo.Text + 1 & "'"
                'Dim cmd2 As New MySqlCommand(sql2, conn)
                'Dim reader2 As MySqlDataReader
                'reader2 = cmd2.ExecuteReader
                'reader2.Close()

                Dim sql2 = "UPDATE refe2 SET SRNUM = SRNUM + 1"
                Dim cmd2 As New MySqlCommand(sql2, conn)
                Dim reader2 As MySqlDataReader
                reader2 = cmd2.ExecuteReader
                reader2.Close()

                'TextBox_Checker.Clear()

                'ORNumber()

                ReferenceMasterFileMaint.UpdateDGVIew()

                MessageBox.Show("Successfully Saved", GenericName, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Button_Print.Enabled = True
                conn.Close()
            Catch ex As MySqlException
                MessageBox.Show("Error connecting to the server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try

        ElseIf RadioButton1.Checked = False And TextBox_ServiceReq.Text <> "" Then

            Button_Save.Enabled = False

            ORNumber()

            Try
                openMySQL()
                Dim sql As String = "INSERT into servrequ VALUES('" & TextBox_Account.Text & "', '" & ComboBox_ConName.Text & "', '" & TextBox_Address.Text & "', '" & Format(DateTimePicker_TranDate.Value, "yyyy-MM-dd HH:mm:ss") & "', '" & TextBox_SRNo.Text & "', '" & TextBox_ServiceReq.Text & "', '" & TextBox_Received.Text & "', '" & TextBox_Approved.Text & "', '" & TextBox_Reviewed.Text & "', '" & ComboBox_ActionTaken.Text & "', '" & ComboBox_Plumber.Text & "', '" & ComboBox_ServiceStatus.Text & "', NULL, '" & TextBox_Zone.Text & "', '" & Format(DateTimePicker_TranDate.Value, "yyyy-MM-dd HH:mm:ss") & "', NULL, NULL, '" & ComboBox_SizeLine.Text & "', '" & ComboBox_ClassLine.Text & "', '" & TextBox_MeterNo.Text & "', NULL, NULL, '', '', '" & Format(DateTimePicker_Start.Value, "HH:mm") & "', '" & ComboBox_WOC1.Text & "', '" & ComboBox_WOC2.Text & "', '', '" & TextBox_TelNo.Text & "', '" & TextBox_CountSkilled.Text & "', '" & TextBox_CountUnskilled.Text & "', '', NULL, NULL, NULL, NULL, '', '', '0', '', '')"
                Dim cmd As New MySqlCommand(sql, conn)
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader
                reader.Close()

                If TextBox_DateInst.Text <> "" Then
                    Dim sql1 = "UPDATE servrequ SET DTINSTLD = '" & Format(DateTimePicker_DateInst.Value, "yyyy-MM-dd") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_Completed.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET COMPLETED = '" & Format(DateTimePicker_Completed.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_Start.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET DATESTART = '" & Format(DateTimePicker_Start.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_ActStarted.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET ACTSTART = '" & Format(DateTimePicker_ActStarted.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If DateTimePicker_Finished.Visible = True Then
                    Dim sql1 = "UPDATE servrequ SET DATEFINISHED = '" & Format(DateTimePicker_Finished.Value, "yyyy-MM-dd HH:mm:ss") & "' WHERE SRNO = '" & TextBox_SRNo.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                ''
                If TextBox_ServiceReq.Text = "TEMPO. DISCONNECTION" Then
                    Dim sql1 = "UPDATE concfile SET REMARKS = 'DISCONNECTED', DISCODATE = '" & Format(DateTimePicker_Completed.Value, "yyyy-MM-dd") & "', STATUS = 'INACTIVE' WHERE ACCTNUM = '" & TextBox_Account.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If

                If TextBox_ServiceReq.Text = "FOR RECONNECTION" Then
                    'Dim sql2x As String = "INSERT into discreco VALUES('" & TextBox_Account.Text & "', '" & ComboBox_ConName.Text & "', '" & TextBox_Address.Text & "', '" & Format(DateTimePicker_TranDate.Value, "yyyy-MM-dd") & "', '" & TextBox_MeterNo.Text & "', '" & TextBox_Reading.Text & "', 'RECONNECTED', '" & TextBox_Zone.Text & "', NULL)"
                    'Dim cmd2x As New MySqlCommand(sql2x, conn)
                    'Dim reader2x As MySqlDataReader
                    'reader2x = cmd2x.ExecuteReader
                    'reader2x.Close()

                    Dim sql1 = "UPDATE concfile SET REMARKS = 'RECONNECTED', STATUS = 'ACTIVE' WHERE ACCTNUM = '" & TextBox_Account.Text & "'"
                    Dim cmd1 As New MySqlCommand(sql1, conn)
                    Dim reader1 As MySqlDataReader
                    reader1 = cmd1.ExecuteReader
                    reader1.Close()
                End If
                ''

                'Dim sql2 = "UPDATE refe2 SET SRNUM = '" & TextBox_SRNo.Text + 1 & "'"
                'Dim cmd2 As New MySqlCommand(sql2, conn)
                'Dim reader2 As MySqlDataReader
                'reader2 = cmd2.ExecuteReader
                'reader2.Close()

                Dim sql2 = "UPDATE refe2 SET SRNUM = SRNUM + 1"
                Dim cmd2 As New MySqlCommand(sql2, conn)
                Dim reader2 As MySqlDataReader
                reader2 = cmd2.ExecuteReader
                reader2.Close()

                'TextBox_Checker.Clear()

                'ORNumber()

                ReferenceMasterFileMaint.UpdateDGVIew()

                MessageBox.Show("Successfully Saved", GenericName, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Button_Print.Enabled = True
                conn.Close()
            Catch ex As MySqlException
                MessageBox.Show("Error connecting to the server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try

        End If

    End Sub

    Private Sub CB_TasteOdor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_TempoDisco.CheckedChanged
        If CB_TempoDisco.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_NoWater.Checked = False
            CB_ReRead.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "TEMPO. DISCONNECTION"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_ReRead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_ReRead.CheckedChanged
        If CB_ReRead.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_NoWater.Checked = False
            CB_TempoDisco.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "FOR RECONNECTION"
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_NoWater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_NoWater.CheckedChanged
        If CB_NoWater.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "NO WATER"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_MeterLeak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_MeterLeak.CheckedChanged
        If CB_MeterLeak.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_NoWater.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "METER LEAK"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_LowPress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_LowPress.CheckedChanged
        If CB_LowPress.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_MeterLeak.Checked = False
            CB_NoWater.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "LOW PRESSURE"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_HighPress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_StandPipe.CheckedChanged
        If CB_StandPipe.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "STAND PIPE LEAK"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_HighCon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_BallValve.CheckedChanged
        If CB_BallValve.Checked = True Then
            CB_Dirty.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "BUSTED BALL VALVE"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CB_Dirty_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Dirty.CheckedChanged
        If CB_Dirty.Checked = True Then
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_NoWater.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "DIRTY WATER"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub TextBox_Others_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Others.Click
        CB_BallValve.Checked = False
        CB_StandPipe.Checked = False
        CB_LowPress.Checked = False
        CB_MeterLeak.Checked = False
        CB_NoWater.Checked = False
        CB_ReRead.Checked = False
        CB_TempoDisco.Checked = False
        CB_Dirty.Checked = False
        CheckBox_CheckLeaks.Checked = False
        CheckBox_SLLeak.Checked = False
        CheckBox_TailPiece.Checked = False
        CheckBox_TappLeak.Checked = False
        CheckBox_Cutoff.Checked = False
        CheckBox_SurvRelo.Checked = False
        CheckBox_SurvReco.Checked = False

        TextBox_ServiceReq.Text = TextBox_Others.Text

        ComboBox_ClassLine.SelectedIndex = -1
        ComboBox_SizeLine.SelectedIndex = -1
    End Sub

    Private Sub TextBox_Others_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Others.TextChanged
        CB_BallValve.Checked = False
        CB_StandPipe.Checked = False
        CB_LowPress.Checked = False
        CB_MeterLeak.Checked = False
        CB_NoWater.Checked = False
        CB_ReRead.Checked = False
        CB_TempoDisco.Checked = False
        CB_Dirty.Checked = False
        CheckBox_CheckLeaks.Checked = False
        CheckBox_SLLeak.Checked = False
        CheckBox_TailPiece.Checked = False
        CheckBox_TappLeak.Checked = False
        CheckBox_Cutoff.Checked = False
        CheckBox_SurvRelo.Checked = False
        CheckBox_SurvReco.Checked = False

        TextBox_ServiceReq.Text = TextBox_Others.Text
        Label_Read.Hide()
        TextBox_Reading.Hide()

        ComboBox_ClassLine.SelectedIndex = -1
        ComboBox_SizeLine.SelectedIndex = -1
    End Sub

    Private Sub TextBox21_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.Click
        Button1.Show()
        DateTimePicker_Finished.Show()
        TextBox1.Hide()
    End Sub

    Private Sub TextBox22_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.Click
        Button2.Show()
        DateTimePicker_Start.Show()
        TextBox2.Hide()
    End Sub

    Private Sub TextBox23_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.Click
        Button3.Show()
        DateTimePicker_Completed.Show()
        TextBox3.Hide()
    End Sub

    Private Sub TextBox24_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.Click
        Button4.Show()
        DateTimePicker_ActStarted.Show()
        TextBox4.Hide()
    End Sub

    Private Sub Button21_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button1.Hide()
        DateTimePicker_Finished.Hide()
        TextBox1.Show()
    End Sub

    Private Sub Button22_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button2.Hide()
        DateTimePicker_Start.Hide()
        TextBox2.Show()
    End Sub

    Private Sub Button23_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button3.Hide()
        DateTimePicker_Completed.Hide()
        TextBox3.Show()
    End Sub

    Private Sub Button24_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Button4.Hide()
        DateTimePicker_ActStarted.Hide()
        TextBox4.Show()
    End Sub

    Private Sub Button_View_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_View.Click
        ServiceRequestFile.MdiParent = Main
        ServiceRequestFile.WindowState = FormWindowState.Normal
        ServiceRequestFile.Show()

        ServiceRequestFile.RadioButton1.Checked = True
        If RadioButton1.Checked = True Then
            ServiceRequestFile.ComboBox_ConName.Text = TextBox_ConName.Text
        Else
            ServiceRequestFile.ComboBox_ConName.Text = ComboBox_ConName.Text
        End If
        'If ServiceRequestFile.Panel_Cons.Visible = True Then
        ServiceRequestFile.UpdateDGVIew1()
        'Else
        '    ServiceRequestFile.UpdateDGVIew()
        'End If
    End Sub

    Private Sub TextBox_Zone_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Zone.TextChanged
        If TextBox_Zone.Text = "" Then
            TextBox_Zone.Text = "0"
        End If
    End Sub

    Private Sub Button_Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Print.Click

        EffDate = ""

        Try
            openMySQL()
            Dim strsqlchanged As String = "SELECT EFFDATE FROM concfile WHERE ACCTNUM = '" & TextBox_Account.Text & "' and effdate < '2007-01-01'"
            Dim txcda As MySqlDataAdapter = New MySqlDataAdapter(strsqlchanged, conn)
            Dim txcds As DataSet = New DataSet
            txcda.Fill(txcds, "concfile")
            Dim txcdt As DataTable = txcds.Tables("concfile")

            Dim row As DataRow

            For Each row In txcdt.Rows

                If row("EFFDATE") Is System.DBNull.Value Then 'NO ACTION IF OUTPUT FROM DATABASE IS NULL

                Else
                    EffDate = row("EFFDATE")
                End If

            Next row

            conn.Close()
        Catch ex As Exception

            MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try


        Button_Print.Enabled = False

        Try
            openMySQL()
            Dim sql = "DROP TABLE IF EXISTS dummyservicerequest"
            Dim cmd As New MySqlCommand(sql, conn)
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader
            reader.Close()

            Dim sql1 = "CREATE TABLE dummyservicerequest SELECT ACCTNUM, NAME, ADDRESS, TELNO, DATE, DTINSTLD, SRNO, SERVICE, DATESTART, RECEIVED, APPROVED, REVIEWED, ACTION, PLUMBER, MTRNO, WO1, WO2, CLASSLINE, CLASSSIZE FROM servrequ WHERE SRNO = '" & TextBox_SRNo.Text & "'"
            Dim cmd1 As New MySqlCommand(sql1, conn)
            Dim reader1 As MySqlDataReader
            reader1 = cmd1.ExecuteReader
            reader1.Close()

            '
            Dim sqls = "DROP TABLE IF EXISTS dummyservicerequest1"
            Dim cmds As New MySqlCommand(sqls, conn)
            Dim readers As MySqlDataReader
            readers = cmds.ExecuteReader
            readers.Close()

            Dim sql1s = "CREATE TABLE dummyservicerequest1 SELECT ACCTNUM FROM concfile WHERE ACCTNUM = '" & TextBox_Account.Text & "' and effdate < '2007-01-01'"
            Dim cmd1s As New MySqlCommand(sql1s, conn)
            Dim reader1s As MySqlDataReader
            reader1s = cmd1s.ExecuteReader
            reader1s.Close()

            If EffDate = "" Then
                Dim sqlsX = "INSERT INTO dummyservicerequest1 VALUES ('NEW CONN.')"
                Dim cmdsX As New MySqlCommand(sqlsX, conn)
                Dim readersX As MySqlDataReader
                readersX = cmdsX.ExecuteReader
                readersX.Close()

            Else
                Dim sqlsX = "UPDATE dummyservicerequest1 SET ACCTNUM = 'RE-REGS.'"
                Dim cmdsX As New MySqlCommand(sqlsX, conn)
                Dim readersX As MySqlDataReader
                readersX = cmdsX.ExecuteReader
                readersX.Close()
            End If


            ORNumber()

            conn.Close()

        Catch ex As MySqlException
            MessageBox.Show("Error connecting to the server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        PrintSetup.MdiParent = Main
        PrintSetup.WindowState = FormWindowState.Maximized
        PrintSetup.Show()
        PrintSetup.CrystalReportViewer1.ReportSource = CrystalAddress + "CRServiceRequest.rpt"

        Button_Save.Enabled = True

    End Sub

    Private Sub TextBox_Reading_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Reading.TextChanged
        If TextBox_Reading.Text = "" Then
            TextBox_Reading.Text = "0"
            TextBox_Reading.SelectAll()
        End If

    End Sub

    Private Sub TextBox_Reading_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Reading.DoubleClick

        ConsumerStatusFile.MdiParent = Main
        ConsumerStatusFile.WindowState = FormWindowState.Normal
        ConsumerStatusFile.Visible = True

        ConsumerStatusFile.TextBox_Account.Text = TextBox_Account.Text

    End Sub


    Private Sub TextBox_in_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Reading.MouseDown
        TextBox_Reading.SelectAll()
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_SLLeak.CheckedChanged
        If CheckBox_SLLeak.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "S/L LEAK"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CheckBox_CheckLeaks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_CheckLeaks.CheckedChanged
        If CheckBox_CheckLeaks.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "CHECK LEAKS"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CheckBox_TailPiece_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_TailPiece.CheckedChanged
        If CheckBox_TailPiece.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "TAIL PIECE LEAK"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CheckBox_TappLeak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_TappLeak.CheckedChanged
        If CheckBox_TappLeak.Checked = True Then
            CB_Dirty.Checked = False
            CB_BallValve.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "TAPPING POINT LEAK"
            Label_Read.Hide()
            TextBox_Reading.Hide()
        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub ComboBox_ClassLine_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_ClassLine.SelectedIndexChanged
        If ComboBox_ClassLine.Text = "SERVICE LINE" Then
            ComboBox_SizeLine.Text = "1/2"
        ElseIf ComboBox_ClassLine.Text = "EXPANSION" Then
            ComboBox_SizeLine.Text = "0"
        ElseIf ComboBox_ClassLine.Text = "OTHERS" Then
            ComboBox_SizeLine.Text = "0"
        ElseIf ComboBox_ClassLine.Text = "WATER METER" Then
            ComboBox_SizeLine.Text = "1/2"
        Else
            ComboBox_SizeLine.SelectedIndex = -1
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_WOC1.SelectedIndexChanged
        If ComboBox_WOC1.Text = "BREAKAGE/CUT OFF" Then
            ComboBox_WOC2.Text = "REPAIR AND MAINTENANCE"
            ComboBox_ClassLine.SelectedIndex = -1
            ComboBox_SizeLine.SelectedIndex = -1
        ElseIf ComboBox_WOC1.Text = "LEAKS" Then
            ComboBox_WOC2.Text = "REPAIR AND MAINTENANCE"
            ComboBox_ClassLine.SelectedIndex = -1
            ComboBox_SizeLine.SelectedIndex = -1
        ElseIf ComboBox_WOC1.Text = "CUSTOMER SERVICE" Then
            ComboBox_WOC2.SelectedIndex = -1
            ComboBox_ClassLine.SelectedIndex = -1
            ComboBox_SizeLine.SelectedIndex = -1
        End If
    End Sub

    Private Sub ComboBox_WOC2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_WOC2.SelectedIndexChanged
        If ComboBox_WOC2.Text = "EXPANSION" Then
            ComboBox_SizeLine.Text = "0"
            ComboBox_ClassLine.Text = "EXPANSION"
        Else
            ComboBox_ClassLine.SelectedIndex = -1
            ComboBox_SizeLine.SelectedIndex = -1
        End If
    End Sub

    Private Sub DateTimePicker_Completed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker_Completed.ValueChanged
        DateTimePicker_ActStarted.Value = DateTimePicker_Completed.Value
        DateTimePicker_Finished.Value = DateTimePicker_Completed.Value
    End Sub

    Private Sub DateTimePicker_TranDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker_TranDate.ValueChanged
        DateTimePicker_Start.Value = DateTimePicker_TranDate.Value
    End Sub

    Private Sub DateTimePicker_Start_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker_Start.ValueChanged
        DateTimePicker_TranDate.Value = DateTimePicker_Start.Value
    End Sub

    Private Sub ComboBox_Plumber_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_Plumber.TextChanged, ComboBox_Plumber.SelectedIndexChanged

        If ComboBox_Plumber.Text <> "" Then

            Try
                openMySQL()
                Dim strsqlchanged As String = "SELECT CLASS FROM plumber WHERE NAME = '" & ComboBox_Plumber.Text & "'"
                Dim txcda As MySqlDataAdapter = New MySqlDataAdapter(strsqlchanged, conn)
                Dim txcds As DataSet = New DataSet
                txcda.Fill(txcds, "plumber")
                Dim txcdt As DataTable = txcds.Tables("plumber")

                Dim row As DataRow

                For Each row In txcdt.Rows

                    Dim PLUMCLASS As String

                    PLUMCLASS = row("CLASS")

                    If PLUMCLASS = "SKILLED" Then
                        TextBox_CountSkilled.Text = 1
                        TextBox_CountUnskilled.Text = 0
                    ElseIf PLUMCLASS = "UNSKILLED" Then
                        TextBox_CountUnskilled.Text = 1
                        TextBox_CountSkilled.Text = 0
                    Else
                        TextBox_CountUnskilled.Text = 0
                        TextBox_CountSkilled.Text = 0
                    End If

                Next row

                conn.Close()
            Catch ex As Exception

                MessageBox.Show("Error connecting to server: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End Try

        Else
            TextBox_CountUnskilled.Text = 0
            TextBox_CountSkilled.Text = 0

        End If

    End Sub

    Private Sub CheckBox_Cutoff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_Cutoff.CheckedChanged
        If CheckBox_Cutoff.Checked = True Then
            CB_Dirty.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CB_BallValve.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_SurvRelo.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "CUT OFF S/L"
            Label_Read.Hide()
            TextBox_Reading.Hide()

        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CheckBox_SurvRelo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_SurvRelo.CheckedChanged
        If CheckBox_SurvRelo.Checked = True Then
            CB_Dirty.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CB_BallValve.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvReco.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "SURVEY FOR RELOCATION"
            Label_Read.Hide()
            TextBox_Reading.Hide()

        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub

    Private Sub CheckBox_SurvReco_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_SurvReco.CheckedChanged
        If CheckBox_SurvReco.Checked = True Then
            CB_Dirty.Checked = False
            CB_StandPipe.Checked = False
            CB_LowPress.Checked = False
            CB_MeterLeak.Checked = False
            CB_ReRead.Checked = False
            CB_TempoDisco.Checked = False
            CB_NoWater.Checked = False
            CB_BallValve.Checked = False
            CheckBox_CheckLeaks.Checked = False
            CheckBox_SLLeak.Checked = False
            CheckBox_TailPiece.Checked = False
            CheckBox_TappLeak.Checked = False
            CheckBox_SurvRelo.Checked = False
            CheckBox_Cutoff.Checked = False

            TextBox_Others.Clear()
            TextBox_ServiceReq.Text = "SURVEY FOR RECONNECTION"
            Label_Read.Hide()
            TextBox_Reading.Hide()

        End If

        If CheckBox_TappLeak.Checked = False And CheckBox_TailPiece.Checked = False And CheckBox_SLLeak.Checked = False And CheckBox_CheckLeaks.Checked = False And CB_BallValve.Checked = False And CB_StandPipe.Checked = False And CB_LowPress.Checked = False And CB_MeterLeak.Checked = False And CB_NoWater.Checked = False And CB_ReRead.Checked = False And CB_TempoDisco.Checked = False And CB_Dirty.Checked = False And CheckBox_SurvRelo.Checked = False And CheckBox_SurvReco.Checked = False And CheckBox_Cutoff.Checked = False Then
            TextBox_ServiceReq.Text = ""
        End If
    End Sub
End Class