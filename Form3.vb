
Imports Microsoft.Reporting
Imports Microsoft.ReportingServices
Imports System.Data.SqlClient



Public Class Form3
    Public DV1 As Date
    Public a2 As Date
    Public a3 As Date


    Private Sub Prevent_Duplicates()

        '1st level check from app date picker
        'SQL  
        Dim SQLCONNECTION1 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1 As System.Data.SqlClient.SqlDataReader
        CMD1.CommandType = System.Data.CommandType.Text   'command syntax

        DV1 = DateTimePicker1.Value.Date

        CMD1.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}'", DV1)

        SQLCONNECTION1.Open()
        CMD1.Connection = SQLCONNECTION1
        SQLREADER1 = CMD1.ExecuteReader



        If SQLREADER1.HasRows Then
            'CHECK FOR DUPLICATE POSTINGS
            MsgBox("Automation File for the specified settlement date already uploaded", vbOKOnly, "Interswitch Automation")
            Button4.Visible = True
            Button3.Visible = False
            Button2.Visible = False
            ReportViewer1.Visible = False
            GoTo 1
        Else
            SQLCONNECTION1.Close()
            Generate_report()
            Exit Sub
        End If


1:      SQLCONNECTION1.Close()





    End Sub



    Private Sub Generate_report()

        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''The 3 lines below were code for extracting date from excel
        Dim CMD1 As New System.Data.OleDb.OleDbCommand
        CMD1.CommandType = System.Data.CommandType.Text   'command syntax
        Dim OLEDBREADER As OleDb.OleDbDataReader
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim fBrowse As New OpenFileDialog
        With fBrowse
            .Filter = "Excel files(*.xlsx)|*.xlsx|All files (*.*)|*.*"
            .FilterIndex = 1
            .Title = "Import data from Excel file"
        End With
        If fBrowse.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim fname As String
            fname = fBrowse.FileName
            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & fname & " '; " & "Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)

            'A CODE SHLD EXTRACT THE DATE FROM THE EXCEL SCHEDULE AND CHECK IF IT ALREADY EXISTS ON THE APPs RECORDS
            '2nd level check from file to be uploaded
            CMD1.CommandText = String.Format("select [SETTLEMENT_DATE] from [Sheet1$]") 'selects date from excel file
            CMD1.Connection = MyConnection
            MyConnection.Open() 'opens the connection 
            OLEDBREADER = CMD1.ExecuteReader
            If OLEDBREADER.HasRows Then
                While (OLEDBREADER.Read)
                    a2 = OLEDBREADER.GetValue(OLEDBREADER.GetOrdinal("SETTLEMENT_DATE")).ToString
                    Check4date()
                    If a2 = a3 Then
                        GoTo 2
                    Else
                        GoTo 1
                    End If
                End While
            End If


1:          MyCommand.TableMappings.Add("Table", "Test")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            MyConnection.Close()
            For Each Drr As DataRow In DtSet.Tables(0).Rows
                Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=./;Initial Catalog=INTERSWITCH;Integrated Security=True;")
                Dim CMD1RP As New System.Data.SqlClient.SqlCommand
                SQLCONNECTION1RP.Open()
                CMD1RP.CommandText = String.Format("INSERT INTO TRANSFER2014 (STATUS, ACCOUNT, ACCT_TYPE,AMT,FEE,SETTLEMENT_DATE) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')", Drr(0).ToString, Drr(1).ToString, Drr(2).ToString, Drr(3).ToString, Drr(4).ToString, Drr(5).ToString)
                CMD1RP.Connection = SQLCONNECTION1RP
                CMD1RP.ExecuteNonQuery()
                SQLCONNECTION1RP.Close()
            Next
            MsgBox("Successfully Saved")
            Button4.Visible = True
            Button3.Visible = False
            Button2.Visible = False
            ReportViewer1.Visible = False
            'DateTimePicker1.Visible = False
            Exit Sub
        Else
            MsgBox("No Excel File was uploaded")
            Button4.Visible = False
            Button3.Visible = False
            Button2.Visible = False
            ReportViewer1.Visible = False
            'DateTimePicker1.Visible = False
            Exit Sub
        End If





2:      MsgBox("A Computation file with date" & "  " & a3 & "  " & "has already been uploaded")
        Exit Sub


    End Sub


    Private Sub Check4date()

        'This code prevents the excel of same date from being uploaded twice

        'SQL
        Dim SQLCONNECTION6 As New System.Data.SqlClient.SqlConnection("Data Source=./;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD6 As New System.Data.SqlClient.SqlCommand
        CMD6.CommandType = System.Data.CommandType.Text   'command syntax
        Dim SQLREADER As SqlDataReader

        CMD6.CommandText = String.Format("select [SETTLEMENT_DATE] from [TRANSFER2014] where [SETTLEMENT_DATE] ='{0}'", a2) 'selects all account numbers in the table
        CMD6.Connection = SQLCONNECTION6
        SQLCONNECTION6.Open() 'opens the connection 
        SQLREADER = CMD6.ExecuteReader
        If SQLREADER.HasRows = True Then
            While (SQLREADER.Read)
                a3 = SQLREADER.GetValue(SQLREADER.GetOrdinal("SETTLEMENT_DATE")).ToString
            End While
        Else
            GoTo 1 'since a3 has no date value
        End If

1:      SQLCONNECTION6.Close()

    End Sub








    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Button4.Visible = False
        Button3.Visible = False
        Button2.Visible = False
        ReportViewer1.Visible = False
        Prevent_Duplicates()

    End Sub

    Private Sub Report_Viewer()


        Me.INTERSWITCHDataSet.Clear()
        Me.INTERSWITCHDataSet.SETTLEMENT.Clear()
        Me.TRANSFER2014TableAdapter.ClearBeforeFill = True



        Dim dataSet1 As DataSet = New DataSet("DataSet1")
        dataSet1 = TRANSFER2014BindingSource.DataSource
        dataSet1.EnforceConstraints = False


        Dim DT As Date



        DT = DateTimePicker1.Value.Date

        Me.TRANSFER2014TableAdapter.FillBy(Me.INTERSWITCHDataSet.TRANSFER2014, DT)
        Me.ReportViewer1.RefreshReport()
        ReportViewer1.Visible = True






    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Form1.Show()
        Me.Hide()

    End Sub

    
    
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Button1.Visible = True
        Button3.Visible = True
        Button2.Visible = True
        'DateTimePicker1.Visible = True


        Report_Viewer()

        MsgBox("Click Next to commence Settlement,If the upload results shown do not add up to the computation file you may use the delete button and re-upload correctly")


    End Sub

    Private Sub Second_Delete()

        'SQL  
        Dim SQLCONNECTION2D As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2D As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2D As System.Data.SqlClient.SqlDataReader
        CMD2D.CommandType = System.Data.CommandType.Text   'command syntax

        DV1 = DateTimePicker1.Value.Date

        CMD2D.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}'", DV1)

        SQLCONNECTION2D.Open()
        CMD2D.Connection = SQLCONNECTION2D
        SQLREADER2D = CMD2D.ExecuteReader



        If SQLREADER2D.HasRows Then
            'CHECK IF DATA WITH DATE SPECIFIED EXISTS
            First_Delete()
            SQLCONNECTION2D.Close()
        Else
            MsgBox("Automation File for Settlement Date" & "  " & DV1 & "  " & "could no longer be found it may have been deleted or upload is pending")
            GoTo 1
        End If


1:      SQLCONNECTION2D.Close()

    End Sub


    Private Sub First_Delete()

        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDB1 As New System.Data.SqlClient.SqlCommand

       

        'DELETE COMPUTATION FILE
        SQLCONNECTION3.Open()
        CMDB1.CommandText = String.Format("DELETE FROM TRANSFER2014 WHERE SETTLEMENT_DATE = '{0}'", dv1)
        CMDB1.Connection = SQLCONNECTION3
        CMDB1.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        Button3.Visible = False
        Button4.Visible = False
        'DateTimePicker1.Visible = False
        ReportViewer1.Visible = False
        Button1.Visible = True

        MsgBox("Computation File for Settlement Date" & "  " & DV1 & "  " & "Succesfully Deleted")

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Second_Delete()

    End Sub



End Class