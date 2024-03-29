Public Class Form2
    Public DV1 As Date
    Public ISW_PYBL As String
    Public OAB As String
    Public ECT As String
    Public MC_ATM As Double
    Public MC_WEB As String
    Public MC_POS As String
    Dim ISW_POS As String
    Public QT1 As String  '237
    Public QT2 As String  '238
    Public CSH_CRD_ACC As String
    Public MC_POS_AMT As Double
    Public MC_TRXN_REF As String
    Public MC_TRXN_DATE As String

    Public CC_AMT As Double         'FCMB CASHCARD
    Public CC_TRXN_REF As String
    Public CC_TRXN_DATE As String

    Public FCC_AMT As Double       'FINBANK CASHCARD
    Public FCC_TRXN_REF As String
    Public FCC_TRXN_DATE As String

    Public FC_AMT As Double           'FCMB FLASHWALLET
    Public FC_TRXN_REF As String
    Public FC_TRXN_DATE As String

    Public FFW_AMT As Double           'FINBANK FLASHWALLET
    Public FFW_TRXN_REF As String
    Public FFW_TRXN_DATE As String


    Public EC_AMT As Double            'FCMB CARDS ON FCMB POS TERMINAL
    Public EC_TRXN_REF As String
    Public EC_TRXN_DATE As String


    Public FB_EC_AMT As Double           'FINBANK CARDS ON FCMB POS TERMINAL
    Public FB_EC_TRXN_REF As String
    Public FB_EC_TRXN_DATE As String


    Public PP_AMT As Double            'FCMB POS PAYABLE
    Public PP_TRXN_REF As String
    Public PP_TRXN_DATE As String


    Public FPP_AMT As Double          'FINBANK POS PAYABLE
    Public FPP_TRXN_REF As String
    Public FPP_TRXN_DATE As String


    Public MW_AMT As Double
    Public V_IP As Double
    Public QT_VIP As Double

    Public QT_OA1 As Double
    Public QT_OA2 As Double


    Public QT_OA2_FB As Double
    Public QT_VIP_FB As Double

    Public IB_AMT As Double
    Public N96_AMT As Double
    Public HMFB As Double
    Public FB_INC As Double


    Public VRV_IP_FB As Double




    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form1.Show()
    End Sub




    Private Sub MasterCard_pos()



        'SQL 
        Dim SQLCONNECTION1MC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1MC As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1MC As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String

        '""

        CARD_TYPE = "MASTERCARD"
        CHANNEL = "POS"
        DV1 = DateTimePicker1.Value 'NO NEED TO CALL THIS PARAMETER AGAIN APPWIDE
        CMD1MC.CommandText = String.Format("SELECT * FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [CHANNEL]= '{1}' AND [SETTLE DATE] = '{2}'", CARD_TYPE, CHANNEL, DV1)

        SQLCONNECTION1MC.Open()
        CMD1MC.Connection = SQLCONNECTION1MC
        SQLREADER1MC = CMD1MC.ExecuteReader


        If SQLREADER1MC.HasRows Then
            GoTo 1
        Else
            MsgBox("There are no Mastercard recharge transactions done via POS Terminal", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1MC.Close()
            Exit Sub
        End If
1:      While (SQLREADER1MC.Read)
            MC_POS_AMT = Convert.ToDouble(SQLREADER1MC.GetValue(SQLREADER1MC.GetOrdinal("AMOUNT")))
            MC_TRXN_DATE = SQLREADER1MC.GetValue(SQLREADER1MC.GetOrdinal("TRXN DATE")).ToString()
            MC_TRXN_REF = Convert.ToString(SQLREADER1MC.GetValue(SQLREADER1MC.GetOrdinal("TRXN REF")))
            INSERT_MC_POS()
        End While
        MsgBox("Mastercard recharge transactions done via POS Terminal Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTION1MC.Close()




    End Sub

    Private Sub INSERT_MC_POS()
        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "0000079997"
        V2 = "D"
        'V3 = fee
        V4 = MC_TRXN_REF & " " & "RCHRG" & " " & Convert.ToString(MC_TRXN_DATE)
        V5 = "999"
        'V6 = "D"
        V7 = MC_POS_AMT



        'INSERT MC POS PAYABLE RECHARGE
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()


    End Sub


    Private Sub Cashcard_FCMB()
        'SQL 
        Dim SQLCONNECTION1CC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1CC As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1CC As System.Data.SqlClient.SqlDataReader

        Dim BIN As String
        'Dim CHANNEL As String

        '""

        BIN = "506100"
        'CHANNEL = "POS"
        DV1 = DateTimePicker1.Value
        CMD1CC.CommandText = String.Format("SELECT * FROM [RECHARGE_FCMB] WHERE [BIN] = '{0}' AND [SETTLE DATE] = '{1}'", BIN, DV1)

        SQLCONNECTION1CC.Open()
        CMD1CC.Connection = SQLCONNECTION1CC
        SQLREADER1CC = CMD1CC.ExecuteReader


        If SQLREADER1CC.HasRows Then
            GoTo 1  'THIS IS NECESSARY TO AVOID ERRORS WHEN USING A SUM QUERY, BUT WAS ADOPTED APPWIDE
        Else
            MsgBox("There are no Cashcard recharge transactions", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1CC.Close()
            Exit Sub
        End If
1:      While (SQLREADER1CC.Read)
            CC_AMT = Convert.ToDouble(SQLREADER1CC.GetValue(SQLREADER1CC.GetOrdinal("AMOUNT")))
            CC_TRXN_DATE = SQLREADER1CC.GetValue(SQLREADER1CC.GetOrdinal("TRXN DATE")).ToString()
            CC_TRXN_REF = Convert.ToString(SQLREADER1CC.GetValue(SQLREADER1CC.GetOrdinal("PAN")))
            INSERT_CC()
        End While
        MsgBox("Cashcard recharge transactions Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTION1CC.Close()








    End Sub



    Private Sub INSERT_CC()


        'SQL
        Dim SQLCONNECTION2CC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI99924430138"
        V2 = "D"
        'V3 = fee
        V4 = CC_TRXN_REF 'PAN
        V5 = "999"
        'V6 = "D"
        V7 = CC_AMT



        'INSERT CASHCARD RECHARGE SETTLEMENT
        SQLCONNECTION2CC.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2CC
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2CC.Close()





    End Sub


    Private Sub Flash_wallet()

        'SQL 
        Dim SQLCONNECTION1FC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1FC As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1FC As System.Data.SqlClient.SqlDataReader
        Dim BIN As String
        'Dim CHANNEL As String

        '""

        BIN = "506138"
        'CHANNEL = "POS"
        DV1 = DateTimePicker1.Value
        CMD1FC.CommandText = String.Format("SELECT SUM(AMOUNT) AS FC_AMT FROM [RECHARGE_FCMB] WHERE [BIN] = '{0}' AND [SETTLE DATE] = '{1}'", BIN, DV1)

        SQLCONNECTION1FC.Open()
        CMD1FC.Connection = SQLCONNECTION1FC
        SQLREADER1FC = CMD1FC.ExecuteReader

        Try
            If SQLREADER1FC.HasRows Then
                GoTo 1  'THIS IS NECESSARY TO AVOID ERRORS WHEN USING A SUM QUERY, BUT WAS ADOPTED APPWIDE
            Else
                MsgBox("There are no flash wallet recharge transactions", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION1FC.Close()
                Exit Sub
            End If
1:          While (SQLREADER1FC.Read)
                FC_AMT = Convert.ToDouble(SQLREADER1FC.GetValue(SQLREADER1FC.GetOrdinal("FC_AMT")))
                'FC_TRXN_DATE = SQLREADER1FC.GetValue(SQLREADER1FC.GetOrdinal("TRXN DATE")).ToString()
                'FC_TRXN_REF = Convert.ToString(SQLREADER1FC.GetValue(SQLREADER1FC.GetOrdinal("PAN")))
                Insert_FC()
            End While
            MsgBox("Flash wallet recharge transactions Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1FC.Close()
        Catch EX As Exception
            FC_AMT = 0
        End Try




    End Sub

    Private Sub Insert_FC()

        'flashwallet 506138

        'SQL
        Dim SQLCONNECTION2CC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.Date
        V5 = "999"
        'V6 = "D"
        V7 = FC_AMT



        'INSERT CASHCARD RECHARGE SETTLEMENT
        SQLCONNECTION2CC.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2CC
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2CC.Close()

    End Sub


    Private Sub Etcc_FCMB()



        'SQL 
        Dim SQLCONNECTION1ET As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1ET As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1ET As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim ISSUER As String

        '""

        CARD_TYPE = "VERVE"
        CHANNEL = "POS"
        TRM_OWNER = "First City Monument Bank"
        ISSUER = "First City Monumental Bank"
        DV1 = DateTimePicker1.Value
        CMD1ET.CommandText = String.Format("SELECT * FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [CHANNEL]= '{1}' AND [TERMINAL OWNER]='{2}' AND [ISSUER]= '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, CHANNEL, TRM_OWNER, ISSUER, DV1)

        SQLCONNECTION1ET.Open()
        CMD1ET.Connection = SQLCONNECTION1ET
        SQLREADER1ET = CMD1ET.ExecuteReader


        If SQLREADER1ET.HasRows Then
            GoTo 1
        Else
            MsgBox("There are no verve recharge transactions done via POS Terminal", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1ET.Close()
            Exit Sub
        End If
1:      While (SQLREADER1ET.Read)
            EC_AMT = Convert.ToDouble(SQLREADER1ET.GetValue(SQLREADER1ET.GetOrdinal("AMOUNT")))
            EC_TRXN_DATE = SQLREADER1ET.GetValue(SQLREADER1ET.GetOrdinal("TRXN DATE")).ToString()
            EC_TRXN_REF = Convert.ToString(SQLREADER1ET.GetValue(SQLREADER1ET.GetOrdinal("TRXN REF")))
            INSERT_ETCC_FCMB()
        End While
        MsgBox("FCMB Verve recharge transactions done via FCMB POS Terminal Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTION1ET.Close()










    End Sub

    Private Sub INSERT_ETCC_FCMB()


        'SQL
        Dim SQLCONNECTION2ET As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "9999999602"
        V2 = "D"
        'V3 = fee
        V4 = EC_TRXN_REF & " " & EC_TRXN_DATE 'REF
        V5 = "999"
        'V6 = "D"
        V7 = EC_AMT



        'INSERT ETCC TRANSACTIONS   (FCMB VERVE ON FCMB POS TERMINALS)
        SQLCONNECTION2ET.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2ET
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2ET.Close()



    End Sub


    Private Sub ISW_POS_PYBL()

        'SQL 
        Dim SQLCONNECTION1PP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1PP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1PP As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim ISSUER As String

        '""

        CARD_TYPE = "VERVE"
        CHANNEL = "POS"
        TRM_OWNER = "First City Monument Bank"
        ISSUER = "First City Monumental Bank"
        DV1 = DateTimePicker1.Value
        CMD1PP.CommandText = String.Format("SELECT * FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [CHANNEL]= '{1}' AND [TERMINAL OWNER] <>'{2}' AND [ISSUER] = '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, CHANNEL, TRM_OWNER, ISSUER, DV1)

        SQLCONNECTION1PP.Open()
        CMD1PP.Connection = SQLCONNECTION1PP
        SQLREADER1PP = CMD1PP.ExecuteReader


        If SQLREADER1PP.HasRows Then
            GoTo 1
        Else
            MsgBox("There are no FCMB verve recharge transactions done via NON FCMB POS Terminal", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1PP.Close()
            Exit Sub
        End If
1:      While (SQLREADER1PP.Read)
            PP_AMT = Convert.ToDouble(SQLREADER1PP.GetValue(SQLREADER1PP.GetOrdinal("AMOUNT")))
            PP_TRXN_DATE = SQLREADER1PP.GetValue(SQLREADER1PP.GetOrdinal("TRXN DATE")).ToString()
            PP_TRXN_REF = Convert.ToString(SQLREADER1PP.GetValue(SQLREADER1PP.GetOrdinal("TRXN REF")))
            PP_INSERT()
        End While
        MsgBox("FCMB Verve recharge transactions done via NON FCMB POS Terminal Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTION1PP.Close()



        'Select * from RECHARGE_FCMB  where [SETTLE DATE]   = '2014-06-26'
        'AND [TERMINAL OWNER] <> 'First City Monument Bank'
        'AND [CARD_TYPE] ='VERVE' AND [CHANNEL]='POS'AND [ISSUER] ='First City Monumental Bank'
        'AND [SETTLE DATE] ='2014-06-26'


    End Sub


    Private Sub PP_INSERT()

        'PP IS ISW POS PAYABLE

        'SQL
        Dim SQLCONNECTION2PP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "0000039999"
        V2 = "D"
        'V3 = fee
        V4 = PP_TRXN_REF & " " & "RCHG" & PP_TRXN_DATE 'REF
        V5 = "999"
        'V6 = "D"
        V7 = PP_AMT



        'INSERT POS PAYABLE RECHARGE
        SQLCONNECTION2PP.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2PP
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2PP.Close()




    End Sub


    Private Sub MC_Web_Recharge()


        'SQL 
        Dim SQLCONNECTION1MW As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1MW As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1MW As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        'Dim CHANNEL As String
        Dim TRM_OWNER As String
        'Dim ISSUER As String

        '""

        CARD_TYPE = "MASTERCARD"
        TRM_OWNER = "QuickTeller Website"
        DV1 = DateTimePicker1.Value
        CMD1MW.CommandText = String.Format("SELECT SUM(AMOUNT) AS MW_T FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] = '{1}'  AND [SETTLE DATE] = '{2}'", CARD_TYPE, TRM_OWNER, DV1)

        SQLCONNECTION1MW.Open()
        CMD1MW.Connection = SQLCONNECTION1MW
        SQLREADER1MW = CMD1MW.ExecuteReader


        Try
            If SQLREADER1MW.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no MasterCard Web recharge transactions done via QuickTeller", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION1MW.Close()
                Exit Sub
            End If
1:          While (SQLREADER1MW.Read)
                MW_AMT = Convert.ToDouble(SQLREADER1MW.GetValue(SQLREADER1MW.GetOrdinal("MW_T")))
                MW_Insert()
            End While
            MsgBox("MasterCard Web recharge transactions Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1MW.Close()
        Catch ex As Exception
            MW_AMT = 0
        End Try




        'Select * from RECHARGE_FCMB  where [SETTLE DATE]   = '2014-06-26'
        'AND [TERMINAL OWNER] = 'QuickTeller Website'
        'AND [CARD_TYPE] ='MASTERCARD' 
        'AND [SETTLE DATE] ='2014-06-26'


    End Sub

    Private Sub MW_Insert()




        'MW IS MCWEB RECHARGE

        'SQL
        Dim SQLCONNECTION2MW As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI99924430253"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        'V6 = "D"
        V7 = MW_AMT



        'INSERT MW
        SQLCONNECTION2MW.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2MW
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2MW.Close()







    End Sub


    Private Sub MC_ATM_RCH()

        'SQL 
        Dim SQLCONNECTION1MW As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1MW As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1MW As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        'Dim ISSUER As String

        '""
        CHANNEL = "POS"
        CARD_TYPE = "MASTERCARD"
        TRM_OWNER = "QuickTeller Website"
        DV1 = DateTimePicker1.Value
        CMD1MW.CommandText = String.Format("SELECT SUM(AMOUNT) AS MC_ATM FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] <> '{1}'  AND [CHANNEL] <> '{2}' AND [SETTLE DATE] = '{3}'", CARD_TYPE, TRM_OWNER, CHANNEL, DV1)

        SQLCONNECTION1MW.Open()
        CMD1MW.Connection = SQLCONNECTION1MW
        SQLREADER1MW = CMD1MW.ExecuteReader

        Try
            If SQLREADER1MW.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no MasterCard ATM recharge transactions done via QuickTeller", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION1MW.Close()
                Exit Sub
            End If
1:          While (SQLREADER1MW.Read)
                MC_ATM = Convert.ToDouble(SQLREADER1MW.GetValue(SQLREADER1MW.GetOrdinal("MC_ATM")))
                MC_ATM_INSERT()
            End While
            MsgBox("MasterCard ATM recharge transactions Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1MW.Close()
        Catch ex As Exception
            MC_ATM = 0
        End Try




        'Select sum(AMOUNT) as MC_ATM from RECHARGE_FCMB  where [SETTLE DATE]   = '2014-06-26'
        'AND [TERMINAL OWNER] <> 'QuickTeller Website'
        'AND [CARD_TYPE] ='MASTERCARD' 
        'AND [CHANNEL]  <>'POS'
        'AND [SETTLE DATE] ='2014-06-26'




    End Sub

    Private Sub MC_ATM_INSERT()



        'MASTERCARD RECHARGE ON ATM

        'SQL
        Dim SQLCONNECTION2MW As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI99924430254"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        'V6 = "D"
        V7 = MC_ATM



        'INSERT MW
        SQLCONNECTION2MW.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2MW
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2MW.Close()


    End Sub

    Private Sub Verve_IP()


        'SQL 
        Dim SQLCONNECTION1IP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1IP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1IP As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim TRMNL_ID As String
        'Dim ISSUER As String

        '""
        CHANNEL = "POS"
        CARD_TYPE = "VERVE"
        TRM_OWNER = "QuickTeller Website"
        TRMNL_ID = "3FMI0001"
        DV1 = DateTimePicker1.Value
        CMD1IP.CommandText = String.Format("SELECT SUM(AMOUNT) AS VRV_IP FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] <> '{1}'  AND [CHANNEL] <> '{2}' AND [TERMINAL ID] <> '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, TRM_OWNER, CHANNEL, TRMNL_ID, DV1)

        SQLCONNECTION1IP.Open()
        CMD1IP.Connection = SQLCONNECTION1IP
        SQLREADER1IP = CMD1IP.ExecuteReader

        Try
            If SQLREADER1IP.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no MasterCard ATM recharge transactions done via QuickTeller", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION1IP.Close()
                Exit Sub
            End If
1:          While (SQLREADER1IP.Read)
                V_IP = Convert.ToDouble(SQLREADER1IP.GetValue(SQLREADER1IP.GetOrdinal("VRV_IP")))
                Verve_Isw_pybl_insert()
            End While
            MsgBox("ROU ATM recharge transactions Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1IP.Close()
        Catch ex As Exception
            V_IP = 0
        End Try



        'Select sum(AMOUNT) as VRV_IP from RECHARGE_FCMB  where [SETTLE DATE]   = '2014-06-26'
        'AND [TERMINAL OWNER] <> 'QuickTeller Website'
        'AND [CARD_TYPE] ='VERVE' 
        'AND [TERMINAL ID] <> '3FMI0001'
        'AND [CHANNEL]  <>'POS'


    End Sub


    Private Sub Verve_Isw_pybl_insert()



        'ISW PAYABLE ROU RECHARGE ON ATM

        'SQL
        Dim SQLCONNECTION2IP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        'V6 = "D"
        V7 = V_IP



        'INSERT MW
        SQLCONNECTION2IP.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2IP
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2IP.Close()

    End Sub


    Private Sub QT_Verve_IP()


        'SQL 
        Dim SQLCONNECTION3IP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3IP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3IP As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        'Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim TRMNL_ID As String
        'Dim ISSUER As String

        '""
        'CHANNEL = "POS"
        CARD_TYPE = "VERVE"
        TRM_OWNER = "QuickTeller Website"
        TRMNL_ID = "4QTL0001"
        DV1 = DateTimePicker1.Value
        CMD3IP.CommandText = String.Format("SELECT SUM(AMOUNT) AS QT_VRV_IP FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] = '{1}'  AND [TERMINAL ID] = '{2}' AND [SETTLE DATE] = '{3}'", CARD_TYPE, TRM_OWNER, TRMNL_ID, DV1)

        SQLCONNECTION3IP.Open()
        CMD3IP.Connection = SQLCONNECTION3IP
        SQLREADER3IP = CMD3IP.ExecuteReader

        Try
            If SQLREADER3IP.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no MasterCard ATM recharge transactions done via QuickTeller", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION3IP.Close()
                Exit Sub
            End If
1:          While (SQLREADER3IP.Read)
                QT_VIP = Convert.ToDouble(SQLREADER3IP.GetValue(SQLREADER3IP.GetOrdinal("QT_VRV_IP")))
                INSERT_QT_VIP()
            End While
            MsgBox("ROU ATM recharge transactions Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION3IP.Close()
        Catch ex As Exception
            QT_VIP = 0
        End Try

        'Select sum(AMOUNT) as QT_IP from RECHARGE_FCMB  where [SETTLE DATE]   = '2014-06-26'
        'AND [CARD_TYPE] ='VERVE'
        'AND [TERMINAL ID] ='4QTL0001'


    End Sub

    Private Sub INSERT_QT_VIP()


        'ISW PAYABLE ROU RECHARGE ON QTELLER

        'SQL
        Dim SQLCONNECTION4IP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        'V6 = "D"
        V7 = QT_VIP



        'INSERT MW
        SQLCONNECTION4IP.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION4IP
        CMD3.ExecuteNonQuery()
        SQLCONNECTION4IP.Close()



    End Sub


    Private Sub OAB_3FMI0001()
        'SQL 
        Dim SQLCONNECTION1OA As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1OA As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1OA As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        'Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim TRMNL_ID As String
        Dim BIN As String
        'Dim ISSUER As String

        '""
        'CHANNEL = "POS"
        CARD_TYPE = "FLOAT"
        TRM_OWNER = "QuickTeller Website"
        TRMNL_ID = "3FMI0001"
        BIN = "628051"


        DV1 = DateTimePicker1.Value
        CMD1OA.CommandText = String.Format("SELECT SUM(AMOUNT) AS QT_OAB FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] = '{1}'  AND [TERMINAL ID] = '{2}' AND [BIN]= '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, TRM_OWNER, TRMNL_ID, BIN, DV1)

        SQLCONNECTION1OA.Open()
        CMD1OA.Connection = SQLCONNECTION1OA
        SQLREADER1OA = CMD1OA.ExecuteReader

        Try
            If SQLREADER1OA.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no recharge transactions done via QuickTeller on Terminal 3FMI0001", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION1OA.Close()
                Exit Sub
            End If
1:          While (SQLREADER1OA.Read)
                QT_OA1 = Convert.ToDouble(SQLREADER1OA.GetValue(SQLREADER1OA.GetOrdinal("QT_OAB")))
                INSERT_QT_OAB1()
            End While
            MsgBox("VERVE recharge transactions ON QUICKTELLER TERMINAL 3FMI0001 Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1OA.Close()
        Catch ex As Exception
            QT_OA1 = 0
        End Try


        'Select sum(AMOUNT) as QT_IP from RECHARGE_FCMB  where [SETTLE DATE]   = '2014-06-26'
        'AND [CARD_TYPE] ='VERVE'
        'AND [TERMINAL ID] ='4QTL0001'


    End Sub

    Private Sub INSERT_QT_OAB1()

        'ISW PAYABLE ROU RECHARGE ON QTELLER

        'SQL
        Dim SQLCONNECTION2OA As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand
        Dim CMD4 As New System.Data.SqlClient.SqlCommand
        Dim CMD5 As New System.Data.SqlClient.SqlCommand

        Dim VSWP1 As String
        Dim VSWP2 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        Dim V6 As String
        Dim V7 As Double
        Dim CM As String

        VSWP1 = "NGNLI99924430238"
        VSWP2 = "NGNLI99924430237"
        V1 = "NGNLI99924430087"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        V6 = "C"
        V7 = QT_OA1
        CM = "RECHARGE 3FMI0001"


        'INSERT QT_3FMI0001
        SQLCONNECTION2OA.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2OA
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2OA.Close()



        'RECHARGE SWEEP ' DR 238
        SQLCONNECTION2OA.Open()
        CMD4.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", VSWP1, V2, V7, V4, V5, CM, DV1)
        CMD4.Connection = SQLCONNECTION2OA
        CMD4.ExecuteNonQuery()
        SQLCONNECTION2OA.Close()

        'RECHARGE SWEEP ' CR 237
        SQLCONNECTION2OA.Open()
        CMD4.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", VSWP2, V6, V7, V4, V5, CM, DV1)
        CMD4.Connection = SQLCONNECTION2OA
        CMD4.ExecuteNonQuery()
        SQLCONNECTION2OA.Close()


        'THIS SWEEP IS BECAUSE OF THE FFN EXPLANATION BELOW








    End Sub

    Private Sub OAB_3BOL0001()

        'SQL 
        Dim SQLCONNECTION3OA As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3OA As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3OA As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        'Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim TRMNL_ID As String
        Dim BIN As String
        'Dim ISSUER As String

        '""
        'CHANNEL = "POS"
        CARD_TYPE = "VERVE"
        TRM_OWNER = "QuickTeller Website"
        TRMNL_ID = "3BOL0001"
        BIN = "506108"


        DV1 = DateTimePicker1.Value
        CMD3OA.CommandText = String.Format("SELECT SUM(AMOUNT) AS QT_OAB FROM [RECHARGE_FCMB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] = '{1}'  AND [TERMINAL ID] = '{2}' AND [BIN]= '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, TRM_OWNER, TRMNL_ID, BIN, DV1)

        SQLCONNECTION3OA.Open()
        CMD3OA.Connection = SQLCONNECTION3OA
        SQLREADER3OA = CMD3OA.ExecuteReader

        Try
            If SQLREADER3OA.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no recharge transactions done via QuickTeller on Terminal 3BOL0001", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION3OA.Close()
                Exit Sub
            End If
1:          While (SQLREADER3OA.Read)
                QT_OA2 = Convert.ToDouble(SQLREADER3OA.GetValue(SQLREADER3OA.GetOrdinal("QT_OAB")))
                INSERT_QT_OAB2()
            End While
            MsgBox("VERVE recharge transactions ON QUICKTELLER TERMINAL 3BOL0001 Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION3OA.Close()
        Catch ex As Exception

            QT_OA2 = 0
        End Try

    End Sub

    Private Sub INSERT_QT_OAB2()



        'ISW PAYABLE ROU RECHARGE ON QTELLER

        'SQL
        Dim SQLCONNECTION4OA As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand
       

        
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        Dim V6 As String
        Dim V7 As Double
        Dim CM As String

        
        V1 = "NGNLI99924430087"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        V6 = "C"
        V7 = QT_OA2
        CM = "RECHARGE 3BOL0001"


        'INSERT QT_3FMI0001
        SQLCONNECTION4OA.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION4OA
        CMD3.ExecuteNonQuery()
        SQLCONNECTION4OA.Close()


    End Sub


    'FINBANK'''''''''''''''''''''''''''''''''''''

    Private Sub OAB_3BOL0001_FINBANK()



        'SQL 
        Dim SQLCONNECTION1FB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1FB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1FB As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        'Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim TRMNL_ID As String
        'Dim BIN As String
        'Dim ISSUER As String

        '""
        'CHANNEL = "POS"
        CARD_TYPE = "VERVE"
        TRM_OWNER = "QuickTeller Website"
        TRMNL_ID = "3BOL0001"
        'BIN = "506114"


        DV1 = DateTimePicker1.Value
        CMD1FB.CommandText = String.Format("SELECT SUM(AMOUNT) AS QT_OAB FROM [RECHARGE_FB] WHERE [TERMINAL OWNER] = '{0}' AND [CARD_TYPE]  = '{1}'  AND [TERMINAL ID] = '{2}' AND [SETTLE DATE] = '{3}'", TRM_OWNER, CARD_TYPE, TRMNL_ID, DV1)

        SQLCONNECTION1FB.Open()
        CMD1FB.Connection = SQLCONNECTION1FB
        SQLREADER1FB = CMD1FB.ExecuteReader

        Try
            If SQLREADER1FB.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no recharge transactions done via QuickTeller on Terminal 3BOL0001 FOR FINBANK", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION1FB.Close()
                Exit Sub
            End If
1:          While (SQLREADER1FB.Read)
                QT_OA2_FB = Convert.ToDouble(SQLREADER1FB.GetValue(SQLREADER1FB.GetOrdinal("QT_OAB")))
                INSERT_QT_OAB_FINBANK()
            End While
            MsgBox("VERVE recharge transactions ON QUICKTELLER TERMINAL 3BOL0001 FOR FINBANK Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1FB.Close()
        Catch ex As System.Exception
            QT_OA2_FB = 0
            'INSERT_QT_OAB_FINBANK()
        End Try


    End Sub



    Private Sub INSERT_QT_OAB_FINBANK()



        'ISW PAYABLE ROU RECHARGE ON QTELLER

        'SQL
        Dim SQLCONNECTION2FB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand



        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        Dim V6 As String
        Dim V7 As Double
        Dim CM As String


        V1 = "NGNLI99924430087"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        V6 = "C"
        V7 = QT_OA2_FB
        CM = "RECHARGE 3BOL0001"


        'INSERT QT_3FMI0001
        SQLCONNECTION2FB.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2FB
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2FB.Close()


    End Sub


    Private Sub QT_VERVE_IP_FINBANK()


        'SQL 
        Dim SQLCONNECTION3FB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3FB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3FB As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        'Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim TRMNL_ID As String
        'Dim ISSUER As String

        '""
        'CHANNEL = "POS"
        CARD_TYPE = "VERVE"
        TRM_OWNER = "QuickTeller Website"
        TRMNL_ID = "4QTL0001"
        DV1 = DateTimePicker1.Value.Date
        CMD3FB.CommandText = String.Format("SELECT SUM(AMOUNT) AS QT_VRV_IP FROM [RECHARGE_FB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] = '{1}'  AND [TERMINAL ID] = '{2}' AND [SETTLE DATE] = '{3}'", CARD_TYPE, TRM_OWNER, TRMNL_ID, DV1)

        SQLCONNECTION3FB.Open()
        CMD3FB.Connection = SQLCONNECTION3FB
        SQLREADER3FB = CMD3FB.ExecuteReader

        Try
            If SQLREADER3FB.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no ATM recharge transactions done via QuickTeller", vbOKOnly, "Interswitch Automation")
                SQLCONNECTION3FB.Close()
                Exit Sub
            End If
1:          While (SQLREADER3FB.Read)
                QT_VIP_FB = Convert.ToDouble(SQLREADER3FB.GetValue(SQLREADER3FB.GetOrdinal("QT_VRV_IP")))
                INSERT_QT_VERVE_FINBANK()
            End While
            MsgBox("QUICKTELLER recharge transactions FOR FINBANK Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION3FB.Close()
        Catch ex As System.Exception
            QT_VIP_FB = 0
            'INSERT_QT_VERVE_FINBANK()
        End Try




    End Sub

    Private Sub INSERT_QT_VERVE_FINBANK()


        'ISW PAYABLE ROU RECHARGE ON QTELLER

        'SQL
        Dim SQLCONNECTION4FB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        'V6 = "D"
        V7 = QT_VIP_FB



        'INSERT MW
        SQLCONNECTION4FB.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION4FB
        CMD3.ExecuteNonQuery()
        SQLCONNECTION4FB.Close()

    End Sub


    Private Sub CASHCARD_FINBANK()


        'SQL 
        Dim SQLCONNECTIONFBCC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDFBCC As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERFBCC As System.Data.SqlClient.SqlDataReader

        Dim BIN As String
        'Dim CHANNEL As String

        '""

        BIN = "506100"
        'CHANNEL = "POS"
        DV1 = DateTimePicker1.Value
        CMDFBCC.CommandText = String.Format("SELECT * FROM [RECHARGE_FB] WHERE [BIN] = '{0}' AND [SETTLE DATE] = '{1}'", BIN, DV1)

        SQLCONNECTIONFBCC.Open()
        CMDFBCC.Connection = SQLCONNECTIONFBCC
        SQLREADERFBCC = CMDFBCC.ExecuteReader


        If SQLREADERFBCC.HasRows Then
            GoTo 1  'THIS IS NECESSARY TO AVOID ERRORS WHEN USING A SUM QUERY, BUT WAS ADOPTED APPWIDE
        Else
            MsgBox("There are no Cashcard recharge transactions", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONFBCC.Close()
            Exit Sub
        End If
1:      While (SQLREADERFBCC.Read)
            FCC_AMT = Convert.ToDouble(SQLREADERFBCC.GetValue(SQLREADERFBCC.GetOrdinal("AMOUNT")))
            FCC_TRXN_DATE = SQLREADERFBCC.GetValue(SQLREADERFBCC.GetOrdinal("TRXN DATE")).ToString()
            FCC_TRXN_REF = Convert.ToString(SQLREADERFBCC.GetValue(SQLREADERFBCC.GetOrdinal("PAN")))
            INSERT_CC_FINBANK()
        End While
        MsgBox("Cashcard recharge transactions for Finbank Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTIONFBCC.Close()




    End Sub


    Private Sub INSERT_CC_FINBANK()


        'SQL
        Dim SQLCONNECTION2FBCC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI99924430138"
        V2 = "D"
        'V3 = fee
        V4 = FCC_TRXN_REF 'PAN
        V5 = "999"
        'V6 = "D"
        V7 = FCC_AMT



        'INSERT CASHCARD RECHARGE SETTLEMENT
        SQLCONNECTION2FBCC.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2FBCC
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2FBCC.Close()




    End Sub


    Private Sub FW_FINBANK()


        'SQL 
        Dim SQLCONNECTIONFINFC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDFINFC As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERFINFC As System.Data.SqlClient.SqlDataReader

        Dim BIN As String
        'Dim CHANNEL As String

        '""

        BIN = "506138"
        'CHANNEL = "POS"
        DV1 = DateTimePicker1.Value
        CMDFINFC.CommandText = String.Format("SELECT SUM(AMOUNT) AS FFW_AMT FROM [RECHARGE_FB] WHERE [BIN] = '{0}' AND [SETTLE DATE] = '{1}'", BIN, DV1)

        SQLCONNECTIONFINFC.Open()
        CMDFINFC.Connection = SQLCONNECTIONFINFC
        SQLREADERFINFC = CMDFINFC.ExecuteReader

        Try
            If SQLREADERFINFC.HasRows Then
                GoTo 1  'THIS IS NECESSARY TO AVOID ERRORS WHEN USING A SUM QUERY, BUT WAS ADOPTED APPWIDE
            Else
                MsgBox("There are no flash wallet recharge transactions", vbOKOnly, "Interswitch Automation")
                SQLCONNECTIONFINFC.Close()
                Exit Sub
            End If
1:          While (SQLREADERFINFC.Read)
                FFW_AMT = Convert.ToDouble(SQLREADERFINFC.GetValue(SQLREADERFINFC.GetOrdinal("FFW_AMT")))
                'FFW_TRXN_DATE = SQLREADERFINFC.GetValue(SQLREADERFINFC.GetOrdinal("TRXN DATE")).ToString()
                'FFW_TRXN_REF = Convert.ToString(SQLREADERFINFC.GetValue(SQLREADERFINFC.GetOrdinal("PAN")))
                INSERT_FW_FINBANK()
            End While
            MsgBox("Flash wallet recharge transactions Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONFINFC.Close()
        Catch EX As Exception
            FFW_AMT = 0
        End Try



    End Sub


    Private Sub INSERT_FW_FINBANK()



        'flashwallet 506138

        'SQL
        Dim SQLCONNECTION2CC As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.Date
        V5 = "999"
        'V6 = "D"
        V7 = FFW_AMT



        'INSERT CASHCARD RECHARGE SETTLEMENT
        SQLCONNECTION2CC.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2CC
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2CC.Close()





    End Sub


    Private Sub Etcc_FINBANK()

        'RCHRG ON FCMB/FIN POS TERMINAL

        'SQL 
        Dim SQLCONNECTION1ET As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1ET As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1ET As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim ISSUER As String

        '""

        CARD_TYPE = "VERVE"
        CHANNEL = "POS"
        TRM_OWNER = "First City Monument Bank"
        ISSUER = "First Inland Bank Plc"
        DV1 = DateTimePicker1.Value
        CMD1ET.CommandText = String.Format("SELECT * FROM [RECHARGE_FB] WHERE [CARD_TYPE] = '{0}' AND [CHANNEL]= '{1}' AND [TERMINAL OWNER]='{2}' AND [ISSUER]= '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, CHANNEL, TRM_OWNER, ISSUER, DV1)

        SQLCONNECTION1ET.Open()
        CMD1ET.Connection = SQLCONNECTION1ET
        SQLREADER1ET = CMD1ET.ExecuteReader


        If SQLREADER1ET.HasRows Then
            GoTo 1
        Else
            MsgBox("There are no verve recharge transactions done via POS Terminal", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1ET.Close()
            Exit Sub
        End If
1:      While (SQLREADER1ET.Read)
            FB_EC_AMT = Convert.ToDouble(SQLREADER1ET.GetValue(SQLREADER1ET.GetOrdinal("AMOUNT")))
            FB_EC_TRXN_DATE = SQLREADER1ET.GetValue(SQLREADER1ET.GetOrdinal("TRXN DATE")).ToString()
            FB_EC_TRXN_REF = Convert.ToString(SQLREADER1ET.GetValue(SQLREADER1ET.GetOrdinal("TRXN REF")))
            INSERT_ETCC_FINBANK()
        End While
        MsgBox("FCMB Verve recharge transactions done via FCMB POS Terminal Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTION1ET.Close()




    End Sub

    Private Sub INSERT_ETCC_FINBANK()


        'SQL
        Dim SQLCONNECTION2ET As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "9999999602"
        V2 = "D"
        'V3 = fee
        V4 = FB_EC_TRXN_REF & " " & "RCHRG" & " " & FB_EC_TRXN_DATE 'REF
        V5 = "999"
        'V6 = "D"
        V7 = FB_EC_AMT



        'INSERT ETCC TRANSACTIONS   (FCMB VERVE ON FCMB POS TERMINALS)
        SQLCONNECTION2ET.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2ET
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2ET.Close()


    End Sub


    'THERE MIGHT BE A NEED TO WRITE CODE FOR POS TERMINAL OWNED BY FINBANK HERE TOO



    Private Sub PP_FINBANK()

        'SQL 
        Dim SQLCONNECTION1PP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1PP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1PP As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        Dim ISSUER As String

        '""

        CARD_TYPE = "VERVE"
        CHANNEL = "POS"
        TRM_OWNER = "First City Monument Bank"
        ISSUER = "First Inland Bank Plc"
        DV1 = DateTimePicker1.Value
        CMD1PP.CommandText = String.Format("SELECT * FROM [RECHARGE_FB] WHERE [CARD_TYPE] = '{0}' AND [CHANNEL]= '{1}' AND [TERMINAL OWNER] <>'{2}' AND [ISSUER] = '{3}' AND [SETTLE DATE] = '{4}'", CARD_TYPE, CHANNEL, TRM_OWNER, ISSUER, DV1)

        SQLCONNECTION1PP.Open()
        CMD1PP.Connection = SQLCONNECTION1PP
        SQLREADER1PP = CMD1PP.ExecuteReader


        If SQLREADER1PP.HasRows Then
            GoTo 1
        Else
            MsgBox("There are no FCMB verve recharge transactions done via NON FCMB POS Terminal", vbOKOnly, "Interswitch Automation")
            SQLCONNECTION1PP.Close()
            Exit Sub
        End If
1:      While (SQLREADER1PP.Read)
            FPP_AMT = Convert.ToDouble(SQLREADER1PP.GetValue(SQLREADER1PP.GetOrdinal("AMOUNT")))
            FPP_TRXN_DATE = SQLREADER1PP.GetValue(SQLREADER1PP.GetOrdinal("TRXN DATE")).ToString()
            FPP_TRXN_REF = Convert.ToString(SQLREADER1PP.GetValue(SQLREADER1PP.GetOrdinal("TRXN REF")))
            INSERT_PP_FINBANK()
        End While
        MsgBox("FCMB Verve recharge transactions done via NON FCMB POS Terminal Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTION1PP.Close()

    End Sub


    Private Sub INSERT_PP_FINBANK()

        'PP IS ISW POS PAYABLE

        'SQL
        Dim SQLCONNECTION2PP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "0000039999"
        V2 = "D"
        'V3 = fee
        V4 = FPP_TRXN_REF & " " & "RCHRG" & FPP_TRXN_DATE 'REF
        V5 = "999"
        'V6 = "D"
        V7 = FPP_AMT



        'INSERT POS PAYABLE RECHARGE
        SQLCONNECTION2PP.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2PP
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2PP.Close()






    End Sub

    Private Sub VERVE_IP_FINBANK()

        'RCHRGE TRANSACTIONS HERE ARE DONE ON NON QUICKTELLER,WITH FINBANK VERVE CARDS ON FCMB AND NON FCMB ATM 

        'SQL 
        Dim SQLCONNECTIONVFB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDVFB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERVFB As System.Data.SqlClient.SqlDataReader

        Dim CARD_TYPE As String
        Dim CHANNEL As String
        Dim TRM_OWNER As String
        'Dim TRMNL_ID As String
        'Dim ISSUER As String

        '""
        CHANNEL = "POS"
        CARD_TYPE = "VERVE"
        TRM_OWNER = "QuickTeller Website"
        'TRMNL_ID = "3FMI0001"
        DV1 = DateTimePicker1.Value
        CMDVFB.CommandText = String.Format("SELECT SUM(AMOUNT) AS VRV_IP FROM [RECHARGE_FB] WHERE [CARD_TYPE] = '{0}' AND [TERMINAL OWNER] <> '{1}'  AND [CHANNEL] <> '{2}' AND [SETTLE DATE] = '{3}'", CARD_TYPE, TRM_OWNER, CHANNEL, DV1)


        SQLCONNECTIONVFB.Open()
        CMDVFB.Connection = SQLCONNECTIONVFB
        SQLREADERVFB = CMDVFB.ExecuteReader

        Try
            If SQLREADERVFB.HasRows Then
                GoTo 1
            Else
                MsgBox("There are no recharge transactions done via ATM FOR FINBANK", vbOKOnly, "Interswitch Automation")
                SQLCONNECTIONVFB.Close()
                Exit Sub
            End If
1:          While (SQLREADERVFB.Read)
                VRV_IP_FB = Convert.ToDouble(SQLREADERVFB.GetValue(SQLREADERVFB.GetOrdinal("VRV_IP")))
                INSERT_VRV_IP_FINBANK()
            End While
            MsgBox("VERVE recharge transactions ON ATM FOR FINBANK Settled", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONVFB.Close()
        Catch ex As System.Exception
            VRV_IP_FB = 0
            'INSERT_QT_OAB_FINBANK()
        End Try


    End Sub


    Private Sub INSERT_VRV_IP_FINBANK()

        'SQL
        Dim SQLCONNECTION2IP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double

        'V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "D"
        'V3 = fee
        V4 = "RECHARGE" & " " & DV1.ToShortDateString  'REF
        V5 = "999"
        'V6 = "D"
        V7 = VRV_IP_FB



        'INSERT MW
        SQLCONNECTION2IP.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION2IP
        CMD3.ExecuteNonQuery()
        SQLCONNECTION2IP.Close()





    End Sub













    'RECHARGE INCOME''''''''''''

    Private Sub RECHARGE_IB_INCOME()



        'SQL 
        Dim SQLCONNECTIONIB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDIB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERIB As System.Data.SqlClient.SqlDataReader

        Dim PARTNER As String
        Dim AMT As Double

        '""

        PARTNER = "FCMB Internet Banking"
        AMT = 0
        DV1 = DateTimePicker1.Value 'NO NEED TO CALL THIS PARAMETER AGAIN APPWIDE
        CMDIB.CommandText = String.Format("SELECT * FROM [RECHARGE_INCOME] WHERE [PARTNER] = '{0}' AND [FINAL SETTLEMENT]> {1} AND [SETTLE DATE] = '{2}'", PARTNER, AMT, DV1)

        SQLCONNECTIONIB.Open()
        CMDIB.Connection = SQLCONNECTIONIB
        SQLREADERIB = CMDIB.ExecuteReader


        If SQLREADERIB.HasRows Then
            GoTo 1
        Else
            MsgBox("There is no recharge Internet banking income", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONIB.Close()
            Exit Sub
        End If
1:      While (SQLREADERIB.Read)
            IB_AMT = Convert.ToDouble(SQLREADERIB.GetValue(SQLREADERIB.GetOrdinal("FINAL SETTLEMENT")))
            Insert_IB_income()
        End While
        MsgBox("Internet Banking Income Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTIONIB.Close()






    End Sub



    Private Sub Insert_IB_income()


        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double
        Dim SBU As String


        SBU = "90A"
        'V0 = "NGNEX99963130009"
        V1 = "NGNIN99954210036"
        V2 = "C"
        'V3 = fee
        V4 = "VTU" & " " & DV1.ToShortDateString
        V5 = "999"
        'V6 = "D"
        V7 = IB_AMT



        'INSERT MC POS PAYABLE RECHARGE
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, V7, V4, V5, SBU, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



    End Sub


    Private Sub RECHARGE_96N_INCOME()

        'SQL 
        Dim SQLCONNECTIONIB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDIB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERIB As System.Data.SqlClient.SqlDataReader

        Dim PARTNER As String
        Dim AMT As Double

        '""

        PARTNER = "First City Monumental Bank"
        AMT = 0
        DV1 = DateTimePicker1.Value 'NO NEED TO CALL THIS PARAMETER AGAIN APPWIDE
        CMDIB.CommandText = String.Format("SELECT * FROM [RECHARGE_INCOME] WHERE [PARTNER] = '{0}' AND [FINAL SETTLEMENT]> {1} AND [SETTLE DATE] = '{2}'", PARTNER, AMT, DV1)

        SQLCONNECTIONIB.Open()
        CMDIB.Connection = SQLCONNECTIONIB
        SQLREADERIB = CMDIB.ExecuteReader


        If SQLREADERIB.HasRows Then
            GoTo 1
        Else
            MsgBox("There is no recharge partner income", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONIB.Close()
            Exit Sub
        End If
1:      While (SQLREADERIB.Read)
            N96_AMT = Convert.ToDouble(SQLREADERIB.GetValue(SQLREADERIB.GetOrdinal("FINAL SETTLEMENT")))
            INSERT_96N_INCOME()
        End While
        MsgBox("FCMB Recharge Income Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTIONIB.Close()




    End Sub


    Private Sub INSERT_96N_INCOME()

        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double
        Dim SBU As String


        SBU = "96N"
        'V0 = "NGNEX99963130009"
        V1 = "NGNIN99954210036"
        V2 = "C"
        'V3 = fee
        V4 = "RECHARGE INCOME" & " " & DV1.ToShortDateString
        V5 = "999"
        'V6 = "D"
        V7 = N96_AMT



        'INSERT MC POS PAYABLE RECHARGE
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, V7, V4, V5, SBU, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



    End Sub



    Private Sub RECHARGE_HASAL_INCOME()

        'SQL 
        Dim SQLCONNECTIONIB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDIB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERIB As System.Data.SqlClient.SqlDataReader

        Dim PARTNER As String
        Dim AMT As Double

        '""

        PARTNER = "HASAL MICROFINANCE BANK"
        AMT = 0
        DV1 = DateTimePicker1.Value 'NO NEED TO CALL THIS PARAMETER AGAIN APPWIDE
        CMDIB.CommandText = String.Format("SELECT * FROM [RECHARGE_INCOME] WHERE [PARTNER] = '{0}' AND [FINAL SETTLEMENT]> {1} AND [SETTLE DATE] = '{2}'", PARTNER, AMT, DV1)

        SQLCONNECTIONIB.Open()
        CMDIB.Connection = SQLCONNECTIONIB
        SQLREADERIB = CMDIB.ExecuteReader


        If SQLREADERIB.HasRows Then
            GoTo 1
        Else
            MsgBox("There is no HASAL MFB income", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONIB.Close()
            Exit Sub
        End If
1:      While (SQLREADERIB.Read)
            HMFB = Convert.ToDouble(SQLREADERIB.GetValue(SQLREADERIB.GetOrdinal("FINAL SETTLEMENT")))
            INSERT_HMFB_INCOME()
        End While
        MsgBox("HASAL MFB Income Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTIONIB.Close()




    End Sub


    Private Sub INSERT_HMFB_INCOME()


        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double
        'Dim SBU As String


        'SBU = "96N"
        'V0 = "NGNEX99963130009"
        V1 = "0525932032"
        V2 = "C"
        'V3 = fee
        V4 = "RECHARGE INCOME" & " " & DV1.ToShortDateString
        V5 = "999"
        'V6 = "D"
        V7 = HMFB



        'INSERT MC POS PAYABLE RECHARGE
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



    End Sub



    Private Sub RECHARGE_FINBANK_INCOME()

        'SQL 
        Dim SQLCONNECTIONIB As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDIB As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERIB As System.Data.SqlClient.SqlDataReader

        Dim PARTNER As String
        Dim AMT As Double

        '""

        PARTNER = "First Inland Bank Plc"
        AMT = 0
        DV1 = DateTimePicker1.Value 'NO NEED TO CALL THIS PARAMETER AGAIN APPWIDE
        CMDIB.CommandText = String.Format("SELECT * FROM [RECHARGE_INCOME] WHERE [PARTNER] = '{0}' AND [FINAL SETTLEMENT]> {1} AND [SETTLE DATE] = '{2}'", PARTNER, AMT, DV1)

        SQLCONNECTIONIB.Open()
        CMDIB.Connection = SQLCONNECTIONIB
        SQLREADERIB = CMDIB.ExecuteReader


        If SQLREADERIB.HasRows Then
            GoTo 1
        Else
            MsgBox("There is no recharge FINBANK income", vbOKOnly, "Interswitch Automation")
            SQLCONNECTIONIB.Close()
            Exit Sub
        End If
1:      While (SQLREADERIB.Read)
            FB_INC = Convert.ToDouble(SQLREADERIB.GetValue(SQLREADERIB.GetOrdinal("FINAL SETTLEMENT")))
            INSERT_FINBANK_RCH_INCOME()
        End While
        MsgBox("FNBANK RECHARGE Income Settled", vbOKOnly, "Interswitch Automation")
        SQLCONNECTIONIB.Close()


    End Sub

    Private Sub INSERT_FINBANK_RCH_INCOME()

        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        'Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V5 As String
        ' Dim V6 As String
        Dim V7 As Double
        Dim SBU As String


        SBU = "96N"
        'V0 = "NGNEX99963130009"
        V1 = "NGNIN99954210036"
        V2 = "C"
        'V3 = fee
        V4 = "RECHARGE INCOME" & " " & DV1.ToShortDateString
        V5 = "999"
        'V6 = "D"
        V7 = FB_INC



        'INSERT MC POS PAYABLE RECHARGE
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, V7, V4, V5, SBU, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



    End Sub














    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        MasterCard_pos()

        Cashcard_FCMB()

        Etcc_FCMB()

        ISW_POS_PYBL()

        MC_Web_Recharge()

        MC_ATM_RCH()

        Verve_IP()

        QT_Verve_IP()

        OAB_3FMI0001()

        OAB_3BOL0001()

        OAB_3BOL0001_FINBANK()

        QT_VERVE_IP_FINBANK()

        CASHCARD_FINBANK()

        FW_FINBANK()

        Etcc_FINBANK()

        PP_FINBANK()

        VERVE_IP_FINBANK()



        MsgBox("Recharge settlement for" & DV1.ToShortDateString & " " & "is complete", vbOKOnly, "Interswitch_Automation")

        'INCOME

        RECHARGE_IB_INCOME()
        RECHARGE_96N_INCOME()
        RECHARGE_HASAL_INCOME()
        RECHARGE_FINBANK_INCOME()

        MsgBox("Recharge Income for" & DV1.ToShortDateString & " " & "is complete", vbOKOnly, "Interswitch_Automation")








    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class