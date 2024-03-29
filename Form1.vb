Option Strict Off
Imports Microsoft.Reporting
Imports Microsoft.ReportingServices

Public Class Form1



    'IFR -ISSUER FEE RECEIVABLE
    'IFP ISSUER FEE PAYABLE
    'AFR ACQUIRER FEE RECEIVABLE
    'AFP ACQUIRER FEE PAYABLE
    Public DV1 As Date
    Public acct_type As String
    Public amount As Double
    Public fee As Double
    Public amount1 As Double
    Public fee1 As Double
    Public amount2 As Double
    Public fee2 As Double
    Public amount3 As Double
    Public fee3 As Double
    Public amount4 As Double
    Public fee4 As Double
    Public amount5 As Double
    Public fee5 As Double
    Public amount6 As Double
    Public fee6 As Double
    Public amount7 As Double
    Public fee7 As Double
    Public amount8 As Double
    Public fee8 As Double
    Public amount9 As Double
    Public fee9 As Double
    Public amount10 As Double
    Public fee10 As Double

    Public fee11 As Double
    Public fee12 As Double
    Public fee13 As Double
    Public fee14 As Double
    Public fee15 As Double
    Public fee16 As Double
    Public fee17 As Double
    Public fee18 As Double
    Public fee19 As Double
    Public fee20 As Double

    Public fee21 As Double
    Public fee22 As Double
    Public fee23 As Double
    Public fee24 As Double
    Public fee25 As Double
    Public fee26 As Double
    Public fee27 As Double
    Public fee28 As Double
    Public fee29 As Double
    Public fee30 As Double

    Public fee31 As Double
    Public fee32 As Double
    Public fee33 As Double
    Public fee34 As Double
    Public fee35 As Double
    Public fee36 As Double
    Public fee37 As Double
    Public fee38 As Double
    Public fee39 As Double
    Public fee40 As Double


    Public QT_SVA_IFP As Double
    Public QT_SVA_FR As Double 'SVA FEE RCV

    Public Fc1 As Double 'CARDLESS VARIABLES
    Public fc2 As Double
    Public fc3 As Double
    Public fc4 As Double

    Public Fc5 As Double 'CARDLESS VARIABLES VERVE TOKEN 
    Public fc6 As Double
    Public fc7 As Double
    Public fc8 As Double

    Public FC9 As Double  'CARDLESS VARIABLES NON-CARD
    Public FC10 As Double
    Public FC11 As Double
    Public FC12 As Double

    Public fee_m1 As Double
    Public fee_m2 As Double
    Public amount_m1 As Double
    Public amount_m2 As Double

    Public amount11 As Double
    Public amount12 As Double
    Public amount13 As Double
    Public amount14 As Double
    Public amount15 As Double
    Public amount16 As Double
    Public amount17 As Double
    Public amount18 As Double
    Public amount19 As Double
    Public amount20 As Double
    Public amount21 As Double 'CARDLESS VARIABLES 21 AND 22
    Public amount22 As Double
    Public amount23 As Double 'CARDLESS VARIABLES VERVE TOKEN 23 AND 24
    Public amount24 As Double
    Public amount25 As Double 'SVA AMOUNT PAYABLE (IN PAYMENT GATEWAY)
    Public amount26 As Double 'CARDLESS VARIABLES NON-CARD 26 AND 27
    Public amount27 As Double
    Public amount28 As Double 'POS TRANSFERS 28 AND 29
    Public amount29 As Double

    Public Merchant_fee As Double 'POS CASHBACK MERCHANT FEE RCV

    Public Mcard_rou_count As Single
    Public AMOUNT30 As Double
    Public AMOUNT31 As Double

    Public bill_afr As Double
    Public bill_ifp As Double
    Public bill_ifr As Double
    Public bill_sva_fr As Double

    Public vit_ar As Double
    Public vit_ap As Double
    Public VIT_ifp As Double
    Public VIT_ifr As Double

    Public fee_ap_rcv As Double

    Public PPC_aP As Double
    Public PPC_aR As Double
    Public PPC_IFR As Double
    Public PPC_IFP As Double
    Public Mcard_dollar As Double












    Private Sub Eliminate_duplication()
        'SQL  
        Dim SQLCONNECTION1 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1 As System.Data.SqlClient.SqlDataReader
        CMD1.CommandType = System.Data.CommandType.Text   'command syntax

        DV1 = DateTimePicker1.Value.Date

        CMD1.CommandText = String.Format("SELECT * FROM [SETTLEMENT] WHERE [DATE] = '{0}'", DV1)

        SQLCONNECTION1.Open()
        CMD1.Connection = SQLCONNECTION1
        SQLREADER1 = CMD1.ExecuteReader



        If SQLREADER1.HasRows Then
            'CHECK FOR DUPLICATE POSTINGS
            MsgBox("Complete Report Already Generated for the specified settlement date", vbOKOnly, "Interswitch Automation")
            GoTo 1
        Else
            SQLCONNECTION1.Close()
            Generate_report()
        End If


1:      SQLCONNECTION1.Close()



    End Sub


    Private Sub Generate_report()


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
            'CALL ALL SETTLEMENT ITEMS HERE
            MasterCard_Dollar_Settlement()
            CALL_MOBILITY()
            CALL_NOU()
            CALL_ROU()
            CALL_Autopay()
            CALL_PAYMENT_GATEWAY()
            CALL_UPPERLINK_TRANSFERS()
            CALL_REMOTE_POS_ACQUIRER_FEES_2()
            CHECK_LATE_POS_CH_IFR()
            Call_Remote_Pos_Issuer_fees_2()
            CALL_POS_REWARD_MONEY_SPEND()
            Call_remote_pos_cardholder2()
            CALL_PREPAID_CARDLOAD()
            CALL_ATM_CARDLESS()
            CALL_ATM_CARDLESS_VERVE_TOKEN()
            CALL_ATM_CARDLESS_NON_CARD()
            CALL_Remote_Web()
            CALL_Remote_web_fees()
            CALL_Remote_web_fees_iso()
            Web_Acquired_amounts()
            Web_Acquired_Acquirer_fee_payable()
            CALL_Web_Acquired_Fee_Receivable()
            CHECK_LATE_POS_AP()
            CHECK_LATE_POS_AR()
            CHECK_LATE_POS_AFR()
            CHECK_LATE_POS_AFP()
            CALL_POS_TRANSFERS()
            Call_POS_CASHBACK_Merchant_fee()
            CALL_RELATIONAL_TRANSFER()
            CALL_VERVE_INTERNATIONAL()
            CALL_BILLPAYMENT()
            'INSERT SETTLEMENT RECORDS HERE
            Insert_MasterCard_Dollar_Settlement()
            MOBILITY_Insert()
            NOU_INSERT()
            INSERT_ROU()
            Autopay_Insert()
            PG_INSERT()
            UPPERLINK_INSERT()
            CALL_ATM_CARDLESS_INSERT()
            CALL_ATM_CARDLESS_VT_INSERT()
            CALL_ATM_CARDLESS_NON_CARD_INSERT()
            Remote_Pos_Acquirer_Insert()
            Remote_Pos_Issuer_Insert()
            POS_REWARD_MONEY_SPEND_INSERT()
            REMOTE_Pos_cardholder_issuer_fee_insert()
            Insert_Remote_Web_Amounts()
            Insert_Remote_WEB_all_fees()
            Insert_Web_Acquired_amounts()
            Insert_Web_Acquired_Fee_Receivable()
            VERVE_INTERNATIONAL_INSERT()
            CALL_POS_TRANSFER_INSERT()
            Insert_POS_CASHBACK_MERCHANT_FEE()
            RELATIONAL_TRANSFERS_INSERT()
            BILLPAYMENT_INSERT()
            INSERT_PREPAID_CARD()

            MsgBox("Complete Report Generated", vbOKOnly, "Interswitch Automation")
        Else
            MsgBox("There is no settlement report for the date specified", vbOKOnly, "Interswitch Automation")
        End If


        SQLCONNECTION1.Close()

    End Sub

    Private Sub CALL_POS_TRANSFERS()

        'SQL 
        Dim SQLCONNECTION2 As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2 As New System.Data.SqlClient.SqlCommand
        Dim CMD3 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2 As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String
        Dim acct_type1 As String

        'POS TRANSFERS

        account = "POS TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        acct_type1 = "AMOUNT RECEIVABLE"

        CMD2.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)
        CMD3.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type1)

        'for amount payable
        Try
            SQLCONNECTION2.Open()
            CMD2.Connection = SQLCONNECTION2
            SQLREADER2 = CMD2.ExecuteReader


            If SQLREADER2.HasRows Then
                While (SQLREADER2.Read)
                    amount28 = SQLREADER2.GetValue(SQLREADER2.GetOrdinal("AMT")).ToString()
                End While
            End If
        Catch EX As Exception
            MsgBox("There is no pos transfer amount payable")
            SQLCONNECTION2.Close()
        End Try

        SQLCONNECTION2.Close()


        'for amount recivable
        Try
            SQLCONNECTION2.Open()
            CMD3.Connection = SQLCONNECTION2
            SQLREADER2 = CMD3.ExecuteReader


            If SQLREADER2.HasRows Then
                While (SQLREADER2.Read)
                    amount29 = SQLREADER2.GetValue(SQLREADER2.GetOrdinal("AMT")).ToString()
                End While
            End If
        Catch EX As Exception
            MsgBox("There is no pos transfer amount receivable")
            SQLCONNECTION2.Close()
        End Try

        SQLCONNECTION2.Close()



    End Sub

    Private Sub CALL_POS_TRANSFER_INSERT()

        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V3 As Double
        Dim V4 As String
        Dim V5 As String
        Dim V6 As String
        Dim V7 As Double

        V0 = "NGNAS99916151005"
        V1 = "NGNLI99924434011"
        V2 = "D"
        V3 = amount28
        V4 = "POS TRSFR SETTLEMENT" & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "C"
        V7 = amount29



        'INSERT AP (POS TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V3, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



        'INSERT AR (POS TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V0, V6, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()


    End Sub

    Private Sub Call_POS_CASHBACK_Merchant_fee()


        'SQL 
        Dim SQLCONNECTION2 As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2 As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String


        'POS PURCHASE WITH CASHBACK

        account = "POS PURCHASE WITH CASHBACK"
        status = "CURRENT"
        acct_type = "MERCHANT FEE RECEIVABLE"


        CMD2.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)



        Try
            SQLCONNECTION2.Open()
            CMD2.Connection = SQLCONNECTION2
            SQLREADER2 = CMD2.ExecuteReader


            If SQLREADER2.HasRows Then
                While (SQLREADER2.Read)
                    Merchant_fee = SQLREADER2.GetValue(SQLREADER2.GetOrdinal("FEE")).ToString()
                End While
            End If
        Catch EX As Exception
            MsgBox("There is no merchant receivable for POS Cash Back")
            SQLCONNECTION2.Close()
        End Try

        SQLCONNECTION2.Close()


    End Sub

    Private Sub Insert_POS_CASHBACK_MERCHANT_FEE()

        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        Dim V0 As String
        Dim V2 As String  'STATUS
        Dim V3 As Double
        Dim V4 As String
        Dim V5 As String
        Dim V6 As String


        V0 = "XXXXX"
        V2 = "C"
        V3 = Merchant_fee
        V4 = "POS CASHBACK MERCHANT FEE " & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "CONFIRM MERCHANT TO BE CREDITED FROM POS REPORT"



        'INSERT POS CASHBACK MERCHANT FEE 
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V2, V3, V4, V5, V6, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



    End Sub



    Private Sub CALL_POS_REWARD_MONEY_SPEND()


        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader




        Dim status As String
        Dim account As String


        '"REWARD MONEY (SPEND) POS FEE SETTLEMENT"

        ''START AFR
        account = "REWARD MONEY (SPEND) POS FEE SETTLEMENT"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"
        CMD1RP.CommandText = String.Format("SELECT [FEE] FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                fee21 = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION1RP.Close()
        '''''''''''''''''''''''''''''''''''''''''''END AFR'''''''''''''''''

        ''START IFP

        account = "REWARD MONEY (SPEND) POS FEE SETTLEMENT"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD1RP.CommandText = String.Format("SELECT [FEE] FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                fee22 = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION1RP.Close()

        ''''''''''''''''''''''''''''''''END IFP''''''''''''''''



    End Sub


    Private Sub POS_REWARD_MONEY_SPEND_INSERT()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim V1 As String
        Dim Vcon1 As String
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String


        SBU = "93Q"
        CMNT = "REWARD MONEY SETTLEMENT"


        V0 = "NGNIN99954210036"
        Vcon = "C"


        V1 = "NGNEX99963130009"
        Vcon1 = "D"

        V4 = "POS REWARD MONEY" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT AFR RWD MONEY
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee21, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()

        'INSERT IFP RWD MONEY 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V1, Vcon1, fee22, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()





    End Sub


    Private Sub REMOTE_Pos_cardholder_issuer_fee_insert()


        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        'Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String


        'SBU = "93Q"
        CMNT = "THE REMOTE POS REPORT SHOULD BE USED TO RAISE CARDHOLDER FEES"


        V0 = "XXXXXXXXXX"
        Vcon = "C"

        V4 = "LUMP:CARDHOLDER REWARD" & " " & DV1.ToShortDateString
        V5 = "999"

        'FEE37 IS A NET CARDHOLDER FEES INCLUDING LATE POS REVERSALS


        'INSERT NET REMOTE POS CARDHOLDER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, Vcon, fee37, V4, V5, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()





    End Sub


    Private Sub Call_remote_pos_cardholder2()

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        Dim Test As Double


        account = "POS"
        account1 = "REWARD MONEY (SPEND) POS FEE SETTLEMENT"
        status = "CURRENT"
        acct_type = "CARDHOLDER_ISSUER FEE RECEIVABLE"


        'CMD1RP.CommandText = String.Format("SELECT [FEE] FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)
        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) as 'TEST' FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type, status) 'SUM AS 'TEST I.E NO NEED TO DIM TEST AS DOUBLE
        Try
            SQLCONNECTION1RP.Open()
            CMD1RP.Connection = SQLCONNECTION1RP
            SQLREADER1RP = CMD1RP.ExecuteReader


            If SQLREADER1RP.HasRows Then
                While (SQLREADER1RP.Read)
                    Test = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                    fee23 = Test
                End While
            End If

            'Test = Convert.ToDouble(CMD1RP.ExecuteScalar)
            'fee23 = Test


        Catch ex As Exception
            'call the check_late_pos here
            fee37 = fee23 + fee36  'THIS IS THE NET CARDHOLDER FEES INCLUDING LATE REVERSALS
            MsgBox("Cardholder fee is" & " " & Test)
            MsgBox("There are no Cardholder Issuer Fee Receivable")
            SQLCONNECTION1RP.Close()
            Exit Sub
        End Try


        'Call the check_late_pos here
        fee37 = fee23 + fee36  'THIS IS THE NET CARDHOLDER FEES INCLUDING LATE REVERSALS
        MsgBox("Cardholder fee is" & " " & Test)
        SQLCONNECTION1RP.Close()





    End Sub





    Private Sub CALL_REMOTE_POS_ACQUIRER_FEES_2()

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        Dim acct_type1 As String
        Dim acct_type2 As String
        Dim TEST As Double



        account = "POS"
        account1 = "REWARD MONEY (SPEND) POS FEE SETTLEMENT"
        status = "CURRENT"
        acct_type = "ACQUIRER"
        acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"
        acct_type2 = "MASTERCARD LOCAL PROCESSING BILLING(POS PURCHASE)" 'recent addition

        CMD1RP.CommandText = String.Format("Select SUM(FEE) As TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCT_TYPE]<>'{3}' AND [ACCT_TYPE]<>'{4}' AND [ACCT_TYPE] LIKE '%{5}%'  AND [STATUS]='{6}'", DV1, account, account1, acct_type1, acct_type2, acct_type, status)



        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE <>'CARDHOLDER_ISSUER FEE RECEIVABLE' AND ACCOUNT <> 'REWARD MONEY (SPEND) POS FEE SETTLEMENT'
        'AND ACCT_TYPE LIKE'%ACQUIRER%' AND SETTLEMENT_DATE='2014-07-11'





        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                fee19 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()





    End Sub


    Private Sub Remote_Pos_Acquirer_Insert()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = ""
        Vcon = ""
        SBU = "93Q"
        CMNT = "THIS IS A NET REMOTE POS ACQUIRER FEES"

        If fee19 > 0 Then
            V0 = "NGNIN99954210036"
            Vcon = "C"
        End If
        If fee19 < 0 Then
            V0 = "NGNEX99963130009"
            Vcon = "D"
        End If

        V4 = "LUMP:POS ACQ FEE" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee19, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()







    End Sub




    Private Sub Call_Remote_Pos_Issuer_fees_2()

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        Dim account2 As String
        Dim acct_type1 As String
        Dim TEST As Double



        account = "POS"
        account1 = "REWARD MONEY (SPEND) POS FEE SETTLEMENT"
        status = "CURRENT"
        acct_type = "ISSUER"
        acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"
        account2 = "MASTERCARD LOCAL PROCESSING BILLING(POS PURCHASE)"

        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCOUNT]<>'{3}' AND [ACCT_TYPE]<>'{4}' AND [ACCT_TYPE] LIKE '%{5}%'  AND [STATUS]='{6}'", DV1, account, account1, account2, acct_type1, acct_type, status)



        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE <>'CARDHOLDER_ISSUER FEE RECEIVABLE' AND ACCOUNT <> 'REWARD MONEY (SPEND) POS FEE SETTLEMENT'
        'AND ACCT_TYPE LIKE'%ISSUER%' AND SETTLEMENT_DATE='2014-07-11'





        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                fee20 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()



    End Sub

    Private Sub Remote_Pos_Issuer_Insert()

        'SQL
        Dim SQLCONNECTION2RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = ""
        Vcon = ""
        SBU = "96N"
        CMNT = "THIS IS A NET REMOTE POS ISSUER FEES"

        If fee20 > 0 Then
            V0 = "NGNIN99954210036"
            Vcon = "C"
        End If
        If fee20 < 0 Then
            V0 = "NGNEX99963130009"
            Vcon = "D"
        End If

        V4 = "LUMP:POS ISSUER FEES" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ISSUER FEES 
        SQLCONNECTION2RP.Open()
        CMD2RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee20, V4, V5, SBU, CMNT, DV1)
        CMD2RP.Connection = SQLCONNECTION2RP
        CMD2RP.ExecuteNonQuery()
        SQLCONNECTION2RP.Close()


    End Sub


    Private Sub CALL_PAYMENT_GATEWAY()

        'PAYMENT_GATEWAY
        'SQL 
        Dim SQLCONNECTION1PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1PG As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"PAYMENT_GATEWAY TRANSFERS"

        account = "OTHER TRANSFERS"  'OR "OTHER TRANSFERS(NON GENERIC PLATFORM)" THE NON GENERIC PLATFORM IS WHY WE USED LIKE IN THE PG QUERY STATEMENT
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD1PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1PG.Open()
        CMD1PG.Connection = SQLCONNECTION1PG
        SQLREADER1PG = CMD1PG.ExecuteReader


        If SQLREADER1PG.HasRows Then
            While (SQLREADER1PG.Read)
                amount10 = SQLREADER1PG.GetValue(SQLREADER1PG.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1PG.Close()

        ''''''''''''''''''''''''''''END AR OTHER TRANSFERS'''''''''''''''''''''''''''''''


        '''''''''''''''''''''''''START AR QUICKTELLER TRANSFERS(SVA)''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION2PG As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2PG As System.Data.SqlClient.SqlDataReader



        '"QUICKTELLER TRANSFERS(SVA)"

        account = "QUICKTELLER TRANSFERS(SVA)"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2PG.Open()
        CMD2PG.Connection = SQLCONNECTION2PG
        SQLREADER2PG = CMD2PG.ExecuteReader


        If SQLREADER2PG.HasRows Then
            While (SQLREADER2PG.Read)
                amount11 = SQLREADER2PG.GetValue(SQLREADER2PG.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2PG.Close()

        ''''''''''''''''''''''''''''END QUICKTELLER TRANSFERS(SVA) AR'''''''''''''''''''''''''''''''





        '''''''''''''''''''''''''START AP QUICKTELLER TRANSFERS(SVA)''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTIONQT_SVA_AP As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDQT_SVA_AP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERQT_SVA_AP As System.Data.SqlClient.SqlDataReader



        '"QUICKTELLER TRANSFERS(SVA)"  AMOUNT PAYABLE

        account = "QUICKTELLER TRANSFERS(SVA)"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMDQT_SVA_AP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTIONQT_SVA_AP.Open()
        CMDQT_SVA_AP.Connection = SQLCONNECTIONQT_SVA_AP
        SQLREADERQT_SVA_AP = CMDQT_SVA_AP.ExecuteReader


        If SQLREADERQT_SVA_AP.HasRows Then
            While (SQLREADERQT_SVA_AP.Read)
                AMOUNT25 = SQLREADERQT_SVA_AP.GetValue(SQLREADERQT_SVA_AP.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTIONQT_SVA_AP.Close()


        ''''''''''''''''''''''''''''END QUICKTELLER TRANSFERS(SVA) AP'''''''''''''''''''''''''''''''




        '''''''''''''''''''''''''START AP QUICKTELLER WEB TRANSFERS''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION4PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4PG As System.Data.SqlClient.SqlDataReader



        'QUICKTELLER WEB TRANSFERS

        account = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD4PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4PG.Open()
        CMD4PG.Connection = SQLCONNECTION4PG
        SQLREADER4PG = CMD4PG.ExecuteReader


        If SQLREADER4PG.HasRows Then
            While (SQLREADER4PG.Read)
                amount12 = SQLREADER4PG.GetValue(SQLREADER4PG.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION4PG.Close()

        ''''''''''''''''''''''''''''END AP QUICKTELLER WEB TRANSFERS''''''''''''''''''''''''''''''



        '''''''''''''''''''''''''START AR QUICKTELLER WEB TRANSFERS''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION5PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5PG As System.Data.SqlClient.SqlDataReader



        'QUICKTELLER WEB TRANSFERS

        account = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD5PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION5PG.Open()
        CMD5PG.Connection = SQLCONNECTION5PG
        SQLREADER5PG = CMD5PG.ExecuteReader


        If SQLREADER5PG.HasRows Then
            While (SQLREADER5PG.Read)
                amount13 = SQLREADER5PG.GetValue(SQLREADER5PG.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION5PG.Close()

        ''''''''''''''''''''''''''''END AR QUICKTELLER WEB TRANSFERS''''''''''''''''''''''''''''''





        '''''''''''''''''''''''''START REMITA INITIATED TRANSFERS''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION6PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD6PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER6PG As System.Data.SqlClient.SqlDataReader



        'REMITA TRANSFERS

        account = "REMITA TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD6PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION6PG.Open()
        CMD6PG.Connection = SQLCONNECTION6PG
        SQLREADER6PG = CMD6PG.ExecuteReader


        If SQLREADER6PG.HasRows Then
            While (SQLREADER6PG.Read)
                amount14 = SQLREADER6PG.GetValue(SQLREADER6PG.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION6PG.Close()

        ''''''''''''''''''''''''''''END REMITA INITIATED TRANSFERS''''''''''''''''''''''''''''''



        '''''''''''''''''''''''''START REMITA RECEIVED TRANSFERS''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION9PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD9PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER9PG As System.Data.SqlClient.SqlDataReader



        'REMITA TRANSFERS

        account = "REMITA TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD9PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION9PG.Open()
        CMD9PG.Connection = SQLCONNECTION9PG
        SQLREADER9PG = CMD9PG.ExecuteReader


        If SQLREADER9PG.HasRows Then
            While (SQLREADER9PG.Read)
                amount15 = SQLREADER9PG.GetValue(SQLREADER9PG.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION9PG.Close()

        ''''''''''''''''''''''''''''END REMITA RECEIVED TRANSFERS''''''''''''''''''''''''''''''


        '''''''''''''''''END AMOUNTS PG
        '''''''''''''''''
        ''''''''''''''''' START FEES PG




        '''''''''''''''''''''''''START PAYMENT_GATEWAY Fees''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION3PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3PG As System.Data.SqlClient.SqlDataReader



        '"QUICKTELLER TRANSFERS(SVA)" IFR

        account = "QUICKTELLER TRANSFERS(SVA)"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD3PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3PG.Open()
        CMD3PG.Connection = SQLCONNECTION3PG
        SQLREADER3PG = CMD3PG.ExecuteReader


        If SQLREADER3PG.HasRows Then
            While (SQLREADER3PG.Read)
                fee15 = SQLREADER3PG.GetValue(SQLREADER3PG.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3PG.Close()

        '''''''''''''''''''''''''''''''END '"QUICKTELLER TRANSFERS(SVA)" IFR''''''''''''''''''''''''


        'START "QUICKTELLER TRANSFERS(SVA)" IFP

        Dim SQLCONNECTION8PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD8PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER8PG As System.Data.SqlClient.SqlDataReader

        '"QUICKTELLER TRANSFERS(SVA)" IFP

        account = "QUICKTELLER TRANSFERS(SVA)"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD8PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION8PG.Open()
        CMD8PG.Connection = SQLCONNECTION8PG
        SQLREADER8PG = CMD8PG.ExecuteReader


        If SQLREADER8PG.HasRows Then
            While (SQLREADER8PG.Read)
                QT_SVA_IFP = SQLREADER8PG.GetValue(SQLREADER8PG.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION8PG.Close()

        '''''''''''''''''''''''''''''''END '"QUICKTELLER TRANSFERS(SVA)" IFP''''''''''''''''''''''''



        'START "QUICKTELLER TRANSFERS(SVA)" SVA_FR

        Dim SQLCONNECTION10PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD10PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER10PG As System.Data.SqlClient.SqlDataReader



        '"QUICKTELLER TRANSFERS(SVA)" SVA FEE RECEIVABLE

        account = "QUICKTELLER TRANSFERS(SVA)"
        status = "CURRENT"
        acct_type = "SVA FEE RECEIVABLE"
        CMD10PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION10PG.Open()
        CMD10PG.Connection = SQLCONNECTION10PG
        SQLREADER10PG = CMD10PG.ExecuteReader


        If SQLREADER10PG.HasRows Then
            While (SQLREADER10PG.Read)
                QT_SVA_FR = SQLREADER10PG.GetValue(SQLREADER10PG.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION10PG.Close()

        '''''''''''''''''''''''''''''''END '"QUICKTELLER TRANSFERS(SVA)" SVA FEE RECEIVABLE''''''''''''''''''''''''



        'QUICKTELLER WEB TRANSFERS FEES

        Dim SQLCONNECTION7PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD7PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER7PG As System.Data.SqlClient.SqlDataReader



        '"QUICKTELLER WEB TRANSFERS" IFR

        account = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD7PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION7PG.Open()
        CMD7PG.Connection = SQLCONNECTION7PG
        SQLREADER7PG = CMD7PG.ExecuteReader


        If SQLREADER7PG.HasRows Then
            While (SQLREADER7PG.Read)
                fee16 = SQLREADER7PG.GetValue(SQLREADER7PG.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION7PG.Close()

        '''''''''''''''''''''''''''''''''' '"END QUICKTELLER WEB TRANSFERS" IFR''''''''''''''''



        '''''''''''''''''''''''''''''''''''''"START QUICKTELLER WEB TRANSFERS" IFP'''''''''''''''''

        Dim SQLCONNECTION18PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD18PG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER18PG As System.Data.SqlClient.SqlDataReader

        'QUICKTELLER WEB TRANSFERS" IFP

        account = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD18PG.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION18PG.Open()
        CMD18PG.Connection = SQLCONNECTION18PG
        SQLREADER18PG = CMD18PG.ExecuteReader


        If SQLREADER18PG.HasRows Then
            While (SQLREADER18PG.Read)
                fee17 = SQLREADER18PG.GetValue(SQLREADER18PG.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION18PG.Close()



        ''''''''''''''''''' '"END QUICKTELLER WEB TRANSFERS" IFP



        ''''''''''''''''''''''''''''END PAYMENT_GATEWAY fees'''''''''''''''''''''''''''''''



    End Sub

    Private Sub PG_INSERT()



        'SQL
        Dim SQLCONNECTION2PG As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2PG As New System.Data.SqlClient.SqlCommand
        Dim CMD3PG As New System.Data.SqlClient.SqlCommand
        Dim CMD4PG As New System.Data.SqlClient.SqlCommand
        Dim CMD5PG As New System.Data.SqlClient.SqlCommand
        Dim CMD6PG As New System.Data.SqlClient.SqlCommand
        Dim CMD7PG As New System.Data.SqlClient.SqlCommand
        Dim CMD8PG As New System.Data.SqlClient.SqlCommand
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V4c As String
        Dim V4d As String
        Dim V4e As String
        Dim V4f As String
        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        Dim CMNT As String
        Dim V0 As String
        Dim Vstat As String
        Dim OAB As String

        OAB = "NGNLI99924430087"
        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "C"
        V4 = "LUMP:PGW FEES" & " " & DV1.ToShortDateString
        V4a = "LUMP:PGW OTHER TRANSFERS" & " " & DV1.ToShortDateString
        V4b = "LUMP:PGW QUICKTELLER TRANSFERS(SVA)" & " " & DV1.ToShortDateString
        V4c = "LUMP:PGW QUICKTELLER WEB TRANSFERS(RCVBLE)" & " " & DV1.ToShortDateString
        V4d = "LUMP:PGW QUICKTELLER WEB TRANSFERS(PAYBLE)" & " " & DV1.ToShortDateString
        V4e = "LUMP:PGW REMITA PAYBLE" & " " & DV1.ToShortDateString
        V4f = "LUMP:PGW REMITA RCVBLE" & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "D"
        SBU = "94L"
        CMNT = "THIS IS A NET PG FEES"
        fee18 = 1 * (fee15 + fee16 + fee17 + QT_SVA_FR + QT_SVA_IFP) 'NET PG FEE
        Vstat = ""

        If fee18 > 0 Then
            Vstat = "C"
        End If
        If fee18 < 0 Then
            Vstat = "D"
        End If






        'INSERT NET PG FEES 
        SQLCONNECTION2PG.Open()
        CMD3PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vstat, fee18, V4, V5, SBU, CMNT, DV1)
        CMD3PG.Connection = SQLCONNECTION2PG
        CMD3PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()



        'INSERT OTHER TRANSFERS (AMOUNT RECEIVABLE)
        SQLCONNECTION2PG.Open()
        CMD4PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount10, V4a, V5, DV1)
        CMD4PG.Connection = SQLCONNECTION2PG
        CMD4PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()

        'INSERT QUICKTELLER TRANSFERS(SVA) (AMOUNT RECEIVABLE)
        SQLCONNECTION2PG.Open()
        CMD2PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount11, V4b, V5, DV1)
        CMD2PG.Connection = SQLCONNECTION2PG
        CMD2PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()

        'INSERT QUICKTELLER TRANSFERS(SVA) (AMOUNT PAYABLE)
        SQLCONNECTION2PG.Open()
        CMD2PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, AMOUNT25, V4b, V5, DV1)
        CMD2PG.Connection = SQLCONNECTION2PG
        CMD2PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()


        'INSERT QUICKTELLER WEB TRANSFERS (AMOUNT RECEIVABLE)
        SQLCONNECTION2PG.Open()
        CMD5PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount13, V4c, V5, DV1)
        CMD5PG.Connection = SQLCONNECTION2PG
        CMD5PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()

        'INSERT QUICKTELLER WEB TRANSFERS (AMOUNT PAYABLE)
        SQLCONNECTION2PG.Open()
        CMD6PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, amount12, V4d, V5, DV1)
        CMD6PG.Connection = SQLCONNECTION2PG
        CMD6PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()


        'INSERT REMITA TRANSFERS (AMOUNT PAYABLE)
        SQLCONNECTION2PG.Open()
        CMD7PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", OAB, V6, amount14, V4e, V5, DV1)
        CMD7PG.Connection = SQLCONNECTION2PG
        CMD7PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()


        'INSERT REMITA TRANSFERS (AMOUNT RECEIVABLE)
        SQLCONNECTION2PG.Open()
        CMD7PG.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount15, V4f, V5, DV1)
        CMD7PG.Connection = SQLCONNECTION2PG
        CMD7PG.ExecuteNonQuery()
        SQLCONNECTION2PG.Close()




    End Sub


    Private Sub CALL_Autopay()

        'AUTOPAY
        'SQL 
        Dim SQLCONNECTION1AUTOPAY As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1AUTOPAY As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"AUTOPAY TRANSFERS"

        account = "AUTOPAY TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD1AUTOPAY.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1AUTOPAY.Open()
        CMD1AUTOPAY.Connection = SQLCONNECTION1AUTOPAY
        SQLREADER1AUTOPAY = CMD1AUTOPAY.ExecuteReader


        If SQLREADER1AUTOPAY.HasRows Then
            While (SQLREADER1AUTOPAY.Read)
                amount8 = SQLREADER1AUTOPAY.GetValue(SQLREADER1AUTOPAY.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1AUTOPAY.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''


        '''''''''''''''''''''''''START AR''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION2AUTOPAY As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2AUTOPAY As System.Data.SqlClient.SqlDataReader



        '"AUTOPAY TRANSFERS"

        account = "AUTOPAY TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2AUTOPAY.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2AUTOPAY.Open()
        CMD2AUTOPAY.Connection = SQLCONNECTION2AUTOPAY
        SQLREADER2AUTOPAY = CMD2AUTOPAY.ExecuteReader


        If SQLREADER2AUTOPAY.HasRows Then
            While (SQLREADER2AUTOPAY.Read)
                amount9 = SQLREADER2AUTOPAY.GetValue(SQLREADER2AUTOPAY.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2AUTOPAY.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''




        '''''''''''''''''''''''''START Autopay Fees''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION3AUTOPAY As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3AUTOPAY As System.Data.SqlClient.SqlDataReader



        '"AUTOPAY TRANSFERS"

        account = "AUTOPAY TRANSFER FEES"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD3AUTOPAY.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3AUTOPAY.Open()
        CMD3AUTOPAY.Connection = SQLCONNECTION3AUTOPAY
        SQLREADER3AUTOPAY = CMD3AUTOPAY.ExecuteReader


        If SQLREADER3AUTOPAY.HasRows Then
            While (SQLREADER3AUTOPAY.Read)
                fee14 = SQLREADER3AUTOPAY.GetValue(SQLREADER3AUTOPAY.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3AUTOPAY.Close()




        Dim SQLCONNECTION4AUTOPAY As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4AUTOPAY As System.Data.SqlClient.SqlDataReader



        '"AUTOPAY TRANSFERS"

        account = "AUTOPAY TRANSFER FEES"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD4AUTOPAY.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4AUTOPAY.Open()
        CMD4AUTOPAY.Connection = SQLCONNECTION4AUTOPAY
        SQLREADER4AUTOPAY = CMD4AUTOPAY.ExecuteReader


        If SQLREADER4AUTOPAY.HasRows Then
            While (SQLREADER4AUTOPAY.Read)
                fee_ap_rcv = SQLREADER4AUTOPAY.GetValue(SQLREADER4AUTOPAY.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4AUTOPAY.Close()


        fee14 = fee14 + fee_ap_rcv


        ''''''''''''''''''''''''''''END Autopay fees'''''''''''''''''''''''''''''''

























    End Sub

    Private Sub Autopay_Insert()




        'SQL
        Dim SQLCONNECTION2AUTOPAY As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim CMD3AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim CMD4AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim CMD5AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim CMD6AUTOPAY As New System.Data.SqlClient.SqlCommand
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        Dim CMNT As String
        Dim V0 As String




        V0 = "XXXXXXXXXX"
        V1 = "NGNLI00124430030"
        V2 = "C"
        V4 = "AUTOPAY PYBLE " & " " & DV1.ToShortDateString
        V4a = "AUTOPAY RCVBLE" & " " & DV1.ToShortDateString
        V4b = "AUTOPAY FEES" & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "D"
        SBU = "96N"
        CMNT = "RAISE THE NECESSARY AUTOPAY CREDIT AND DEBIT ENTRIES USING THE PDF REPORT FOR AUTOPAY"




        'INSERT AP (AUTOPAY) 
        SQLCONNECTION2AUTOPAY.Open()
        CMD3AUTOPAY.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, amount8, V4, V5, DV1)
        CMD3AUTOPAY.Connection = SQLCONNECTION2AUTOPAY
        CMD3AUTOPAY.ExecuteNonQuery()
        SQLCONNECTION2AUTOPAY.Close()


        'INSERT AR (AUTOPAY)
        SQLCONNECTION2AUTOPAY.Open()
        CMD4AUTOPAY.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount9, V4a, V5, DV1)
        CMD4AUTOPAY.Connection = SQLCONNECTION2AUTOPAY
        CMD4AUTOPAY.ExecuteNonQuery()
        SQLCONNECTION2AUTOPAY.Close()


        'INSERT AUTOPAY FEES
        SQLCONNECTION2AUTOPAY.Open()
        CMD2AUTOPAY.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, fee14, V4b, V5, CMNT, DV1)
        CMD2AUTOPAY.Connection = SQLCONNECTION2AUTOPAY
        CMD2AUTOPAY.ExecuteNonQuery()
        SQLCONNECTION2AUTOPAY.Close()










    End Sub


    Private Sub CALL_ROU()

        'ROU
        'SQL 
        Dim SQLCONNECTION1ROU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1ROU As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1ROU As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"ATM WITHDRAWAL (REGULAR)"

        account = "ATM WITHDRAWAL (REGULAR)"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD1ROU.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1ROU.Open()
        CMD1ROU.Connection = SQLCONNECTION1ROU
        SQLREADER1ROU = CMD1ROU.ExecuteReader


        If SQLREADER1ROU.HasRows Then
            While (SQLREADER1ROU.Read)
                amount6 = SQLREADER1ROU.GetValue(SQLREADER1ROU.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1ROU.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''

        'ISSUER FEE PAYABLE (VERVE)

        'SQL 
        Dim SQLCONNECTION3ROU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3ROU As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3ROU As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (REGULAR)"

        account = "ATM WITHDRAWAL (VERVE BILLING)"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD3ROU.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3ROU.Open()
        CMD3ROU.Connection = SQLCONNECTION3ROU
        SQLREADER3ROU = CMD3ROU.ExecuteReader


        If SQLREADER3ROU.HasRows Then
            While (SQLREADER3ROU.Read)
                fee10 = SQLREADER3ROU.GetValue(SQLREADER3ROU.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3ROU.Close()


        ''''''''''''''''''''''''''''''END IFP VERVE'''''''''''''''''''''''''''

        'ISSUER FEE PAYABLE (VERVE AND MASTERCARD COMBINED)

        'SQL 
        Dim SQLCONNECTION4ROU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4ROU As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4ROU As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (REGULAR)"

        account = "ATM WITHDRAWAL (REGULAR)"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD4ROU.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ROU.Open()
        CMD4ROU.Connection = SQLCONNECTION4ROU
        SQLREADER4ROU = CMD4ROU.ExecuteReader


        If SQLREADER4ROU.HasRows Then
            While (SQLREADER4ROU.Read)
                fee13 = SQLREADER4ROU.GetValue(SQLREADER4ROU.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ROU.Close()

        '''''''''''''''''''''''''''''''''IFP COMBINED END''''''''''''''''''''''''''''''''''''''''















        'ROU REDUCTION
        Dim V9 As Single
        Dim V10 As Double
        Dim V11 As Double
        V9 = fee10 / 5  'VERVE ROU COUNT
        V10 = (60 * V9)  '60 CHRGED BY ISW * VERVE ROW COUNT

        V11 = fee10 + V10  'i.e. (VERVE BILLING AT N5 + V10) or (v9 *65)
        amount7 = V11

        'fee11 = V11 '(ROU reduction back to N60 charged,credit expense back)

        'MasterCard Row count
        Dim Mcard_expense As Double
        Mcard_expense = Math.Abs(fee13) - (Math.Abs(V9) * 60) 'no need to divide by 90 which is amount charged by interswitch for Mcard ROU transactions
        fee12 = (-1 * Mcard_expense)
        Mcard_rou_count = (Mcard_expense / 55)     'NOW 55 USED TO BE 90 PER MCARD VERVE ATM WD





    End Sub


    Private Sub INSERT_ROU()

        'SQL
        Dim SQLCONNECTION2ROU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ROU As New System.Data.SqlClient.SqlCommand
        Dim CMD3ROU As New System.Data.SqlClient.SqlCommand
        Dim CMD4ROU As New System.Data.SqlClient.SqlCommand
        Dim CMD5ROU As New System.Data.SqlClient.SqlCommand
        Dim CMD6ROU As New System.Data.SqlClient.SqlCommand
        Dim VA As String
        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        'Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        Dim CMNT As String
        Dim CMNT1 As String
        'Dim V7 As Double

        VA = "NGNEX99963130012"
        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "C"
        'V3 = amount7
        V4 = "ROU PAYBLE" & " " & DV1.ToShortDateString
        V4a = "ROU FEES" & " " & DV1.ToShortDateString
        V4b = "MASTERCARD ROU FEES" & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "D"
        'V7 = amount
        SBU = "96N"
        CMNT = "THE ROU REPORT IS REQUIRED FOR MASTERCARD AND CASHCARD"
        CMNT1 = "CONFIRM THE MASTERCARD ROW COUNT FROM ROU PDF REPORT" & "  " & Mcard_rou_count.ToString


        'INSERT AP (ROU)
        SQLCONNECTION2ROU.Open()
        CMD2ROU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V6, amount6, V4, V5, CMNT, DV1)
        CMD2ROU.Connection = SQLCONNECTION2ROU
        CMD2ROU.ExecuteNonQuery()
        SQLCONNECTION2ROU.Close()

        'INSERT IFP (ROU) 'ATM WITHDRAWAL (VERVE BILLING)
        'SQLCONNECTION2ROU.Open()
        'CMD3ROU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V6, fee10, V4, V5, SBU, DV1)
        'CMD3ROU.Connection = SQLCONNECTION2ROU
        'CMD3ROU.ExecuteNonQuery()
        'SQLCONNECTION2ROU.Close()

        'INSERT ROU COUNT*65
        SQLCONNECTION2ROU.Open()
        CMD4ROU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, amount7, V4a, V5, DV1)
        CMD4ROU.Connection = SQLCONNECTION2ROU
        CMD4ROU.ExecuteNonQuery()
        SQLCONNECTION2ROU.Close()




        'INSERT ROU REDUCTION (ROU COUNT *15)
        'SQLCONNECTION2ROU.Open()
        'CMD5ROU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V2, fee11, V4, V5, SBU, DV1)
        'CMD5ROU.Connection = SQLCONNECTION2ROU
        'CMD5ROU.ExecuteNonQuery()
        'SQLCONNECTION2ROU.Close()



        'INSERT ROU MASTERCARD EXPENSE
        SQLCONNECTION2ROU.Open()
        CMD6ROU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", VA, V6, fee12, V4b, V5, SBU, CMNT1, DV1)
        CMD6ROU.Connection = SQLCONNECTION2ROU
        CMD6ROU.ExecuteNonQuery()
        SQLCONNECTION2ROU.Close()







    End Sub


    Private Sub CALL_NOU()

        'NOU
        'SQL 
        Dim SQLCONNECTION1NOU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1NOU As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1NOU As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"ATM WITHDRAWAL (REGULAR)"

        account = "ATM WITHDRAWAL (REGULAR)"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD1NOU.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1NOU.Open()
        CMD1NOU.Connection = SQLCONNECTION1NOU
        SQLREADER1NOU = CMD1NOU.ExecuteReader


        If SQLREADER1NOU.HasRows Then
            While (SQLREADER1NOU.Read)
                amount4 = SQLREADER1NOU.GetValue(SQLREADER1NOU.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1NOU.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''


        ''''''''''''''''''''''''''''START AFP'''''''''''''''''''''''''''''''
        'SQL 
        Dim SQLCONNECTION2NOU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2NOU As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2NOU As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (REGULAR)"

        account = "ATM WITHDRAWAL (VERVE BILLING)"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE PAYABLE"
        CMD2NOU.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2NOU.Open()
        CMD2NOU.Connection = SQLCONNECTION2NOU
        SQLREADER2NOU = CMD2NOU.ExecuteReader


        If SQLREADER2NOU.HasRows Then
            While (SQLREADER2NOU.Read)
                fee7 = SQLREADER2NOU.GetValue(SQLREADER2NOU.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2NOU.Close()
        '''''''''''''''''''''''''''''''AFP END''''''''''''''''''''''''

        'ACQUIRER FEE RECEIVABLE

        'SQL 
        Dim SQLCONNECTION3NOU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3NOU As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3NOU As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (REGULAR)"

        account = "ATM WITHDRAWAL (REGULAR)"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"
        CMD3NOU.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3NOU.Open()
        CMD3NOU.Connection = SQLCONNECTION3NOU
        SQLREADER3NOU = CMD3NOU.ExecuteReader


        If SQLREADER3NOU.HasRows Then
            While (SQLREADER3NOU.Read)
                fee8 = SQLREADER3NOU.GetValue(SQLREADER3NOU.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3NOU.Close()

        'NOU REDUCTION
        'Dim V9 As Single
        'Dim V11 As Single
        'V9 = -1 * (fee7 / 6.5)  'NOU COUNT
        'V11 = fee8 / 55 'FEE7 IS A NEGATIVE VALUE
        'fee9 = (V11 - V9) * 55 '

        Dim V10 As Double
        V10 = fee8 + (fee7 * 1) 'FEE7 IS A NEGATIVE VALUE
        amount5 = V10 '(NET OF AFR-AFP(verve billing))








    End Sub

    Private Sub NOU_INSERT()

        'SQL
        Dim SQLCONNECTION2NOU As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2NOU As New System.Data.SqlClient.SqlCommand
        Dim CMD3NOU As New System.Data.SqlClient.SqlCommand
        Dim CMD4NOU As New System.Data.SqlClient.SqlCommand
        Dim CMD5NOU As New System.Data.SqlClient.SqlCommand
        Dim VA As String
        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V4a As String
        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        'Dim V7 As Double

        VA = "NGNAS00116150028"
        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V2 = "C"
        V3 = amount4
        V4 = "NOU RCVBLE" & " " & DV1.ToShortDateString
        V4a = "NOU FEES" & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "D"
        'V7 = amount
        SBU = "96N"

        'INSERT AR (NOU)
        SQLCONNECTION2NOU.Open()
        CMD2NOU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", VA, V2, V3, V4, V5, DV1)
        CMD2NOU.Connection = SQLCONNECTION2NOU
        CMD2NOU.ExecuteNonQuery()
        SQLCONNECTION2NOU.Close()

        'INSERT AFP (NOU) 'ATM WITHDRAWAL (VERVE BILLING)
        'SQLCONNECTION2NOU.Open()
        'CMD3NOU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, fee7, V4, V5, SBU, DV1)
        'CMD3NOU.Connection = SQLCONNECTION2NOU
        'CMD3NOU.ExecuteNonQuery()
        'SQLCONNECTION2NOU.Close()

        'INSERT NET AFR (THIS IS LESS AFP(VERVE BILLING)
        SQLCONNECTION2NOU.Open()
        CMD4NOU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", VA, V2, amount5, V4a, V5, DV1)
        CMD4NOU.Connection = SQLCONNECTION2NOU
        CMD4NOU.ExecuteNonQuery()
        SQLCONNECTION2NOU.Close()




        'INSERT NOU REDUCTION (NOU COUNT *20)
        'SQLCONNECTION2NOU.Open()
        'CMD5NOU.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, fee9, V4, V5, DV1)
        'CMD5NOU.Connection = SQLCONNECTION2NOU
        'CMD5NOU.ExecuteNonQuery()
        'SQLCONNECTION2NOU.Close()

    End Sub

    Private Sub CALL_UPPERLINK_TRANSFERS()

        'SQL 
        Dim SQLCONNECTION2 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2 As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        'ATM TRANSFERS

        account = "UPPERLINK TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        Try
            SQLCONNECTION2.Open()
            CMD2.Connection = SQLCONNECTION2
            SQLREADER2 = CMD2.ExecuteReader


            If SQLREADER2.HasRows Then
                While (SQLREADER2.Read)
                    AMOUNT30 = SQLREADER2.GetValue(SQLREADER2.GetOrdinal("AMT")).ToString()
                End While
            End If
        Catch EX As Exception
            MsgBox("There is no Upper Link Transfers")
            SQLCONNECTION2.Close()
            Exit Sub
        End Try

        SQLCONNECTION2.Close()


    End Sub

    Private Sub UPPERLINK_INSERT()


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
        V1 = "NGNLI00124430030"
        V2 = "C"
        'V3 = fee
        V4 = "UPPERLINK TRSFR SETTLEMENT" & " " & DV1.ToShortDateString
        V5 = "999"
        'V6 = "D"
        V7 = AMOUNT30



        'INSERT AR (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()

    End Sub


    Private Sub CALL_RELATIONAL_TRANSFER()


        'SQL 
        Dim SQLCONNECTION2RT As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2RT As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2RT As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        'ATM TRANSFERS

        account = "RELATIONAL TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2RT.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2RT.Open()
        CMD2RT.Connection = SQLCONNECTION2RT
        SQLREADER2RT = CMD2RT.ExecuteReader


        If SQLREADER2RT.HasRows Then
            GoTo 1
        Else
            SQLCONNECTION2RT.Close()
            AMOUNT31 = 0
            Exit Sub
        End If


1:      While (SQLREADER2RT.Read)
            AMOUNT31 = SQLREADER2RT.GetValue(SQLREADER2RT.GetOrdinal("AMT")).ToString()
        End While
        SQLCONNECTION2RT.Close()







    End Sub


    Private Sub RELATIONAL_TRANSFERS_INSERT()

        'SQL
        Dim SQLCONNECTION3RT As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3RT As New System.Data.SqlClient.SqlCommand


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
        V2 = "C"
        'V3 = fee
        V4 = "RELATIONAL TRSFR SETTLEMENT" & " " & DV1.ToShortDateString
        V5 = "999"
        'V6 = "D"
        V7 = AMOUNT31



        'INSERT AR (ATM TRANSFERS)
        SQLCONNECTION3RT.Open()
        CMD3RT.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, V7, V4, V5, DV1)
        CMD3RT.Connection = SQLCONNECTION3RT
        CMD3RT.ExecuteNonQuery()
        SQLCONNECTION3RT.Close()




    End Sub


    Private Sub CALL_MOBILITY()

        'SQL 
        Dim SQLCONNECTION2 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2 As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        'ATM TRANSFERS

        account = "ATM TRANSFERS"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"
        CMD2.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2.Open()
        CMD2.Connection = SQLCONNECTION2
        SQLREADER2 = CMD2.ExecuteReader


        If SQLREADER2.HasRows Then
            While (SQLREADER2.Read)
                fee = SQLREADER2.GetValue(SQLREADER2.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2.Close()

        ''''''''''''''''''''''''''''END AFR'''''''''''''''''''''''''''''''



        ''''''''''''''''START AP''''''''''''''''''

        Dim SQLCONNECTION2AP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2AP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2AP As System.Data.SqlClient.SqlDataReader


        'ATM TRANSFERS

        account = "ATM TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD2AP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2AP.Open()
        CMD2AP.Connection = SQLCONNECTION2AP
        SQLREADER2AP = CMD2AP.ExecuteReader


        If SQLREADER2AP.HasRows Then
            While (SQLREADER2AP.Read)
                amount = SQLREADER2AP.GetValue(SQLREADER2AP.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2AP.Close()

        '''''''''''''''END AP''''''''''''''''''''''''''''




        ''''''''''''''''START AR''''''''''''''''''

        Dim SQLCONNECTION2AR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2AR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2AR As System.Data.SqlClient.SqlDataReader


        'ATM TRANSFERS

        account = "ATM TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2AR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2AR.Open()
        CMD2AR.Connection = SQLCONNECTION2AR
        SQLREADER2AR = CMD2AR.ExecuteReader


        If SQLREADER2AR.HasRows Then
            While (SQLREADER2AR.Read)
                amount1 = SQLREADER2AR.GetValue(SQLREADER2AR.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2AR.Close()

        ''''''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''


        '''''''''''''''START IFP''''''''''''''''''


        Dim SQLCONNECTION2IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFP As System.Data.SqlClient.SqlDataReader


        'ATM TRANSFERS

        account = "ATM TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD2IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFP.Open()
        CMD2IFP.Connection = SQLCONNECTION2IFP
        SQLREADER2IFP = CMD2IFP.ExecuteReader


        If SQLREADER2IFP.HasRows Then
            While (SQLREADER2IFP.Read)
                fee1 = SQLREADER2IFP.GetValue(SQLREADER2IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION2IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFR As System.Data.SqlClient.SqlDataReader


        'ATM TRANSFERS

        account = "ATM TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD2IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFR.Open()
        CMD2IFR.Connection = SQLCONNECTION2IFR
        SQLREADER2IFR = CMD2IFR.ExecuteReader


        If SQLREADER2IFR.HasRows Then
            While (SQLREADER2IFR.Read)
                fee2 = SQLREADER2IFR.GetValue(SQLREADER2IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFR.Close()


        ''''''''''''''''''''''''''''''''''END IFR''''''''''''''''''''''''''''''''

        '''''''''''''''''''''''''''CARDHOLDER ITEMS START HERE'''''''''''''''''''''''''''''''''''''


        '''''''''''''''START IFP''''''''''''''''''


        Dim SQLCONNECTION3IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3IFP As System.Data.SqlClient.SqlDataReader


        'CARDHOLDER

        account = "CARD HOLDER ACCOUNT TRANSFER"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD3IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3IFP.Open()
        CMD3IFP.Connection = SQLCONNECTION3IFP
        SQLREADER3IFP = CMD3IFP.ExecuteReader


        If SQLREADER3IFP.HasRows Then
            While (SQLREADER3IFP.Read)
                fee3 = SQLREADER3IFP.GetValue(SQLREADER3IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION3IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3IFR As System.Data.SqlClient.SqlDataReader


        'CARDHOLDER

        account = "CARD HOLDER ACCOUNT TRANSFER"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD3IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3IFR.Open()
        CMD3IFR.Connection = SQLCONNECTION3IFR
        SQLREADER3IFR = CMD3IFR.ExecuteReader


        If SQLREADER3IFR.HasRows Then
            While (SQLREADER3IFR.Read)
                fee4 = SQLREADER3IFR.GetValue(SQLREADER3IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3IFR.Close()


        ''''''''''''''''''''''''''''''CARDHOLDER ENDS HERE''''''''''''''''''''''''''



        '''''''''''''''''''''''QUICKTELLER MOBILE STARTS HERE''''''''''''''''''''''''''''
        'QUICKTELLER MOBILE AP


        Dim SQLCONNECTION4AP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4AP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4AP As System.Data.SqlClient.SqlDataReader




        account = "QUICKTELLER MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD4AP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4AP.Open()
        CMD4AP.Connection = SQLCONNECTION4AP
        SQLREADER4AP = CMD4AP.ExecuteReader


        If SQLREADER4AP.HasRows Then
            While (SQLREADER4AP.Read)
                amount2 = SQLREADER4AP.GetValue(SQLREADER4AP.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION4AP.Close()

        '''''''''''''''END AP''''''''''''''''''''''''''''




        ''''''''''''''''START AR''''''''''''''''''

        Dim SQLCONNECTION4AR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4AR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4AR As System.Data.SqlClient.SqlDataReader


        '"QUICKTELLER MOBILE TRANSFERS" AR

        account = "QUICKTELLER MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD4AR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4AR.Open()
        CMD4AR.Connection = SQLCONNECTION4AR
        SQLREADER4AR = CMD4AR.ExecuteReader


        If SQLREADER4AR.HasRows Then
            While (SQLREADER4AR.Read)
                amount3 = SQLREADER4AR.GetValue(SQLREADER4AR.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION4AR.Close()


        ''''''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''

        'QT_MOBILE_TRANSFER FEES

        Dim SQLCONNECTION4IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4IFP As System.Data.SqlClient.SqlDataReader




        account = "QUICKTELLER MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD4IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4IFP.Open()
        CMD4IFP.Connection = SQLCONNECTION4IFP
        SQLREADER4IFP = CMD4IFP.ExecuteReader


        If SQLREADER4IFP.HasRows Then
            While (SQLREADER4IFP.Read)
                fee5 = SQLREADER4IFP.GetValue(SQLREADER4IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION4IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4IFR As System.Data.SqlClient.SqlDataReader


        'ATM TRANSFERS

        account = "QUICKTELLER MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD4IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4IFR.Open()
        CMD4IFR.Connection = SQLCONNECTION4IFR
        SQLREADER4IFR = CMD4IFR.ExecuteReader


        If SQLREADER4IFR.HasRows Then
            While (SQLREADER4IFR.Read)
                fee6 = SQLREADER4IFR.GetValue(SQLREADER4IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4IFR.Close()


        ''''''''''''''''''''''''''''''''''END IFR''''''''''''''''''''''''''''''''













        ''''''''''''''''''''''''''''''''END QUICKTELLERMOBILE REPORTS HERE'''''''''''''''''''''''''





        ''''''''''START MOBILE TRANSFERS HERE




        'MOBILE_TRANSFER FEES

        Dim SQLCONNECTION5IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5IFP As System.Data.SqlClient.SqlDataReader




        account = "MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD5IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION5IFP.Open()
        CMD5IFP.Connection = SQLCONNECTION5IFP
        SQLREADER5IFP = CMD5IFP.ExecuteReader


        If SQLREADER5IFP.HasRows Then
            While (SQLREADER5IFP.Read)
                fee_m1 = SQLREADER5IFP.GetValue(SQLREADER5IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION5IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION5IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5IFR As System.Data.SqlClient.SqlDataReader


        'MOBILE_TRANSFER Fees

        account = "MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD5IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION5IFR.Open()
        CMD5IFR.Connection = SQLCONNECTION5IFR
        SQLREADER5IFR = CMD5IFR.ExecuteReader


        If SQLREADER5IFR.HasRows Then
            While (SQLREADER5IFR.Read)
                fee_m2 = SQLREADER5IFR.GetValue(SQLREADER5IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION5IFR.Close()

        ''''''''''''''''''''''''''''''''''''''''''''''''
        'MOBILE_TRANSFER AMOUNTS

        Dim SQLCONNECTION5AP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5AP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5AP As System.Data.SqlClient.SqlDataReader




        account = "MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD5AP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION5AP.Open()
        CMD5AP.Connection = SQLCONNECTION5AP
        SQLREADER5AP = CMD5AP.ExecuteReader


        If SQLREADER5AP.HasRows Then
            While (SQLREADER5AP.Read)
                amount_m1 = SQLREADER5AP.GetValue(SQLREADER5AP.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION5AP.Close()


        ''''''''''''''''''''''''''''''''''END AP''''''''''''''''''''''''''''''''




        '''''''''''''''START AR''''''''''''''''''


        Dim SQLCONNECTION5AR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5AR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5AR As System.Data.SqlClient.SqlDataReader


        'MOBILE_TRANSFER Fees

        account = "MOBILE TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD5AR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION5AR.Open()
        CMD5AR.Connection = SQLCONNECTION5AR
        SQLREADER5AR = CMD5AR.ExecuteReader


        If SQLREADER5AR.HasRows Then
            While (SQLREADER5AR.Read)
                amount_m2 = SQLREADER5AR.GetValue(SQLREADER5AR.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION5AR.Close()


        ''''END AR'''''''''''''''''''''''


        ''''''''''''''''''''END MOBILE TRANSFERS HERE














    End Sub


    Private Sub MOBILITY_Insert()

        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand
        Dim CMD3A As New System.Data.SqlClient.SqlCommand
        Dim CMD3B As New System.Data.SqlClient.SqlCommand
        Dim CMD3C As New System.Data.SqlClient.SqlCommand
        Dim CMD3D As New System.Data.SqlClient.SqlCommand
        Dim CMD3E As New System.Data.SqlClient.SqlCommand
        Dim CMD3F As New System.Data.SqlClient.SqlCommand
        Dim CMD3G As New System.Data.SqlClient.SqlCommand
        Dim CMD3H As New System.Data.SqlClient.SqlCommand
        Dim CMD3I As New System.Data.SqlClient.SqlCommand
        Dim CMD3J As New System.Data.SqlClient.SqlCommand
        Dim CMD3K As New System.Data.SqlClient.SqlCommand
        Dim CMD3L As New System.Data.SqlClient.SqlCommand
        Dim CMD3M As New System.Data.SqlClient.SqlCommand



        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V1A As String
        Dim V2 As String  'STATUS
        Dim V3 As Double  'FEE
        Dim V4a As String
        Dim V4b As String
        Dim V4c As String
        Dim V4d As String
        Dim V4e As String
        Dim V4f As String
        Dim V4g As String
        Dim V4h As String
        Dim V4i As String
        Dim V4j As String
        Dim V4k As String
        Dim V4l As String
        Dim V4m As String
        Dim V4n As String

        Dim V5 As String
        Dim V6 As String
        Dim V7 As Double

        Dim SBU As String
        Dim SBU1 As String

        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V1A = "NGNIN99954210036"
        V2 = "C"
        V3 = fee
        V4a = "ATM TRNSFR SETTL RCV" & " " & DV1.ToShortDateString
        V4b = "ATM TRNSFR ACQ FEE" & " " & DV1.ToShortDateString
        V4c = "ATM TRNSFR SETTL PYBL" & " " & DV1.ToShortDateString
        V4d = "ATM TRNSFR ISS FEE PYBLE" & " " & DV1.ToShortDateString
        V4e = "ATM TRNSFR ISS FEE RCVBLE" & " " & DV1.ToShortDateString
        V4f = "MBLTY CARDHOLDER FEES" & " " & DV1.ToShortDateString
        V4g = "QTELLER MOBILE PYBLE" & " " & DV1.ToShortDateString
        V4h = "QTELLER MOBILE RCVBL" & " " & DV1.ToShortDateString
        V4i = "QTELLER MOBILE ISS FEE PYBL" & " " & DV1.ToShortDateString
        V4j = "QTELLER MOBILE ISS FEE RCVBL" & " " & DV1.ToShortDateString
        V4k = "MOBILE TRSF PYBLE" & " " & DV1.ToShortDateString
        V4l = "MOBILE TRSF RCVBL" & " " & DV1.ToShortDateString
        V4m = "MOBILE TRSF ISS FEE PYBL" & " " & DV1.ToShortDateString
        V4n = "MOBILE TRSF ISS FEE RCVBL" & " " & DV1.ToShortDateString



        V5 = "999"
        V6 = "D"
        V7 = amount
        SBU = "96N"
        SBU1 = "94L"


        'INSERT AFR (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1A, V2, V3, V4b, V5, SBU1, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        'INSERT AP (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3A.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, V7, V4c, V5, DV1)
        CMD3A.Connection = SQLCONNECTION3
        CMD3A.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT AR (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3B.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount1, V4a, V5, DV1)
        CMD3B.Connection = SQLCONNECTION3
        CMD3B.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        'INSERT IFP (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3C.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, fee1, V4d, V5, SBU, DV1)
        CMD3C.Connection = SQLCONNECTION3
        CMD3C.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFR (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3D.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1A, V2, fee2, V4e, V5, SBU, DV1)
        CMD3D.Connection = SQLCONNECTION3
        CMD3D.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''MOBILITY CARDHOLDER INSERTING''''''''''''''''''''''''''''''''''''''

        'INSERT NET FEES IFP&IFR (CARDHOLDER) 
        Dim V8 As Double
        'SBU1 IS 94L
        V8 = fee3 + fee4 'NET FEE
        SQLCONNECTION3.Open()
        CMD3E.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, V8, V4f, V5, SBU1, DV1)
        CMD3E.Connection = SQLCONNECTION3
        CMD3E.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        ''''''''''''''''''''''''''''''''''MOBILITY CARDHOLDER INSERTING ENDS HERE''''''''''''''''''''''''''''''''''''''




        ''''''''''''QUICKTELLER MOBILE INSERT STARTS HERE'''''''''''''''''''''''''


        'INSERT AP (QT MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3F.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, amount2, V4g, V5, DV1)
        CMD3F.Connection = SQLCONNECTION3
        CMD3F.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT AR (QT MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3G.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount3, V4h, V5, DV1)
        CMD3G.Connection = SQLCONNECTION3
        CMD3G.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFP (QT MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3H.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, fee5, V4i, V5, DV1)
        CMD3H.Connection = SQLCONNECTION3
        CMD3H.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFR (QT MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3I.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, fee6, V4j, V5, DV1)
        CMD3I.Connection = SQLCONNECTION3
        CMD3I.ExecuteNonQuery()
        SQLCONNECTION3.Close()



        'INSERT AP (MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3J.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, amount_m1, V4k, V5, DV1)
        CMD3J.Connection = SQLCONNECTION3
        CMD3J.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT AR (MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3K.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, amount_m2, V4l, V5, DV1)
        CMD3K.Connection = SQLCONNECTION3
        CMD3K.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFP (MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3L.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V6, fee_m1, V4m, V5, DV1)
        CMD3L.Connection = SQLCONNECTION3
        CMD3L.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFR (MOBILE TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3M.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V1, V2, fee_m2, V4n, V5, DV1)
        CMD3M.Connection = SQLCONNECTION3
        CMD3M.ExecuteNonQuery()
        SQLCONNECTION3.Close()








    End Sub


    Private Sub CALL_BILLPAYMENT()




        'SQL 
        Dim SQLCONNECTION2 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2 As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        'BILLPAYMENT

        account = "BILLPAYMENT"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"
        CMD2.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2.Open()
        CMD2.Connection = SQLCONNECTION2
        SQLREADER2 = CMD2.ExecuteReader


        If SQLREADER2.HasRows Then
            While (SQLREADER2.Read)
                bill_afr = SQLREADER2.GetValue(SQLREADER2.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2.Close()

        ''''''''''''''''''''''''''''END AFR'''''''''''''''''''''''''''''''










        '''''''''''''''START IFP''''''''''''''''''


        Dim SQLCONNECTION2IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFP As System.Data.SqlClient.SqlDataReader


        'BILLPAYMENT

        account = "BILLPAYMENT"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD2IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFP.Open()
        CMD2IFP.Connection = SQLCONNECTION2IFP
        SQLREADER2IFP = CMD2IFP.ExecuteReader


        If SQLREADER2IFP.HasRows Then
            While (SQLREADER2IFP.Read)
                bill_ifp = SQLREADER2IFP.GetValue(SQLREADER2IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION2IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFR As System.Data.SqlClient.SqlDataReader


        'BILLPAYMENT

        account = "BILLPAYMENT"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD2IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFR.Open()
        CMD2IFR.Connection = SQLCONNECTION2IFR
        SQLREADER2IFR = CMD2IFR.ExecuteReader


        If SQLREADER2IFR.HasRows Then
            While (SQLREADER2IFR.Read)
                bill_ifr = SQLREADER2IFR.GetValue(SQLREADER2IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFR.Close()



        ''''''''''''''''''''''''''''''''''END IFR''''''''''''''''''''''''''''''''





        '''''''''''''''START SVA_FR''''''''''''''''''


        Dim SQLCONNECTION2SVA As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2SVA As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2SVA As System.Data.SqlClient.SqlDataReader


        'BILLPAYMENT

        account = "BILLPAYMENT"
        status = "CURRENT"
        acct_type = "SVA FEE RECEIVABLE"
        CMD2SVA.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2SVA.Open()
        CMD2SVA.Connection = SQLCONNECTION2SVA
        SQLREADER2SVA = CMD2SVA.ExecuteReader


        If SQLREADER2SVA.HasRows Then
            While (SQLREADER2SVA.Read)
                bill_sva_fr = SQLREADER2SVA.GetValue(SQLREADER2SVA.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2SVA.Close()



        ''''''''''''''''''''''''''''''''''END SVA_FR''''''''''''''''''''''''''''''''

    End Sub



    Private Sub BILLPAYMENT_INSERT()

        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")

        Dim CMDB1 As New System.Data.SqlClient.SqlCommand
        Dim CMDB2 As New System.Data.SqlClient.SqlCommand
        Dim CMDB3 As New System.Data.SqlClient.SqlCommand
        Dim CMDB4 As New System.Data.SqlClient.SqlCommand

        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V3 As Double  'FEE
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V4c As String


        Dim V5 As String
        Dim V6 As String
        'Dim V7 As Double
        Dim SBU1 As String
        Dim SBU2 As String
        Dim SBU3 As String

        V0 = "NGNEX99963130009"
        V1 = "NGNIN99954210036"
        V2 = "C"
        V3 = fee
        V4 = "BILLS PYMNT ACQ FEE RCV" & " " & DV1.ToShortDateString
        V4a = "BILLS PYMNT ISS FEE PYBL" & " " & DV1.ToShortDateString
        V4b = "BILLS PYMNT ISS FEE RCVBL" & " " & DV1.ToShortDateString
        V4c = "SVA FEE RCVBLE" & " " & DV1.ToShortDateString

        V5 = "999"
        V6 = "D"

        SBU1 = "94L"
        SBU2 = "96N"
        SBU3 = "90A"



        'INSERT AFR (BILLPAYMENT)
        SQLCONNECTION3.Open()
        CMDB1.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, bill_afr, V4, V5, SBU1, DV1)
        CMDB1.Connection = SQLCONNECTION3
        CMDB1.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        'INSERT IFP (BILLPAYMENT)
        SQLCONNECTION3.Open()
        CMDB2.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, bill_ifp, V4a, V5, SBU2, DV1)
        CMDB2.Connection = SQLCONNECTION3
        CMDB2.ExecuteNonQuery()
        SQLCONNECTION3.Close()



        'INSERT IFR (BILLPAYMENT)
        SQLCONNECTION3.Open()
        CMDB3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, bill_ifr, V4b, V5, SBU2, DV1)
        CMDB3.Connection = SQLCONNECTION3
        CMDB3.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT SVA_FR (BILLPAYMENT)
        SQLCONNECTION3.Open()
        CMDB4.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, bill_sva_fr, V4c, V5, SBU3, DV1)
        CMDB4.Connection = SQLCONNECTION3
        CMDB4.ExecuteNonQuery()
        SQLCONNECTION3.Close()








    End Sub


    Private Sub CALL_VERVE_INTERNATIONAL()



        'SQL 
        Dim SQLCONNECTION2VR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2VR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2VR As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        'VERVE INTL TRANSFERS

        account = "VERVE INTL TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2VR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2VR.Open()
        CMD2VR.Connection = SQLCONNECTION2VR
        SQLREADER2VR = CMD2VR.ExecuteReader


        If SQLREADER2VR.HasRows Then
            While (SQLREADER2VR.Read)
                vit_ar = SQLREADER2VR.GetValue(SQLREADER2VR.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2VR.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''


        Dim SQLCONNECTION2VP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2VP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2VP As System.Data.SqlClient.SqlDataReader

        

        'VERVE INTL TRANSFERS

        account = "VERVE INTL TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD2VP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2VP.Open()
        CMD2VP.Connection = SQLCONNECTION2VP
        SQLREADER2VP = CMD2VP.ExecuteReader


        If SQLREADER2VP.HasRows Then
            While (SQLREADER2VP.Read)
                vit_ap = SQLREADER2VP.GetValue(SQLREADER2VP.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2VP.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''







        '''''''''''''''START IFP''''''''''''''''''


        Dim SQLCONNECTION2IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFP As System.Data.SqlClient.SqlDataReader


        'VERVE INTL TRANSFERS

        account = "VERVE INTL TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD2IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFP.Open()
        CMD2IFP.Connection = SQLCONNECTION2IFP
        SQLREADER2IFP = CMD2IFP.ExecuteReader


        If SQLREADER2IFP.HasRows Then
            While (SQLREADER2IFP.Read)
                VIT_ifp = SQLREADER2IFP.GetValue(SQLREADER2IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION2IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFR As System.Data.SqlClient.SqlDataReader


        'VERVE INTL TRANSFERS

        account = "VERVE INTL TRANSFERS"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD2IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFR.Open()
        CMD2IFR.Connection = SQLCONNECTION2IFR
        SQLREADER2IFR = CMD2IFR.ExecuteReader


        If SQLREADER2IFR.HasRows Then
            While (SQLREADER2IFR.Read)
                VIT_ifr = SQLREADER2IFR.GetValue(SQLREADER2IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFR.Close()


    End Sub


    Private Sub VERVE_INTERNATIONAL_INSERT()

        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")

        Dim CMDB1 As New System.Data.SqlClient.SqlCommand
        Dim CMDB2 As New System.Data.SqlClient.SqlCommand
        Dim CMDB3 As New System.Data.SqlClient.SqlCommand
        Dim CMDB4 As New System.Data.SqlClient.SqlCommand

        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V3 As Double  'FEE
        Dim V4 As String
        Dim v4a As String
        Dim v4b As String
        Dim v4c As String
        Dim V5 As String
        Dim V6 As String
        Dim V7 As String
        'Dim SBU1 As String
        Dim SBU2 As String
        'Dim SBU3 As String

        V0 = "NGNLI00124430030"
        V1 = "NGNIN99954210036"
        V7 = "NGNEX99963130009"
        V2 = "C"
        V3 = fee
        V4 = "VERVE INT RCVBL" & " " & DV1.ToShortDateString
        v4a = "VERVE INT PYBL" & " " & DV1.ToShortDateString
        v4b = "VERVE INTNL ISS FEE PYBL" & " " & DV1.ToShortDateString
        v4c = "VERVE INTNL ISS FEE RCVBL" & " " & DV1.ToShortDateString
        V5 = "999"
        V6 = "D"

        'SBU1 = "94L"
        SBU2 = "96N"
        'SBU3 = "90A"



        'INSERT AR (VERVE INTL TRANSFERS)
        SQLCONNECTION3.Open()
        CMDB1.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V0, V2, vit_ar, V4, V5, DV1)
        CMDB1.Connection = SQLCONNECTION3
        CMDB1.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        'INSERT AP (VERVE INTL TRANSFERS)
        SQLCONNECTION3.Open()
        CMDB2.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V0, V6, vit_ap, v4a, V5, DV1)
        CMDB2.Connection = SQLCONNECTION3
        CMDB2.ExecuteNonQuery()
        SQLCONNECTION3.Close()



        'INSERT IFP (VERVE INTL TRANSFERS)
        SQLCONNECTION3.Open()
        CMDB3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V7, V6, VIT_ifp, v4b, V5, SBU2, DV1)
        CMDB3.Connection = SQLCONNECTION3
        CMDB3.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFR (VERVE INTL TRANSFERS)
        SQLCONNECTION3.Open()
        CMDB4.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, VIT_ifr, v4c, V5, SBU2, DV1)
        CMDB4.Connection = SQLCONNECTION3
        CMDB4.ExecuteNonQuery()
        SQLCONNECTION3.Close()






    End Sub


    Private Sub CALL_PREPAID_CARDLOAD()

        'SQL 
        Dim SQLCONNECTION2PPC_AP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2PPC_AP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2PPC_AP As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        'PREPAID CARDLOAD

        account = "PREPAID CARDLOAD"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD2PPC_AP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2PPC_AP.Open()
        CMD2PPC_AP.Connection = SQLCONNECTION2PPC_AP
        SQLREADER2PPC_AP = CMD2PPC_AP.ExecuteReader


        If SQLREADER2PPC_AP.HasRows Then
            While (SQLREADER2PPC_AP.Read)
                PPC_aP = SQLREADER2PPC_AP.GetValue(SQLREADER2PPC_AP.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2PPC_AP.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''




        Dim SQLCONNECTION2PPC_AR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2 As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2PPC_AR As System.Data.SqlClient.SqlDataReader



        'PREPAID CARDLOAD

        account = "PREPAID CARDLOAD"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2PPC_AR.Open()
        CMD2.Connection = SQLCONNECTION2PPC_AR
        SQLREADER2PPC_AR = CMD2.ExecuteReader


        If SQLREADER2PPC_AR.HasRows Then
            While (SQLREADER2PPC_AR.Read)
                PPC_aR = SQLREADER2PPC_AR.GetValue(SQLREADER2PPC_AR.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2PPC_AR.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''







        '''''''''''''''START IFP''''''''''''''''''


        Dim SQLCONNECTION2IFP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFP As System.Data.SqlClient.SqlDataReader


        'PREPAID CARDLOAD

        account = "PREPAID CARDLOAD"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD2IFP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFP.Open()
        CMD2IFP.Connection = SQLCONNECTION2IFP
        SQLREADER2IFP = CMD2IFP.ExecuteReader


        If SQLREADER2IFP.HasRows Then
            While (SQLREADER2IFP.Read)
                PPC_ifp = SQLREADER2IFP.GetValue(SQLREADER2IFP.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFP.Close()


        ''''''''''''''''''''''''''''''''''END IFP''''''''''''''''''''''''''''''''




        '''''''''''''''START IFR''''''''''''''''''


        Dim SQLCONNECTION2IFR As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2IFR As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2IFR As System.Data.SqlClient.SqlDataReader


        'PREPAID CARDLOAD

        account = "PREPAID CARDLOAD"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD2IFR.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2IFR.Open()
        CMD2IFR.Connection = SQLCONNECTION2IFR
        SQLREADER2IFR = CMD2IFR.ExecuteReader


        If SQLREADER2IFR.HasRows Then
            While (SQLREADER2IFR.Read)
                PPC_ifr = SQLREADER2IFR.GetValue(SQLREADER2IFR.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION2IFR.Close()



        ''''''''''''''''''''''''''''''''''END IFR''''''''''''''''''''''''''''''''





    End Sub



    Private Sub INSERT_PREPAID_CARD()

        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")

        Dim CMDB1 As New System.Data.SqlClient.SqlCommand
        Dim CMDB2 As New System.Data.SqlClient.SqlCommand
        Dim CMDB3 As New System.Data.SqlClient.SqlCommand
        Dim CMDB4 As New System.Data.SqlClient.SqlCommand

        Dim V0 As String
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V3 As Double  'FEE
        Dim V4 As String
        Dim v4a As String
        Dim v4b As String
        Dim v4c As String
        Dim V5 As String
        Dim V6 As String
        Dim V7 As String
        'Dim SBU1 As String
        Dim SBU2 As String
        'Dim SBU3 As String

        V0 = "NGNLI99924430087"
        V1 = "NGNIN99954210036"
        V7 = "NGNEX99963130009"
        V2 = "C"
        V3 = fee
        V4 = "PREPAID CARDLOAD RCV" & " " & DV1.ToShortDateString
        v4a = "PREPAID CARDLOAD PYBL" & " " & DV1.ToShortDateString
        v4b = "PREPAID CARDLOAD ISS FEE PYBL" & " " & DV1.ToShortDateString
        v4c = "PREPAID CARDLOAD ISS FEE RCV" & " " & DV1.ToShortDateString






        V5 = "999"
        V6 = "D"

        'SBU1 = "94L"
        SBU2 = "96N"
        'SBU3 = "90A"



        'INSERT AR (PREPAID CARDLOAD)
        SQLCONNECTION3.Open()
        CMDB1.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V0, V2, PPC_aR, V4, V5, DV1)
        CMDB1.Connection = SQLCONNECTION3
        CMDB1.ExecuteNonQuery()
        SQLCONNECTION3.Close()

        'INSERT AP (PREPAID CARDLOAD)
        SQLCONNECTION3.Open()
        CMDB2.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", V0, V6, PPC_aP, v4a, V5, DV1)
        CMDB2.Connection = SQLCONNECTION3
        CMDB2.ExecuteNonQuery()
        SQLCONNECTION3.Close()



        'INSERT IFP (PREPAID CARDLOAD)
        SQLCONNECTION3.Open()
        CMDB3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V7, V6, PPC_IFP, v4b, V5, SBU2, DV1)
        CMDB3.Connection = SQLCONNECTION3
        CMDB3.ExecuteNonQuery()
        SQLCONNECTION3.Close()


        'INSERT IFR (PREPAID CARDLOAD)
        SQLCONNECTION3.Open()
        CMDB4.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, PPC_IFR, v4c, V5, SBU2, DV1)
        CMDB4.Connection = SQLCONNECTION3
        CMDB4.ExecuteNonQuery()
        SQLCONNECTION3.Close()


    End Sub


    Private Sub CALL_ATM_CARDLESS()

        'ATM WITHDRAWAL CARDLESS
        'SQL 
        Dim SQLCONNECTION1ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1ATM_CARDLESS As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"ATM WITHDRAWAL (CARDLESS)"

        account = "ATM WITHDRAWAL (CARDLESS)"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD1ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1ATM_CARDLESS.Open()
        CMD1ATM_CARDLESS.Connection = SQLCONNECTION1ATM_CARDLESS
        SQLREADER1ATM_CARDLESS = CMD1ATM_CARDLESS.ExecuteReader


        If SQLREADER1ATM_CARDLESS.HasRows Then
            While (SQLREADER1ATM_CARDLESS.Read)
                amount21 = SQLREADER1ATM_CARDLESS.GetValue(SQLREADER1ATM_CARDLESS.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1ATM_CARDLESS.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''


        '''''''''''''''''''''''''START AR''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION2ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (CARDLESS)"

        account = "ATM WITHDRAWAL (CARDLESS)"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD2ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        SQLREADER2ATM_CARDLESS = CMD2ATM_CARDLESS.ExecuteReader


        If SQLREADER2ATM_CARDLESS.HasRows Then
            While (SQLREADER2ATM_CARDLESS.Read)
                amount22 = SQLREADER2ATM_CARDLESS.GetValue(SQLREADER2ATM_CARDLESS.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2ATM_CARDLESS.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''




        '''''''''''''''''''''''''START Cardless Fees''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION3ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (CARDLESS)"

        account = "ATM WITHDRAWAL (CARDLESS)"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD3ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3ATM_CARDLESS.Open()
        CMD3ATM_CARDLESS.Connection = SQLCONNECTION3ATM_CARDLESS
        SQLREADER3ATM_CARDLESS = CMD3ATM_CARDLESS.ExecuteReader


        If SQLREADER3ATM_CARDLESS.HasRows Then
            While (SQLREADER3ATM_CARDLESS.Read)
                Fc1 = SQLREADER3ATM_CARDLESS.GetValue(SQLREADER3ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3ATM_CARDLESS.Close()




        Dim SQLCONNECTION4ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (CARDLESS)"

        account = "ATM WITHDRAWAL (CARDLESS)"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                fc2 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()


        '"ATM WITHDRAWAL (CARDLESS)"

        account = "ATM WITHDRAWAL (CARDLESS)"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"

        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                fc3 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()



        '"ATM WITHDRAWAL (CARDLESS)"

        account = "ATM WITHDRAWAL (CARDLESS)"
        status = "CURRENT"
        acct_type = "SCHEME OWNER ISSUER FEE RECEIVABLE"


        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                fc4 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()




        ''''''''''''''''''''''''''''END ATM WITHDRAWAL (CARDLESS)'''''''''''''''''''''''''''''''




    End Sub


    Private Sub CALL_ATM_CARDLESS_INSERT()



        'SQL
        Dim SQLCONNECTION2ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD3ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD4ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD5ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD6ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V4c As String
        Dim v4d As String
        Dim v4e As String


        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        Dim SBU1 As String
        Dim CMNT As String
        Dim V0 As String
        Dim V1a As String
        Dim V1b As String




        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V1a = "NGNAS00116150028"
        V1b = "NGNIN99954210036"
        V2 = "C"
        V4 = "ATM_WD CARDLESS PYBLE " & " " & DV1.ToShortDateString
        V4a = "ATM_WD CARDLESS RCVBLE" & " " & DV1.ToShortDateString
        V4b = "ATM_WD CARDLESS IFP" & " " & DV1.ToShortDateString
        V4c = "ATM_WD CARDLESS IFR" & " " & DV1.ToShortDateString
        v4d = "ATM_WD CARDLESS AFR" & " " & DV1.ToShortDateString
        v4e = "ATM_WD CARDLESS SCHM OWNR FEES IFR" & " " & DV1.ToShortDateString



        V5 = "999"
        V6 = "D"
        SBU = "96N"
        SBU1 = "94L"
        CMNT = "ALWAYS CONFIRM FROM THE ROU REPORT WHERE CARDLESS TRANSACTIONS DEFAULT TO"




        'INSERT AP (ATM_CARDLESS) 
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD3ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V6, amount21, V4, V5, CMNT, DV1)
        CMD3ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD3ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT AR (ATM_CARDLESS)
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1a, V2, amount22, V4a, V5, CMNT, DV1)
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD4ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'IFP
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD2ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, Fc1, V4b, V5, SBU, DV1)
        CMD2ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD2ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()



        'INSERT ATM_CARDLESS FEES 'IFR
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, fc2, V4c, V5, SBU, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'AFR
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, fc3, v4d, V5, SBU1, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'SCHEME OWNER FEE
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, fc4, v4e, V5, SBU, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()





    End Sub

    Private Sub CALL_ATM_CARDLESS_NON_CARD()

        'ATM WITHDRAWAL CARDLESS
        'SQL 
        Dim SQLCONNECTION1ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1ATM_CARDLESS As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"ATM WITHDRAWAL (Cardless:Non-Card Generated)"

        account = "ATM WITHDRAWAL (Cardless:Non-Card Generated)"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD1ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1ATM_CARDLESS.Open()
        CMD1ATM_CARDLESS.Connection = SQLCONNECTION1ATM_CARDLESS
        SQLREADER1ATM_CARDLESS = CMD1ATM_CARDLESS.ExecuteReader


        If SQLREADER1ATM_CARDLESS.HasRows Then
            While (SQLREADER1ATM_CARDLESS.Read)
                amount26 = SQLREADER1ATM_CARDLESS.GetValue(SQLREADER1ATM_CARDLESS.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1ATM_CARDLESS.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''


        '''''''''''''''''''''''''START AR''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION2ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (Cardless:Non-Card Generated)"

        account = "ATM WITHDRAWAL (Cardless:Non-Card Generated)"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD2ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        SQLREADER2ATM_CARDLESS = CMD2ATM_CARDLESS.ExecuteReader


        If SQLREADER2ATM_CARDLESS.HasRows Then
            While (SQLREADER2ATM_CARDLESS.Read)
                amount27 = SQLREADER2ATM_CARDLESS.GetValue(SQLREADER2ATM_CARDLESS.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2ATM_CARDLESS.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''




        '''''''''''''''''''''''''START Cardless_Non_Card Fees''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION3ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (Cardless:Non-Card Generated)

        account = "ATM WITHDRAWAL (Cardless:Non-Card Generated)"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD3ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3ATM_CARDLESS.Open()
        CMD3ATM_CARDLESS.Connection = SQLCONNECTION3ATM_CARDLESS
        SQLREADER3ATM_CARDLESS = CMD3ATM_CARDLESS.ExecuteReader


        If SQLREADER3ATM_CARDLESS.HasRows Then
            While (SQLREADER3ATM_CARDLESS.Read)
                FC9 = SQLREADER3ATM_CARDLESS.GetValue(SQLREADER3ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3ATM_CARDLESS.Close()




        Dim SQLCONNECTION4ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (Cardless:Non-Card Generated)

        account = "ATM WITHDRAWAL (Cardless:Non-Card Generated)"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                FC10 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()


        '"ATM WITHDRAWAL (Cardless:Non-Card Generated)

        account = "ATM WITHDRAWAL (Cardless:Non-Card Generated)"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"

        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                FC11 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()



        '"ATM WITHDRAWAL (Cardless:Non-Card Generated)

        account = "ATM WITHDRAWAL (Cardless:Non-Card Generated)"
        status = "CURRENT"
        acct_type = "SCHEME OWNER ISSUER FEE RECEIVABLE"


        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                FC12 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()




        ''''''''''''''''''''''''''''END ATM WITHDRAWAL (Cardless:Non-Card Generated))'''''''''''''''''''''''''''''''







    End Sub


    Private Sub CALL_ATM_CARDLESS_NON_CARD_INSERT()

        'NC IS NON-CARD
        'SQL
        Dim SQLCONNECTION2ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD3ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD4ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD5ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD6ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V4c As String
        Dim v4d As String
        Dim v4e As String


        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        Dim SBU1 As String
        Dim CMNT As String
        Dim V0 As String
        Dim V1a As String
        Dim V1b As String




        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V1a = "NGNAS00116150028"
        V1b = "NGNIN99954210036"
        V2 = "C"
        V4 = "ATM_WD CARDLESS_NC PYBLE " & " " & DV1.ToShortDateString
        V4a = "ATM_WD CARDLESS_NC RCVBLE" & " " & DV1.ToShortDateString
        V4b = "ATM_WD CARDLESS_NC IFP" & " " & DV1.ToShortDateString
        V4c = "ATM_WD CARDLESS_NC IFR" & " " & DV1.ToShortDateString
        v4d = "ATM_WD CARDLESS_NC AFR" & " " & DV1.ToShortDateString
        v4e = "ATM_WD CARDLESS_NC SCHM OWNR FEES IFR" & " " & DV1.ToShortDateString



        V5 = "999"
        V6 = "D"
        SBU = "96N"
        SBU1 = "94L"
        CMNT = "ALWAYS CONFIRM FROM THE ROU REPORT WHERE CARDLESS TRANSACTIONS DEFAULT TO"




        'INSERT AP (ATM_CARDLESS) 
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD3ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V6, amount26, V4, V5, CMNT, DV1)
        CMD3ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD3ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT AR (ATM_CARDLESS)
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1a, V2, amount27, V4a, V5, CMNT, DV1)
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD4ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'IFP
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD2ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, FC9, V4b, V5, SBU, DV1)
        CMD2ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD2ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()



        'INSERT ATM_CARDLESS FEES 'IFR
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, FC10, V4c, V5, SBU, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'AFR
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, FC11, v4d, V5, SBU1, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'SCHEME OWNER FEE
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, FC12, v4e, V5, SBU, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()





    End Sub

    Private Sub CALL_ATM_CARDLESS_VERVE_TOKEN()

        'ATM WITHDRAWAL CARDLESS
        'SQL 
        Dim SQLCONNECTION1ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1ATM_CARDLESS As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String

        '"ATM WITHDRAWAL (Cardless:Paycode Verve Token)"

        account = "ATM WITHDRAWAL (Cardless:Paycode Verve Token)"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        CMD1ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION1ATM_CARDLESS.Open()
        CMD1ATM_CARDLESS.Connection = SQLCONNECTION1ATM_CARDLESS
        SQLREADER1ATM_CARDLESS = CMD1ATM_CARDLESS.ExecuteReader


        If SQLREADER1ATM_CARDLESS.HasRows Then
            While (SQLREADER1ATM_CARDLESS.Read)
                amount23 = SQLREADER1ATM_CARDLESS.GetValue(SQLREADER1ATM_CARDLESS.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION1ATM_CARDLESS.Close()

        ''''''''''''''''''''''''''''END AP'''''''''''''''''''''''''''''''


        '''''''''''''''''''''''''START AR''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION2ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (Cardless:Paycode Verve Token)"

        account = "ATM WITHDRAWAL (Cardless:Paycode Verve Token)"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        CMD2ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD2ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        SQLREADER2ATM_CARDLESS = CMD2ATM_CARDLESS.ExecuteReader


        If SQLREADER2ATM_CARDLESS.HasRows Then
            While (SQLREADER2ATM_CARDLESS.Read)
                amount24 = SQLREADER2ATM_CARDLESS.GetValue(SQLREADER2ATM_CARDLESS.GetOrdinal("AMT")).ToString()
            End While
        End If
        SQLCONNECTION2ATM_CARDLESS.Close()

        ''''''''''''''''''''''''''''END AR'''''''''''''''''''''''''''''''




        '''''''''''''''''''''''''START Cardless_Verve_Token Fees''''''''''''''''''''''''''''''''''

        Dim SQLCONNECTION3ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (Cardless:Paycode Verve Token)"

        account = "ATM WITHDRAWAL (Cardless:Paycode Verve Token)"
        status = "CURRENT"
        acct_type = "ISSUER FEE PAYABLE"
        CMD3ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION3ATM_CARDLESS.Open()
        CMD3ATM_CARDLESS.Connection = SQLCONNECTION3ATM_CARDLESS
        SQLREADER3ATM_CARDLESS = CMD3ATM_CARDLESS.ExecuteReader


        If SQLREADER3ATM_CARDLESS.HasRows Then
            While (SQLREADER3ATM_CARDLESS.Read)
                Fc5 = SQLREADER3ATM_CARDLESS.GetValue(SQLREADER3ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION3ATM_CARDLESS.Close()




        Dim SQLCONNECTION4ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4ATM_CARDLESS As System.Data.SqlClient.SqlDataReader



        '"ATM WITHDRAWAL (Cardless:Paycode Verve Token)"

        account = "ATM WITHDRAWAL (Cardless:Paycode Verve Token)"
        status = "CURRENT"
        acct_type = "ISSUER FEE RECEIVABLE"
        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                fc6 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()


        '"ATM WITHDRAWAL (Cardless:Paycode Verve Token)"

        account = "ATM WITHDRAWAL (Cardless:Paycode Verve Token)"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"

        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                fc7 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()



        '"ATM WITHDRAWAL (Cardless:Paycode Verve Token)"

        account = "ATM WITHDRAWAL (Cardless:Paycode Verve Token)"
        status = "CURRENT"
        acct_type = "SCHEME OWNER ISSUER FEE RECEIVABLE"


        CMD4ATM_CARDLESS.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT]= '{1}' AND [STATUS]='{2}' AND [ACCT_TYPE]='{3}'", DV1, account, status, acct_type)

        SQLCONNECTION4ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION4ATM_CARDLESS
        SQLREADER4ATM_CARDLESS = CMD4ATM_CARDLESS.ExecuteReader


        If SQLREADER4ATM_CARDLESS.HasRows Then
            While (SQLREADER4ATM_CARDLESS.Read)
                fc8 = SQLREADER4ATM_CARDLESS.GetValue(SQLREADER4ATM_CARDLESS.GetOrdinal("FEE")).ToString()
            End While
        End If
        SQLCONNECTION4ATM_CARDLESS.Close()




        ''''''''''''''''''''''''''''END ATM WITHDRAWAL (Cardless:Paycode Verve Token)'''''''''''''''''''''''''''''''





    End Sub


    Private Sub CALL_ATM_CARDLESS_VT_INSERT()

        'VT IS VERVE TOKEN
        'SQL
        Dim SQLCONNECTION2ATM_CARDLESS As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD3ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD4ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD5ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim CMD6ATM_CARDLESS As New System.Data.SqlClient.SqlCommand
        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V4 As String
        Dim V4a As String
        Dim V4b As String
        Dim V4c As String
        Dim v4d As String
        Dim v4e As String


        Dim V5 As String
        Dim V6 As String
        Dim SBU As String
        Dim SBU1 As String
        Dim CMNT As String
        Dim V0 As String
        Dim V1a As String
        Dim V1b As String




        V0 = "NGNEX99963130009"
        V1 = "NGNLI00124430030"
        V1a = "NGNAS00116150028"
        V1b = "NGNIN99954210036"
        V2 = "C"
        V4 = "ATM_WD CARDLESS_VT PYBLE " & " " & DV1.ToShortDateString
        V4a = "ATM_WD CARDLESS_VT RCVBLE" & " " & DV1.ToShortDateString
        V4b = "ATM_WD CARDLESS_VT IFP" & " " & DV1.ToShortDateString
        V4c = "ATM_WD CARDLESS_VT IFR" & " " & DV1.ToShortDateString
        v4d = "ATM_WD CARDLESS_VT AFR" & " " & DV1.ToShortDateString
        v4e = "ATM_WD CARDLESS_VT SCHM OWNR FEES IFR" & " " & DV1.ToShortDateString



        V5 = "999"
        V6 = "D"
        SBU = "96N"
        SBU1 = "94L"
        CMNT = "ALWAYS CONFIRM FROM THE ROU REPORT WHERE CARDLESS TRANSACTIONS DEFAULT TO"




        'INSERT AP (ATM_CARDLESS) 
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD3ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V6, amount23, V4, V5, CMNT, DV1)
        CMD3ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD3ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT AR (ATM_CARDLESS)
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD4ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1a, V2, amount24, V4a, V5, CMNT, DV1)
        CMD4ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD4ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'IFP
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD2ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, V6, Fc5, V4b, V5, SBU, DV1)
        CMD2ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD2ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()



        'INSERT ATM_CARDLESS FEES 'IFR
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, fc6, V4c, V5, SBU, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'AFR
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, fc7, v4d, V5, SBU1, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()


        'INSERT ATM_CARDLESS FEES 'SCHEME OWNER FEE
        SQLCONNECTION2ATM_CARDLESS.Open()
        CMD5ATM_CARDLESS.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1b, V2, fc8, v4e, V5, SBU, DV1)
        CMD5ATM_CARDLESS.Connection = SQLCONNECTION2ATM_CARDLESS
        CMD5ATM_CARDLESS.ExecuteNonQuery()
        SQLCONNECTION2ATM_CARDLESS.Close()






    End Sub

    Private Sub CALL_Remote_Web()



        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double



        account = "WEB"
        account1 = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT PAYABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(AMT) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type, status)



        'SELECT SUM(AMT) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%WEB%' 
        'AND ACCOUNT <>'QUICKTELLER WEB TRANSFERS' 
        'AND ACCT_TYPE ='AMOUNT PAYABLE' AND SETTLEMENT_DATE='2014-07-21'





        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                amount16 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()









    End Sub


    Private Sub Insert_Remote_Web_Amounts()


        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        'Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = "XXXXXXXXXX"
        Vcon = "D"
        'SBU = "93Q"
        CMNT = "REMOTE WEB REPORT IS REQUIRED"



        V4 = "REMOTE WEB" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, Vcon, amount16, V4, V5, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()






    End Sub


    Private Sub CALL_Remote_web_fees()


        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        Dim account2 As String
        'Dim acct_type1 As String
        Dim TEST As Double



        account = "WEB"
        account1 = "QUICKTELLER WEB TRANSFERS"
        account2 = "MASTERCARD LOCAL PROCESSING BILLING(WEB PURCHASE)" 'RECENT INCLUSION IN SETTLEMENT
        status = "CURRENT"
        acct_type = "ISSUER"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}' AND [ACCOUNT] <> '{3}' AND [ACCT_TYPE] LIKE '%{4}%' AND [STATUS]='{5}'", DV1, account, account1, account2, acct_type, status)



        ' ---WEB (ISSUER FEE RCV+PAYABLE)
        ' SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%WEB%' 
        'AND ACCOUNT <>'QUICKTELLER WEB TRANSFERS' AND ACCT_TYPE LIKE'%ISSUER%' AND SETTLEMENT_DATE='2014-07-21'


        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                fee24 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()








    End Sub



    Private Sub CALL_Remote_web_fees_iso()

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader

        Dim status As String
        Dim account As String
        Dim account1 As String
        Dim account2 As String
        'Dim acct_type1 As String
        Dim TEST As Double



        account = "WEB"
        account1 = "QUICKTELLER WEB TRANSFERS"
        account2 = "MASTERCARD LOCAL PROCESSING BILLING(WEB PURCHASE)" 'RECENT INCLUSION IN SETTLEMENT
        status = "CURRENT"
        acct_type = "ISO"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCOUNT] <> '{3}' AND [ACCT_TYPE] LIKE '%{4}%' AND [STATUS]='{5}'", DV1, account, account1, account2, acct_type, status)



        ' ---WEB (ISO FEE)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%WEB%' 
        'AND ACCOUNT <>'QUICKTELLER WEB TRANSFERS' AND ACCT_TYPE LIKE'%ISO%' AND SETTLEMENT_DATE='2014-07-21'


        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                fee25 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()

        'ADD THE REMOTE WEB FEES+ISO FEES HERE
        'fee26 = fee24 + fee25

        'FEES NOW SPLIT   ISO (FEE25)  RWEB(FEE24)





    End Sub


    Private Sub Insert_Remote_WEB_all_fees()


        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim CMD2RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim CMNT1 As String
        Dim V3 As String
        Dim V4 As String
        Dim V5 As String

        V0 = "NGNIN99954210036"
        Vcon = "C"
        SBU = "96N"
        CMNT = "REMOTE WEB FEES"
        CMNT1 = "REMOTE WEB ISO FEES"

        'FEE25 IS (ISSUER WEB) 'SEE WEB ISO MODULE

        V3 = "REMOTE WEB ISO FEES" & " " & DV1.ToShortDateString
        V4 = "REMOTE WEB ISSUER FEES" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee24, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD2RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee25, V3, V5, SBU, CMNT1, DV1)
        CMD2RP.Connection = SQLCONNECTION1RP
        CMD2RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()



    End Sub


    Private Sub Web_Acquired_amounts()
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double



        account = "WEB"
        account1 = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "AMOUNT RECEIVABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(AMT) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type, status)


        '--WEB ACQUIRED AMT
        'SELECT SUM(AMT) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%WEB%' 
        'AND ACCOUNT <>'QUICKTELLER WEB TRANSFERS' 
        'AND ACCT_TYPE ='AMOUNT RECEIVABLE' AND SETTLEMENT_DATE='2014-07-21'



        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                amount17 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()





    End Sub


    Private Sub Web_Acquired_Acquirer_fee_payable()

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double



        account = "WEB"
        account1 = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE PAYABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type, status)


        '--WEB ACQ FEE PAYABLE
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%WEB%' 
        'AND ACCOUNT <>'QUICKTELLER WEB TRANSFERS' AND ACCT_TYPE ='ACQUIRER FEE PAYABLE' AND SETTLEMENT_DATE='2014-07-21'


        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                fee27 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()

        'sum all web_acq_amounts_here
        amount18 = amount17 + fee27 'new amount 17 is total web acquired fee




    End Sub


    Private Sub Insert_Web_Acquired_amounts()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        'Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = "XXXXXXXXXX"
        Vcon = "C"
        'SBU = "93Q"
        CMNT = "WEB ACQUIRED REPORT IS REQUIRED TO CREDIT MERCHANTS"

        'AMT17 IS NET WEB ACQ AMOUNT

        V4 = "WEB ACQUIRED SETTLEMENT" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, Vcon, amount18, V4, V5, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()






    End Sub


    Private Sub CALL_Web_Acquired_Fee_Receivable()

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double


        account = "WEB"
        account1 = "QUICKTELLER WEB TRANSFERS"
        status = "CURRENT"
        acct_type = "ACQUIRER FEE RECEIVABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}'  AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type, status)



        ' ---WEB(ACQUIRER FEE RCV)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='CURRENT' AND ACCOUNT LIKE '%WEB%' 
        'AND ACCOUNT <>'QUICKTELLER WEB TRANSFERS' AND ACCT_TYPE ='ACQUIRER FEE RECEIVABLE' AND SETTLEMENT_DATE='2014-07-21'


        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader


        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
                fee28 = TEST
            End While
        End If
        SQLCONNECTION1RP.Close()


    End Sub


    Private Sub Insert_Web_Acquired_Fee_Receivable()


        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = "NGNIN99954210036"
        Vcon = "C"
        SBU = "92G"
        CMNT = "WEB ACQUIRER RECEIVABLE FEES"

        'FEE28 IS WEB ACQUIRER FEE RECEIVABLE

        V4 = "WEB ACQUIRER RCV FEES" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee28, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()



    End Sub

    Private Sub CALL_LATE_POS_REVERSAL_AMT_PAYABLE()
        'AMOUNT PAYABLE



        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double


        account = "POS"
        'account1 = "QUICKTELLER WEB TRANSFERS"
        status = "LATE"
        acct_type = "AMOUNT PAYABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(AMT) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)

        'SELECT SUM(AMT) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='AMOUNT PAYABLE' AND SETTLEMENT_DATE='2014-07-21'


        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader
        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
            End While
            amount20 = TEST   'AMOUNT PYBLE IS AMOUNT GIVEN BACK TO CUSTOMER
        End If
        SQLCONNECTION1RP.Close()



    End Sub

    Private Sub INSERT_LATE_REVERSAL_AMT_PAYABLE()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        'Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = "XXXXXXXXXX"
        Vcon = "C"
        'SBU = "92G"
        CMNT = "REMOTE POS REPORT IS REQUIRED TO CREDIT CUSTOMER"

        'AMT 20 IS LATE POS AMT PAYABLE

        V4 = "LATE POS AMT PYBLE" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, Vcon, amount20, V4, V5, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()




    End Sub


    Private Sub CALL_LATE_POS_REVERSAL_AMT_RCV()

        'AMOUNT RCV



        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double


        account = "POS"
        'account1 = "QUICKTELLER WEB TRANSFERS"
        status = "LATE"
        acct_type = "AMOUNT RECEIVABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1RP.CommandText = String.Format("SELECT SUM(AMT+FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)

        'SELECT SUM(AMT+FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='AMOUNT RECEIVABLE' AND SETTLEMENT_DATE='2014-07-21'


        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader
        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
            End While
            amount19 = TEST   'AMOUNT RCV IS AMOUNT NET FEE A DEBIT TO THE POS MERCHANT
        End If
        SQLCONNECTION1RP.Close()





    End Sub

    Private Sub INSERT_LATE_REVERSAL_AMT_RCV()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        'Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = "XXXXXXXXXX"
        Vcon = "D"
        'SBU = "92G"
        CMNT = "POS ACQUIRED REPORT IS REQUIRED TO DEBIT MERCHANT"

        'AMT 19 IS LATE POS AMT RECEIVABLE

        V4 = "RVSL:LATE POS AMT RCVBL" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, Vcon, amount19, V4, V5, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()

    End Sub

    Private Sub CALL_LATE_POS_IFP_IFR()

        'ISSUER FEE PAYABLE AND RECEIVABLE

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader
        Dim CMD2RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2RP As System.Data.SqlClient.SqlDataReader
        Dim CMD3RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        Dim account1 As String
        Dim acct_type1 As String
        Dim acct_type2 As String
        Dim TEST As Double
        Dim TEST1 As Double
        Dim TEST2 As Double

        account = "POS"
        account1 = "MASTERCARD LOCAL PROCESSING BILLING(POS PURCHASE)"
        status = "LATE"
        acct_type = "AMOUNT PAYABLE"
        acct_type1 = "ISSUER FEE PAYABLE"
        acct_type2 = "ISSUER FEE RECEIVABLE"


        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}' AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type, status)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='AMOUNT PAYABLE' AND SETTLEMENT_DATE='2014-07-21'



        CMD2RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST1 FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%' AND [ACCOUNT] <> '{2}' AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type1, status)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='ISSUER FEE PAYABLE' AND SETTLEMENT_DATE='2014-07-21'


        CMD3RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST2 FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCOUNT] <> '{2}' AND [ACCT_TYPE]='{3}' AND [STATUS]='{4}'", DV1, account, account1, acct_type2, status)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='ISSUER FEE RECEIVABLE' AND SETTLEMENT_DATE='2014-07-21'

        ''''''''''''''''''''''''''''''''''''FEE UNDER AP

        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader

        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
            End While
            fee29 = TEST   'FEE UNDER AMOUNT PAYABL
        End If
        SQLCONNECTION1RP.Close()



        ''''''''''''''''''''''''''''''''''''''FEE UNDER IFP

        SQLCONNECTION1RP.Open()
        CMD2RP.Connection = SQLCONNECTION1RP
        SQLREADER2RP = CMD2RP.ExecuteReader

        If SQLREADER2RP.HasRows Then
            While (SQLREADER2RP.Read)
                TEST1 = SQLREADER2RP.GetValue(SQLREADER2RP.GetOrdinal("TEST1")).ToString
            End While
            fee30 = TEST1  'IFP
        End If
        SQLCONNECTION1RP.Close()


        ''''''''''''''''''''''''''''''''''''''FEE UNDER IFR

        SQLCONNECTION1RP.Open()
        CMD3RP.Connection = SQLCONNECTION1RP
        SQLREADER3RP = CMD3RP.ExecuteReader

        If SQLREADER3RP.HasRows = True Then
            While (SQLREADER3RP.Read)
                TEST2 = SQLREADER3RP.GetValue(SQLREADER3RP.GetOrdinal("TEST2")).ToString()
            End While
            fee31 = TEST2   'FEE IFR
        End If
        SQLCONNECTION1RP.Close()




        fee32 = fee29 + fee30 + fee31 'NET LATE ISSUER FEE



    End Sub

    Private Sub INSERT_LATE_POS_IFP_IFR()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = ""
        Vcon = ""


        If fee32 >= 0 Then
            V0 = "NGNIN99954210036"
            Vcon = "C"
        End If

        If fee32 < 0 Then
            V0 = "NGNEX99963130009"
            Vcon = "D"
        End If




        SBU = "96N"
        CMNT = "LATE POS NET ISSUER FEES"

        'FEE32 IS LATE POS NET IFP AND IFR

        V4 = "RVSL:LATE POS ISSUER FEES" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V0, Vcon, fee32, V4, V5, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()





    End Sub


    Private Sub CALL_LATE_POS_REVERSAL_AFP()

        'AFP 

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader




        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        Dim TEST As Double



        account = "POS"
        'account1 = "QUICKTELLER WEB TRANSFERS"
        status = "LATE"
        acct_type = "ACQUIRER FEE PAYABLE"



        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='ACQUIRER FEE PAYABLE' AND SETTLEMENT_DATE='2014-07-21'


        ''''''''''''''''''''''''''''''''''''FEE UNDER AFP
        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader



        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
            End While
            fee33 = TEST   'AMOUNT RCV IS AMOUNT NET FEE A DEBIT TO THE POS MERCHANT
        End If
        SQLCONNECTION1RP.Close()










    End Sub


    Private Sub CALL_LATE_POS_REVERSAL_AFR()

        'LATE AFR

        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1RP As System.Data.SqlClient.SqlDataReader




        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        Dim TEST As Double
        'Dim TEST1 As Double


        account = "POS"
        'account1 = "QUICKTELLER WEB TRANSFERS"
        status = "LATE"
        acct_type = "ACQUIRER FEE RECEIVABLE"



        CMD1RP.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)
        'SELECT SUM(FEE) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='ACQUIRER FEE PAYABLE' AND SETTLEMENT_DATE='2014-07-21'



        ''''''''''''''''''''''''''''''''''''FEE UNDER AFR
        SQLCONNECTION1RP.Open()
        CMD1RP.Connection = SQLCONNECTION1RP
        SQLREADER1RP = CMD1RP.ExecuteReader



        If SQLREADER1RP.HasRows Then
            While (SQLREADER1RP.Read)
                TEST = SQLREADER1RP.GetValue(SQLREADER1RP.GetOrdinal("TEST")).ToString()
            End While
            fee34 = TEST   'ACQ FEE RCV
        End If
        SQLCONNECTION1RP.Close()









    End Sub

    Private Sub INSERT_LATE_POS_AFP()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = ""
        Vcon = ""


        If fee33 >= 0 Then
            V0 = "NGNIN99954210036"
            Vcon = "C"
        End If

        If fee33 < 0 Then
            V0 = "NGNEX99963130009"
            Vcon = "D"
        End If




        SBU = "93Q"
        CMNT = "LATE POS ACQUIRER FEES PAYABLE"

        'FEE35 IS LATE POS NET AFP AND AFR

        V4 = "LATE POS ACQ FEE PYBL" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee33, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()


    End Sub


    Private Sub INSERT_LATE_POS_AFR()

        'SQL
        Dim SQLCONNECTION1RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1RP As New System.Data.SqlClient.SqlCommand
        Dim V0 As String    'account (Income or Expense)
        Dim Vcon As String  'credit status
        Dim SBU As String
        Dim CMNT As String
        Dim V4 As String
        Dim V5 As String

        V0 = ""
        Vcon = ""


        If fee34 >= 0 Then
            V0 = "NGNIN99954210036"
            Vcon = "C"
        End If

        If fee34 < 0 Then
            V0 = "NGNEX99963130009"
            Vcon = "D"
        End If




        SBU = "93Q"
        CMNT = "LATE POS ACQUIRER FEES RECEIVABLE"

        'FEE35 IS LATE POS NET AFP AND AFR

        V4 = "RVSL:LATE POS ACQ FEES RCVBL" & " " & DV1.ToShortDateString
        V5 = "999"

        'INSERT NET REMOTE POS ACQUIRER FEES 
        SQLCONNECTION1RP.Open()
        CMD1RP.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[SBU],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}')", V0, Vcon, fee34, V4, V5, SBU, CMNT, DV1)
        CMD1RP.Connection = SQLCONNECTION1RP
        CMD1RP.ExecuteNonQuery()
        SQLCONNECTION1RP.Close()





    End Sub


    Private Sub CALL_LATE_POS_CH_IFR()

        '"CARDHOLDER_ISSUER FEE RECEIVABLE"



        Dim SQLCONNECTION10RP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD10RP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER10RP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        'Dim TEST10 As Double


        account = "POS"
        status = "LATE"
        acct_type = "CARDHOLDER_ISSUER FEE RECEIVABLE"


        CMD10RP.CommandText = String.Format("SELECT SUM(FEE) AS 'TEST10' FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)

        'SELECT SUM(AMT) AS TEST FROM TRANSFER2014  WHERE STATUS ='LATE' AND ACCOUNT LIKE '%POS%' 
        'AND ACCT_TYPE ='CARDHOLDER_ISSUER FEE RECEIVABLE' AND SETTLEMENT_DATE='2014-07-21'


        Try
            SQLCONNECTION10RP.Open()
            CMD10RP.Connection = SQLCONNECTION10RP
            SQLREADER10RP = CMD10RP.ExecuteReader


            If SQLREADER10RP.HasRows Then
                While (SQLREADER10RP.Read)
                    fee36 = SQLREADER10RP.GetValue(SQLREADER10RP.GetOrdinal("TEST10")).ToString
                End While
                'fee36 IS REVERSAL OF REWARD MONEY FROM CUSTOMER,IF THE CUSTOMER DDNT GET IT THEN PASS TO EXPENSE
            End If
            SQLCONNECTION10RP.Close()

        Catch EX As Exception
            MsgBox("There are no late Cardholder Fees to be reversed from customer's accounts")
            fee36 = 0

        End Try










        'code below may come in handy
        'TEST10 = CStr(TEST10) 'convert a double to string
        'TEST10 = CDbl(TEST10) 'convert a string to double


    End Sub


    Private Sub CHECK_LATE_POS_AP()

        'AMOUNT PAYABLE
        '5 LATE POS ITEMS OCCUR IF AMOUNT PAYABLE IS PRESENT


        Dim SQLCONNECTION1LP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD1LP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER1LP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        'Dim TEST As Double


        account = "POS"
        'account1 = "QUICKTELLER WEB TRANSFERS"
        status = "LATE"
        acct_type = "AMOUNT PAYABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD1LP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)



        SQLCONNECTION1LP.Open()
        CMD1LP.Connection = SQLCONNECTION1LP
        SQLREADER1LP = CMD1LP.ExecuteReader
        If SQLREADER1LP.HasRows Then
            CALL_LATE_POS_REVERSAL_AMT_PAYABLE()
            INSERT_LATE_REVERSAL_AMT_PAYABLE()
            CALL_LATE_POS_IFP_IFR()
            INSERT_LATE_POS_IFP_IFR()
        Else
            MsgBox("There are no late reversal amounts payable", vbOKOnly, "INTERSWITCH AUTOMATION")
        End If
        SQLCONNECTION1LP.Close()





    End Sub


    Private Sub CHECK_LATE_POS_AR()


        'AMOUNT RECEIVABLE
        '4 LATE POS ITEMS OCCUR IF AMOUNT PAYABLE IS PRESENT


        Dim SQLCONNECTION2LP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2LP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2LP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        'Dim TEST As Double


        account = "POS"
        'account1 = "QUICKTELLER WEB TRANSFERS"
        status = "LATE"
        acct_type = "AMOUNT RECEIVABLE"
        'acct_type1 = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD2LP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)



        SQLCONNECTION2LP.Open()
        CMD2LP.Connection = SQLCONNECTION2LP
        SQLREADER2LP = CMD2LP.ExecuteReader
        If SQLREADER2LP.HasRows Then
            CALL_LATE_POS_REVERSAL_AMT_RCV()
            INSERT_LATE_REVERSAL_AMT_RCV()
        Else
            MsgBox("There are no late reversal amounts receivable", vbOKOnly, "INTERSWITCH AUTOMATION")
        End If
        SQLCONNECTION2LP.Close()
















    End Sub


    Private Sub CHECK_LATE_POS_AFP()

        'AF PAYABLE



        Dim SQLCONNECTION3LP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3LP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER3LP As System.Data.SqlClient.SqlDataReader


        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        'Dim TEST As Double


        account = "POS"
        status = "LATE"
        acct_type = "ACQUIRER FEE PAYABLE"


        CMD3LP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)



        SQLCONNECTION3LP.Open()
        CMD3LP.Connection = SQLCONNECTION3LP
        SQLREADER3LP = CMD3LP.ExecuteReader
        If SQLREADER3LP.HasRows Then
            CALL_LATE_POS_REVERSAL_AFP()
            INSERT_LATE_POS_AFP()
        Else
            GoTo 1
        End If
1:      SQLCONNECTION3LP.Close()





    End Sub

    Private Sub CHECK_LATE_POS_AFR()


        'ACQUIRER FEE RECEIVABLE



        Dim SQLCONNECTION4LP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD4LP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER4LP As System.Data.SqlClient.SqlDataReader


        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        'Dim TEST As Double


        account = "POS"
        status = "LATE"
        acct_type = "ACQUIRER FEE RECEIVABLE"


        CMD4LP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)



        SQLCONNECTION4LP.Open()
        CMD4LP.Connection = SQLCONNECTION4LP
        SQLREADER4LP = CMD4LP.ExecuteReader
        If SQLREADER4LP.HasRows Then
            CALL_LATE_POS_REVERSAL_AFR()
            INSERT_LATE_POS_AFR()
        Else
            GoTo 1
        End If
1:      SQLCONNECTION4LP.Close()











    End Sub


    Private Sub CHECK_LATE_POS_CH_IFR()

        '"CARDHOLDER_ISSUER FEE RECEIVABLE"
        'REVERSED CARDHOLDER FEES


        Dim SQLCONNECTION5LP As New System.Data.SqlClient.SqlConnection("Data Source=femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5LP As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5LP As System.Data.SqlClient.SqlDataReader


        'Dim RP_acq1 As Double

        Dim status As String
        Dim account As String
        'Dim account1 As String
        'Dim acct_type1 As String
        'Dim TEST As Double


        account = "POS"
        status = "LATE"
        acct_type = "CARDHOLDER_ISSUER FEE RECEIVABLE"

        CMD5LP.CommandText = String.Format("SELECT * FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'  AND [ACCT_TYPE]='{2}' AND [STATUS]='{3}'", DV1, account, acct_type, status)



        SQLCONNECTION5LP.Open()
        CMD5LP.Connection = SQLCONNECTION5LP
        SQLREADER5LP = CMD5LP.ExecuteReader
        If SQLREADER5LP.HasRows Then
            CALL_LATE_POS_CH_IFR()
        Else
            fee36 = 0
            MsgBox("There are no late cardholder issuer fee receivable", vbOKOnly, "INTERSWITCH AUTOMATION")
        End If
        SQLCONNECTION5LP.Close()

    End Sub

    Private Sub MasterCard_Dollar_Settlement()

        Dim SQLCONNECTIONMC As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDMC As New System.Data.SqlClient.SqlCommand
        Dim SQLREADERMC As System.Data.SqlClient.SqlDataReader




        'Dim status As String
        Dim account As String
        Dim Test As Double

        account = "MASTERCARD "

        CMDMC.CommandText = String.Format("SELECT SUM(FEE) AS TEST FROM [TRANSFER2014] WHERE [SETTLEMENT_DATE] = '{0}' AND [ACCOUNT] LIKE '%{1}%'", DV1, account)


        Try
            SQLCONNECTIONMC.Open()
            CMDMC.Connection = SQLCONNECTIONMC
            SQLREADERMC = CMDMC.ExecuteReader


            If SQLREADERMC.HasRows Then
                While (SQLREADERMC.Read)
                    Test = SQLREADERMC.GetValue(SQLREADERMC.GetOrdinal("TEST")).ToString()
                    Mcard_dollar = Test
                End While
            End If


        Catch ex As Exception
            'Call the error here
            MsgBox("Total MasterCard Dollar Expense is" & " " & Test)
            MsgBox("There are no MasterCard Dollar Expense")
            SQLCONNECTIONMC.Close()
            Exit Sub
        End Try

        MsgBox("Total MasterCard Dollar Expense is" & " " & Mcard_dollar)
        SQLCONNECTIONMC.Close()





    End Sub


    Private Sub Insert_MasterCard_Dollar_Settlement()


        'SQL
        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD3 As New System.Data.SqlClient.SqlCommand


        Dim V1 As String  'ACCOUNT
        Dim V2 As String  'STATUS
        Dim V4 As String
        Dim V5 As String
        Dim CMNT As String
        Dim V7 As Double


        V1 = "NGNEX99963130012"
        V7 = Mcard_dollar ' The total MasterCard Dollar Settlement Value
        V2 = "D"

        If V7 >= 0 Then
            V2 = "C"
        End If

        V4 = "MASTERCARD DOLLAR SETTLEMENT" & " " & DV1.ToShortDateString
        V5 = "999"
        CMNT = "The Entries should be raised using the Breakdown of the Dollar Settlement using appropriate SBU code for Issuer and Acquirer"




        'INSERT AR (ATM TRANSFERS)
        SQLCONNECTION3.Open()
        CMD3.CommandText = String.Format("insert into SETTLEMENT ([ACCOUNT],[STATUS],[AMOUNT],[NARRATION],[SOL],[COMMENTS],[DATE]) VALUES ('{0}','{1}',{2},'{3}','{4}','{5}','{6}')", V1, V2, V7, V4, V5, CMNT, DV1)
        CMD3.Connection = SQLCONNECTION3
        CMD3.ExecuteNonQuery()
        SQLCONNECTION3.Close()



    End Sub




    Private Sub Report_viewer()




        Me.INTERSWITCHDataSet.Clear()
        Me.INTERSWITCHDataSet.SETTLEMENT.Clear()
        Me.SETTLEMENTTableAdapter.ClearBeforeFill = True



        Dim dataSet1 As DataSet = New DataSet("DataSet1")
        dataSet1 = SETTLEMENTBindingSource.DataSource
        dataSet1.EnforceConstraints = False


        Dim DT As Date



        DT = DateTimePicker1.Value.Date




        Me.SETTLEMENTTableAdapter.FillBy(Me.INTERSWITCHDataSet.SETTLEMENT, DT)
        Me.ReportViewer1.RefreshReport()



        'Me.SETTLEMENTTableAdapter.FillBy(INTERSWITCHDataSet.SETTLEMENT, DT)

        'TODO: This line of code loads data into the 'INTERSWITCHDataSet.SETTLEMENT' table. You can move, or remove it, as needed.
        'Me.SETTLEMENTTableAdapter.Fill(Me.INTERSWITCHDataSet.SETTLEMENT)

        ' Me.ReportViewer1.RefreshReport()


    End Sub

    Private Sub Data_grid_view()


        DataGridView1.Rows.Clear()



        Dim SQLCONNECTION5DG As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD5DG As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER5DG As System.Data.SqlClient.SqlDataReader


        ' Dim ACCT As String
        'Dim STAT As String
        'Dim AMT As Double
        'Dim NARR As String
        'Dim SOL As String
        'Dim SBU As String
        'Dim COMMENT As String
        Dim DTE As Date


        'DateTimePicker1.CustomFormat = "YYYY-MM-DD"


        DTE = DateTimePicker1.Value.Date
        CMD5DG.CommandType = System.Data.CommandType.Text

        CMD5DG.CommandText = String.Format("SELECT * FROM [SETTLEMENT] WHERE [DATE] = '{0}'", DTE)





        SQLCONNECTION5DG.Open()
        CMD5DG.Connection = SQLCONNECTION5DG
        SQLREADER5DG = CMD5DG.ExecuteReader
        If SQLREADER5DG.HasRows Then
            While (SQLREADER5DG.Read)
                Dim item1 As New DataGridViewRow
                item1.CreateCells(DataGridView1)
                item1.Cells(0).Value = Convert.ToString(SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("ACCOUNT"))).ToString
                item1.Cells(1).Value = SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("STATUS")).ToString
                item1.Cells(2).Value = Convert.ToDouble(SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("AMOUNT")))
                item1.Cells(3).Value = SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("NARRATION")).ToString
                item1.Cells(4).Value = SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("SOL")).ToString
                item1.Cells(5).Value = SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("SBU")).ToString
                item1.Cells(6).Value = SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("COMMENTS")).ToString
                item1.Cells(7).Value = Convert.ToDateTime(SQLREADER5DG.GetValue(SQLREADER5DG.GetOrdinal("DATE"))).ToShortDateString
                DataGridView1.Rows.Add(item1)
            End While
            Me.DataGridView1.Refresh()
            Me.DataGridView1.Show()

        End If


        LinkLabel1.Visible = True

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Eliminate_duplication()

        DataGridView1.Visible = False
        ReportViewer1.Visible = False

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        LinkLabel2.Visible = True
        



    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Show()
        Me.Hide()
    End Sub



    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub



    Private Sub LinkLabel1_Click(sender As Object, e As EventArgs) Handles LinkLabel1.Click

        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To DataGridView1.ColumnCount - 1
                For k As Integer = 1 To DataGridView1.Columns.Count
                    xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
                Next
            Next
        Next

        xlWorkSheet.SaveAs("D:\vbexcel.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        Dim res As MsgBoxResult
        res = MsgBox("Process completed, Would you like to open file?", MsgBoxStyle.YesNo)
        If (res = MsgBoxResult.Yes) Then
            Process.Start("D:\vbexcel.xlsx")
        End If



    End Sub

    
   

    
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked

        If DataGridView1.Visible = True Then
            Report_viewer()
            Me.ReportViewer1.Show()
            DataGridView1.Visible = False
            LinkLabel1.Visible = False
            Exit Sub
        End If


        If DataGridView1.Visible = False Then
            Data_grid_view()
            Me.DataGridView1.Show()
            Me.ReportViewer1.Visible = False
            LinkLabel1.Visible = True
            Exit Sub
        End If


    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Form3.Show()
        Me.Hide()


    End Sub

    Private Sub Second_Delete()

        'SQL  
        Dim SQLCONNECTION2D As New System.Data.SqlClient.SqlConnection("Data Source=.\;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMD2D As New System.Data.SqlClient.SqlCommand
        Dim SQLREADER2D As System.Data.SqlClient.SqlDataReader
        CMD2D.CommandType = System.Data.CommandType.Text   'command syntax

        DV1 = DateTimePicker1.Value.Date

        CMD2D.CommandText = String.Format("SELECT * FROM [SETTLEMENT] WHERE [DATE] = '{0}'", DV1)


        SQLCONNECTION2D.Open()
            CMD2D.Connection = SQLCONNECTION2D
            SQLREADER2D = CMD2D.ExecuteReader

            If SQLREADER2D.HasRows Then
                'CHECK IF DATA WITH DATE SPECIFIED EXISTS
                First_Delete()
                SQLCONNECTION2D.Close()
            Else
                MsgBox("Report for Settlement Date" & "  " & DV1 & "  " & "could no longer be found it may have been deleted or settlement is pending")
                GoTo 1
            End If

        ' Catch ex As Exception
        'MsgBox("There is no file with the specified date to be deleted")
        'SQLCONNECTION2D.Close()
        'End Try

1:      SQLCONNECTION2D.Close()

    End Sub

    Private Sub First_Delete()

        'allows user to delete settlement report 

        Dim SQLCONNECTION3 As New System.Data.SqlClient.SqlConnection("Data Source = femioladipo;Initial Catalog=INTERSWITCH;Integrated Security=True;")
        Dim CMDB1 As New System.Data.SqlClient.SqlCommand

        'DV1 = DateTimePicker1.Value.Date

        CMDB1.CommandText = String.Format("DELETE FROM SETTLEMENT WHERE DATE = '{0}'", DV1)
            CMDB1.Connection = SQLCONNECTION3

            'DELETE COMPUTATION FILE
            SQLCONNECTION3.Open()
            CMDB1.ExecuteNonQuery()
        SQLCONNECTION3.Close()








        DataGridView1.Visible = False
        ReportViewer1.Visible = False
        LinkLabel1.Visible = False
        LinkLabel2.Visible = False

        MsgBox("File for Settlement Date" & "  " & DV1 & "  " & "Succesfully Deleted")

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Second_Delete()
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Form3.Close()
        ' Form2.Close()
    End Sub
End Class
