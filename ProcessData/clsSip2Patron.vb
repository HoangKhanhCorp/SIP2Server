Imports System.Data
Imports System.Data.OracleClient
Imports System.Data.OracleClient.OracleType
Imports System.Data.SqlClient
Imports System.Data.SqlDbType
Imports Libol60.DataAccess.Circulation
Imports Libol60.DataAccess.Admin
Imports Libol60.DataAccess.Common
Public Class clsSip2Patron
    Protected strErrorMsg As String = ""
    Protected intErrorCode As Integer = 0
    Protected strConnectionstring As String = ""
    Protected strDBServer As String = ""

    Protected oraConnection As OracleConnection
    Protected oraCommand As OracleCommand
    Protected oraDataAdapter As New OracleDataAdapter

    Protected sqlConnection As SqlConnection
    Protected sqlCommand As SqlCommand
    Protected sqlDataAdapter As New SqlDataAdapter

    Protected dsData As New DataSet

    Private strTimeZone As String
    Private strCurrencyUnit As String
    Dim strLibName As String = ""
    Protected intUserID As Integer = 1
    Private strUserRight As String

    ' *****************************************************************************************************
    ' End declare variables
    ' Declare public properties
    ' *****************************************************************************************************
    ' UserID property
    Public Property UserID() As Integer
        Get
            Return intUserID
        End Get
        Set(ByVal Value As Integer)
            intUserID = Value
        End Set
    End Property

    ' ErrorMsg property
    Public ReadOnly Property ErrorMsg() As String
        Get
            ErrorMsg = strErrorMsg
        End Get
    End Property

    ' ErrorCode property
    Public ReadOnly Property ErrorCode() As Integer
        Get
            ErrorCode = intErrorCode
        End Get
    End Property

    Public Property Connectionstring() As String
        Get
            Return strConnectionstring
        End Get
        Set(ByVal Value As String)
            strConnectionstring = Value
        End Set
    End Property
    ' strDBServer property
    Public Property DBServer() As String
        Get
            Return strDBServer
        End Get
        Set(ByVal Value As String)
            strDBServer = Value
        End Set
    End Property

    Public Property TimeZone() As String
        Get
            Return strTimeZone
        End Get
        Set(ByVal Value As String)
            strTimeZone = Value
        End Set
    End Property

    Public Property LibName() As String
        Get
            Return strLibName
        End Get
        Set(ByVal Value As String)
            strLibName = Value
        End Set
    End Property
    Public Property CurrencyUnit() As String
        Get
            Return strCurrencyUnit
        End Get
        Set(ByVal Value As String)
            strCurrencyUnit = Value
        End Set
    End Property

    ' *****************************************************************************************************
    ' End declare properties
    ' Implement methods here
    ' *****************************************************************************************************

    ' Initialize method
    ' Purpose: init all objects
    Public Sub Initialize()
        Select Case UCase(strDBServer)
            Case "ORACLE"
                oraConnection = New OracleConnection(strConnectionstring)
                oraCommand = New OracleCommand
                oraCommand.Connection = oraConnection
            Case "SQLSERVER"
                sqlConnection = New SqlConnection(strConnectionstring)
                sqlCommand = New SqlCommand
                sqlCommand.Connection = sqlConnection
        End Select
    End Sub
    ' OpenConnection method
    ' Purpose: open connection
    Public Sub OpenConnection()
        Select Case UCase(strDBServer)
            Case "ORACLE"
                If oraConnection.State = ConnectionState.Closed Then
                    oraConnection.Open()
                End If
            Case "SQLSERVER"
                If sqlConnection.State = ConnectionState.Closed Then
                    sqlConnection.Open()
                End If
        End Select
    End Sub

    Public Function CheckOpenConnection() As Boolean
        CheckOpenConnection = True
        Call OpenConnection()
        Call CloseConnection()
        Try

        Catch ex As Exception
            CheckOpenConnection = False
        End Try
    End Function


    ' CloseConnection method
    ' Purpose: close connection
    Public Sub CloseConnection()
        Select Case UCase(strDBServer)
            Case "ORACLE"
                If Not oraConnection.State = ConnectionState.Closed Then
                    oraConnection.Close()
                End If
            Case "SQLSERVER"
                If Not sqlConnection.State = ConnectionState.Closed Then
                    sqlConnection.Close()
                End If
        End Select
    End Sub


    ' thuc hien cau lenh sql : ket qua lay ra mot datatable
    ' cau lenh sql thuong la cau lenh select 
    Public Function GetData(ByVal strSql As String) As DataTable
        Call OpenConnection()
        Select Case UCase(strDBServer)
            Case "ORACLE"
                With oraCommand
                    .CommandText = strSql
                    .CommandType = CommandType.Text
                    Try
                        oraDataAdapter.SelectCommand = oraCommand
                        oraDataAdapter.Fill(dsData, "tblResult")
                        GetData = dsData.Tables("tblResult")
                        dsData.Tables.Remove("tblResult")
                    Catch OraEx As OracleException
                        strErrorMsg = OraEx.Message.ToString
                        intErrorCode = OraEx.Code
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
            Case "SQLSERVER"
                With sqlCommand
                    .CommandText = strSql
                    .CommandType = CommandType.Text
                    Try
                        sqlDataAdapter.SelectCommand = sqlCommand
                        sqlDataAdapter.Fill(dsData, "tblResult")
                        GetData = dsData.Tables("tblResult")
                        dsData.Tables.Remove("tblResult")
                    Catch sqlClientEx As SqlException
                        strErrorMsg = sqlClientEx.Message.ToString
                        intErrorCode = sqlClientEx.Number
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
        End Select
        Call CloseConnection()
    End Function

    ' thuc hien cau lenh sql 
    ' cau lenh sql thuong la cau lenh insert, update, delete 
    Public Sub ExecuteQuery(ByVal strSql As String)
        Call OpenConnection()
        Select Case UCase(strDBServer)
            Case "ORACLE"
                With oraCommand
                    .CommandText = strSql
                    .CommandType = CommandType.Text
                    Try
                        .ExecuteNonQuery()
                    Catch OraEx As OracleException
                        strErrorMsg = OraEx.Message.ToString
                        intErrorCode = OraEx.Code
                    End Try
                End With
            Case "SQLSERVER"
                With sqlCommand
                    .CommandText = strSql
                    .CommandType = CommandType.Text
                    Try
                        .ExecuteNonQuery()
                    Catch sqlClientEx As SqlException
                        strErrorMsg = sqlClientEx.Message.ToString
                        intErrorCode = sqlClientEx.Number
                    End Try
                End With
        End Select
        Call CloseConnection()
    End Sub
    '--------------------------------------------------
    ' purpose : bo dau trong cau tieng viet
    ' in : xau can bo
    ' out : xau da duoc bo
    ' creator :
    '--------------------------------------------------
    Public Shared Function DecodeUTF8(ByVal strUTFString As String) As String

        Dim bytString() As Byte = StringToByteArray(strUTFString)
        Dim chrChars() As Char
        Dim i As Int32
        Dim strReturn As String = ""
        chrChars = System.Text.Encoding.UTF8.GetChars(bytString)

        For i = 0 To UBound(chrChars)
            strReturn += chrChars(i)
        Next
        Return strReturn
    End Function

    Public Shared Function StringToByteArray(ByVal strString As String) As Byte()
        Dim bytBuffer() As Byte
        Dim i As Int32

        ReDim bytBuffer(Len(strString) - 1)
        For i = 0 To UBound(bytBuffer)
            bytBuffer(i) = CByte(Asc(Mid(strString, i + 1, 1)))
        Next

        Return bytBuffer
    End Function
    Private Function CutVietnameseAccent(ByVal strInputs As String) As String
        Dim strNoAccentChar As String
        Dim strOutput As String = ""
        Dim strInput As String = ""
        Dim inti As Integer
        If strInputs & "" = "" Then
            CutVietnameseAccent = ""
            Exit Function
        End If

        strInputs = DecodeUTF8(strInputs)

        For inti = 0 To strInputs.Length - 1
            strInput = strInputs.Chars(inti)
            If InStr("A,Ả,Ạ,Ầ,Ấ,Ẩ,Ẫ,Ậ,Ă,Ằ,Ắ,Ẳ,Ẵ,Ặ,A", strInput) > 0 Then
                strNoAccentChar = "A"
            ElseIf InStr("a,à,á,ả,ã,ạ,â,ầ,ấ,ẩ,ẫ,ậ,ă,ằ,ắ,ẳ,ẵ,ặ", strInput) > 0 Then
                strNoAccentChar = "a"
            ElseIf InStr("E,È,É,Ẻ,Ẽ,Ẹ,Ê,Ề,Ế,Ể,Ễ,Ệ,E", strInput) > 0 Then
                strNoAccentChar = "E"
            ElseIf InStr("e,è,é,ẻ,ẽ,ẹ,ê,ề,ế,ể,ễ,ệ,e", strInput) > 0 Then
                strNoAccentChar = "e"
            ElseIf InStr("U,Ù,Ú,Ủ,Ũ,Ụ,Ư,Ừ,Ứ,Ử,Ữ,Ự,U,U", strInput) > 0 Then
                strNoAccentChar = "U"
            ElseIf InStr("u,ù,ú,ủ,ũ,ụ,ư,ừ,ứ,ử,ữ,ự,u", strInput) > 0 Then
                strNoAccentChar = "u"
            ElseIf InStr("I,Ì,Í,Ỉ,Ĩ,Ị,I", strInput) > 0 Then
                strNoAccentChar = "I"
            ElseIf InStr("i,ì,í,ỉ,ĩ,ị", strInput) > 0 Then
                strNoAccentChar = "i"
            ElseIf InStr("O,Ò,Ó,Ỏ,Õ,Ọ,Ô,Ồ,Ố,Ổ,Ỗ,Ộ,Ơ,Ờ,Ớ,Ở,Ỡ,Ợ,O", strInput) > 0 Then
                strNoAccentChar = "O"
            ElseIf InStr("o,ò,ó,ỏ,õ,ọ,ô,ồ,ố,ổ,ỗ,ộ,ơ,ờ,ớ,ở,ỡ,ợ,o", strInput) > 0 Then
                strNoAccentChar = "o"
            ElseIf InStr("Y,Ỳ,Ý,Ỷ,Ỹ,Ỵ,Y", strInput) > 0 Then
                strNoAccentChar = "Y"
            ElseIf InStr("y,ỳ,ý,ỷ,ỹ,ỵ", strInput) > 0 Then
                strNoAccentChar = "y"
            ElseIf InStr("Đ,Ä", strInput) > 0 Then
                strNoAccentChar = "D"
            ElseIf InStr("đ", strInput) > 0 Then
                strNoAccentChar = "d"
            Else
                strNoAccentChar = strInput
            End If
            strOutput = strOutput & strNoAccentChar
        Next
        CutVietnameseAccent = strOutput.Trim
    End Function

    Private Sub LoadFields(ByVal arrLabels() As String, ByRef colf As Collection, ByVal req As String)
        Dim i As Integer
        Dim strPairs
        Dim strFieldID, strFieldVal, strSIP2IDs As String
        strPairs = Split(req, "|")
        For i = 0 To UBound(strPairs)
            If strPairs(i).ToString.Length > 2 Then
                strFieldID = strPairs(i).ToString.Substring(0, 2)
                strFieldVal = strPairs(i).ToString.Substring(2, Len(strPairs(i)) - 2)
                If Not colf.Contains(strFieldID) Then
                    colf.Add(Item:=strFieldVal, Key:=strFieldID)
                End If

            End If


            strSIP2IDs = strSIP2IDs & strFieldID & ","
        Next
        For i = LBound(arrLabels) To UBound(arrLabels)
            If InStr(strSIP2IDs, arrLabels(i)) = 0 Then
                colf.Add(Item:="", Key:=arrLabels(i))
            End If
        Next
    End Sub

    Private Function PrintISOTime(ByVal dtmDate As Date, ByVal tz As String) As String
        PrintISOTime = Year(dtmDate) & CStr(Month(dtmDate)).PadLeft(2, "0") & CStr(dtmDate.Day).PadLeft(2, "0") & tz & CStr(Hour(dtmDate)).PadLeft(2, "0") & CStr(Minute(dtmDate)).PadLeft(2, "0") & CStr(Second(dtmDate)).PadLeft(2, "0")
    End Function

    Private Function ParseISOTime(ByVal strISODate) As Date
        Try
            ParseISOTime = CDate(Mid(strISODate, 5, 2) & "/" & Mid(strISODate, 7, 2) & "/" & Mid(strISODate, 1, 4) & " " & Mid(strISODate, 13, 2) & ":" & Mid(strISODate, 15, 2) & ":" & Mid(strISODate, 17, 2))
        Catch ex As Exception
            ParseISOTime = CDate(Mid(strISODate, 7, 2) & "/" & Mid(strISODate, 5, 2) & "/" & Mid(strISODate, 1, 4) & " " & Mid(strISODate, 13, 2) & ":" & Mid(strISODate, 15, 2) & ":" & Mid(strISODate, 17, 2))
        End Try

    End Function

    Private Function GetContent(ByVal strCopynumber As String) As DataTable
        '   Dim tblContent As DataTable = GetData("SELECT A.Content,B.Note,D.Code,C.Symbol from Field200s A,HOLDING B, Holding_Location C, Holding_library D where A.ItemID=B.ItemID and (B.CopyNumber)='" & strCopynumber & "' and A.FieldCode = '245' And B.LocationID=C.ID and C.LibID=D.ID")
        Dim tblContent As DataTable = GetData("SELECT A.Gia_tri as Content,isnull(B.Note,'') as Note,D.Ten_viet_tat as Code,C.Ten_viet_tat_kho as Symbol from Field200s A,Ma_xep_gia B, Kho C, Thu_vien D where A.Tai_lieu_id=B.Tai_lieu_id and (B.Ma_xep_gia)='" & strCopynumber & "' and A.Truong_ID = 448 And B.Kho_ID=C.Kho_ID and C.Thu_vien_ID=D.Thu_vien_ID")

        If tblContent.Rows.Count > 0 Then
            If IsDBNull(tblContent.Rows(0).Item("Note")) Then
                tblContent.Rows(0).Item("Note") = ""
            End If
        End If
        Return tblContent
    End Function
    Private Function ConvertDateBack(ByVal dt As Date) As String
        Return Left(dt.ToString, Len(dt.ToString) - 3)
    End Function
    Public Function PatronEnable(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strTransDate, strOutput As String
        Dim arrTemp() As String = {"AO", "AA", "AD", "AC"}
        Dim tblResult As DataTable
        strTransDate = Left(req, 18)
        req = Right(req, Len(req) - 18)
        Call LoadFields(arrTemp, colFields, req)
        strOutput = "26              001" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA")
        'tblResult = Me.GetData("SELECT FirstName,MiddleName,LastName, Password FROM CIR_PATRON WHERE Code = '" & colFields("AA") & "'")
        tblResult = Me.GetData("SELECT Ho_ten as Fullname,Mat_khau as Password FROM Ban_doc WHERE So_the = '" & colFields("AA") & "'")
        If Not tblResult Is Nothing AndAlso tblResult.Rows.Count > 0 Then
            'If tblResult.Rows(0).Item("Password") = colFields("AD") Then
            '    strOutput = strOutput & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("FirstName") & " " & tblResult.Rows(0).Item("MiddleName") & " " & tblResult.Rows(0).Item("LastName")) & "|BLY|CQY"
            'Else
            '    strOutput = strOutput & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("FirstName") & " " & tblResult.Rows(0).Item("MiddleName") & " " & tblResult.Rows(0).Item("LastName")) & "|BLY|CQN"
            'End If
            If tblResult.Rows(0).Item("Password") = colFields("AD") Then
                strOutput = strOutput & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("Fullname")) & "|BLY|CQY"
            Else
                strOutput = strOutput & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("Fullname")) & "|BLY|CQN"
            End If
        Else
            strOutput = strOutput & "|AE|BLN|CQN"
        End If
        tblResult = Nothing
        PatronEnable = strOutput
    End Function

    Public Function RenewAll(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strTransDate, strOutput As String
        Dim intRenewedCount, intUnRenewedCount As Integer
        Dim dtmNewDueDate As Date
        Dim strRenewedItems, strUnRenewedItems, strTitle As String
        Dim arrTemp() As String = {"AO", "AA", "AD", "AC", "BO"}
        Dim tblResult As DataTable
        Dim tblResult1 As DataTable

        strTransDate = Left(req, 18)
        req = Right(req, Len(req) - 18)
        Call LoadFields(arrTemp, colFields, req)
        ' tblResult = Me.GetData("SELECT ID FROM CIR_PATRON WHERE Code = '" & colFields("AA") & "' AND Password = '" & colFields("AD") & "'")
        tblResult = Me.GetData("SELECT ID FROM Ban_doc WHERE So_the = '" & colFields("AA") & "' AND Mat_khau = '" & colFields("AD") & "'")
        If Not tblResult Is Nothing AndAlso tblResult.Rows.Count > 0 Then
            tblResult1 = Me.GetData("SELECT CIR_LOAN.ItemID, RenewCount, Renewals, CIR_LOAN.ID, RenewalPeriod, TimeUnit, DueDate , CIR_LOAN.LocationID, CIR_LOAN.CopyNumber, Field200s.Content FROM CIR_LOAN, HOLDING, Field200s, CIR_LOAN_TYPE WHERE Field200s.ItemID = CIR_LOAN.ItemID AND HOLDING.ItemID = CIR_LOAN.ItemID  AND HOLDING.CopyNumber = CIR_LOAN.CopyNumber AND CIR_LOAN.LoanMode = CIR_LOAN_TYPE.ID AND FieldCode = '245' AND PatronID = " & tblResult.Rows(0).Item("ID"))
            intRenewedCount = 0
            intUnRenewedCount = 0
            If Not tblResult1 Is Nothing Then
                Dim i As Integer
                For i = 0 To tblResult1.Rows.Count - 1
                    strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblResult1.Rows(0).Item("Content")))
                    If InStr(strTitle, "/") > 0 Then
                        strTitle = Trim(Left(strTitle, InStr(strTitle, "/") - 1))
                    End If
                    strTitle = strTitle & ". Item Id: " & tblResult1.Rows(0).Item("CopyNumber")
                    If Not IsDBNull(tblResult1.Rows(0).Item("RenewCount")) AndAlso CLng(tblResult1.Rows(0).Item("Renewals")) <= CLng(tblResult1.Rows(0).Item("RenewCount")) Then
                        intUnRenewedCount = intUnRenewedCount + 1
                        strUnRenewedItems = strUnRenewedItems & "|BN" & strTitle
                    ElseIf CDate(tblResult1.Rows(0).Item("DueDate")) < Now Then
                        intUnRenewedCount = intUnRenewedCount + 1
                        strUnRenewedItems = strUnRenewedItems & "|BN" & strTitle
                    ElseIf (CInt(tblResult1.Rows(0).Item("TimeUnit")) = 1 And DateDiff("d", Now, CDate(tblResult1.Rows(0).Item("DueDate"))) >= 3) Or (CInt(tblResult1.Rows(0).Item("TimeUnit")) = 2 And DateDiff("h", Now, CDate(tblResult1.Rows(0).Item("DueDate"))) >= 1) Then
                        intUnRenewedCount = intUnRenewedCount + 1
                        strUnRenewedItems = strUnRenewedItems & "|BN" & strTitle
                    Else
                        Dim tblResult3 As DataTable
                        tblResult3 = GetData("SELECT * FROM CIR_HOLDING WHERE ItemID = " & tblResult1.Rows(0).Item("ItemID") & " AND InTurn = 0 AND (CopyNumber = '' OR CopyNumber IS NULL OR  CopyNumber = '" & tblResult1.Rows(0).Item("CopyNumber") & "')")
                        If Not tblResult3 Is Nothing AndAlso tblResult3.Rows.Count > 0 Then
                            intUnRenewedCount = intUnRenewedCount + 1
                            strUnRenewedItems = strUnRenewedItems & "|BN" & strTitle
                        Else
                            intRenewedCount = intRenewedCount + 1
                            strRenewedItems = strRenewedItems & "|BM" & strTitle
                            dtmNewDueDate = RenewItem(tblResult1.Rows(0).Item("ID"), CInt(tblResult1.Rows(0).Item("RenewalPeriod")), CInt(tblResult1.Rows(0).Item("TimeUnit")), CDate(tblResult1.Rows(0).Item("DueDate")))
                        End If
                    End If
                Next
                strOutput = "661" & CStr(intRenewedCount).PadLeft(4, "0") & CStr(intUnRenewedCount).PadLeft(4, "0") & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & strRenewedItems & strUnRenewedItems
            End If
        Else
            strOutput = "66000000000" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AFPatron information is invalid"
        End If
        RenewAll = strOutput
    End Function

    Public Function Renew(ByVal req As String) As String
        Dim colFields As New Collection
        Dim dtmNewDueDate As Date
        Dim strTransDate, strThirdPartyAllowed, strNoBlock, strNBDueDate As String
        Dim strSelectedFields, strSelectedTabs, strJoinSQL As String
        Dim strOutput As String
        Dim arrTemp() As String = {"AO", "AA", "AD", "AB", "AJ", "AC", "CH", "BO"}
        Dim tblResult As DataTable
        Dim tblResult1 As DataTable

        strThirdPartyAllowed = Left(req, 1)
        strNoBlock = Mid(req, 2, 1)
        strTransDate = Mid(req, 3, 18)
        strNBDueDate = Mid(req, 21, 18)
        req = Right(req, Len(req) - 38)
        Call LoadFields(arrTemp, colFields, req)
        Dim intSoNgayGiaHan As Int32 = 10
        'Check xem lai
        ' If UCase(strNoBlock) = "Y" Then
        If 1 = 1 Then
            If strNBDueDate.Trim = "" Then
                strNBDueDate = Now.AddDays(8).ToString("yyyyMMdd    hhmmss")
            End If
            dtmNewDueDate = ParseISOTime(strNBDueDate)
            If Not colFields("AB") = "" Then
                If Not strDBServer = "ORACLE" Then
                    tblResult = Me.GetData("select isnull(Ngay_tra,getdate()) as NgayTra ,isnull(So_luot_gia_han,0) as SoLuot,l.Renewals,l.RenewalPeriod from An_pham_cho_muon m left join LoanType l on m.LoanType=l.LoanTypeID where Ma_xep_gia='" & colFields("AB") & "'")
                    strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB")
                    If tblResult Is Nothing Then
                        strOutput = strOutput & "|AJ|AH|AFNo item specified"
                        Return strOutput
                        Exit Function
                    End If
                    If tblResult.Rows.Count = 0 Then
                        strOutput = strOutput & "|AJ|AH|AFNo item specified"
                        Return strOutput
                        Exit Function
                    End If
                    If Date.Parse(tblResult.Rows(0)("NgayTra")) < Now Then
                        strOutput = strOutput & "|AJ|AH|AFSach qua han, khong duoc Gia han."
                        Return strOutput
                        Exit Function
                    End If
                    If tblResult.Rows(0)("SoLuot") >= tblResult.Rows(0)("Renewals") Then
                        strOutput = strOutput & "|AJ|AH|AFQua luot Gia han."
                        Return strOutput
                        Exit Function
                    End If
                    dtmNewDueDate = Date.Parse(tblResult.Rows(0)("NgayTra")).AddDays(CInt(tblResult.Rows(0)("RenewalPeriod")))
                    strNBDueDate = PrintISOTime(dtmNewDueDate, TimeZone)
                    ' Me.ExecuteQuery("update An_pham_cho_muon set Ngay_tra='" & ConvertDateBack(dtmNewDueDate) & "',So_luot_gia_han=ISNULL(So_luot_gia_han,0)+1 where So_the_ID in (select id from Ban_doc where so_the like '" & colFields("AA") & "') and Ma_xep_gia='" & colFields("AJ") & "'")
                    Me.ExecuteQuery(" update  An_pham_cho_muon Set Ngay_tra='" & dtmNewDueDate.ToString("MM/dd/yyyy hh:mm:sss") & "', So_luot_gia_han = ISNULL(So_luot_gia_han,0)+1 where  So_the_ID in (select id from Ban_doc where so_the like '" & colFields("AA") & "') and Ma_xep_gia='" & colFields("AB") & "'")

                End If
                strOutput = "301YUY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & colFields("AJ") & "|AH" & strNBDueDate
                'ElseIf Not colFields("AJ") = "" Then
                '    If Not strDBServer = "ORACLE" Then
                '        ' Me.ExecuteQuery("UPDATE CIR_LOAN SET RenewCount = RenewCount + 1, DueDate = '" & ConvertDateBack(dtmNewDueDate) & "' WHERE PatronID IN (SELECT ID FROM CIR_PATRON WHERE Code = '" & colFields("AA") & "') AND ItemID IN (SELECT ItemID FROM ITEM WHERE Code = '" & colFields("AJ") & "')")
                '        ' Me.ExecuteQuery("update An_pham_cho_muon set Ngay_tra='" & ConvertDateBack(dtmNewDueDate) & "',So_luot_gia_han=ISNULL(So_luot_gia_han,0)+1 where  Ma_xep_gia='" & colFields("AJ") & "'")
                '        Me.ExecuteQuery("update  An_pham_cho_muon set Ngay_tra=DATEADD(D," & intSoNgayGiaHan.ToString() & ",Ngay_tra), So_luot_gia_han = ISNULL(So_luot_gia_han,0)+1 where Ma_xep_gia='" & colFields("AJ") & "'")

                '    End If
                '    strOutput = "301YUY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & colFields("AJ") & "|AH" & PrintISOTime(dtmNewDueDate, TimeZone)
            Else
                strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFNo item specified"
                Return strOutput
                Exit Function
            End If
        Else
            If Not strDBServer = "ORACLE" Then
                ' strSelectedFields = "CIR_PATRON.Password, CIR_LOAN_TYPE.Renewals, HOLDING.LocationID, CIR_LOAN_TYPE.RenewalPeriod, CIR_LOAN_TYPE.TimeUnit, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.FixedFee, CIR_LOAN.*"
                strSelectedFields = "CIR_PATRON.Mat_khau Password, CIR_LOAN_TYPE.Renewals, HOLDING.Kho_ID LocationID, CIR_LOAN_TYPE.RenewalPeriod, CIR_LOAN_TYPE.TimeUnit, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.FixedFee, CIR_LOAN.*"
            Else
                strSelectedFields = "CIR_PATRON.Password, CIR_LOAN_TYPE.Renewals, HOLDING.LocationID, CIR_LOAN_TYPE.RenewalPeriod, CIR_LOAN_TYPE.TimeUnit, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.FixedFee, CIR_LOAN.ID, To_Char(DueDate, 'mm/dd/yyyy hh24:mi:ss') AS DueDate, CIR_LOAN.RenewCount, CIR_LOAN.LoanMode, CIR_LOAN.ItemID"
            End If
            strSelectedTabs = "Ban_doc CIR_PATRON,LoanType CIR_LOAN_TYPE,An_pham_cho_muon CIR_LOAN,Ma_xep_gia HOLDING"
            ' strJoinSQL = "CIR_LOAN.PatronID = CIR_PATRON.ID AND CIR_LOAN.LoanTypeID = CIR_LOAN_TYPE.ID AND HOLDING.CopyNumber = CIR_LOAN.CopyNumber AND HOLDING.ItemID = CIR_LOAN.ItemID AND Code = '" & colFields("AA") & "'"
            strJoinSQL = "CIR_LOAN.So_the_ID = CIR_PATRON.ID AND CIR_LOAN.LoanType = CIR_LOAN_TYPE.LoanTypeID AND HOLDING.Ma_xep_gia = CIR_LOAN.Ma_xep_gia AND HOLDING.Tai_lieu_ID = CIR_LOAN.Tai_lieu_ID  AND So_the = '" & colFields("AA") & "'"
            If Not colFields("AB") = "" Then
                strJoinSQL = strJoinSQL & " And (CIR_LOAN.Ma_xep_gia) = '" & CStr(colFields("AB")) & "'"
            ElseIf Not colFields("AJ") = "" Then
                strSelectedTabs = strSelectedTabs & ",Tai_lieu ITEM"
                strJoinSQL = strJoinSQL & " AND CIR_LOAN.Tai_lieu_id = ITEM.Tai_lieu_id AND CIR_LOAN.Ma_xep_gia = '" & colFields("AJ") & "'"
            Else
                strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFNo item specified"
                Renew = strOutput
                Exit Function
            End If
            tblResult = Me.GetData("SELECT " & strSelectedFields & " FROM " & strSelectedTabs & " WHERE " & strJoinSQL)
            If Not tblResult Is Nothing AndAlso tblResult.Rows.Count > 0 Then
                If Not IsDBNull(tblResult.Rows(0).Item("Password")) AndAlso (Not tblResult.Rows(0).Item("Password") = colFields("AD")) Then
                    strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFWrong patron password"
                ElseIf Not IsDBNull(tblResult.Rows(0).Item("RenewCount")) AndAlso CLng(tblResult.Rows(0).Item("Renewals")) <= CLng(tblResult.Rows(0).Item("RenewCount")) Then
                    strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFNo more renewal allowed"
                ElseIf CDate(tblResult.Rows(0).Item("DueDate")) < Now Then
                    strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFItem is overdue"
                ElseIf (CInt(tblResult.Rows(0).Item("TimeUnit")) = 1 And DateDiff("d", Now, CDate(tblResult.Rows(0).Item("DueDate"))) >= 3) Or (CInt(tblResult.Rows(0).Item("TimeUnit")) = 2 And DateDiff("h", Now, CDate(tblResult.Rows(0).Item("DueDate"))) >= 1) Then
                    strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFToo early to make renew request for this item"
                Else
                    tblResult1 = Me.GetData("SELECT * FROM CIR_HOLDING WHERE ItemID = " & tblResult.Rows(0).Item("ItemID") & " AND InTurn = 0 AND (CopyNumber = '' OR CopyNumber = '" & colFields("AB") & "')")
                    If Not tblResult1 Is Nothing AndAlso tblResult1.Rows.Count = 0 Then
                        If tblResult.Rows(0).Item("Fee") > 0 And Not UCase(colFields("BO")) = "Y" And CInt(tblResult.Rows(0).Item("FixedFee")) = 0 Then
                            If CInt(tblResult.Rows(0).Item("TimeUnit")) = 1 Then
                                strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|BH" & CurrencyUnit & "|BV" & tblResult.Rows(0).Item("Fee") & "|AJ|AH|AFFee per day"
                            Else
                                strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|BH" & CurrencyUnit & "|BV" & tblResult.Rows(0).Item("Fee") & "|AJ|AH|AFFee per hour"
                            End If
                        Else
                            dtmNewDueDate = RenewItem(tblResult.Rows(0).Item("ID"), CInt(tblResult.Rows(0).Item("RenewalPeriod")), CInt(tblResult.Rows(0).Item("TimeUnit")), CDate(tblResult.Rows(0).Item("DueDate")))
                            strOutput = "301YUY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & colFields("AJ") & "|AH" & PrintISOTime(dtmNewDueDate, TimeZone)
                        End If
                    Else
                        strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFItem is on hold"
                    End If
                End If
            Else
                strOutput = "300NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH|AFPatron does not check out the specified item or invalid patron"
            End If
        End If
        Renew = strOutput
    End Function

    Public Function ItemStatusUpdate(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strTransDate, strOutput As String
        Dim arrTemp() As String = {"AO", "AB", "AC", "CH"}
        Dim tblResult As DataTable
        strTransDate = Left(req, 18)
        req = Right(req, Len(req) - 18)
        Call LoadFields(arrTemp, colFields, req)
        tblResult = Me.GetData("SELECT ID FROM HOLDING WHERE CopyNumber = '" & colFields("AB") & "'")
        If tblResult.Rows.Count > 0 Then
            Me.ExecuteQuery("UPDATE HOLDING SET Note = '" & Replace(colFields("CH"), "'", "''") & "' WHERE ID = " & tblResult.Rows(0).Item("ID"))
            strOutput = "201" & PrintISOTime(Now, TimeZone) & "AB" & colFields("AB") & "|CH" & colFields("CH") & "|AFItem status updated successfully"
        Else
            strOutput = "200" & PrintISOTime(Now, TimeZone) & "AB" & colFields("AB") & "|CH" & colFields("CH") & "|AFItem identifier does not exist"
        End If
        ItemStatusUpdate = strOutput
    End Function


    Public Function ItemInformation(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strTransDate, strCirStatus, strOwner, strFeeType, strFee, strNote As String
        Dim strDueDate, strReturnDate, strPickUpDate, strTitle, strMediaType, strOutput As String
        Dim curFee As Double
        Dim dblDocID As Double
        Dim intQueueLen As Integer
        Dim arrTemp() As String = {"AO", "AB", "AC"}
        Dim tblResult As DataTable
        Dim strPermanentLoc, strRecallDate As String

        '1720200220    203549AO|ABDATH102724|AZF762
        strTransDate = Left(req, 18)
        req = Right(req, Len(req) - 18)

        Call LoadFields(arrTemp, colFields, req)
        Dim strItemDKCB As String = colFields("AB")
        If InStr(strItemDKCB, "Item ID:") > 8 Then
            strItemDKCB = strItemDKCB.Substring(InStr(strItemDKCB, "Item ID:") + 8, strItemDKCB.Length - InStr(strItemDKCB, "Item ID:") - 8).Trim
        End If
        '  tblResult = Me.GetData("SELECT 0 AS Missing, ItemID, CIR_LOAN_TYPE.TimeUnit, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.FixedFee, HOLDING.OnHold, HOLDING.InCirculation, HOLDING.InUsed, Symbol, Code , Name , LockedReason FROM HOLDING, HOLDING_LOCATION, HOLDING_LIBRARY, CIR_LOAN_TYPE WHERE CIR_LOAN_TYPE.ID = HOLDING.LoanTypeID AND HOLDING_LIBRARY.ID = HOLDING.LibID AND HOLDING.LocationID = HOLDING_LOCATION.ID AND CopyNumber = '" & colFields("AB") & "' UNION SELECT 1 AS Missing, ItemID, CIR_LOAN_TYPE.TimeUnit, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.FixedFee, 0 AS OnHold, 0 As InCirculation, 0 AS InUsed, Symbol, Code, Name, HOLDING_REMOVE_REASON.Reason AS LockedReason FROM HOLDING_REMOVED LEFT JOIN HOLDING_REMOVE_REASON ON HOLDING_REMOVED.REASON=HOLDING_REMOVE_REASON.ID, HOLDING_LOCATION, HOLDING_LIBRARY, CIR_LOAN_TYPE WHERE CIR_LOAN_TYPE.ID = HOLDING_REMOVED.LoanTypeID AND HOLDING_LIBRARY.ID = HOLDING_REMOVED.LibID AND HOLDING_REMOVED.LocationID = HOLDING_LOCATION.ID AND CopyNumber = '" & colFields("AB") & "'")
        tblResult = Me.GetData("SELECT 0 AS Missing,Tai_lieu_ID ItemID, CIR_LOAN_TYPE.TimeUnit, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.FixedFee, HOLDING.OnHold, HOLDING.InCirculation, HOLDING.InUsed,HOLDING_LOCATION.Ten_viet_tat_kho Symbol,Ma_xep_gia Code ,HOLDING_LOCATION.Kho Name , LockedReason FROM Ma_xep_gia HOLDING,Kho HOLDING_LOCATION,Thu_vien HOLDING_LIBRARY,LoanType CIR_LOAN_TYPE WHERE CIR_LOAN_TYPE.LoanTypeID = HOLDING.LoanTypeID AND HOLDING_LIBRARY.Thu_vien_ID = HOLDING.Ten_thu_vien_ID AND HOLDING.Kho_ID = HOLDING_LOCATION.Kho_ID AND Ma_xep_gia like '" & strItemDKCB & "' ")
        If tblResult.Rows.Count > 0 Then
            dblDocID = CDbl(tblResult.Rows(0).Item("ItemID"))
            If CInt(tblResult.Rows(0).Item("Missing")) = 1 Then
                strCirStatus = "13"
            ElseIf Not CInt(tblResult.Rows(0).Item("OnHold")) = 0 Then
                strCirStatus = "02" '"06"
            ElseIf CInt(tblResult.Rows(0).Item("InCirculation")) = 0 Then
                strCirStatus = "08" ' "06"
            ElseIf Not CInt(tblResult.Rows(0).Item("InUsed")) = 0 Then

                strCirStatus = "06" ' "04"
            Else
                'Dang cai nhau voi thang Duc la o day, 03 hay 04 nha
                strCirStatus = "03"
                'strCirStatus = "03"
            End If
            strNote = ""
            If Not IsDBNull(tblResult.Rows(0).Item("LockedReason")) Then
                strNote = tblResult.Rows(0).Item("LockedReason")
            End If
            strOwner = CutVietnameseAccent(tblResult.Rows(0).Item("Code"))
            If Not tblResult.Rows(0).Item("Name") = "" Then
                strOwner = CutVietnameseAccent(tblResult.Rows(0).Item("Name")) & " (" & strOwner & ")"
            End If
            strPermanentLoc = CutVietnameseAccent(tblResult.Rows(0).Item("Name") & "") '& "--" & CutVietnameseAccent(tblResult.Rows(0).Item("Symbol") & "")
            curFee = tblResult.Rows(0).Item("Fee")
            If curFee > 0 Then
                strFeeType = "06"
                If Not CInt(tblResult.Rows(0).Item("FixedFee")) = 0 Then
                    If CInt(tblResult.Rows(0).Item("TimeUnit")) = 1 Then
                        strFee = CStr(curFee) & " per day"
                    Else
                        strFee = CStr(curFee) & " per hour"
                    End If
                Else
                    strFee = CStr(curFee)
                End If
            Else
                strFee = "0"
                strFeeType = "01"
            End If
            Dim tblResult1 As DataTable
            tblResult1 = Me.GetData("Select isnull(Ngay_tra,getdate()) As DueDate, RecallDate FROM an_pham_cho_muon WHERE ma_xep_gia = '" & strItemDKCB & "'")
            If tblResult1.Rows.Count > 0 Then
                strDueDate = CDate(tblResult1.Rows(0).Item("DueDate")).ToString("dd/MM/yyyy") ' PrintISOTime(CDate(tblResult1.Rows(0).Item("DueDate")), TimeZone)
                If Not IsDBNull(tblResult1.Rows(0).Item("RecallDate")) AndAlso Not tblResult1.Rows(0).Item("RecallDate") = "" Then
                    strRecallDate = PrintISOTime(CDate(tblResult1.Rows(0).Item("RecallDate")), TimeZone)
                End If
            End If
            tblResult1 = Nothing
            tblResult1 = Me.GetData("SELECT Gia_tri as Content FROM Field200s WHERE Truong_id=448 AND Tai_lieu_id = " & dblDocID)
            If tblResult1.Rows.Count > 0 Then
                strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblResult1.Rows(0).Item("Content"))).Trim()
                If InStr(strTitle, "/") > 0 Then
                    strTitle = Left(strTitle, InStr(strTitle, "/") - 1).Trim
                End If
            End If
            'If Not strDBServer = "ORACLE" Then
            '    Me.ExecuteQuery("DELETE FROM Giu_cho WHERE Thoi_diem_ket_thuc < GetDate() OR Thoi_diem_het_han < GETDATE()")
            'Else
            '    Me.ExecuteQuery("DELETE FROM CIR_HOLDING WHERE ExpiredDate < SYSDATE OR TimeOutDate < SYSDATE")
            'End If
            tblResult1 = Nothing
            tblResult1 = Me.GetData("SELECT Thoi_diem_het_han as TimeOutDate FROM Giu_cho WHERE (Ma_xep_gia = '' OR Ma_xep_gia = '" & strItemDKCB & "') AND Tai_lieu_ID = " & dblDocID & " ORDER BY Thoi_diem_het_han")
            ' SELECT TimeOutDate FROM CIR_HOLDING WHERE (CopyNumber = '' OR CopyNumber is not null OR CopyNumber = '" & colFields("AB") & "') AND ItemID = " & dblDocID & " ORDER BY TimeOutDate")
            intQueueLen = tblResult1.Rows.Count
            If intQueueLen > 0 Then
                strPickUpDate = PrintISOTime(CDate(tblResult1.Rows(0).Item("TimeOutDate")), TimeZone)
            End If
            strOutput = "18" & strCirStatus & "00" & strFeeType & PrintISOTime(Now, TimeZone) & "AH" & strDueDate & "|AB" & colFields("AB")
            If intQueueLen > 0 Then
                strOutput = strOutput & "|CF" & CStr(intQueueLen)
            End If
            'Dua ngay het han ra truoc |AB
            'If Not strDueDate = "" Then
            '    strOutput = strOutput & "|AH" & strDueDate
            'End If
            If Not strRecallDate = "" Then
                strOutput = strOutput & "|CJ" & strDueDate
            End If
            If Not strPickUpDate = "" Then
                strOutput = strOutput & "|CM" & strPickUpDate
            End If
            '   strOutput = strOutput & "|AJ" & strTitle & "|BG" & strOwner & "|BH" & CurrencyUnit & "|BV" & strFee & "|AQ" & strPermanentLoc
            strOutput = strOutput & "|AJ" & strTitle & "|BG" & LibName & "|CK001|AQ" & strPermanentLoc
            If Not strNote = "" Then
                strOutput = strOutput & "|AF" & strNote
            End If

            '1808000119980723 115615CF00000|ABItemBook|AJTitle For Item Book|SIP2 Developer’s Guide June 7, 1999 CK003|AQPermanent Location For ItemBook, Language 1|APCurrent Location ItemBook|CHFree-form text With New item Property|AY0AZC05B
        Else
            strOutput = "18010001" & PrintISOTime(Now, TimeZone) & "AB" & colFields("AB") & "|AJ|AFItem identifier does not exist"
        End If

        'ServercheckSip2: Ten truong: TDMU, Noi (Location): VN,DKCB: TK02412 Dang tai lieu Book (CK001). Han tra 4/9/2021, Note: This book belongs to TDM's Library. Sercurity Marker:01 (None) voi 02 (M Tattle-Tape Security Strip)
        'Good
        '1803010120210327    045201AH4/1/2021|ABCN2010004967|AJNhung van de co ban ve bao ve an ninh quoc gia va giu gin trat tu an toan xa hoi :  Giao trinh dung cho bac Dai hoc CAND  1|BGThis book belongs to QA's Library|CK001|AY1AZBBEB




        ItemInformation = strOutput
    End Function

    Public Function FeePaid(ByVal req As String) As String
        Dim colFields As New Collection
        Dim blnPaymentAccepted As Boolean
        Dim strTransDate, strSQL, strMessage, strReason As String
        Dim lngPatronID As Long
        Dim curRate As Double
        Dim arrTemp() As String = {"AO", "AA", "AD", "CG", "BK", "AC", "BV"}
        Dim tblResult As DataTable
        Dim strFeeType, strPaymentType, strCurrency As String
        Dim objDCommonBusiness As New clsDCommonBusiness
        objDCommonBusiness.DBServer = strDBServer
        objDCommonBusiness.ConnectionString = strConnectionstring
        Call objDCommonBusiness.Initialize()
        Dim objDAccountTransaction As New clsDAccountTransaction
        objDAccountTransaction.DBServer = strDBServer
        objDAccountTransaction.ConnectionString = strConnectionstring
        Call objDAccountTransaction.Initialize()

        strTransDate = Left(req, 18)
        strFeeType = Mid(req, 19, 2)
        strPaymentType = Mid(req, 21, 2)
        strCurrency = Mid(req, 23, 3)
        req = Right(req, Len(req) - 25)
        Call LoadFields(arrTemp, colFields, req)
        Dim dtmTransDate As Date
        dtmTransDate = ParseISOTime(strTransDate)

        blnPaymentAccepted = False
        If IsNumeric(colFields("BV")) Then
            If CDbl(colFields("BV")) > 0 Then
                objDCommonBusiness.CurrencyCode = strCurrency
                tblResult = objDCommonBusiness.GetCurrency
                If tblResult.Rows.Count > 0 Then
                    curRate = CDbl(tblResult.Rows(0).Item("Rate"))
                Else
                    curRate = 1
                End If
                tblResult = Nothing
                tblResult = Me.GetData("SELECT ID FROM CIR_PATRON WHERE Code = '" & colFields("AA") & "'")
                If tblResult.Rows.Count > 0 Then
                    Select Case strFeeType
                        Case "01"
                            strReason = "Other/unknown"
                        Case "02"
                            strReason = "Administrative"
                        Case "03"
                            strReason = "Damage"
                        Case "04"
                            strReason = "Overdue"
                        Case "05"
                            strReason = "Processing"
                        Case "06"
                            strReason = "Rental"
                        Case "07"
                            strReason = "Replacement"
                        Case "08"
                            strReason = "Computer access charge"
                        Case "09"
                            strReason = "Hold fee"
                        Case Else
                            strReason = "Other/unknown"
                    End Select
                    strReason = "Fee type: " & strReason
                    Select Case strPaymentType
                        Case "00"
                            strReason = strReason & ". Payment type: Cash"
                        Case "01"
                            strReason = strReason & ". Payment type: VISA"
                        Case "02"
                            strReason = strReason & ". Payment type: Credit card"
                        Case Else
                            strReason = strReason & ". Payment type: Not specified"
                    End Select
                    objDAccountTransaction.PatronCode = colFields("AA")
                    objDAccountTransaction.Amount = colFields("BV")
                    objDAccountTransaction.CreatedDate = ConvertDateBack(dtmTransDate)
                    objDAccountTransaction.Reason = strReason
                    objDAccountTransaction.Rate = curRate
                    objDAccountTransaction.Currency = strCurrency
                    'lelinh
                    'objDAccountTransaction.CreateNewFine()
                    If Not strDBServer = "ORACLE" Then
                        strSQL = "SELECT isnull(SUM(Amount*Rate),0) as TotalDebt FROM CIR_FINE WHERE PatronCode = '" & colFields("AA") & "'"
                    Else
                        strSQL = "SELECT NVL(SUM(Amount*Rate),0) as TotalDebt FROM CIR_FINE WHERE PatronCode = '" & colFields("AA") & "'"
                    End If
                    tblResult = Nothing
                    tblResult = Me.GetData(strSQL)
                    Me.ExecuteQuery("UPDATE CIR_PATRON SET Debt = " & CStr(-1 * tblResult.Rows(0).Item("TotalDebt")) & " WHERE (CODE) = '" & CStr(colFields("AA")) & "'")
                    blnPaymentAccepted = True
                End If
            Else
                strMessage = "Invalid fee amount"
            End If
        Else
            strMessage = "Invalid fee amount"
        End If
        If blnPaymentAccepted Then
            FeePaid = "38Y" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA")
        Else
            FeePaid = "38N" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AF" & strMessage
        End If
        objDCommonBusiness.Dispose()
        objDCommonBusiness = Nothing
        objDAccountTransaction.Dispose()
        objDAccountTransaction = Nothing
    End Function


    Public Function EndPatronSession(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strTransDate As String
        Dim arrTemp() As String = {"AO", "AA", "AD", "AC"}
        strTransDate = Left(req, 18)
        req = Right(req, Len(req) - 18)
        Call LoadFields(arrTemp, colFields, req)
        EndPatronSession = "36Y" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AFEnd Session"

        'SIP 2 Server
        '36Y20210326    060726AO|AA1725202010040|AY3AZF5FF
    End Function

    Public Function CheckIn(ByVal req As String) As String
        Dim colFields As New Collection
        Dim arrTemp() As String = {"AP", "AO", "AB", "AC", "CH", "BI"}
        Dim tblResult As DataTable
        Dim strReturnDate, strTransDate, strNoBlock, strTitle, strPermanentLoc As String
        Dim dtmReturnDate As Date
        Dim strMessageOut As String = ""
        Dim objDLoanTransaction As New clsDLoanTransaction
        Dim intRetval As Integer

        objDLoanTransaction.DBServer = strDBServer
        objDLoanTransaction.ConnectionString = strConnectionstring
        Call objDLoanTransaction.Initialize()

        strNoBlock = Left(req, 1)
        strTransDate = Mid(req, 2, 18)
        strReturnDate = Mid(req, 20, 18)
        req = Right(req, Len(req) - 37)
        Call LoadFields(arrTemp, colFields, req)
        dtmReturnDate = ParseISOTime(strReturnDate)

        ' check copynumber and patron
        objDLoanTransaction.UserID = intUserID
        objDLoanTransaction.PatronCode = ""
        objDLoanTransaction.CopyNumber = colFields("AB")
        intRetval = objDLoanTransaction.CheckCopyNumber()
        Select Case intRetval
            Case 0
                strMessageOut = "Item identifier does not exist"
            Case 1 ' Copynumber doesn't exists
                strMessageOut = "Item identifier does not exist"
            Case 2 ' Copynumber is locked or not in circulation
                strMessageOut = "Copynumber is locked or not in circulation"
            Case 3 ' Copynumber is on loan
                'Ok for check in if has error
                strMessageOut = "Check in error"
            Case 4 ' Copynumber is on hold
                strMessageOut = "Copynumber is on hold"
            Case 5 ' Librarian has not permission to manage location of the CopyNumber
                strMessageOut = "Librarian has not permission to manage location of the CopyNumber"
            Case 6 ' Librarian has not permission to manage location of Patron
                strMessageOut = "Librarian has not permission to manage location of Patron"
            Case 9 ' Librarian has not permission to manage location of Patron
                strMessageOut = "Copynumber is not in circulation"
            Case 12 ' Sach qua han khong duoc tra
                strMessageOut = "Sach qua han, vui long lien he Quay."
        End Select
        If intRetval <> 3 Then
            'Truong hop khong cho tra
            If intRetval = 9 Then
                tblResult = Me.GetContent(colFields("AB"))
                strTitle = ""
                If Not tblResult.Rows.Count = 0 Then
                    strTitle = Trim(TrimSubFieldCodes(CutVietnameseAccent(TheDisplayOne(tblResult.Rows(0).Item("Content")))))
                End If
                If Not UCase(strNoBlock) = "Y" Then
                    CheckIn = "100NUY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AB" & colFields("AB") & "|AQ" & strPermanentLoc & "|AFItem is not listed as checkout|AJ" & strTitle
                Else
                    tblResult = Me.GetData("select tai_lieu_id as ItemID,InUsed from ma_xep_gia where (ma_xep_gia)='" & CStr(colFields("AB")) & "'")
                    If Not tblResult.Rows(0).Item("InUsed") = 0 Then
                        Me.ExecuteQuery("UPDATE so_luong_an_pham SET So_ban_roi = So_ban_roi + 1 WHERE So_ban_roi < So_ban AND Tai_lieu_id = " & tblResult.Rows(0).Item("ItemID"))
                        Me.ExecuteQuery("UPDATE ma_xep_gia SET InUsed = 0 WHERE (ma_xep_gia)='" & CStr(colFields("AB")) & "'")
                    End If
                    CheckIn = "101YUN" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AB" & colFields("AB") & "|AQ" & strPermanentLoc & "|AJ" & strTitle
                End If
            Else
                CheckIn = "100NUY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AB" & colFields("AB") & "|AQ|AF" & strMessageOut
            End If
            Call objDLoanTransaction.Dispose()
            objDLoanTransaction = Nothing
            Exit Function
        End If
        ' start checkin
        objDLoanTransaction.CheckInDate = ConvertDateBack(dtmReturnDate)
        intRetval = objDLoanTransaction.CheckIn(colFields("AB"), 0)
        tblResult = Me.GetContent(colFields("AB"))
        strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblResult.Rows(0).Item("Content")))
        strPermanentLoc = CutVietnameseAccent(tblResult.Rows(0).Item("Code") & "") & "--" & CutVietnameseAccent(tblResult.Rows(0).Item("Symbol") & "")
        If intRetval = 0 Then
            'check in ok
            CheckIn = "101YUN" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AB" & colFields("AB") & "|AQ" & strPermanentLoc & "|AJ" & strTitle
        Else
            CheckIn = "100NUY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AB" & colFields("AB") & "|AQ" & strPermanentLoc & "|AJ" & strTitle & "|AF" & strMessageOut
        End If
        Call objDLoanTransaction.Dispose()
        objDLoanTransaction = Nothing

        'ServercheckSip2: Ten truong: TDMU, Noi (Location): VN,DKCB: TK02412 Dang tai lieu Book (CK001). Han tra 4/9/2021, Note: This book belongs to TDM's Library. Sercurity Marker:01 (None) voi 02 (M Tattle-Tape Security Strip)
        '101YNN20210326    055558AO|ABTK02412|AQ|AH|AJnhan de tai lieu 13|CK001|AY0AZEB31
    End Function

    Private Function CreateHold(ByVal strCardNo As String, ByVal strPassword As String, ByVal strValidDateReq As String, ByVal strCopyNumber As String) As Integer
        Me.OpenConnection()
        Select Case UCase(strDBServer)
            Case "SQLSERVER"
                With sqlCommand
                    .CommandText = "SP_CIR_HOLDING_CREATE"
                    .CommandType = CommandType.StoredProcedure
                    Try
                        With .Parameters
                            .Add(New SqlParameter("@strCardNo", SqlDbType.VarChar, 50)).Value = strCardNo
                            .Add(New SqlParameter("@strPassWord", SqlDbType.VarChar, 50)).Value = strPassword
                            .Add(New SqlParameter("@strValidDate", SqlDbType.VarChar, 30)).Value = strValidDateReq
                            .Add(New SqlParameter("@strCopyNumber", SqlDbType.VarChar, 32)).Value = strCopyNumber
                            .Add(New SqlParameter("@intRET", SqlDbType.Int)).Direction = ParameterDirection.Output
                        End With
                        .ExecuteNonQuery()
                        CreateHold = .Parameters("@intRET").Value
                    Catch sqlClientEx As SqlException
                        strErrorMsg = sqlClientEx.Message.ToString
                        intErrorCode = sqlClientEx.Number
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
            Case "ORACLE"
                With oraCommand
                    .CommandText = "CIRCULATION.SP_CIR_HOLDING_CREATE"
                    .CommandType = CommandType.StoredProcedure
                    Try
                        With .Parameters
                            .Add(New OracleParameter("strCardNo", OracleType.VarChar, 50)).Value = strCardNo
                            .Add(New OracleParameter("strPassWord", OracleType.VarChar, 50)).Value = strPassword
                            .Add(New OracleParameter("strValidDate", OracleType.VarChar, 10)).Value = strValidDateReq
                            .Add(New OracleParameter("strCopyNumber", OracleType.VarChar, 32)).Value = strCopyNumber
                            .Add(New OracleParameter("intRET", OracleType.Number)).Direction = ParameterDirection.Output
                        End With
                        .ExecuteNonQuery()
                        CreateHold = .Parameters("intRET").Value
                    Catch OraEx As OracleException
                        strErrorMsg = OraEx.Message.ToString
                        intErrorCode = OraEx.Code
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
        End Select
        Me.CloseConnection()
    End Function

    Private Function UpdateHold(ByVal strCardNo As String, ByVal strPassword As String, ByVal strValidDateReq As String, ByVal strCopyNumber As String) As Integer
        Me.OpenConnection()
        Select Case UCase(strDBServer)
            Case "SQLSERVER"
                With sqlCommand
                    .CommandText = "SP_CIR_HOLDING_UPDATE"
                    .CommandType = CommandType.StoredProcedure
                    Try
                        With .Parameters
                            .Add(New SqlParameter("@strCardNo", SqlDbType.VarChar, 50)).Value = strCardNo
                            .Add(New SqlParameter("@strPassWord", SqlDbType.VarChar, 50)).Value = strPassword
                            .Add(New SqlParameter("@strValidDate", SqlDbType.VarChar, 30)).Value = strValidDateReq
                            .Add(New SqlParameter("@strCopyNumber", SqlDbType.VarChar, 32)).Value = strCopyNumber
                            .Add(New SqlParameter("@intRET", SqlDbType.Int)).Direction = ParameterDirection.Output
                        End With
                        .ExecuteNonQuery()
                        UpdateHold = .Parameters("@intRET").Value
                    Catch sqlClientEx As SqlException
                        strErrorMsg = sqlClientEx.Message.ToString
                        intErrorCode = sqlClientEx.Number
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
            Case "ORACLE"
                With oraCommand
                    .CommandText = "CIRCULATION.SP_CIR_HOLDING_UPDATE"
                    .CommandType = CommandType.StoredProcedure
                    Try
                        With .Parameters
                            .Add(New OracleParameter("strCardNo", OracleType.VarChar, 50)).Value = strCardNo
                            .Add(New OracleParameter("strPassWord", OracleType.VarChar, 50)).Value = strPassword
                            .Add(New OracleParameter("strValidDate", OracleType.VarChar, 10)).Value = strValidDateReq
                            .Add(New OracleParameter("strCopyNumber", OracleType.VarChar, 32)).Value = strCopyNumber
                            .Add(New OracleParameter("intRET", OracleType.Number)).Direction = ParameterDirection.Output
                        End With
                        .ExecuteNonQuery()
                        UpdateHold = .Parameters("intRET").Value
                    Catch OraEx As OracleException
                        strErrorMsg = OraEx.Message.ToString
                        intErrorCode = OraEx.Code
                    Finally
                        .Parameters.Clear()
                    End Try

                End With
        End Select
        Me.CloseConnection()
    End Function

    Private Function DeleteHold(ByVal strCardNo As String, ByVal strPassword As String, ByVal strCopyNumber As String) As Integer
        Me.OpenConnection()
        Select Case UCase(strDBServer)
            Case "SQLSERVER"
                With sqlCommand
                    .CommandText = "SP_CIR_HOLDING_DELETE"
                    .CommandType = CommandType.StoredProcedure
                    Try
                        With .Parameters
                            .Add(New SqlParameter("@strCardNo", SqlDbType.VarChar, 50)).Value = strCardNo
                            .Add(New SqlParameter("@strPassWord", SqlDbType.VarChar, 50)).Value = strPassword
                            .Add(New SqlParameter("@strCopyNumber", SqlDbType.VarChar, 32)).Value = strCopyNumber
                            .Add(New SqlParameter("@intRET", SqlDbType.Int)).Direction = ParameterDirection.Output
                        End With
                        .ExecuteNonQuery()
                        DeleteHold = .Parameters("@intRET").Value
                    Catch sqlClientEx As SqlException
                        strErrorMsg = sqlClientEx.Message.ToString
                        intErrorCode = sqlClientEx.Number
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
            Case "ORACLE"
                With oraCommand
                    .CommandText = "CIRCULATION.SP_CIR_HOLDING_DELETE"
                    .CommandType = CommandType.StoredProcedure
                    Try
                        With .Parameters
                            .Add(New OracleParameter("strCardNo", OracleType.VarChar, 50)).Value = strCardNo
                            .Add(New OracleParameter("strPassWord", OracleType.VarChar, 50)).Value = strPassword
                            .Add(New OracleParameter("strCopyNumber", OracleType.VarChar, 32)).Value = strCopyNumber
                            .Add(New OracleParameter("intRET", OracleType.Number)).Direction = ParameterDirection.Output
                        End With
                        .ExecuteNonQuery()
                        DeleteHold = .Parameters("intRET").Value
                    Catch OraEx As OracleException
                        strErrorMsg = OraEx.Message.ToString
                        intErrorCode = OraEx.Code
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
        End Select
        Me.CloseConnection()
    End Function

    Public Function Hold(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strHoldMode, strTransDate, strOutput, strExpDate As String
        Dim dtmTransDate, dtmExpDate As Date
        Dim intHoldQuota As Integer
        Dim dblDocID, dblID As Double
        Dim arrTemp() As String = {"BW", "BS", "BY", "AO", "AA", "AD", "AB", "AJ", "AC", "BO"}
        Dim tblResult As DataTable
        Dim intRetval As Integer = 0
        Dim strMessage As String = ""

        strHoldMode = Left(req, 1)
        strTransDate = Mid(req, 2, 18)
        req = Right(req, Len(req) - 19)
        Call LoadFields(arrTemp, colFields, req)
        dtmExpDate = ParseISOTime(colFields("BW"))
        Select Case strHoldMode
            Case "+" ' Add
                intRetval = CreateHold(colFields("AA"), colFields("AD"), ConvertDateBack(dtmExpDate), colFields("AB"))
                Select Case intRetval
                    Case 0
                        strMessage = "Hold request is added"
                    Case 1
                        strMessage = "Patron information is not valid"
                    Case 2
                        strMessage = "Patron card is expired"
                    Case 3
                        strMessage = "You have already placed a hold request on this item"
                    Case 4
                        strMessage = "You are mot allowed to place more hold requests"
                    Case 5
                        strMessage = "You are mot allowed to place more hold requests"
                    Case 6
                        strMessage = "Item identifier or title identifier is not valid"
                End Select
                If intRetval = 0 Then
                    strOutput = "161N" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & colFields("AJ") & "|AF" & strMessage
                Else
                    strOutput = "160N" & PrintISOTime(Now, TimeZone) & "|AO" & colFields("AO") & "|AA" & colFields("AA") & "|AF" & strMessage
                End If
            Case "-" ' Delete
                intRetval = DeleteHold(colFields("AA"), colFields("AD"), colFields("AB"))
                Select Case intRetval
                    Case 0
                        strMessage = "Hold request is removed"
                    Case 1
                        strMessage = "Patron information is not valid"
                    Case 2
                        strMessage = "Patron card is expired"
                    Case 3
                        strMessage = "No hold request matched"
                    Case 4
                        strMessage = "You are mot allowed to place more hold requests"
                    Case 5
                        strMessage = "You are mot allowed to place more hold requests"
                    Case 6
                        strMessage = "Item identifier or title identifier is not valid"
                End Select
                If intRetval = 0 Then
                    strOutput = "161Y" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AF" & strMessage
                Else
                    strOutput = "160N" & PrintISOTime(Now, TimeZone) & "|AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AF" & strMessage
                End If

            Case "*" ' Update
                intRetval = UpdateHold(colFields("AA"), colFields("AD"), ConvertDateBack(dtmExpDate), colFields("AB"))
                Select Case intRetval
                    Case 0
                        strMessage = "Hold request is updated"
                    Case 1
                        strMessage = "Patron information is not valid"
                    Case 2
                        strMessage = "Patron card is expired"
                    Case 3
                        strMessage = "No hold request matched"
                    Case 4
                        strMessage = "You are mot allowed to place more hold requests"
                    Case 5
                        strMessage = "You are mot allowed to place more hold requests"
                    Case 6
                        strMessage = "Item identifier or title identifier is not valid"
                End Select
                If intRetval = 0 Then
                    strOutput = "161N" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AF" & strMessage
                Else
                    strOutput = "160N" & PrintISOTime(Now, TimeZone) & "|AO" & colFields("AO") & "|AA" & colFields("AA") & "|AF" & strMessage
                End If
        End Select
        Hold = strOutput
    End Function

    Public Function SCStatus(ByVal req As String) As String
        SCStatus = "98YYYYYY500003" & PrintISOTime(Now, TimeZone) & "2.00AO|AM" & LibName & "|BXYYYYYYYYYYYYYYYY|ANVN"

        'ServercheckSip2: Ten truong: TDMU, Noi (Location): VN
        '98YYYYYY50000320210326    0538312.00AO|AMTDMU|BXYYYYYYYYYYYYYYYY|ANVN|AY6AZEB3D

    End Function
    Private Sub BlockPatronUpdate(ByVal strCodePatron As String, ByVal strFromDate As String, ByVal intNumberOfDays As Integer)
        Me.OpenConnection()
        Select Case UCase(strDBServer)
            Case "SQLSERVER"
                With sqlCommand
                    .CommandText = "SP_CIR_PATRON_BLOCK_INSERT"
                    .CommandType = CommandType.StoredProcedure
                    With .Parameters
                        .Add(New SqlParameter("@strCodePatron", SqlDbType.VarChar, 50)).Value = strCodePatron
                        .Add(New SqlParameter("@strFromDate", SqlDbType.VarChar, 30)).Value = strFromDate
                        .Add(New SqlParameter("@intNumberOfDays", SqlDbType.Int)).Value = intNumberOfDays
                    End With
                    Try
                        .ExecuteNonQuery()
                    Catch sqlClientEx As SqlException
                        strErrorMsg = sqlClientEx.Message.ToString
                        intErrorCode = sqlClientEx.Number
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
            Case "ORACLE"
                With oraCommand
                    .CommandText = "CIRCULATION.SP_CIR_PATRON_BLOCK_INSERT"
                    .CommandType = CommandType.StoredProcedure
                    With .Parameters
                        .Add(New OracleParameter("strCodePatron", OracleType.VarChar, 50)).Value = strCodePatron
                        .Add(New OracleParameter("strFromDate", OracleType.VarChar, 30)).Value = strFromDate
                        .Add(New OracleParameter("intNumberOfDays", OracleType.Number)).Value = intNumberOfDays
                    End With
                    Try
                        .ExecuteNonQuery()
                    Catch OraEx As OracleException
                        strErrorMsg = OraEx.Message.ToString
                        intErrorCode = OraEx.Code
                    Finally
                        .Parameters.Clear()
                    End Try
                End With
        End Select
        Me.CloseConnection()
    End Sub

    Public Function BlockPatron(ByVal req As String) As String
        Dim colFields As New Collection
        Dim arrTemp() As String = {"AO", "AL", "AA", "AC"}
        Dim strCardRetained, strTransDate, strMessage, strBlockDate As String
        Dim dtmTransDate As Date
        Dim lngPatronID As Long

        strCardRetained = Left(req, 1)
        strTransDate = Mid(req, 2, 18)
        dtmTransDate = ParseISOTime(strTransDate)
        strBlockDate = ConvertDateBack(dtmTransDate)
        req = Right(req, Len(req) - 19)
        Call LoadFields(arrTemp, colFields, req)
        Dim tblResult As DataTable
        strMessage = ""
        If UCase(strCardRetained) = "Y" Then
            strMessage = "Card retained by ShelfCheck system"
        End If
        Call BlockPatronUpdate(colFields("AA"), strBlockDate, 0)
        BlockPatron = PatronStatusRequest("001" & strTransDate & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AC" & colFields("AC") & "|AD", False, strMessage)
    End Function

    Private Function CreateID(ByVal strTableName As String, Optional ByVal strFieldName As String = "ID") As Integer
        Dim strsql As String
        If strDBServer = "ORACLE" Then
            strsql = "SELECT NVL(Max(" & strFieldName & "), 0) + 1 as ID FROM " & strTableName
        Else
            strsql = "SELECT ISNULL(Max(" & strFieldName & "), 0) + 1 as ID FROM " & strTableName
        End If
        Dim tblTemp As DataTable
        tblTemp = Me.GetData(strsql)
        CreateID = tblTemp.Rows(0).Item("ID")
    End Function

    Public Function CheckOut(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strNBDueDate, strTransDate, strNoBlock, strCheckOutDate, strDueDate, strTitle, strResponse As String
        Dim strNote, strSCRenewalPolicy As String
        Dim dtmNBDueDate, dtmCheckOutDate, dtmTransDate As Date
        Dim arrTemp() As String = {"AO", "AA", "AD", "AB", "BI", "AC", "CH", "BO"}
        Dim tblResult As DataTable
        Dim strPatronCode As String
        Dim strCopyNumber As String
        Dim strMessageOut As String = ""
        Dim objDLoanTransaction As New clsDLoanTransaction

        objDLoanTransaction.DBServer = strDBServer
        objDLoanTransaction.ConnectionString = strConnectionstring
        Call objDLoanTransaction.Initialize()

        strSCRenewalPolicy = Left(req, 1)
        strNoBlock = Mid(req, 2, 1)
        strTransDate = Mid(req, 3, 18)
        'Check lai nha
        '  strNBDueDate = Mid(req, 21, 18)
        strNBDueDate = ""
        If strNBDueDate.Trim <> "" Then
            dtmNBDueDate = ParseISOTime(strNBDueDate)
        Else
            strNBDueDate = ""
        End If

        dtmCheckOutDate = ParseISOTime(strTransDate)
        dtmTransDate = dtmCheckOutDate
        req = Right(req, Len(req) - 38)
        Call LoadFields(arrTemp, colFields, req)
        Dim intRetval As Integer
        strPatronCode = colFields("AA")
        strCopyNumber = colFields("AB")

        ' check copynumber and patron
        objDLoanTransaction.UserID = intUserID
        objDLoanTransaction.PatronCode = strPatronCode
        objDLoanTransaction.CopyNumber = strCopyNumber
        intRetval = objDLoanTransaction.CheckCopyNumber()
        Select Case intRetval
            Case 0
                ' Ok for checkout
                strMessageOut = "Error Unknow"
            Case 1 ' Copynumber doesn't exists
                strMessageOut = "Copynumber does not exists"
            Case 2 ' Copynumber is locked or not in circulation
                strMessageOut = "Copynumber is locked or not in circulation"
            Case 3 ' Copynumber is on loan
                strMessageOut = "Copynumber is on loan"
            Case 4 ' Copynumber is on hold
                strMessageOut = "Copynumber is on hold"
            Case 5 ' Librarian has not permission to manage location of the CopyNumber
                strMessageOut = "Librarian has not permission to manage location of the CopyNumber"
            Case 6 ' Librarian has not permission to manage location of Patron
                strMessageOut = "Lib has not permission to manage location of Patron or Patron does not exist."
            Case 7 ' Patron is locked
                strMessageOut = "Patron is locked"
            Case 8 ' Patron is locked
                strMessageOut = "Sach doc tai cho KHONG duoc muon ve."
            Case 9 ' Patron is locked
                strMessageOut = "Tai lieu da muon, KHONG duoc muon lai trong vong 24h."
            Case 10 ' Patron is locked
                strMessageOut = "Muon qua han ngach."
            Case 11 ' Patron is locked
                strMessageOut = "Tai lieu duoc muon cung nhan de voi cuon dang muon."
        End Select
        ' information is incorrect
        If strNoBlock.ToUpper = "N" Then
            If intRetval > 0 Then
                strResponse = "120NNN" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH" & PrintISOTime(Now, TimeZone) & "|AF" & strMessageOut
                Call objDLoanTransaction.Dispose()
                objDLoanTransaction = Nothing
                CheckOut = strResponse
                Exit Function
            End If
            ' start Check out
            objDLoanTransaction.PatronCode = strPatronCode
            objDLoanTransaction.CopyNumber = strCopyNumber
            objDLoanTransaction.LoanMode = 1
            If Not strNBDueDate = "" Then
                objDLoanTransaction.DueDate = ConvertDateBack(dtmNBDueDate)
            Else
                objDLoanTransaction.DueDate = ""
            End If
            objDLoanTransaction.CheckOutDate = ConvertDateBack(dtmCheckOutDate)
            intRetval = objDLoanTransaction.CheckOut(1)

            If intRetval > 0 Then
                'Check out Ok
                tblResult = Me.GetContent(strCopyNumber)
                strTitle = ""
                strNote = ""
                If Not tblResult.Rows.Count = 0 Then
                    strTitle = Trim(TrimSubFieldCodes(CutVietnameseAccent(TheDisplayOne(tblResult.Rows(0).Item("Content")))))
                    strNote = CutVietnameseAccent(tblResult.Rows(0).Item("Note"))
                End If
                '  strResponse = "121NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & strTitle & "|AH" & strNBDueDate & "|CH" & strNote
                If strNBDueDate = "" Then
                    'strNBDueDate = PrintISOTime(Now.AddDays(10), TimeZone)
                    'Check lai nha
                    strNBDueDate = Now.AddDays(10).ToString("dd/MM/yyyy")
                End If
                'strResponse = "121NNY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & strTitle & "|AH" & strNBDueDate & "|CH" & strNote
                '121NNY20210326    165632AO|AA100220066|AB100012024|AJLuat an ninh quAc gia: QuAc hoi thong qua ngay 03 thAng 12 nam 2004|AH4/19/2021|CK001|AY4AZE9F9
                '121NNY20210326    171816AO|AA100220066|AB100012024|AJLuat an ninh quAc gia: QuAc hoi thong qua ngay 03 thAng 12 nam 2004|AH4/19/2021|CK001|CH|AY|AZD5C7

                strResponse = "121NNY" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & strTitle & "|AH" & strNBDueDate & "|CK001|CH" & strNote
            Else
                'Check out error
                strResponse = "120NNN" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH" & PrintISOTime(Now, TimeZone) & "|AF" & strMessageOut
            End If
            strErrorMsg = objDLoanTransaction.ErrorMsg
        Else
            If Not (intRetval = 0 Or intRetval = 6) Then
                strResponse = "120NNN" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH" & PrintISOTime(Now, TimeZone) & "|AF" & strMessageOut
                Call objDLoanTransaction.Dispose()
                objDLoanTransaction = Nothing
                CheckOut = strResponse
                Exit Function
            End If
            Dim lngPatronID As Integer
            Dim blnCheckedOut As Boolean
            Dim strExpiredDate As String
            ' Verify if the item is already checked out
            If intRetval = 3 Then
                blnCheckedOut = True
            Else
                blnCheckedOut = False
            End If
            If Not strDBServer = "ORACLE" Then
                strCheckOutDate = "CONVERT(datetime, '" & ConvertDateBack(dtmTransDate) & "', 103)"
                strDueDate = "CONVERT(datetime, '" & ConvertDateBack(dtmNBDueDate) & "', 103)"
            Else
                strCheckOutDate = "CONVERTDATE('" & ConvertDateBack(dtmTransDate) & "')"
                strDueDate = "CONVERTDATE('" & ConvertDateBack(dtmNBDueDate) & "')"
            End If

            ' If patron record does not exist, then add new patron record
            tblResult = Me.GetData("SELECT ID FROM Ban_doc WHERE (So_the) = '" & strPatronCode & "'")
            If Not tblResult.Rows.Count = 0 Then
                lngPatronID = CLng(tblResult.Rows(0).Item("ID"))
            Else
                '================================ Tao moi Ban doc----------------------------------

                '    If Not strDBServer = "ORACLE" Then
                '        strExpiredDate = "CONVERT(datetime, '" & ConvertDateBack(DateAdd(DateInterval.Year, 4, dtmTransDate)) & "', 103)"
                '    Else
                '        strExpiredDate = "CONVERTDATE('" & ConvertDateBack(DateAdd(DateInterval.Year, 4, dtmTransDate)) & "')"
                '    End If
                '    lngPatronID = Me.CreateID("CIR_PATRON")
                '    tblResult = Me.GetData("SELECT ID FROM CIR_DIC_ETHNIC")
                '    Dim intEthnicID As Integer = tblResult.Rows(0).Item("ID")
                '    tblResult = Me.GetData("SELECT ID FROM CIR_DIC_EDUCATION")
                '    Dim intEducationID As Integer = tblResult.Rows(0).Item("ID")
                '    tblResult = Me.GetData("SELECT ID FROM CIR_DIC_OCCUPATION")
                '    Dim intOccupationID As Integer = tblResult.Rows(0).Item("ID")
                '    Me.ExecuteQuery("INSERT INTO CIR_PATRON (ID,FirstName,MiddleName,LastName, Code,  ValidDate,ExpiredDate, Sex, Debt, Password, Status,PatronGroupID,Note,EthnicID,EducationID,OccupationID) VALUES (" & lngPatronID & ", 'Unknown','','Patron', '" & strPatronCode & "', " & strCheckOutDate & ", " & strExpiredDate & ", 0, 0, '" & colFields("AD") & "', 1,1,'Patron is created by Sip2Server'," & intEthnicID & "," & intEducationID & "," & intOccupationID & ")")
            End If
            ' If the item is already checked out
            If blnCheckedOut Then
                tblResult = Me.GetData("SELECT So_the_id as PatronID, ID FROM An_pham_cho_muon WHERE Ma_xep_gia = '" & colFields("AB") & "'")
                If Not tblResult.Rows.Count = 0 Then
                    ' If the item is already checked out by the same patron then actually it is renewed
                    If CLng(tblResult.Rows(0).Item("PatronID")) = CLng(lngPatronID) Then
                        '   Me.ExecuteQuery("UPDATE CIR_LOAN SET DueDate = " & strDueDate & ", RenewCount = RenewCount + 1 WHERE ID = " & tblResult.Rows(0).Item("ID"))
                        tblResult = Me.GetContent(strCopyNumber)
                        strTitle = ""
                        strNote = ""
                        If Not tblResult.Rows.Count = 0 Then
                            strTitle = Trim(TrimSubFieldCodes(CutVietnameseAccent(TheDisplayOne(tblResult.Rows(0).Item("Content")))))
                            strNote = CutVietnameseAccent(tblResult.Rows(0).Item("Note"))
                        End If
                        strResponse = "121YUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & strTitle & "|AH" & strNBDueDate & "|CH" & strNote
                    End If
                End If
            Else
                objDLoanTransaction.PatronCode = strPatronCode
                objDLoanTransaction.CopyNumber = strCopyNumber
                objDLoanTransaction.LoanMode = 1
                If Not strDueDate = "NULL" Then
                    objDLoanTransaction.DueDate = ConvertDateBack(dtmNBDueDate)
                Else
                    objDLoanTransaction.DueDate = "NULL"
                End If
                objDLoanTransaction.CheckOutDate = ConvertDateBack(dtmCheckOutDate)
                intRetval = objDLoanTransaction.CheckOut(1)

                If intRetval > 0 Then
                    'Check out Ok
                    tblResult = Me.GetContent(strCopyNumber)
                    strTitle = ""
                    strNote = ""
                    If Not tblResult.Rows.Count = 0 Then
                        strTitle = Trim(TrimSubFieldCodes(CutVietnameseAccent(TheDisplayOne(tblResult.Rows(0).Item("Content")))))
                        strNote = CutVietnameseAccent(tblResult.Rows(0).Item("Note"))
                    End If
                    strResponse = "121NUU" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ" & strTitle & "|AH" & strNBDueDate & "|CH" & strNote
                Else
                    'Check out error
                    strResponse = "120NNN" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & colFields("AA") & "|AB" & colFields("AB") & "|AJ|AH" & PrintISOTime(Now, TimeZone) & "|AF" & strMessageOut
                End If
            End If
        End If
        Call objDLoanTransaction.Dispose()
        objDLoanTransaction = Nothing
        'ServercheckSip2: Ten truong: TDMU, Noi (Location): VN, Dang tai lieu Book. Han tra 4/9/2021
        '121NNY20210326    054408AO|AA123|ABTK02412|AJnhan de tai lieu 1|AH4/9/2021|CK001|AY2AZE951


        '121NNY20210326    165449AO|AA100220066|AB100012024|AJLuat an ninh quAc gia: QuAc hoi thong qua ngay 03 thAng 12 nam 2004|AH20210326    165314|CH
        '121NNY20210326    165632AO|AA100220066|AB100012024|AJLuat an ninh quAc gia: QuAc hoi thong qua ngay 03 thAng 12 nam 2004|AH4/9/2021|CK001|AY4AZE9F9

        CheckOut = strResponse
    End Function


    Public Function PatronStatusRequest(ByVal req As String, ByVal blnPWDRequired As Boolean, Optional ByVal strMesage As String = "") As String
        Dim colFields As New Collection
        Dim objXe As XCrypt.XCryptEngine
        objXe = New XCrypt.XCryptEngine(XCrypt.XCryptEngine.AlgorithmType.MD5)

        Dim strLang, strTransDate, strResponse As String
        Dim arrTemp() As String = {"AO", "AA", "AD", "AC"}
        Dim tblResult As DataTable

        strLang = Left(req, 3)
        strTransDate = Mid(req, 4, 18)
        req = Right(req, Len(req) - 21)
        Call LoadFields(arrTemp, colFields, req)

        tblResult = Me.GetData("SELECT so_the as Code,Ho_ten as Fullname,Mat_khau as Password,Debt FROM Ban_doc WHERE (So_the) = '" & Replace(colFields("AA"), "'", "''") & "'")
        If tblResult.Rows.Count = 0 Then
            strResponse = "24              001" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA|AE|BLN"
        Else
            If blnPWDRequired Then
                If Not tblResult.Rows(0).Item("Password") & "" = objXe.Encrypt(colFields("AD")) Then
                    strResponse = "24              001" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & tblResult.Rows(0).Item("Code") & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("FullName")) & "|BLY|CQN"
                Else
                    strResponse = "24              001" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & tblResult.Rows(0).Item("Code") & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("FullName")) & "|BLY|CQY|BH" & CurrencyUnit & "|BV" & tblResult.Rows(0).Item("Debt")
                End If
            Else
                strResponse = "24              001" & PrintISOTime(Now, TimeZone) & "AO" & colFields("AO") & "|AA" & tblResult.Rows(0).Item("Code") & "|AE" & CutVietnameseAccent(tblResult.Rows(0).Item("FullName")) & "|BLY|CQY|BH" & CurrencyUnit & "|BV" & tblResult.Rows(0).Item("Debt")
            End If
        End If

        If strMesage <> "" Then
            strResponse = strResponse & "|AF" & strMesage
        End If

        'ServercheckSip2: Ten truong: TDMU, Noi (Location): VN
        '24              00020210326    054105AO|AA123|AENguyen Hoang Chinh|AY8AZEE6F  -----------------Good

        PatronStatusRequest = strResponse
    End Function

    Public Function PatronInformation(ByVal req As String) As String
        Dim colFields As New Collection
        Dim strLang, strTransDate, strResponse, strSummary, strTitle As String
        Dim strHomeAdd, strEmailAdd, strHomePhone, strItemDetails As String
        Dim intPatronID As Integer
        Dim intHoldCounts, intOverdueCounts, intChargeCounts, intFineCounts, intRecallCounts, intUHoldCounts As Integer
        Dim intHoldLimit, intStartPos, intEndPos As Integer
        Dim curDebt As Double = 0
        Dim dtmBasedDate As Date
        Dim arrTemp() As String = {"AO", "AA", "AD", "BQ", "BP", "AC"}

        strLang = Left(req, 3)
        strTransDate = Mid(req, 4, 18)
        strSummary = Mid(req, 22, 10)
        req = Right(req, Len(req) - 31)
        Call LoadFields(arrTemp, colFields, req)
        If IsNumeric(colFields("BP")) Then
            intStartPos = CInt(colFields("BP"))
        Else
            intStartPos = 1
        End If
        If IsNumeric(colFields("BQ")) Then
            intEndPos = CInt(colFields("BQ"))
        Else
            intEndPos = 10
        End If

        Dim tblResult As DataTable
        Dim strPatronCode As String = colFields("AA")

        Dim objDPatron As New clsDPatron
        objDPatron.DBServer = strDBServer
        objDPatron.ConnectionString = strConnectionstring
        Call objDPatron.Initialize()

        Dim objDLoanTransaction As New clsDLoanTransaction
        objDLoanTransaction.DBServer = strDBServer
        objDLoanTransaction.ConnectionString = strConnectionstring
        Call objDLoanTransaction.Initialize()

        objDPatron.FullName = ""
        objDPatron.PatronCode = strPatronCode
        tblResult = objDPatron.GetPatronInfor

        If tblResult Is Nothing Then
            strResponse = "64              001" & PrintISOTime(Now, TimeZone) & "000000000000000000000000AO" & colFields("AO") & "|AA" & colFields("AA") & "|AE|BLN"
            Return strResponse

        End If
        If tblResult.Rows.Count = 0 Then
                strResponse = "64              001" & PrintISOTime(Now, TimeZone) & "000000000000000000000000AO" & colFields("AO") & "|AA" & colFields("AA") & "|AE|BLN"
            Else
                'get fullname
                Dim strFullname As String = tblResult.Rows(0).Item("FullName") & ""
            'strFullname = strFullname.Substring(strFullname.IndexOf(">") + 1)
            'strFullname = CutVietnameseAccent(Left(strFullname, strFullname.Length - 4))
            strFullname = CutVietnameseAccent(strFullname)
            '  If Not IsDBNull(tblResult.Rows(0).Item("Password")) AndAlso Not tblResult.Rows(0).Item("Password") = colFields("AD") Then
            If 1 = 0 Then
                'Chinh: khong set kiem tra password
                '64              00020200319    112836000200020000000000000000AO|AA123|AEJohn Doe|BZ5|AY5AZEC5C
                '64              00020200319    115045000100020000000000000000AO|AA1|AEJohn Doe|BZ3|ATOverdue items count 1|ATOverdue items count 2|AY7AZDB38

                ' strResponse = "64              001" & PrintISOTime(Now, TimeZone) & "000000000000000000000000AO" & colFields("AO") & "|AA" & strPatronCode & "|AE" & CutVietnameseAccent(strFullname) & "|BLY|CQN"
                strResponse = "64              001" & PrintISOTime(Now, TimeZone) & "000300030000000000000000AO" & colFields("AO") & "|AA" & strPatronCode & "|AE" & CutVietnameseAccent(strFullname) & "|BZ5|BLY|CQN"
            Else
                Decimal.TryParse(tblResult.Rows(0).Item("Debt"), curDebt)
                intHoldLimit = tblResult.Rows(0).Item("InLibraryQuota")
                strEmailAdd = tblResult.Rows(0).Item("Email") & ""
                strHomePhone = tblResult.Rows(0).Item("Telephone") & ""
                strHomeAdd = CutVietnameseAccent(tblResult.Rows(0).Item("Address") & "")
                objDLoanTransaction.PatronCode = strPatronCode
                Dim tblTempCounts As DataTable
                Dim i As Integer
                intUHoldCounts = 0
                ' Get unavailable infor
                '    tblTempCounts = Me.GetData("SELECT CopyNumber, Content FROM CIR_HOLDING, Field200s WHERE CIR_HOLDING.ItemID = Field200s.ItemID AND FieldCode = '245' AND (PatronCode) = ('" & strPatronCode & "') ORDER BY CIR_HOLDING.ID")
                'tblTempCounts = Me.GetData("SELECT Ma_xep_gia as CopyNumber,dbo.fnGetSubField(f.Gia_tri,'$a','$') as Content FROM An_pham_cho_muon m, Field200s f WHERE m.Tai_lieu_ID = f.Tai_lieu_ID AND f.Truong_ID = '448' AND (So_the_ID) = (" & tblResult.Rows(0).Item("ID") & ")") ' ORDER BY m.ID

                'intUHoldCounts = 0 ' tblTempCounts.Rows.Count

                'Dim i As Integer
                'If InStr(strSummary, "Y") = 6 Then
                '    If intStartPos < tblTempCounts.Rows.Count And intStartPos > 0 Then
                '        ' For i = intStartPos To intEndPos
                '        For i = 0 To intUHoldCounts - 1
                '            If tblResult.Rows.Count = 0 Then
                '                Exit For
                '            End If
                '            strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblTempCounts.Rows(i).Item("Content")))
                '            If InStr(strTitle, "/") > 0 Then
                '                strTitle = Left(strTitle, InStr(strTitle, "/") - 1)
                '            End If
                '            '      strItemDetails = strItemDetails & "|CD" & strTitle & ". Item ID: " & tblTempCounts.Rows(i).Item("CopyNumber")
                '            strItemDetails = strItemDetails & "|CD" & tblTempCounts.Rows(i).Item("CopyNumber")
                '        Next
                '    End If
                'End If
                '    Dim strSqlCounts As String = "SELECT CopyNumber, Content, RecallDate, DueDate, CIR_LOAN_TYPE.Fee, CIR_LOAN_TYPE.OverdueFine FROM CIR_LOAN, CIR_LOAN_TYPE, Field200s,CIR_PATRON WHERE Field200s.ItemID = CIR_LOAN.ItemID AND CIR_LOAN.LoanTypeID = CIR_LOAN_TYPE.ID AND FieldCode = '245' AND CIR_LOAN.LoanMode IN (1, 2) AND PatronID = CIR_PATRON.ID AND (CIR_PATRON.CODE)='" & strPatronCode & "'"
                'Dim strSqlCounts As String = "SELECT Ma_xep_gia as CopyNumber,dbo.fnGetSubField(f.Gia_tri,'$a','$') as Content, RecallDate,ngay_tra as DueDate, l.Fee,0 as OverdueFine FROM An_pham_cho_muon m, LoanType l, Field200s f,ban_doc b WHERE  m.Tai_lieu_ID = f.Tai_lieu_ID AND m.LoanType = l.LoanTypeID AND Truong_ID = '448' AND  m.So_the_ID = b.ID AND (So_the)='" & strPatronCode & "'"
                Dim strSqlCounts As String = "SELECT Ma_xep_gia as CopyNumber,dbo.fnGetSubField(f.Gia_tri,'$a','$') as Content, RecallDate,ngay_tra as DueDate, 0 as Fee,0 as OverdueFine FROM An_pham_cho_muon m,  Field200s f,ban_doc b WHERE  m.Tai_lieu_ID = f.Tai_lieu_ID AND Truong_ID = '448' AND  m.So_the_ID = b.ID AND (So_the)='" & strPatronCode & "'"


                tblTempCounts = Me.GetData(strSqlCounts) ' & " ORDER BY m.ID")
                If Not tblTempCounts.Rows.Count = 0 Then
                    intHoldCounts = tblTempCounts.Rows.Count

                    ' If InStr(strSummary, "Y") = 1 Then
                    If InStr(strSummary, "Y") = 0 Then
                        '  If intStartPos < intHoldCounts And intStartPos > 0 Then
                        For i = 0 To intHoldCounts - 1
                            If tblTempCounts.Rows.Count = 0 Then
                                Exit For
                            End If
                            strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblTempCounts.Rows(i).Item("Content")))
                            If InStr(strTitle, "/") > 0 Then
                                strTitle = Left(strTitle, InStr(strTitle, "/") - 1)
                            End If
                            strItemDetails = strItemDetails & "|AS" & strTitle & ". Item ID: " & tblTempCounts.Rows(i).Item("CopyNumber")
                            '  strItemDetails = strItemDetails & "|AS" & tblTempCounts.Rows(i).Item("CopyNumber")
                        Next
                        ' End If
                    End If
                    If InStr(strSummary, "Y") = 5 Then
                        tblTempCounts = Me.GetData(strSqlCounts & " and RecallDate is not null ") 'ORDER BY m.ID
                        intRecallCounts = tblTempCounts.Rows.Count

                        'If intStartPos < tblTempCounts.Rows.Count And intStartPos > 0 Then
                        '    For i = intStartPos To intEndPos
                        For i = 0 To intRecallCounts - 1
                            If tblTempCounts.Rows.Count = 0 Then
                                Exit For
                            End If
                            'strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblTempCounts.Rows(i).Item("Content")))
                            'If InStr(strTitle, "/") > 0 Then
                            '    strTitle = Left(strTitle, InStr(strTitle, "/") - 1)
                            'End If
                            '   strItemDetails = strItemDetails & "|BU" & strTitle & ". Item ID: " & tblTempCounts.Rows(i).Item("CopyNumber")
                            strItemDetails = strItemDetails & "|BU" & tblTempCounts.Rows(i).Item("CopyNumber")
                        Next
                        '   End If
                    End If

                    tblTempCounts = Me.GetData(strSqlCounts & " and Fee > 0 ") 'ORDER BY m.ID
                    If Not tblTempCounts Is Nothing Then
                        intChargeCounts = tblTempCounts.Rows.Count
                        If InStr(strSummary, "Y") = 3 Then
                            'If intStartPos < tblTempCounts.Rows.Count And intStartPos > 0 Then
                            '    For i = intStartPos To intEndPos
                            For i = 0 To intChargeCounts - 1
                                If tblTempCounts.Rows.Count = 0 Then
                                    Exit For
                                End If
                                'strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblTempCounts.Rows(i).Item("Content")))
                                'If InStr(strTitle, "/") > 0 Then
                                '    strTitle = Left(strTitle, InStr(strTitle, "/") - 1)
                                'End If
                                strItemDetails = strItemDetails & "|AU" & strTitle & ". Item ID: " & tblTempCounts.Rows(i).Item("CopyNumber")
                            Next
                            '  End If
                        End If

                End If

                    If strDBServer = "ORACLE" Then
                        tblTempCounts = Me.GetData(strSqlCounts & " and to_char(DueDate,'mm/dd/yyyy') < to_char(sysdate,'mm/dd/yyyy') ORDER BY CIR_LOAN.ID")
                    Else
                        tblTempCounts = Me.GetData(strSqlCounts & " and ngay_tra > getdate()") ' ORDER BY m.ID
                    End If
                    If Not tblTempCounts Is Nothing Then
                        intOverdueCounts = tblTempCounts.Rows.Count
                        If InStr(strSummary, "Y") = 2 Then
                            'If intStartPos < tblTempCounts.Rows.Count And intStartPos > 0 Then
                            '    For i = intStartPos To intEndPos
                            For i = 0 To intOverdueCounts - 1
                                If tblTempCounts.Rows.Count = 0 Then
                                    Exit For
                                End If
                                strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblTempCounts.Rows(i).Item("Content")))
                                If InStr(strTitle, "/") > 0 Then
                                    strTitle = Left(strTitle, InStr(strTitle, "/") - 1)
                                End If
                                '   strItemDetails = strItemDetails & "|AT" & strTitle & ". Item ID: " & tblTempCounts.Rows(i).Item("CopyNumber")
                                strItemDetails = strItemDetails & "|AT" & tblTempCounts.Rows(i).Item("CopyNumber")
                            Next
                            '   End If
                        End If
                    End If
                    If strDBServer = "ORACLE" Then
                        tblTempCounts = Me.GetData(strSqlCounts & " and to_char(DueDate,'mm/dd/yyyy') < to_char(sysdate,'mm/dd/yyyy') AND OverdueFine > 0 ORDER BY CIR_LOAN.ID")
                    Else
                        tblTempCounts = Me.GetData(strSqlCounts & " and ngay_tra > getdate() AND OverdueFine > 0") ' ORDER BY m.ID
                    End If
                    If Not tblTempCounts Is Nothing Then
                        intFineCounts = tblTempCounts.Rows.Count
                        If InStr(strSummary, "Y") = 4 Then
                            'If intStartPos < tblTempCounts.Rows.Count And intStartPos > 0 Then
                            '    For i = intStartPos To intEndPos
                            For i = 0 To intFineCounts - 1
                                If tblTempCounts.Rows.Count = 0 Then
                                    Exit For
                                End If
                                strTitle = CutVietnameseAccent(TrimSubFieldCodes(tblTempCounts.Rows(i).Item("Content")))
                                If InStr(strTitle, "/") > 0 Then
                                    strTitle = Left(strTitle, InStr(strTitle, "/") - 1)
                                End If
                                strItemDetails = strItemDetails & "|AV" & strTitle & ". Item ID: " & tblTempCounts.Rows(i).Item("CopyNumber")
                            Next
                            '   End If
                        End If
                    End If
                Else
                        intHoldCounts = 0
                    intOverdueCounts = 0
                    intFineCounts = 0
                    intChargeCounts = 0
                End If
                'strResponse = "64              001" & PrintISOTime(Now, TimeZone) & CStr(intHoldCounts).PadLeft(4, "0") & CStr(intOverdueCounts).PadLeft(4, "0") & CStr(intChargeCounts).PadLeft(4, "0") & CStr(intFineCounts).PadLeft(4, "0") & CStr(intRecallCounts).PadLeft(4, "0") & CStr(intUHoldCounts).PadLeft(4, "0") & "AO" & colFields("AO") & "|AA" & strPatronCode & "|AE" & strFullname & "|BZ" & CStr(intHoldLimit).PadLeft(4, "0") & intHoldLimit & "|BLY|CQY|BH" & CurrencyUnit & "|BV" & curDebt & "|BD" & strHomeAdd & "|BE" & strEmailAdd & "|BF" & strHomePhone & strItemDetails
                strResponse = "64              000" & PrintISOTime(Now, "    ") & CStr(intHoldCounts).PadLeft(4, "0") & CStr(intOverdueCounts).PadLeft(4, "0") & CStr(intChargeCounts).PadLeft(4, "0") & CStr(intFineCounts).PadLeft(4, "0") & CStr(intRecallCounts).PadLeft(4, "0") & CStr(intUHoldCounts).PadLeft(4, "0") & "AO" & colFields("AO") & "|AA" & strPatronCode & "|AE" & strFullname & "|BZ" & intHoldLimit & strItemDetails
            End If
        End If
        '64              00020210326    164748000200010000000000000000AO|AA100220066|AELu Van Ngoc|AY3AZEB93 -----Good
        '64              00020210326    164556000000000000000000000000AO|AA100220066|AELu Van Ngoc|BZ10|
        '64              00020210327    044538000000000000000000000000AO|AA000230187|AELu Van Ngoc|AY0AZE700

        PatronInformation = strResponse
    End Function

    Public Function Login(ByVal req As String) As String
        Dim objXe As XCrypt.XCryptEngine
        objXe = New XCrypt.XCryptEngine(XCrypt.XCryptEngine.AlgorithmType.MD5)
        Dim colFields As New Collection
        Dim strUIDAlgorithm, strPWDAlgorithm As String
        Dim arrTemp() As String = {"CN", "CO", "CP"}
        Dim objDUser As New clsDUser
        Dim objDRole As New clsDRole
        objDUser.DBServer = strDBServer
        objDUser.ConnectionString = strConnectionstring
        Call objDUser.Initialize()
        objDRole.DBServer = strDBServer
        objDRole.ConnectionString = strConnectionstring
        Call objDRole.Initialize()

        strUIDAlgorithm = Left(req, 1)
        strPWDAlgorithm = Mid(req, 2, 1)
        req = Right(req, Len(req) - 2)
        Call LoadFields(arrTemp, colFields, req)

        Dim tblResult As DataTable
        Dim strPassword As String = objXe.Encrypt(Replace(colFields("CO"), "'", "''"), "pl")
        Dim strUser As String = Replace(colFields("CN"), "'", "''")
        objDUser.UserName = strUser
        objDUser.UserPass = strPassword
        tblResult = objDUser.GetUserLogin
        If Not tblResult.Rows.Count = 0 Then
            intUserID = tblResult.Rows(0).Item("ID")
            objDRole.UID = intUserID
            Dim tblRights As DataTable
            Dim i As Integer
            tblRights = objDRole.GetRights
            If Not tblRights Is Nothing Then
                If tblRights.Rows.Count > 0 Then
                    For i = 0 To tblRights.Rows.Count - 1
                        strUserRight = strUserRight & tblRights.Rows(i).Item(0) & ","
                    Next
                End If
            End If
            Login = "941"
        Else
            intUserID = 0
            Login = "940"
        End If
        'Server check: checksum, no checksum
        '941AY1AZFDFC
        '941AY2AZFDFB

        objDRole.Dispose()
        objDRole = Nothing
        objDUser.Dispose()
        objDUser = Nothing
    End Function



    Private Function TheDisplayOne(ByVal Title) As String
        Dim OutOfBracket, inBracketPart As String
        Dim dispTitle As String
        Dim i As Integer

        Title = Trim(Title)
        OutOfBracket = True
        dispTitle = ""
        inBracketPart = ""
        For i = 1 To Len(Title)
            If Mid(Title, i, 1) = "<" Then
                OutOfBracket = False
            End If
            If Mid(Title, i, 1) = ">" Then
                OutOfBracket = True
                If Not inBracketPart = "" Then
                    If InStr(inBracketPart, "=") > 0 Then
                        dispTitle = dispTitle & Left(inBracketPart, InStr(inBracketPart, "=") - 1)
                    Else
                        dispTitle = dispTitle & inBracketPart
                    End If
                    inBracketPart = ""
                End If
            End If
            If OutOfBracket Then
                If Mid(Title, i, 1) <> ">" Then
                    dispTitle = dispTitle & Mid(Title, i, 1)
                End If
            Else
                If Mid(Title, i, 1) <> "<" Then
                    inBracketPart = inBracketPart & Mid(Title, i, 1)
                End If
            End If
        Next
        If Not inBracketPart = "" Then
            If InStr(inBracketPart, "|") > 0 Then
                dispTitle = dispTitle & Left(inBracketPart, InStr(inBracketPart, "=") - 1)
            Else
                dispTitle = dispTitle & inBracketPart
            End If
            inBracketPart = ""
        End If
        TheDisplayOne = dispTitle
    End Function

    Private Function TrimSubFieldCodes(ByVal s As String) As String
        Dim j As Long
        Dim o As String
        o = ""
        Do While Len(s) > 0
            j = InStr(s, "$")
            If j > 0 Then
                o = o & Left(s, j - 1) & " "
                s = Right(s, Len(s) - j - 1)
            Else
                o = o & s
                s = ""
            End If
        Loop
        TrimSubFieldCodes = o
    End Function


    Private Function RenewItem(ByVal dblLoanID As Integer, ByVal intAddTime As Integer, ByVal intTU As Integer, ByVal dtmDueDate As Date) As Date
        Dim tblRenewItem As DataTable
        Dim objDRenew As New clsDRenew
        objDRenew.DBServer = strDBServer
        objDRenew.ConnectionString = strConnectionstring
        Call objDRenew.Initialize()
        objDRenew.RenewItems(dblLoanID, intAddTime, intTU, ConvertDateBack(dtmDueDate))
        tblRenewItem = Me.GetData("SELECT DueDate FROM CIR_LOAN where id=" & dblLoanID)
        objDRenew.Dispose()
        objDRenew = Nothing
        Return tblRenewItem.Rows(0).Item("DueDate")
    End Function

    ' Dispose method
    ' Purpose: release all objects
    ' Input: boolean value IsDisposing
    Public Overridable Overloads Sub Dispose(ByVal IsDisposing As Boolean)
        If IsDisposing Then
            ' Release unmanaged resources.
            If Not dsData Is Nothing Then
                dsData.Dispose()
                dsData = Nothing
            End If
            If Not oraConnection Is Nothing Then
                oraConnection.Dispose()
                oraConnection = Nothing
            End If
            If Not oraCommand Is Nothing Then
                oraCommand.Dispose()
                oraCommand = Nothing
            End If
            If Not sqlConnection Is Nothing Then
                sqlConnection.Dispose()
                sqlConnection = Nothing
            End If
            If Not sqlCommand Is Nothing Then
                sqlCommand.Dispose()
                sqlCommand = Nothing
            End If
            If Not sqlDataAdapter Is Nothing Then
                sqlDataAdapter.Dispose()
                sqlDataAdapter = Nothing
            End If
            If Not oraDataAdapter Is Nothing Then
                oraDataAdapter.Dispose()
                oraDataAdapter = Nothing
            End If
        End If
    End Sub

    ' Finalize method
    ' Purpose: call Dispose(False).
    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

End Class
