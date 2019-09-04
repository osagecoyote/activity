Imports Microsoft.Win32
Imports System.Data
Imports System.Data.SqlClient

Module Module1

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "  Class Members "

    'Property vars;
    Private p_applicationpath As String = ""
    Private p_savefilepath As String = ""
    'Connection vars;
    Private cn As SqlConnection

#End Region

#Region "  Enumerations "

    Public Enum CalendarMonths As Integer
        Null = 0
        January = 1
        February = 2
        March = 3
        April = 4
        May = 5
        June = 6
        July = 7
        August = 8
        September = 9
        October = 10
        November = 11
        December = 12
    End Enum

#End Region

#Region "  Functions "

    Function ConvertCardinalMonthToString(ByVal emonth As Int32) As String
        Try
            Dim enumeratemonths() As String = [Enum].GetNames(GetType(CalendarMonths))
            Return enumeratemonths(emonth)
        Catch ex As Exception
            Return "Null"
        End Try
    End Function

    Function ConvertMonthStringToCardinal(ByVal emonthstring As String) As Int32
        Try
            Dim months As CalendarMonths
            Return CInt([Enum].Parse(months.GetType, emonthstring, True))
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function FindFirstOfMonth(ByVal edate As Date) As Date
        'this routine returns the first date's day of the given month's date
        Dim starting As Date
        Dim days As Int32
        Dim month As Int32 = edate.Month
        Dim year As Int32 = edate.Year
        Try
            'calc the beginning day of the current month
            starting = CDate(month & "/01/" & year)
            Return starting
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function FindLastOfMonth(ByVal edate As Date) As Date
        'this routine returns the last date's day of the given month's date
        Dim starting, ending As Date
        Dim days As Int32
        Dim month As Int32 = edate.Month
        Dim year As Int32 = edate.Year
        Try
            'calc the beginning day of the current month
            starting = CDate(month & "/01/" & year)
            'calc the last day of the month by extracting the number of days in the month
            days = starting.DaysInMonth(year, month)
            ending = CDate(month & "/" & days & "/" & year & " 11:59:59 PM")
            Return ending
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function FormatExpenditureCode(ByVal code As String) As String
        'returns a formatted expenditure code
        If code.Trim.Length <> 26 Then Return ""
        Dim tempcode As String
        tempcode = code.Substring(0, 1)
        tempcode += "-" & code.Substring(1, 2)
        tempcode += "-" & code.Substring(3, 3)
        tempcode += "-" & code.Substring(6, 4)
        tempcode += "-" & code.Substring(10, 3)
        tempcode += "-" & code.Substring(13, 3)
        tempcode += "-" & code.Substring(16, 4)
        tempcode += "-" & code.Substring(20, 3)
        tempcode += "-" & code.Substring(23, 3)
        Return tempcode
    End Function

    Function FormatRevenueCode(ByVal code As String) As String
        'returns a formatted revenue code
        If code.Trim.Length <> 16 Then Return ""
        Dim tempcode As String
        tempcode = code.Substring(0, 1)
        tempcode += "-" & code.Substring(1, 2)
        tempcode += "-" & code.Substring(3, 3)
        tempcode += "-" & code.Substring(6, 4)
        tempcode += "-" & code.Substring(10, 3)
        tempcode += "-" & code.Substring(13, 3)
        Return tempcode
    End Function

    Function Format3CharNumber(ByVal number As String) As String
        Dim x As Int32
        Dim str As String
        Try
            x = CInt(number)
            If x > 999 Then x = 1
            str = x.ToString.Format("{0:D3}", x)
            Return str
        Catch ex As Exception
            'if any error occurs, return the first available number
            Return "001"
        End Try
    End Function

    Function Format4CharNumber(ByVal number As String) As String
        Dim x As Int32
        Dim str As String
        Try
            x = CInt(number)
            If x > 9999 Then x = 1
            str = x.ToString.Format("{0:D4}", x)
            Return str
        Catch ex As Exception
            'if any error occurs, return the first available number
            Return "0001"
        End Try
    End Function

    Function Format5CharNumber(ByVal number As String) As String
        Dim x As Int32
        Dim str As String
        Try
            x = CInt(number)
            str = x.ToString.Format("{0:D5}", x)
            Return str
        Catch ex As Exception
            'if any error occurs, return the first available number
            Return "00001"
        End Try
    End Function

    Function Format8CharNumber(ByVal number As String) As String
        Dim x As Int32
        Dim str As String
        Try
            x = CInt(number)
            str = x.ToString.Format("{0:D8}", x)
            Return str
        Catch ex As Exception
            'if any error occurs, return the first available number
            Return "00000001"
        End Try
    End Function

#End Region

#Region "  Permissions "

    Friend Function GetPermissions(ByVal econnectionstring As String, ByVal euserkey As Int32, ByVal eassembly As Int32, ByVal eadministrator As Boolean) As Int32
        Dim SSQL As String
        Dim cmd As SqlCommand

        'if admin, then skip;
        If eadministrator Then Return 2

        Try
            'get the permissions for this module;
            SSQL = "SELECT perm_rw FROM upermissions WHERE user_autoinc_key = @p1 and assm_id = @p2"
            cn = New SqlConnection(econnectionstring)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", euserkey)
            cmd.Parameters.Add("@p2", eassembly)
            cn.Open()
            Return CInt(cmd.ExecuteScalar)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            cmd.Dispose()
        End Try
    End Function

#End Region

#Region "  Properties "

    Friend ReadOnly Property ApplicationPath() As String
        Get
            Return p_applicationpath
        End Get
    End Property

    Friend Property SaveFilePath() As String
        Get
            Return p_savefilepath
        End Get
        Set(ByVal Value As String)
            If Not Value.EndsWith("\") Then Value += "\"
            p_savefilepath = Value
        End Set
    End Property

#End Region

#Region "  Registry Win32 "

    'registry values;
    Private _rgnode As String = "SOFTWARE"
    Private _rgcompany As String = "ADPC"
    Private _rgproduct As String = "Activity Fund.Net"
    Private _rgmain As String = "AF_Main"
    Private _rgreporting As String = "AF_Reporting"

    Friend Sub GetRegistry()
        Dim rgSoftware As RegistryKey
        Dim rgMain As RegistryKey
        Dim rgModule As RegistryKey
        Dim _rg32node As String = "ADPC\Activity Fund.Net\AF_Reporting"
        Dim _rg64node As String = "Wow6432Node\ADPC\Activity Fund.Net\AF_Reporting"
        Dim _rg32mainnode As String = "ADPC\Activity Fund.Net\AF_Main"
        Dim _rg64mainnode As String = "Wow6432Node\ADPC\Activity Fund.Net\AF_Main"

        Try
            'get key to HKEY_LOCAL_MACHINE/SOFTWARE/;
            rgSoftware = Registry.LocalMachine.OpenSubKey("SOFTWARE", False)
            If rgSoftware Is Nothing Then Exit Sub

            'Test for 64 bit for AF_Reporting.
            rgModule = rgSoftware.OpenSubKey(_rg64node, False)
            If rgModule Is Nothing Then
                'Test for 32 bit and exit if not found.
                rgModule = rgSoftware.OpenSubKey(_rg32node, False)
                If rgModule Is Nothing Then Exit Sub
            End If

            'Test for 64 bit for AF_Main.
            rgMain = rgSoftware.OpenSubKey(_rg64mainnode, False)
            If rgMain Is Nothing Then
                'Test for 32 bit and exit if not found.
                rgMain = rgSoftware.OpenSubKey(_rg32mainnode, False)
                If rgMain Is Nothing Then Exit Sub
            End If

            'Get the startup path from AF_Main.
            p_applicationpath = CType(rgMain.GetValue("ApplicationPath", ""), String)
            'Get the positive pay path from AF_Reporting.
            Module1.SaveFilePath = CType(rgModule.GetValue("PositivePayPath", Module1.ApplicationPath), String)

            'Close resource.
            If Not rgSoftware Is Nothing Then rgSoftware.Close()
            If Not rgMain Is Nothing Then rgMain.Close()
            If Not rgModule Is Nothing Then rgModule.Close()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Friend Sub SetRegistry()
        Dim rgSoftware As RegistryKey
        Dim rgModule As RegistryKey
        Dim _rg32node As String = "ADPC\Activity Fund.Net\AF_Reporting"
        Dim _rg64node As String = "Wow6432Node\ADPC\Activity Fund.Net\AF_Reporting"

        Try
            'get key to HKEY_LOCAL_MACHINE/SOFTWARE/;
            rgSoftware = Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            If rgSoftware Is Nothing Then Exit Sub

            'Test for 64 bit for AF_Reporting.
            rgModule = rgSoftware.OpenSubKey(_rg64node, True)
            If rgModule Is Nothing Then
                'Test for 32 bit and exit if not found.
                rgModule = rgSoftware.OpenSubKey(_rg32node, True)
                If rgModule Is Nothing Then Exit Sub
            End If

            'Set values in the registry;
            rgModule.SetValue("PositivePayPath", Module1.SaveFilePath)

            'Close resource.
            If Not rgSoftware Is Nothing Then rgSoftware.Close()
            If Not rgModule Is Nothing Then rgModule.Close()
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

End Module
