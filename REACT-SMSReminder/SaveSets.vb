Imports System.Runtime.InteropServices
Imports System.Text

Module SaveSets
	'Get the configuration file
    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, _
                ByVal lpKeyName As String, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As StringBuilder, _
                ByVal nSize As Integer, _
                ByVal lpFileName As String) As Integer

	'Load the specific value from config file
    Public Function Load(ByVal Section As String, ByVal Key As String) As String
        Dim res As Integer

        Dim strFileName As String = ""
        Dim sb As StringBuilder
        sb = New StringBuilder(500)
        strFileName = Environment.CurrentDirectory & "\Configuration.ini"

        res = GetPrivateProfileString(Section, Key, "", sb, sb.Capacity, strFileName)

        Load = Trim(sb.ToString)
    End Function
End Module
