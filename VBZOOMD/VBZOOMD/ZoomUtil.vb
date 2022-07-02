Public Class ZoomUtil
    Private Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" (ByVal ptr As Integer) As Integer

    Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal dest As String, ByVal ptr As Integer) As Integer
    Private Declare Function lstrcpybyte Lib "KERNEL32" Alias "lstrcpyA" (ByRef dest As Byte, ByVal ptr As Integer) As Integer

    Private Declare Sub StrCpy Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal dest As String, ByVal ptr As Integer, ByVal size As Integer)
    Private Declare Sub BytCpy Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef dest As Byte, ByVal ptr As Integer, ByVal size As Integer)

    Public Const ERR_INVALID_QUERY_TYPE As Integer = 101
    Public Const ERR_ZOOM_ERROR As Integer = 102
    Public Const ERR_INVALID_QUERY As Integer = 103
    Public Const ERR_NO_RECORD As Integer = 104
    Public Const ERR_INVALID_OPTION_NAME As Integer = 105
    Public Const ERR_INVALID_SCAN_FIELD As Integer = 106
    Public Const ERR_INVALID_CONNECTION As Integer = 107
    Public Const ERR_NOT_SUPPORTED As Integer = 108
    Public Const ERR_DEPRECATED_FUNCTION As Integer = 109

    Public Const ERR_UNSUPPORTED_RAW_TYPE As Integer = 111
    Public Const ERR_INVALID_OPTION_VALUE As Integer = 112

    Public Const ERR_WIN_API As Integer = 180

    'custom options
    Public Const X_IGNOREMARCCHARACTERENCODING As String = "X-VBZOOM-ignoreMARCCharacterEncoding"
    Public Const X_ASSUMEDMARCCHARACTERENCODING As String = "X-VBZOOM-assumedMARCCharacterEncoding"

    'Supported MARC Character Encodings
    Public Const MARC_8 As String = "marc-8"
    Public Const UTF_8 As String = "utf-8"
    Public Const ISO_8859_1 As String = "iso-8859-1"

    Public Const Z_PQF As Integer = 0
    Public Const Z_CQL As Integer = 1

    ''' <summary>
    ''' Given a C pointer to a string, copy the string to a managed String object
    ''' </summary>
    ''' <param name="ptr">C-style memory pointer</param>
    ''' <param name="s">Length or -1 if it is null-terminated</param>
    ''' <returns></returns>
    Public Shared Function CopyString(ptr As Integer, s As Integer) As String
        If ptr = 0 Then Return ""

        Dim bytes() As Byte = ZoomUtil.CopyByteArray(ptr, s)

        'trim the null string terminator
        ReDim Preserve bytes(bytes.Count - 2)

        Return Text.Encoding.GetEncoding(UTF_8).GetString(bytes)
    End Function

    ''' <summary>
    ''' Given a C pointer, copy the bytes to a managed Byte Array
    ''' </summary>
    ''' <param name="ptr">C-style memory pointer</param>
    ''' <param name="s">Length or -1 if it is null-terminated</param>
    ''' <returns></returns>
    Public Shared Function CopyByteArray(ptr As Integer, s As Integer) As Byte()
        Dim ret() As Byte = Nothing

        If ptr <> 0 Then
            Dim iptr As New IntPtr(ptr)
            If s = -1 Then
                'count the bytes until the null terminator
                s = 0
                Do Until 0 = Runtime.InteropServices.Marshal.ReadByte(iptr, s)
                    s += 1
                Loop
            End If
            ReDim ret(s)
            Runtime.InteropServices.Marshal.Copy(iptr, ret, 0, s)
        End If

        Return ret
    End Function



    Public Shared Sub ValidateOption(who As Object, name As String, Optional value As String = "")
        'following are the various options which can be set for the YAZ ZOOM API
        Select Case name
            'connection options
            Case "implementationName"
            Case "implementationId"
            Case "implementationVersion"
            Case "user"
            Case "group"
            Case "pass"
            Case "password"
            Case "host"
            Case "proxy"
            Case "async"
            Case "maximumRecordSize"
            Case "preferredMessageSize"
            Case "lang"
            Case "charset"
            Case "targetImplementationId"
            Case "targetImplementationName"
            Case "targetImplementationVersion"
            Case "serverImplementationId"
            Case "serverImplementationName"
            Case "serverImplementationVersion"
            Case "databaseName"
            Case "piggyback"
            Case "smallSetUpperBound"
            Case "largeSetLowerBound"
            Case "mediumSetPresentNumber"
            Case "smallSetElementSetName"
            Case "mediumSetElementSetName"

                'resultset options
            Case "start"
            Case "count"
            Case "presentChunk"
            Case "step"
            Case "elementSetName"
            Case "preferredRecordSyntax"
            Case "schema"
            Case "setname"

                'scanset options
            Case "number"
            Case "position"
            Case "stepSize"
            Case "scanStatus"

                'undocumented options
            Case "timeout"

                'custom resultset options
            Case X_IGNOREMARCCHARACTERENCODING
                If TypeName(who) <> "ZoomResultSet" Then Throw New ZoomException(who, ERR_INVALID_OPTION_NAME, name)

            Case X_ASSUMEDMARCCHARACTERENCODING
                If TypeName(who) <> "ZoomResultSet" Then
                    Throw New ZoomException(who, ERR_INVALID_OPTION_NAME, name)
                End If
                If Not String.IsNullOrWhiteSpace(value) Then
                    If value <> UTF_8 And value <> MARC_8 And value <> ISO_8859_1 Then
                        Throw New ZoomException(who, ERR_INVALID_OPTION_VALUE, value, name)
                    End If
                End If
            Case Else
                Throw New ZoomException(who, ERR_INVALID_OPTION_NAME, name)
        End Select
    End Sub

End Class
