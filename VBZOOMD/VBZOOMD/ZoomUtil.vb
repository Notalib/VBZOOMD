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


    Public Shared Function CopyString(ptr As Integer, s As Integer) As String
        Dim l As Integer
        Dim ret As String

        If ptr = 0 Then
            ret = ""
        Else
            If s = -1 Then
                l = lstrlen(ptr)
                ret = New String(vbNullChar, l)

                l = lstrcpy(ret, ptr)

                If l = 0 Then
                    Throw New ZoomException(Nothing, ERR_WIN_API, "Unable to copy string")
                End If

            Else
                ret = New String(vbNullChar, s)

                StrCpy(ret, ptr, s)
            End If

        End If


        Return ret
    End Function


    Public Shared Function CopyByteArray(ptr As Integer, s As Integer) As Byte()
        Dim l As Integer
        Dim ret() As Byte = Nothing

        If ptr <> 0 Then
            If s = -1 Then
                l = lstrlen(ptr)
                ReDim ret(l)

                l = lstrcpybyte(ret(0), ptr)

                If l = 0 Then
                    Throw New ZoomException(Nothing, ERR_WIN_API, "Unable to copy bytes")
                End If

            Else
                ReDim ret(s)

                BytCpy(ret(0), ptr, s)
            End If

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
