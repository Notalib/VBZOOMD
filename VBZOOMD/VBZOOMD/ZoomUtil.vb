Public Module ZoomUtil
    Private Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" (ByVal ptr As Integer) As Integer

    Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal dest As String, ByVal ptr As Integer) As Integer
    Private Declare Function lstrcpybyte Lib "KERNEL32" Alias "lstrcpyA" (ByRef dest As Byte, ByVal ptr As Integer) As Integer

    Private Declare Function lstrcpyn Lib "KERNEL32" Alias "lstrcpynA" (ByVal dest As String, ByVal ptr As Integer, ByVal size As Integer) As Integer

    Public Declare Sub RtlMoveMemory Lib "KERNEL32" (ByRef dest As Object, ByRef ptr As Object, ByVal size As Integer)
    Private Declare Sub StrCpy Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal dest As String, ByVal ptr As Integer, ByVal size As Integer)
    Private Declare Sub BytCpy Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef dest As Byte, ByVal ptr As Integer, ByVal size As Integer)

    Public Const MY_ERR_OFFSET = 100

    Public Const ERR_INVALID_QUERY_TYPE = 101
    Public Const ERR_ZOOM_ERROR = 102
    Public Const ERR_INVALID_QUERY = 103
    Public Const ERR_NO_RECORD = 104
    Public Const ERR_INVALID_OPTION_NAME = 105
    Public Const ERR_INVALID_SCAN_FIELD = 106
    Public Const ERR_INVALID_FIELD_SPEC = 107
    Public Const ERR_NOT_SUPPORTED = 108
    Public Const ERR_INTERNAL_ERR = 109
    Public Const ERR_UNKNOWN_ENCODING = 110
    Public Const ERR_UNSUPPORTED_RAW_TYPE = 111
    Public Const ERR_INVALID_OPTION_VALUE = 112

    Public Const ERR_MARC_MISSING_RT = 131
    Public Const ERR_MARC_MISSING_DIR_FT = 132
    Public Const ERR_MARC_INVALID_ENTRY_MAP = 133
    Public Const ERR_MARC_MISSING_VAR_FT = 134
    Public Const ERR_MARC_MISSING_CONTROL_NUMBER = 135
    Public Const ERR_XML = 136
    Public Const ERR_MARC_UNINITIALIZED = 137
    Public Const ERR_MARC_XML = 138
    Public Const ERR_MARC_IMPL_DEFINED = 139
    Public Const ERR_MARC_IND = 140
    Public Const ERR_MARC_LEADER = 141

    Public Const ERR_INVALID_TAG = 160
    Public Const ERR_INVALID_BER_CLASS = 161
    Public Const ERR_UNTERMINATED_INDEFINITE = 162

    Public Const ERR_WIN_API = 180
    Public Const ERR_MSXML_4_0_NOT_INSTALLED = 181

    Public Const WRN_INVALID_CHAR = 700
    Public Const WRN_LENGTH_IMPL_DEFINED = 701
    Public Const WRN_MARC_8_TO_UNICODE = 702
    Public Const WRN_MARC_INVALID_IND = 703

    'custom options
    Public Const X_IGNOREMARCCHARACTERENCODING = "X-VBZOOM-ignoreMARCCharacterEncoding"
    Public Const X_ASSUMEDMARCCHARACTERENCODING = "X-VBZOOM-assumedMARCCharacterEncoding"

    'MARC LEADER/09 values
    Public Const LDR_09_MARC_8 = " "
    Public Const LDR_09_UTF_8 = "a"

    'Supported MARC Character Encodings
    Public Const MARC_8 = "marc-8"
    Public Const UTF_8 = "utf-8"
    Public Const ISO_8859_1 = "iso-8859-1"



    Public Structure Z_External
        Dim direct_reference As Integer 'pointer to Odr_oid (array of integers terminated by -1)
        Dim indirect_reference As Integer 'pointer to integer
        Dim descriptor As Integer 'pointer to byte
        Dim which As Integer
        'Generic types
        'Z_External_single 0
        'Z_External_octet 1
        'Z_External_arbitrary 2
        'Specific types
        'Z_External_sutrs 3
        'Z_External_explainRecord 4
        'Z_External_resourceReport1 5
        'Z_External_resourceReport2 6
        'Z_External_promptObject1 7
        'Z_External_grs1 8
        'Z_External_extendedService 9
        'Z_External_itemOrder 10
        'Z_External_diag1 11
        'Z_External_espec1 12
        'Z_External_summary 13
        'Z_External_OPAC 14
        'Z_External_searchResult1 15
        'Z_External_update 16
        'Z_External_dateTime 17
        'Z_External_universeReport 18
        'Z_External_ESAdmin 19
        'Z_External_update0 20
        'Z_External_userInfo1 21
        'Z_External_charSetandLanguageNegotiation 22
        'Z_External_acfPrompt1 23
        'Z_External_acfDes1 24
        'Z_External_acfKrb1 25
        Dim u As Integer 'pointer to the actual record (C union type)
    End Structure

    Public Structure odr_oct
        Dim buf As Integer 'pointer to char (string)
        Dim len_ As Integer
        Dim size As Integer
    End Structure

    Public Function CopyString(ptr As Integer, s As Integer) As String
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


        CopyString = ret
    End Function


    Public Function CopyByteArray(ptr As Integer, s As Integer) As Byte()
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



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Purpose: Return a string which is a copy of the bytes contained
    '         at the buffer pointer to by ptr, using the semantics of
    '         the Mid function
    '
    'Inputs:  ptr   memory pointer to the start of the buffer
    '         s     1-based offset into buffer
    '         l     the number of bytes to copy
    '
    'Returns: string copy of buffer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function MidPtr(ptr As Integer, s As Integer, l As Integer) As String
        Dim ret As String

        ret = CopyString(ptr + s - 1, l)

        MidPtr = ret
    End Function


    Public Sub ValidateOption(who As Object, name As String, Optional value As String = "")
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
End Module
