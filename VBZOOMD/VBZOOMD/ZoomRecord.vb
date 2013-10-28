Imports System.Xml

Public Class ZoomRecord
    Implements IDisposable

    Public Enum ZoomRawDataType
        RawDataDefault = 0
        RawDataString = vbString
        RawDataByteArray = vbByte + vbArray
        RawDataPointer = vbInteger
        RawDataLength = -1
    End Enum

    Private bIgnoreMARCCharacterEncoding As Boolean
    Private sAssumedMARCCharacterEncoding As String

    Private sYAZCharsetFrom As String
    Private sYAZCharsetTo As String

    Private hZOOM_record As Integer


    Private Declare Function ZOOM_record_clone Lib "YAZ5.dll" Alias "_ZOOM_record_clone@4" (ByVal rec As Integer) As Integer
    Private Declare Sub ZOOM_record_destroy Lib "YAZ5.dll" Alias "_ZOOM_record_destroy@4" (ByVal rec As Integer)
    Private Declare Function ZOOM_record_get Lib "YAZ5.dll" Alias "_ZOOM_record_get@12" (ByVal rec As Integer, ByVal typ As String, ByRef size As Integer) As Integer



    Friend Sub frndSetMARCCharcterEncoding(ignore As Boolean, assumed As String)
        bIgnoreMARCCharacterEncoding = ignore
        sAssumedMARCCharacterEncoding = assumed
    End Sub

    Public Function GetRecordSyntax() As String
        Dim l As Integer
        Dim s As Integer

        l = ZOOM_record_get(hZOOM_record, "syntax" & priGetYAZCharsetFromTo(), s)

        Return ZoomUtil.CopyString(l, -1)
    End Function


    Private Function priGetYAZCharsetFromTo() As String
        Dim ret As String
        If Len(sYAZCharsetFrom) > 0 Then
            ret = ";charset=" & sYAZCharsetFrom
            If Len(sYAZCharsetTo) > 0 Then
                ret = ret & "," & sYAZCharsetTo
            End If
        Else
            ret = ""
        End If

        priGetYAZCharsetFromTo = ret
    End Function

    Public Function XMLRecord() As String
        Dim l As Integer
        Dim s As Integer
        Dim ret As String = ""

        l = ZOOM_record_get(hZOOM_record, "xml" & priGetYAZCharsetFromTo(), s)

        If l <> 0 Then
            ret = ZoomUtil.CopyString(l, s)
        End If

        Return ret
    End Function

    Public Function RenderRecord() As String
        Dim l As Integer
        Dim s As Integer
        Dim ret As String = ""

        l = ZOOM_record_get(hZOOM_record, "render" & priGetYAZCharsetFromTo(), s)

        If l <> 0 Then
            ret = ZoomUtil.CopyString(l, s)
        End If

        Return ret
    End Function


    Public ReadOnly Property Database() As String
        Get
            Dim l As Integer
            Dim s As Integer

            l = ZOOM_record_get(hZOOM_record, "database" & priGetYAZCharsetFromTo(), s)

            Return ZoomUtil.CopyString(l, -1)

        End Get
    End Property

    Public Function RawData(Optional ByRef typ As ZoomRawDataType = ZoomRawDataType.RawDataDefault) As Object
        Dim l As Integer
        Dim s As Integer
        Dim ret As Object

        l = ZOOM_record_get(hZOOM_record, "raw" & priGetYAZCharsetFromTo(), s)

        If typ = ZoomRawDataType.RawDataLength Then
            ret = s

        ElseIf s <> -1 Then
            If typ = ZoomRawDataType.RawDataDefault Or typ = ZoomRawDataType.RawDataString Then
                typ = ZoomRawDataType.RawDataString
                ret = ZoomUtil.CopyString(l, s)

            ElseIf typ = ZoomRawDataType.RawDataByteArray Then
                typ = ZoomRawDataType.RawDataByteArray
                ret = ZoomUtil.CopyByteArray(l, s)

            ElseIf typ = ZoomRawDataType.RawDataPointer Then
                typ = ZoomRawDataType.RawDataPointer
                ret = l

            Else
                Throw New ZoomException(Me, ZoomUtil.ERR_UNSUPPORTED_RAW_TYPE, typ)

            End If
        Else
            typ = ZoomRawDataType.RawDataPointer
            ret = l

        End If

        Return ret
    End Function

    Friend Sub frndSetZoomRecordHandle(h As Integer)
        Dim syntax As String

        If hZOOM_record <> 0 Then
            ZOOM_record_destroy(hZOOM_record)
            hZOOM_record = 0
        End If

        hZOOM_record = ZOOM_record_clone(h)

        syntax = Me.GetRecordSyntax

    End Sub

    Public Function XMLData() As XmlDocument
        Dim retXML As New XmlDocument
        Dim rs As String

        rs = GetRecordSyntax()

        Select Case rs

            Case "USmarc"
                retXML.LoadXml(Me.XMLRecord)

            Case "text-XML", "XML", "application-XML"
                retXML.LoadXml(Me.XMLRecord)

            Case "OPAC"
                retXML.LoadXml(Me.XMLRecord)

            Case Else
                Throw New ZoomException(Me, ZoomUtil.ERR_NOT_SUPPORTED, Me.GetRecordSyntax)
                retXML = Nothing

        End Select

        Return retXML
    End Function

    Public Sub New()
        hZOOM_record = 0
        bIgnoreMARCCharacterEncoding = False
        sAssumedMARCCharacterEncoding = ZoomUtil.MARC_8
        sYAZCharsetFrom = ""
        sYAZCharsetTo = ""

    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' No managed objects to dispose
            End If

            If hZOOM_record <> 0 Then
                ZOOM_record_destroy(hZOOM_record)
            End If
            hZOOM_record = 0
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
