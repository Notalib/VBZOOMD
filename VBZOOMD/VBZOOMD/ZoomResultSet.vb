Public Class ZoomResultSet
    Implements IDisposable

    Private hZOOM_resultset As Integer
    Private zConnection As ZoomConnection

    Private bIgnoreMARCCharacterEncoding As Boolean
    Private sAssumedMARCCharacterEncoding As String

    Private Declare Sub ZOOM_resultset_destroy Lib "YAZ5.dll" Alias "_ZOOM_resultset_destroy@4" (ByVal r As Integer)

    Private Declare Sub ZOOM_resultset_option_set Lib "YAZ5.dll" Alias "_ZOOM_resultset_option_set@12" (ByVal r As Integer, ByVal key As String, ByVal val As String)

    Private Declare Function ZOOM_resultset_option_get Lib "YAZ5.dll" Alias "_ZOOM_resultset_option_get@8" (ByVal r As Integer, ByVal key As String) As Integer

    Private Declare Function ZOOM_resultset_size Lib "YAZ5.dll" Alias "_ZOOM_resultset_size@4" (ByVal r As Integer) As Integer

    Private Declare Sub ZOOM_resultset_records Lib "YAZ5.dll" Alias "_ZOOM_resultset_records@16" (ByVal r As Integer, ByRef recs As Integer, ByVal start As Integer, ByVal count As Integer)

    Private Declare Function ZOOM_resultset_record Lib "YAZ5.dll" Alias "_ZOOM_resultset_record@8" (ByVal r As Integer, ByVal pos As Integer) As Integer


    Public ReadOnly Property AdditionalInfo() As String
        Get
            Return zConnection.AdditionalInfo
        End Get
    End Property

    Public ReadOnly Property ErrorCode() As Integer
        Get
            Return zConnection.ErrorCode
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return zConnection.ErrorMessage
        End Get
    End Property

    Friend Sub frndSetResultSetHandle(h As Integer)

        If hZOOM_resultset <> 0 Then
            ZOOM_resultset_destroy(hZOOM_resultset)
            zConnection.frndCheckForError()
            hZOOM_resultset = 0
        End If

        hZOOM_resultset = h

    End Sub


    Public Function GetRecord(which As Integer) As ZoomRecord
        Dim hzr As Integer
        Dim zr As New ZoomRecord

        hzr = ZOOM_resultset_record(hZOOM_resultset, which)
        zConnection.frndCheckForError()

        If hzr = 0 Then
            Throw New ZoomException(Me, ZoomUtil.ERR_NO_RECORD, which)
        Else
            zr.frndSetMARCCharcterEncoding(bIgnoreMARCCharacterEncoding, sAssumedMARCCharacterEncoding)
            zr.frndSetZoomRecordHandle(hzr)
        End If

        GetRecord = zr
    End Function

    Public Function GetSize() As Integer
        GetSize = ZOOM_resultset_size(hZOOM_resultset)
        zConnection.frndCheckForError()
    End Function


    Friend Sub frndSetConnection(c As ZoomConnection)
        zConnection = c
    End Sub

    Public Property ZoomOption(name As String) As String
        Get
            Dim ret As Integer
            Dim v As String

            ZoomUtil.ValidateOption(Me, name)

            Select Case name

                Case ZoomUtil.X_IGNOREMARCCHARACTERENCODING
                    v = CStr(bIgnoreMARCCharacterEncoding * -1)

                Case ZoomUtil.X_ASSUMEDMARCCHARACTERENCODING
                    v = sAssumedMARCCharacterEncoding

                Case Else

                    ret = ZOOM_resultset_option_get(hZOOM_resultset, name)
                    zConnection.frndCheckForError()

                    v = ZoomUtil.CopyString(ret, -1)

            End Select

            Return v
        End Get
        Set(value As String)
            ZoomUtil.ValidateOption(Me, name, value)

            Select Case name
                Case ZoomUtil.X_IGNOREMARCCHARACTERENCODING
                    bIgnoreMARCCharacterEncoding = CBool(value)

                Case ZoomUtil.X_ASSUMEDMARCCHARACTERENCODING
                    sAssumedMARCCharacterEncoding = CStr(value)

                Case Else

                    ZOOM_resultset_option_set(hZOOM_resultset, CStr(name), CStr(value))
                    zConnection.frndCheckForError()

            End Select
        End Set
    End Property

    Public Sub New()
        bIgnoreMARCCharacterEncoding = False
        sAssumedMARCCharacterEncoding = ZoomUtil.MARC_8
        hZOOM_resultset = 0
    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                zConnection = Nothing
            End If

            If hZOOM_resultset <> 0 Then
                ZOOM_resultset_destroy(hZOOM_resultset)
                hZOOM_resultset = 0
            End If
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

