Public Class ZoomScanSet
    Implements IDisposable

    Private hZOOM_scanset As Integer
    Private zConnection As ZoomConnection

    Private Declare Sub ZOOM_scanset_destroy Lib "YAZ5.dll" Alias "_ZOOM_scanset_destroy@4" (ByVal Scan As Integer)

    Private Declare Sub ZOOM_scanset_option_set Lib "YAZ5.dll" Alias "_ZOOM_scanset_option_set@12" (ByVal Scan As Integer, ByVal key As String, ByVal val As String)

    Private Declare Function ZOOM_scanset_option_get Lib "YAZ5.dll" Alias "_ZOOM_resultset_option_get@8" (ByVal Scan As Integer, ByVal key As String) As Integer

    Private Declare Function ZOOM_scanset_size Lib "YAZ5.dll" Alias "_ZOOM_scanset_size@4" (ByVal Scan As Integer) As Integer

    Private Declare Function ZOOM_scanset_term Lib "YAZ5.dll" Alias "_ZOOM_scanset_term@16" (ByVal Scan As Integer, ByVal pos As Integer, ByRef occ As Integer, ByRef length As Integer) As Integer




    Friend Sub frndSetConnection(c As ZoomConnection)
        zConnection = c
    End Sub

    Friend Sub frndSetScanSetHandle(h As Integer)

        If hZOOM_scanset <> 0 Then
            ZOOM_scanset_destroy(hZOOM_scanset)
            zConnection.frndCheckForError()
            hZOOM_scanset = 0
        End If

        hZOOM_scanset = h

    End Sub



    Public Function GetSize() As Integer
        GetSize = ZOOM_scanset_size(hZOOM_scanset)
        zConnection.frndCheckForError()
    End Function



    Public Function GetTerm(which As Integer) As String
        Dim l As Integer
        Dim occ As Integer
        Dim length As Integer

        l = ZOOM_scanset_term(hZOOM_scanset, which, occ, length)
        zConnection.frndCheckForError()

        GetTerm = ZoomUtil.CopyString(l, -1)

    End Function

    Public Function GetField(which As Integer, field As String) As String
        Dim l As Integer
        Dim occ As Integer
        Dim length As Integer
        Dim trm As String

        l = ZOOM_scanset_term(hZOOM_scanset, which, occ, length)
        zConnection.frndCheckForError()

        trm = ZoomUtil.CopyString(l, -1)

        Select Case CStr(FIELD)
            Case "freq"
                GetField = occ
            Case "display"
                GetField = trm
            Case "attrs"
                GetField = ""
            Case "alt"
                GetField = ""
            Case "other"
                GetField = ""
            Case Else
                Throw New ZoomException(Me, ZoomUtil.ERR_INVALID_SCAN_FIELD, CStr(field))
        End Select

    End Function

    Public Property ZoomOption(name As String) As String
        Get
            Dim ret As Integer
            Dim v As String

            ZoomUtil.ValidateOption(Me, name)

            ret = ZOOM_scanset_option_get(hZOOM_scanset, name)
            zConnection.frndCheckForError()

            v = ZoomUtil.CopyString(ret, -1)

            Return v
        End Get
        Set(value As String)

            ZoomUtil.ValidateOption(Me, name, value)


            ZOOM_scanset_option_set(hZOOM_scanset, name, value)
            zConnection.frndCheckForError()

        End Set
    End Property


    Public ReadOnly Property ErrorCode() As Integer
        Get
            ErrorCode = zConnection.ErrorCode
        End Get
    End Property


    Public ReadOnly Property ErrorMessage() As String
        Get
            ErrorMessage = zConnection.ErrorMessage
        End Get
    End Property


    Public ReadOnly Property AdditionalInfo() As String
        Get
            Return zConnection.AdditionalInfo
        End Get
    End Property

    Public Sub New()
        hZOOM_scanset = 0
    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                zConnection = Nothing
            End If

            If hZOOM_scanset <> 0 Then
                ZOOM_scanset_destroy(hZOOM_scanset)
                zConnection.frndCheckForError()
                hZOOM_scanset = 0
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
