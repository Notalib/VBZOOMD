Public Class ZoomQuery
    Implements IDisposable

    Private hZOOM_query As Integer

    Private sQry As String
    Private eTyp As Integer

    Private Declare Function ZOOM_query_create Lib "YAZ5.dll" Alias "_ZOOM_query_create@0" () As Integer

    Private Declare Sub ZOOM_query_destroy Lib "YAZ5.dll" Alias "_ZOOM_query_destroy@4" (ByVal q As Integer)

    Private Declare Function ZOOM_query_prefix Lib "YAZ5.dll" Alias "_ZOOM_query_prefix@8" (ByVal q As Integer, ByVal str As String) As Integer

    Private Declare Function ZOOM_query_cql Lib "YAZ5.dll" Alias "_ZOOM_query_cql@8" (ByVal q As Integer, ByVal str As String) As Integer

    Private Declare Function ZOOM_query_sortby Lib "YAZ5.dll" Alias "_ZOOM_query_sortby@8" (ByVal q As Integer, ByVal criteria As String) As Integer

    Friend Function frndGetQueryHandle() As Integer
        frndGetQueryHandle = hZOOM_query
    End Function


    Friend Sub frndSetQueryString(s As String)
        Dim ret As Integer

        If Len(s) = 0 Then
            Throw New ZoomException(Me, ZoomUtil.ERR_INVALID_QUERY, s)
        End If

        eTyp = 0
        ret = ZOOM_query_prefix(hZOOM_query, s)

        If ret = -1 Then
            Throw New ZoomException(Me, ZoomUtil.ERR_INVALID_QUERY, s)
        End If

        sQry = s
    End Sub

    Friend Sub frndSetQueryStringCQL(s As String)
        Dim ret As Integer

        If Len(s) = 0 Then
            Throw New ZoomException(Me, ZoomUtil.ERR_INVALID_QUERY, s)
        End If

        eTyp = 1
        ret = ZOOM_query_cql(hZOOM_query, s)

        If ret = -1 Then
            Throw New ZoomException(Me, ZoomUtil.ERR_INVALID_QUERY, s)
        End If

        sQry = s
    End Sub


    Public ReadOnly Property QueryString() As String
        Get
            Return sQry
        End Get
    End Property

    Public ReadOnly Property QueryType() As Integer
        Get
            Return eTyp
        End Get
    End Property

    Public Sub New()
        hZOOM_query = ZOOM_query_create
    End Sub

    Public Sub New(qry As String)
        Me.New(qry, ZoomUtil.Z_PQF)
    End Sub

    Public Sub New(qry As String, typ As Integer)
        Me.New()

        Select Case typ
            Case ZoomUtil.Z_PQF
                frndSetQueryString(qry)

            Case ZoomUtil.Z_CQL
                frndSetQueryStringCQL(qry)

            Case Else
                Throw New ZoomException(Me, ZoomUtil.ERR_INVALID_QUERY_TYPE, CInt(typ))

        End Select
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' No managed objects to dispose.
            End If

            If hZOOM_query <> 0 Then
                ZOOM_query_destroy(hZOOM_query)
            End If
            hZOOM_query = 0

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
