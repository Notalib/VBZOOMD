Public Class ZoomConnection
    Implements IDisposable


    Private bConnected As Boolean

    Private sHost As String
    Private lPortNum As Integer

    Private hZOOM_connection As Integer

    Private Declare Function ZOOM_connection_new Lib "YAZ5.dll" Alias "_ZOOM_connection_new@8" (ByVal Host As String, ByVal portnum As Integer) As Integer

    Private Declare Function ZOOM_connection_create Lib "YAZ5.dll" Alias "_ZOOM_connection_create@4" (ByVal Options As Integer) As Integer

    Private Declare Sub ZOOM_connection_connect Lib "YAZ5.dll" Alias "_ZOOM_connection_connect@12" (ByVal c As Integer, ByVal Host As String, ByVal portnum As Integer)

    Private Declare Sub ZOOM_connection_destroy Lib "YAZ5.dll" Alias "_ZOOM_connection_destroy@4" (ByVal c As Integer)

    Private Declare Sub ZOOM_connection_option_set Lib "YAZ5.dll" Alias "_ZOOM_connection_option_set@12" (ByVal c As Integer, ByVal key As String, ByVal val As String)

    Private Declare Function ZOOM_connection_option_get Lib "YAZ5.dll" Alias "_ZOOM_connection_option_get@8" (ByVal c As Integer, ByVal key As String) As Integer

    Private Declare Function ZOOM_connection_error Lib "YAZ5.dll" Alias "_ZOOM_connection_error@12" (ByVal c As Integer, ByRef cp As Integer, ByRef addinfo As Integer) As Integer
    Private Declare Function ZOOM_connection_errcode Lib "YAZ5.dll" Alias "_ZOOM_connection_errcode@4" (ByVal c As Integer) As Integer
    Private Declare Function ZOOM_connection_errmsg Lib "YAZ5.dll" Alias "_ZOOM_connection_errmsg@4" (ByVal c As Integer) As Integer
    Private Declare Function ZOOM_connection_addinfo Lib "YAZ5.dll" Alias "_ZOOM_connection_addinfo@4" (ByVal c As Integer) As Integer

    Private Declare Function ZOOM_connection_search Lib "YAZ5.dll" Alias "_ZOOM_connection_search@8" (ByVal c As Integer, ByVal q As Integer) As Integer

    Private Declare Function ZOOM_connection_search_pqf Lib "YAZ5.dll" Alias "_ZOOM_connection_search_pqf@8" (ByVal c As Integer, ByVal q As String) As Integer

    Private Declare Function ZOOM_connection_scan Lib "YAZ5.dll" Alias "_ZOOM_connection_scan@8" (ByVal c As Integer, ByVal q As String) As Integer

    Public ReadOnly Property Host() As String
        Get
            Return sHost
        End Get
    End Property


    Public ReadOnly Property Port() As Integer
        Get
            Return lPortNum
        End Get
    End Property


    Private Sub priCheckConnection()
        If bConnected = False Then
            ZOOM_connection_connect(hZOOM_connection, sHost, lPortNum)
            frndCheckForError()
            bConnected = True
        End If
    End Sub

    Public ReadOnly Property ErrorCode() As Integer
        Get
            Return ZOOM_connection_errcode(hZOOM_connection)
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Dim m As Integer

            m = ZOOM_connection_errmsg(hZOOM_connection)

            Return CopyString(m, -1)
        End Get
    End Property

    Public ReadOnly Property AdditionalInfo() As String
        Get
            Dim m As Integer

            m = ZOOM_connection_addinfo(hZOOM_connection)

            Return CopyString(m, -1)
        End Get
    End Property

    Friend Sub frndSetHost(h As String)
        sHost = h
    End Sub

    Friend Sub frndSetPortNum(p As Integer)
        lPortNum = p
    End Sub


    Friend Sub frndCheckForError()
        Dim zec As Integer

        zec = ErrorCode

        If zec <> 0 Then
            Throw New ZoomException(Me, ERR_ZOOM_ERROR, zec, ErrorMessage, AdditionalInfo)
        End If
    End Sub

    'Public Function Search(myqry As Object) As ZoomResultSet
    '    Dim qry As ZoomQuery
    '    Dim hzrs As Integer
    '    Dim zrs As New ZoomResultSet

    '    qry = myqry

    '    priCheckConnection()

    '    hzrs = ZOOM_connection_search(hZOOM_connection, qry.frndGetQueryHandle)
    '    frndCheckForError()

    '    zrs.frndSetConnection(Me)
    '    zrs.frndSetResultSetHandle(hzrs)

    '    Return zrs
    'End Function

    'Public Function Scan(qry As Object) As ZoomScanSet
    '    Dim hzss As Integer
    '    Dim zss As New ZoomScanSet

    '    priCheckConnection()

    '    hzss = ZOOM_connection_scan(hZOOM_connection, CStr(qry))
    '    frndCheckForError()

    '    zss.frndSetConnection(Me)
    '    zss.frndSetScanSetHandle(hzss)

    '    Return zss
    'End Function

    Public Property ZoomOption(name As String) As String
        Get
            Dim ret As Integer
            Dim v As String

            ValidateOption(Me, name)

            ret = ZOOM_connection_option_get(hZOOM_connection, name)
            frndCheckForError()

            v = CopyString(ret, -1)

            Return v
        End Get
        Set(value As String)
            ValidateOption(Me, name, value)

            ZOOM_connection_option_set(hZOOM_connection, name, value)
            frndCheckForError()

        End Set
    End Property

    Public Sub New()
        Me.New("localhost", 210)
    End Sub

    Public Sub New(host As String)
        Me.New(host, 210)
    End Sub


    Public Sub New(host As String, port As Integer)
        sHost = host
        lPortNum = port

        'create a new connection w/o actually connecting
        'connection is deferred until the first search or scan is performed
        hZOOM_connection = ZOOM_connection_create(0)
        frndCheckForError()

    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            If hZOOM_connection <> 0 Then
                ZOOM_connection_destroy(hZOOM_connection)
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
