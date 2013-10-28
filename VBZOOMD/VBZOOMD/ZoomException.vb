Public Class ZoomException
    Inherits Exception

    Public Property ZoomErrorCode

    Public Sub New(who As Object, errorCode As Integer, ParamArray params() As Object)

        MyBase.New(String.Format("Zoom Error: {0} {1}", errorCode, String.Join(", ", params.ToArray)))
        ZoomErrorCode = errorCode
    End Sub
End Class
