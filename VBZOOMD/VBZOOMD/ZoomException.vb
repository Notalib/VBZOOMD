Public Class ZoomException
    Inherits Exception

    Public Property ZoomErrorCode

    Public Sub New(who As Object, errorCode As Integer, ParamArray params() As Object)
        MyBase.New(String.Format("Zoom Error: {0} {1} {2} {3} {4} {5}", errorCode, params))
        ZoomErrorCode = errorCode
    End Sub
End Class
