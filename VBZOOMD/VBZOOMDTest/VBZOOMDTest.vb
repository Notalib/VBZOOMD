Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports VBZOOMD
Imports System.Xml

<TestClass()> Public Class VBZOOMDTest

    ''' <summary>
    ''' Basic Query and USMARC Bib ID Search and Diacritics
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub TestZoom6()
        Dim zconn As New ZoomConnection("z3950.loc.gov", 7090)
        zconn.ZoomOption("databaseName") = "Voyager"
        zconn.ZoomOption("preferredRecordSyntax") = "USmarc"
        Dim zquery As New ZoomQuery("@attr 1=12 16358795") 'bib id search
        Dim zrs As ZoomResultSet = zconn.Search(zquery)
        Dim rec As String = zrs.GetRecord(0).RenderRecord
        Assert.IsTrue(rec.Contains("16358795"))
        Assert.IsTrue(rec.Contains("Gesänge aus Nur Eine löst den Zauberspruch, oder, Wer ist glücklich"))
    End Sub


    ''' <summary>
    ''' Basic Query and USMARC Render Record
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub TestZoom1()
        Dim zconn As New ZoomConnection("z3950.loc.gov", 7090)
        zconn.ZoomOption("databaseName") = "Voyager"
        zconn.ZoomOption("preferredRecordSyntax") = "USmarc"
        Dim zquery As New ZoomQuery("@attr 1=7 0253333490") 'isbn search
        Dim zrs As ZoomResultSet = zconn.Search(zquery)
        Dim rec As String = zrs.GetRecord(0).RenderRecord
        Assert.IsTrue(rec.Contains("0253333490"))
    End Sub

    ''' <summary>
    ''' Basic Query and USMARC XML Record
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub TestZoom2()
        Dim zconn As New ZoomConnection("z3950.loc.gov", 7090)
        zconn.ZoomOption("databaseName") = "Voyager"
        zconn.ZoomOption("preferredRecordSyntax") = "USmarc"
        Dim zquery As New ZoomQuery("@attr 1=7 0253333490") 'isbn search
        Dim zrs As ZoomResultSet = zconn.Search(zquery)
        Dim rec As String = zrs.GetRecord(0).XMLRecord
        Assert.IsTrue(rec.Contains("0253333490"))
    End Sub

    ''' <summary>
    ''' Basic Query and USMARC XML Data
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub TestZoom3()
        Dim zconn As New ZoomConnection("z3950.loc.gov", 7090)
        zconn.ZoomOption("databaseName") = "Voyager"
        zconn.ZoomOption("preferredRecordSyntax") = "USmarc"
        Dim zquery As New ZoomQuery("@attr 1=7 0253333490") 'isbn search
        Dim zrs As ZoomResultSet = zconn.Search(zquery)
        Dim rec As XmlDocument = zrs.GetRecord(0).XMLData
        Dim xmlStr As String = rec.OuterXml
        Assert.IsTrue(xmlStr.Contains("0253333490"))
    End Sub

    ''' <summary>
    ''' Basic Query and OPAC XML Data
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub TestZoom4()
        Dim zconn As New ZoomConnection("z3950.loc.gov", 7090)
        zconn.ZoomOption("databaseName") = "Voyager"
        zconn.ZoomOption("preferredRecordSyntax") = "OPAC"
        Dim zquery As New ZoomQuery("@attr 1=7 0253333490") 'isbn search
        Dim zrs As ZoomResultSet = zconn.Search(zquery)
        Dim rec As XmlDocument = zrs.GetRecord(0).XMLData
        Dim xmlStr As String = rec.OuterXml
        Assert.IsTrue(xmlStr.Contains("0253333490"))
    End Sub

    ''' <summary>
    ''' Test Connection Close Exception
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()> <ExpectedException(GetType(ZoomException))>
    Public Sub TestZoom5()
        Dim zconn As New ZoomConnection("z3950.loc.gov", 7090)
        zconn.ZoomOption("databaseName") = "Voyager"
        zconn.ZoomOption("preferredRecordSyntax") = "USmarc"
        Dim zquery As New ZoomQuery("@attr 1=7 0253333490") 'isbn search
        Dim zrs As ZoomResultSet = zconn.Search(zquery)
        Dim rec As String = zrs.GetRecord(0).RenderRecord
        Assert.IsTrue(rec.Contains("0253333490"))

        zconn.Close()

        zrs = zconn.Search(zquery)
        rec = zrs.GetRecord(0).RenderRecord
        Assert.IsTrue(rec.Contains("0253333490"))


    End Sub

End Class