Imports System.IO
Imports VBZOOMD
Imports System.Xml
Imports System.Xml.Xsl
Imports System.Net

Public Class HomeController
    Inherits System.Web.Mvc.Controller

    Dim postscript As String = ""

    Function Index() As ActionResult
        Return View()
    End Function

    Function One(id As String) As ActionResult
        Dim cnum As String = id 'Request.PathInfo
        Dim barcode As String = ""

        'for testing 
        'cnum = "/4231937"
        'cnum = "/912 B41p.opc"
        'cnum = "/G8300182 .M3"
        'cnum = "/G4071 .P3 1863 .C4"
        'cnum = "/G3700s.C2 1882 .U5 Peoria, Ill."
        'cnum = "/1099891.opc"
        'cnum = "/30112024085216"
        'cnum = "/738 C817 v.12__COLON__75.opac"


        Dim okExts As String() = {".xml", ".marc", ".mrc", ".txt", ".text", ".opc", ".opac", ".mod", ".mods", ".dc", ".htm", ".html"}
        Dim ext As String = Path.GetExtension(cnum).ToLower
        If ext = String.Empty Or Not okExts.Contains(ext) Then
            cnum = cnum & ".xml"
            ext = ".xml"
        End If

        'TODO: The VBZOOMD Library no longer support the GetField method, so the queries will not work
        Dim qs As String = Server.UrlDecode(Request.QueryString.ToString)
        Dim query As String = ""
        Dim xmlns As String = ""
        Dim qtype As String = ""
        If qs.Length > 0 Then
            qtype = ParseQueryString(qs, query, xmlns)
        End If

        If cnum.Length > 1 AndAlso okExts.Contains(ext) Then

            'cnum = cnum.Substring(1)


            Dim marc_id As String = Path.GetFileNameWithoutExtension(cnum) 'cnum.Substring(0, cnum.Length - 4)
            marc_id = DecodeId(marc_id)

            Dim id2 As String = CloseDigits(marc_id)
            Dim id3 As String = TruncateIll(marc_id)
            Dim id4 As String = TruncateIll(id2)
            Dim id5 As String = DropPrefix(marc_id)

            If marc_id Like "##############" Then
                'probably a barcode
                barcode = marc_id
                marc_id = GetBibIdForBarcode(marc_id)
                If marc_id = barcode Then
                    barcode = ""
                End If
            End If

            Dim qry(2) As String
            Dim qnum As Integer
            If IsNumeric(marc_id) Then
                qnum = 0
                qry(0) = "@attr 1=12" & " """ & marc_id & """" 'bibid
            Else
                qnum = 2
                qry(0) = "@attr 3=1 @attr 1=13" & " """ & marc_id & """" 'Dewey classification
                qry(1) = "@attr 3=1 @attr 1=16" & " """ & marc_id & """" 'LC classification
                qry(2) = "@attr 3=3 @attr 1=16" & " """ & marc_id & """" 'LC classification
                If Not String.IsNullOrEmpty(id2) Then
                    ReDim Preserve qry(5)
                    qnum = 5
                    qry(3) = "@attr 3=1 @attr 1=13" & " """ & id2 & """" 'Dewey classification
                    qry(4) = "@attr 3=1 @attr 1=16" & " """ & id2 & """" 'LC classification
                    qry(5) = "@attr 3=3 @attr 1=16" & " """ & id2 & """" 'LC classification
                End If
                If Not String.IsNullOrEmpty(id3) Then
                    ReDim Preserve qry(8)
                    qnum = 8
                    qry(6) = "@attr 3=1 @attr 1=13" & " """ & id3 & """" 'Dewey classification
                    qry(7) = "@attr 3=1 @attr 1=16" & " """ & id3 & """" 'LC classification
                    qry(8) = "@attr 3=3 @attr 1=16" & " """ & id3 & """" 'LC classification
                End If
                If Not String.IsNullOrEmpty(id4) Then
                    ReDim Preserve qry(11)
                    qnum = 11
                    qry(9) = "@attr 3=1 @attr 1=13" & " """ & id4 & """" 'Dewey classification
                    qry(10) = "@attr 3=1 @attr 1=16" & " """ & id4 & """" 'LC classification
                    qry(11) = "@attr 3=3 @attr 1=16" & " """ & id4 & """" 'LC classification
                End If
                If Not String.IsNullOrEmpty(id5) Then
                    ReDim Preserve qry(14)
                    qnum = 14
                    qry(12) = "@attr 3=1 @attr 1=13" & " """ & id5 & """" 'Dewey classification
                    qry(13) = "@attr 3=1 @attr 1=16" & " """ & id5 & """" 'LC classification
                    qry(14) = "@attr 3=3 @attr 1=16" & " """ & id5 & """" 'LC classification
                End If
            End If



            Dim zc As New ZoomConnection(ConfigurationManager.AppSettings.Item("ZHost"), ConfigurationManager.AppSettings.Item("ZPort"))
            zc.ZoomOption("databaseName") = ConfigurationManager.AppSettings.Item("ZDatabase")
            If ext = ".opc" Or ext = ".opac" Then
                zc.ZoomOption("preferredRecordSyntax") = "OPAC"
            Else
                zc.ZoomOption("preferredRecordSyntax") = "USmarc"
            End If
            zc.ZoomOption("user") = ConfigurationManager.AppSettings.Item("ZUser")
            zc.ZoomOption("elementSetName") = ConfigurationManager.AppSettings.Item("ZElementSetName")
            zc.ZoomOption("timeout") = ConfigurationManager.AppSettings.Item("ZTimeout")

            Dim zquery As ZoomQuery = Nothing
            Dim zrs As ZoomResultSet = Nothing
            For tries As Integer = 0 To qnum
                Try
                    zquery = New ZoomQuery(qry(tries))
                    zrs = zc.Search(zquery)
                Catch ex As Exception
                    If zrs IsNot Nothing Then zrs.Dispose()
                    zrs = Nothing
                    If zquery IsNot Nothing Then zquery.Dispose()
                    zquery = Nothing
                    If tries >= qnum Then
                        ReturnHTTPStatus(502, "Z39.50 Error: " & ex.Message)
                        Return New HttpStatusCodeResult(502, "Bad Gateway--Z39.50")
                        If zc IsNot Nothing Then
                            zc.Close()
                            zc.Dispose()
                        End If
                        zc = Nothing
                        Exit Function
                    End If
                End Try
                If zrs IsNot Nothing AndAlso zrs.GetSize > 0 Then
                    postscript = postscript & "<?query " & qry(tries).Replace("-", " ") & " ?>" & vbCrLf
                    Exit For
                End If
            Next

            If zrs.GetSize > 0 Then
                postscript = postscript & "<?count " & zrs.GetSize & " ?>" & vbCrLf
                Dim zr As ZoomRecord = zrs.GetRecord(0)

                If ext = ".txt" Or ext = ".text" Then
                    Response.ContentType = "text/plain"
                    If qtype = "marcpath" Then
                        Try
                            Response.Write(zr.GetField(query))
                        Catch ex As Exception
                            ReturnHTTPStatus(400, "MARCPath Error: " & ex.Message)
                        End Try
                    ElseIf qtype = "xpath1" Then
                        ReturnHTTPStatus(400, "XPointer scheme 'xpath1' is not supported for plain text MARC records; use 'marcpath' instead")
                    ElseIf qtype <> "" Then
                        ReturnHTTPStatus(400, "Unrecognized XPointer scheme '" & qtype & "'; use 'marcpath' instead")
                    Else
                        Response.Write(zr.RenderRecord)
                    End If
                ElseIf ext = ".xml" Or ext = ".mrc" Or ext = ".marc" Or ext = ".opc" Or ext = ".opac" Then
                    Response.ContentType = "application/xml"
                    If qtype = "xpath1" Then
                        Try
                            Dim xfrag As String = zr.GetField(query, xmlns)
                            If xfrag.StartsWith("<record") Then
                                Response.Write(xfrag)
                            Else
                                Response.Write("<f:fcs xmlns:f='http://www.w3.org/2001/02/xml-fragment' ")
                                Response.Write("extref='http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd' ")
                                Response.Write("parentref='http://" & Request.Url.Host & Request.Path & "'>" & vbCrLf)
                                Response.Write(xfrag)
                                Response.Write("</f:fcs>")
                            End If
                        Catch ex As Exception
                            ReturnHTTPStatus(400, "XPath Error: " & ex.Message)
                        End Try
                    ElseIf qtype = "marcpath" Then
                        ReturnHTTPStatus(400, "XPointer scheme 'marcpath' is not supported for XML MARC records; use 'xpath1' instead")
                    ElseIf qtype <> "" Then
                        ReturnHTTPStatus(400, "Unrecognized XPointer scheme '" & qtype & "'; use 'xpath1' instead")
                    Else
                        Try
                            Dim xwr As New XmlTextWriter(Response.OutputStream, Text.Encoding.UTF8)
                            xwr.Formatting = Formatting.Indented
                            zr.XMLData.Save(xwr)
                            'Response.Write(zr.XMLData.OuterXml)
                        Catch ex As Exception
                            If ext = ".opc" Or ext = ".opac" Then
                                ReturnHTTPStatus(500, "The OPAC format is not be available for this record try the MARC XML format instead:  <a href='" & marc_id & ".xml'>" & marc_id & ".xml</a>")
                            Else
                                ReturnHTTPStatus(500, "Internal Error: " & ex.Message & "  Try the SUTRS format instead:  <a href='" & marc_id & ".txt'>" & marc_id & ".txt</a>")
                            End If
                        End Try
                    End If
                ElseIf ext = ".mod" Or ext = ".mods" Then
                    Response.ContentType = "application/xml"

                    Dim xml As New XmlDocument
                    xml.LoadXml(zr.XMLData.OuterXml)

                    Dim xsl As New XslCompiledTransform
                    xsl.Load(HttpContext.Server.MapPath("~/app_data/MARC21slim2MODS3-4.xsl"))

                    xsl.Transform(New XmlNodeReader(xml), Nothing, Response.OutputStream)

                ElseIf ext = ".htm" Or ext = ".html" Then
                    Response.ContentType = "text/html"

                    Dim xml As New XmlDocument
                    xml.LoadXml(zr.XMLData.OuterXml)

                    Dim xsl As New XslCompiledTransform
                    xsl.Load(HttpContext.Server.MapPath("~/app_data/MARC21slim2English.xsl"))

                    xsl.Transform(New XmlNodeReader(xml), Nothing, Response.OutputStream)

                ElseIf ext = ".dc" Then
                    Response.ContentType = "application/xml"

                    Dim xml As New XmlDocument
                    xml.LoadXml(zr.XMLData.OuterXml)

                    Dim xsl As New XslCompiledTransform
                    xsl.Load(HttpContext.Server.MapPath("~/app_data/MARC21slim2OAIDC.xsl"))

                    xsl.Transform(New XmlNodeReader(xml), Nothing, Response.OutputStream)

                Else
                    ReturnHTTPStatus(404, "Could not find record number " & marc_id)
                End If

                If zr IsNot Nothing Then zr.Dispose()
                If zrs IsNot Nothing Then zrs.Dispose()
                If zquery IsNot Nothing Then zquery.Dispose()
                If zc IsNot Nothing Then
                    zc.Close()
                    zc.Dispose()
                End If

                zr = Nothing
                zrs = Nothing
                zquery = Nothing
                zc = Nothing

            Else
                ReturnHTTPStatus(404, "Could not find record number " & marc_id)
            End If
        Else
            ReturnHTTPStatus(404, "Could not find record number " & cnum)
        End If
        Response.Write(vbCrLf & postscript)
        Return New EmptyResult
    End Function

    Private Sub ReturnHTTPStatus(ByVal code As Long, ByVal msg As String)
        Response.StatusCode = code
        If msg.Contains("</") Then
            Response.ContentType = "text/html"
        Else
            Response.ContentType = "text/plain"
        End If
        Response.Write(msg)
    End Sub

    Private Function DecodeId(id As String) As String
        id = Server.UrlDecode(id)
        id = id.Replace("__COLON__", ":")
        id = id.Replace("__colon__", ":")
        id = id.Replace("__Colon__", ":")
        Return id
    End Function

    Private Function ParseQueryString(ByVal qs As String, ByRef query As String, ByRef xmlns As String) As String
        Dim ret As String = ""
        Dim fnd As Boolean = False
        Dim reXMLNS As New Regex("([^=]+)\s*=\s*(\S+)")
        Dim reXPointer As New Regex("([^(]+)\(([^\)]+)\)")

        Dim ms As MatchCollection
        Dim m As Match
        Dim m2 As Match

        ms = reXPointer.Matches(qs)

        For Each m In ms
            If m.Groups(1).Value = "xpath1" And Not (fnd) Then
                fnd = True
                ret = m.Groups(1).Value
                query = m.Groups(2).Value
            ElseIf m.Groups(1).Value = "marcpath" And Not (fnd) Then
                fnd = True
                ret = m.Groups(1).Value
                query = m.Groups(2).Value
            ElseIf m.Groups(1).Value = "xmlns" Then
                m2 = reXMLNS.Match(m.Groups(2).Value)
                If m2.Success Then
                    xmlns = xmlns & "xmlns:" & m2.Groups(1).Value & "='" & m2.Groups(2).Value & "' "
                End If
            ElseIf Not (fnd) Then
                ret = m.Groups(1).Value
                query = m.Groups(2).Value
            End If
        Next

        Return ret
    End Function

    Private Function GetBibIdForBarcode(barcode As String) As String
        Dim hreq As HttpWebRequest = HttpWebRequest.Create(String.Format(ConfigurationManager.AppSettings.Item("BarcodeURL"), barcode))
        Dim cnt As Integer = 0
        Try
            Dim hres As HttpWebResponse = hreq.GetResponse
            Dim bibids As String = ""

            Dim strr As StreamReader = New StreamReader(hres.GetResponseStream)
            Do Until strr.EndOfStream
                Dim ln As String = strr.ReadLine
                If Not String.IsNullOrEmpty(ln.Trim) Then
                    bibids = ln.Trim
                    cnt = cnt + 1
                End If
            Loop
            hres.Close()
            If cnt > 0 Then
                postscript = postscript & "<?barcode " & barcode & " ?>" & vbCrLf
                postscript = postscript & "<?bibcount " & cnt & " ?>" & vbCrLf
                Return bibids
            Else
                Return barcode
            End If
        Catch wex As WebException
            Return barcode
        End Try

    End Function

    ''' <summary>
    ''' Any digits separated by just whitespace should be closed up
    ''' </summary>
    ''' <param name="num"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CloseDigits(num As String) As String
        Dim re As New Regex("(\d)\s+(\d)")

        If re.IsMatch(num) Then
            Return re.Replace(num, "$1$2")
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Some call numbers seem to be missing the trailing state designator, such as G3700s.C2 1882 .U5 Peoria, Ill.
    ''' </summary>
    ''' <param name="num"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TruncateIll(num As String) As String
        Dim re As New Regex(",?\s+Ill\.?\s*$")

        If re.IsMatch(num) Then
            Return re.Replace(num, "")
        Else
            Return ""
        End If

    End Function

    Private Function DropPrefix(num As String) As String
        If num.StartsWith("X") Then
            Return Mid(num, 2).Trim
        Else
            Return ""
        End If
    End Function
End Class