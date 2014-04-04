<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head><title>
	Get Bib Record Web Service
</title></head>
<body>
    <h1>
      Get Bib Record Web Service
    </h1>

    <p>
    This is a web service used to retrieve data from the UIUC Voyager Library Catalog given a Bib Id, Call Number, or Barcode Number.  It supports one action with three kinds of 
    identifiers and four possible return formats.  The URL format is as follows:  <b>/getmarc/one.aspx/identifier.ext[?v=true]</b>.
    </p>

    <p>
      This will retrieve a Bibligraphic Record corresponding to the given identifier in the format corresponding to given ext value.  The identifier can be one of the following:
    </p>

    <dl>
    <dt>Bib ID</dt>
    <dd>Any value which is a valid integer is assumed to be a Bib Id.  This is the most reliable way to retreive a MARC record since a Bib Id corresponds to one and only 
    one MARC Record.</dd>

    <dt>Barcode Number</dt>
    <dd>Any value of exactly 14 digits is assumed to be a Barcode Number.  There may be cases, such as when multiple bibliographic or intellectual volumes are 
    bound in a single physical volume, when a barcode corresponds to multiple Bib Records.  When this happens only one of the Bib Records will be returned, usually the
    one with largest Bib Id.  For an alternate service that deals with Barcodes see the <a href="/barcode/">Barcode Web Service</a></dd>

    <dt>Call Number</dt>
    <dd>Any other value is assumed to be a Call Number. For the same reason given above for Barcode Numbers, this may result in multiple Bib Records; only one of which will be 
    returned.  Call Numbers can also be problematic because of inconsistent punctuation, spaces, and prefixes.  Usually prefixes like Q or F or quarto or folio should be
    omitted.  Where possible normalized spacing and punctuation should be used.  Finally, because of web server restructions, any colon character contained in a Call Number must
    be spelled out as __COLON__.  For example, this call number, 738 C817 v.12:75, would be converted to this, 738 C817 v.12__COLON__75, before used in a URL.</dd>
    </dl>

    <p>
    The format extension value can be one of the following:
    </p>

    <dl>
    <dt>marc</dt><dd>This extension will cause records formatted as MARC XML to be returned.</dd>
    <dt>opac</dt><dd>This extension will cause records formatted as OPAC XML to be returned.  These are similar to MARC XML except they include holdings and
    circulation data.</dd>
    <dt>text</dt><dd>This extension will cause records formatted as plain text to be returned.</dd>
    <dt>mods</dt><dd>This extension will cause records formatted as MODS XML to be returned.  These are the MARC XML transformed into MODS using the standard
     XSLT from the Library of Congress.</dd>
    <dt>dc</dt><dd>This extension will cause records formatted as OAI DC (Dublin Core) to be returned.  These are the MARC XML transformed into DC using the standard
     XSLT from the Library of Congress.</dd>
   </dl>
   <p>If the extension is ommitted, marc is assumed.</p>

    <p>
        The optional [?v=true] parameter indicates that the MARC XML returned from the Z39.50 target should be validated against the XML schema before further processing.
        Note that only the source MARC XML will be validated, even if the selected return type is MODS, DC, or some other XML schema.
    </p>

    <h2>Here are some examples:</h2>
    <h3>Retreive records based on a Bib Id</h3>
    <ul>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.marc"})%>">one/1099891.marc</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.marc", .v = True})%>">one/1099891.marc?v=true</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.text"})%>">one/1099891.text</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.opac"})%>">one/1099891.opac</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.mods"})%>">one/1099891.mods</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.dc"})%>">one/1099891.dc</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "1099891.html"})%>">one/1099891.html</a></li>
    </ul>

    <h3>Retreive records based on a Barcode Number</h3>
    <ul>
      <li><a href="<%:Url.Action("one", New With {.id = "30112024718303.marc"})%>">one/30112024718303.marc</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "30112024718303.text"})%>">one/30112024718303.text</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "30112024718303.opac"})%>">one/30112024718303.opac</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "30112024718303.mods"})%>">one/30112024718303.mods</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "30112024718303.dc"})%>">one/30112024718303.dc</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "30112024718303.html"})%>">one/30112024718303.html</a></li>
    </ul>

    <h3>Retreive records based on a Call Number</h3>
    <p>This call number is actually "738 C817 v.12:75", but because of the colon is must be turned into "738 C817 v.12__COLON__75".  Also, this is an example of a Call Number
    that corresponds to multiple Bib Records where only one of those records will be returned.</p>
    <ul>
      <li><a href="<%:Url.Action("one", New With {.id = "738 C817 v.13__COLON__4.marc"})%>">one/738 C817 v.13__COLON__4.marc</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "738 C817 v.13__COLON__4.text"})%>">one/738 C817 v.13__COLON__4.text</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "738 C817 v.13__COLON__4.opac"})%>">one/738 C817 v.13__COLON__4.opac</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "738 C817 v.13__COLON__4.mods"})%>">one/738 C817 v.13__COLON__4.mods</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "738 C817 v.13__COLON__4.dc"})%>">one/738 C817 v.13__COLON__4.dc</a></li>
      <li><a href="<%:Url.Action("one", New With {.id = "738 C817 v.13__COLON__4.html"})%>">one/738 C817 v.13__COLON__4.html</a></li>
    </ul>
    <p>
    All of the formats which are returned will have one or more of the folloowing XML processing instructions (PIs) appended to the end:
    </p>
    <dl>
      <dt>&lt;?barcode ____ ?></dt><dd>If the identifier was a barcode, this PI will contain that barcode value.</dd>
      <dt>&lt;?bibcount ____ ?></dt><dd>If the identifier was a barcode, this PI will contain the number of Bib Records correspodning to that barcode value.</dd>
      <dt>&lt;?query ____ ?></dt><dd>This will contain the value of the actually Z39.50 query in <a href-"http://www.indexdata.com/yaz/doc/tools.html#PQF">PQF syntax</a>.</dd>
      <dt>&lt;?count ____ ?></dt><dd>This will contain the number of records found by the PQF query, even though only one of those results will be returned.</dd>
    </dl>
</body>
</html>