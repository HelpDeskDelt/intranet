<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->


<BR>
<%
Response.Buffer = True 
On Error Resume Next

strSQL="select * from Empleados where FechaEgreso is null"
rsUpdateEntry.Open strSQL, adoCon
rsUpdateEntry2.Open strSQL, adoCon
NoCampos=1
Dim campos(20)


Response.Write("<table border=0 cellpadding=3 cellspacing=2 align=center>")

Response.Write("</tr>")
For Each rsUpdateEntry in rsUpdateEntry.Fields
    Campos(NoCampos)=rsUpdateEntry.Name
    Response.Write "<td>" & Campos(NoCampos) & "</td>"
    NoCampos=NoCampos+1
Next
Response.Write("</tr>")


While Not rsUpdateEntry2.EOF
    Response.Write "<tr>"
    For i=1 to NoCampos
      Response.Write "<td>" & rsUpdateEntry2(campos(i)) & "</td>"      
    Next
    rsUpdateEntry2.MoveNext
    Response.Write "</tr>"
Wend
rsUpdateEntry2.close

Response.Write("</table>")


%>




</body>  


  <!-- menu script itself. you should not modify this file -->
  <script language="JavaScript" src="/menuj/menu.js"></script>
  <script language="JavaScript" src="/menuj/menu_items.js"></script>
  <script language="JavaScript" src="/menuj/menu_tpl.js"></script>
</p>
<br>
<script language="JavaScript">
  <!--//

  // some table cell or other HTML element. Always put it before </body>
  // each menu gets two parameters (see demo files)
  // 1. items structure
  // 2. geometry structure

  new menu (MENU_ITEMS, MENU_POS);
  
  //-->
</script>

