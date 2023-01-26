<!--#include virtual="/config/conn.asp"-->

<%
   Response.Expires = -1
   vString="monto"&request("ID")
   Response.write("<input type='text' name='"&vString&"' value='"&request("monto")&"'>")     
%>

