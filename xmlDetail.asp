<!--#include virtual="/config/conn.asp"-->


<html>
<!DOCTYPE html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-13">
<title>Intranet Deltaflow</title>

<link rel="stylesheet" href="/CSS/style.css">



</head>
<body>	





<P class='td-c azul fontmed'>CONCEPTOS DE <B><%=request("UUID")%>&nbsp; <input type="button" value="cerrar" onclick="hidediv('data')"> </B></H1>

<%
   Response.Expires = -1
  
    strSQL="select * from xml where UUID='"& request("UUID") &"' and carga is not null"
   'response.write strSQL  
	rsUpdateEntry.Open strSQL, adoCon4

	response.write("<P style='text-align: text-align: left'>"&  rsUpdateEntry("CONCEPTOS1") &  rsUpdateEntry("CONCEPTOS2") &  rsUpdateEntry("CONCEPTOS3") &"</P>")
    response.write("<P>&nbsp; </P>")

   'response.write(request("UUID"))
   rsUpdateEntry.close


%>

<P class='td-c'><input type="button" value="cerrar" onclick="hidediv('data')">