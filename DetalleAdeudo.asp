<!--#include virtual="/config/conn.asp"-->

<%

strSQL=" "

response.write strSQL
'rsUpdateEntry.Open strSQL, adoCon


CreateTable(900) 

response.write("<tr>")
response.write("<td class='td-c subtitulo'>#</td>")
response.write("<td class='td-c subtitulo'>Tipo Cotizacion</td>")
response.write("<td class='td-c subtitulo'>Fecha</td>")
response.write("<td class='td-c subtitulo'>Socio Negocio</td>")
response.write("<td class='td-c subtitulo'>Codigo</td>")
response.write("<td class='td-c subtitulo'>Cantidad</td>")
response.write("<td class='td-c subtitulo'>Precio</td>")
response.write("<td class='td-c subtitulo'>Subtotal</td>")
response.write("<td class='td-c subtitulo'>CFDi</td>")
response.write("<td class='td-c subtitulo'>Pedido</td>")
response.write("</tr>")



'while not rsUpdateEntry.EOF
   
  'rsUpdateEntry.moveNext
  
'wend
'rsUpdateEntry.close

 
%>
<input type="button" value="cerrar" onclick="hidediv('detalle')"> </P>   