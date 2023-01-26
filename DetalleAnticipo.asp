<!--#include virtual="/config/conn.asp"-->

<%

strSQL="select * from SolicitudAnticipos where Docnum="&request("ID")
rsUpdateEntry.Open strSQL, adoCon4

'response.write strSQL

CreateTable(600) 

response.write("<tr><td class='td-c subtitulo strong' colspan=2>DETALLE DE LA SOLICITUD DE ANTICIPO</td></tr>")
response.write("<tr><td class='td-r subtitulo'># Solicitud</td><td class='td-l'>"&rsUpdateEntry("ID") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Pedido</td><td class='td-l'>"&rsUpdateEntry("DocNum") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Razon Social</td><td class='td-l'>"&rsUpdateEntry("RS") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Codigo SN</td><td class='td-l'>"&rsUpdateEntry("CardCode") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Socio Negocio</td><td class='td-l'>"&rsUpdateEntry("CardName") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Asesor</td><td class='td-l'>"&rsUpdateEntry("SlpName") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>eMail Asesor</td><td class='td-l'>"&rsUpdateEntry("SlpEmail") &"</td></tr>")
'response.write("<tr><td class='td-r subtitulo'>Otros emails</td><td class='td-l'>"&rsUpdateEntry("Email") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Total Documento</td><td class='td-l'>"&rsUpdateEntry("Total") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Anticipo</td><td class='td-l'><font color=red>"&rsUpdateEntry("Anticipo") &"</font></td></tr>")
response.write("<tr><td class='td-r subtitulo'>Uso CFDi</td><td class='td-l'>"&rsUpdateEntry("Code") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Clave Metodo Pago</td><td class='td-l'>"&rsUpdateEntry("PayMethCod") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Instrucciones</td><td class='td-l'>"&rsUpdateEntry("Notas") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Fecha Solicitud</td><td class='td-l'>"&rsUpdateEntry("RequestDate") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Fecha respuesta</td><td class='td-l'>"&rsUpdateEntry("AnswerDate") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'># Anticipo en SAP</td><td class='td-l'>"&rsUpdateEntry("ODPI") &"</td></tr>")
closeTable
rsUpdateEntry.close
 
%>
<input type="button" value="cerrar" onclick="hidediv('detalle')"> </P>   