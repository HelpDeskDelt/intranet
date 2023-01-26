<!--#include virtual="/config/conn.asp"-->

<%

strSQL="select *,dbo.fn_GetMonthName(fechaDeposito,'spanish') as Fecha,'$ ' + CONVERT(VARCHAR,CONVERT(MONEY,importe),1)  as money from Pagos where ID="&request("ID")
rsUpdateEntry.Open strSQL, adoCon4

'response.write strSQL

CreateTable(600) 

response.write("<tr><td class='td-c subtitulo strong' colspan=2>DETALLE DE CONFIRMACION EN BANCOS</td></tr>")
response.write("<tr><td class='td-r subtitulo'># transaccion</td><td class='td-l'>"&rsUpdateEntry("ID") &"</td></tr>")

response.write("<tr><td class='td-r subtitulo'>Banco</td><td class='td-l'>"&rsUpdateEntry("Banco") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Fecha Deposito</td><td class='td-l'>"&rsUpdateEntry("fecha") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Importe</td><td class='td-l'>"&rsUpdateEntry("money") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Moneda</td><td class='td-l'>"&rsUpdateEntry("Moneda") &"</td></tr>")
response.write("<tr><td class='td-r subtitulo'>Fecha Captura</td><td class='td-l'>"&rsUpdateEntry("Logdate") &"</td></tr>")
closeTable
rsUpdateEntry.close
 
%>
<input type="button" value="cerrar" onclick="hidediv('detalle')"> </P>   