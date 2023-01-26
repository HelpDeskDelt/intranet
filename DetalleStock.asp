<!--#include virtual="/config/conn.asp"-->

<%

strSQL="select * from _inv2 where ItemCode='"&request("codigo") &"' and Warehouse='"&request("almacen") &"' "
'response.write strSQL
rsUpdateEntry.Open strSQL, adoCon4

DoTitulo("Detalle del stock en razones sociales")
CreateTable(600) 

response.write("<tr>")
response.write("<td class='td-c subtitulo'>Razon Social</td>")
response.write("<td class='td-c subtitulo'>Codigo</td>")
response.write("<td class='td-c subtitulo'>Familia</td>")
response.write("<td class='td-c subtitulo'>Almacen</td>")
response.write("<td class='td-c subtitulo'>Stock</td>")
response.write("</tr>")


while not rsUpdateEntry.EOF
   response.write("<tr>")
   response.write("<td class='td-c fontMed'>"& rsUpdateEntry("RS") &"</td>")  
   response.write("<td class='td-l fontMed'>"& rsUpdateEntry("ItemCode") &"</td>")   
   response.write("<td class='td-c fontMed'>"& rsUpdateEntry("ItmsGrpNam") &"</td>")  
   response.write("<td class='td-c fontMed'>"& rsUpdateEntry("warehouse") &"</td>")  
   response.write("<td class='td-r fontMed'>"& rsUpdateEntry("stock") &"</td>")  
   response.write("</tr>")
  rsUpdateEntry.moveNext 
wend
rsUpdateEntry.close
closeTable

%>
<input type="button" value="cerrar" onclick="hidediv('stock')"> </P>   

