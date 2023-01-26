<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->




<%
  Response.Expires = -1
  response.redirect("ShowContent.asp?contenido=65")

  response.write("<P>&nbsp;</P><P>&nbsp;</P><center><table width=900px>")
  strSQL="select Rate,RateDate from ORTT where year(RateDate)=year(getdate()) and month(RateDate)=month(getdate()) and day(RateDate)=day(getdate())"
  'response.write strSQL
  rsUpdateEntry.Open strSQL, adoCon2 'DMX'

  vRate=""
  if request("frate") <>"" then
          vRate=request("frate")
  else
          vRate=trim(rsUpdateEntry("Rate"))
  end if

  response.write ("<form action='/modules/costeo.asp' method='post'>")
  response.write("<tr><td>El tipo de cambio para el d&iacute;a de hoy es: <input type='text' name='frate' id='frate' value='"& vRate &"'></td></tr>")
  response.write("<tr><td><B>Pegue el listado de sus partidas en esta &aacute;rea:</B></td></tr>")
  'response.write ("<input type='hidden' name='frate' id='frate' value='"&  trim(rsUpdateEntry("Rate")) &"'>")
  rsUpdateEntry.close    %>
  
  <tr><td>
  <textarea rows="20" cols="25" id="fpartidas" name="fpartidas"><%=request("fpartidas")%></textarea></td></tr>

  <tr><td>
  <P align=left>Almac&eacute;n: <select name="falmacen">
      <%
        if request("falmacen")="MXL" then
           response.write("<option value=2 selected>MXL</option>")
        else
           response.write("<option value=2>MXL</option>")
        end if
        
         if request("falmacen")="SJR" then
           response.write("<option value=3 selected>SJR</option>")
        else
           response.write("<option value=3>SJR</option>")
        end if
       %>

     </select>
  	<input type="submit" value="mostrar costos"></td></tr>
  </form></table>


<%
rate=""
rate=request("frate")
'response.write(rate &"<BR>")

partidas=""
partidas=trim(request("fpartidas"))
partida=""
Pos=1
longitud=0
strSQL=""

while Pos<>0
	pos=InStr(partidas,chr(13))
	
	if pos=0 then
        partida=trim(partidas) 
	else
	    partida=trim(mid(Partidas,1,Pos-1))
        Partidas=mid(Partidas,Pos+2,2000)
    end if
   
    if len(trim(partida))>1 then
       strSQL=strSQL & "'"& trim(partida) &"',"    
    end if    
    'response.write( strSQL &"<BR>")
wend


Pos=len(strSQL)
if Pos>0 then
    strSQL="select a.ItemCode,a.ItemName,CASE when b.Price=0 THEN 0 ELSE b.Price END as COSTO_USD,CASE when b.Price=0 THEN 0 ELSE round(b.Price*1.05,2) END as mas_cinco from OITM a inner join ITM1 b on b.PriceList="& trim(request("falmacen"))  &" and a.ItemCode=b.ItemCode where a.ItemCode in (" & mid(strSQL,1,Pos-1) &")"

    response.write("<table width='900px' cellspacing=1 cellpadding=1 border=1>")
    response.write("<tr><td style='width: 900px'>QUERY: " & mid(strSQL,1,249) &"</td></table><P>&nbsp;</P>")
    rsUpdateEntry.Open strSQL, adoCon2   'DMX

    response.write("<H1><B>COSTOS DE VENTA (matriz DMX) </B> </H1>")
    response.write("<table width='650px' cellspacing=1 cellpadding=1 border=1>")
    response.write("<tr><td class=subtitulo>ItemCode</td><td class=subtitulo>ItemName</td><td class=subtitulo>Costo USD</td><td class=subtitulo>Con 5% m&aacute;s</td></tr>")
    
    while not rsUpdateEntry.EOF
       response.write("<tr>")
       response.write("<td>" & rsUpdateEntry("ItemCode")  &"</td>" )
       response.write("<td>" & rsUpdateEntry("ItemName")  &"</td>" )
       response.write("<td>" & rsUpdateEntry("COSTO_USD")  &"</td>" )
       response.write("<td style='color:red'>" & rsUpdateEntry("mas_cinco")  &"</td>" )
       response.write("</tr>")
       rsUpdateEntry.MoveNext
    wend
    rsUpdateEntry.close
    response.write("</table><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>")
   

end if



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


