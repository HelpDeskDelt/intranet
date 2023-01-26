<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->




<%
  Response.Expires = -1

  response.write("<P>&nbsp;</P><center><table width=900px>")
  strSQL="select Rate,RateDate from ORTT where year(RateDate)=year(getdate()) and month(RateDate)=month(getdate()) and day(RateDate)=day(getdate())"
  'response.write strSQL
  rsUpdateEntry.Open strSQL, adoCon2 'DMX'

  vRate=""
  if request("frate") <>"" then
          vRate=request("frate")
  else
          vRate=trim(rsUpdateEntry("Rate"))
  end if

  response.write ("<form action='/modules/IGI_SAP.asp' method='post'>")
  response.write("<tr><td>Indique el tipo de cambio de la importaci&oacute;n: <input type='text' name='frate' id='frate' value='"& vRate &"'></td></tr>")
  response.write("<tr><td><B>Pegue el listado de sus &oacute;rdenes de compra relacionadas a la importaci&oacute;n:</B></td></tr>")
  'response.write ("<input type='hidden' name='frate' id='frate' value='"&  trim(rsUpdateEntry("Rate")) &"'>")
  rsUpdateEntry.close    %>
  
  <tr><td>
  <textarea rows="10" cols="25" id="fpartidas" name="fpartidas"><%=request("fpartidas")%></textarea></td></tr>
  <tr><td>
  
  	<input type="submit" value="retraer % IGI SAP DMX"></td></tr>
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
wend

   if trim(request("fpartidas"))<>"" then
     pos=len(strSQL)
     strSQL=mid(strSQL,1,Pos-1)
     'response.write( strSQL &"<BR>")
    end if

Pos=len(strSQL)
if Pos>0 then
    strSQL="select a.DocNum,a.CardName,b.ItemCode,c.ItemName,c.SuppCatNum,b.Quantity,b.Price,round(b.Price*"&rate &",2) as Price_MXN,d.CstGrpName from OPOR a inner join POR1 b on a.DocEntry=b.DocEntry left join OITM c on b.ItemCode=c.ItemCode left join OARG d on c.CstGrpCode=d.CstGrpCode  where a.DocNum in (" & strSQL & ") order by b.price"
    'response.write strSQL
    rsUpdateEntry.Open strSQL, adoCon2   'DMX

    'response.write("<B>QUERY:</B> " & strSQL &"<BR>")

    response.write("<H1><B>CONFIGURACION DE IGI PARA LOS SIGUIENTES CODIGOS DE DMX</B> </H1>")
    response.write("<table width='1060px' cellspacing=0 cellpadding=3 border=0>")
    response.write("<tr><td class=subtitulo>#</td><td class=subtitulo>PO</td><td class=subtitulo>Proveedor</td><td class=subtitulo>C&oacute;digo</td><td class=subtitulo>Nombre</td><td class=subtitulo># Fabricante</td><td class=subtitulo>CANT</td><td class=subtitulo>Precio USD</td><td class=subtitulo>Precio MXN</td><td class=subtitulo>IGI</td></tr>")
    
    i=1
    while not rsUpdateEntry.EOF
       response.write("<tr>")
       response.write("<td><B>" & i  &"</B></td>" )
       response.write("<td>" & rsUpdateEntry("DocNum")  &"</td>" )
       response.write("<td>" & mid(rsUpdateEntry("CardName"),1,35)  &"</td>" )
       response.write("<td>" & rsUpdateEntry("ItemCode")  &"</td>" )
       response.write("<td>" & mid(rsUpdateEntry("ItemName"),1,35)  &"</td>" )
       response.write("<td>" & rsUpdateEntry("SuppCatNum")  &"</td>" )
        response.write("<td>" & rsUpdateEntry("Quantity")  &"</td>" )
       response.write("<td>" & FormatCurrency(rsUpdateEntry("Price") )  &"</td>" )
       response.write("<td>" & FormatCurrency(rsUpdateEntry("Price_MXN") )  &"</td>" )
       response.write("<td>" & rsUpdateEntry("CstGrpName")  &"</td>" )       
       response.write("</tr>")
       rsUpdateEntry.MoveNext
       i=i+1
    wend
    rsUpdateEntry.close
    response.write("</table><P>&nbsp;</P>")

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


