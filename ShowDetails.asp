<!--#include virtual="/config/conn.asp"-->

<%
response.write ("<BR><BR>")

if request("pedido")<>"" then
     response.write ("<H1>PARTIDAS PENDIENTES POR SUMINISTRAR PEDIDO " & request("rs") &"-" & request("pedido") &"</H1>")
end if

if request("remision")<>"" then
       response.write ("<H1>DETALLE DE LA REMISION " & request("rs") &"-" & request("remision") &"</H1>")
end if

if request("factura")<>"" then
       response.write ("<H1>DETALLE DE LA FACTURA " & request("rs") &"-" & request("factura") &"</H1>")
end if

if request("OC")<>"" then
       response.write ("<H1>DETALLE DE LA ORDEN DE COMPRA " & request("rs") &"-" & request("OC") &"</H1>")
end if
if request("PD")<>"" then
     response.write ("<H1>PARTIDAS DEL PEDIDO " & request("rs") &"-" & request("PD")   )
end if

if request("P")<>"" then
     if request("rs")="DMX" then
          response.write ("<H1>PARTIDAS DEL PEDIDO DMX-"  & request("P")  )
     else
          response.write ("<H1>PARTIDAS DEL PEDIDO DFW-"  & request("P")   )
   end if
end if

'response.write ("<P>&nbsp</P>")

  flag=0
  Cardname=0

  if request("pedido")<>"" then
       strSQL="SELECT (a.DocTotalFc-a.VatSumFc) as subtotal,b.LineNum+1 as lineNum,b.ItemCode,c.ItemName,isnull(b.Project,'SIN PROYECTO') as Proyecto,b.Quantity,b.OpenQty,d.WhsName from ORDR a inner join RDR1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode inner join OWHS d on b.WhsCode=d.WhsCode where b.OpenQty>0 and a.DocNum=" & request("pedido") &" order by LineNum"
       flag=1
  end if


  if request("remision")<>"" then
       strSQL="SELECT (a.DocTotalFc-a.VatSumFc) as subtotal,a.Cardname,b.LineNum+1 as lineNum,b.ItemCode,c.ItemName,isnull(b.Project,'SIN PROYECTO') as Proyecto,b.Quantity,b.OpenQty,d.WhsName from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode inner join OWHS d on b.WhsCode=d.WhsCode where a.DocNum=" & request("remision") &" order by LineNum"  
       flag=1
  end if

   if request("factura")<>"" then
       strSQL="SELECT (a.DocTotalFc-a.VatSumFc) as subtotal,a.Cardname,b.LineNum+1 as lineNum,b.ItemCode,c.ItemName,isnull(b.Project,'SIN PROYECTO') as Proyecto,b.Quantity,b.OpenQty,d.WhsName from OINV a inner join INV1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode inner join OWHS d on b.WhsCode=d.WhsCode where a.DocNum=" & request("factura") &" order by LineNum"
       flag=1
  end if

  if request("OC")<>"" then
       strSQL="SELECT (a.DocTotalFc-a.VatSumFc) as subtotal,b.LineNum+1 as lineNum,b.ItemCode,c.ItemName,isnull(b.Project,'SIN PROYECTO') as Proyecto,b.Quantity,b.OpenQty,d.WhsName from OPOR a inner join POR1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode inner join OWHS d on b.WhsCode=d.WhsCode where a.DocNum=" & request("OC") &" order by LineNum"       
       flag=1
  end if

if request("PD")<>"" then
       strSQL="SELECT (a.DocTotalFc-a.VatSumFc) as subtotal,a.Cardname,b.LineNum+1 as lineNum,b.ItemCode,c.ItemName,isnull(b.Project,'SIN PROYECTO') as Proyecto,b.Quantity,b.OpenQty,d.WhsName from ORDR a inner join RDR1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode inner join OWHS d on b.WhsCode=d.WhsCode where a.DocNum=" & request("PD") &" order by LineNum"      
       flag=1
       Cardname=1
  end if

if request("P")<>"" then
       strSQL="SELECT (a.DocTotalFc-a.VatSumFc) as subtotal,a.Cardname,b.LineNum+1 as lineNum,b.ItemCode,c.ItemName,isnull(b.Project,'SIN PROYECTO') as Proyecto,b.Quantity,b.OpenQty,d.WhsName from ORDR a inner join RDR1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode inner join OWHS d on b.WhsCode=d.WhsCode where a.DocNum=" & request("P") &" order by LineNum"     
       flag=1
       Cardname=1
  end if

  'response.write strSQL&"&nbsp;&nbsp;&nbsp;"& request("RS")

  select case request("RS")
     case "DMX"  rsUpdateEntry.Open strSQL,adoCon2
     case "DFW"  rsUpdateEntry.Open strSQL,adoCon3
     case "MME"  rsUpdateEntry.Open strSQL,adoCon5
  end select



  
  if not rsUpdateEntry.EOF   then
       response.write( "<BR><B>Subtotal del documento: $ " &  rsUpdateEntry("subtotal") &"</B>" ) 
  end if

  if Cardname=1 and not rsUpdateEntry.EOF   then
      response.write( "<BR>" &  rsUpdateEntry("Cardname") &"</H1>") 
  end if
   
  if flag=1 then
         
          response.write("<table width='700px' cellpadding=5 cellSpacing=1 align=center border=0")

          response.write("<tr>")
          response.write("<td width='2%' class='subtitulo'>#</td>") 
          response.write("<td width='7%' class='subtitulo' >C&oacute;digo</td>")
          response.write("<td width='20%' class='subtitulo' >Nombre del c&oacute;digo</td>")
          response.write("<td width='9%' class='subtitulo'>Proyecto</td>")
          response.write("<td width='8%' class='subtitulo'>Ordenada</td>")
          if request("pedido")<>"" or request("OC")<>"" then
             response.write("<td width='8%' class='subtitulo'>Pendiente</td>")
          end if
          response.write("<td width='14%' class='subtitulo'>Almac&eacute;n</td></tr>")
          
          while not rsUpdateEntry.EOF
             response.write("<tr>")
             response.write("<td class='td-l'>"& rsUpdateEntry("lineNum") &"</td>") 
             response.write("<td class='td-l'>"& rsUpdateEntry("ItemCode") &"</td>") 
             response.write("<td class='td-l'>"& rsUpdateEntry("ItemName") &"</td>") 
             response.write("<td class='td-c'>"& rsUpdateEntry("Proyecto") &"</td>") 
             response.write("<td class='td-r'>"& rsUpdateEntry("Quantity") &"</td>") 
             if request("pedido")<>"" or request("OC")<>""  then
                response.write("<td class='td-r'>"& rsUpdateEntry("OpenQty") &"</td>") 
             end if
             response.write("<td class='td-r'>"& rsUpdateEntry("WhsName") &"</td>") 

             response.write("</tr>")
            rsUpdateEntry.Movenext
          wend
          response.write("</table>")
          rsUpdateEntry.close

   end if
%>



<P>&nbsp</P>

<% if request("control")=2 then    %>
      <input type="button" value="cerrar" onclick="hidediv('BoxRoundedDetail2')">     <%
else   %>
      <input type="button" value="cerrar" onclick="hidediv('BoxRoundedDetail')">     <%
end if
%>

<P>&nbsp</P>


