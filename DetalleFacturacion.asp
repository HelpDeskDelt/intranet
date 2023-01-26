<!--#include virtual="/config/conn.asp"-->

<%

strSQL="set dateformat ymd;select a1.DocNum as Factura,a1.U_TIPOCOTIZACION as TIPO,.dbo.fn_GetMonthName(a1.DocDate ,'spanish') as 'Fecha', a1.cardname as SocioNegocio,a2.itemcode as codigo,a2.Quantity as cantidad,a2.Price as Precio,round(a2.Quantity*a2.Price,2) as subtotal,'factura' as CFDi,x.pedido from PRODUCTIVA_DMX.dbo.OINV a1 inner join PRODUCTIVA_DMX.dbo.INV1 a2 on a1.docentry=a2.docentry left join SBOTemp.dbo.Notificaciones x on x.tipo='factura' and a1.docnum=x.DocNum where a1.CardName not like '%DELTAFLOW%' and a1.CANCELED='N' and a2.ItemCode='"& request("codigo") &"' and a1.docdate>='"& request("fechai") &"' and a1.docdate<='"& request("fechaf") &"' UNION select a1.DocNum,a1.U_TIPOCOTIZACION,.dbo.fn_GetMonthName(a1.DocDate ,'spanish') as 'Fecha',a1.cardname,a2.itemcode,a2.Quantity,a2.Price,round(a2.Quantity*a2.Price,2) as subtotal,'factura' as CFDi,x.pedido from PRODUCTIVA_DFW.dbo.OINV a1 inner join PRODUCTIVA_DFW.dbo.INV1 a2 on a1.docentry=a2.docentry left join SBOTemp.dbo.Notificaciones x on x.tipo='factura' and a1.docnum=x.DocNum where a1.CardName not like '%DELTAFLOW%' and a1.CANCELED='N' and a2.ItemCode='"& request("codigo") &"' and a1.docdate>='"& request("fechai") &"' and a1.docdate<='"& request("fechaf") &"' UNION select a1.DocNum,a1.U_TIPOCOTIZACION,.dbo.fn_GetMonthName(a1.DocDate ,'spanish') as 'Fecha',a1.cardname,a2.itemcode,-1*a2.Quantity,a2.Price,round(-1*a2.Quantity*a2.Price,2) as subtotal,'NotaCredito' as CFDi,0 from PRODUCTIVA_DMX.dbo.ORIN a1 inner join PRODUCTIVA_DMX.dbo.RIN1 a2 on a1.docentry=a2.docentry where a1.CardName not like '%DELTAFLOW%' and a1.CANCELED='N' and a2.ItemCode='"& request("codigo") &"' and a1.docdate>='"& request("fechai") &"' and a1.docdate<='"& request("fechaf") &"' UNION select a1.DocNum,a1.U_TIPOCOTIZACION,.dbo.fn_GetMonthName(a1.DocDate ,'spanish') as 'Fecha',a1.cardname,a2.itemcode,-1*a2.Quantity,a2.Price,round(-1*a2.Quantity*a2.Price,2) as subtotal,'NotaCredito' as CFDi,0 from PRODUCTIVA_DFW.dbo.ORIN a1 inner join PRODUCTIVA_DFW.dbo.RIN1 a2 on a1.docentry=a2.docentry where a1.CardName not like '%DELTAFLOW%' and a1.CANCELED='N' and a2.ItemCode='"& request("codigo") &"' and a1.docdate>='"& request("fechai") &"' and a1.docdate<='"& request("fechaf") &"' "

'response.write strSQL
rsUpdateEntry.Open strSQL, adoCon


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

vSubtotal=""
vCantidad=0

while not rsUpdateEntry.EOF
   response.write("<tr>")
   response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("factura") &"</td>")
    response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("tipo") &"</td>")
    response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("fecha") &"</td>")
    response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("SocioNegocio") &"</td>")
    response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("codigo") &"</td>")
    response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("cantidad") &"</td>")
    vCantidad=vCantidad + int(trim(  rsUpdateEntry("cantidad") ))
    response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("Precio") &"</td>")
    response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("subtotal") &"</td>")
    vSubtotal= vSubtotal & trim( rsUpdateEntry("subtotal") )
    response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("CFDi") &"</td>")
    response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("pedido") &"</td>")
   response.write("</tr>")
  rsUpdateEntry.moveNext
  if not rsUpdateEntry.EOF then
        vSubtotal= vSubtotal & "+"
  end if
wend
  rsUpdateEntry.close

  strSQL=" select sum(" & vSubtotal &") as calculo"
  'response.write strSQL
  rsUpdateEntry.Open strSQL, adoCon
  separador
  response.write("<tr><td colspan=5 class='td-r'>Total: </td><td class='td-r'>" & vCantidad &"</td><td>&nbsp;</td><td class='td-r'>"& rsUpdateEntry("calculo") &"</td><td colspan=2>&nbsp;</td></tr>")
   rsUpdateEntry.close
  response.write("</table><P align='center'>")
%>
<input type="button" value="cerrar" onclick="hidediv('detalle')"> </P>   