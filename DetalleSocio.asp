<!--#include virtual="/config/conn.asp"-->

<%

select case request("estatus")
case 1 Titulo="Detalle pendiente de suministrar para "& request("CardName")
case 2
case 3 
end select

response.write("<center>")

DoTitulo(titulo)

strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_suministro' AND TABLE_SCHEMA = 'dbo')     drop table   dbo._suministro;"
rsUpdateEntry7.Open strSQL, adoCon4

       
strSQL="SELECT T0.DocNum as Pedido,T1.LineNum+1 as Linea,T1.ItemCode as Codigo,T1.Dscription as Descripcion,cast(T1.OpenQty as int) as Cantidad,T1.Price as Precio_venta,round(T1.OpenQty*T1.Price,2) as subtotal,  stock=( select cast( isnull( (select SUM(Y.InQty)-SUM(Y.OutQty) from  PRODUCTIVA_DMX.dbo.OINM Y  WITH (NOLOCK) where y.ItemCode=T1.ItemCode) ,0) +  isnull( (select SUM(Y.InQty)-SUM(Y.OutQty) from  PRODUCTIVA_DFW.dbo.OINM Y  WITH (NOLOCK) where y.ItemCode=T1.ItemCode) ,0) as int ) ),isnull( (select sum(OpenQty) from PRODUCTIVA_DMX.dbo.POR1 where LineStatus='O' and ItemCode= T1.ItemCode ), 0 ) as compradas,isnull( (select sum(OpenQty) from PRODUCTIVA_DMX.dbo.POR1 where LineStatus='O' and ItemCode=T1.ItemCode and U_UBICACIONMATERIAL='EN FABRICA SIN FACTURAR USD' ), 0 ) as liberar,0 as Pdtecomprar into SBOTemp.dbo._suministro  FROM ORDR T0 WITH (NOLOCK) inner join RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry  WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' and substring(T1.ItemCode,1,2)<>'CM'  and T0.CardCode='" & request("CardCode") &"'"
'response.write strSQL

if request("RS")="DMX" then
    rsUpdateEntry6.Open strSQL, adoCon2
else
    rsUpdateEntry6.Open strSQL, adoCon3
end if

strSQL="update _suministro set Pdtecomprar=cantidad-stock-compradas where cantidad>(stock+compradas) "
rsUpdateEntry5.Open strSQL, adoCon4

strSQL="select * from _suministro order by Pedido,linea"
rsUpdateEntry.Open strSQL, adoCon4

CreateTable(980) 

response.write("<tr>")
response.write("<td class='td-c subtitulo'>Pedido</td>")
response.write("<td class='td-c subtitulo'>linea</td>")
response.write("<td class='td-c subtitulo'>Codigo</td>")
response.write("<td class='td-c subtitulo'>Descripcion</td>")
response.write("<td class='td-c subtitulo'>Cant Pdte</td>")
response.write("<td class='td-c subtitulo'>Precio venta</td>")

response.write("<td class='td-c subtitulo'>Subtotal</td>")
response.write("<td class='td-c subtitulo'>En Stock</td>")
response.write("<td class='td-c subtitulo'>Comprado</td>")
response.write("<td class='td-c subtitulo'>Pdte liberar</td>")
response.write("<td class='td-c subtitulo'>Pdte comprar</td>")

response.write("</tr>")

vSubtotal=""

i=1

while not rsUpdateEntry.EOF
   response.write("<tr>")
    if i=2 then
          vstring=" style='background-color: #E1DFDF' "
    else
         vstring=" style='background-color: #FFFFFF;' "
    end if
    response.write("<td class='td-c fonttiny' "& vstring &">"& rsUpdateEntry("Pedido") &"</td>")
    response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("Linea") &"</td>")
    response.write("<td class='td-l fonttiny' "& vstring &">"& rsUpdateEntry("Codigo") &"</td>")
    response.write("<td class='td-l fonttiny' "& vstring &">"& rsUpdateEntry("Descripcion") &"</td>")
  
    response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("cantidad") &"</td>")
   
    response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("Precio_venta") &"</td>")
    response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("subtotal") &"</td>")
    vSubtotal= vSubtotal & trim( rsUpdateEntry("subtotal") )
    
    response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("stock") &"</td>")

    response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("compradas") &"</td>")

    if int(trim( rsUpdateEntry("liberar")  ))>0 then
        response.write("<td class='td-r fonttiny' "& vstring &"><font color=red>"& rsUpdateEntry("liberar") &"</font></td>")
    else
        response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("liberar") &"</td>")
    end if

    if int(trim( rsUpdateEntry("Pdtecomprar")  ))>0 then
             response.write("<td class='td-r fonttiny' "& vstring &"><font color=red>"& rsUpdateEntry("Pdtecomprar") &"</font></td>")
    else
             response.write("<td class='td-r fonttiny' "& vstring &">"& rsUpdateEntry("Pdtecomprar") &"</td>")
    end if

   
   response.write("</tr>")
  rsUpdateEntry.moveNext
  i=i+1

  if not rsUpdateEntry.EOF then        
        vSubtotal= vSubtotal & "+"
  end if

  if i=3 then 
       i=1
  end if

wend
rsUpdateEntry.close


strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_suministro' AND TABLE_SCHEMA = 'dbo')     drop table   dbo._suministro;"
rsUpdateEntry7.Open strSQL, adoCon4



  'strSQL=" select sum(" & vSubtotal &") as calculo"
  'response.write strSQL
  'rsUpdateEntry.Open strSQL, adoCon
  'separador
  'response.write("<tr><td colspan=5 class='td-r'>Total: </td><td class='td-r'>" & vCantidad &"</td><td>&nbsp;</td><td class='td-r'>"& rsUpdateEntry("calculo") &"</td><td colspan=2>&nbsp;</td></tr>")
   'rsUpdateEntry.close

  response.write("</table><P align='center'>")
%>
<input type="button" value="cerrar" onclick="hidediv('detalle')"> </P>   