<!--#include virtual="/config/conn.asp"-->


<%
if request("mode")=1 then

   strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as Subtotal,sum(subtotal) as subtotal2 from HistoricoVentas"&request("anio") &" where asesor='"& request("asesor") &"' "
   'response.write strSQL
   rsUpdateEntry3.Open strSQL, adoCon4


   strSQL="select aux.SocioNegocio,aux.Subtotal,round(CAST(aux.subtotal AS FLOAT)/CAST("& rsUpdateEntry3("subtotal2") &" AS FLOAT)*100,1)  as Pareto from ( select CardName as SocioNegocio,sum(subtotal)as Subtotal from HistoricoVentas"&request("anio")&" where asesor='"& request("asesor") &"'  group by CardName ) as aux order by Pareto desc"
   'response.write strSQL
   rsUpdateEntry.Open strSQL, adoCon4
   rsUpdateEntry2.Open strSQL, adoCon4
   
   
   response.write("<P>&nbsp;</P><H1>"& request("anio") &".- "&request("asesor") &"</H1>")



   CreateTable(800)

   CamposRS1
   ShowValoresRS2

   response.write("<tr><td class='subtitulo td-r'>Total</td><td class='td-r'>" & rsUpdateEntry3("subtotal") &"</td><td class='td-c'>&nbsp;</td></tr>" )

   closeTable

%> <input type="button" value="ocultar" onclick="hidediv('asesor')"></P>   <%

end if



if request("mode")=2 then

   response.write("<P>&nbsp;</P><H1>2018-2020.- " & request("asesor") &"</H1>")
   strSQL="select distinct(cardname) from HistoricoVentas2018 where asesor='"& request("asesor") &"' UNION select distinct(cardname) from HistoricoVentas2019 where asesor='"& request("asesor") &"' and CardName not in (select cardname from HistoricoVentas2018 where asesor='"& request("asesor") &"') UNION select distinct(cardname) from HistoricoVentas2020 where asesor='"& request("asesor") &"' and CardName not in (select cardname from HistoricoVentas2018 where asesor='"& request("asesor") &"') and CardName not in (select cardname from HistoricoVentas2019 where asesor='"& request("asesor") &"')  "

   rsUpdateEntry.Open strSQL, adoCon4
   CreateTable(700)

   i=1
   while not rsUpdateEntry.EOF               ''$ ' + CONVERT(VARCHAR,CONVERT(MONEY,MontPsto),1) as MontPsto,'
      

       strSQL="select '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,( (select cast(isnull(sum(subtotal),0) as float) from HistoricoVentas2018 where asesor='"& request("asesor") &"' and cardname='"& rsUpdateEntry("cardname") &"') + (select cast(isnull(sum(subtotal),0) as float) from HistoricoVentas2019 where asesor='"& request("asesor") &"' and cardname='"& rsUpdateEntry("cardname") &"') + (select cast(isnull(sum(subtotal),0) as float) from HistoricoVentas2020 where asesor='"& request("asesor") &"' and cardname='"& rsUpdateEntry("cardname") &"') )),1) as calculo "

       if i=1 then 
            'response.write(strSQL &"<BR>")  
       end if
        rsUpdateEntry2.Open strSQL, adoCon4

       response.write("<tr><td class='td-r fontsmall'>"& rsUpdateEntry("cardname") &"</td><td class='td-r fontsmall'>"& rsUpdateEntry2("calculo") &"</td></tr>")

       rsUpdateEntry.MoveNext
       i=i+1
       rsUpdateEntry2.close
   wend
   rsUpdateEntry.close

   closeTable
   %> <input type="button" value="ocultar" onclick="hidediv('asesor')"> <P>&nbsp;</P>  <%

end if

%>  