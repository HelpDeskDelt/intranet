<!--#include virtual="/config/conn.asp"-->

<%
Dim columnas(7)     

select case trim(request("action"))
case "1" 'suministrar'
         titulo="Pendiente por suministrar"
         doTitulo(titulo)

          strSQL="SELECT a.DocNum as Pedido,(b.LineNum+1) as Linea,b.ItemCode,b.Dscription,d.ItmsGrpNam,b.OpenQty,CONVERT(VARCHAR,CONVERT(MONEY, round(b.OpenQty*b.Price,2) ),1) as subtotal  FROM ORDR a WITH (NOLOCK) inner join RDR1 b WITH (NOLOCK) on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode left join OITB d on c.ItmsGrpCod=d.ItmsGrpCod where a.CANCELED='N' AND b.LineStatus='O' and a.CardCode='"& request("SN") &"'"

          'response.write strSQL
          rsUpdateEntry5.Open strSQL, adoCon2
          rsUpdateEntry.Open strSQL, adoCon2



          CreateTable(900) 
          Response.Write("<tr>")
          
               i=1
               For Each rsUpdateEntry in rsUpdateEntry.Fields
                      Response.Write("<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>")
                      columnas(i)=rsUpdateEntry.Name
                      i=i+1
               Next    
          Response.Write("</tr>")     

          while not rsUpdateEntry5.EOF
               Response.Write("<tr>")
               for i=1 to 7
                    Response.write("<td class='td-r fontmed'>"& rsUpdateEntry5(columnas(i)) &"</td>")              
               next     
               Response.Write("</tr>")
               rsUpdateEntry5.MoveNext
          wend
          rsUpdateEntry5.close
          closeTable
 
case "2"
         titulo="Remisionado por facturar"
         doTitulo(titulo)

         strSQL="SELECT a.DocNum as Pedido,(b.LineNum+1) as Linea,b.ItemCode,b.Dscription,d.ItmsGrpNam,b.OpenQty,CONVERT(VARCHAR,CONVERT(MONEY, round(b.OpenQty*b.Price,2) ),1) as subtotal  FROM ODLN a WITH (NOLOCK) inner join DLN1 b WITH (NOLOCK) on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode left join OITB d on c.ItmsGrpCod=d.ItmsGrpCod where a.CANCELED='N' AND b.LineStatus='O' and a.CardCode='"& request("SN") &"'"

          'response.write strSQL
          rsUpdateEntry5.Open strSQL, adoCon2
          rsUpdateEntry.Open strSQL, adoCon2



          CreateTable(900) 
          Response.Write("<tr>")
          
               i=1
               For Each rsUpdateEntry in rsUpdateEntry.Fields
                      Response.Write("<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>")
                      columnas(i)=rsUpdateEntry.Name
                      i=i+1
               Next    
          Response.Write("</tr>")     

          while not rsUpdateEntry5.EOF
               Response.Write("<tr>")
               for i=1 to 7
                    Response.write("<td class='td-r fontmed'>"& rsUpdateEntry5(columnas(i)) &"</td>")              
               next     
               Response.Write("</tr>")
               rsUpdateEntry5.MoveNext
          wend
          rsUpdateEntry5.close
          closeTable




case "3"
         titulo="Credito vencido"
         doTitulo(titulo)



end select
%>
<P align='center'><input type="button" value="cerrar" onclick="hidediv('ShowData')"></P>   