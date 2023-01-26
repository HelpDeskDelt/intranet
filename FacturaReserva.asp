<!--#include virtual="/config/conn-pruebas.asp"-->
<!--#include virtual="/header.asp"-->
<!--#include virtual="/modules/contenidos.asp"-->

<%
   response.write("<P align='center'>&nbsp;</P>")

     if request("msg")<>"" then  
           response.write("<P class='alert td-c'>" & request("msg") &"</P>")
   end if

   response.write("<H1 class='td-c'>REVALUACION DE INVENTARIO POR DIFERENCIA DE TIPO DE CAMBIO</H1><P>&nbsp</P>")
   response.write ("<form id='TC' action='FacturaReserva.asp' method='post'>")

   select case request("action")
   case "1"  response.write ("<input type='hidden' name='action' value='2'>")
             createtable(440)            
             response.write("<tr><td class='td-r subtitulo'># Pedimento </td><td class='td-l'><input type='text' class='anchox2 td-l' name='fpedimento' maxlength='15' placeholder='15 digitos'>")
             response.write("<tr><td class='td-c' colspan=2>&nbsp;</td></tr>")
             response.write("<tr><td class='td-c' colspan=2><input type='submit' value='continuar'></td></tr>")

             

   case "2"      
                   if len(trim(request("fpedimento")))<>15 then
                          response.redirect("FacturaReserva.asp?action=1&msg=numero de pedimento no existe, verifique")
                    end if

                   response.write ("<input type='hidden' name='action' value='3'>")             

                   strSQL="select a.docnum as 'FactReserva',a.Docdate as 'Fecha',a.DocRate as 'TipoCambio',b.LineNum+1 as Linea,b.ItemCode as 'Codigo',b.Dscription as 'Descripcion',CONVERT(VARCHAR,CONVERT(MONEY,b.Price),1)  as Precio_USD,round(a.DocRate*b.Price,2) as Costo_MXP from OPCH a inner join PCH1 b on a.DocEntry=b.DocEntry where a.Docentry in (SELECT Distinct(T1.BaseEntry)   FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry    inner join OITM A on T1.itemcode=A.itemcode left join OITL T2 on t1.docentry = T2.[ApplyEntry] and T1.[LineNum] = T2.[ApplyLine] and T2.[ApplyType] = 20 left JOIN ITL1 T3 ON T2.LogEntry = T3.LogEntry    left join OBTN T4 on T4.[ItemCode] = T3.[ItemCode] and T3.[MdAbsEntry] = t4.[absentry]    WHERE T4.[DistNumber] = '"& request("fpedimento") &"' and T0.CANCELED='N' AND T1.BaseType=18   ) order by a.Docnum,b.LineNum"

                   'response.write strSQL

                   rsUpdateEntry.Open strSQL, adoCon5 
                   rsUpdateEntry2.Open strSQL, adoCon5

                   if rsUpdateEntry.EOF then    'no entontro ninguna entrada'
                            rsUpdateEntry.close
                            rsUpdateEntry2.close
                            response.redirect("FacturaReserva.asp?action=1&msg=No existen facturas de reserva asociadas a este pedimento, revise")
                    end if

                   createtable(960)
                   response.write("<tr><td colspan=20 class='td-c ALERT'>FACTURAS DE RESERVA RELACIONADAS</TD></TR>")
                   CamposRS1
                   ShowValoresRS2
                   closetable


                   strSQL="SELECT T0.DocNum as 'Entrada',T1.BaseRef as 'F.Reserva',T1.LineNum+1 as Linea, T1.ItemCode as Codigo,T0.CardName as Proveedor,T1.Dscription as Descripcion,cast(T1.Quantity as int) as Cantidad,CONVERT(VARCHAR,CONVERT(MONEY,T1.Price),1)  as Precio,T0.DocRate as TC_Factura,(select Rate from ORTT where RateDate=T0.DocDate) as TC_Import,round(T1.Price*T0.DocRate,2) as CostoActual,round(T1.Price*(select Rate from ORTT where RateDate=T0.DocDate),2)  as CostoRevaluar FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry inner join OITM A on T1.itemcode=A.itemcode  left join OITL T2 on t1.docentry = T2.[ApplyEntry] and T1.[LineNum] = T2.[ApplyLine] and T2.[ApplyType] = 20  left JOIN ITL1 T3 ON T2.LogEntry = T3.LogEntry left join OBTN T4 on T4.[ItemCode] = T3.[ItemCode] and T3.[MdAbsEntry] = t4.[absentry]  WHERE T4.[DistNumber] = '"&request("fpedimento") &"' and T0.CANCELED='N'  order by T0.DocNum,T1.LineNum  "
                    'response.write strSQL
                    rsUpdateEntry3.Open strSQL, adoCon5 
                    rsUpdateEntry4.Open strSQL, adoCon5 

                    if rsUpdateEntry3.EOF then    'no entontro ninguna entrada'
                            rsUpdateEntry3.close
                            rsUpdateEntry4.close
                            response.redirect("FacturaReserva.asp?action=1&msg=No existen entradas con este pedimento")
                    end if

                   createtable(960)
                   response.write("<tr><td colspan=20 class='td-c ALERT'>ENTRADAS PROVENIENTES DE FACTURAS RELACIONADAS</TD></TR>")
                   CamposRS3
                   ShowValoresRS4
                   closetable

                   rsUpdateEntry5.Open strSQL, adoCon5 
                   createtable(90)
                   while not rsUpdateEntry5.EOF
                      response.write("<tr><td class='td-r'>"& rsUpdateEntry5("Codigo") &"</td></tr>")
                      rsUpdateEntry5.MoveNext
                   wend
                   rsUpdateEntry5.MoveFirst
                   closetable
                   createtable(90)
                   while not rsUpdateEntry5.EOF
                      response.write("<tr><td class='td-r'>"& rsUpdateEntry5("CostoRevaluar") &" MXP </td></tr>")
                      rsUpdateEntry5.MoveNext
                   wend
                   rsUpdateEntry5.close
                   closetable



      
   end select
   response.write("</form>")




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

