<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->
<!--#include virtual="/modules/contenidos.asp"-->


<div id="contenido" class="inner-contenido">
<%
Response.Expires = -1
'======================================================================='
      contenido_selector
'======================================================================='      
%>
</div>



</body>  
  <!-- menu script itself. you should not modify this file -->
  <script language="JavaScript" src="/menuj/menu.js"></script>
  <script language="JavaScript" src="/menuj/menu_items.js"></script> 
   
 
  
<!--  footer--------------->
<script language="JavaScript">

  if ($(window).width() < 900) {

        var MENU_POS = [
      {
        
        'height': 20,  // item sizes
        'width': 90,
        'block_top': 88,    //posicion absoluta del Menu intranet
        'block_left':  wndcent({width:600,height:30}).x,    //multiplique 8 menus por width
        
        // offsets between items of the same level
        'top': 0,
        'left': 90,
        'right':90,
        // time in milliseconds before menu is hidden after cursor has gone out
        // of any items
        'hide_delay': 160,
        'expd_delay': 160,
        'css' : {
          'outer': ['m0l0oout', 'm0l0oover'],
          'inner': ['m0l0iout', 'm0l0iover']
        }
      },
      {
        'height': 20,
        'width': 90,  // es la anchura de los items desplegados en cada menu
        'block_top': 20,
        'block_left': 0,
        'top': 20,
        'left': 0,
        'css': {
          'outer' : ['m0l1oout', 'm0l1oover'],
          'inner' : ['m0l1iout', 'm0l1iover']
        }
      },
      {
        'block_top': 5,
        'block_left': 109,
        'css': {
          'outer': ['m0l2oout', 'm0l2oover'],
          'inner': ['m0l1iout', 'm0l2iover']
        }
      }
      ]

  }
  else {

        var MENU_POS = [
      {
        
        'height': 20,  // item sizes
        'width': 150,
        'block_top': 88,    //posicion absoluta del Menu intranet
        'block_left':  wndcent({width:1050,height:30}).x,    //multiplique 8 menus por width
        
        // offsets between items of the same level
        'top': 0,
        'left': 150,
        'right':150,
        // time in milliseconds before menu is hidden after cursor has gone out
        // of any items
        'hide_delay': 160,
        'expd_delay': 160,
        'css' : {
          'outer': ['m0l0oout', 'm0l0oover'],
          'inner': ['m0l0iout', 'm0l0iover']
        }
      },
      {
        'height': 20,
        'width': 150,  // es la anchura de los items desplegados en cada menu
        'block_top': 20,
        'block_left': 0,
        'top': 20,
        'left': 0,
        'css': {
          'outer' : ['m0l1oout', 'm0l1oover'],
          'inner' : ['m0l1iout', 'm0l1iover']
        }
      },
      {
        'block_top': 5,
        'block_left': 109,
        'css': {
          'outer': ['m0l2oout', 'm0l2oover'],
          'inner': ['m0l1iout', 'm0l2iover']
        }
      }
      ]

  }


  <!--//

  // some table cell or other HTML element. Always put it before </body>
  // each menu gets two parameters (see demo files)
  // 1. items structure
  // 2. geometry structure

  new menu (MENU_ITEMS, MENU_POS);
  
  //-->
</script>


</html>



<%
sub ItemsCotizadosHistoricos   'contenido 127'  
   
     doAlert
     response.write("<P>&nbsp;</P>")
     titulo="ANALISIS EN LAS COTIZACIONES DE ARTICULOS (EN HISTORICAS)"
     doTitulo(titulo)
     response.write("<P><H3>PERIODO: 14/MAYO/2013 A 30/ABR/2019  </H3></P>")

     response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=127>")
         response.write("<input type='hidden' name='action' value=2>")

         CreateTable(500)
         response.write("<tr><td class='td-r'>Familia: </td><td class='td-l'><select name='ffamilia'>")
         strSQL="select * from  DMX_OLD.dbo.OITB order by ItmsGrpNam"
         rsUpdateEntry.Open strSQL, adoCon2    'DMX'

         response.write("<option value='0'>sin seleccionar</option>")
         while not rsUpdateEntry.EOF
              if trim(request("ffamilia"))=trim(rsUpdateEntry("ItmsGrpNam")) then
                 response.write("<option value='"& rsUpdateEntry("ItmsGrpNam") &"' SELECTED>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry("ItmsGrpNam") &"'>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              end if
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close

         response.write("</select></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n incluye palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra1' value='"& request("fpalabra1") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n incluye otra palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra2' value='"& request("fpalabra2") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;on <B>NO INCLUYE </B>palabra:  </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra3' value='"& request("fpalabra3") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;on <B>NO INCLUYE </B>palabra: </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra4' value='"& request("fpalabra4") &"'></td></tr>")
         response.write("<tr><td class='td-r'>C&oacute;digo:</td><td class='td-l'><input class='anchox2' type='text' name='fcodigo' value='"& request("fcodigo") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Fecha inicio:</td><td class='td-l'><input type='date' name='ffechaI'></td></tr>")
         response.write("<tr><td class='td-r'>Fecha fin:</td><td class='td-l'><input type='date' name='ffechaF'></td></tr>")
  
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='Mostrar Analisis'></form></td></tr></table>")

  vItems=""
  if request("action")="2" then

        if request("ffamilia")="0" and len(trim(request("fpalabra1")))<4 and len(trim(request("fpalabra2")))<4 and len(trim(request("fpalabra3")))<4 and len(trim(request("fcodigo")))<4  then

               response.redirect("ShowContent.asp?contenido=75&msg=debe ser mas especifico con los criterios de busqueda")
        end if


        'CONCATENA TODOS LOS CODIGOS DE UNA FAMILIA SEPARADO POR COMAS'
        strSQL="SELECT CASE WHEN LEN(I.Items)>1 THEN SUBSTRING(I.Items,0,LEN(I.Items)) ELSE '' END as Items FROM (SELECT (SELECT STUFF(    (SELECT ''''+ ItemCode+''''+',' from DMX_OLD.dbo.OITM where 1=1 "
        if request("ffamilia")<>"0" then
             strSQL2="select ItmsGrpCod from DMX_OLD.dbo.OITB where ItmsGrpNam='"& request("ffamilia") &"' "
             rsUpdateEntry.Open strSQL2, adoCon4
             strSQL=strSQL & " and ItmsGrpCod="& rsUpdateEntry("ItmsGrpCod")             
             rsUpdateEntry.close             
        end if
        if request("fpalabra1")<>"" then
            strSQL=strSQL & " and ItemName like '%"& request("fpalabra1") &"%' "
        end if
        if request("fpalabra2")<>"" then
            strSQL=strSQL & " and ItemName like '%"& request("fpalabra2") &"%' "
        end if
        if request("fpalabra3")<>"" then
            strSQL=strSQL & " and ItemName not like '%"& request("fpalabra3") &"%' "
        end if
        if request("fpalabra4")<>"" then
            strSQL=strSQL & " and ItemName not like '%"& request("fpalabra4") &"%' "
        end if
        if request("fcodigo")<>"" then
            strSQL=strSQL & " and ItemCode='"& request("fcodigo") &"' "
        end if
        strSQL=strSQL & " for XML PATH ('')  ),1,0, '' )       ) as Items  ) as I "    
        'response.write(strSQL & "<BR><BR><BR>") 
        rsUpdateEntry.Open strSQL, adoCon2    'DMX'

        vItems=trim(rsUpdateEntry("Items"))
        rsUpdateEntry.close

        strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_ItemsCotizados' AND TABLE_SCHEMA = 'dbo')     drop table dbo._ItemsCotizados;"
        rsUpdateEntry7.Open strSQL, adoCon4  'SBO'

        strSQL="CREATE TABLE [_ItemsCotizados](    [ItemCode] [varchar](50)  NOT NULL,   [ItemName] [varchar](300) NOT NULL,   [CantPsto] [int] NULL,   [MontPsto] decimal(18,2) NULL,   [CantCompra] [int] NULL,   [MontCompra] decimal(18,2) NULL,  [CantFact] [int] NULL,   [MontFact] decimal(18,2) NULL,  [Bateo] decimal(18,2) NULL,  [Pareto] decimal(18,2) NULL  ) ON [PRIMARY]"
        rsUpdateEntry6.Open strSQL, adoCon4  'SBO'

        
        strSQL= "set dateformat ymd;INSERT INTO _ItemsCotizados (ItemCode,ItemName,CantPsto,MontPsto,CantCompra,MontCompra,CantFact,MontFact) SELECT A.ITEMCODE,A.ItemName,ISNULL(A1.CANTDMX,0)+ISNULL(A5.CANTDFW,0) AS CantPsto,ISNULL(A1.ParaPresupuesto,0)+ISNULL(A5.ParaPresupuesto,0) AS MontPsto,ISNULL(A2.CANTDMX,0)+ISNULL(A6.CANTDFW,0) AS CantCompra,ISNULL(A2.ParaCompra,0)+ISNULL(A6.ParaCompra,0) AS MontCompra, ISNULL(A3.CANTDMX,0)+ISNULL(A4.CANTDMX,0)+ISNULL(A7.CANTDFW,0)+ISNULL(A8.CANTDFW,0) AS CantFact,ISNULL(A3.Facturado,0)+ISNULL(A4.Facturado,0)+ISNULL(A7.Facturado,0)+ISNULL(A8.Facturado,0) AS MontFact FROM DMX_OLD.dbo.OITM A WITH (NOLOCK) "

        'A1 cantidades DMX Para Presupuesto
        strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDMX, ISNULL(sum(C.TOTALFRGN),0) AS ParaPresupuesto FROM DMX_OLD.dbo.OITM O WITH (NOLOCK) JOIN DMX_OLD.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DMX_OLD.dbo.OQUT C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION not LIKE('%COMPRA%') AND C.ItemCode in ("& vItems  &") GROUP BY O.ItemCode) A1 ON A.ITEMCODE = A1.ITEMCODE "

        'A5 Cantidades DFW Para Presupuesto
         strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDFW, ISNULL(sum(C.TOTALFRGN),0) AS ParaPresupuesto  FROM DFW_OLD.dbo.OITM O WITH (NOLOCK) JOIN DFW_OLD.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DFW_OLD.dbo.OQUT C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"' AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION not LIKE('%COMPRA%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A5 ON A.ITEMCODE = A5.ITEMCODE "

        'A2 cantidades DMX Para Compra
        strSQL=strSQL & "LEFT JOIN  (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDMX, ISNULL(sum(C.TOTALFRGN),0) AS ParaCompra  FROM DMX_OLD.dbo.OITM O WITH (NOLOCK) JOIN DMX_OLD.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DMX_OLD.dbo.OQUT C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION LIKE('%COMPRA%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A2 ON A.ITEMCODE = A2.ITEMCODE "

         'A6 Cantidades DFW Para Compra
         strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDFW, ISNULL(sum(C.TOTALFRGN),0) AS ParaCompra  FROM DFW_OLD.dbo.OITM O JOIN DFW_OLD.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DFW_OLD.dbo.OQUT C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"' AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION LIKE('%COMPRA%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A6 ON A.ITEMCODE = A6.ITEMCODE "


        'A3 DMX FACTURADO
        strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDMX, ISNULL(sum(C.TOTALFRGN),0) AS Facturado FROM DMX_OLD.dbo.OITM O WITH (NOLOCK) JOIN DMX_OLD.dbo.INV1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DMX_OLD.dbo.OINV C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A3 ON A.ITEMCODE = A3.ITEMCODE " 


        'A4 DMX NOTA DE CREDITO
         strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, -ISNULL(sum(C.Quantity),0) AS CantDMX, -ISNULL(sum(C.TOTALFRGN),0) AS Facturado FROM DMX_OLD.dbo.OITM O WITH (NOLOCK) JOIN DMX_OLD.dbo.RIN1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DMX_OLD.dbo.ORIN C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A4 ON A.ITEMCODE = A4.ITEMCODE "


         'A7 DFW FACTURADO
         strSQL=strSQL & "LEFT JOIN  (SELECT  O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDFW, ISNULL(sum(C.TOTALFRGN),0) AS Facturado  FROM DFW_OLD.dbo.OITM O WITH (NOLOCK) JOIN DFW_OLD.dbo.INV1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DFW_OLD.dbo.OINV C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A7 ON A.ITEMCODE = A7.ITEMCODE "

         'A8 DFW NOTA DE CREDITO
         strSQL=strSQL & "LEFT JOIN  (SELECT  O.ITEMCODE, -ISNULL(sum(C.Quantity),0) AS CantDFW, -ISNULL(sum(C.TOTALFRGN),0) AS Facturado  FROM DFW_OLD.dbo.OITM O WITH (NOLOCK)  JOIN DFW_OLD.dbo.RIN1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN DFW_OLD.dbo.ORIN C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A8 ON A.ITEMCODE = A8.ITEMCODE "
         

         strSQL=strSQL & "WHERE A.ItemCode in ("& vItems &")"
         rsUpdateEntry5.Open strSQL, adoCon4  'SBO'
         'response.write strSQL

         strSQL="delete _ItemsCotizados where CantPsto=0 and CantCompra=0"
         rsUpdateEntry4.Open strSQL, adoCon4  'SBO'


         strSQL="UPDATE _ItemsCotizados SET Bateo=CASE WHEN CantCompra>0 THEN CAST(CantFact AS FLOAT )/CAST(CantCompra AS FLOAT ) *100  ELSE 0 END, PARETO= case when (select sum(MontFact) from _ItemsCotizados) >0  then CAST(MontFact AS FLOAT ) / CAST( (select sum(MontFact) from _ItemsCotizados) AS FLOAT ) * 100 else 0 end "
         'response.write strSQL
         rsUpdateEntry3.Open strSQL, adoCon4  'SBO'

         response.write("<P>&nbsp;</P>")
         response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='1200px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

         response.write("<tr><td style='height:10px;' colspan=2>&nbsp;</td>")
         response.write("<td style='height:10px;' colspan=4 class='td-c strong'> C O T I Z A D O</td>")
         response.write("<td style='height:10px;' colspan=2 class='td-c strong'> F A C T U R A D O </td>")
         response.write("<td style='height:10px;' colspan=3>&nbsp;</td></tr>")


              

         strSQL="select a.ItemCode as codigo,a.ItemName as Descripcion ,a.CantPsto, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.MontPsto),1) as MontPsto,a.CantCompra,'$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.MontCompra),1) as MontCompra,    a.CantFact, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.MontFact),1) as MontFact, convert(varchar,a.Bateo)+ ' %' as Bateo, a.Pareto,  stock=( select cast( isnull( (select SUM(Y.InQty)-SUM(Y.OutQty) from  PRODUCTIVA_DMX.dbo.OINM Y  WITH (NOLOCK) where y.ItemCode collate SQL_Latin1_General_CP1_CI_AS =a.ItemCode collate SQL_Latin1_General_CP1_CI_AS) ,0) +  isnull( (select SUM(Y.InQty)-SUM(Y.OutQty) from  PRODUCTIVA_DFW.dbo.OINM Y  WITH (NOLOCK) where y.ItemCode collate SQL_Latin1_General_CP1_CI_AS =a.ItemCode collate SQL_Latin1_General_CP1_CI_AS ) ,0) as int ) )  from _ItemsCotizados a order by a.Pareto DESC"

         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4  'SBO'
         rsUpdateEntry2.Open strSQL, adoCon4  'SBO'

         CamposRS1
          
         while not rsUpdateEntry2.EOF

               response.write("<tr><td class='td-l fontmed'>" & mid(rsUpdateEntry2("Codigo"),1,25) & mid(rsUpdateEntry2("Codigo"),26,50) &"</td>")
               response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2("Descripcion") &"</td>")
            
               response.write("<td class='td-r fontmed'>" & replace(replace(FormatCurrency(rsUpdateEntry2("CantPsto")),"$",""),".00","") &"</td>")

               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("MontPsto") &"</td>")

               response.write("<td class='td-r fontmed' style='background-color:#E1DFDF;'>" & rsUpdateEntry2("CantCompra") &"</td>")
               response.write("<td class='td-r fontmed' style='background-color:#E1DFDF;'>" & rsUpdateEntry2("MontCompra") &"</td>")

               response.write("<td class='td-r fontmed'>" & replace(replace(FormatCurrency(rsUpdateEntry2("CantFact")),"$",""),".00","") &"</td>")

               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("MontFact") &"</td>")
               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Bateo") &"</td>")

               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Pareto") &"</td>") 
               response.write("<td class='td-r fontmed'>" & replace(replace(FormatCurrency(rsUpdateEntry2("stock")),"$",""),".00","")  &"</td>")


               '<td class="td-r"><A href="Javascript:sendReq('/modules/DetalleFacturacion.asp','codigo,fechai,fechaf','<%=rsUpdateEntry2("codigo")
               '  =request("ffechaI")    =request("ffechaF")    ','detalle')"> XXX </a></td>       
              

               
               response.write("</tr>")

               rsUpdateEntry2.movenext
        wend
        rsUpdateEntry2.close



         strSQL="select sum(CantPsto) as CantPsto, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(MontPsto)),1) as MontPsto,sum(CantCompra) as CantCompra,'$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(MontCompra)),1) as MontCompra, sum(CantFact) as CantFact, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(MontFact)),1) as MontFact, sum(Pareto) as Pareto from _ItemsCotizados"
         'response.write strSQL
         rsUpdateEntry7.Open strSQL, adoCon4  'SBO'

         response.write("<tr><td colspan=2 class='td-r subtitulo'>TOTAL</td>")
         response.write("<td class='td-r'>" & replace(replace(FormatCurrency(rsUpdateEntry7("CantPsto")),"$",""),".00","")  &"</td>")

         response.write("<td class='td-r'>" & rsUpdateEntry7("MontPsto") &"</td>")

         response.write("<td class='td-r' style='background-color:#E1DFDF;'>" & replace(replace(FormatCurrency(rsUpdateEntry7("CantCompra")),"$",""),".00","")  &"</td>")

         response.write("<td class='td-r' style='background-color:#E1DFDF;'>" & rsUpdateEntry7("MontCompra") &"</td>")

         response.write("<td class='td-r'>" & replace(replace(FormatCurrency(rsUpdateEntry7("CantFact")),"$",""),".00","")  &"</td>")
         
         response.write("<td class='td-r'>" & rsUpdateEntry7("MontFact") &"</td>")
         response.write("<td class='td-r'>&nbsp;</td>")
        
         response.write("<td class='td-r' style='background-color:#E1DFDF;'>&nbsp;</td>")
         response.write("<td class='td-r'>" & rsUpdateEntry7("Pareto") &"</td>")
         response.write("<td class='td-r'>&nbsp;</td>")
         rsUpdateEntry7.close

         response.write("</table>")   %>  <button class="exportToExcel">Exportar a excel</button> </center>

   <script>
              $(function() {
                $(".exportToExcel").click(function(e){
                  var table = $(this).prev('.table2excel');
                  if(table && table.length){
                    var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
                    $(table).table2excel({
                      exclude: ".noExl",
                      name: "Excel Document Name",
                      filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
                      fileext: ".xls",
                      exclude_img: true,
                      exclude_links: true,
                      exclude_inputs: true,
                      preserveColors: preserveColors
                    });
                  }
                });
                
              });
    </script>

   <%

          response.write("</table><P>&nbsp;</P><div id='detalle'>&nbsp;</div>")

          strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_ItemsCotizados' AND TABLE_SCHEMA = 'dbo')     drop table dbo._ItemsCotizados;"
          rsUpdateEntry7.Open strSQL, adoCon4  'SBO'

  end if



end sub


sub feeds   'contenido 118'
     doAlert
     titulo="editar feeds de inicio"
     doTitulo(titulo)
     if request("action")="" then
           
           CreateTable(900)
           response.write("<tr><td class='subtitulo td-c' width='5%'>#</td><td class='subtitulo td-c' width='70%'>feed</td><td class='subtitulo td-c' width='11%'>usuario</td><td class='subtitulo td-c' width='11%'>valido a</td></tr>")

           strSQL="select *,.dbo.fn_GetMonthName(ValidUntil,'spanish') as fecha from feeds where cast( getdate() as date)<=ValidUntil  order by ID desc"  
           'response.write strSQL
           rsUpdateEntry.Open strSQL, adoCon

           while not rsUpdateEntry.EOF
               rowin
               response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=118&action=1&ID="&rsUpdateEntry("ID")&"'><U>"& rsUpdateEntry("ID") &"</U></a></td>")
               response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("feed") &"</td>")
               response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("id_usuario") &"</td>")
               response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("fecha") &"</td>")
               RowOut
               rsUpdateEntry.MoveNext
           wend
           rsUpdateEntry.close
           closeTable

           response.write("<P class='td-c'><a href='ShowContent.asp?contenido=118&action=1'><img src='/img/indicador.jpg'> agregar un feed</a></P>")

     end if

     if request("action")="1" then   'ADD'

          response.write("<form id='feeds' method='POST' action='ShowContent.asp'>")   
          response.write("<input type='hidden' name='contenido' value=118>")
          if request("ID")<>"" then
               response.write("<input type='hidden' name='action' value=3>")
               response.write("<input type='hidden' name='ID' value="&request("ID")&">")

               strSQL="select * from feeds where ID="&request("ID")
               rsUpdateEntry.Open strSQL, adoCon

               vFeed=rsUpdateEntry("feed")
               rsUpdateEntry.close
          else
               response.write("<input type='hidden' name='action' value=2>")
          end if

          CreateTable(900)
          response.write("<tr><td class='subtitulo td-c' width='100%'>introduzca el feed, que sera visible a inicio de intranet</td></tr>")
          response.write("<tr><td class='td-c' width='100%'><textarea name='ffeed' rows=5 cols=120>"& vFeed &"</textarea></td></tr>")
          response.write("<tr><td class='td-c' width='100%'>Este feed ser&aacute; v&aacute;lido hasta: <input type='date' name='ffecha'></td></tr>")
          response.write("<tr><td class='td-c'>&nbsp;</td></tr>")
          response.write("<tr><td class='td-c'><input type='submit' value='agregar'></form></td></tr>")

     end if

      if request("action")="2" then   'ADD UP'
         if request("ffecha")="" then response.redirect("ShowContent.asp?contenido=118&msg=no proporciono una fecha de expiracion del feed")
         if len(request("ffeed"))<10 then response.redirect("ShowContent.asp?contenido=118&msg=la descripcion del feed es muy corta")

         strSQL="insert into feeds (ID,DocDate,LogDate,ValidUntil,ID_usuario,feed) values ((select max(ID)+1 from feeds),cast(getdate() as date),getdate(),'"&request("ffecha") &"','"& session("session_id")&"','"&request("ffeed")&"')"
         response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon

         response.redirect("/home.asp?msg=un feed ha sido agregado")

      end if


      if request("action")="3" then   'EDIT UP'
         if request("ffecha")="" then response.redirect("ShowContent.asp?contenido=118&msg=no proporciono una fecha de expiracion del feed")
         if len(request("ffeed"))<10 then response.redirect("ShowContent.asp?contenido=118&msg=la descripcion del feed es muy corta")

      strSQL="UPDATE feeds set feed='"&request("ffeed")&"',LogDate=getdate(),ID_usuario='"&session("session_id")&"',ValidUntil='"&request("ffecha")&"' where ID="&request("ID")

      response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon

      response.redirect("/home.asp?msg=un feed ha sido actualizado")



      end if


end sub



sub  OpenLinesPurchasing    'ventas   contenido 101'   

   strSQL="SELECT AUX1.*,ISNULL(AUX2.BACKORDER,0) BACKORDER,ISNULL(AUX3.COLOCADO,0) COLOCADO,ISNULL(AUX4.EMBAR_FACT,0) EMBAR_FACT,ISNULL(AUX5.EMBAR_SIN_FACT,0) EMBAR_SIN_FACT,ISNULL(AUX6.FRONTERA_FACT,0) FRONTERA_FACT,ISNULL(AUX7.FRONTERA_SIN_FACT,0) FRONTERA_SIN_FACT FROM    (SELECT 'DMX' as RS,T0.CardName AS 'SOCIO_NEGOCIO',sum( round(T1.OpenQty*T1.Price,2) ) as SUBTOTAL,CONVERT(VARCHAR,CONVERT(MONEY, sum(round(T1.OpenQty*T1.Price,2)) ),1) as subtotal2       FROM [PRODUCTIVA_DMX].dbo.OPOR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry    WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O'    GROUP BY T0.CardName having sum( round(T1.OpenQty*T1.Price,2) )>0 ) AS AUX1 LEFT JOIN   (SELECT A.CardName AS 'SOCIO_NEGOCIO',sum(round(B.OpenQty*B.Price,2)) AS BACKORDER    FROM [PRODUCTIVA_DMX].dbo.OPOR A WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 B WITH (NOLOCK) on A.DocEntry=B.DocEntry    WHERE A.CANCELED='N' and A.DocStatus='O' and B.LineStatus='O' AND B.U_UBICACIONMATERIAL='01'    GROUP BY A.CardName ) AS AUX2 ON AUX1.SOCIO_NEGOCIO=AUX2.SOCIO_NEGOCIO    LEFT JOIN     (SELECT C.CardName AS 'SOCIO_NEGOCIO',sum(round(D.OpenQty*D.Price,2)) AS COLOCADO   FROM [PRODUCTIVA_DMX].dbo.OPOR C WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 D WITH (NOLOCK) on C.DocEntry=D.DocEntry    WHERE C.CANCELED='N' and C.DocStatus='O' and D.LineStatus='O' AND D.U_UBICACIONMATERIAL='02'    GROUP BY C.CardName ) AS AUX3 ON AUX1.SOCIO_NEGOCIO=AUX3.SOCIO_NEGOCIO    LEFT JOIN     (SELECT F.CardName AS 'SOCIO_NEGOCIO',sum(round(G.OpenQty*G.Price,2)) AS EMBAR_FACT   FROM [PRODUCTIVA_DMX].dbo.OPOR F WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 G WITH (NOLOCK) on F.DocEntry=G.DocEntry    WHERE F.CANCELED='N' and F.DocStatus='O' and G.LineStatus='O' AND G.U_UBICACIONMATERIAL='03'    GROUP BY F.CardName ) AS AUX4 ON AUX1.SOCIO_NEGOCIO=AUX4.SOCIO_NEGOCIO    LEFT JOIN     (SELECT H.CardName AS 'SOCIO_NEGOCIO',sum(round(I.OpenQty*I.Price,2)) AS EMBAR_SIN_FACT   FROM [PRODUCTIVA_DMX].dbo.OPOR H WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 I WITH (NOLOCK) on H.DocEntry=I.DocEntry    WHERE H.CANCELED='N' and H.DocStatus='O' and I.LineStatus='O' AND I.U_UBICACIONMATERIAL='04'    GROUP BY H.CardName ) AS AUX5 ON AUX1.SOCIO_NEGOCIO=AUX5.SOCIO_NEGOCIO    LEFT JOIN     (SELECT J.CardName AS 'SOCIO_NEGOCIO',sum(round(K.OpenQty*K.Price,2)) AS FRONTERA_FACT    FROM [PRODUCTIVA_DMX].dbo.OPOR J WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 K WITH (NOLOCK) on J.DocEntry=K.DocEntry    WHERE J.CANCELED='N' and J.DocStatus='O' and K.LineStatus='O' AND K.U_UBICACIONMATERIAL='05'    GROUP BY J.CardName ) AS AUX6 ON AUX1.SOCIO_NEGOCIO=AUX6.SOCIO_NEGOCIO    LEFT JOIN     (SELECT L.CardName AS 'SOCIO_NEGOCIO',sum(round(M.OpenQty*M.Price,2)) AS FRONTERA_SIN_FACT    FROM [PRODUCTIVA_DMX].dbo.OPOR L WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 M WITH (NOLOCK) on L.DocEntry=M.DocEntry    WHERE L.CANCELED='N' and L.DocStatus='O' and M.LineStatus='O' AND M.U_UBICACIONMATERIAL='06'    GROUP BY L.CardName ) AS AUX7 ON AUX1.SOCIO_NEGOCIO=AUX7.SOCIO_NEGOCIO ORDER BY AUX1.SUBTOTAL DESC  "

  'response.write strSQL

   strSQL2="SELECT CONVERT(VARCHAR,CONVERT(MONEY, (SELECT ISNULL( SUM( round(T1.OpenQty*T1.Price,2) ),0) FROM [PRODUCTIVA_DMX].dbo.OPOR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.POR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' )    ),1)   AS CALCULO"

   'response.write strSQL

   titulo="PARTIDAS ABIERTAS ORDENES DE COMPRA"
   DoTitulo(titulo)
   
 
   response.write("<center><table border=1 cellspacing=2 cellpadding=3 align=center width='980px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

   rsUpdateEntry.Open strSQL, adoCon4
   
   rsUpdateEntry3.Open strSQL2, adoCon4

   response.write("<tr><td class='subtitulo'>RS</td>")
   response.write("<td class='subtitulo'>Socio de Negocio</td>")
   response.write("<td class='subtitulo'>Subtotal</td>")

   response.write("<td class='subtitulo'>Backorder</td>")
   response.write("<td class='subtitulo'>Colocado</td>")
   response.write("<td class='subtitulo'>Embarcado <br>Facturado</td>")
   response.write("<td class='subtitulo'>Embarcado <br>Sin Facturar</td>")
   response.write("<td class='subtitulo'>Frontera <br>Facturado</td>")
   response.write("<td class='subtitulo'>Frontera <br>Sin Facturar</td></tr>")

 
    if rsUpdateEntry.EOF then
           NoRegistros
    else         
           while not rsUpdateEntry.EOF
                response.write("<tr><td class='td-c fontmed'>"& rsUpdateEntry("RS") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("SOCIO_NEGOCIO") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("subtotal2") &"</td>")

                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("BACKORDER") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("COLOCADO") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("EMBAR_FACT") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("EMBAR_SIN_FACT") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("FRONTERA_FACT") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("FRONTERA_SIN_FACT") &"</td></tr>")
                rsUpdateEntry.MoveNext
           wend
           rsUpdateEntry.close
           response.write("<tr><td colspan=2 class='td-r subtitulo'>GRAN TOTAL</td><td class='td-r fontmed strong'>" & rsUpdateEntry3("calculo") &"</td></tr>")
           rsUpdateEntry3.close           
           %>
           </table><button class="exportToExcel">Exportar a excel</button> </center>  

           <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>
    
      <%

    end if 
    response.write("<P style='font-size:30px'>&nbsp;</P>")

 
end sub


sub VentasAnioAsesor  'contenido 102'

     titulo="Subtotales Ventas por Asesor en "&year(date())
     doTitulo(titulo)

     strSQL="select distinct(a.asesor),isnull(AUX1.Enero,0) ENE,isnull(AUX2.Febrero,0) FEB,isnull(AUX3.Marzo,0) MAR, isnull(AUX4.Abril,0) ABR,isnull(AUX5.Mayo,0) MAY,isnull(AUX6.Junio,0) JUN,isnull(AUX7.Julio,0) JUL,isnull(AUX8.Agosto,0) AGO,isnull(AUX9.Septiembre,0) SEP, isnull(AUX10.Octubre,0) OCT,isnull(AUX11.Noviembre,0) NOV,isnull(AUX12.Diciembre,0) DIC,TOTAL=(select sum(subtotal) from HistoricoVentas" & year(date()) &" where asesor=a.asesor) from HistoricoVentas" & year(date()) &" a left join  ( select asesor,sum(subtotal) as Enero from HistoricoVentas" & year(date()) &" where month(docdate)=1 group by asesor ) as AUX1 on a.asesor=AUX1.asesor left join  ( select asesor,sum(subtotal) as Febrero from HistoricoVentas" & year(date()) &" where month(docdate)=2 group by asesor ) as AUX2 on a.asesor=AUX2.asesor left join  ( select asesor,sum(subtotal) as Marzo from HistoricoVentas" & year(date()) &" where month(docdate)=3 group by asesor ) as AUX3 on a.asesor=AUX3.asesor left join  ( select asesor,sum(subtotal) as Abril from HistoricoVentas" & year(date()) &" where month(docdate)=4 group by asesor ) as AUX4 on a.asesor=AUX4.asesor left join  ( select asesor,sum(subtotal) as Mayo from HistoricoVentas" & year(date()) &" where month(docdate)=5 group by asesor ) as AUX5 on a.asesor=AUX5.asesor left join  ( select asesor,sum(subtotal) as Junio from HistoricoVentas" & year(date()) &" where month(docdate)=6 group by asesor ) as AUX6 on a.asesor=AUX6.asesor  left join  ( select asesor,sum(subtotal) as Julio from HistoricoVentas" & year(date()) &" where month(docdate)=7 group by asesor ) as AUX7 on a.asesor=AUX7.asesor  left join  ( select asesor,sum(subtotal) as Agosto from HistoricoVentas" & year(date()) &" where month(docdate)=8 group by asesor ) as AUX8 on a.asesor=AUX8.asesor left join  ( select asesor,sum(subtotal) as Septiembre from HistoricoVentas" & year(date()) &" where month(docdate)=9 group by asesor ) as AUX9 on a.asesor=AUX9.asesor  left join  ( select asesor,sum(subtotal) as Octubre from HistoricoVentas" & year(date()) &" where month(docdate)=10 group by asesor ) as AUX10 on a.asesor=AUX10.asesor  left join  ( select asesor,sum(subtotal) as Noviembre from HistoricoVentas" & year(date()) &" where month(docdate)=11 group by asesor ) as AUX11 on a.asesor=AUX11.asesor left join  ( select asesor,sum(subtotal) as Diciembre from HistoricoVentas" & year(date()) &" where month(docdate)=12 group by asesor ) as AUX12 on a.asesor=AUX12.asesor "


     response.write("<center><table border=1 cellspacing=2 cellpadding=3 align=center width='980px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

   rsUpdateEntry.Open strSQL, adoCon4   
   rsUpdateEntry2.Open strSQL, adoCon4

    Dim Campos(14)
    i=1
    rowin
     For Each rsUpdateEntry in rsUpdateEntry.Fields
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
            campos(i)=rsUpdateEntry.Name
            i=i+1
     Next 
    rowout


    while not rsUpdateEntry2.EOF
       rowin
       response.write("<td class='td-r subtitulo'>" & rsUpdateEntry2("asesor") &"</td>")
       for i=1 to 12            
            response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2( meses(i) ) &"</td>")
       next
       response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2("Total") &"</td>")
       rowout
       rsUpdateEntry2.MoveNext
    wend
    rsUpdateEntry2.close

    %>
           </table><button class="exportToExcel">Exportar a excel</button> <P>&nbsp;</P>
     <%
     vAnio=year(date())-1
     titulo="Subtotales Ventas por Asesor en "&vAnio
     doTitulo(titulo)

     strSQL="select distinct(a.asesor),isnull(AUX1.Enero,0) ENE,isnull(AUX2.Febrero,0) FEB,isnull(AUX3.Marzo,0) MAR, isnull(AUX4.Abril,0) ABR,isnull(AUX5.Mayo,0) MAY,isnull(AUX6.Junio,0) JUN,isnull(AUX7.Julio,0) JUL,isnull(AUX8.Agosto,0) AGO,isnull(AUX9.Septiembre,0) SEP, isnull(AUX10.Octubre,0) OCT,isnull(AUX11.Noviembre,0) NOV,isnull(AUX12.Diciembre,0) DIC,TOTAL=(select sum(subtotal) from HistoricoVentas" & vAnio &" where asesor=a.asesor) from HistoricoVentas" & vAnio &" a left join  ( select asesor,sum(subtotal) as Enero from HistoricoVentas" & vAnio &" where month(docdate)=1 group by asesor ) as AUX1 on a.asesor=AUX1.asesor left join  ( select asesor,sum(subtotal) as Febrero from HistoricoVentas" & vAnio &" where month(docdate)=2 group by asesor ) as AUX2 on a.asesor=AUX2.asesor left join  ( select asesor,sum(subtotal) as Marzo from HistoricoVentas" & vAnio &" where month(docdate)=3 group by asesor ) as AUX3 on a.asesor=AUX3.asesor left join  ( select asesor,sum(subtotal) as Abril from HistoricoVentas" & vAnio &" where month(docdate)=4 group by asesor ) as AUX4 on a.asesor=AUX4.asesor left join  ( select asesor,sum(subtotal) as Mayo from HistoricoVentas" &vAnio &" where month(docdate)=5 group by asesor ) as AUX5 on a.asesor=AUX5.asesor left join  ( select asesor,sum(subtotal) as Junio from HistoricoVentas" & vAnio &" where month(docdate)=6 group by asesor ) as AUX6 on a.asesor=AUX6.asesor  left join  ( select asesor,sum(subtotal) as Julio from HistoricoVentas" & vAnio &" where month(docdate)=7 group by asesor ) as AUX7 on a.asesor=AUX7.asesor  left join  ( select asesor,sum(subtotal) as Agosto from HistoricoVentas" & vAnio &" where month(docdate)=8 group by asesor ) as AUX8 on a.asesor=AUX8.asesor left join  ( select asesor,sum(subtotal) as Septiembre from HistoricoVentas" & vAnio &" where month(docdate)=9 group by asesor ) as AUX9 on a.asesor=AUX9.asesor  left join  ( select asesor,sum(subtotal) as Octubre from HistoricoVentas" & vAnio &" where month(docdate)=10 group by asesor ) as AUX10 on a.asesor=AUX10.asesor  left join  ( select asesor,sum(subtotal) as Noviembre from HistoricoVentas" & vAnio &" where month(docdate)=11 group by asesor ) as AUX11 on a.asesor=AUX11.asesor left join  ( select asesor,sum(subtotal) as Diciembre from HistoricoVentas" & vAnio &" where month(docdate)=12 group by asesor ) as AUX12 on a.asesor=AUX12.asesor "


     response.write("<center><table border=1 cellspacing=2 cellpadding=3 align=center width='980px' class='table2excel table2excel_with_colors' data-tableName='table2'>")

   rsUpdateEntry7.Open strSQL, adoCon4   
   rsUpdateEntry2.Open strSQL, adoCon4

   
    i=1
    rowin
     For Each rsUpdateEntry7 in rsUpdateEntry7.Fields
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry7.Name & "</td>"
            campos(i)=rsUpdateEntry7.Name
            i=i+1
     Next 
    rowout


    while not rsUpdateEntry2.EOF
       rowin
       response.write("<td class='td-r subtitulo'>" & rsUpdateEntry2("asesor") &"</td>")
       for i=1 to 12            
            response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2( meses(i) ) &"</td>")
       next
       response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2("Total") &"</td>")
       rowout
       rsUpdateEntry2.MoveNext
    wend
    rsUpdateEntry2.close

     %>
           </table><button class="exportToExcel">Exportar a excel</button> <P>&nbsp;</P>
           </center>  
           <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>
    
<%     


end sub


sub DIOT   'contenido 103'
         titulo="Genere informe DIOT"
         doTitulo(titulo)

         strSQL="SELECT name  FROM sys.databases where name like '%PRODUCTIVA%' and name not like '%test%' and name not like '%prueba%'"
         rsUpdateEntry.Open strSQL, adoCon4     



    if request("action")="" then

         response.write("<form id='DIOT' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=103>")
         response.write("<input type='hidden' name='action' value=1>")


         CreateTable(350)
         rowin
         response.write("<td class='td-r subtitulo'>Base de Datos:</td><td class='td-c'><select name='fBD'>")         
             while not rsUpdateEntry.EOF
                 response.write("<option value='"& rsUpdateEntry("name") &"'>"& rsUpdateEntry("name") &"</option>")
                 rsUpdateEntry.MoveNext
             wend
             rsUpdateEntry.close
         response.write("</select></td></tr>")
         RowIn
         response.write("<td class='td-r subtitulo'>Fecha inicial:</td><td class='td-c'><input type='date' name='ffechaini'></td></tr>")      
         RowIn
         response.write("<td class='td-r subtitulo'>Fecha final:</td><td class='td-c'><input type='date' name='ffechafin'></td></tr>") 
         response.write("<td class='td-c' colspan=2>&nbsp;</td></tr>") 
         response.write("<td class='td-c' colspan=2><input type='submit' value='generar DIOT'></form></td></tr>") 
         response.write("<td class='td-c' colspan=2>&nbsp;</td></tr>") 
         response.write("<td class='td-c' colspan=2>&nbsp;</td></tr>") 
         response.write("<td class='td-c' colspan=2>fecha final debe ser mayor fecha inicial</td></tr>") 
         closeTable

    else
         vfecha1=date()
         vfecha2=date()
         vfecha1=cdate(request("ffechaini"))
         vfecha2=cdate(request("ffechafin"))
         'response.write ( vfecha1 &"<BR>")
         'response.write ( vfecha2 &"<BR>")
         if vfecha2>vfecha1 then
               '----debe Ejecutar DIOT y generar Layout'
               strSQL="EXEC XAM_FISCAL_24 {d '"& request("ffechaini") &"'}, {d '"& request("ffechafin") &"'}, 'N', 'N' , 'Y', 'Y', 'P', 'N'"

               strSQL2="SELECT DISTINCT C01_TIPO_TERCERO,C02_TIPO_OPERACION,CASE WHEN C01_TIPO_TERCERO = '05' THEN '' ELSE C03_RFC  END AS C03_RFC  ,  C04_ID_FISCAL,C05_NOMBRE_EXTRANJERO,C06_PAIS_RESIDENCIA,C07_NACIONALIDAD,       CASE WHEN C08_VALOR_ACTOS_IVA16 = 0 THEN '' ELSE CAST(C08_VALOR_ACTOS_IVA16 as nvarchar(12)) END as C08_VALOR_ACTOS_IVA16,        '' as C09_IVATASA15_DEROGADO, CASE WHEN C10_MONTO_IVA_NO_ACREDITABLE = 0 THEN '' ELSE CAST(C10_MONTO_IVA_NO_ACREDITABLE as nvarchar(12)) END as C10_MONTO_IVA_NO_ACREDITABLE,'' as C11_IVA_PAGADO_11_DEROGADO,        '' as C12_IVA_PAGADO_10_DEROGADO,   CASE WHEN C13_MONTOACTOS_IVA_ESTIMULO_FRONTERA = 0 THEN '' ELSE CAST(C13_MONTOACTOS_IVA_ESTIMULO_FRONTERA as nvarchar(12)) END as C13_MONTOACTOS_IVA_ESTIMULO_FRONTERA,   CASE WHEN C14_MONTOIVAPAGADO11_NOACRED = 0 THEN '' ELSE CAST(C14_MONTOIVAPAGADO11_NOACRED as nvarchar(12)) END as C14_MONTOIVAPAGADO11_NOACRED,         CASE WHEN C15_MONTOIVAPAGADO_FRONTERA_NOACRED = 0 THEN '' ELSE CAST(C15_MONTOIVAPAGADO_FRONTERA_NOACRED as nvarchar(12)) END as C15_MONTOIVAPAGADO_FRONTERA_NOACRED,        CASE WHEN C16_ACT0S_GRAVADOS_IMPORTACION = 0 THEN '' ELSE CAST(C16_ACT0S_GRAVADOS_IMPORTACION as nvarchar(12)) END as C16_ACT0S_GRAVADOS_IMPORTACION,         CASE WHEN C17_IVA_IMPORTACION_NOACRED = 0 THEN '' ELSE CAST(C17_IVA_IMPORTACION_NOACRED as nvarchar(12)) END as C17_IVA_IMPORTACION_NOACRED,        '' as C18_IVA_IMPORTACION_11_DEROGADO,        '' as C19_IVA_NOAC_IMPORTACION_11_DEROGADO,         CASE WHEN C20_VALOR_ACTOS_IMPORT_EXENTO = 0 THEN '' ELSE CAST(C20_VALOR_ACTOS_IMPORT_EXENTO as nvarchar(12)) END C20_VALOR_ACTOS_IMPORT_EXENTO,         CASE WHEN C21_VALOR_ACTOS_IVA_0 = 0 THEN '' ELSE CAST(C21_VALOR_ACTOS_IVA_0 as nvarchar(12)) END as C21_VALOR_ACTOS_IVA_0,        CASE WHEN C22_VALOR_ACTOS_EXENTOS_IVA = 0 THEN '' ELSE CAST(C22_VALOR_ACTOS_EXENTOS_IVA as nvarchar(12)) END C22_VALOR_ACTOS_EXENTOS_IVA,         CASE WHEN C23_IVA_RETENIDO = 0 THEN '' ELSE CAST(C23_IVA_RETENIDO as nvarchar(12)) END as C23_IVA_RETENIDO,         CASE WHEN C24_IVA_BONIF_DEVOL = 0 THEN '' ELSE CAST(C24_IVA_BONIF_DEVOL as nvarchar(12)) END as C24_IVA_BONIF_DEVOL     FROM T_XAM_DIOT_REP;   "
               'response.write ( strSQL &"<BR>")

               select case request("fBD")
               case "PRODUCTIVA_DMX"
                     rsUpdateEntry2.Open strSQL, adoCon2 
                     rsUpdateEntry3.Open strSQL2, adoCon2 
                     rsUpdateEntry4.Open strSQL2, adoCon2 
               case "PRODUCTIVA_DFW"
                     rsUpdateEntry2.Open strSQL, adoCon3   
                     rsUpdateEntry3.Open strSQL2, adoCon3 
                     rsUpdateEntry4.Open strSQL2, adoCon3
               case "PRODUCTIVA_MEIDE"
                     rsUpdateEntry2.Open strSQL, adoCon5 
                     rsUpdateEntry3.Open strSQL2, adoCon5 
                     rsUpdateEntry4.Open strSQL2, adoCon5
               case else
               end select

                response.write("<table border=1 cellspacing=2 cellpadding=3 align='left' class='dataTable'>")  
                Dim Campos(24)
                i=1
                For Each rsUpdateEntry3 in rsUpdateEntry3.Fields
                        Response.Write "<th class='subtitulo td-c fonttiny'>" & rsUpdateEntry3.Name & "</th>"
                        campos(i)=rsUpdateEntry3.Name
                        i=i+1
                Next 
                while not rsUpdateEntry4.EOF
                    rowin
                    for i=1 to 24
                       response.write("<td class='td-r fonttiny'>" & rsUpdateEntry4(campos(i)) &"</td>")
                    next
                    rowout
                    rsUpdateEntry4.MoveNext
                wend
                rsUpdateEntry4.close
                 %>
         </table> <button id="export">genera TXT</button>                                                 &nbsp;&nbsp;&nbsp;   

         
    <script>
          $('#export').click(function() {
          var titles = [];
          var data = [];

          /*
           * Get the table headers, this will be CSV headers
           * The count of headers will be CSV string separator
           */

          $('.dataTable th').each(function() {
            titles.push($(this).text());
          });

          /*
           * Get the actual data, this will contain all the data, in 1 array
           */
          $('.dataTable td').each(function() {
            data.push($(this).text());
          });
          
          /*
           * Convert our data to CSV string
           */
          //var CSVString = prepCSVRow(titles, titles.length, '');
          //CSVString = prepCSVRow(data, titles.length, CSVString);

           var CSVString = prepCSVRow(data, titles.length, '');

          /*
           * Make CSV downloadable
           */
          var downloadLink = document.createElement("a");
          var blob = new Blob(["\ufeff", CSVString]);
          var url = URL.createObjectURL(blob);
          downloadLink.href = url;
          downloadLink.download = "data.txt";

          /*
           * Actually download CSV
           */
          document.body.appendChild(downloadLink);
          downloadLink.click();
          document.body.removeChild(downloadLink);
        });

           /*
        * Convert data array to CSV string
        * @param arr {Array} - the actual data
        * @param columnCount {Number} - the amount to split the data into columns
        * @param initial {String} - initial string to append to CSV string
        * return {String} - ready CSV string
        */
        function prepCSVRow(arr, columnCount, initial) {
          var row = ''; // this will hold data
          var delimeter = '|'; // data slice separator, in excel it's `;`, in usual CSv it's `,`
          var newLine = '|\r\n'; // newline separator for CSV row

          /*
           * Convert [1,2,3,4] into [[1,2], [3,4]] while count is 2
           * @param _arr {Array} - the actual array to split
           * @param _count {Number} - the amount to split
           * return {Array} - splitted array
           */
          function splitArray(_arr, _count) {
            var splitted = [];
            var result = [];
            _arr.forEach(function(item, idx) {
              if ((idx + 1) % _count === 0) {
                splitted.push(item);
                result.push(splitted);
                splitted = [];
              } else {
                splitted.push(item);
              }
            });
            return result;
          }
          var plainArr = splitArray(arr, columnCount);
          // don't know how to explain this
          // you just have to like follow the code
          // and you understand, it's pretty simple
          // it converts `['a', 'b', 'c']` to `a,b,c` string
          plainArr.forEach(function(arrItem) {
            arrItem.forEach(function(item, idx) {
              row += item + ((idx + 1) === arrItem.length ? '' : delimeter);
            });
            row += newLine;
          });
          return initial + row;
        }
    </script>

    <P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>

    <%
         else
        
               response.write("<P class='td-c alert'>Fechas inconsistentes, informe no generado</P>")

         end if

    end if

end sub


sub PAGOS   'CONTENIDO 104'
     doAlert
     titulo="REGISTRAR UN PAGO REALIZADO A FABRICANTE"
     DoTitulo(titulo)

     response.write("<form id='' method='POST' action='ShowContent.asp'>")
     response.write("<input type='hidden' name='contenido' value='104'>")

     select case request("action")
     case "1"       
       response.write("<input type='hidden' name='action' value='2'>")
       CreateTable(350)
       RowIn
       response.write("<td width='60%' class='subtitulo td-r'>Indique raz&oacute;n social:</td>")
       response.write("<td width='40%' class='td-l'><select name='RS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option><option value='MME'>MME</option></select></td>")
       RowOut
       RowIn
       response.write("<td class='subtitulo td-r'>Indique la moneda:</td>")
       response.write("<td class='td-l'><select name='moneda'><option value='USD'>USD</option><option value='MXN'>MXN</option></select></td>")
       RowOut

       RowIn
       response.write("<td class='subtitulo td-r'>Indique naturaleza:</td>")
       response.write("<td class='td-l'><select name='tipo'><option value='EXTRANJERO'>EXTRANJERO</option><option value='NACIONAL'>NACIONAL</option></select></td>")
       response.write("<tr><td colspan=2>&nbsp;</td></tr>")       
       response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='continuar'></form</td></tr>")
       response.write("<tr><td colspan=2>&nbsp;</td></tr>")
       closeTable
       response.write("<P>&nbsp;</P>")
       ShowPagos

      case "2" 
       response.write("<input type='hidden' name='action' value='3'>")
       response.write("<input type='hidden' name='RS' value='"& request("RS") &"'>")
       response.write("<input type='hidden' name='Moneda' value='"& request("moneda") &"'>")
       response.write("<input type='hidden' name='tipo' value='"& request("tipo") &"'>")       

       CreateTable(450)
       RowIn
       response.write("<td width='60%' class='subtitulo td-r'>Seleccione Fabricante:</td>")
       response.write("<td width='40%' class='td-l'><select name='NombreCorto'>")

       strSQL="select * from OCRD where substring(cardcode,1,1)='P' order by cardCode"
       select case request("RS")
       case "DMX"   rsUpdateEntry.Open strSQL, adoCon2
       case "DFW"   rsUpdateEntry.Open strSQL, adoCon3
       case "MME"   rsUpdateEntry.Open strSQL, adoCon5
       end select
       
       while not rsUpdateEntry.EOF
           response.write("<option value='"& rsUpdateEntry("U_NombreCorto") &"'>"& rsUpdateEntry("U_NombreCorto") &"(" &  rsUpdateEntry("cardcode")  &")</option>")
           rsUpdateEntry.movenext
       wend
       rsUpdateEntry.close
       response.write("</select></td></tr>")
       RowIn
       response.write("<td class='subtitulo td-r'>Indique el monto:</td>")
       response.write("<td class='td-l'><input type='text' name='monto' class='td-r' placeholder='no utilice comas'> "& request("moneda") &"</td>")
       RowIn
       response.write("<td class='subtitulo td-r'>Fecha del pago:</td>")
       response.write("<td class='td-l'><input type='date' name='fecha'></td>")        
       RowIn
       response.write("<td class='subtitulo td-r'>Es para pago en <br> compras de contado:</td>")
       response.write("<td class='td-l'><input type='checkbox' name='contado'></td>")    
       RowOut 
       separador
       
       response.write("<tr><td colspan=2>&nbsp;</td></tr>") 
       response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='continuar, subir comprobante'></form</td></tr>")
       response.write("<tr><td colspan=2>&nbsp;</td></tr>")
       
       closeTable


      case "3" 
      if request("monto")<>"" and request("fecha")<>"" then
           CardCode=""
           strSQL="select * from OCRD where substring(cardcode,1,1)='P' and U_NombreCorto='"& request("NombreCorto") &"' "
           select case request("RS")
           case "DMX"   rsUpdateEntry.Open strSQL, adoCon2
           case "DFW"   rsUpdateEntry.Open strSQL, adoCon3
           case "MME"   rsUpdateEntry.Open strSQL, adoCon5
           end select

           if not rsUpdateEntry.EOF then
                  CardCode=rsUpdateEntry("CardCode")
           end if
           rsUpdateEntry.close

          strSQL="select max(ID)+1 as calculo from PAGOS"
          rsUpdateEntry.Open strSQL, adoCon
          vID=rsUpdateEntry("calculo")
          rsUpdateEntry.close

          strSQL="insert into PAGOS (ID,modulo,RS,moneda,tipo,NombreCorto,CardCode,Monto,DocDate,contado,LOGDATE,ID_USUARIO) values ((select max(ID)+1 as calculo from PAGOS),'proveedor','"& request("RS") &"','"& request("moneda") &"','"& request("tipo") &"','"& request("NombreCorto") &"','"& CardCode &"',"& request("monto") &",'"& request("Fecha") 

           if request("contado")="on" or request("contado")="true"  then
              strSQL= strSQL &"',1,getdate(),'"& Session("session_id")  &"')"
           else
               strSQL= strSQL &"',0,getdate(),'"& Session("session_id") &"')"
           end if

          response.write(strSQL&"<br>")
          rsUpdateEntry.Open strSQL, adoCon
          response.redirect("http://192.168.0.206/repositorio/PAYMENTS/upload/uploadform-Payments.asp?ID=" & vID)


      else
          response.redirect("ShowContent.asp?contenido=104&action=1&msg=envio datos incompletos")
      end if

     case "4" 
          response.write("<a href='ShowContent.asp?contenido=104&action=5&ID="& request("ID") &"'>Confirme desea borrar PAYMENT no. "& request("ID") &"</a>")
     case "5" 
          response.write "va a borrar payment no. " & request("ID") 
          strSQL= "DELETE PAGOS where ID="& request("ID") 
          rsUpdateEntry.Open strSQL, adoCon
          response.redirect("ShowContent.asp?contenido=104&action=1&msg=se elimino payment no."& request("ID")  )

     end select

end sub


sub ShowPagos

      titulo="Ultimos 10 pagos realizados"
      doTitulo(Titulo)
    
      strSQL="select TOP 10 ID,.dbo.fn_GetMonthNameUS(DocDate,'english') as Fecha,RS,Tipo,NombreCorto,CONVERT(VARCHAR,CONVERT(MONEY,Monto),1) as Monto,ID_USUARIO,nombre_archivo,year(DocDate) as ANIO,month(docdate) as MES from PAGOS order by ID desc "   

      Const adCmdText = &H0001
      Const adOpenStatic = 3

      rsUpdateEntry5.Open strSQL, adoCon
      rsUpdateEntry6.Open strSQL, adoCon

      Dim Campos(20)
      CreateTable(900)

         i=1
         For Each rsUpdateEntry5 in rsUpdateEntry5.Fields                        
            campos(i)=rsUpdateEntry5.Name
            i=i+1           
         Next     
         
         RowIn
         for i=1 to 8           
            Response.Write "<th class='td-c'>" & campos(i) & "</th>"               
         next      
         Response.Write "<th class='td-c'>&nbsp;</th>"            
         RowOut
         vMes=""

         while not rsUpdateEntry6.EOF   
             rowin           
             for i=1 to 7
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry6(campos(i)) &"</td>")
             next     
             if int(trim(rsUpdateEntry6("MES")))<10 then 
                 vMes="0"&trim(rsUpdateEntry6("MES")) 
             else vMes=trim(rsUpdateEntry6("MES")) end if     
             response.write("<td class='td-r fonttiny'><a href='http://192.168.0.206/repositorio/PAYMENTS/"&rsUpdateEntry6("ANIO")&"/"&vMes&"/"& rsUpdateEntry6(campos(8))&"' target='pagos'>"& rsUpdateEntry6(campos(8)) &"</a></td>")
             Response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=104&action=4&ID=" & rsUpdateEntry6("ID") &"'><img src='/img/b_drop.png' with='11px' alt='borrar' title='borrar' height='11px' border='0'></a></td>")
             rsUpdateEntry6.MoveNext                         
             RowOut
         wend
         rsUpdateEntry6.close
         closeTable

end sub



sub informeCombustibles  'contenido 105'
    doAlert
    titulo="INFORME DE CONSUMO COMBUSTIBLE <BR><H3>(basado en captura de vouchers)</H3>"
    if request("fVendor")<>"" then
        if int(trim(request("fVendor")))>0 then 
              if int(trim(request("fVendor")))=1 then titulo="INFORME DE CONSUMO COMBUSTIBLE <BR><H3>(basado en captura de vouchers ENERCARD)</H3>"
              if int(trim(request("fVendor")))=2 then titulo="INFORME DE CONSUMO COMBUSTIBLE <BR><H3>(basado en captura de vouchers EFECTIVALE)</H3>"
        end if
    else
        titulo="INFORME DE CONSUMO COMBUSTIBLE <BR><H3>(basado en captura de vouchers)</H3>"
    end if
    
    doTitulo(titulo)

    'response.write("<P>&nbsp;</P>")
    response.write("<form id='consumo' method='POST' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value='105'>")
    response.write("<input type='hidden' name='action' value='1'>")

    if request("action")="" then
      CreateTable(360)
      rowin
      response.write("<th class='td-r'>Fecha inicial:</th><td class='td-l fontmed'><input type='date' name='ffechaini'></td></tr>")
      rowin
      response.write("<th class='td-r'>Fecha final:</th><td class='td-l fontmed'><input type='date' name='ffechafin'></td></tr>")
      rowin
      rowin
      response.write("<th class='td-r'>Proveedor:</th><td class='td-l fontmed'><select name='fVendor'><option value='0'>todos</option><option value='1'>enercard</option><option value='2'>efectivale</option></select></td></tr>")
      response.write("<td colspan=2>&nbsp;</td></tr>")
      rowin
      response.write("<td colspan=2 class='td-c'><input type='submit' value='continuar'></form></td></tr>")
      closeTable
    end if

    if request("action")="1" then         
         strSQL="select case when cast('"& request("ffechafin") &"' as date)>cast('"& request("ffechaini") &"' as date) then 1 else 0 end as calculo "         
         rsUpdateEntry.Open strSQL, adoCon
         vCalculo=rsUpdateEntry("calculo")
         rsUpdateEntry.close

         if vCalculo=1 then   'valid'
                strSQL="select a.TarjetaCombustible as Tarjeta,a.Estacion,a.ID_viatico as Viatico,case when substring(a.Motivo,1,16)='Combustible por ' then substring(a.Motivo,17,1000) else a.Motivo end as Motivo,.dbo.fn_GetMonthName(a.DocDate,'spanish') as Fecha,a.ID_unidad as Unidad,a.Chofer,a.ID_USUARIO,a.Litros,c.Combustible,b.Motivo as Asunto_Viatico from VoucherGas a left join VIATICOS_R1 b on a.ID_viatico=b.ID left join Flotilla c on a.id_unidad=c.id_unidad where a.DocDate>='"& request("ffechaini") &"' and a.DocDate<='"& request("ffechafin") &"' "

                if int(trim(request("fVendor")))>0 then
                    if int(trim(request("fVendor")))=1 then strSQL=strSQL&"and len(a.TarjetaCombustible)>7 "
                    if int(trim(request("fVendor")))=2 then strSQL=strSQL&"and len(a.TarjetaCombustible)=7 "
                end if

                strSQL=strSQL&" order by a.ID_unidad,a.DocDate,a.ID_viatico,a.LogDate"
                'response.write strSQL
                rsUpdateEntry.Open strSQL, adoCon
                rsUpdateEntry2.Open strSQL, adoCon

                response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='1140px' class='table2excel table2excel_with_colors' data-tableName='table1'>")  

                Dim Campos(11) 
         i=1
         
         For Each rsUpdateEntry in rsUpdateEntry.Fields                        
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next     
         
         RowIn
         for i=1 to 4           
            Response.Write "<th class='td-c'>" & campos(i) & "</th>"               
         next         
         response.write("<th class='td-c' width='90px'>" & campos(5) & "</th>")
         for i=6 to 11           
            Response.Write "<th class='td-c'>" & campos(i) & "</th>"               
         next
         RowOut

         vsuma="select "
         i=1
         while not rsUpdateEntry2.EOF 
             vUnidad=rsUpdateEntry2("unidad")           
             if i=1 then 
                 strSQL="select * from flotilla where id_unidad='"& rsUpdateEntry2("unidad") &"' "                
                 rsUpdateEntry3.Open strSQL, adoCon
                 response.write("<tr><td colspan=11 class='td-l strong' style='padding-left:80px'>"& rsUpdateEntry2("unidad") &" | " & rsUpdateEntry3("descripcion") &"</td></tr>")
                 rsUpdateEntry3.close
             end if
             
             rowin
             for i=1 to 11
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
             next    
             vsuma=vsuma & "+" & trim(rsUpdateEntry2("litros"))  
             vcombustible=rsUpdateEntry2("combustible") 
             rsUpdateEntry2.MoveNext                         
             RowOut
   
             if not rsUpdateEntry2.EOF then
                 if vUnidad<>rsUpdateEntry2("unidad") then
                     separador
                     vsuma=vsuma&" as calculo "
                     'response.write(vsuma&"<br>")
                     rsUpdateEntry3.Open vsuma, adoCon                    
                     response.write("<tr><td colspan=8 class='td-r strong'>TOTAL:</td><td colspan=3 class='td-l strong'>"& rsUpdateEntry3("calculo") &" Litros de "& rsUpdateEntry2("combustible") &"</td></tr>") 
                     rsUpdateEntry3.close               
                     strSQL="select * from flotilla where id_unidad='"& rsUpdateEntry2("unidad") &"' "                
                     rsUpdateEntry4.Open strSQL, adoCon
                     response.write("<tr><td colspan=11 class='td-l strong' style='padding-left:80px'>"& rsUpdateEntry2("unidad") &" | " & rsUpdateEntry4("descripcion") &"</td></tr>")
                     rsUpdateEntry4.close
                     vsuma="select "
                 end if
             else
                     separador
                     vsuma=vsuma&" as calculo "
                     'response.write(vsuma&"<br>")
                     rsUpdateEntry3.Open vsuma, adoCon                               
                     response.write("<tr><td colspan=8 class='td-r strong'>TOTAL:</td><td colspan=3 class='td-l strong'>"& rsUpdateEntry3("calculo") &" Litros de "& vcombustible &"</td></tr>")                     
                     rsUpdateEntry3.close      

             end if
         wend
         
         rsUpdateEntry2.close
         %>
         </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp;

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <%

         else
                response.redirect("ShowContent.asp?contenido=105&msg=fecha final debe ser superior a fecha inicial")
         end if

    end if

end sub


sub Resguardos  'contenido 106'
     doAlert
     response.write("<P>&nbsp;</P>")
     if request("action")="" then 
        titulo="REGISTRE RESGUARDOS DE MERCANCIAS"
        DoTitulo(titulo)
     end if

   

    if request("action")="" then
      response.write("<form id='resguardo' method='POST' action='ShowContent.asp'>")
      response.write("<input type='hidden' name='contenido' value='106'>")
      response.write("<input type='hidden' name='action' value='1'>")
      CreateTable(700)

      rowin
      response.write("<th class='td-r' width='35%'>Raz&oacute;n Social:</th><td class='td-l fontmed' width='15%'><select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option><option value='MME'>MME</option></select></td>")     

      response.write("<th class='td-r' width='30%'>Tipo de documento SAP:</th><td class='td-l fontmed' width='20%'><select name='ftipo'><option value='factura'>facturas</option><option value='remision'>remisiones</option></select></td>")
      RowOut
      RowIn
         response.write("<th class='td-r'># documentos SAP del embarque (mismo cliente):</th><td class='td-l fontmed'><select name='fcantidad'>")
          for i=1 to 30 
              response.write("<option value='"&i&"'>"&i&"</option>")
          next
          response.write("</select></td>")      
     
         response.write("<th class='td-r'>Fecha desde que se resguarda:</th><td class='td-l fontmed'><input type='date' name='ffechai'></td></tr>")
      RowOut
       
      rowin
      response.write("<td colspan=4 class='td-c'><input type='submit' value='continuar'></form></td></tr>")
      closeTable
    end if



    if request("action")="1" then
      if request("ffechai")="" then response.redirect("ShowContent.asp?contenido=106&msg=la fecha no puede ir en blanco")
      response.write("<form id='resguardo' method='POST' action='ShowContent.asp'>")
      response.write("<input type='hidden' name='contenido' value='106'>")
      response.write("<input type='hidden' name='action' value='2'>")

    
      response.write("<input type='hidden' name='RS' value='"& request("fRS") &"'>")
      response.write("<input type='hidden' name='cantidad' value='"& request("fcantidad") &"'>")      
      response.write("<input type='hidden' name='fechai' value='"& request("ffechai") &"'>")  
      response.write("<input type='hidden' name='ftipo' value='"& request("ftipo") &"'>")  
      CreateTable(840)
               
      rowin
      response.write("<th class='td-r'>Socio de negocio:</th><td class='td-l fontmed'><select name='fSN'>")

      strSQL="select * from repositorio where RazonSocial='"&request("fRS") &"' and substring(cardcode,1,1)='C'  order by cardcode desc"
      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon4

      while not rsUpdateEntry.EOF
        response.write("<option value='"& rsUpdateEntry("cardCode") &"'>"& rsUpdateEntry("Ruta") &"</option>")
        rsUpdateEntry.MoveNext
      wend
      rsUpdateEntry.close
      response.write("</select></td></tr>")  

      for i=1 to int(trim(request("fcantidad")))
         rowin
         response.write("<th class='td-r'>Do0cumento SAP No."& i&":</th><td class='td-l fontmed'><input type='numeric' name='factura"&i&"'></td></tr>")       
      next

      rowin
        response.write("<th class='td-r'>Instrucciones: </th><td class='td-l fontmed'> <textarea rows='4' cols='80' name='finstruccion'></textarea></td></tr>")
        response.write("<th class='td-r'>Pr&oacute;xima fecha de revisi&oacute;n:</th><td class='td-l fontmed'> <input type='date' name='ffechaprox'></td></tr>")


      rowin
      response.write("<td colspan=2 class='td-c'>&nbsp;</td></tr>")      
      rowin
      response.write("<td colspan=2 class='td-c'><input type='submit' value='registrar'></form></td></tr>")
      closeTable

    end if


    if request("action")="2" then

          strSQL=request("factura1")
          for i=2 to int(trim(request("cantidad")))
               strSQL= strSQL & ","
               vstring="factura"&i
               strSQL= strSQL & request(vstring)               
          next
          facturas=strSQL

          strSQL="select count(*) as calculo from Notificaciones where RazonSocial='"&request("RS")&"' and tipo='"&request("ftipo")&"' and DocNum in ("& facturas &")"
          response.write(strSQL&"<br>")
          rsUpdateEntry.Open strSQL, adoCon4
          vstring=trim(rsUpdateEntry("calculo"))
          rsUpdateEntry.close

          if int(vstring)<>int(trim(request("cantidad"))) then   response.redirect("ShowContent.asp?contenido=106&msg=uno o mas numeros de documentos SAP no existen para la Razon Social especificada")

           strSQL="select count(*) as calculo from Notificaciones where RazonSocial='"&request("RS")&"' and tipo='"&request("ftipo")&"' and DocNum in ("& facturas &") and cardCode='"& request("fSN") &"' "
          response.write(strSQL&"<br>")
          rsUpdateEntry.Open strSQL, adoCon4
          vstring=trim(rsUpdateEntry("calculo"))
          rsUpdateEntry.close

          if int(vstring)<>int(trim(request("cantidad"))) then  response.redirect("ShowContent.asp?contenido=106&msg=uno o mas numeros de documentos de SAP no corresponden al Socio de Negocio no. " & request("fSN") )

          if trim(request("ffechaprox"))="" then response.redirect("ShowContent.asp?contenido=106&msg=la fecha proxima de revision no puede ir vacia")

          strSQL="select case when cast('"&trim(request("ffechaprox")) &"' as date)>cast(getdate() as date) then 1 else 0 end as calculo"
          response.write(strSQL&"<br>")
          rsUpdateEntry.Open strSQL, adoCon4
          vstring=trim(rsUpdateEntry("calculo"))
          rsUpdateEntry.close

          if int(vstring)=0 then  response.redirect("ShowContent.asp?contenido=106&msg=la fecha proxima de revision debe ser a futuro" )

          strSQL="select * from Notificaciones where RazonSocial='"&request("RS")&"' and tipo='"&request("ftipo")&"' and DocNum in ("& facturas &")"
          response.write(strSQL&"<br>")
          rsUpdateEntry.Open strSQL, adoCon4   

          strSQL="select * from Notificaciones where RazonSocial='"&request("RS")&"' and modulo='ventas' and tipo='pedido' and DocNum="&rsUpdateEntry("Pedido")
          response.write(strSQL&"<br>")
          rsUpdateEntry2.Open strSQL, adoCon4

          select case request("ftipo")
            case "factura"     strSQL="select a.*,b.WhsCode from OINV a inner join INV1 b on a.docentry=b.docentry where a.Docnum=" & rsUpdateEntry("Docnum")    
            case "remision"    strSQL="select a.*,b.WhsCode from ODLN a inner join DLN1 b on a.docentry=b.docentry where a.Docnum=" & rsUpdateEntry("Docnum")
          end select
         
          select case request("RS")
           case "DMX"  rsUpdateEntry3.Open strSQL,adoCon2
           case "DFW"  rsUpdateEntry3.Open strSQL,adoCon3
           case "MME"  rsUpdateEntry3.Open strSQL,adoCon5
           end select
           vDestino=rsUpdateEntry3("U_Destino")
           vAlmacen=rsUpdateEntry3("WhsCode")
           rsUpdateEntry3.close

          strSQL="insert into Resguardos (ID,DocDate,RS,CardCode,CardName,LogDate,NextDate,SlpEmail,U_Destino,Cantidad,Tipo,Documentos,instrucciones,WhsCode,Pedido) values ((select max(ID) from Resguardos)+1,cast(getdate() as date),'"&request("RS")&"','"& request("fSN") &"','"&rsUpdateEntry("cardName")&"',getdate(),'"&request("ffechaprox")&"','"&rsUpdateEntry2("SlpEmail")&"','"&UCASE(vDestino)&"',"& trim(request("cantidad")) &",'"& request("ftipo") &"','"& facturas &"','"& request("finstruccion") &"','"& vAlmacen &"',"&rsUpdateEntry("Pedido")&")"
          response.write(strSQL&"<br>")
          rsUpdateEntry.close
          rsUpdateEntry2.close
          rsUpdateEntry3.Open strSQL, adoCon4 

          response.redirect("ShowContent.asp?contenido=106&msg=el resguardo ha quedado registrado" )
                 
    end if
    response.write("<P>&nbsp;</P>")



    if request("action")="3" then
          strSQL="select * from resguardos where id="&request("ID")
          rsUpdateEntry3.Open strSQL, adoCon4 

          titulo="edite los datos del resguardo no."&request("ID")
          doTitulo(titulo)

          response.write("<form id='resguardo' method='POST' action='ShowContent.asp'>")
          response.write("<input type='hidden' name='contenido' value='106'>")
          response.write("<input type='hidden' name='action' value='4'>")
      
          response.write("<input type='hidden' name='ID' value='"& rsUpdateEntry3("ID") &"'>")


          CreateTable(960)
          RowIn
            response.write("<td class='subtitulo td-r' width='30%'>Fecha en que inicio el resguardo:</td>")
            response.write("<td class='td-l' width='70%'>"& rsUpdateEntry3("DocDate") &"</td>")
          RowOut
          RowIn
          response.write("<td class='subtitulo td-r'>Razon Social:</td>")
          response.write("<td class='td-l'>"& rsUpdateEntry3("RS") &"</td>")
          RowOut
           RowIn
          response.write("<td class='subtitulo td-r'>Socio de negocio:</td>")
          response.write("<td class='td-l'>"& rsUpdateEntry3("CardName") &"</td>")
          RowOut
          RowIn
          response.write("<td class='subtitulo td-r'>Destino:</td>")
          response.write("<td class='td-l'>"& rsUpdateEntry3("U_Destino") &"</td>")
          RowOut
          'RowIn
          'response.write("<td class='subtitulo td-r'>Email asesor ventas:</td>")
          'response.write("<td class='td-l'>"& rsUpdateEntry3("SlpEmail") &"</td>")
          'RowOut
          RowIn
          response.write("<td class='subtitulo td-r'># Docs SAP:</td>")
          response.write("<td class='td-l'>"& rsUpdateEntry3("Cantidad") &"</td>")
          RowOut
           RowIn
          response.write("<td class='subtitulo td-r'>Tipo de documentos SAP:</td>")
          response.write("<td class='td-l'>"& rsUpdateEntry3("TIPO") &"</td>")
          RowOut
          RowIn
          response.write("<td class='subtitulo td-r'>Documentos:</td>")
          response.write("<td class='td-l'>"& rsUpdateEntry3("Documentos") &"</td>")
          RowOut

          RowIn
          response.write("<td class='subtitulo td-r'>Instrucciones:</td>")
          response.write("<td class='td-l'> <textarea rows='4' cols='80' name='finstruccion'>"&rsUpdateEntry3("instrucciones") &"</textarea></td>")
          RowOut
          RowIn
          response.write("<td class='subtitulo td-r'>Fecha de revision: ("& rsUpdateEntry3("NextDate") &")</td>")
          response.write("<td class='td-l'><input type='date' name='fnextDate'></td>")
          RowOut
          RowIn
          response.write("<td class='subtitulo td-r'>Email responsable:</td>")
          response.write("<td class='td-l'><input type='text' name='femail' value='"& rsUpdateEntry3("SlpEmail")&"'  style='width: 640px'></td>")
          RowOut
          RowIn
          response.write("<td class='td-c'>&nbsp;</td>")
          RowOut
          RowIn
          response.write("<td class='td-c' colspan=2><input type='submit' value='actualizar'></td>")
          RowOut
          RowIn
          response.write("<td class='td-c'>&nbsp;</td>")
          RowOut
          rsUpdateEntry3.close
          closeTable

          response.write("</form>")
    else
          showResguardos
    end if


    
       if request("action")="4" then   'edit UP'
              vNextDate=trim(request("fnextDate"))            
              if vNextDate<>"" then
                    strSQL="select case when '"& vNextDate &"'>cast(getdate() as date) then 1 else 0 end as calculo"
                    rsUpdateEntry5.Open strSQL, adoCon4 
                    if int(trim(rsUpdateEntry5("calculo")))=1 then
                          strSQL="update Resguardos set instrucciones='"& request("finstruccion") &"',NextDate='"&vNextDate&"',SlpEmail='"& request("femail") &"' where ID="&request("ID")      
                    else
                          strSQL="update Resguardos set instrucciones='"& request("finstruccion") &"',SlpEmail='"& request("femail") &"'  where ID="&request("ID")
                    end if  
                    rsUpdateEntry5.close       
              else
                    strSQL="update Resguardos set instrucciones='"& request("finstruccion") &"',SlpEmail='"& request("femail") &"'  where ID="&request("ID")
              end if
              response.write strSQL
              rsUpdateEntry5.Open strSQL, adoCon4 
              response.redirect("ShowContent.asp?contenido=106&msg=su resguardo ha sido actualizado")


     end if
    

    if request("action")="9" then   'BORRAR CONFIRM'
             
         response.write("<P>&nbsp;</P>")

         response.write("<form id='resguardo' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='106'>")
         response.write("<input type='hidden' name='action' value='10'>")
         response.write("<input type='hidden' name='ID' value='"&request("ID") &"'>")
        
         response.write("<P class='td-c'><select name='opcion'><option value='1'>esta seguro desea borrar registro</option><select></P>")          
         response.write("<P class='td-c'><input type='submit' value='borrar'></form></P>")


    end if





    if request("action")="10" then   'BORRAR UP'

        strSQL="update resguardos set activo=0 where ID="&request("ID")
        response.write strSQL
        rsUpdateEntry7.Open strSQL, adoCon4

        response.redirect("ShowContent.asp?contenido=106&msg=el reguardo ha sido concluido")

    end if

    
    

end sub


sub showResguardos  'contenido 106'
    titulo="RESGUARDOS REGISTRADOS"
    DoTitulo(titulo)

    strSQL="select ID,.dbo.fn_GetMonthName(DocDate,'spanish') as Inicio,RS,CardName as SocioNegocio,U_Destino as Destino,Tipo,Documentos,SlpEmail,.dbo.fn_GetMonthName(NextDate,'spanish') as ProxRevision,WhsCode as Almacen,Pedido,Instrucciones  from Resguardos where Activo=1 "    

    if request("ID")<>"" then    strSQL=strSQL &"and ID="&request("ID")
  
    strSQL=strSQL &"  order by SocioNegocio,Pedido"   
    'response.write strSQL
    rsUpdateEntry.Open strSQL, adoCon4
    rsUpdateEntry2.Open strSQL, adoCon4

    if request("ID")<>"" then
           CreateTable(1100)
    else
        response.write("<table width='1100px' border=1 cellpadding=4 cellspacing=2 align=center  class='table2excel table2excel_with_colors' data-tableName='table1'>") 
    end if
    

      Dim Campos(14) 
         i=1         
         For Each rsUpdateEntry in rsUpdateEntry.Fields                        
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next   

         RowIn
         For i=1 to 6
            response.Write("<th class='td-c'>" & campos(i) & "</th>")
         next  
         For i=8 to 11
            response.Write("<th class='td-c'>" & campos(i) & "</th>")
         next  
         RowOut


    if rsUpdateEntry2.EOF then
        NoRegistros
    end if

    while not rsUpdateEntry2.EOF
        RowIn
        for i=1 to 3
            response.write("<td class='td-r fonttiny' width='50px'>"& rsUpdateEntry2(campos(i)) &"</td>")
        next
        response.write("<td class='td-c fonttiny strong' width='100px'>"& rsUpdateEntry2("SocioNegocio") &"</td>")
        for i=5 to 6
            response.write("<td class='td-r fonttiny' width='80px'>"& rsUpdateEntry2(campos(i)) &"</td>")
        next
        'for i=7 to 7
            'response.write("<td class='td-r fonttiny' width='100px'> "& rsUpdateEntry2(campos(i)) &"</td>")
        'next
        for i=8 to 9
            response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=106&action=3&ID="& rsUpdateEntry2(campos(1)) &"'>"& rsUpdateEntry2(campos(i)) &"</a></td>")
        next
        for i=10 to 11
            response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
        next
         Response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=106&action=9&ID="&rsUpdateEntry2("ID")&"'><img src='/img/b_drop.png' with='8px' alt='borrar' title='borrar' height='8px' border=0></a></td>")      
        rowout
        RowIn
        response.write("<td class='td-r fonttiny strong' colspan=4>[Documentos:]</td><td class='td-l fonttiny'  colspan=7>"&rsUpdateEntry2("documentos")&"</td>")
        rowout
        RowIn
        response.write("<td class='td-r fonttiny strong' colspan=4>[Ultimas Instrucciones]</td><td class='td-l fonttiny' style='background-color:#E0E0E0' colspan=7><a href='ShowContent.asp?contenido=106&action=3&ID="& rsUpdateEntry2(campos(1)) &"'>"&rsUpdateEntry2(campos(12)) &"</a></td>")
        rowout
        rsUpdateEntry2.movenext
        separador
    wend
    rsUpdateEntry2.close

     if request("ID")<>"" then
           closeTable
    else  %>
        </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp; 

            <script>
              $(function() {
                $(".exportToExcel").click(function(e){
                  var table = $(this).prev('.table2excel');
                  if(table && table.length){
                    var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
                    $(table).table2excel({
                      exclude: ".noExl",
                      name: "Excel Document Name",
                      filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
                      fileext: ".xls",
                      exclude_img: true,
                      exclude_links: true,
                      exclude_inputs: true,
                      preserveColors: preserveColors
                    });
                  }
                });
                
              });
            </script>
    <%
    end if

end sub  



sub EliminaFechaEmbarque   'contenido 107'   
  
    if request("action")="" then

        response.write("<P>&nbsp;</P>")
        titulo="Elimina Fecha de embarques"       
        DoTitulo(titulo)
        response.write("<form id='utilidad' method='POST' action='ShowContent.asp'>")
        response.write("<input type='hidden' name='contenido' value=107>")

    

       response.write("<input type='hidden' name='action' value=1>")
         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option><option value='MME'>MME</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Pedido: <input class='anchoSmall' type='numeric' name='fpedido' value=" & request("fpedido") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='borrar fechas de embarque'></form></P>")

    else
         response.write("<P>&nbsp;</P>")
         response.write("<P class='alert td-c'>Se eliminaron fechas de embarque</P>")
         titulo="Elimina Fecha de embarques del pedido: " & request("fpedido")
         DoTitulo(titulo)
   
         'QUERY PARA DATOS DEL PEDIDO'
         strSQL="DELETE SAP_FECHAS_EMBARQUE where RS='"& request("fRS") &"' and DocNum="&  request("fpedido")
         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4
        

    end if
 

end sub  



sub AnalisisUtilidadHistorico   'contenido 108'

    doAlert
    titulo="HISTORICO ANALISIS DE UTILIDAD (POR PEDIDO)"
    DoTitulo(titulo)

    strSQL="select TOP(2000) RS,Pedido,PcrtUtilEstm as Util_Estmda,ProvisionFleteMXN as Prov_Flete,Envios,Viaticos,Combustibles,AjusteFlete,UtilidadMXN,UtilidadAjustada,PcrtUtil,PcrtUtilAjustada as UtilAjustada,ID_USUARIO as usuario,case when PcrtUtilEstm<PcrtUtilAjustada then 1 else 0 end as estatus from AnalisisUtilidadPedido order by RS,Pedido desc"

    Const adCmdText = &H0001
    Const adOpenStatic = 3
    nPage=0
    registros=0

     'response.write strSQL 
         rsUpdateEntry.Open strSQL, adoCon4
         rsUpdateEntry2.Open strSQL, adoCon4, adOpenStatic, adCmdText

         rsUpdateEntry2.PageSize = 20 
         nPageCount = rsUpdateEntry2.PageCount

         CreateTable(1200)

         if not rsUpdateEntry.EOF then
               if request("vPage")<>"" then
                    nPage = int(trim(request("vPage")))
                else
                    nPage=1
                end if
              rsUpdateEntry2.AbsolutePage=npage

                response.write("<tr><td colspan=3 class='td-r'><B><U>PAGINAS:</U></B></td><td colspan=8 class='td-l'>")
                for i=1 to nPageCount step 1
                      if i<>npage then  
                          response.write("<a href='/modules/ShowContent.asp?contenido="& request("contenido") &"&vPage="&i &"'>") 
                          response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
                      end if
                    if i=35 or i=70 or i=105 or i=140 then             
                         response.write("<br>")
                    end if
                next
                response.write("</td></tr>")
  
        else

            NoRegistros

        end if


         Dim Campos(14) 
         i=1
         RowIn
         For Each rsUpdateEntry in rsUpdateEntry.Fields            
            'Response.Write("<th>" & rsUpdateEntry.Name & "</th>")
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next   
         For i=1 to 13
            response.Write("<th class='td-c'>" & campos(i) & "</th>")
         next     
         
         RowOut

         while not rsUpdateEntry2.EOF AND registros<=20
             rowin
             Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(1)) &"</td>")
             Response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=85&action=1&fRS=DMX&fPedido="& rsUpdateEntry2(campos(2)) &"' target='utilidad'>"& rsUpdateEntry2(campos(2)) &"</a></td>")
             Response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry2(campos(3)) &"</td>")
             
             for i=4 to 11
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
             next  
             if trim(rsUpdateEntry2("estatus"))="1" then
                 Response.write("<td class='td-r fonttiny strong'><font color='green'>"& rsUpdateEntry2(campos(12)) &"</font></td>") 
             else
                 Response.write("<td class='td-r fonttiny strong'><font color='red'>"& rsUpdateEntry2(campos(12)) &"</font></td>")
             end if

             Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(13)) &"</td>")
             


             rsUpdateEntry2.MoveNext
             registros=registros+1
             RowOut
         wend
         
         rsUpdateEntry2.close
         closeTable

end sub



sub EstimarFlete  'contenido 109'
    doAlert
    titulo="Estime el flete usando nuestras unidades"
    doTitulo(titulo)

    if request("action")="" then    'elija la entidad federativa destino'
         response.write("<div id='destino'>")

         CreateTable(500)
         rowin
         response.write("<td class='subtitulo td-r' width='35%'>Entidad Federativa destino</td>")
         strSQL="select * from cat_entidades where id_entidad<>'00' order by entidad"
         rsUpdateEntry.Open strSQL,adoCon
         response.write("<td class='td-l' width='65%'>")  %>

         <select id="entidad_destino" onchange="Javascript:PassValueChangeDiv('entidad_destino','destino','/modules/FormChangeState3.asp')">   <%'elemento,divn,modulo 

         response.write("<option value='0' SELECTED>seleccione</option>")
         while not rsUpdateEntry.EOF
              response.write("<option value='"&rsUpdateEntry("id_entidad")&"'>"&rsUpdateEntry("entidad")&"</option>")
              rsUpdateEntry.movenext
         wend
         rsUpdateEntry.close
         response.write("</select></td>")
         rowout
         closeTable
         response.write("</div>")
    end if



    if request("action")="1" then  
         response.write("<form id='peajes' method='post' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='109'>")  
         response.write("<input type='hidden' name='action' value=2>")

         select case request("falmacen")
         case "SJR" vEntidadOrigen="21"
                    vCiudadOrigen="San Juan del Rio"
         case "MXL" vEntidadOrigen="02"
                    vCiudadOrigen="Mexicali"
         case "MTY" vEntidadOrigen="18"
                    vCiudadOrigen="Monterrey"
         end select

        vString=trim(request("fciudad_destino"))
        vPos=inStr(vString,"*")
        vString1=mid(vString,1,Vpos-1)
        vString2=mid(vString,Vpos+1,30)

         strSQL="select * from Peajes where Entidad_Origen='"&vEntidadOrigen&"' and ciudad_origen='" & vCiudadOrigen &"' and Entidad_Destino='" & request("entidad_destino") &"' and mpio_destino='" & vString1 &"' and ciudad_destino='"& vString2 &"'"
         'response.write strSQL
         rsUpdateEntry.Open strSQL,adoCon
         if rsUpdateEntry.EOF then
             rsUpdateEntry.close
             strSQL="insert into Peajes (ID,Entidad_Origen,ciudad_origen,Entidad_Destino,mpio_Destino,ciudad_destino) values ((select max(ID) from Peajes)+1,'"&vEntidadOrigen&"','"&vCiudadOrigen&"','"&request("entidad_destino")&"','"&vString1&"','"&vString2&"')"
             'response.write strSQL
             rsUpdateEntry2.Open strSQL,adoCon

             strSQL="select * from Peajes where Entidad_Origen='"&vEntidadOrigen&"' and ciudad_origen='" & vCiudadOrigen &"' and Entidad_Destino='" & request("entidad_destino") &"' and mpio_destino='" & vString1 &"' and ciudad_destino='"& vString2 &"'"
             rsUpdateEntry.Open strSQL,adoCon            
         end if

         strSQL="select * from Flotilla where id_unidad='"& request("funidad")&"'"
         session("id_unidad")=request("funidad")

         rsUpdateEntry2.Open strSQL,adoCon
         vEjes=0
         vEjes=int(rsUpdateEntry2("Total_Llantas"))/2-1
         vString=trim(vEjes)
         vPeaje=0
         vString=rsUpdateEntry(vString)
         vPeaje=int(trim(vString))

         if vPeaje=0 then
             response.write("<div id='peaje'>")

                response.write("<P class='alert td-c'>A&uacute;n no se captura el peaje con estos ejes en esta ruta en intranet , por favor capt&uacute;relo</P>")  
                response.write("<P class='td-c'>Utilice la p&aacute;gina <a href='http://gaia.inegi.org.mx/mdm-client/' target='peaje'><U>INEGI Mapa Digital de Mexico</U> </a> para conocer estos costos </P>")  

                CreateTable(600)
                RowIn
                response.write("<td class='td-c subtitulo'>Origen</td>")
                response.write("<td class='td-c subtitulo'>Entidad <br>Destino</td>")
                response.write("<td class='td-c subtitulo'>Ciudad <br> destino</td>")
                response.write("<td class='td-c subtitulo'>"&vEjes& " Ejes <font color=red> (oneway) </font></td>")
                response.write("<td class='td-c subtitulo'>&nbsp;</td>")
                rowout
                RowIn
                strSQL="select * from cat_entidades where id_entidad='"& request("entidad_destino") &"'"
                rsUpdateEntry7.Open strSQL,adoCon

                response.write("<input type='hidden' id='var1' value='"& vEntidadOrigen &"'>") 
                response.write("<input type='hidden' id='var2' value='"& vCiudadOrigen&"'>")
                response.write("<input type='hidden' id='var3' value='"& request("entidad_destino") &"'>")
                response.write("<input type='hidden' id='var4' value='"& vString1&"'>")  'mpio destino'
                response.write("<input type='hidden' id='var5' value='"& vString2&"'>")  'ciudad destino'
                response.write("<input type='hidden' id='var6' value='"& vEjes &"'>") 'ejes'


                response.write("<td class='td-c'>"& vCiudadOrigen &"</td>")
                response.write("<td class='td-c'>"& rsUpdateEntry7("entidad") &"</td>")
                rsUpdateEntry7.close
                response.write("<td class='td-c'>"& vString2 &"</td>")
                response.write("<td class='td-c'><input type='text' id='var7'></td>")  'variable para capturar peajes'
                response.write("<td class='td-c'>")
                %>
                <input type="button" value="actualizar" onclick="JavaScript:Pass7VarsChangeDiv('peaje','/modules/UpdatePeaje.asp','1')"></td>
                <%
                rowout
                closeTable

             response.write("</div>")

         else
               response.redirect("ShowContent.asp?contenido=109&action=2&ID=" & rsUpdateEntry("ID") &"&ejes=" &vEjes )
         end if
    end if


    if request("action")="2" then  
        '--ruta- ID  ejes: request("ejes")'
        strSQL="select *,["& request("ejes") &"]*2 as Peajes from Peajes where ID="&request("ID")
        rsUpdateEntry.Open strSQL,adoCon

        strSQL="select a.*,b.PrecioLitro from Flotilla a inner join Combustibles b on a.Combustible=b.Combustible where a.id_unidad='"& session("id_unidad") &"'"
        'response.write strSQL
        rsUpdateEntry2.Open strSQL,adoCon

        response.write("<input type='hidden' id='var1' value='"& request("ID") &"'>")        
        response.write("<input type='hidden' id='var2' value='"& session("id_unidad") &"'>")        

        CreateTable(750)
        rowin               
        response.write("<td class='subtitulo td-r'>Unidad a utilizar</td><td colspan=3 class='td-l'><b>"& session("id_unidad") &"</b> " & rsUpdateEntry2("descripcion")  &"</td>")
        rowout
        RowIn
        response.write("<td class='subtitulo td-r' width='25%'>Maxima Carga</td><td class='td-l' width='25%'>"& rsUpdateEntry2("MaxCarga") &" kgs</td>")
        response.write("<td class='subtitulo td-r' width='25%'>Plataforma</td><td class='td-l' width='25%'>"&  rsUpdateEntry2("Plataforma")  &"</td>")  
        rowout
        rowin            
        response.write("<td class='subtitulo td-r'>Ejes</td><td class='td-l'>"& request("ejes") &"</td>")  
        response.write("<td class='subtitulo td-r'>kms x litro "& rsUpdateEntry2("combustible") &"</td><td class='td-l'>"& rsUpdateEntry2("kms_litro") &"</td>")        
        rowout
        rowin            
        response.write("<td class='subtitulo td-r'>Precio por litro</td><td class='td-l'>"& rsUpdateEntry2("PrecioLitro") &"</td>")  
        response.write("<td class='subtitulo td-r'>&nbsp;</td><td class='td-l'>&nbsp;</td>")        
        rowout
        rowin            
        response.write("<td class='subtitulo td-r'>Gasto por peajes <font color=red>(roundtrip)</font></td><td class='td-l'> <font color=red> - "& FormatCurrency(rsUpdateEntry("Peajes")) &"</font></td>")  
        response.write("<td class='td-r subtitulo strong'>Ciudad Origen</td><td class='td-l'>"& rsUpdateEntry("ciudad_origen") &"</td>")        
        rowout
        rowin       
        response.write("<td colspan=2 class='fontmed td-c'><img src='/img/down.gif' border=0 alt='' title=''> ")
        select case  rsUpdateEntry2("clave_sucursal")  
        case "001"  'mexicali'
               response.write("<a class='fontmed' href='https://www.google.com.mx/maps/dir//Deltaflow,+Av.+Del+Arrayan,+Divisi%C3%B3n+Dos+%23421,+El+Sahuaro,+21395+Mexicali,+B.C./@32.6121594,-115.3860765,17z/data=!4m9!4m8!1m0!1m5!1m1!1s0x80d77153e80c52ff:0xfd7eea5f787a5dbb!2m2!1d-115.3838825!2d32.6121549!3e0' target='ruta'>calcule ruta con google maps</a>")
        case "002"  'Monterrey
               response.write("<a class='fontmed' href='https://www.google.com.mx/maps/dir//Deltaflow,+Avenida+Universidad+%23299,+Col,+Banthi,+76804+San+Juan+del+R%C3%ADo,+Qro./@20.3869379,-99.9463833,17z/data=!4m9!4m8!1m0!1m5!1m1!1s0x85d30b13bc514ae1:0xb0507822ecc88203!2m2!1d-99.9441823!2d20.386973!3e0' target='ruta'>calcule ruta con google maps</a>")
        case "003"  'SJR
                 response.write("<a class='fontmed' href='https://www.google.com.mx/maps/dir//Deltaflow,+Avenida+Universidad+%23299,+Col,+Banthi,+76804+San+Juan+del+R%C3%ADo,+Qro./@20.3869379,-99.9463833,17z/data=!4m9!4m8!1m0!1m5!1m1!1s0x85d30b13bc514ae1:0xb0507822ecc88203!2m2!1d-99.9441823!2d20.386973!3e0' target='ruta'>calcule ruta con google maps</a>")
        end select 

        response.write("</td><td class='td-r subtitulo strong'>Ciudad Destino</td><td class='td-l'>"& rsUpdateEntry("ciudad_destino") &"</td>")
        rowout
        rowin
        response.write("<input type='hidden' id='var3' value='"& rsUpdateEntry("Peajes") &"'>")
        response.write("<td class='subtitulo td-r'>Ingrese kms del viaje (one way)</td><td class='td-l'><input type='number' id='var4'></td>")    
         response.write("<td class='subtitulo td-r'>Ingrese subtotal USD cotizacion</td><td class='td-l'><input type='number' id='var5'></td>")
        rowout  
     
        response.write("<input type='hidden' id='var6' value='6'>")
        response.write("<input type='hidden' id='var7' value='7'>")
        %>
        <tr><td class="td-c" colspan="4"><input type="button" value="calcular flete" onclick="JavaScript:Pass7VarsChangeDiv('ruta','/modules/UpdatePeaje.asp','2')"></td></tr>  <%       
        closeTable
        response.write("<div id='ruta'>&nbsp;</div>")

        rsUpdateEntry.close
        rsUpdateEntry2.close

    end if
          
    CloseTable
    
end sub




sub Peajes    'contenido 110
    doAlert
     response.write("<form id='peaje' method='post' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value=110>")


    if request("action")="" then
          titulo="Costos de peajes por Eje de unidad"
          doTitulo(titulo) 

          strSQL="select b.entidad as Estado_destino,a.* from Peajes a inner join cat_entidades b on a.Entidad_Destino=b.id_entidad order by a.ciudad_origen,a.ciudad_destino "
          rsUpdateEntry.Open strSQL,adoCon

          CreateTable(700)
          rowin
          response.write("<td class='td-r subtitulo'>Origen</td>")
          response.write("<td class='td-r subtitulo'>Entidad destino</td>")
          response.write("<td class='td-r subtitulo'>Ciudad destino</td>")
          response.write("<td class='td-r subtitulo'>1 eje</td>")
          response.write("<td class='td-r subtitulo'>2 ejes</td>")
          response.write("<td class='td-r subtitulo'>4 ejes</td>")  

          RowOut
          while not rsUpdateEntry.EOF
              rowin
              response.write("<td class='td-r fontmed'>" & rsUpdateEntry("ciudad_origen") &"</td>")
              response.write("<td class='td-r fontmed'>" & rsUpdateEntry("Estado_destino") &"</td>")
              response.write("<td class='td-r fontmed'>" & rsUpdateEntry("ciudad_destino") &"</td>")
              response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=110&action=1&ID=" & rsUpdateEntry("ID") &"'>" & rsUpdateEntry("1") &"</a></td>")
              response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=110&action=1&ID=" & rsUpdateEntry("ID") &"'>" & rsUpdateEntry("2") &"</a></td>")
              response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=110&action=1&ID=" & rsUpdateEntry("ID") &"'>" & rsUpdateEntry("4") &"</a></td>")        
              rsUpdateEntry.MoveNext
              rowout
          wend
          rsUpdateEntry.close
          closeTable          
    end if

    if request("action")="1" then

          
          titulo="Actualizando costos de peajes por Eje de unidad <br><h3>Ingrese peajes en <font color=red>oneway</font></h3>"
          doTitulo(titulo) 

          response.write("<input type='hidden' name='action' value=2>")
          response.write("<input type='hidden' name='ID' value="& request("ID") &">")

          strSQL="select b.entidad as Estado_destino,a.* from Peajes a inner join cat_entidades b on a.Entidad_Destino=b.id_entidad where ID="&request("ID")
          rsUpdateEntry.Open strSQL,adoCon

          CreateTable(700)
          rowin
          response.write("<td class='td-r subtitulo'>Origen</td>")
          response.write("<td class='td-r subtitulo'>Entidad destino</td>")
          response.write("<td class='td-r subtitulo'>Ciudad destino</td>")
          response.write("<td class='td-c subtitulo'>1 eje</td>")
          response.write("<td class='td-c subtitulo'>2 ejes</td>")
          response.write("<td class='td-c subtitulo'>4 ejes</td>")  
          RowOut
          rowin
              response.write("<td class='td-r fontmed'>" & rsUpdateEntry("ciudad_origen") &"</td>")
              response.write("<td class='td-r fontmed'>" & rsUpdateEntry("Estado_destino") &"</td>")
              response.write("<td class='td-r fontmed'>" & rsUpdateEntry("ciudad_destino") &"</td>")
              response.write("<td class='td-r fontmed'><input type='text' name='f1' value='" & rsUpdateEntry("1") &"'></td>")
              response.write("<td class='td-r fontmed'><input type='text' name='f2' value='" & rsUpdateEntry("2") &"'></td>")
              response.write("<td class='td-r fontmed'><input type='text' name='f4' value='" & rsUpdateEntry("4") &"'></td>")                           
          rowout
          rowin
          response.write("<td colspan=3>&nbsp;</td><td colspan=3 class='td-c'><input type='submit' value='actualizar'></form></td>")
          rsUpdateEntry.close
          closeTable

    end if

    if request("action")="2" then

          strSQL="select b.entidad as Estado_destino,a.* from Peajes a inner join cat_entidades b on a.Entidad_Destino=b.id_entidad where ID="&request("ID")
          rsUpdateEntry.Open strSQL,adoCon

          msg="los peajes para la ciudad: "&rsUpdateEntry("ciudad_destino") &" fueron actualizados"

          vEje1=replace(request("f1"),"..",".")
          vEje1=replace(request("f1"),",","")
          if int(vEje1)<=0 then
              vEje1=trim(rsUpdateEntry("1"))
          end if
          vEje2=replace(request("f2"),"..",".")
          vEje2=replace(request("f2"),",","")
           if int(vEje2)<=0 then
              vEje2=trim(rsUpdateEntry("2"))
          end if
          vEje4=replace(request("f4"),"..",".")
          vEje4=replace(request("f4"),",","")
          if int(vEje4)<=0 then
              vEje4=trim(rsUpdateEntry("4"))
          end if

          rsUpdateEntry.close
          strSQL="update peajes set [1]="&vEje1&",[2]="&vEje2&",[4]="&vEje4&" where ID="&request("ID")
          response.write strSQL
          rsUpdateEntry.Open strSQL,adoCon

          response.redirect("ShowContent.asp?contenido=110&msg=" &msg)
          
    end if
    response.write("<P>&nbsp;</P><P>&nbsp;</P>")

end sub


sub Codigos   'contenido 112'

     doAlert
     response.write("<BR>")
     titulo="DATOS MAESTROS DE CODIGOS<br><font class='fonttiny'><a class='fonttiny' href='ShowContent.asp?contenido=112&cambiar=1'>cambiar de empresa</a></font>"
     DoTitulo(titulo)
     response.write("<P class='td-c fonttiny'>")

    if request("action")="" then     
     response.write("<form id='codigos' method='post' action='ShowContent.asp'>")
     response.write("<input type='hidden' name='contenido' value='112'>")
     response.write("<input type='hidden' name='action' value='1'>")

     response.write("<BR><BR><BR><P class='td-c'>Elija la raz&oacute;n social: <select name='fRS'>")
     response.write("<option value='DMX'>DMX</option>")
     response.write("<option value='DFW'>DFW</option>")  
     response.write("<option value='MME'>MME</option>")
     response.write("</select></P><BR><BR><BR><input type='submit' value='continuar'></form><BR><BR><BR>")
    end if

    if request("cambiar")="1" then
         session("RS")=""
    end if

    if request("action")="1" then

       if request("fRS")<>"" then
           select case request("fRS")
           case "DMX"  session("RS")="DMX"
           case "DFW"  session("RS")="DFW"
           case "MME"  session("RS")="MME"
           end select
       end if

     Const adCmdText = &H0001
     Const adOpenStatic = 3
     nPage=0
     registros=1

     strSQL="select * from OITM where validFor='Y' and ItemCode not like 'SI_%' and ItemCode not like 'COM%'  and ItemCode not like 'CM%' "
     if request("filtro")="" then
         if request("fcodigo")<>"" then
              strSQL=strSQL  &" and ItemCode like '%"& trim(request("fcodigo")) &"%' "
         end if
         if request("fdesc")<>"" then
              strSQL=strSQL  &" and ItemName like '%"& trim(request("fdesc")) &"%' "
         end if
         if request("fUnidad")<>"" then
              strSQL=strSQL  &" and U_CFDi_ClaveUnidad like '%"& trim(request("fUnidad")) &"%' "
         end if          
         if request("fProdServ")<>"" then
              strSQL=strSQL  &" and U_CFDi_ClaveProdServ like '%"& trim(request("fProdServ")) &"%' "
         end if
      else
         select case request("filtro")
         case "1" strSQL=strSQL  &"and ( U_CFDi_ClaveUnidad is null or U_CFDi_ClaveUnidad='' or len(U_CFDi_ClaveUnidad)>3    )"
         case "2" strSQL=strSQL  &"and ( U_CFDi_ClaveProdServ is null or U_CFDi_ClaveProdServ='' )"
         end select
      end if
     'response.write StrSQL&"&nbsp;&nbsp;&nbsp;&nbsp;"&session("RS")

           select case session("RS")
           case "DMX"  rsUpdateEntry.Open strSQL,adoCon2, adOpenStatic, adCmdText
           case "DFW"  rsUpdateEntry.Open strSQL,adoCon3, adOpenStatic, adCmdText
           case "MME"  rsUpdateEntry.Open strSQL,adoCon5, adOpenStatic, adCmdText
           end select

     rsUpdateEntry.PageSize = 50
     nPageCount = rsUpdateEntry.PageCount
      
     CreateTable(900)  

     if not rsUpdateEntry.EOF then
       if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry.AbsolutePage=npage


      response.write("<tr><td class='td-r'><B><U>PAGINAS:</U></B></td><td colspan=3 class='td-l fonttiny'>")
      for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a CLASS='fonttiny' href='/modules/ShowContent.asp?action=1&contenido="& request("contenido") &"&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
          if i=30 or i=60 or i=90 or i=120 or i=150 then             
               response.write("<br>")
          end if
      next
      response.write("</td></tr><tr><td colspan=4>&nbsp;</td></tr>")  
  else
      NoRegistros
  end if

     rowin
     response.write("<td class='subtitulo td-c FontMed'>Codigo</td>")
     response.write("<td class='subtitulo td-c FontMed'>Descripcion</td>")
     response.write("<td class='subtitulo td-c FontMed'><a href='ShowContent.asp?contenido=112&action=1&filtro=1'>CFDi Clave unidad</a></td>")
     response.write("<td class='subtitulo td-c FontMed'><a href='ShowContent.asp?contenido=112&action=1&filtro=2'>CFDi clave ProdServicio</a></td>")
     response.write("<td class='td-c FontMed' width='50px'>&nbsp;</td>")
     rowout

     response.write("<form id='codigos' method='post' action='ShowContent.asp'>")
     response.write("<input type='hidden' name='contenido' value='112'>")
     response.write("<input type='hidden' name='action' value='1'>")

     rowin
     response.write("<td class='td-c'><input type='text' name='fcodigo'></td>")
     response.write("<td class='td-l'><input type='text' name='fdesc'></td>")
     response.write("<td class='td-c'><input type='text' name='fUnidad'></td>")
     response.write("<td class='td-c'><input type='text' name='fProdServ'></td>")
     response.write("<td class='td-c'><input type='submit' value='Buscar'></form></td>")
     rowout
     separador

     while not rsUpdateEntry.EOF AND registros<=50
          RowIn
          response.write("<td class='td-c FontMed'>" & rsUpdateEntry("ItemCode") &"</td>")
          response.write("<td class='td-l fonttiny'>" & rsUpdateEntry("ItemName") &"</td>")
          response.write("<td class='td-c FontMed'>" & rsUpdateEntry("U_CFDi_ClaveUnidad") &"</td>")
          response.write("<td class='td-c FontMed'>" & rsUpdateEntry("U_CFDi_ClaveProdServ") &"</td>")
          rowout
          rsUpdateEntry.MoveNext
          registros=registros+1
     wend
     rsUpdateEntry.close

     closeTable
     response.write("<BR><BR><BR>")
    end if

end sub


sub CotizacionesAbiertas   'contenido 113'

     doAlert
     titulo="cotizaciones abiertas"
     doTitulo(titulo)
     response.write("<P>&nbsp;</P>")

    response.write("<form id='cotizaciones' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=113>")

         if request("action")="" then

             response.write("<input type='hidden' name='action' value=1>")

             CreateTable(500)
             response.write("<tr><td class='td-r'>Asesor: </td><td class='td-l'><select name='fasesor'>")
             response.write("<option value='0' SELECTED>todos</option>")

             strSQL="select * from OSLP where SlpName not like '%disponible%' and ACTIVE='Y' and Slpcode>0 and SlpName not like '%empresas%'"
             rsUpdateEntry.Open strSQL, adoCon2    'DMX'

             while not rsUpdateEntry.EOF
                  response.write("<option value='"& rsUpdateEntry("SlpName") &"'>"&  rsUpdateEntry("SlpName") &"</option>")          
                  rsUpdateEntry.MoveNext
             wend
             rsUpdateEntry.close

             response.write("</select></td></tr>")

             
             response.write("<tr><td class='td-r'>Fecha inicio:</td><td class='td-l'><input type='date' name='ffechaI'></td></tr>")
             response.write("<tr><td class='td-r'>Fecha fin:</td><td class='td-l'><input type='date' name='ffechaF'></td></tr>")

             response.write("<tr><td class='td-r'>Tipo de cotizaci&oacute;n: </td><td class='td-l'><select name='ftipo'>")
             response.write("<option value='1000' SELECTED>todos</option>")
             response.write("<option value='1'>COMPRA</option>")
             response.write("<option value='0'>OTRO</option>")        
             response.write("</select></td></tr>")
      
             response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='continuar'></form></td></tr></table>")
         end if

         if request("action")="1" then

            if request("ffechaI")="" or request("ffechaF")="" then
                  response.redirect("ShowContent.asp?contenido=113&msg=debe especificar un rago de fecha")
            else
                  strSQL="select case when cast('"& request("ffechaF") &"' as date)> cast('"& request("ffechaI") &"' as date) then 1 else 0 end as Flag"
                  rsUpdateEntry.Open strSQL, adoCon4
                  if trim(rsUpdateEntry("flag"))="0" then
                       rsUpdateEntry.close
                       response.redirect("ShowContent.asp?contenido=113&msg=fecha final debe ser mayor que fecha inicial")
                  else
                       rsUpdateEntry.close
                  end if
            end if

            strSQL="SELECT 'DMX' as RS,T1.Docnum as Cotizacion,T1.DocDate as Fecha,T1.cardname as SocioNegocio,(T1.DocTotalSy-T1.VatSumSy) as subtotal,T1.U_PROYECTO as Proyecto,S.SlpName as Asesor,T1.U_TIPOCOTIZACION as Tipo FROM PRODUCTIVA_DMX.dbo.OQUT T1 inner join PRODUCTIVA_DMX.dbo.OSLP S ON T1.SlpCode = S.SlpCode WHERE CANCELED='N' and DocStatus='O' "

            strSQL= strSQL &"and T1.DocDate>='"& request("ffechaI") &"' and T1.DocDate<='"& request("ffechaF") &"' "
            if request("fasesor")<>"" then
               strSQL= strSQL &"and S.Slpname='"& request("fasesor") &"' "
            end if
            select case request("ftipo")   '1 COMPRA   0 NO COMPRA'
                case "1" strSQL= strSQL &"and T1.U_TIPOCOTIZACION='COMPRA' "
                case "0" strSQL= strSQL &"and T1.U_TIPOCOTIZACION!='COMPRA' "
            end select

            strSQL= strSQL &"UNION "

            strSQL= strSQL & "SELECT 'DFW' as RS,T1.Docnum as Cotizacion,T1.DocDate as Fecha,T1.cardname as SocioNegocio,(T1.DocTotalSy-T1.VatSumSy) as subtotal,T1.U_PROYECTO as Proyecto,S.SlpName as Asesor,T1.U_TIPOCOTIZACION as Tipo FROM PRODUCTIVA_DFW.dbo.OQUT T1 inner join PRODUCTIVA_DFW.dbo.OSLP S ON T1.SlpCode = S.SlpCode WHERE CANCELED='N' and DocStatus='O' "

            strSQL= strSQL &"and T1.DocDate>='"& request("ffechaI") &"' and T1.DocDate<='"& request("ffechaF") &"' "
            if request("fasesor")<>"" then
               strSQL= strSQL &"and S.Slpname='"& request("fasesor") &"' "
            end if
            select case request("ftipo")   '1 COMPRA   0 NO COMPRA'
                case "1" strSQL= strSQL &"and T1.U_TIPOCOTIZACION='COMPRA' "
                case "0" strSQL= strSQL &"and T1.U_TIPOCOTIZACION!='COMPRA' "
            end select

            strSQL= strSQL &"UNION "

            strSQL= strSQL & "SELECT 'MME' as RS,T1.Docnum as Cotizacion,T1.DocDate as Fecha,T1.cardname as SocioNegocio,(T1.DocTotalSy-T1.VatSumSy) as subtotal,T1.U_PROYECTO as Proyecto,S.SlpName as Asesor,T1.U_TIPOCOTIZACION as Tipo FROM PRODUCTIVA_MEIDE.dbo.OQUT T1 inner join PRODUCTIVA_MEIDE.dbo.OSLP S ON T1.SlpCode = S.SlpCode WHERE CANCELED='N' and DocStatus='O' "

            strSQL= strSQL &"and T1.DocDate>='"& request("ffechaI") &"' and T1.DocDate<='"& request("ffechaF") &"' "
            if request("fasesor")<>"" then
               strSQL= strSQL &"and S.Slpname='"& request("fasesor") &"' "
            end if
            select case request("ftipo")   '1 COMPRA   0 NO COMPRA'
                case "1" strSQL= strSQL &"and T1.U_TIPOCOTIZACION='COMPRA' "
                case "0" strSQL= strSQL &"and T1.U_TIPOCOTIZACION!='COMPRA' "
            end select

            'strSQL= strSQL &"ORDER BY T1.DocDate"
            'response.write strSQL 

            rsUpdateEntry.Open strSQL, adoCon4
            rsUpdateEntry2.Open strSQL, adoCon4

         response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='1100px' class='table2excel table2excel_with_colors' data-tableName='table1'>")  

         Dim Campos(8) 
         i=1
         RowIn
         
         For Each rsUpdateEntry in rsUpdateEntry.Fields            
            Response.Write("<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>")
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next        
         RowOut

         while not rsUpdateEntry2.EOF
             rowin
             Response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2(campos(1)) &"</td>")
             Response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2(campos(2)) &"</td>")
             Response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2(campos(3)) &"</td>")
             Response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2(campos(4)) &"</td>")
             Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(5)) &"</td>")
             Response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2(campos(6)) &"</td>")
             Response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2(campos(7)) &"</td>")
             Response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2(campos(8)) &"</td>")
        
             rsUpdateEntry2.MoveNext
             RowOut
         wend
         
         rsUpdateEntry2.close
         %>
         </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp;

     <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <%
         end if

end sub




sub BlogVentas 
 
  if request("control")="1" and request("fhilo")<>"" then
        vfecha=trim(request("ffechai"))
        if vfecha="" then vfecha=trim(year(date()))&"-"&trim(month(date()))&"-"&trim(day(date()))
        strSQL="insert into BlogVentas (ID,FATHER,Descripcion,Fechai,ID_USUARIO) values ( (select max(ID)+1 from  BlogVentas),0,'"&request("fhilo")&"','"&vfecha&"','"& session("session_id") & "')" 
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon
        response.redirect("ShowContent.asp?contenido=114&msg=se agrego otro hilo de seguimiento")
  end if

  if request("control")="2" and request("father")<>"" then
        strSQL="insert into BlogVentas (ID,FATHER,Fechai,ID_USUARIO,Entrada) values ( (select isnull(max(ID),0)+1 from  BlogVentas where Father="& request("father")&" and Entrada is not null),"&request("father")&",getdate(),'"& session("session_id") & "','"& request("fentrada") &"')" 
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon
        response.redirect("ShowContent.asp?contenido=114&msg=se agrego un seguimiento a un hilo ")
  end if

  doAlert
  Response.Write("<p>&nbsp;</p>")
  titulo="SEGUIMIENTO DE TEMAS DE VENTAS"
  DoTitulo(Titulo)
 

if request("control")="" then
  
  strSQL="select count(ID) as nodos from BlogVentas where FATHER=0 and FechaF is null"
  'response.write strSQL
  rsUpdateEntry.Open strSQL, adoCon
  n=0
  n=rsUpdateEntry("nodos")
  rsUpdateEntry.close

  response.write("<input type='hidden' ID='nodos' value='"& n &"'>")
  
  strSQL="select ID,descripcion+' ('+ .dbo.fn_GetMonthName(fechaI,'spanish')+')' as nodo from BlogVentas where FATHER=0 and FechaF is null order by ID"
  rsUpdateEntry.Open strSQL, adoCon
  'response.write strSQL

  for i=1 to n
      response.write("<input type='hidden' ID='nodo"& i &"' value='"&rsUpdateEntry("ID")&"|"& rsUpdateEntry("nodo") &"'>")
      rsUpdateEntry.MoveNext
  next
  rsUpdateEntry.close


  strSQL="select ID,FATHER,' ('+ .dbo.fn_GetMonthName(fechaI,'spanish')+') '+substring(Entrada,1,200)+' ...' as entrada from BlogVentas where FATHER>0 and FechaF is null order by FATHER, ID DESC "
  'response.write(strSQL&"<br>")
  rsUpdateEntry.Open strSQL, adoCon 

  vnodos=1
  while not rsUpdateEntry.EOF
       response.write("<input type='hidden' ID='entrada"& vnodos &"' value='"&rsUpdateEntry("ID")&"|"& rsUpdateEntry("FATHER") &"|"& rsUpdateEntry("entrada")&"'>")
      rsUpdateEntry.MoveNext   
      vnodos=vnodos+1   
  wend
  rsUpdateEntry.close
  vnodos=vnodos-1
  'response.write(strSQL&"<br>")
  response.write("<input type='hidden' ID='entradas' value='"& vnodos &"'>")
 %>

  
<script type="text/javascript">

  function Node(id, pid, name, url, title, target, icon, iconOpen, open) {
  this.id = id;
  this.pid = pid;
  this.name = name;
  this.url = url;
  this.title = title;
  this.target = target;
  this.icon = icon;
  this.iconOpen = iconOpen;
  this._io = open || false;
  this._is = false;
  this._ls = false;
  this._hc = false;
  this._ai = 0;
  this._p;
};



// Tree object

function dTree(objName) {
  this.config = {
    target         : null,
    folderLinks     : true,
    useSelection    : true,
    useCookies      : true,
    useLines        : false,
    useIcons        : true,
    useStatusText   : false,
    closeSameLevel  : false,
    inOrder         : false
  }


  this.icon = {
    root        : '/img/cuadrito.jpg',
    folder      : '/img/cuadritodese.jpg',
    folderOpen  : '/img/cuadrito2.jpg',
    node        : '/img/cuadritodese.jpg',
    empty       : '/img/white-opacity-40.png',
    line        : '/img/plus.gif',
    join        : '/img/minus1.gif',
    joinBottom  : '/img/minus2.gif',

          plus        : '/img/close.png',
    plusBottom  : '/img/minus4.gif',
    minus       : '/img/open.png', 
    minusBottom : '/img/close.png',
    nlPlus      : '/img/nolines_plus.gif',
    nlMinus     : '/img/nolines_minus.gif'
  };


  this.obj = objName;
  this.aNodes = [];
  this.aIndent = [];
  this.root = new Node(-1);
  this.selectedNode = null;
  this.selectedFound = false;
  this.completed = false;

};

// Adds a new node to the node array
dTree.prototype.add = function(id, pid, name, url, title, target, icon, iconOpen, open) {
  this.aNodes[this.aNodes.length] = new Node(id, pid, name, url, title, target, icon, iconOpen, open);
};


// Open/close all nodes
dTree.prototype.openAll = function() {
  this.oAll(true);
};

dTree.prototype.closeAll = function() {
  this.oAll(false);
};



// Outputs the tree to the page
dTree.prototype.toString = function() {
  var str = '<div class="dtree">\n';
  if (document.getElementById) {
    if (this.config.useCookies) this.selectedNode = this.getSelected();
    str += this.addNode(this.root);
  } else str += 'Browser not supported.';

  str += '</div>';
  if (!this.selectedFound) this.selectedNode = null;
  this.completed = true;
  return str;
};



// Creates the tree structure
dTree.prototype.addNode = function(pNode) {
  var str = '';
  var n=0;
  if (this.config.inOrder) n = pNode._ai;
  for (n; n<this.aNodes.length; n++) {
    if (this.aNodes[n].pid == pNode.id) {
      var cn = this.aNodes[n];
      cn._p = pNode;
      cn._ai = n;
      this.setCS(cn);
      if (!cn.target && this.config.target) cn.target = this.config.target;
      if (cn._hc && !cn._io && this.config.useCookies) cn._io = this.isOpen(cn.id);
      if (!this.config.folderLinks && cn._hc) cn.url = null;
      if (this.config.useSelection && cn.id == this.selectedNode && !this.selectedFound) {
          cn._is = true;
          this.selectedNode = n;
          this.selectedFound = true;
      }
      str += this.node(cn, n);
      if (cn._ls) break;
    }
  }
  return str;
};



// Creates the node icon, url and text
dTree.prototype.node = function(node, nodeId) {
  var str = '<div class="dTreeNode"><br>' + this.indent(node, nodeId);
  if (this.config.useIcons) {
    if (!node.icon) node.icon = (this.root.id == node.pid) ? this.icon.root : ((node._hc) ? this.icon.folder : this.icon.node);
    if (!node.iconOpen) node.iconOpen = (node._hc) ? this.icon.folderOpen : this.icon.node;
    if (this.root.id == node.pid) {
      node.icon = this.icon.root;
      node.iconOpen = this.icon.root;
    }

    str += '<img id="i' + this.obj + nodeId + '" src="' + ((node._io) ? node.iconOpen : node.icon) + '" alt="" />';

  }

  if (node.url) {

    //str += '<a id="s' + this.obj + nodeId + '" class="' + ((this.config.useSelection) ? ((node._is ? 'nodeSel' : 'node')) : 'node') + '" href=""';
    //str += '<a href="" onclick="javascript:RunModule(&#39;'+ node.url +'&#39;);"';
    //el hiperlink de los nodos
    str += '&nbsp;&nbsp;&nbsp;<a href="javascript:sendReq(&#39;'+ node.url +'&#39;);">'+ node.name +'</a>';
    //if (node.title) str += ' title="' + node.title + '"';
    //if (node.target) str += ' target="' + node.target + '"';
    //if (this.config.useStatusText) str += ' onmouseover="window.status=\'' + node.name + '\';return true;" onmouseout="window.status=\'\';return true;" ';
    //if (this.config.useSelection && ((node._hc && this.config.folderLinks) || !node._hc))
    //str += ' onclick="javascript:RunModule(&#39;'+ node.url +'&#39;);"';
    //str += '>';
  }

  else {str += '<a href="javascript: ' + this.obj + '.o(' + nodeId + ');" class="node">'+ node.name + '</a>'  };
  //else if ((!this.config.folderLinks || !node.url) && node._hc && node.pid != this.root.id)

    //str += '<a href="javascript: ' + this.obj + '.o(' + nodeId + ');" class="node">';

  //str += node.name;

  //if (node.url || ((!this.config.folderLinks || !node.url) && node._hc)) str += '</a>';

  str += '</div>';

  if (node._hc) {

    str += '<div id="d' + this.obj + nodeId + '" class="clip" style="display:' + ((this.root.id == node.pid || node._io) ? 'block' : 'none') + ';">';

    str += this.addNode(node);

    str += '</div><br><br>';  //parece que aqui termina el nodo

  }
  this.aIndent.pop();
  return str;
};



// Adds the empty and line icons
dTree.prototype.indent = function(node, nodeId) {   //lo desactive

  //var str = '';
  //if (this.root.id != node.pid) {
  //  for (var n=0; n<this.aIndent.length; n++)

  //    str += '<img src="' + ( (this.aIndent[n] == 1 && this.config.useLines) ? this.icon.line : this.icon.empty ) + '" alt="" />';

  //  (node._ls) ? this.aIndent.push(0) : this.aIndent.push(1);

  //  if (node._hc) {

  //    str += '<a href="javascript: ' + this.obj + '.o(' + nodeId + ');"><img id="j' + this.obj + nodeId + '" src="';

  //    if (!this.config.useLines) str += (node._io) ? this.icon.nlMinus : this.icon.nlPlus;

  //    else str += ( (node._io) ? ((node._ls && this.config.useLines) ? this.icon.minusBottom : this.icon.minus) : ((node._ls && this.config.useLines) ? this.icon.plusBottom : this.icon.plus ) );

  //    str += '" alt="" /></a>';

  //  } else str += '<img src="' + ( (this.config.useLines) ? ((node._ls) ? this.icon.joinBottom : this.icon.join ) : this.icon.empty) + '" alt="" />';
  //}
  //return str;
};



// Checks if a node has any children and if it is the last sibling
dTree.prototype.setCS = function(node) {
  var lastId;
  for (var n=0; n<this.aNodes.length; n++) {
    if (this.aNodes[n].pid == node.id) node._hc = true;
    if (this.aNodes[n].pid == node.pid) lastId = this.aNodes[n].id;
  }
  //if (lastId==node.id) node._ls = true; //desactive que identificara la ultima linea
};



// Returns the selected node
dTree.prototype.getSelected = function() {
  var sn = this.getCookie('cs' + this.obj);
  return (sn) ? sn : null;
};



// Highlights the selected node
dTree.prototype.s = function(id) {
  if (!this.config.useSelection) return;
  var cn = this.aNodes[id];
  if (cn._hc && !this.config.folderLinks) return;
  if (this.selectedNode != id) {

    if (this.selectedNode || this.selectedNode==0) {
      eOld = document.getElementById("s" + this.obj + this.selectedNode);
      eOld.className = "node";
    }

    eNew = document.getElementById("s" + this.obj + id);
    eNew.className = "nodeSel";
    this.selectedNode = id;
    if (this.config.useCookies) this.setCookie('cs' + this.obj, cn.id);
  }
};



// Toggle Open or close

dTree.prototype.o = function(id) {
  var cn = this.aNodes[id];
  this.nodeStatus(!cn._io, id, cn._ls);
  cn._io = !cn._io;
  if (this.config.closeSameLevel) this.closeLevel(cn);
  if (this.config.useCookies) this.updateCookie();
};



// Open or close all nodes
dTree.prototype.oAll = function(status) {
  for (var n=0; n<this.aNodes.length; n++) {
    if (this.aNodes[n]._hc && this.aNodes[n].pid != this.root.id) {
      this.nodeStatus(status, n, this.aNodes[n]._ls)
      this.aNodes[n]._io = status;
    }
  }
  if (this.config.useCookies) this.updateCookie();
};



// Opens the tree to a specific node
dTree.prototype.openTo = function(nId, bSelect, bFirst) {
  if (!bFirst) {
    for (var n=0; n<this.aNodes.length; n++) {
      if (this.aNodes[n].id == nId) {
        nId=n;
        break;
      }
    }
  }

  var cn=this.aNodes[nId];
  if (cn.pid==this.root.id || !cn._p) return;
  cn._io = true;
  cn._is = bSelect;
  if (this.completed && cn._hc) this.nodeStatus(true, cn._ai, cn._ls);
  if (this.completed && bSelect) this.s(cn._ai);
  else if (bSelect) this._sn=cn._ai;
  this.openTo(cn._p._ai, false, true);
};



// Closes all nodes on the same level as certain node
dTree.prototype.closeLevel = function(node) {
  for (var n=0; n<this.aNodes.length; n++) {
    if (this.aNodes[n].pid == node.pid && this.aNodes[n].id != node.id && this.aNodes[n]._hc) {
      this.nodeStatus(false, n, this.aNodes[n]._ls);
      this.aNodes[n]._io = false;
      this.closeAllChildren(this.aNodes[n]);
    }
  }
}



// Closes all children of a node
dTree.prototype.closeAllChildren = function(node) {
  for (var n=0; n<this.aNodes.length; n++) {
    if (this.aNodes[n].pid == node.id && this.aNodes[n]._hc) {
      if (this.aNodes[n]._io) this.nodeStatus(false, n, this.aNodes[n]._ls);
      this.aNodes[n]._io = false;
      this.closeAllChildren(this.aNodes[n]);    
    }
  }
}



// Change the status of a node(open or closed)
dTree.prototype.nodeStatus = function(status, id, bottom) {
  eDiv  = document.getElementById('d' + this.obj + id);
  eJoin = document.getElementById('j' + this.obj + id);
  if (this.config.useIcons) {
    eIcon = document.getElementById('i' + this.obj + id);
    eIcon.src = (status) ? this.aNodes[id].iconOpen : this.aNodes[id].icon;
  }

  eJoin.src = (this.config.useLines)?
  ((status)?((bottom)?this.icon.minusBottom:this.icon.minus):((bottom)?this.icon.plusBottom:this.icon.plus)):
  ((status)?this.icon.nlMinus:this.icon.nlPlus);
  eDiv.style.display = (status) ? 'block': 'none';
};





// [Cookie] Clears a cookie
dTree.prototype.clearCookie = function() {
  var now = new Date();
  var yesterday = new Date(now.getTime() - 1000 * 60 * 60 * 24);
  this.setCookie('co'+this.obj, 'cookieValue', yesterday);
  this.setCookie('cs'+this.obj, 'cookieValue', yesterday);
};



// [Cookie] Sets value in a cookie
dTree.prototype.setCookie = function(cookieName, cookieValue, expires, path, domain, secure) {
  document.cookie =
    escape(cookieName) + '=' + escape(cookieValue)
    + (expires ? '; expires=' + expires.toGMTString() : '')
    + (path ? '; path=' + path : '')
    + (domain ? '; domain=' + domain : '')
    + (secure ? '; secure' : '');
};



// [Cookie] Gets a value from a cookie

dTree.prototype.getCookie = function(cookieName) {
  var cookieValue = '';
  var posName = document.cookie.indexOf(escape(cookieName) + '=');
  if (posName != -1) {
    var posValue = posName + (escape(cookieName) + '=').length;
    var endPos = document.cookie.indexOf(';', posValue);
    if (endPos != -1) cookieValue = unescape(document.cookie.substring(posValue, endPos));
    else cookieValue = unescape(document.cookie.substring(posValue));
  }
  return (cookieValue);
};



// [Cookie] Returns ids of open nodes as a string
dTree.prototype.updateCookie = function() {
  var str = '';
  for (var n=0; n<this.aNodes.length; n++) {
    if (this.aNodes[n]._io && this.aNodes[n].pid != this.root.id) {
      if (str) str += '.';
      str += this.aNodes[n].id;
    }
  }
  this.setCookie('co' + this.obj, str);
};



// [Cookie] Checks if a node id is in a cookie
dTree.prototype.isOpen = function(id) {
  var aOpen = this.getCookie('co' + this.obj).split('.');
  for (var n=0; n<aOpen.length; n++)
    if (aOpen[n] == id) return true;
  return false;
};



// If Push and pop is not implemented by the browser
if (!Array.prototype.push) {
  Array.prototype.push = function array_push() {
    for(var i=0;i<arguments.length;i++)
      this[this.length]=arguments[i];
    return this.length;
  }

};

if (!Array.prototype.pop) {
  Array.prototype.pop = function array_pop() {
    lastElement = this[this.length-1];
    this.length = Math.max(this.length-1,0);
    return lastElement;
  }
};

//<a href="javascript: d.openAll();">Expandir Todo</a>  ABAJO
var d = new dTree('d');

d.add(0,-1,'<a href="javascript: d.closeAll();">Compactar todo</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<a href="Javascript:sendReq(\'blogActions.asp\',\'control\',\'1\',\'dtree-root\')">agregar otro hilo</a><br><br>');
   
   for(nodoID=1;nodoID<=document.getElementById('nodos').value;nodoID++){
            var cadena=document.getElementById('nodo'+nodoID).value;
            var arrayDeCadenas = cadena.split('|');
            d.add(nodoID,0,arrayDeCadenas[1],'blogActions.asp\',\'control,ID\',\'2,'+arrayDeCadenas[0]+'\',\'dtree-root'); 
   } 

   
   for(EntradaID=1;EntradaID<=document.getElementById('entradas').value;EntradaID++){                
            var cadena=document.getElementById('entrada'+EntradaID).value;         
            var arrayDeCadenas = cadena.split('|');      
            //d.add(12,1,'EL SAUZAL','tree.asp?string=001*EL SAUZAL');         
            //d.add(arrayDeCadenas[0],arrayDeCadenas[1],arrayDeCadenas[2],'link'); 
            d.add(arrayDeCadenas[0],arrayDeCadenas[1],arrayDeCadenas[2],'blogActions.asp\',\'control,ID\',\'2,'+arrayDeCadenas[0]+'\',\'dtree-root');
   } 

</script>

 <div class="dtree-root" id="dtree-root">
    <script type="text/javascript">     
         document.write(d);  
    </script>
</div>

<%

  if request("control")="1" then  'ADD UP'
      response.write( request("fhilo")  )

  end if


end if


end sub


sub ComportamientoPagos  'contenido 114'

    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"& request("msg") &"</P>")
    end if
   
   titulo="Comportamiento de pagos (lineas de credito)"
   DoTitulo(titulo)


         response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=114>")
         response.write("<input type='hidden' name='action' value=1>")

         strSQL="select 'DFW-'+substring(CardName,1,30) as NombreCorto,'DFW' as RS,cardCode from PRODUCTIVA_DFW.dbo.OCRD where DebtLine>0 and substring(cardCode,1,1)='C' and cardname not like '%deltaflow%' and cardname not like '%meide%'  UNION select 'DMX-'+substring(CardName,1,30) as NombreCorto,'DMX' as RS,cardCode from PRODUCTIVA_DMX.dbo.OCRD where DebtLine>0 and substring(cardCode,1,1)='C' and cardname not like '%deltaflow%' and cardname not like '%meide%'"

         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4 

         response.write("<table width='800px' cellpadding=4 cellspacing=3 align='center'><tr><td class='td-r'>Socio de negocio: </td><td class='td-l'><select name='fSN'>")

         response.write("<option value='0'>sin seleccionar</option>")

         while not rsUpdateEntry.EOF              
                  response.write("<option value='"& rsUpdateEntry("RS") &"*" & rsUpdateEntry("cardCode") &"'>"&  rsUpdateEntry("NombreCorto") &"</option>")
              
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close

         response.write("</select> </td></tr>")
       

         response.write("<tr><td class='td-r'>Fecha inicial: </td><td class='td-l'><input type='date' name='ffechai'></td></tr>")
         response.write("<tr><td class='td-r'>Fecha final: </td><td class='td-l'><input type='date' name='ffechaf'></td></tr>")
        
          
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='mostrar informe'></form></td></tr></table>")


    if request("action")=1 then
          if request("ffechai")="" or request("ffechaf")="" then
                response.redirect("ShowContent.asp?contenido=114&msg=debe especificar un periodo de fechas")
          end if
          if request("fSN")="0" then
                response.redirect("ShowContent.asp?contenido=114&msg=debe especificar un socio de negocio")
          end if

          strSQL="select case when cast('"& request("ffechaf") &"' as date)> cast('"& request("ffechai") &"' as date) then 1 else 0 end as estatus"
          'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon4 
          vstring=trim(rsUpdateEntry("estatus"))
          rsUpdateEntry.close

          if vstring="0" then response.redirect("ShowContent.asp?contenido=114&msg=fecha final debe ser superior a fecha inicial")

           vString=request("fSN")
           vPos=inStr(vString,"*")
           vString1=mid(vString,1,Vpos-1)
           vString2=mid(vString,Vpos+1,30)

          strSQL="SELECT T0.DocNum,T0.DocTotalFC,T0.U_CFDi_FormaPago,T0.DocDate,T0.DocDueDate,DATEDIFF(day,T0.DocDate,T0.DocDueDate) as DiasCredito,T0.U_CFDi_TimbreUUID,TC0.DocNum as Pago,TC0.DocDate as FechaPago,case when DateDiff(DAY,TC0.DocDate,T0.DocDueDate)<0 THEN cast( DateDiff(DAY,TC0.DocDate,T0.DocDueDate)*-1 as varchar)+' dias despues de vencimiento' else cast( DateDiff(DAY,TC0.DocDate,T0.DocDueDate) as varchar)+' dias antes vencimiento' end FROM OINV T0  LEFT JOIN RCT2 TC1 ON T0.DocEntry = TC1.baseAbs AND TC1.InvType = 13 LEFT JOIN ORCT TC0 ON TC0.DocEntry = TC1.DocNum WHERE T0.U_CFDi_FormaPago=99 and  T0.CANCELED='N' and   T0.DocDate>='"& request("ffechai") &"' and T0.DocDate<='"& request("ffechaf") &"' and T0.CardCode='"& vString2 &"' and   TC0.Canceled='N' order by T0.DocNum"

          'response.write(strSQL&"<br>")
          select case vString1
          case "DFW" rsUpdateEntry.Open strSQL, adoCon3 
                     rsUpdateEntry2.Open strSQL, adoCon3 
          case "DMX" rsUpdateEntry.Open strSQL, adoCon2 
                     rsUpdateEntry2.Open strSQL, adoCon2 
          end select
        
          response.write("<P>&nbsp;</P>")
          CreateTable(1000)

          Dim Campos(10)
          i=1
          RowIn
          For Each rsUpdateEntry in rsUpdateEntry.Fields
                  Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
                  campos(i)=rsUpdateEntry.Name
                  i=i+1
          Next 
          rowout

          if rsUpdateEntry2.EOF then NoRegistros

     while not rsUpdateEntry2.EOF        
           rowin
             for i=1 to 10
               Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
             next  
           rsUpdateEntry2.MoveNext
     wend
     closeTable
     rsUpdateEntry2.close
   
    end if
end sub


sub CostoInterE  'contenido 115'

 if request("fmes")<>"" then
           vmes=request("fmes")
           vanio=request("fanio") 
     else
           vmes=month(date())
           vanio=year(date())
  end if

  Titulo="HISTORICO EN COSTOS DE TRANSFERENCIA: "& meses(vmes) &", "&vanio&"<BR style='font-size:3px;height:3px'><form id='costos' method='POST' action='ShowContent.asp'><input type='hidden' name='contenido' value='115'><select name='fmes'>"

       for i=1 to 12 step 1
          if month(date())=i then
            Titulo=Titulo&"<option value='"&i&"' SELECTED>"&meses(i)&"</option>"
          else
                Titulo=Titulo&"<option value='"&i&"'>"&meses(i)&"</option>"
          end if
       next
       Titulo=Titulo&"</select>&nbsp;&nbsp;&nbsp;<select name='fanio'>"
       for i=year(date()) to year(date())-1 step -1
          Titulo=Titulo&"<option value='"&i&"'>"& i &"</option>"
       next
       Titulo=Titulo&"</select>&nbsp;&nbsp;<input type='submit' value='mostrar'></form>"

       strSQL="select   round(   sum( (b.Price*a.DocRate-GrossBuyPr)*b.Quantity) / sum(b.Price*a.DocRate*b.Quantity) * 100 ,2)  as calculo from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry where year(a.DocDate)="& vanio &" and month(a.DocDate)="& vmes &" and a.CANCELED='N' and a.CardName like '%deltaflow%' "
       rsUpdateEntry.Open strSQL, adoCon2  'DMX'

       Dotitulo(Titulo)  

       response.write("<P class='td-c'>% Global de transferencia del periodo (DMX-DFW): " & rsUpdateEntry("calculo") &" %</P>")
       rsUpdateEntry.close
              
       strSQL="select   a.DocNum as Remision,a.DocDate,sum( (b.Price*a.DocRate-GrossBuyPr)*b.Quantity) as Utilidad,sum(b.Price*a.DocRate*b.Quantity) as Ingresos,round(   sum( (b.Price*a.DocRate-GrossBuyPr)*b.Quantity) / sum(b.Price*a.DocRate*b.Quantity) * 100 ,2) as MUB from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry where year(a.DocDate)="& vanio &" and month(a.DocDate)="& vmes &" and a.CardName like '%deltaflow%' and a.CANCELED='N' group by a.DocNum,a.DocDate order by a.DocNum"

       rsUpdateEntry.Open strSQL, adoCon2
       rsUpdateEntry2.Open strSQL, adoCon2
       
       CreateTable(700)

       Dim Campos(5) 
         i=1
         RowIn
         For Each rsUpdateEntry in rsUpdateEntry.Fields            
            Response.Write("<th class='subtitulo td-c'>" & rsUpdateEntry.Name & "</th>")
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next        
         RowOut

         if rsUpdateEntry2.EOF then
                 NoRegistros
         else         
          
            while not rsUpdateEntry2.EOF

                response.write("<tr>")
                response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2(campos(1)) &"</td>")
                response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2(campos(2)) &"</td>")
                response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2(campos(3)) &"</td>")
                response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2(campos(4)) &"</td>")
                response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2(campos(5)) &"</td>")
                rsUpdateEntry2.MoveNext
            wend
            rsUpdateEntry2.close


          end if


end sub


sub ShowXmlsCombustibles   'CONTENIDO 116'

    doAlert
    titulo="ESTADOS DE CUENTA COMBUSTIBLES"
    Dotitulo(Titulo)

    response.write("<P class='td-c'>")  %>
    <input type="button" onclick="window.location='http://192.168.0.206/uploadform-combustibles.asp'" value='agregar'></P>
    <%
    titulo="Ultimos ejercicios realizados"
    doTitulo(titulo)

  Dim strSQL
  strSQL="select  dbo.fn_GetMonthName(fecha,'spanish') as FechaTimbre,TipoComprobante as Tipo,MetodoPago as Metodo,FormaPago,Moneda,Folio,Subtotal,Impuestos,Total,Emisor,RFCReceptor,UUID,Substring(Concepto,1,23) Concepto,NombreXML as EmitirInforme from xml where CardCode='A000015' order by Fecha desc"
  'Response.write strSQL

  Const adCmdText = &H0001
  Const adOpenStatic = 3

  Dim nPage,registros,nPageCount
  nPage=0
  registros=1

  rsUpdateEntry.Open strSQL, adoCon4
  rsUpdateEntry2.Open strSQL, adoCon4, adOpenStatic, adCmdText

  rsUpdateEntry2.PageSize = 10 
  nPageCount = rsUpdateEntry2.PageCount

  Dim Campos(14)
  Dim Contador
  contador=1

  CreateTable(1200)
        if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry2.AbsolutePage=npage
      response.write("<tr><td colspan=3 class='td-r'>PAGINAS:</td><td colspan=10>")
       for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='ShowContent.asp?contenido=116&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
        next
        response.write("</td></tr>")


         For Each rsUpdateEntry in rsUpdateEntry.Fields
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
            campos(contador)=rsUpdateEntry.Name
            contador=contador+1
         Next

         while not rsUpdateEntry2.EOF AND registros<=10
           rowin
             for contador=1 to 13
               Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(contador)) &"</td>")
             next
             
            
             Response.write("<td class='td-r fonttiny'><a href='/modules/ShowContent.asp?contenido=94&xml="& rsUpdateEntry2("EmitirInforme") &"'><img src='/images/alert.gif' border=0 alt='informe' title='informe' width='14px' height='14px'></a></td>")

           rowout
             rsUpdateEntry2.MoveNext
             registros=registros+1
         wend
         rsUpdateEntry2.close

 
  closeTable

end sub




sub DirectorioClientes  'contenido 117'
      doAlert

      titulo="Directorio de SN con l&iacute;neas de cr&eacute;dito"
      DoTitulo(titulo)      

      strSQL="select TOP 1 AUX.CardName from SBOTemp.dbo.Repositorio a inner join  ( select 'DMX' as RazonSocial,CardCode,CardName from PRODUCTIVA_DMX.dbo.OCRD   where substring(cardcode,1,1)='C' and CardName not like '%deltaflow%' and  CardName not like '%meide%' and DebtLine>0   UNION   select 'DFW' as RazonSocial,cardcode,CardName   from PRODUCTIVA_DFW.dbo.OCRD    where substring(cardcode,1,1)='C' and CardName not like '%deltaflow%' and  CardName not like '%meide%' and DebtLine>0   ) as AUX on a.RazonSocial=aux.RazonSocial collate SQL_Latin1_General_CP1_CI_AS and a.CardCode=AUX.CardCode collate SQL_Latin1_General_CP1_CI_AS where AUX.CardName collate SQL_Latin1_General_CP1_CI_AS not in (select cardname from DirectorioClientes)  "      
      rsUpdateEntry.Open strSQL, adoCon4

      if not rsUpdateEntry.EOF then
        strSQL="insert into DirectorioClientes (ID,CardName) values ((select MAX(ID)+1 from DirectorioClientes),'"& rsUpdateEntry("cardname") &"') "
        rsUpdateEntry2.Open strSQL, adoCon4        
      end if
      rsUpdateEntry.close

      strSQL="UPDATE b set b.municipio=c.municipio from PRODUCTIVA_DMX.dbo.OCRD a inner join SBOTemp.dbo.DirectorioClientes b on a.CardName collate SQL_Latin1_General_CP1_CI_AS=b.cardName  inner join INTRANET.dbo.CodigosPostales c on a.ZipCode collate SQL_Latin1_General_CP1_CI_AS=c.CP  where substring(a.cardcode,1,1)='C' and b.municipio is null"
      rsUpdateEntry2.Open strSQL, adoCon4

      strSQL="UPDATE b set b.municipio=c.municipio from PRODUCTIVA_DFW.dbo.OCRD a inner join SBOTemp.dbo.DirectorioClientes b on a.CardName collate SQL_Latin1_General_CP1_CI_AS=b.cardName  inner join INTRANET.dbo.CodigosPostales c on a.ZipCode collate SQL_Latin1_General_CP1_CI_AS=c.CP  where substring(a.cardcode,1,1)='C' and b.municipio is null"
      rsUpdateEntry2.Open strSQL, adoCon4

      if request("action")="1" then
           strSQL="UPDATE DirectorioClientes set Dueno='"& request("fduenio") &"',RepLegal='"& request("fRepLegal") &"',CuentasPagar='"& request("fcuentas") &"',Abogado='"& request("fabogado") &"',Gerentecompras='"& request("fGerenteCompras") &"',Comprador='"& request("fcomprador") &"',Suministro='"& request("fsuministro") &"',Presupuestos='"& request("fpresupuestos") &"' where ID="& request("ID")
           response.write strSQL
           rsUpdateEntry.Open strSQL, adoCon4
           response.redirect("ShowContent.asp?contenido=117&msg=datos del socio de negocio actualizado")

      end if


      strSQL="select * from DirectorioClientes "
      if request("ID")<>"" then
             strSQL= strSQL&" where ID=" & request("ID")
      end if
      strSQL= strSQL&" order by CardName"

      rsUpdateEntry.Open strSQL, adoCon4
      response.write("<table width='1300px' cellspacing=2 cellpadding=3 border=1>")      
      RowIn  
      response.write("<th class='subtitulo td-c'>Socio de Negocio</th>")
      response.write("<th class='subtitulo td-c'>Municipio</th>")
      response.write("<th class='subtitulo td-c'>Due&ntilde;o</th>")
      response.write("<th class='subtitulo td-c'>Rep. Legal</th>")
      response.write("<th class='subtitulo td-c'>CuentasPagar</th>")
      response.write("<th class='subtitulo td-c'>Abogado</th>")
      response.write("<th class='subtitulo td-c'>Gerente Compras</th>")
      response.write("<th class='subtitulo td-c'>Compradores</th>")
      response.write("<th class='subtitulo td-c'>Presupuestos</th>")
      response.write("<th class='subtitulo td-c'>Suministro</th>")      
      rowout

      while not rsUpdateEntry.EOF
           RowIn
           response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=117&ID="& rsUpdateEntry("ID") &"'>" & rsUpdateEntry("cardname") &"</a></td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("municipio") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("Dueno") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("RepLegal") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("CuentasPagar") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("Abogado") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("GerenteCompras") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("Comprador") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("Presupuestos") &"</td>")
           response.write("<td class='td-r fonttiny'>" & rsUpdateEntry("Suministro") &"</td>")
           rsUpdateEntry.MoveNext
           rowout
      wend

      closeTable


      if request("ID")<>"" then
         rsUpdateEntry.MoveFirst
         response.write("<form id='Directorio' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=117>")
         response.write("<input type='hidden' name='ID' value='"& request("ID") &"'>")
         response.write("<input type='hidden' name='action' value=1>")

         CreateTable(800)
        
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Due&ntilde;o:</td><td><input type='text' name='fduenio' value='"& rsUpdateEntry("Dueno")&"' class='anchox4'></td>")
         RowOut
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Representante Legal:</td><td><input type='text' name='fRepLegal' value='"& rsUpdateEntry("RepLegal")&"' class='anchox4'></td>")
         RowOut
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Cuentas por pagar:</td><td><input type='text' name='fCuentas' value='"& rsUpdateEntry("CuentasPagar")&"' class='anchox4'></td>")
         RowOut
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Abogado:</td><td><input type='text' name='fabogado' value='"& rsUpdateEntry("Abogado")&"' class='anchox4'></td>")
         RowOut
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Gerente de Compras:</td><td><input type='text' name='fGerenteCompras' value='"& rsUpdateEntry("GerenteCompras")&"' class='anchox4'></td>")
         RowOut
         
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Compradores:</td><td><input type='text' name='fcomprador' value='"& rsUpdateEntry("Comprador")&"' class='anchox4'></td>")
         RowOut
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Presupuestos:</td><td><input type='text' name='fpresupuestos' value='"& rsUpdateEntry("Presupuestos")&"' class='anchox4'></td>")
         RowOut
         RowIn
         response.write("<td class='subtitulo fontmed td-r'>Encargado del suministro:</td><td><input type='text' name='fsuministro' value='"& rsUpdateEntry("Suministro")&"' class='anchox4'></td>")
         RowOut

         response.write("<td class='td-c' colspan=2>&nbsp;</td></tr>") 
         response.write("<td class='td-c' colspan=2><input type='submit' value='actualizar'></form></td></tr>") 
         closeTable

      end if


      rsUpdateEntry.close


end sub




Sub DetalleFactura   'contenido 119'

    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then  response.write("<P class='alert td-c'>"&  request("msg") &"</P>")   

    titulo="Detalle de una factura"
    DoTitulo(titulo)

    if request("action")="" then

         response.write("<form id='inv' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=119>")
         response.write("<input type='hidden' name='action' value=1>")


         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option><option value='MEIDE'>MME</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Pedido: <input class='anchoSmall' type='numeric' name='ffactura' value="&request("ffactura") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

     else

         strSQL="select c.SlpName as asesor,(x.DocTotalSy-x.VatSumSy) as 'subtotal',x.DocTotalSy,x.VatSumSy as IVA,x.*,y.*,case when U_FECHALMACEN is not null THEN 1 else 0 end as 'flag',case when x.U_CFDi_SolCancela=0 then 'activo' else 'cancelado' end as 'timbre',round(y.Quantity*y.price,2) as SubPartida,x.CANCELED from OINV x inner join INV1 y on x.docentry=y.docentry left join OSLP c on x.SlpCode=c.slpcode where x.DocNum=" &request("ffactura") &" order by y.lineNum"
         'response.write strSQL

         select case request("fRS")
             case "DMX"  rsUpdateEntry.Open strSQL, adoCon2 'DMX'
             case "DFW"  rsUpdateEntry.Open strSQL, adoCon3 'DFW'
             case "MME"  rsUpdateEntry.Open strSQL, adoCon5 'MME'
         end select
      
         CreateTable(1200)
         rowin
         response.write("<td class='td-c'>")

         if not rsUpdateEntry.EOF then

         response.write("<table border=1 width='400px' cellpadding=3 cellspacing=2>")        
         
         response.write("<tr><td class='subtitulo td-r' width='100px'>Factura:</td><td class='td-l'>"& request("fRS") &"-" &request("ffactura") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Socio negocio:</td><td class='td-l'>"& rsUpdateEntry("cardname") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Fecha:</td><td class='td-l'>"& rsUpdateEntry("DocDate") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Referencia:</td><td class='td-l'>"& rsUpdateEntry("numAtCard") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Proyecto:</td><td class='td-l'>"& rsUpdateEntry("Project") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Peso Total:</td><td class='td-l'>"& rsUpdateEntry("U_PESO_GLOBAL") &" Kg</td></tr>")
          response.write("<tr><td class='subtitulo td-r'>Seguimiento:</td><td class='td-l'>"& rsUpdateEntry("U_COMINTERNO") &"</td></tr>")

         response.write("</table></td><td class='td-c'>")
         response.write("<table border=1 width='500px' cellpadding=3 cellspacing=2>")
               response.write("<tr><td class='subtitulo td-r'>Comentarios:</td><td class='td-l'>"& mid(rsUpdateEntry("Comments"),1,38)&"<br>" & mid(rsUpdateEntry("Comments"),39,200) &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Subtotal:</td><td class='td-l'>"&  rsUpdateEntry("subTotal") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>IVA:</td><td class='td-l'>"& rsUpdateEntry("IVA") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Total:</td><td class='td-l'>"& rsUpdateEntry("DocTotalSy") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Asesor:</td><td class='td-l'>"& rsUpdateEntry("asesor") &"</td></tr>")
         response.write("</table></td> ")      

         response.write("<td class='td-c'>")
         response.write("<table border=1 width='300px' cellpadding=3 cellspacing=2>")
               response.write("<tr><td class='subtitulo td-r' width='50%'>Timbre:</td><td class='td-l' width='50%'>"& rsUpdateEntry("timbre") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>UsoCFDi:</td><td class='td-l'>"&  rsUpdateEntry("U_CFDi_UsoCFDi") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Metodo Pago:</td><td class='td-l'>"& rsUpdateEntry("U_CFDi_MetodoPago") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Forma de pago:</td><td class='td-l'>"& rsUpdateEntry("U_CFDi_FormaPago") &"</td></tr>")

                 response.write("<tr><td class='subtitulo td-r'>Estatus:</td>")
                if rsUpdateEntry("CANCELED")="Y" then
              response.write("<td class='td-l' style='background-color:red;color:white'>FACT CANCELADA</td></tr>")
                else
              response.write("<td class='td-l'>NO CANCELADO</td></tr>")
                end if
         response.write("</table></td></tr> ")      


         else
               response.write("<tr><td colspan=8>no existen facturas con esos criterios</td></tr>")

         end if


         response.write("</table>")
         response.write("<P>&nbsp;</P>")
         CreateTable(870)


         response.write("<tr><th class='subtitulo td-c'>#</th><th class='subtitulo td-c'>Codigo</th><th class='subtitulo td-c'>Descripcion</th><th class='subtitulo td-c'>Cantidad</th><th class='subtitulo td-c'>Precio</th><th class='subtitulo td-c'>Subtotal</th><th class='subtitulo td-c'>Almacen <br>cotizado</th><th class='subtitulo td-c'>F embarque</th></tr>")
          
         n=1 
         while not rsUpdateEntry.EOF
             if n=1 then
                  vString=" style='background-color: white;' "
             else
                  vString=" style='background-color: #E1DFDF;' "
             end if
             response.write("<tr>")
             response.write("<td class='td-c' "& vString &">" & rsUpdateEntry("lineNum")+1 &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & rsUpdateEntry("ItemCode") &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & left(rsUpdateEntry("Dscription"),50) &"</td>")
             
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("Quantity") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("Price") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("SubPartida") &"</td>")
          
             response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("WhsCode") &"</td>")
             response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("Shipdate") &"</td>")

             if rsUpdateEntry("flag")=1 then
                  response.write("<td class='td-c fontmed alert'>" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
             else
                  response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
             end if

             response.write("</tr>")
             rsUpdateEntry.MoveNext
             n=n+1
             if n=3 then  n=1
   
         wend
         rsUpdateEntry.close
         closeTable

    end if

end sub

sub ChangeStatusPO 'contenido 120'

    doAlert
    response.write("<P>&nbsp;</P>")
    titulo="Cambiar Ubicacion material en todas partidas de una PO"
    DoTitulo(titulo)


    if request("action")="" then

         response.write("<form id='PO' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=120>")
         response.write("<input type='hidden' name='action' value=1>")

         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;# PO: <input class='anchoSmall' type='numeric' name='fPO' value="&request("fPO") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

    end if

    if request("action")="1" then
         response.write("<form id='PO' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=120>")
         response.write("<input type='hidden' name='action' value=2>")
         response.write("<input type='hidden' name='RS' value="&request("fRS")&">")
         

         strSQL="select * from OPOR where Docnum="&request("fPO")
         strSQL2="select * from [@ESTATUS_COMPRA]"
         select case request("fRS")
             case "DMX"  rsUpdateEntry.Open strSQL, adoCon2 'DMX'
                         rsUpdateEntry2.Open strSQL2, adoCon2 'DMX'
             case "DFW"  rsUpdateEntry.Open strSQL, adoCon3 'DFW'
                         rsUpdateEntry2.Open strSQL2, adoCon3 'DFW'
             case "MME"  rsUpdateEntry.Open strSQL, adoCon5 'MME'
                         rsUpdateEntry2.Open strSQL2, adoCon5 'MME'
         end select
      
         if  rsUpdateEntry.EOF then response.redirect("ShowContent.asp?contenido=120&msg=PO No. "&request("fPO") &" no existe")

          response.write("<input type='hidden' name='DocEntry' value="& rsUpdateEntry("DocEntry")  &">")

          CreateTable(400)
          RowIn
            response.write("<td class='td-r subtitulo' width='35%'>RS</td><td class='td-l'>" & request("fRS") &"</td>")
          rowout
          RowIn
            response.write("<td class='td-r subtitulo'>Cardcode:</td><td class='td-l'>" & rsUpdateEntry("CardCode") &"</td>")
          rowout
          RowIn
            response.write("<td class='td-r subtitulo'>Fabricante:</td><td class='td-l'>" & rsUpdateEntry("CardName") &"</td>")
          rowout
          RowIn
            response.write("<td class='td-r subtitulo'>Fecha de la PO:</td><td class='td-l'>" & rsUpdateEntry("DocDate") &"</td>")
          rowout
          
          RowIn
            response.write("<td class='td-r subtitulo'>Cambiar estatus a:</td><td class='td-l'><select name='festatus'>")
            while not rsUpdateEntry2.EOF
                response.write("<option value='"&rsUpdateEntry2("code") &"'>"& rsUpdateEntry2("Name")  &"</option>")
                rsUpdateEntry2.movenext
            wend
            rsUpdateEntry2.close
            response.write("</select></td>")
          rowout
          RowIn
                 response.write("<td colspan=2>&nbsp;</td>")
          rowout
          RowIn
                 response.write("<td colspan=2 class='td-c'><input type='submit' value='actualizar'></form></td>")
          rowout
          rsUpdateEntry.close

    end if

     if request("action")="2" then
          strSQL="UPDATE POR1 SET U_UBICACIONMATERIAL='"&request("festatus")&"' where DocEntry="&request("DocEntry")
         

          select case request("RS")
             case "DMX"   rsUpdateEntry.Open strSQL, adoCon2 'DMX' 
                          response.write("DMX: "& strSQL  )
             case "DFW"   rsUpdateEntry.Open strSQL, adoCon3 'DFW'   
                          response.write("DFW: "& strSQL  )                     
             case "MME"   rsUpdateEntry.Open strSQL, adoCon5 'MME'    
                          response.write("MME: "& strSQL  )                     
         end select

         response.redirect("ShowContent.asp?contenido=120&msg=La PO en cuestion ha sido actualizada")

     end if

end sub


sub TranfStock   'contenido 121'

    doAlert
    response.write("<P>&nbsp;</P>")
    titulo="Revisa el valor mercancia por documentos de SAP"
    DoTitulo(titulo)


    if request("action")="" then

         response.write("<form id='transfer' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=121>")
         response.write("<input type='hidden' name='action' value=1>")

         CreateTable(400)
         RowIn
         response.write("<td class='td-r subtitulo' width='50%'>Raz&oacute;n Social:</td>")
         response.write("<td class='td-l' width='50%'><select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option></select></td>")
         RowOut

         RowIn
         response.write("<td class='td-r subtitulo'>Tipo de documento:</td>")
         response.write("<td class='td-l'><select name='ftipo'>")
            response.write("<option value='ODLN'>remision</option>")
            response.write("<option value='OINV'>facturas</option>")
            response.write("<option value='OWTR'>trans</option>")
            response.write("</select></td>")
         RowOut

         for i=1 to 20 
            RowIn
                response.write("<td class='td-r subtitulo'># doc SAP "&i&"</td>")
                response.write("<td class='td-l'><input class='anchoSmall' type='numeric' name='doc"&i&"'></td>")
            RowOut
         next
 
         response.write("<tr><td colspan=2>&nbsp;</td></tr><tr><td colspan=2 class='td-c'><input type='submit' value='continuar'></form>")
         closeTable

    end if

    n=1
    if request("action")="1" then
         dim transfer(20)
         for i=1 to 20
               vstring="doc"&i
               if trim(request(vString))<>"" then
                   n=n+1
                   transfer(i)=trim(request(vString))
               end if
         next
         vstring=""
         for i=1 to N
             vString=vString & transfer(i) &","
         next
         vString=replace(vString,",,","")    
         vtipo=mid(request("ftipo"),2,4) &"1"  

         if request("ftipo")="OWTR" then         

            strSQL="SELECT z.DocNum,z.DocCur,CASE when AUX.Total<0 then -1*AUX.Total else AUX.Total end as Total FROM   "& request("ftipo") &" z INNER JOIN ( SELECT  a.DocNum,sum( round(b.Quantity*b.StockPrice,2) ) as Total  FROM "&request("ftipo")&" a inner join "&vtipo&" b on a.DocEntry=b.DocEntry    WHERE a.DocNum in ("& vString &")     GROUP BY a.DocNum ) as AUX on z.DocNum=AUX.DocNum WHERE z.DocNum in ("& vString&") "
         else
            strSQL="SELECT  a.DocNum,a.DocCur,a.DocTotalFc as Total  FROM "&request("ftipo")&" a  WHERE a.DocNum in ("& vString &")  "

       end if
         'response.write strSQL

         select case request("fRS")
            case "DMX"  rsUpdateEntry.Open strSQL, adoCon2    'DMX'
            case "DFW"  rsUpdateEntry.Open strSQL, adoCon3
            case "MME"  rsUpdateEntry.Open strSQL, adoCon5
          end select 

          if rsUpdateEntry.EOF then NoRegistros
          CreateTable(500)
          rowin
          response.write("<td class='td-c subtitulo'># doc</td><td class='td-c subtitulo'>Moneda</td><td class='td-c subtitulo'>Total</td>")
          RowOut
          strSQL="select sum("
          flag=0

          while not rsUpdateEntry.EOF
              RowIn
              response.write("<td class='td-r'>"& rsUpdateEntry("DocNum") &"</td>")
              response.write("<td class='td-r'>"& rsUpdateEntry("DocCur") &"</td>")
              response.write("<td class='td-r'>"& FormatCurrency(rsUpdateEntry("Total")) &"</td>")              
              RowOut
              strSQL=strSQL&"+"&rsUpdateEntry("Total")
              rsUpdateEntry.MoveNext
              flag=1
          wend
          rsUpdateEntry.close
          closeTable

          if flag=1 and request("ftipo")="OWTR" then
              strSQL=strSQL&") as calculo,(select Rate from PRODUCTIVA_DMX.dbo.ORTT where RateDate=cast(GETDATE() as date) ) as Rate"
              rsUpdateEntry2.Open strSQL, adoCon4
              response.write("<P class='td-c'>El valor total del traslado es: <B>"& FormatCurrency(rsUpdateEntry2("calculo")) &" MXN</B></P>")
              

              strSQL="select round(cast("& rsUpdateEntry2("calculo") &" as float)/"& rsUpdateEntry2("Rate") &",2) as calculo"
              rsUpdateEntry3.Open strSQL, adoCon4
              response.write("<P class='td-c'>El valor total del traslado en (TC "& rsUpdateEntry2("Rate") &") es: <B>" & FormatCurrency(rsUpdateEntry3("calculo"))  &" USD</B></P>") 
              rsUpdateEntry2.close
              rsUpdateEntry3.close

          end if

          if  flag=1 and request("ftipo")<>"OWTR" then
              strSQL=strSQL&") as calculo"
              rsUpdateEntry3.Open strSQL, adoCon4
              response.write("<P class='td-c'>El valor total del traslado es: <B>" & FormatCurrency(rsUpdateEntry3("calculo"))  &" USD</B></P>")               
              rsUpdateEntry3.close
          end if

          response.write("<P class='td-c'>L&iacute;mite de cobertura m&aacute;xima por Embarque (Zurich: 110934497): <font color=red>90,000 USD </font></B></P>")
              response.write("<P class='td-c'>L&iacute;mite de peso m&aacute;ximo por Embarque (Zurich: 110934497): <font color=red>25 toneladas </font></B></P>") 

    end if

end sub



sub ViaticoEdit   'contenido 122'
    doAlert
    response.write("<P>&nbsp;</P>")
    titulo="Agregue una remision a un viatico"
    DoTitulo(titulo)


    if request("action")="" then

         response.write("<form id='transfer' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=122>")
         response.write("<input type='hidden' name='action' value=1>")

         CreateTable(700)
         RowIn
         response.write("<td class='td-r subtitulo' width='25%'>No. de viatico:</td>")
         response.write("<td class='td-l' width='75%'><select name='fID'>")

         strSQL="select * from VIATICOS_R1 where Helpdesk<4 and helpdesk!=2 order by ID desc"
         rsUpdateEntry.Open strSQL, adoCon
         while not rsUpdateEntry.EOF              
              response.write("<option value='"&rsUpdateEntry("ID")&"'>"& rsUpdateEntry("ID")&"-" & rsUpdateEntry("notas") &"</option>")
              rsUpdateEntry.MoveNext
          wend
          rsUpdateEntry.close
          response.write("</select></td></tr>")
   
          response.write("<tr><td colspan=2>&nbsp;</td></tr><tr><td colspan=2 class='td-c'><input type='submit' value='continuar'></form>")
          closeTable

    end if

    if request("action")="1" then

         strSQL="select a.*,cardname=(select top 1 cardname from SBOTemp.dbo.Notificaciones where modulo='ventas' and tipo='pedido' and RS=a.RS and DocNum=a.Pedido) from ViaticoRemision a where a.ID="&request("fID")
         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon

         response.write("<form id='transfer' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=122>")
         response.write("<input type='hidden' name='action' value=2>")
         response.write("<input type='hidden' name='ID' value="& request("fID") &">")
         response.write("<input type='hidden' name='RS' value="& rsUpdateEntry("RS") &">")

         CreateTable(570)
         RowIn
         response.write("<td class='td-r subtitulo' width='25%'>Viatico:</td>")
         response.write("<td class='td-l' width='75%'>"& request("fID") &"</td>")
         RowOut
         RowIn
         response.write("<td class='td-r subtitulo'>Razon Social:</td>")
         response.write("<td class='td-l'>"& rsUpdateEntry("RS") &"</td>")
         RowOut
         RowIn
         response.write("<td class='td-r subtitulo'>Socio Negocio:</td>")
         response.write("<td class='td-l'>"& rsUpdateEntry("cardName") &"</td>")
         RowOut
         RowIn
         response.write("<td class='td-r subtitulo'>Remisiones <br>agregadas:</td>")
         response.write("<td class='td-l'>")
             while not rsUpdateEntry.EOF
                 response.write("remision: "& rsUpdateEntry("remision") &" | Pedido: "&  rsUpdateEntry("Pedido") &"<br>")
                 rsUpdateEntry.MoveNext
             wend             
             response.write("</td>")
         RowOut
         RowIn
         response.write("<td class='td-r subtitulo'>Agregar remision:</td>")
         response.write("<td class='td-l'><input type='number' name='fremision'></td>")
         RowOut
         RowIn
         response.write("<td class='td-r subtitulo'>&nbsp;</td>")
         response.write("<td class='td-l'><input type='submit' value='agregar'></form></td>")
         RowOut

         rsUpdateEntry.close
         closeTable
    end if


    if request("action")="2" then

         strSQL="select * from Notificaciones where modulo='ventas' and RazonSocial='"&request("RS")&"' and tipo='remision' and Docnum="&request("fremision")
         response.write   strSQL       
         rsUpdateEntry.Open strSQL, adoCon4

         if rsUpdateEntry.EOF then response.redirect("ShowContent.asp?contenido=122&msg=remision no existe")
         '9234         

         strSQL="select * from ViaticoRemision where remision="&request("fremision") &" and RS='"&request("RS")&"' "
         rsUpdateEntry2.Open strSQL, adoCon

         if not rsUpdateEntry2.EOF then response.redirect("ShowContent.asp?contenido=122&msg=remision ya se encuentra agregada en un viatico")
   
         strSQL="insert into ViaticoRemision (ID,Pedido,Remision,RS) values ("&request("ID")&","&rsUpdateEntry("Pedido")&","&request("fremision")&",'"&request("RS")&"')"
         rsUpdateEntry.close
         rsUpdateEntry2.close
         rsUpdateEntry7.Open strSQL, adoCon

         strSQL="UPDATE VIATICOS_R1 set Motivo=(select dbo.fn_EnlistaRemisiones("&request("ID")&") ) where ID="&request("ID")
         rsUpdateEntry5.Open strSQL, adoCon

         response.redirect("ShowContent.asp?contenido=122&msg=la remision ha sido agregada al viatico")

    end if

end sub


sub EstatuSuministro  'contenido 123'

    doAlert
    response.write("<P>&nbsp;</P>")

  if request("action")="" then

         titulo="Estatus del suministro"
         DoTitulo(titulo)

         response.write("<form id='suministro' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=123>")
         response.write("<input type='hidden' name='action' value=1>")

         strSQL="SELECT * FROM (SELECT 'DMX' as RS1 UNION  SELECT 'DFW' as RS1) as AUX1 inner join (  SELECT 'DFW' as RS,Cardcode,CardName FROM PRODUCTIVA_DFW.dbo.ORDR WHERE substring(CardCode,1,1)='C' and CardName not like '%DELTAFLOW%' AND CANCELED='N' and DocStatus='O' group by CardCode,CardName   UNION   SELECT 'DFW' as RS,Cardcode,CardName FROM PRODUCTIVA_DFW.dbo.ODLN WHERE substring(CardCode,1,1)='C' and CardName not like '%DELTAFLOW%' AND CANCELED='N' and DocStatus='O' group by CardCode,CardName   UNION   SELECT 'DMX' as RS,Cardcode,CardName FROM PRODUCTIVA_DMX.dbo.ORDR WHERE substring(CardCode,1,1)='C' and CardName not like '%DELTAFLOW%' AND CANCELED='N' and DocStatus='O' group by CardCode,CardName   UNION   SELECT 'DMX' as RS,Cardcode,CardName FROM PRODUCTIVA_DMX.dbo.ODLN WHERE substring(CardCode,1,1)='C' and CardName not like '%DELTAFLOW%' AND CANCELED='N' and DocStatus='O' group by CardCode,CardName   ) AS AUX2 on AUX1.RS1=AUX2.RS group by AUX1.RS1,AUX2.RS,AUX2.CardCode,AUX2.CardName order by AUX1.RS1,AUX2.RS,AUX2.CardCode,AUX2.CardName  "

          rsUpdateEntry.Open strSQL, adoCon4

         response.write("<P class='td-c'>Elija un Socio de Negocio: <select name='fSN'>")
         while not rsUpdateEntry.EOF
              response.write(" <option value='"&rsUpdateEntry("RS")&"*"&rsUpdateEntry("Cardcode")&"'>"&rsUpdateEntry("RS")&"-"&rsUpdateEntry("cardName")& "</option>")
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close
         response.write("</select></P>")
         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

  end if



  if request("action")="1" then
        vString=request("fSN")
        vPos=inStr(vString,"*")
        vRS=mid(vString,1,Vpos-1)
        vCardCode=mid(vString,Vpos+1,30)

    strSQL="select * from OCRD where cardcode='"&vCardCode&"' "
    'response.write strSQL
    select case vRS
            case "DMX"  rsUpdateEntry7.Open strSQL, adoCon2    'DMX'
            case "DFW"  rsUpdateEntry7.Open strSQL, adoCon3
            case "MME"  rsUpdateEntry7.Open strSQL, adoCon5
          end select 

    titulo="Estatus del suministro SN: "& rsUpdateEntry7("CardName") &"<br><font class='fonttiny'>(renglones en grises, es remisionado no facturado)</font>"
    DoTitulo(titulo)
    rsUpdateEntry7.close
 
    strSQL="Exec [SP_EstatuSuministro] '"&vRS&"','"&vCardCode&"'"
    'response.write strSQL
    rsUpdateEntry6.Open strSQL, adoCon4


    'NUMERO DE DIFERENTES PROYECTOS Y SUS NOMBRES'
    strSQL="select distinct(U_PROYECTO) as Proyecto from _suministro"    
    'response.write strSQL     
    rsUpdateEntry.Open strSQL, adoCon4

    Dim Proyectos(50)  
    Dim Pedidos(50)  
    Dim Destinos(50)

    i=0
    while not rsUpdateEntry.EOF
      i=i+1
      Proyectos(i)=rsUpdateEntry("Proyecto")
      rsUpdateEntry.MoveNext
    wend
    rsUpdateEntry.Close
    NumProyectos=i

    response.write("<table border=1 align=center class='table2excel table2excel_with_colors' data-tableName='table1'>")  'TABLA PRINCIPAL'  

    for n=1 to i 'hasta el numero de proyectos'
      
        strSQL="select distinct(Pedido) as Pedido,U_Destino from _suministro where U_PROYECTO='"& Proyectos(n) &"'" 
        'response.write strSQL     
        rsUpdateEntry2.Open strSQL, adoCon4

        i=0
        while not rsUpdateEntry2.EOF
           i=i+1
           Pedidos(i)=rsUpdateEntry2("Pedido")
           Destinos(i)=rsUpdateEntry2("U_Destino")
           rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.Close
        NumPedidos=i


       
        for P=1 to NumPedidos
              
              response.write("<tr><td colspan=20 class=' subtitulo strong fontmed' style='vertical-align:top'>Pedido: "& Pedidos(P) & " | Destino: "& Destinos(P) &"<BR>")

              strSQL="select * from _suministro where Pedido='"& Pedidos(P) &"' and Remision=0 order by Line"
              'response.write(strSQL&"<br>")
              rsUpdateEntry3.Open strSQL, adoCon4
 
               titulosSuministros
    
               while not rsUpdateEntry3.EOF                 
                   RowIn                   
                   response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("Pedido") &"</td>")
                   response.write("<td class='td-l fonttiny'>-</td>")
                   response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("Referencia") &"&nbsp;&nbsp;</td>")                
                   response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("Shipdate") &"</td>")
                   response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("Line") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("ItemCode") &"</td>")
                   response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("Descripcion") &"</td>")
                   response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("Peso") &"</td>")
                   response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("WhsCode") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("Quantity") &"</td>")                   
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("MXL") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("SJR") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("MTY") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("TR") &"</td>")

                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("RESG-MXL") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("RESG-SJR") &"</td>")
                   response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("RESG-MTY") &"</td>")
                
                   if int(trim(rsUpdateEntry3("stock"))) then
                      response.write("<td class='td-r fonttiny' style='background-color:#C4DAC4'>"& rsUpdateEntry3("stock") &"</td>")
                   else
                      response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("stock") &"</td>")
                   end if
                
                   response.write("<td class='td-r fonttiny'>"& Formatcurrency(rsUpdateEntry3("total")) &"</td>")
                   RowOut
                  rsUpdateEntry3.MoveNext                                                                      
               wend
               rsUpdateEntry3.close
              
               strSQL="select * from _suministro where Pedido='"& Pedidos(P) &"' and Remision>0 order by Line"
               'response.write(strSQL&"<br>")
               rsUpdateEntry3.Open strSQL, adoCon4

               while not rsUpdateEntry3.EOF                 
                   RowIn                   
                   response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("Pedido") &"</td>")
                   response.write("<td class='td-l fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Remision") &"</td>")
                   response.write("<td class='td-l fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Referencia") &"&nbsp;&nbsp;</td>")                
                   response.write("<td class='td-c fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Shipdate") &"</td>")
                   response.write("<td class='td-c fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Line") &"</td>")
                   response.write("<td class='td-r fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("ItemCode") &"</td>")
                   response.write("<td class='td-l fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Descripcion") &"</td>")
                   response.write("<td class='td-l fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Peso") &"</td>")
                   response.write("<td class='td-c fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("WhsCode") &"</td>")
                   response.write("<td class='td-r fonttiny' style='background-color:#DFDBDB'>"& rsUpdateEntry3("Quantity") &"</td>") 
                   response.write("<td class='td-r fonttiny' style='background-color:#DFDBDB' colspan=8>&nbsp;</td>")
                   response.write("<td class='td-r fonttiny' style='background-color:#DFDBDB'>"& Formatcurrency(rsUpdateEntry3("total")) &"</td>") 
                   RowOut
                  rsUpdateEntry3.MoveNext                                                                      
               wend
               rsUpdateEntry3.close
         

        next
        RowOut

    next

     %>

   </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp;    

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>
    <P>&nbsp;</P><P>&nbsp;</P>
    <%

  end if

end sub

sub titulosSuministros
    rowin
      response.write("<td subtitulo class='td-l fonttiny strong'>PED</td>")     
      response.write("<td subtitulo class='td-l fonttiny strong'>REM</td>") 
      response.write("<td subtitulo class='td-l fonttiny strong'>Referencia</td>")   
      response.write("<td subtitulo class='td-c fonttiny strong'>F Embrqe</td>")
       response.write("<td subtitulo class='td-c fonttiny strong'>#</td>")
      response.write("<td subtitulo class='td-c fonttiny strong'>Codigo</td>")
      response.write("<td subtitulo class='td-c fonttiny strong'>Descripcion</td>")
      response.write("<td subtitulo class='td-c fonttiny strong'>Peso Unit</td>")
      response.write("<td subtitulo class='td-c fonttiny strong'>ALM</td>")
      response.write("<td subtitulo class='td-r fonttiny strong'>Pdte</td>")
      

      response.write("<td subtitulo class='td-r fonttiny strong'>MXL</td>")
      response.write("<td subtitulo class='td-r fonttiny strong'>SJR</td>")
      response.write("<td subtitulo class='td-r fonttiny strong'>MTY</td>")
      response.write("<td subtitulo class='td-r fonttiny strong'>TR</td>")

      response.write("<td subtitulo class='td-r fonttiny strong'>RESG-MXL</td>")
      response.write("<td subtitulo class='td-r fonttiny strong'>RESG-SJR</td>")
      response.write("<td subtitulo class='td-r fonttiny strong'>RESG-MTY</td>")

      response.write("<td subtitulo class='td-r fonttiny strong'>T.Stock</td>")

      response.write("<td subtitulo class='td-r fonttiny strong'>Importe</td>")
    RowOut
end sub


sub RemisionDetalle  'contenido 124

    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
    end if

    


    if request("action")="" then
         titulo="Mostrar el detalle de una remision"
         DoTitulo(titulo)

         response.write("<form id='inv' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=124>")
         response.write("<input type='hidden' name='action' value=1>")


         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option><option value='MEIDE'>MME</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Pedido: <input class='anchoSmall' type='numeric' name='fremision' value="&request("fremision") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

    else
         titulo="Mostrar el detalle de la remision: "&request("fremision")
          DoTitulo(titulo)
    
         strSQL="select * from Notificaciones where RazonSocial='"& request("fRS") &"' and tipo='Remision' and docnum="&request("fremision")        
         rsUpdateEntry5.Open strSQL, adoCon4
         vPedido=trim(rsUpdateEntry5("Pedido"))
         rsUpdateEntry5.close
         
         strSQL="select c.SlpName as asesor,x.numAtCard,x.U_COMINTERNO,x.U_PESO_GLOBAL,x.Project,x.docnum,x.docdate,x.cardcode,x.cardname,x.Comments,(x.DocTotalSy-x.VatSumSy) as 'subtotal',x.DocTotalSy,x.VatSumSy as IVA,y.*,case when U_FECHALMACEN is not null THEN 1 else 0 end as 'flag','R:\"&request("fRS")&"\'+R.Ruta+'\'+convert(varchar,"&vPedido&") as repositorio,x.CANCELED from ODLN x inner join DLN1 y on x.docentry=y.docentry left join OSLP c on x.SlpCode=c.slpcode left join SBOTemp.dbo.Repositorio R on R.RazonSocial='"& request("fRS") &"' and x.cardcode=R.CardCode where x.DocNum=" &request("fremision") &" order by y.lineNum"
       
         'response.write strSQL

         if request("fRS")="DMX" then
                 rsUpdateEntry.Open strSQL, adoCon2 'DMX'
         else
                 rsUpdateEntry.Open strSQL, adoCon3 'DFW'
         end if
      
         CreateTable(900)
         rowin
         response.write("<td class='td-c' width='50%'>")
           

        if not rsUpdateEntry.EOF then

         response.write("<table border=1 width='430px' cellpadding=3 cellspacing=2>")        
         
         response.write("<tr><td class='subtitulo td-r' width='100px'>Pedido:</td><td class='td-l'>"& request("fRS") &"-" &vPedido &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Socio negocio:</td><td class='td-l'>"& rsUpdateEntry("cardname") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Fecha:</td><td class='td-l'>"& rsUpdateEntry("DocDate") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Referencia:</td><td class='td-l'>"& rsUpdateEntry("numAtCard") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Proyecto:</td><td class='td-l'>"& rsUpdateEntry("Project") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Peso Total:</td><td class='td-l'>"& rsUpdateEntry("U_PESO_GLOBAL") &" Kg</td></tr>")
          response.write("<tr><td class='subtitulo td-r'>Seguimiento:</td><td class='td-l'>"& rsUpdateEntry("U_COMINTERNO") &"</td></tr>")
         response.write("</table></td><td class='td-c' width='50%'>")
         response.write("<table border=1 width='430px' cellpadding=3 cellspacing=2>")
               response.write("<tr><td class='subtitulo td-r'>Comentarios:</td><td class='td-l'>"& rsUpdateEntry("Comments") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Subtotal:</td><td class='td-l'>"&  rsUpdateEntry("subTotal") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>IVA:</td><td class='td-l'>"& rsUpdateEntry("IVA") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Total:</td><td class='td-l'>"& rsUpdateEntry("DocTotalSy") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Asesor:</td><td class='td-l'>"& rsUpdateEntry("asesor") &"</td></tr>")
              
               response.write("<tr><td class='subtitulo td-r'>Repositorio:</td><td class='td-l'>"& rsUpdateEntry("repositorio") &"</td></tr>")

                response.write("<tr><td class='subtitulo td-r'>Estatus:</td>")
                if rsUpdateEntry("CANCELED")="Y" then
              response.write("<td class='td-l' style='background-color:red;color:white'>REM CANCELADA</td></tr>")
                else
              response.write("<td class='td-l'>NO CANCELADO</td></tr>")
                end if

                strSQL="select a.*,b.portador from envios a left join cat_portadores b on a.id_portador=b.id_portador where a.Docnum="&rsUpdateEntry("docnum")
                rsUpdateEntry5.Open strSQL, adoCon4

                if not rsUpdateEntry5.EOF then

                  response.write("<tr><td class='subtitulo td-r'>Envio(s):</td><td class='td-l'>")
                  while not rsUpdateEntry5.EOF
                      response.write( rsUpdateEntry5("Portador")  & " | rastreo: " & rsUpdateEntry5("rastreo") &    "<img src='/images/tracking.png' width='30px' border='0' title='rastree'></td></tr>")
                      rsUpdateEntry5.MoveNext
                      if not rsUpdateEntry5.EOF then response.write("<BR>")
                  wend
                  

                end if
                rsUpdateEntry5.close
              
         response.write("</table></td></tr> ")      

         else
               response.write("<tr><td colspan=8>no existen pedidos con esos criterios</td></tr>")

         end if

         
               


       
         response.write("</table>")
         response.write("<P>&nbsp;</P>")
         CreateTable(1200)

          
         response.write("<tr><th class='subtitulo td-c'>#</td><th class='subtitulo td-c'>Codigo</td><th class='subtitulo td-c'>Descripcion</td><th class='subtitulo td-c'>Estatus</td><th class='subtitulo td-c'>Cantidad</td><th class='subtitulo td-c'>Almacen <br>cotizado</td><th class='subtitulo td-c'>F embarque</td><th class='subtitulo td-c'>actualizado</td></tr>")
          
         n=1 
         while not rsUpdateEntry.EOF
             if n=1 then
                  vString=" style='background-color: white;' "
             else
                  vString=" style='background-color: #E1DFDF;' "
             end if
             response.write("<tr>")
             response.write("<td class='td-c' "& vString &">" & rsUpdateEntry("lineNum")+1 &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & rsUpdateEntry("ItemCode") &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & left(rsUpdateEntry("Dscription"),80) &"</td>")
             if trim(rsUpdateEntry("LineStatus"))="O" then
                  response.write("<td class='td-c fontmed' "& vString &">sin facturar</td>")
             else
                  response.write("<td class='td-c fontmed' "& vString &">facturada</td>")
             end if
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("Quantity") &"</td>")

            
             response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("WhsCode") &"</td>")
             response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("Shipdate") &"</td>")

             if rsUpdateEntry("flag")=1 then
                  response.write("<td class='td-c fontmed alert'>" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
             else
                  response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
             end if

             

             response.write("</tr>")
             rsUpdateEntry.MoveNext
             n=n+1
             if n=3 then
                 n=1
             end if
         wend
         rsUpdateEntry.close
         closeTable

    end if

end sub


sub EstatusProveedores   'contenido 125'
    
  doAlert
  response.write("<P>&nbsp;</P>")
  titulo="<a href='ShowContent.asp?contenido=125'>Estatus del Cr&eacute;dito con Fabricantes</a>"
  DoTitulo(titulo)

  if session("Fabricantes")<>"on" then
      strSQL="exec SP_Proveedores_intranet"  
      rsUpdateEntry6.Open strSQL, adoCon4
      session("Fabricantes")="on"
  end if
  

  strSQL="select  * from _TEMP_PROV2 "
  if request("Cardcode")<>"" then   strSQL= strSQL&"where CardCode='"&request("Cardcode")&"' "
  strSQL= strSQL&" order by Total Desc"
  'response.write strSQL
  rsUpdateEntry.Open strSQL, adoCon4

  strSQL="select sum(Total) as Total,CONVERT(VARCHAR,CONVERT(MONEY,sum(Total)),1) as Total2 from _TEMP_PROV2"
  rsUpdateEntry2.Open strSQL, adoCon4
  
  CreateTable(900)
  rowin
  
  response.write("<th class='subtitulo td-c'>Fabricante</th>")
  response.write("<th class='subtitulo td-c'>Dias Credito</th>")
  response.write("<th class='subtitulo td-c'>Limite</th>")
  response.write("<th class='subtitulo td-c'>Pdte Suministro</th>")
  response.write("<th class='subtitulo td-c'>Vencido</th>")
  response.write("<th class='subtitulo td-c'>Corriente</th>")
  response.write("<th class='subtitulo td-c'>Total</th>")
  RowOut

  while not rsUpdateEntry.EOF
       rowin       
       response.write("<td class='td-l'><a href='ShowContent.asp?contenido=125&cardcode="&rsUpdateEntry("cardCode")&"&cardname="&rsUpdateEntry("cardName")&"'>" & rsUpdateEntry("CardName") &"</a></td>")
       response.write("<td class='td-c'>" & rsUpdateEntry("ExtraDays") &"</td>")
       response.write("<td class='td-c'>" & Formatcurrency(rsUpdateEntry("Limite")) &"</td>")
       response.write("<td class='td-r'>" & Formatcurrency(rsUpdateEntry("Suministrar")) &"</td>")
       response.write("<td class='td-r'>" & Formatcurrency(rsUpdateEntry("Vencido")) &"</td>")
       response.write("<td class='td-r'>" & Formatcurrency(rsUpdateEntry("Corriente")) &"</td>")
       response.write("<td class='td-r'>" & Formatcurrency(rsUpdateEntry("Total")) &"</td>")
       rowout
       rsUpdateEntry.MoveNext
  wend
  rsUpdateEntry.close
  if request("Cardcode")="" then
     separador
     rowin
     response.write("<td class='td-r strong' colspan=6>TOTAL</td><td class='td-r strong'>" & rsUpdateEntry2("total2") &"</td>")
     rowout
  end if
  rsUpdateEntry2.close
  closeTable()


  if request("Cardcode")<>"" then
   
    titulo="Estatus con el Fabricante: " & request("CardName")
    doTitulo(titulo)

    CreateTable(1200)  'tabla total'
       response.write("<tr><td width='50%' class='td-c' style='vertical-align:top'>")

       CreateTable(590)
       RowIn
       response.write("<td colspan=3 class='td-c strong subtitulo'>PENDIENTE POR SUMINISTRAR</td></tr>")
       RowOut

       strSQL="select distinct(CardCode) as Cardcode from _TEMP_PROV3  where U_CardCode='"&request("Cardcode")&"' "
       rsUpdateEntry3.Open strSQL, adoCon4

      RowIn
        response.write("<td class='td-c subtitulo'>Socio de negocio</td>")
        response.write("<td class='td-c subtitulo'>Pedidos de venta</td>")
        response.write("<td class='td-c subtitulo'>Subtotal (a Costo Venta)</td>")
      RowOut
     
      while not rsUpdateEntry3.EOF
        strSQL="select CardName,Pedidos=( select convert(varchar,DocNum)+'  ' from _TEMP_PROV3 where CardCode='"& rsUpdateEntry3("cardCode")&"' and U_CardCode='"&request("Cardcode")&"' group by DocNum FOR XML PATH ('')  )   ,sum(subtotal) as Subtotal from _TEMP_PROV3  where CardCode='"& rsUpdateEntry3("cardCode")&"' and U_CardCode='"&request("Cardcode")&"' group by CardName"       
        rsUpdateEntry4.Open strSQL, adoCon4

        RowIn
        'response.write("<td class='td-l fonttiny' colspan=3>"& strSQL &"</td>")
        response.write("<td class='td-l fonttiny'>"& rsUpdateEntry4("cardname") &"</td>")
        response.write("<td class='td-l fonttiny'>"& rsUpdateEntry4("Pedidos") &"</td>")
        response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("subTotal") &"</td>")
        rsUpdateEntry4.close
        RowOut
        rsUpdateEntry3.MoveNext
      wend
      rsUpdateEntry3.close
      closeTable()
      response.write("</td><td width='50%' class='td-c' style='vertical-align:top'>")
      CreateTable(590)
      RowIn
      response.write("<td colspan=2 class='td-c strong subtitulo'>XXXX</td></tr>")
      RowOut
      closeTable()
      response.write("</td></tr>")
      closeTable()

  end if

end sub


sub PotterRoemer 'contenido 126'

     strSQL="exec [dbo].[SP_StockConsignado]"
     rsUpdateEntry4.Open strSQL, adoCon4

     response.redirect("/home.asp?msg=se ha enviado el informe stock Potter Roemer")


end sub


sub ControlPagos 'contenido 128'
     response.write("<P>&nbsp;</P>")
     doAlert
     titulo="CONTROL DE PAGOS A FABRICANTE"
     DoTitulo(titulo)


      strSQL="select ID,.dbo.fn_GetMonthNameUS(DocDate,'english') as Fecha,RS,Tipo,NombreCorto,'$ ' + CONVERT(VARCHAR,CONVERT(MONEY,Monto),1) as Monto  ,ID_USUARIO,CantFacturas,Facturas,nombre_archivo,year(DocDate) as ANIO,month(docdate) as MES from PAGOS order by ID desc "

      'response.write strSQL
      Const adCmdText = &H0001
      Const adOpenStatic = 3

      rsUpdateEntry5.Open strSQL, adoCon
      rsUpdateEntry6.Open strSQL, adoCon, adOpenStatic, adCmdText  

      rsUpdateEntry6.PageSize = 10 
      nPageCount = rsUpdateEntry6.PageCount


      Dim Campos(20)
      CreateTable(900)

      if not rsUpdateEntry6.EOF then
           if request("vPage")<>"" then
                nPage = int(trim(request("vPage")))
           else
                nPage=1
           end if
           rsUpdateEntry6.AbsolutePage=npage

           response.write("<tr><td colspan=3 class='td-r'><B><U>PAGINAS:</U></B></td><td colspan=5 class='td-l'>")
           for i=1 to nPageCount step 1
                if i<>npage then  
                    response.write("<a href='/modules/ShowContent.asp?contenido=104&action=1&vPage="& i &"'>") 
                    response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
                end if
              if i=50 or i=100 or i=150 or i=200 then             
                   response.write("<br>")
              end if
          next
          response.write("</td></tr>")
      end if


         i=1
         For Each rsUpdateEntry5 in rsUpdateEntry5.Fields                        
            campos(i)=rsUpdateEntry5.Name
            i=i+1           
         Next     
         
         RowIn
         for i=1 to 8           
            Response.Write "<th class='td-c'>" & campos(i) & "</th>"               
         next      
         Response.Write "<th class='td-c' colspan=2>Controles</th>"  
         'Response.Write "<th class='td-c'>&nbsp;</th>"         
         RowOut

         vMes=""
         while not rsUpdateEntry6.EOF   
             rowin           
             for i=1 to 8
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry6(campos(i)) &"</td>")
             next    
             if trim(rsUpdateEntry6("Facturas"))<>"" then
                  Response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=129&Mode=1&ID="&rsUpdateEntry6("ID")&"'><img src='/img/check-email.png' border=0 alt='recibir email y revisar' title='recibir email y revisar'></a></td>") 
                  Response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=129&Mode=2&ID="&rsUpdateEntry6("ID")&"'><img src='/img/EmailOut2.png' border=0 alt='enviar email a Fabricante' title='enviar email a Fabricante'></a></td>") 
             else
                  Response.write("<td class='td-r fonttiny'>N/A</td>") 
                  Response.write("<td class='td-r fonttiny'>N/A</td>") 
             end if 
            
             RowOut
             RowIn
             response.write("<td class='td-l fonttiny' colspan=5><B>[FACTURAS:] </B>"& rsUpdateEntry6(campos(9)) &"</td>")

             if int(trim(rsUpdateEntry6("MES")))<10 then 
                 vMes="0"&trim(rsUpdateEntry6("MES")) 
             else vMes=trim(rsUpdateEntry6("MES")) end if  

             response.write("<td class='td-l fonttiny' colspan=5><B>[ARCHIVO:] </B><a href='http://192.168.0.206/repositorio/PAYMENTS/"&rsUpdateEntry6("ANIO")&"/"&vMes&"/' target='pagos'>"& rsUpdateEntry6(campos(10)) &"</a></td>")   
             RowOut
             separador
             rsUpdateEntry6.MoveNext                         
             
         wend
         rsUpdateEntry6.close
         closeTable


end sub


sub SendEmailVendorDraft   'contenido 129'
     if request("Mode")="1" then
          titulo="enviar Email comprobante de pago a fabricante (para revision)"
     end if
      if request("Mode")="2" then
          titulo="enviar Email comprobante de pago a fabricante"
     end if
     doTitulo(titulo)

       %>

     <form id="fabricante" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="130">
     <input type="hidden" name="mode" value="<%=request("mode")%>">
     <input type="hidden" name="ID" value="<%=request("ID")%>">
      

     Destinadarios:  <br>   <%

     if request("Mode")="1" then
         response.write("<input style='width: 900px' type='text' name='femail' value='ajauregui@deltaflow.com.mx;'><br><br>")
         response.write("<input type='submit' value='enviar email para revision'>")
     end if
     if request("Mode")="2" then
         strSQL="select * from PAGOS where ID="&request("ID")
         rsUpdateEntry.Open strSQL, adoCon

         strSQL="select * from OCRD where cardcode='"& rsUpdateEntry("cardCode") &"' "
         select case rsUpdateEntry("RS")
            case "DMX" rsUpdateEntry2.Open strSQL, adoCon2
            case "DFW" rsUpdateEntry2.Open strSQL, adoCon3
            case "MME" rsUpdateEntry2.Open strSQL, adoCon5
          end select

         response.write("<input style='width: 900px' type='text' name='femail' value='"& rsUpdateEntry2("U_CFDi_Email") &"'><br><br>")
         rsUpdateEntry2.close
         response.write("<input type='submit' value='enviar email para revision'>")
     end if
     response.write("</form>")

end sub




sub SendEmailVendor  'contenido 130'

    if request("Mode")="" then response.redirect("ShowContent.asp?contenido=128&msg=ocurrio un error interno, vuelva a intentarlo")

    strSQL="EXEC  [dbo].[SP_SendEmailVendor]  @ID = "& request("ID") &", @Mode = "& request("mode") &", @Destinatarios = N'"& request("femail") &"'"    
    response.write strSQL
    rsUpdateEntry5.Open strSQL, adoCon

    response.redirect("ShowContent.asp?contenido=128&msg=un comprobante de pago a fabricante ha sido enviado por email")

end sub





sub   SolicitarPagar  'contenido 131'  
      'limpia los registros no confirmados'
      strSQL="delete facturas where xml is null and DATEDIFF(mi,logdate,getdate())>5"
      rsUpdateEntry.Open strSQL, adoCon4

      doAlert
      titulo="Solicitar pagar una facura de servicios/articulos en PPD"
      response.write("<P class='fontmed'>&nbsp;</P>")
      doTitulo(titulo)

      if request("action")="" then

          response.write("<P class='fontmed strong'>REQUISITOS ESTABLECIDOS POR CONTABILIDAD:<br>")
          CreateTable(600)
          rowin
          response.write("<td>* Que el socio de negocio ya se encuentre registrado en SAP, de no ser as&iacute; tramitar su alta") 
          RowOut
          rowin
          response.write("<td>* Contar con el archivo xml del gasto ejercido") 
          RowOut
          rowin
          response.write("<td>* Contar con el archivo pdf del gasto ejercido")  
          RowOut     
          rowin
          response.write("<td>* Una factura a cr&eacute;dito se programar&aacute; a pago el  d&iacute;a h&aacute;bil anterior a su vencimiento. ")  
            'Una factura al contado se programar&aacute; a pago, al pr&oacute;ximo jueves h&aacute;bil o el &uacute;ltimo dia h&aacute;bil del mes lo que ocurra primero.'

          RowOut   
          closeTable()
         
          response.write("<form id='Pagos' method='POST' action='ShowContent.asp'>")
          response.write("<input type='hidden' name='contenido' value=131>")
          response.write("<input type='hidden' name='action' value=1>")
          response.write("<div id=DivRS>")
            
            response.write("<P class='fontmed'>&nbsp;</P><P class='fontmed strong'>REQUISITOS PARA DAR DE ALTA UN SOCIO DE NEGOCIO:<br>")
            CreateTable(600)
            RowIn
            response.write("<td>* Constancia de Situacion Fiscal (CSF) no mayor de 30 d&iacute;as") 
            RowOut
            rowin
            response.write("<td>* Opini&oacute;n de cumplimiento en -positivo-") 
            RowOut
            RowIn
            response.write("<td>* Car&aacute;tula de banco, cuentas de recepci&oacute;n de pago")   
            RowOut
            rowin
            response.write("<td>* Contacto y cuenta de email para comprobantes de pago o gesti&oacute;n de complementos")     
            RowOut
            closeTable()
            
            response.write("<P class='td-c'>Raz&oacute;n Social:")
            %> 

            <select name='RS' ID='RS' onchange="PassValueChangeDiv('RS','DivRS','PagoFacturaForm.asp')">
                 <option value=0>seleccione</option>
                 <option value='DMX'>DMX</option>
                 <option value='DFW'>DFW</option>
                 <option value='MME'>MME</option></select>
            <%
          response.write("</div>")
          response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

     end if

      
     if request("action")="1" and request("RS")<>"" and request("SN")<>"" then   'antes del uploading'
          response.write(request("RS")&"<BR>")
          response.write(request("SN")&"<BR>")

          if request("fmasivo")="on" OR request("fmasivo")="true"  then response.redirect("ShowContent.asp?contenido=131&action=4&RS="&request("RS")&"&SN="&request("SN") )

          strSQL="select * from OCRD where cardcode='"& request("SN") &"' "
          select case request("RS")
            case "DMX" rsUpdateEntry.Open strSQL, adoCon2
            case "DFW" rsUpdateEntry.Open strSQL, adoCon3
            case "MME" rsUpdateEntry.Open strSQL, adoCon5
          end select
          vRFC=trim(rsUpdateEntry("LicTradNum"))
          rsUpdateEntry.close

          strSQL="insert into Facturas (ID,RS,CardCode,id_usuario,LogDate,LicTradNum) values (isnull((select max(ID) from Facturas),0)+1,'"&request("RS")&"','"&request("SN")&"','"&session("session_id")&"',getdate(),'"& vRFC &"')"
          rsUpdateEntry5.Open strSQL, adoCon4

          strSQL="select max(ID) as ID from Facturas where id_usuario='"&session("session_id")&"' and cast(LogDate as Date)=cast(GETDATE() as date)"
          rsUpdateEntry.Open strSQL, adoCon4

          response.redirect("http://192.168.0.206/repositorio/ADMON/upload/uploadform-Payments.asp?Mode=1&ID="&rsUpdateEntry("ID")&"&RS="&request("RS")&"&cardCode="&request("SN")  )

     end if


     if request("action")="2" then  'termino el uploading se preguntara ASAP y notas'

               'notas after uploading'
               strSQL="select * from FACTURAS where ID="&request("ID")
               rsUpdateEntry5.Open strSQL, adoCon4

               response.write("<P class='td-c'><H3>Editando solicitud no. "&request("ID") &"</H3></P>")

               response.write("<form id='Pagos' method='POST' action='ShowContent.asp'>")
               response.write("<input type='hidden' name='contenido' value=131>")
              
               response.write("<input type='hidden' name='action' value='3'>")
               response.write("<input type='hidden' name='ID' value='"&request("ID")&"'>")
               response.write("<P>&nbsp;</P>")
               'response.write("<P><input type='checkbox' name='fASAP'> Pagar lo mas inmediato posible?<BR>") 
               response.write("Agregar nota: <input type='text' class='anchox4' name='fNotas' value='"&rsUpdateEntry5("notas") &"'></P>")
               rsUpdateEntry5.close
               response.write("<P>&nbsp;</P>")
               response.write("<P><input type='submit' value='finalizar'></form></P>")


     end if

     
     if request("action")="3" then  'termino el uploading se preguntara ASAP y notas'

                    'if request("fASAP")="true" or request("fASAP")="on" then
                        'strSQL="UPDATE FACTURAS set notas='"&request("fNotas")&"',ASAP=1,DocDate=getdate() where ID="&request("ID")
                    'else
                        'strSQL="UPDATE FACTURAS set notas='"&request("fNotas")&"' where ID="&request("ID")
                    'end if

                    strSQL="UPDATE FACTURAS set notas='"&request("fNotas")&"' where ID="&request("ID")
                    response.write strSQL
                    rsUpdateEntry7.Open strSQL, adoCon4

                    'LO DISEÑE PARA PUE -QUEDARA PENDIENTE'
                    'strSQL="exec [dbo].[SP_ProgramaPagoFact] "&request("ID")
                    'rsUpdateEntry5.Open strSQL, adoCon4             
 
                    response.redirect("ShowContent.asp?contenido=132&Mode=1&ID="&request("ID")&"&msg=se actualizaron las notas de la solicitud de pago no."&request("ID")  )

     end if

     if request("action")="4" then  'form carga masiva'
            if session("session_id")="" then response.redirect("ShowContent.asp?contenido=131&msg=la sesion caduco")

            vID=0

            if month(date())<10 then 
               vRepositorio="0"&trim(month(date()))
            else
                vRepositorio=trim(month(date()))
            end if
            vRepositorio=trim(year(date()))&"/"&vRepositorio

            strSQL="insert into CARGASFACTURAS (ID,RS,CardCode,LogDate,id_usuario) values ((select isnull(max(ID),0)+1 from CARGASFACTURAS),'"&request("RS")&"','"&request("SN")&"',getdate(),'"& session("session_id") &"') "
            rsUpdateEntry.Open strSQL, adoCon4

            strSQL="select max(ID) as calculo from CARGASFACTURAS where id_usuario='"& session("session_id") &"' "
            rsUpdateEntry.Open strSQL, adoCon4
            vID=rsUpdateEntry("calculo")
            rsUpdateEntry.close

            strSQL="UPDATE CARGASFACTURAS set Repositorio='"& vRepositorio&"/"&trim(vID) &"' where ID="&vID
            rsUpdateEntry.Open strSQL, adoCon4

            strSQL="EXEC xp_cmdshell  'MD d:\fileserver\wwwroot\repositorio\ADMON\CargasMasivas\" & replace(vRepositorio,"/","\") &"\"& trim(vID) &"'"
            rsUpdateEntry2.Open strSQL, adoCon4
            'response.write strSQL

            CreateTable(350)
            RowIn
            response.write("<td class='td-c strong FontMed'>Coloque todas las facturas a programar en la siguiente ubicaci&oacute;n:  <button><a target='_self' href ='localexplorer:R:\ADMON\CargasMasivas\" & replace(vRepositorio,"/","\") &"\"& trim(vID) &"' style='text-decoration: none;'>Abrir ubicaci&oacute;n</a></button><br><font color=red>R:\ADMON\CargasMasivas\" & replace(vRepositorio,"/","\") &"\"& trim(vID) &"</font><br><br>cuando haya terminado de cargar todas las facturas de click en el siguiente bot&oacute;n: <br><a href='http://192.168.0.206/repositorio/ADMON/xmls_cargas.asp?ID="& vID &"&usuario="& session("session_id") &"'><img src='/img/continuar.png' border=0></a></td>"  )
            RowOut
            closeTable()
     end if

     if request("action")="5" then  'form carga masiva'
             if session("session_id")="" then response.redirect("ShowContent.asp?contenido=131&msg=la sesion caduco")


            response.write("<P class='td-c'><H3>Los siguientes XML califican para programarse a pago: </H3></P>")

            strSQL="select TipoComprobante as Tipo,MetodoPago as Metodo,FormaPago as Forma,Moneda,UsoCFDi as USO,Folio,Emisor,RegimenFiscal as Regimen,version,UUID,Total,Receptor from CARGASFACTURAS  where ID=0 and id_usuario='"& session("session_id") &"'"
            'response.write strSQL
            rsUpdateEntry.Open strSQL, adoCon4
            rsUpdateEntry2.Open strSQL, adoCon4

            Dim Campos(12) 
            i=1
         
             For Each rsUpdateEntry in rsUpdateEntry.Fields            
                'Response.Write("<th>" & rsUpdateEntry.Name & "</th>")
                campos(i)=rsUpdateEntry.Name
                i=i+1           
             Next   


            CreateTable(900)
             RowIn
               For i=1 to 12
                  response.Write("<th class='td-c fonttiny'>" & campos(i) & "</th>")
               next  
             RowOut 

             while not rsUpdateEntry2.EOF
                RowIn
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Tipo") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Metodo") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Forma") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Moneda") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("USO") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Folio") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Emisor") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Regimen") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("version") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("UUID") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("total") &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("Receptor") &"</td>")
                RowOut
                rsUpdateEntry2.MoveNext
             wend
             rsUpdateEntry2.close

            closeTable()
            response.write("<P class='td-c'>Si est&aacute; seguro(a) desea programar estas facturas para pago, d&eacute; click en bot&oacute;n: <br><a href='showContent.asp?contenido=131&action=6&ID="& request("ID") &"'><img src='/img/continuar.png' border=0></a></P>")

  
     end if


      if request("action")="6" then  'form carga masiva'
             if session("session_id")="" then response.redirect("ShowContent.asp?contenido=131&msg=la sesion caduco")

             strSQL="exec SP_CopiaArchivosFacturasMasivo "& request("ID") &",'"& session("session_id") &"'"
             rsUpdateEntry.Open strSQL, adoCon4

             strSQL="SELECT * from CARGASFACTURAS where ID="& request("ID")
             rsUpdateEntry.Open strSQL, adoCon4
 
             response.redirect("ShowContent.asp?contenido=132&mode=1&RS="& rsUpdateEntry("RS") &"&cardcode="& rsUpdateEntry("cardCode") &"&msg=todas las facturas seleccionadas fueron programadas para pago")

      end if


end sub




sub   ControlPagosFacturas  'contenido 132' 
        busquedas=0  
        if request("femisor")<>"" or  request("fUUID")<>"" or request("ffolio")<>""  or request("fusuario")<>"" and request("vRFC")<>"" OR request("RFC")<>"" then
              busquedas=1
        end if
        
        SaltoLinea

        'limpia la programacion de pago facturas no confirmados'
        strSQL="delete facturas where (xml is null and DATEDIFF(mi,logdate,getdate())>5 ) OR MetodoPago='PUE' "
        rsUpdateEntry.Open strSQL, adoCon4
 
        'limpia la aplicacion de pagos no detallada en montos'
        strSQL="delete PagosFacturas where DATEDIFF(mi,logdate,getdate())>5 and (Comprobante is null or ImpPago is null or ImpPago=0) "          
        rsUpdateEntry.Open strSQL, adoCon4
        
        doAlert
        response.write("<P class='td-c'>")               
        if request("mode")="" then response.redirect("ShowContent.asp?contenido=132&mode=1")   

        if request("mode")="1" then
             response.write("<P class='td-c'>")
             response.write("<a href='ShowContent.asp?contenido=131'>Solicitar pagar</a> | ")
             response.write("<a href='ShowContent.asp?contenido=134'>Pagos Realizados</a> | ")
             response.write("<a href='ShowContent.asp?contenido=132&mode=2'>Facturas Pagadas</a> | ") 
             response.write("<a href='ShowContent.asp?contenido=133'>Complementos recibidos</a>")  
             response.write("</P>")   
             titulo="<a href='ShowContent.asp?contenido=132&mode=1'>Programaci&oacute;n de pago  facturas PPD</a>"
             doTitulo(titulo)
                    
        else          
             response.write("<P class='td-c'>")
             response.write("<a href='ShowContent.asp?contenido=131'>Solicitar pagar</a> | ")
             response.write("<a href='ShowContent.asp?contenido=134'>Pagos Realizados</a> | ")
             response.write("<a href='ShowContent.asp?contenido=132&mode=1'>Facts sin pago | </a>")
             response.write("<a href='ShowContent.asp?contenido=133'>Complementos recibidos</a>") 
             response.write("</P>")      
             titulo="<a href='ShowContent.asp?contenido=132&mode=2'>Facturas Pagadas</a>"
             doTitulo(titulo)
              
        end if

        
        '=================================================================================================================='
        response.write("<form id='facturas' method='POST' action='ShowContent.asp'>")
        response.write("<input type='hidden' name='contenido' value='132'>")
        response.write("<input type='hidden' name='Mode' value='"&request("Mode")&"'>")
        '=================================================================================================================='


        if request("RFC")="" then
        response.write("<P class='fontmed'><font class='subtitulo'>BUSQUEDA DE FACTURAS</font><br><br><font class='fonttiny'>")

        response.write("[emisor:] <input type='text' name='femisor' value='"&request("femisor")&"' style='width:90px'>")
        response.write("[UUID:] <input type='text' name='fUUID' value='"&request("fUUID")&"' style='width:320px' maxlength='36'> ")
        response.write("[folio:] <input type='text' name='fFolio' value='"&request("fFolio")&"' style='width:90px'> ")
        

          strSQL="select distinct(id_usuario) as id_usuario from FACTURAS where FechaPagada is null"
          rsUpdateEntry5.Open strSQL, adoCon4

          response.write("[usuario:] <select name='fusuario'>")
          response.write("<option value=''>&nbsp;</option>")
          while not rsUpdateEntry5.EOF
              response.write("<option value='"&rsUpdateEntry5("id_usuario")&"'> "&rsUpdateEntry5("id_usuario")&"</option>")
              rsUpdateEntry5.MoveNext
          wend
          rsUpdateEntry5.close
          response.write("</select>")
          if request("mode")="2" then   response.write("[filtro:] <select name='ffiltro'><option value=''>todas</option><option value='1'>sin complemento</option></select>")
          response.write("</font><input type='submit'  value='buscar'></P><br>")
        end if

        if request("RFC")<>"" then      'cuando es RFC apenas va a consolidar'         
             response.write("<input type='hidden' name='vRFC' value='"&request("RFC")&"'>")  
             response.write("<P class='td-c alert'>Se muestran en pantalla facturas con el mismo RFC para que seleccione una o varias que desea pagar</P>")            
        end if


        if request("vRFC")<>"" then  

             strSQL="select * From Facturas where RFCEmisor='"&request("vRFC")&"' and FechaPagada is null "
             rsUpdateEntry5.Open strSQL, adoCon4
             facturas=""
            
             vFlag=0
             vfecha=""
             while not rsUpdateEntry5.EOF
                     if request("ffecha")="" then  
                         vfecha=trim(year(date()))&"-"&trim(month(date()))&"-"&trim(day(date()))
                     else
                         vfecha=request("ffecha")
                     end if

                     vString="Monto" & rsUpdateEntry5("ID")                     
                     if request(vString)<>""  then
                          Monto=replace(replace(trim(request(vString)),",",""),"$","")
                          facturas=facturas &rsUpdateEntry5("ID") & "*"
                          vFlag=1
                          strSQL="INSERT INTO [dbo].[PAGOSFACTURAS] (ImpPago,SaldoInsoluto,UUID_Factura,LogDate,DocDate) "
                          strSQL=strSQL &"values ("&Monto&","&rsUpdateEntry5("Total")&",'"&rsUpdateEntry5("UUID")&"',getdate(),'"&vfecha&"')"
                          response.write(strSQL&"<br>")
                          rsUpdateEntry6.Open strSQL, adoCon4
                     end if
                     rsUpdateEntry5.MoveNext
              wend
              rsUpdateEntry5.close

              if vFlag=0 then
                       response.redirect("ShowContent.asp?contenido=132&Mode=1&RFC="&request("vRFC")&"&msg=no se indico por lo menos un monto para aplicar")
              else
                       facturas=mid(facturas,1,len(facturas)-1)
                       response.write(facturas&"<br>")

                       if request("fpagado")="true" or request("fpagado")="on" then

                            comprobante="No-disponible-"&facturas&".png"

                            strSQL="UPDATE a set a.Comprobante='"&comprobante&"',a.CopiarCombrobante=1 FROM PagosFacturas a inner join Facturas b on a.UUID_Factura=b.UUID WHERE b.ID in (" & facturas &")"      
                            Response.write(strSQL&"<br>")
                            rsUpdateEntry6.Open strSQL, adoCon4

                            Response.redirect("http://192.168.0.204/modules/ShowContent.asp?contenido=134&comprobante="&comprobante&"&msg=su pago ha sido aplicado")

                       else
                           response.redirect("http://192.168.0.206/Repositorio/ADMON/upload/uploadform-Comprobantes.asp?ID="&facturas)
                       end if
              end if
              
            

        end if



        'QUERY GLOBAL DE FACTURAS'
        strSQL="select ID,RS,CardCode,LicTradNum as RFC_SAP,TipoComprobante as Tipo,MetodoPago as Metodo,FormaPago as Forma,Moneda,UsoCFDi as Uso,Folio,RegimenFiscal as Regimen,version,.dbo.fn_GetMonthName(DocDate,'spanish') as Fecha,case when DocDueDate is null then '' else .dbo.fn_GetMonthName(DocDueDate,'spanish')  end as Vencimiento,Total,Emisor,UUID,Notas, case when FechaPagada is null then 0 else 1 end as Flag,pdf,xml,case when DocDueDate is not null then case when DocDueDate<=cast(getdate() as date) then 1 else 0 end else 0 end as vencida,Repositorio,RFCEmisor,id_usuario,.dbo.fn_GetMonthName(LogDate,'spanish') as FechaSolicitud,ArchivoSoporte   from Facturas WHERE MetodoPago='PPD' "

          FlagMontos=0
        
         if request("ID")<>"" then   strSQL=strSQL&" and ID="&request("ID")
         if request("RFC")<>"" then  strSQL=strSQL&" and RFCEmisor='"&request("RFC")&"' "        
         if request("UUID")<>"" then  strSQL=strSQL&" and UUID='"&request("UUID")&"' "
         if request("comprobante")<>"" then  strSQL=strSQL&" and Comprobante='"&request("comprobante")&"'"
         if request("fusuario")<>"" then     strSQL=strSQL&" and id_usuario='"&request("fusuario")&"'"

        if session("session_id")="AKGONZALEZ" then

            if session("session_id")="AKGONZALEZ" then  strSQL=strSQL&" and id_usuario in ('AKGONZALEZ','KCARLOS','MRANGEL') "

        else
             if session("session_id")<>"ACRUZ" and session("session_id")<>"MCABANILLAS" and session("session_id")<>"YOROZCO" and session("session_id")<>"" then  strSQL=strSQL&" and id_usuario='"&session("session_id")&"'"
        
        end if

        

        if request("RS")<>"" then        strSQL=strSQL&" and RS='"&request("RS") &"' "
        if request("Cardcode")<>"" then  strSQL=strSQL&" and cardCode='"&request("Cardcode") &"' "

        'busquedas'
         if len(trim(request("femisor")))>=3 then  
                strSQL=strSQL&" and ( RFCEmisor like '%"&request("femisor")&"%' OR emisor like '%"&request("femisor")&"%') "
                busquedas=1
         end if
         if len(trim(request("fUUID")))>=3 then  
                strSQL=strSQL&" and UUID like '%"&request("fUUID")&"%' "
                busquedas=1
         end if
         
         if len(trim(request("fFolio")))>=3 then  
                strSQL=strSQL&" and folio like '%"&request("fFolio")&"%' "
                busquedas=1
         end if
         
         if request("MODE")="1" then strSQL=strSQL&" and FechaPagada is null "
         if request("MODE")="2" then strSQL=strSQL&" and FechaPagada is not null "
         
         if request("ffiltro")="1" then 
              strSQL2="select stuff( (select ','+convert(varchar,b.ID) from PAGOSFACTURAS a inner join FACTURAS b on a.UUID_Factura=b.UUID where a.IdDocumento is null for XML PATH ('') ),1,1,'') as calculo"
              rsUpdateEntry5.Open strSQL2, adoCon4

              strSQL=strSQL&" and ID in ("&rsUpdateEntry5("calculo")&") "
              rsUpdateEntry5.close
          end if

         strSQL=strSQL&" order by ID desc" 

          
        Const adCmdText = &H0001
        Const adOpenStatic = 3
        nPage=0
        registros=0
      
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon4
        rsUpdateEntry2.Open strSQL, adoCon4, adOpenStatic, adCmdText
             
        MultiplesPagos=0
        if len(trim(request("fUUID")))>=36 and not rsUpdateEntry2.EOF then  
                 MultiplesPagos=1
                strSQL="select * from PAGOSFACTURAS where Comprobante=(select Comprobante from PAGOSFACTURAS WHERE UUID_Factura='"& request("fUUID") &"') and UUID_Factura<>'"& request("fUUID") &"' "
                'response.write strSQL 
                rsUpdateEntry5.Open strSQL, adoCon4               
        end if

        if busquedas=1 then
           rsUpdateEntry2.PageSize = 10000 
           vPageSize=10000
        else
           rsUpdateEntry2.PageSize = 10
           vPageSize=10
        end if

        nPageCount = rsUpdateEntry2.PageCount

        CreateTable(1160)

   
    if busquedas=0 then

        if not rsUpdateEntry.EOF then
            
               if request("vPage")<>"" then
                    nPage = int(trim(request("vPage")))
                else
                    nPage=1
                end if
                rsUpdateEntry2.AbsolutePage=npage

                response.write("<tr><td colspan=3 class='td-r'><B><U>PAGINAS: </U></B></td><td colspan=8 class='td-l'>")
                for i=1 to nPageCount step 1
                      if i<>npage then  
                          response.write("<a href='/modules/ShowContent.asp?contenido="& request("contenido") &"&Mode="&Request("mode")&"&vPage="&i &"'>") 
                          response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
                      end if
                    if i=50 or i=100 or i=150 or i=200 then             
                         response.write("<br>")
                    end if
                next
                response.write("</td></tr>")

        else

            NoRegistros

        end if

    end if
   



         veces=0
         Dim Campos(30) 
         i=1
         
         For Each rsUpdateEntry in rsUpdateEntry.Fields            
            'Response.Write("<th>" & rsUpdateEntry.Name & "</th>")
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next   

         RowIn
         For i=1 to 15
            response.Write("<th class='td-c fonttiny'>" & campos(i) & "</th>")
         next     
         if FlagMontos=1 then       
             response.Write("<th colspan=2 class='td-c' class='fonttiny'>Monto aplicar en factrura</th>")
         else
             response.Write("<th colspan=2 class='td-c' class='fonttiny'>Controles</th>")
         end if
         RowOut 

         '===========query pagos
         strSQL="select TOP(1) TipoComprobante as Tipo, Folio, UUID, xml as PDF, version, RFCEmisor, Emisor,  RegimenFiscal as Regimen, RFCRecptor,  Receptor,IdDocumento,  TipoCambioP as TC, MonedaP, FormaDePagoP as Forma,  ImpPagado, ImpSaldoAnt, ImpSaldoInsoluto,  NumParcialidad as Parcialidad, Comprobante, UUID_Factura,ImpPago,SaldoInsoluto from PagosFacturas"
         rsUpdateEntry7.Open strSQL, adoCon4
         Dim Campos2(30) 
         i=1
         For Each rsUpdateEntry7 in rsUpdateEntry7.Fields                        
            campos2(i)=rsUpdateEntry7.Name
            i=i+1           
         Next    

        
         while not rsUpdateEntry2.EOF AND registros<vPageSize
               rowin
               if request("mode")=1 then
                   response.write("<td class='fonttiny td-r'>"& rsUpdateEntry2("ID") &" <a href='ShowContent.asp?contenido=132&mode=1&action=6&ID="&rsUpdateEntry2(campos(1))&"'><img src='/img/b_drop.png' alt='eliminar programacion de pago' title='eliminar programacion de pago'></a></td>")
                else
                  response.write("<td class='fonttiny td-r'>"& rsUpdateEntry2("ID") &" </td>")

                end if

               for i=2 to 13
                  response.write("<td class='fonttiny td-c'>"& rsUpdateEntry2(campos(i)) &"</td>")
               next
               if rsUpdateEntry2("vencida")="1" then
                   response.write("<td class='fonttiny td-r' width='90px' style='color:white; background-color: red'>"& rsUpdateEntry2(campos(14)) &"</td>")
               else
                   response.write("<td class='fonttiny td-r' width='90px' style='background-color: #E4FEE4'>"& rsUpdateEntry2(campos(14)) &"</td>")
               end if

              

                response.write("<td class='fonttiny td-r'><B>$ "& rsUpdateEntry2("total") &"</B></td>")


               if int(trim(rsUpdateEntry2("flag")))=0 then
                  response.write("<td class='fonttiny td-c'>")

                  if request("RFC")<>"" then
                      vString="consolida"&rsUpdateEntry2("ID") 
                           vString="div"&rsUpdateEntry2("ID")
                           vString2="consolida"&rsUpdateEntry2("ID")
                           %>
                           <td class="fonttiny td-c" colspan=2>
                               <div id="<%=vString%>">
                                    check para aplicarle pago
                                    <input type="checkbox" name="<%=vString2%>" onclick="Javascript:sendReq('/modules/PPD_Montos.asp','ID,Monto','<%=rsUpdateEntry2("ID")%>,<%=rsUpdateEntry2("Total")%>','div<%=rsUpdateEntry2("ID")%>')">
                               </div>
                             </td>                          
                           </tr>   <%                   
                  else
                     if FlagMontos=0 then
                        response.write("<td class='fonttiny td-c'><a href='ShowContent.asp?contenido=132&Mode=1&RFC="&rsUpdateEntry2("RFCEmisor")&"'><img src='/img/consolidar.png' width='60px' alt='aplicar un pago' title='aplicar un pago'></a></td>")                     
                     end if

                  end if
                                                                
  
               else
                   response.write("<td class='fonttiny td-c' colspan=2><img src='/images/paid.png'></td>")
               end if
               RowOut

                

               RowIn
                  response.write("<td class='fonttiny td-l' colspan=3><B>[EMISOR:] </B>"& mid(rsUpdateEntry2("Emisor"),1,18) &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry2("RFCEmisor") &"</td>")

                   if trim(rsUpdateEntry2("RFC_SAP"))=trim(rsUpdateEntry2("RFCEmisor")) then
                          response.write("<td class='fonttiny td-c'><font style='background-color: #FFFFC0'>RFC OK</td>")
                   else
                          response.write("<td class='fonttiny td-c'>RFC ERROR!</td>")
                   end if
                   response.write("<td class='fonttiny td-c'>&nbsp;</td>")

                  response.write("<td class='fonttiny td-l' colspan=7><B>[UUID:] </B><a href='http://192.168.0.206/repositorio/ADMON/FacturasProvisionadas/"& rsUpdateEntry2("repositorio") &"/" &rsUpdateEntry2("xml")&"' target='xml'>"& rsUpdateEntry2("UUID") &" <img src='/img/descargue.png' border=0></a></td>")
                  response.write("<td class='fonttiny td-l' colspan=2><B>[PDF:] </B><a href='http://192.168.0.206/repositorio/ADMON/FacturasProvisionadas/"&rsUpdateEntry2("repositorio")& "/" &rsUpdateEntry2("PDF")&"' target='pdf'><img src='/images/FilePDF2.png' border=0></a></td>")

                  if request("mode")=1 then
                     if len(rsUpdateEntry2("ArchivoSoporte"))>3 then
                        response.write("<td class='fonttiny td-c' colspan=2 rowspan=2><a href='http://192.168.0.206/repositorio/ADMON/DocumentoSoporte/"&rsUpdateEntry2("ArchivoSoporte")&"' target='soporte'><img src='/images/DocSoporte2.png' border=0 height='36px' alt='descargue papel de trabajo' title='descargue papel de trabajo'></td>")

                        if request("action")<>"2" then
                          response.write("<td class='fonttiny td-l' rowspan=2><a href='ShowContent.asp?contenido=132&mode=1&action=2&ID="&rsUpdateEntry2("ID")&"'><img src='/images/revision.png' border=0 height='36px' alt='solicitar revision' title='solicitar revision'></td>")
                          end if

                     else
                          response.write("<td class='fonttiny td-c' colspan=2 rowspan=2><a href='http://192.168.0.206/repositorio/ADMON/upload/uploadform-Soporte.asp?ID="&rsUpdateEntry2("ID")&"'><img src='/images/DocSoporte1.png' border=0 height='36px' alt='cargue documento soporte' title='cargue documento soporte'></td>")
                     end if
                  else
                     response.write("<td class='fonttiny td-c' colspan=2 rowspan=2>&nbsp;</td>")
                  end if
                RowOut


               Rowin
                      response.write("<td class='fonttiny td-l' colspan=6><B>[Notas:] </B>")

                      if trim(rsUpdateEntry2("Notas"))<>"" then
                           response.write("<font style='background-color: #FFFFC0'><a href='ShowContent.asp?contenido=131&action=2&ID="&rsUpdateEntry2("ID")&"'>" &rsUpdateEntry2("Notas") &"</a></font></td>")
                      else
                           response.write("<a href='ShowContent.asp?contenido=131&action=2&ID="&rsUpdateEntry2("ID")&"'><U>agregue notas</U></A></font></td>")
                      end if

                      response.write("<td class='fonttiny td-l' colspan=6><B>[Solicitante:] </B>"& rsUpdateEntry2("id_usuario") &"</font></td>")
                      response.write("<td class='fonttiny td-l' colspan=3><B>[Solicitado en:] </B>"& rsUpdateEntry2("FechaSolicitud") &"</font></td>")

                      'response.write("<td class='fonttiny td-c' colspan=2>XXXX</td>")

               RowOut
               '=============================================================================================
                strSQL="select TipoComprobante as Tipo, replace(Folio,',','<BR>') Folio, UUID, xml as PDF, version, RFCEmisor, Emisor,  RegimenFiscal as Regimen, RFCRecptor,  Receptor,IdDocumento,  TipoCambioP as TC, MonedaP, FormaDePagoP as Forma,  ImpPagado, ImpSaldoAnt, ImpSaldoInsoluto,  NumParcialidad as Parcialidad, Comprobante, UUID_Factura,ImpPago,SaldoInsoluto,Repositorio,case when CHARINDEX(',',UUID)>1 then 1 else 0 end as VariosComplementos from PagosFacturas where UUID_Factura='"&rsUpdateEntry2("UUID")&"' "
                'response.write("<tr><td colspan=19>"&strSQL&"</td></tr>")
                rsUpdateEntry4.Open strSQL, adoCon4

                response.write("<tr><td colspan=19 class='td-c strong'>  - - - - - - - - - - - - - - - - - - - P A G O S - - - - - - - - - - - - - - - - - - - - - R E A L I Z A D O S - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - </td></tr>")
               

               if not rsUpdateEntry4.EOF then
                   RowIn
                   response.write("<td class='td-c' colspan=19>")
                   response.write("<table width=1160 borde=1 cellspacing=1 cellpadding=2 align=center>")
                   
                   RowIn
                     Response.Write("<td class='td-c'>&nbsp;</td>")
                     for n=1 to 18
                          Response.Write("<td class='subtitulo td-c'>" & campos2(n) & "</td>")  
                     next    
                   RowOut

                   while not rsUpdateEntry4.EOF
                        Rowin
                      response.write("<td class='subtitulo td-l fonttiny'>tesoreria:</td>")
                      response.write("<td colspan=6 class='td-l fonttiny'><B>[Comprobante:]</B><a href='http://192.168.0.206/repositorio/ADMON/PagoFacturas/"& rsUpdateEntry4("repositorio")&"/"& rsUpdateEntry4("Comprobante")&"' target=comprobante>"&rsUpdateEntry4("Comprobante") &"</a></td>")
                      response.write("<td colspan=5 class='td-l fonttiny'><B>[UUID Pagado:]</B> "& rsUpdateEntry4("UUID_Factura") &"</td>")
                      response.write("<td colspan=3 class='td-r fonttiny'><B>pago por --></B></td><td class='td-r fonttiny'>"& rsUpdateEntry4("ImpPago") &"<td>&nbsp;</td></td><td class='td-r fonttiny'>"& rsUpdateEntry4("SaldoInsoluto") &"</td>")
                   RowOut
                   RowIn  
                        response.write("<td class='subtitulo td-l fonttiny'>XML<br>complto:</td>")
                        if rsUpdateEntry4("Tipo")<>"" then                      
                            Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("Tipo") &"</td>")
                            Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4(campos2(2)) &"</td>") 

                            '--------------------------------descarga xml---------------------------------' 

                           if  rsUpdateEntry4("VariosComplementos")=1 then
                              vString1a=trim(rsUpdateEntry4(campos2(4)))
                            
                              Response.write("<td class='td-r fonttiny'>"  )
                              for i=1 to int(len(vString1a)/8)   'veces'
                                  vpos=inStr(vString1a,",") 
                                  if vpos>0 then
                                     vString2a=mid(vString1a,1,vpos-1) 
                                     vString1a=mid(vString1a,vpos+1,8000)
                                  else
                                     vString2a=mid(vString1a,vpos+1,8000)
                                  end if
                                  'response.write(vString2a&"<BR>")

                                  response.write("<a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&vString2a&"' target='complemento'>"& vString2a &"</a><br>") 
                              next
                              Response.write("</td>")


                          else
                              Response.write("<td class='td-r fonttiny'><a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&rsUpdateEntry4(campos2(4))&"' target='complemento'>"& rsUpdateEntry4(campos2(4)) &"</a></td>") 

                          end if

                          if  rsUpdateEntry4("VariosComplementos")=1 then
                              vString1a=replace(rsUpdateEntry4(campos2(4)),".xml",".pdf")
                              Response.write("<td class='td-r fonttiny'>"  )
                              for i=1 to int(len(vString1a)/8)   'veces'
                                  vpos=inStr(vString1a,",") 
                                  if vpos>0 then
                                     vString2a=mid(vString1a,1,vpos-1) 
                                     vString1a=mid(vString1a,vpos+1,8000)
                                  else
                                     vString2a=mid(vString1a,vpos+1,8000)
                                  end if
                                  'response.write(vString2a&"<BR>")

                                  response.write("<a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&vString2a&"'>"& vString2a &"<img src='/images/FilePDF2.png' border=0></a><br>") 
                              next
                              Response.write("</td>")
    
                          else
                            Response.write("<td class='td-r fonttiny'><a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&replace(rsUpdateEntry4(campos2(4)),".xml",".pdf")&"' target='complemento'>"& replace(rsUpdateEntry4(campos2(4)),".xml",".pdf") &"<img src='/images/FilePDF2.png' border=0></a></td>")   
                          end if

                            for i=5 to 18                  
                              Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4(campos2(i)) &"</td>")                                 
                            next  
                        else
                            Response.write("<td class='td-c fonttiny' colspan=20> - - - - - - - - - -informacion de complemento PENDIENTE   - - - - - - - - - - </td>")
                        end if
                       rsUpdateEntry4.MoveNext
                   wend                   
                   closeTable()
                   response.write("</td>")
                   RowOut

               else
                   NoRegistros
               end if
               rsUpdateEntry4.close



               '=============================================================================================
               response.write("<tr><td colspan=19>&nbsp;</td></tr>")
               separador
               rsUpdateEntry2.MoveNext
               registros=registros+1
         wend
         rsUpdateEntry2.close
         closeTable()

         if FlagMontos=1 then
               response.write("<P class='td-c fonttiny'><input type='submit' value='aplicar pago(s)'> * no utilice comas</P>")
               response.write("<P class='td-c fonttiny'>&nbsp;</P>")

         end if

      if MultiplesPagos=1 then
         if not rsUpdateEntry5.EOF then 
              rsUpdateEntry5.close
              strSQL="SELECT STUFF( (select ','+''''+UUID_Factura+'''' from PAGOSFACTURAS where Comprobante=(select Comprobante from PAGOSFACTURAS WHERE UUID_Factura='"& request("fUUID") &"') and UUID_Factura<>'"& request("fUUID") &"' for XML PATH ('') ),1,1, '' ) as calculo"
              rsUpdateEntry5.Open strSQL, adoCon4

              strSQL="SELECT *,.dbo.fn_GetMonthName(Docdate,'spanish') as fecha,.dbo.fn_GetMonthName(DocDuedate,'spanish') as fechavence from FACTURAS where UUID in ("& rsUpdateEntry5("calculo") &")"
              'response.write strSQL
              rsUpdateEntry5.close
              rsUpdateEntry5.Open strSQL, adoCon4

              response.write("<P class='td-c'><H3>Otras facturas pagadas en esta misma operacion</H3></P>")
              CreateTable(1160)
              RowIn
                For i=1 to 16
                    response.Write("<th class='td-c fonttiny'>" & campos(i) & "</th>")
                next
              RowOut
              while not rsUpdateEntry5.EOF
                   RowIn
                   response.write("<td class='fonttiny td-c'><a href='showContent.asp?contenido=132&mode=2&fUUID="&rsUpdateEntry5("UUID")&"' target='"& rsUpdateEntry5("ID")&"'>"& rsUpdateEntry5("ID") &"</a></td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("RS") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("cardcode") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("LicTradNum") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("TipoComprobante") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("MetodoPago") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("FormaPago") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("Moneda") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("UsoCFDi") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("folio") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("RegimenFiscal") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("version") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("Fecha") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("FechaVence") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("total") &"</td>")
                   response.write("<td class='fonttiny td-c'>"& rsUpdateEntry5("emisor") &"</td>")
                   RowOut
                   rsUpdateEntry5.MoveNext
              wend
              closeTable()

         end if
         rsUpdateEntry5.close

       end if



      
         if request("ID")="" then
             if request("RFC")<>"" then
                 CreateTable(500)
                 rowin
                 response.write("<td class='td-r subtitulo'>FECHA DEL PAGO</td><td class='td-c'><input type='date' name='ffecha'></td>")
                 RowOut
                 rowin
                 response.write("<td class='td-r subtitulo'>Se pag&oacute; con TDC &oacute; <br>un tercero realiz&oacute; el pago</td><td class='td-c'><input type='checkbox' name='fpagado'></td>")
                 RowOut
                 rowin
                     response.write("<td colspan=2>&nbsp;</td>")
                 RowOut
                 RowIn
                 response.write("<td colspan=2 class='td-c'><input type='submit' value='continuar'></form></td>")
                 RowOut
                 closeTable()
             end if
         end if       

              'COSTEO
              if request("action")="2" then
                        CreateTable(1160)
                        RowIn
                        response.write("<td class='td-c subtitulo'>Solicite a CONTABILIDAD revision de su costeo</td>")
                        RowOut
                        SaltoLinea


                        SaltoLinea                        
                        RowIn
                        response.write("<td class='td-c'><input type='submit' value='enviar'>")
                        RowOut
                        closeTable()

              end if
  

              if request("action")="6" then
                        response.write("<P class=alert>Est&aacute; seguro desea borrar este registro?</P>")
                        response.write("<P>&nbsp;</P>")
                        response.write("<P><a href='ShowContent.asp?contenido=132&mode=1&action=7&ID="&request("ID")&"'>[BORRAR]</a></P>")
                        response.write("<P><a href='ShowContent.asp?contenido=132&mode=1'>[CANCELAR]</a></P>")

              end if

              if request("action")="7" then
                        strSQL="DELETE Facturas WHERE ID="&request("ID")
                        response.write strSQL
                        rsUpdateEntry3.Open strSQL, adoCon4
                        response.redirect("ShowContent.asp?contenido=132&msg=una programacion de pago fue eliminada")

              end if



end sub


sub ComplementosRecibidos  'contenido 133'
 
    flag=0
    doAlert
    response.write("<P>&nbsp;</P>")
    response.write("<P class='td-c'>")
    response.write("<a href='ShowContent.asp?contenido=131'>Solicitar pagar</a> | ")
    response.write("<a href='ShowContent.asp?contenido=132&mode=1'>Facts sin pago | </a>")
    response.write("<a href='ShowContent.asp?contenido=134'>Pagos Realizados | </a>")
    response.write("<a href='ShowContent.asp?contenido=132&mode=2'>Facturas Pagadas</a>") 
    response.write("</P>")   

    titulo="Complementos recibidos en <br><font color=blue>complementos@deltaflow.com.mx</font>"
    DoTitulo(titulo)

    response.write("<form id='complementos' method='POST' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value='133'>")

    response.write("<P class='fontmed'><font class='subtitulo'>BUSQUEDA DE COMPLEMENTOS</font><br><br>")
    response.write("<font class='fonttiny'>")
    response.write("[emisor:] <input type='text' name='femisor' value='"&request("femisor")&"' style='width:240px'>")
    response.write("[UUID:] <input type='text' name='fUUID' value='"&request("fUUID")&"' style='width:320px' maxlenght='36'>")
    response.write("[estatus:] <select name='festatus'><option value=''>seleccione</option><option value='0'>no vinculado</option></select>")
    response.write("</font><input type='submit'  value='buscar'></P><br></form>")


    strSQL="select UUID as UUID_PAGO,RFCEmisor,Emisor,RFCRecptor,Receptor,TipoCambioP,MonedaP,FormaDePagoP,ImpPagado,ImpSaldoAnt,ImpSaldoInsoluto,IDDocumento,EmailFrom,NombreArchivo,Match from complementos where 1=1 "

    if len(request("femisor"))>=3 then 
         strSQL=strSQL& " and (RFCEmisor like '%"& request("femisor") &"%' or Emisor like '%"& request("femisor") &"%') "
    end if
    if len(request("fUUID"))>=3 then 
         strSQL=strSQL& " and (UUID like '%"& request("fUUID") &"%' or IdDocumento like '%"& request("fUUID") &"%') "
    end if
    if trim(request("festatus"))="0" then 
         strSQL=strSQL& " and Match=0 "
    end if

    strSQL=strSQL&" order by NombreArchivo desc"
    'response.write strSQL

    Const adCmdText = &H0001
    Const adOpenStatic = 3
    nPage=0
    

    rsUpdateEntry.Open strSQL, adoCon4
    rsUpdateEntry2.Open strSQL, adoCon4, adOpenStatic, adCmdText   

    rsUpdateEntry2.PageSize = 6 
    nPageCount = rsUpdateEntry2.PageCount 

    if not rsUpdateEntry.EOF then
       CreateTable(1200)
               if request("vPage")<>"" then
                    nPage = int(trim(request("vPage")))
                else
                    nPage=1
                end if
              rsUpdateEntry2.AbsolutePage=npage

                response.write("<tr><td class='td-r' width='5%'><B><U>PAGINAS:</U></B></td><td class='td-l' width='95%'>")
                for i=1 to nPageCount step 1
                      if i<>npage then  
                          response.write("<a href='/modules/ShowContent.asp?contenido="& request("contenido") &"&vPage="&i &"'>") 
                          response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
                      end if
                    if i=35 or i=70 or i=105 or i=140 then             
                         response.write("<br>")
                    end if
                next
                response.write("</td></tr>")
       closeTable()
    end if



    Dim Campos(20) 
         i=1         
         For Each rsUpdateEntry in rsUpdateEntry.Fields                        
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next   

    HeaderTableComplementos

    if len(request("fUUID"))>=3 and not rsUpdateEntry2.EOF then
       strSQL="select count(*) as calculo from Complementos where UUID='"& rsUpdateEntry2("UUID_PAGO") &"' and IdDocumento<>'"& rsUpdateEntry2("IdDocumento") &"'"
       'response.write strSQL
       rsUpdateEntry3.Open strSQL, adoCon4

       if int(trim(rsUpdateEntry3("calculo")))>0 then
          flag=1
          strSQL="select UUID as UUID_PAGO,RFCEmisor,Emisor,RFCRecptor,Receptor,TipoCambioP,MonedaP,FormaDePagoP,ImpPagado,ImpSaldoAnt,ImpSaldoInsoluto,IDDocumento,EmailFrom,NombreArchivo,Match from complementos where UUID='"& rsUpdateEntry2("UUID_PAGO") &"' and IdDocumento<>'"& rsUpdateEntry2("IdDocumento") &"'"
       end if
       rsUpdateEntry3.close
    end if

    bodyTableComplementos
    closeTable()
    rsUpdateEntry2.close
    response.write("<P>&nbsp;</P>")
    if flag=1 then
        response.write("<H4>Otras facturas pagadas con el mismo complemento</H4>")
        HeaderTableComplementos
        RowIn
        rsUpdateEntry2.Open strSQL, adoCon4
        bodyTableComplementos
        closeTable()
        rsUpdateEntry2.close
        response.write("<P>&nbsp;</P>")
    end if

end sub


sub HeaderTableComplementos
    CreateTable(1200)     
         RowIn
           response.write("<td class='subtitulo td-c fonttiny'>RFCEmisor</td><td class='subtitulo td-c fonttiny'>Emisor</td><td class='subtitulo td-c fonttiny'>RFCRecptor</td><td class='subtitulo td-c fonttiny'>Receptor</td><td class='subtitulo td-c fonttiny'>TipoCambioP</td><td class='subtitulo td-c fonttiny'>MonedaP</td><td class='subtitulo td-c fonttiny'>FormaDePagoP</td><td class='subtitulo td-c fonttiny'>ImpPagado</td><td class='subtitulo td-c fonttiny'>ImpSaldoAnt</td><td class='subtitulo td-c fonttiny'>ImpSaldoInsoluto</td>")  
         RowOut
end sub

sub bodyTableComplementos
        registros=1

        while not rsUpdateEntry2.EOF  AND registros<=8
             rowin
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("RFCEmisor") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("Emisor") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("RFCRecptor") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("Receptor") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("TipoCambioP") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("MonedaP") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("FormaDePagoP") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("ImpPagado") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("ImpSaldoAnt") &"</td>")
             response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("ImpSaldoInsoluto") &"</td>")
     
             RowOut
             response.write("<tr><td class='td-l fonttiny' colspan=4><B>[UUID DEL PAGO:]</B> <a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&rsUpdateEntry2("NombreArchivo")&"' target='xml'>"&rsUpdateEntry2("UUID_PAGO") &" <img src='/img/descargue.png' border=0 title='descargue xml' alt='descargue xml'></a></td>" )
             response.write("<td class='td-l fonttiny' colspan=6><B>[Remitente:] </B>"& rsUpdateEntry2("EmailFrom")  &"</td></tr>" )
             response.write("<tr><td class='td-l fonttiny' colspan=4><B>[UUID FACTURA PAGADA:]</B> "&rsUpdateEntry2("IDDocumento") &"</td>" )  
                 if  rsUpdateEntry2("Match")="1" then
                    response.write("<td class='td-l fonttiny' colspan=6 style='background-color:#E0FCE0'>VINCULADO</td></tr>")
                 else
                    response.write("<td class='td-l fonttiny' colspan=6 style='background-color:#F9F9DF'>PDTE VINCULAR</td></tr>")
                 end if      
             rsUpdateEntry2.MoveNext
             registros=registros+1
             response.write("<tr><td class='td-l fonttiny' colspan=10>&nbsp;</td></tr>" )  
             separador
             
        wend                      
end sub


sub ControlPagosEfectuados 'contenido 134  

       doAlert
       response.write("<P>&nbsp;</P>")      
       response.write("<P class='td-c'>")
       response.write("<a href='ShowContent.asp?contenido=131'>Solicitar pagar</a> | ")
       response.write("<a href='ShowContent.asp?contenido=132&mode=1'>Facts sin pago | </a>")
       response.write("<a href='ShowContent.asp?contenido=132&mode=2'>Facturas Pagadas</a> | ") 
       response.write("<a href='ShowContent.asp?contenido=133'>Complementos recibidos</a>")  
       response.write("</P>")    

       titulo="Pagos Realizados"
       DoTitulo(titulo)      

       response.write("<form id='complementos' method='POST' action='ShowContent.asp'>")
       response.write("<input type='hidden' name='contenido' value='134'>")

        response.write("<P class='fontmed'><br>")
        response.write("<font class='fonttiny'>")
        response.write("[Nombre Comprobante:] <input type='text' name='fcomprobante' value='"&request("fcomprobante")&"' style='width:240px'>") 
        response.write("[UUID:] <input type='text' name='fUUID' value='"&request("fUUID")&"' style='width:300px'>")
        response.write("[filtro:] <select name='ffiltro'><option value=''>todas</option><option value='1'>sin complemento</option></select>")
        response.write("</font><input type='submit'  value='buscar'></P><br></form>")


       strSQL="select TipoComprobante as Tipo, Folio, UUID, xml as PDF, version, RFCEmisor, Emisor,  RegimenFiscal as Regimen, RFCRecptor,  Receptor,IdDocumento,  TipoCambioP as TC, MonedaP, FormaDePagoP as Forma,  ImpPagado, ImpSaldoAnt, ImpSaldoInsoluto,  NumParcialidad as Parcialidad, Comprobante, UUID_Factura,ImpPago,SaldoInsoluto     from PagosFacturas where 1=1 "

       if request("comprobante")<>""  then strSQL=strSQL&" and Comprobante='"& request("Comprobante") &"' "
       if request("fcomprobante")<>"" then strSQL=strSQL&" and Comprobante like '%"& request("fComprobante") &"%' "
       if request("fUUID")<>""        then strSQL=strSQL&" and UUID_Factura like '%"& request("fUUID") &"%' "
       if request("ffiltro")="1"      then strSQL=strSQL&" and IdDocumento is null "

       Const adCmdText = &H0001
       Const adOpenStatic = 3
       nPage=0

       'response.write strSQL
       rsUpdateEntry.Open strSQL, adoCon4
       rsUpdateEntry2.Open strSQL, adoCon4, adOpenStatic, adCmdText 

       rsUpdateEntry2.PageSize = 10 
       nPageCount = rsUpdateEntry2.PageCount 
       
       CreateTable(1160)
       if not rsUpdateEntry.EOF then
       
               if request("vPage")<>"" then
                    nPage = int(trim(request("vPage")))
                else
                    nPage=1
                end if
                rsUpdateEntry2.AbsolutePage=npage

                response.write("<tr><td class='td-r' colspan=2><B><U>PAGINAS:</U></B></td><td class='td-l' colspan=17>")
                for i=1 to nPageCount step 1
                      if i<>npage then  
                          response.write("<a href='/modules/ShowContent.asp?contenido="& request("contenido") &"&vPage="&i &"'>") 
                          response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
                      end if
                    if i=35 or i=70 or i=105 or i=140 then             
                         response.write("<br>")
                    end if
                next
                response.write("</td></tr>")       
       end if

       Dim Campos(30) 
         i=1
         
         For Each rsUpdateEntry in rsUpdateEntry.Fields                        
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next      

         RowIn
                Response.Write("<td class='td-c'>&nbsp;</td>")
           for i=1 to 17
                Response.Write("<td class='subtitulo td-c'>" & campos(i) & "</td>")  
           next    
                Response.Write("<td class='subtitulo td-c'>controles</td>")  
         RowOut

         registros=1
         while not rsUpdateEntry2.EOF AND registros<=10            
             Rowin
                response.write("<td class='subtitulo td-l fonttiny'>tesoreria:</td>")
                response.write("<td colspan=6 class='td-l fonttiny'><B>[Comprobante:]</B><a href='http://192.168.0.206/repositorio/ADMON/PagoFacturas/"&rsUpdateEntry2("Comprobante")&"' target=comprobante>"& rsUpdateEntry2("Comprobante") &"</a></td>")
                response.write("<td colspan=5 class='td-l fonttiny'><B>[UUID Factura:]</B> "& rsUpdateEntry2("UUID_Factura") &"</td>")
                response.write("<td colspan=3 class='td-r fonttiny'><B>pago--></B></td><td class='td-r fonttiny'>"& rsUpdateEntry2("ImpPago") &"<td>&nbsp;</td></td><td class='td-r fonttiny'>"& rsUpdateEntry2("SaldoInsoluto") &"</td>")

                response.write("<td class='td-c fonttiny'><a href='showContent.asp?contenido=134&action=1&fcomprobante="&rsUpdateEntry2("Comprobante")&"&UUID="&rsUpdateEntry2("UUID_Factura")&"'><img src='/img/email.png' alt='re-envie email del pago' title='re-envie email del pago'></a></td>")

             RowOut
             RowIn  
                  response.write("<td class='subtitulo td-l fonttiny'>XML<br>vinculado:</td>")
                  if rsUpdateEntry2("Tipo")<>"" then                      
                      Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("Tipo") &"</td>")
                      Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(2)) &"</td>")  
                      Response.write("<td class='td-r fonttiny'><a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&rsUpdateEntry2(campos(4))&"' target='complemento'>"& rsUpdateEntry2(campos(3)) &"</a></td>") 

                      Response.write("<td class='td-r fonttiny'><a href='http://192.168.0.206/repositorio/ADMON/ComplementosDescargados/"&replace(rsUpdateEntry2(campos(4)),".xml",".pdf")&"' target='complemento'>"& replace(rsUpdateEntry2(campos(4)),".xml",".pdf") &"</a></td>")   
                      for i=5 to 18                  
                        Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")                                 
                      next  
                  else
                      Response.write("<td class='td-c fonttiny' colspan=20> - - - - - - - - - -  p e n d i e n t e  - - - - - - - - - - </td>")
                  end if

                     
             RowOut
             rsUpdateEntry2.MoveNext
             registros=registros+1
             response.write("<tr><td class='td-l fonttiny' colspan=30>&nbsp;</td></tr>" )  
             separador             
         wend         
         rsUpdateEntry2.close

         closeTable()
         response.write("<P>&nbsp;</P>")


         if request("action")="1" then
               strSQL="select * from FACTURAS where UUID='"& request("UUID") &"' "
               rsUpdateEntry4.Open strSQL, adoCon4

               strSQL="exec SP_CopiaArchivosFacturas @ID="&rsUpdateEntry4("ID") &",@Mode=2"
               rsUpdateEntry4.close
               'rsUpdateEntry5.Open strSQL, adoCon4
               response.write strSQL

               'response.redirect("ShowContent.asp?contenido=134&msg=el comprobante de pago fue enviado por email&comprobante="&request("fComprobante"))
         end if

end sub




sub CostoContableOnStock  'contenido 199'

    doAlert
    titulo="Costo Contable (en USD) Partidas con Stock"
    DoTitulo(titulo)

    if request("action")="" then
        response.write("<form id='RS' action='ShowContent.asp' method='POST'><input type='hidden' name='contenido' value=199><input type='hidden' name='action' value=1>")

        response.write("<P class='td-c'>RS donde se encuentra el stock: <select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option><option value='MME'>MME</option></select><br><br>")

        response.write("Almac&eacute;n: <select name='fALM'><option value='001-Mxl'>MXL</option><option value='002-SJR'>SJR</option><option value='003-Mty'>MTY</option></select><br><br><br>")

        response.write("<input type='submit' value='continuar'></P></form>")
    end if

    if request("action")=1 then 

        response.write("<form id='RS' action='ShowContent.asp' method='POST'>")
        response.write("<input type='hidden' name='contenido' value=199>")
        response.write("<input type='hidden' name='action' value=2>") 
        response.write("<input type='hidden' name='fRS' value='"& request("fRS") &"'>")
        response.write("<input type='hidden' name='fALM' value='"& request("fALM") &"'>")

        response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='330px'>")
        response.write("<tr><td class='td-r subtitulo'>Razon Social Origen: </td><td class='td-l'>" & request("fRS") &"</td></tr>")   
         response.write("<tr><td class='td-r subtitulo'>Almac&eacute;n: </td><td class='td-l'>" & request("fALM") &"</td></tr>")   
        response.write("</table>")
        response.write("<P style='font-size: 4px'>&nbsp;</P><center><table border=1 cellspacing=2 cellpadding=3 align=center width='500px'>")
        strSQL="select Rate,RateDate from ORTT where year(RateDate)=year(getdate()) and month(RateDate)=month(getdate()) and day(RateDate)=day(getdate())"
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon2 'DMX'

        vRate=""
        if rsUpdateEntry.EOF then
                vRate="20"
        else        
                vRate=trim(rsUpdateEntry("Rate"))
        end if
        rsUpdateEntry.close
        response.write("<tr><td class='td-r subtitulo'>El TC para el d&iacute;a es: <input type='text' name='frate' value='"& vRate &"'></td></tr></table>")
        response.write("<P class='td-c'>Pegue el listado de sus partidas en esta &aacute;rea:<BR><textarea rows=30 cols=25 name='fpartidas'></textarea></P>")


        response.write("<P class='td-c'><BR><input type='submit' value='continuar'></P></form>")

   end if


   if request("action")=2 then   'usuario ya pegó el listado de partidas'
      
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
        vPos=len(strSQL)
        'response.write  mid(strSql,1,vPos-1)

        strSQL="SELECT ItemCode as Codigo,round(AvgPrice/"&request("frate")&",2) as CostoContable FROM PRODUCTIVA_"&request("fRS")&".dbo.OITW  WHERE WhsCode='"&request("fALM")&"' and ItemCode in ( select a.ItemCode  from PRODUCTIVA_"&request("fRS")&".dbo.OINM a  WHERE  WhsCode='"&request("fALM")&"' and  a.ItemCode in ("&mid(strSql,1,vPos-1)&")   GROUP BY a.itemcode having SUM(a.InQty)-SUM(a.OutQty) >0  ) "

        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon4
        rsUpdateEntry2.Open strSQL, adoCon4

        Dim Campos(20) 
             i=1

             response.write("<table border=1 cellspacing=2 cellpadding=3 width='500px' class='table2excel table2excel_with_colors' data-tableName='table1'>")  

             RowIn
             for Each rsUpdateEntry in rsUpdateEntry.Fields  
                 campos(i)=rsUpdateEntry.Name     
                 Response.Write("<th class='subtitulo td-c'>" & campos(i) & "</th>")                
                 i=i+1           
             Next        
             RowOut             

             while not rsUpdateEntry2.EOF 
                RowIn
                for x=1 to 2
                     response.write("<td class='td-r fontbig'>" & rsUpdateEntry2(campos(x)) &"</td>")
                next
                RowOut
                rsUpdateEntry2.MoveNext
             wend   
             rsUpdateEntry2.close   %>

              </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp;   

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <%
             

     end if



end sub

sub IngresosMesSinEmail 'contenido 135'   
       doAlert()
       SaltoLinea
       titulo="Procese los ingresos del mes <BR>(no envia email, solo procesa los ingresos)"
       doTitulo(titulo)
       SaltoLinea

       response.write("<P class='td-c fontbig'><a href='ShowContent.asp?contenido=135&action=1'>click para procesar manualmente los ingresos del mes</a></P>")
       
       if request("action")="1" then
            strSQL="exec SP_Historico_ventas 0"
            rsUpdateEntry2.Open strSQL, adoCon4

            response.redirect("/modules/ShowContent.asp?contenido=6&msg=se procesaron manualmente los ingresos")

       end if

end sub



sub ProgramacionPagos  'contenido 136

      doAlert
      titulo="Programacion de Pagos"
      doTitulo(titulo)

      response.write("<form id='facturas' method='POST' action='ShowContent.asp'>")
      response.write("<input type='hidden' name='contenido' value='136'>")      

      strSQL="select ID,.dbo.fn_GetMonthName(LogDate,'spanish') as Fecha,RS,cardcode,substring(Emisor,1,30) Emisor,Notas,Folio,UUID,.dbo.fn_GetMonthName(DocDate,'spanish') FechaFact,.dbo.fn_GetMonthName(DocDueDate,'spanish') FachaVence,UsoCFDI,total,Moneda  from Facturas  where FechaPagada is null "

      if request("ffecha")<>"" then strSQL=strSQL & " and LogDate>='"& request("ffecha") &"' "
      strSQL=strSQL & " order by ID desc "

      'response.write(strSQL&"<BR><BR><BR>")
      rsUpdateEntry.Open strSQL, adoCon4   
      rsUpdateEntry2.Open strSQL, adoCon4

       response.write("<P class='td-c'>[no antes de esta fecha:] <input type='date' name='ffecha'> <input type='submit' value='aplicar filtro'></form></P>")

       response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='1400px' class='table2excel table2excel_with_colors' data-tableName='table1'>")
 
      Dim Campos(20)
      i=1
      rowin
       For Each rsUpdateEntry in rsUpdateEntry.Fields
              Response.Write("<td class='subtitulo td-l'>" & rsUpdateEntry.Name & "</td>")
              campos(i)=rsUpdateEntry.Name
              i=i+1
       Next 
      rowout


    while not rsUpdateEntry2.EOF
       rowin       
       for i=1 to 13            
            response.write("<td class='td-l fonttiny'>" & rsUpdateEntry2( campos(i) ) &"</td>")
       next
       'response.write("<td class='td-l fonttiny'>" & rsUpdateEntry2("Moneda") &"</td>")
       rowout
       rsUpdateEntry2.MoveNext
    wend
    rsUpdateEntry2.close
    
         response.write("</table>")   %>  <button class="exportToExcel">Exportar a excel</button> </center>

   <script>
              $(function() {
                $(".exportToExcel").click(function(e){
                  var table = $(this).prev('.table2excel');
                  if(table && table.length){
                    var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
                    $(table).table2excel({
                      exclude: ".noExl",
                      name: "Excel Document Name",
                      filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
                      fileext: ".xls",
                      exclude_img: true,
                      exclude_links: true,
                      exclude_inputs: true,
                      preserveColors: preserveColors
                    });
                  }
                });
                
              });
    </script>

   <%


 end sub 



sub CostoVentaTuberias  'contenido 140

      doAlert
      response.write("<P>&nbsp;</P>")
      titulo="Precio base cotizacion en tuberias " & request("fcodigo")
      doTitulo(titulo)

      strSQL="select OITM.itemcode,OITM.itemname,AUX.Cantidad from OITM  inner join ( select a.itemcode,a.itemname,    (select isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DFW.dbo.OINM WITH (NOLOCK)  where ItemCode=a.ItemCode),0) + isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DMX.dbo.OINM WITH (NOLOCK)  where ItemCode=a.ItemCode),0) ) as Cantidad      from OITM a       where substring(a.itemcode,1,3)='MEC' and substring(a.itemname,1,3)='TUB'  ) as AUX on OITM.ItemCode=AUX.ItemCode where AUX.Cantidad>0    "
      rsUpdateEntry.Open strSQL, adoCon2

      response.write("<form id='CostoVenta' method='POST' action='ShowContent.asp'>")
      response.write("<input type='hidden' name='contenido' value=140>")

      if request("action")="" then
            response.write("<input type='hidden' name='action' value=1>")

            CreateTable(800)
            RowIn
            response.write("<td class='td-r subtitulo' width='30%'>Seleccione c&oacute;digo:</td>")
            response.write("<td class='td-l' width='30%'><select name='fcodigo'>")

            while not rsUpdateEntry.EOF
                response.write("<option value='"&rsUpdateEntry("ItemCode")&"'>"&rsUpdateEntry("ItemCode")&" | "& rsUpdateEntry("ItemName") &"</option>")
                rsUpdateEntry.MoveNext
            wend
            rsUpdateEntry.close
            response.write("</select> </td>")
            RowOut
         
             response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='continuar'></form></td></tr></table>")

        end if

        if request("action")="1" then
            response.write("<input type='hidden' name='action' value=2>")
            strSQL="exec SP_CostoVentaTuberias '"&request("fcodigo") &"'"
            'response.write strSQL
            rsUpdateEntry7.Open strSQL, adoCon4

            strSQL="select * from _CostoVentaTub"
            'response.write strSQL
            rsUpdateEntry2.Open strSQL, adoCon4

            CreateTable(500)
            rowin
            response.write("<td class='td-c subtitulo'>almacen</td><td class='td-c subtitulo'>stock</td><td class='td-c subtitulo'>Precio <br>stock</td><td class='td-c subtitulo'>nuevo precio<br>de compra</td><td class='td-c subtitulo'>Cantidad requerida</td>")
            RowOut
            RowIn    
            response.write("<td class='td-r subtitulo'>MXL</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("MXL")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MXL")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MXL_N")&"</td>")
            response.write("<td class='td-c'><input type='number' name='fmxl'></td>")
            RowOut
            RowIn    
            response.write("<td class='td-r subtitulo'>SJR</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("SJR")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_SJR")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_SJR_N")&"</td>")
            response.write("<td class='td-c'><input type='number' name='fsjr'></td>")
            RowOut
            RowIn    
            response.write("<td class='td-r subtitulo'>MTY</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("MTY")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MTY")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MTY_N")&"</td>")
            response.write("<td class='td-c'><input type='number' name='fmty'></td>")
            RowOut
            RowIn
            response.write("<td colspan=5>&nbsp;</td>")
            RowOut
            Rowin
            response.write("<td colspan=4>&nbsp;</td>")
            response.write("<td class='td-c'><input type='submit' value='Mostrar Precio Promedio'></form</td>")
            RowOut
            closeTable()
        end if

        if request("action")="2" then
            vMxl="0"
            vSJR="0"
            vMTY="0"
            if request("fmxl")="" and  request("fsjr") ="" and request("fmty") ="" then response.redirect("ShowContent.asp?contenido=140&msg=no indico cantidades")
   
            if request("fmxl")<>"" then vMxl=request("fmxl")
            if request("fsjr")<>"" then vSjr=request("fsjr")
            if request("fmty")<>"" then vMty=request("fmty")

            strSQL="UPDATE _CostoVentaTub set MXL_REQ="&vMxl &",SJR_REQ="&vSjr&",MTY_REQ="&vMty
            'response.write strSQL
            rsUpdateEntry3.Open strSQL, adoCon4

            response.redirect("showContent.asp?contenido=140&action=3")
        end if

        if request("action")="3" then

            strSQL="select *,case when MXL_REQ>MXL then round( (MXL*CV_MXL+(MXL_REQ-MXL)*CV_MXL_N)/MXL_REQ ,2) else CV_MXL end MXL_AVR,         case when SJR_REQ>SJR then round( (SJR*CV_SJR+(SJR_REQ-SJR)*CV_SJR_N)/SJR_REQ ,2) else CV_SJR end SJR_AVR,     case when MTY_REQ>MTY then round( (MTY*CV_MTY+(MTY_REQ-MTY)*CV_MTY_N)/MTY_REQ ,2) else CV_MTY end MTY_AVR from _CostoVentaTub "
            'response.write strSQL
            rsUpdateEntry2.Open strSQL, adoCon4

            CreateTable(600)
            rowin
            response.write("<td class='td-c subtitulo'>almacen</td><td class='td-c subtitulo'>stock</td><td class='td-c subtitulo'>Precio <br>stock</td><td class='td-c subtitulo'>nuevo precio<br>de compra</td><td class='td-c subtitulo'>Cantidad requerida</td><td class='td-c subtitulo'>Precio Promedio</td>")
            RowOut
            RowIn    
            response.write("<td class='td-r subtitulo'>MXL</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("MXL")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MXL")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MXL_N")&"</td>")
            response.write("<td class='td-c'>"&rsUpdateEntry2("MXL_REQ")&"</td>")
            response.write("<td class='td-c'>"&rsUpdateEntry2("MXL_AVR")&"</td>")
            RowOut
            RowIn    
            response.write("<td class='td-r subtitulo'>SJR</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("SJR")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_SJR")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_SJR_N")&"</td>")
            response.write("<td class='td-c'>"&rsUpdateEntry2("SJR_REQ")&"</td>")
            response.write("<td class='td-c'>"&rsUpdateEntry2("SJR_AVR")&"</td>")
            RowOut
            RowIn    
            response.write("<td class='td-r subtitulo'>MTY</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("MTY")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MTY")&"</td>")
            response.write("<td class='td-r'>"&rsUpdateEntry2("CV_MTY_N")&"</td>")
            response.write("<td class='td-c'>"&rsUpdateEntry2("MTY_REQ")&"</td>")
            response.write("<td class='td-c'>"&rsUpdateEntry2("MTY_AVR")&"</td>")
            RowOut            
            closeTable()

        end if
end sub



sub FechasEmbarque 'contenido 139'

   response.write("<P>&nbsp;</P>")
   doAlert
   titulo="Actualizar fechas de embarque "&request("fpedido")
   DoTitulo(titulo)

   response.write("<form id='inv' method='POST' action='ShowContent.asp'>")
   response.write("<input type='hidden' name='contenido' value=139>")
       
    if request("action")="" then 
         response.write("<input type='hidden' name='action' value=1>")

         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option><option value='MEIDE'>MME</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Pedido: <input class='anchoSmall' type='numeric' name='fpedido' value="&request("fpedido") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

    end if

    if request("action")="1" then 
     
         if request("lineNum")<>"" then
             response.write("<input type='hidden' name='action' value=2>")
             response.write("<input type='hidden' name='fRS' value="&request("fRS")&">")
             response.write("<input type='hidden' name='fpedido' value="&request("fpedido")&">")
             response.write("<input type='hidden' name='linenum' value="&request("lineNum")&">")

         else
             response.write("<input type='hidden' name='action' value=1>")
         end if

         HeaderPedido

         CreateTable(1000)
         rowin
         response.write("<th class='subtitulo td-c'>#</th><th class='subtitulo td-c'>Codigo</th><th class='subtitulo td-c'>Descripcion</th><th class='subtitulo td-c'>Estatus</th><th class='subtitulo td-c'>Pedida</th><th class='subtitulo td-c'>Pendiente</th><th class='subtitulo td-c'>Almacen</th><th class='subtitulo td-c'>Fecha Embarque</th>")
         RowOut         

         n=1 
         while not rsUpdateEntry.EOF
             if n=1 then
                  vString=" style='background-color:white;' "
             else
                  vString=" style='background-color:#E1DFDF;!Important' "
             end if
             
             rowin
             response.write("<td class='td-c' "& vString &">" & rsUpdateEntry("lineNum")+1 &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & rsUpdateEntry("ItemCode") &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & left(rsUpdateEntry("Dscription"),80) &"</td>")
             if trim(rsUpdateEntry("LineStatus"))="O" then
                  response.write("<td class='td-c fontmed' "& vString &"><font style='color:white;background-color:red'>abierta</font</td>")
             else
                  response.write("<td class='td-c fontmed' "& vString &">cerrada</td>")
             end if
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("Quantity") &"</td>")

             if  int(trim(rsUpdateEntry("OpenQty")))>0 then
                 response.write("<td class='td-c' "& vString &"><font color='red'>" & rsUpdateEntry("OpenQty") &"</font></td>")
             else
                 response.write("<td class='td-c' "& vString &"><font color='green'>" & rsUpdateEntry("OpenQty") &"</font></td>")
             end if
             response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("WhsCode") &"</td>")
             
             if request("lineNum")<>"" then
                  response.write("<td class='td-c fontmed' "& vString &"><input type='date' name='fembarque'></td>")      
             else
                 if trim(rsUpdateEntry("U_FECHALMACEN"))<>"" then
                    response.write("<td class='td-c fontmed' "& vString &"><a href='ShowContent.asp?contenido=139&action=1&fRS="&request("fRS")&"&fpedido="&request("fPedido")&"&lineNum="&rsUpdateEntry("lineNum")&"   '>" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
                     
                 else   
                     response.write("<td class='td-c fontmed' "& vString &"><a href='ShowContent.asp?contenido=139&action=1&fRS="&request("fRS")&"&fpedido="&request("fPedido")&"&lineNum="&rsUpdateEntry("lineNum")&"   '>agregar fecha</td>")
                    
                 end if
             end if

             rsUpdateEntry.MoveNext
             n=n+1
             if n=3 then
                 n=1
             end if

             RowOut
         wend
         rsUpdateEntry.close
         response.write("<tr><td class='td-c' colspan=9>&nbsp;</td></tr>")
         response.write("<tr><td class='td-c' colspan=7>&nbsp;</td><td class='td-c'><input type='submit' value='actualizar'></form></td></tr></table>")
    end if

    if request("action")="2" then 

        if request("fembarque")="" then response.redirect("ShowContent.asp?contenido=139&action=1&fRS="&request("fRS")&"&fpedido="&request("fpedido")&"&msg=envio una fecha en blanco")

        strSQL="select DATEDIFF(DAY,cast(GETDATE() as date),'"&request("fembarque")&"') as calculo"
        rsUpdateEntry5.Open strSQL, adoCon4

        if int(trim(rsUpdateEntry5("calculo")))<1 then
              rsUpdateEntry5.close
              response.redirect("ShowContent.asp?contenido=139&action=1&fRS="&request("fRS")&"&fpedido="&request("fpedido")&"&msg=registre fechas a futuro")
        else
              rsUpdateEntry5.close
        end if

        strSQL="select DATEPART(dw,'"&request("fembarque")&"') as calculo"
        rsUpdateEntry5.Open strSQL, adoCon4
         
        if int(trim(rsUpdateEntry5("calculo")))=7 or int(trim(rsUpdateEntry5("calculo")))=1 then 
             rsUpdateEntry5.close
             response.redirect("ShowContent.asp?contenido=139&action=1&fRS="&request("fRS")&"&fpedido="&request("fpedido")&"&msg=no registre en sabado o domingo")
        else
              rsUpdateEntry5.close
        end if
       
    
        strSQL="UPDATE b SET b.U_FECHALMACEN='"&request("fembarque") &"' from PRODUCTIVA_"&request("fRS")&".dbo.ORDR a inner join PRODUCTIVA_"&request("fRS")&".dbo.RDR1 b on a.DocEntry=b.DocEntry where a.docnum="&request("fpedido")&" and b.LineNum="&request("lineNum")
        response.write strSQL
        rsUpdateEntry6.Open strSQL, adoCon4

        response.redirect("ShowContent.asp?contenido=139&action=1&fRS="&request("fRS")&"&fpedido="&request("fpedido")&"&msg=se ha actualizado una fecha de embarque")

    end if

end sub

'=========================================================================='
%>



     

       