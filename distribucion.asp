<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->

<script src="/config/amcharts.js" type="text/javascript"></script>
<script src="/config/serial.js" type="text/javascript"></script>
<script src="/config/amstock.js" type="text/javascript"></script>


<%
 if  not Session("session_modulo") ="DistribucionCuentas" then
   response.redirect("/home.asp?msg=vuelva a inciar sesion")
 end if

 Dim asesor(6)
    asesor(1)="Martin Gibrann Parra Leal"
    asesor(2)="Alicia Mayorquin"
    asesor(3)="L.B.E. Julio Escobar"
    asesor(4)="Reynaldo Cardona"
    asesor(5)="Héctor Escobar"
    asesor(6)="Daniel Osvaldo Vazquez Trejo"


  Dim trimestres(4)
    trimestres(1)="(1,2,3)"
    trimestres(2)="(4,5,6)"
    trimestres(3)="(7,8,9)"
    trimestres(4)="(10,11,12)"


    strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_Sales' AND TABLE_SCHEMA = 'dbo')     drop table dbo._sales;"
    rsUpdateEntry.Open strSQL, adoCon4  'SBO'


    strSQL="CREATE TABLE [_Sales] ( [Asesor] [varchar](300) NOT NULL, [Cardname] [varchar](1000) NOT NULL, [Subtotal] decimal(18,2) NULL, [Pareto] decimal(18,4) NULL , [Anio] [int] NOT NULL )  ON [PRIMARY]"
    rsUpdateEntry2.Open strSQL, adoCon4  'SBO'


    strSQL="insert into _Sales select asesor,cardname,sum(subtotal) as subtotal,cast(sum(subtotal) as float)/cast((select sum(subtotal) from HistoricoVentas2018) as float)  as Pareto,2018 as Anio  from HistoricoVentas2018   group by asesor,cardname   having sum(subtotal)>=((select sum(subtotal) from HistoricoVentas2018)*.01) UNION  select asesor,cardname,sum(subtotal) as subtotal,cast(sum(subtotal) as float)/cast((select sum(subtotal) from HistoricoVentas2019) as float)  as Pareto,2019   from HistoricoVentas2019 group by asesor,cardname having sum(subtotal)>=((select sum(subtotal) from HistoricoVentas2019)*.01)    UNION  select asesor,cardname,sum(subtotal) as subtotal,cast(sum(subtotal) as float)/cast((select sum(subtotal) from HistoricoVentas2020) as float)  as Pareto,2020  from HistoricoVentas2020 group by asesor,cardname having sum(subtotal)>=((select sum(subtotal) from HistoricoVentas2020)*.01)   UNION  select asesor,cardname,sum(subtotal) as subtotal,cast(sum(subtotal) as float)/cast((select sum(subtotal) from HistoricoVentas2021) as float)  as Pareto,2021  from HistoricoVentas2021 group by asesor,cardname having sum(subtotal)>=((select sum(subtotal) from HistoricoVentas2021)*.01)"
    'response.write strSQL
    rsUpdateEntry3.Open strSQL, adoCon4  'SBOTemp'    


    Const adCmdText = &H0001
    Const adOpenStatic = 3
    strSQL="select distinct(cardname) as cardname,asesor,sum(pareto) as Pareto from _sales  where asesor<>'Otros asesores' group by cardname,asesor order by Pareto DESC"
    rsUpdateEntry5.Open strSQL, adoCon4, adOpenStatic, adCmdText

    nTotalGraficas=rsUpdateEntry5.RecordCount
    response.write("<input type='hidden' id='nTotalGraficas' value="&nTotalGraficas&">")

    if request("cardname")<>"" then
         vcardname=request("cardname")
    else
        vcardname=rsUpdateEntry5("cardname")
    end if

   
    i=1
    for anio=2018 to 2021
            for P=1 to 4
                strSQL="select sum(Subtotal) as calculo from HistoricoVentas"&anio&" where cardname='"& vcardname &"' and month(DocDate) in "& trimestres(P)
                'response.write(strSQL &"<BR>")
                rsUpdateEntry6.Open strSQL, adoCon4

                response.write("<input type='hidden' id='variable"&i&"' value='"& trim(rsUpdateEntry6("calculo")) &"'>")  
                rsUpdateEntry6.close               
                i=i+1
            next
    next
      
 

    titulo="HIST&Oacute;RICO DE VENTAS 2018-2021"
    dotitulo(titulo)
    'response.write("<P>&nbsp;</P><H1><a href='distribucion.asp' class='H1'></a></H1>")
    'response.write("<P>&nbsp;</P>")

    strSQL="select distinct (asesor) as asesor from HistoricoVentas2018  UNION  select distinct (asesor) as asesor from HistoricoVentas2019  where asesor not in (select asesor from HistoricoVentas2018)  UNION  select distinct (asesor) as asesor from HistoricoVentas2020  where asesor not in (select asesor from HistoricoVentas2018) and asesor not in (select asesor from HistoricoVentas2019) UNION  select distinct (asesor) as asesor from HistoricoVentas2021  where asesor not in (select asesor from HistoricoVentas2019) and asesor not in (select asesor from HistoricoVentas2020)  "

    'response.write strSQL

    rsUpdateEntry7.Open strSQL, adoCon4
    CreateTable(900)

    Dim Total_anio(4)
    Dim Total_anio_(4)
   

    n=1
    for i=2018 to 2021 step 1          
          strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as calculo2,sum(subtotal) as calculo from HistoricoVentas"&i 
          rsUpdateEntry4.Open strSQL, adoCon4 
          Total_anio(n)=trim(rsUpdateEntry4("calculo"))
          Total_anio_(n)=trim(rsUpdateEntry4("calculo2"))  'en formato de moneda'
          'response.write("<br>" & Total_anio(n) &"-" & Total_anio_(n) &"<BR>" )
          rsUpdateEntry4.close
          n=n+1
    next

    response.write("<tr>")
    response.write("<td class='subtitulo td-c' width='13%'>Asesor</td><td class='subtitulo td-c' width='10%'>2018</td><td class='subtitulo td-c' width='4%'>%</td><td class='subtitulo td-c' width='10%'>2019</td><td class='subtitulo td-c' width='4%'>%</td><td class='subtitulo td-c' width='10%'>2020</td><td class='subtitulo td-c' width='4%'>%</td><td class='subtitulo td-c' width='10%'>2021</td><td class='subtitulo td-c' width='4%'>%</td><td class='subtitulo td-c'  width='11%'>Total</td></tr>")

    
    while not rsUpdateEntry7.EOF
        response.write("<tr><td class='subtitulo td-r'>" & rsUpdateEntry7("asesor") &"</td>")

        strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as calculo,sum(subtotal) as calculo2,round( cast( sum(subtotal) as float)/cast("& Total_anio(1) &" as float) , 2) as Pareto from HistoricoVentas2018 where asesor='"& rsUpdateEntry7("asesor") &"' "
        'response.write(strSQL &"<BR>")
        rsUpdateEntry.Open strSQL, adoCon4

        strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as calculo,sum(subtotal) as calculo2,round( cast( sum(subtotal) as float)/cast("& Total_anio(2) &" as float) , 2) as Pareto  from HistoricoVentas2019 where asesor='"& rsUpdateEntry7("asesor") &"' "
        rsUpdateEntry2.Open strSQL, adoCon4
        'response.write(strSQL &"<BR>")

        strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as calculo,sum(subtotal) as calculo2,round( cast( sum(subtotal) as float)/cast("& Total_anio(3) &" as float) ,2) as Pareto  from HistoricoVentas2020 where asesor='"& rsUpdateEntry7("asesor") &"' "
        rsUpdateEntry3.Open strSQL, adoCon4   
        'response.write(strSQL &"<BR>")  

        strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as calculo,sum(subtotal) as calculo2,round( cast( sum(subtotal) as float)/cast("& Total_anio(4) &" as float) ,2) as Pareto  from HistoricoVentas2021 where asesor='"& rsUpdateEntry7("asesor") &"' "
        rsUpdateEntry4.Open strSQL, adoCon4   
        'response.write(strSQL &"<BR>")  

        %>
        <td class='td-r'>
             <A href="Javascript:sendReq('DistCuentas.asp','anio,asesor,mode','2018,<%=rsUpdateEntry7("asesor")%>,1','asesor')"> <%=rsUpdateEntry("calculo")%> </a></td>  

         <td class='td-r'><%=rsUpdateEntry("Pareto")%></td>
        <td class='td-r'>
            <A href="Javascript:sendReq('DistCuentas.asp','anio,asesor,mode','2019,<%=rsUpdateEntry7("asesor")%>,1','asesor')"> <%=rsUpdateEntry2("calculo")%> </a></td>  
          <td class='td-r'><%=rsUpdateEntry2("Pareto")%></td>
        <td class='td-r'>
            <A href="Javascript:sendReq('DistCuentas.asp','anio,asesor,mode','2020,<%=rsUpdateEntry7("asesor")%>,1','asesor')"> <%=rsUpdateEntry3("calculo")%> </a></td>  
               <td class='td-r'><%=rsUpdateEntry3("Pareto")%></td>

        <td class='td-r'>
            <A href="Javascript:sendReq('DistCuentas.asp','anio,asesor,mode','2021,<%=rsUpdateEntry7("asesor")%>,1','asesor')"> <%=rsUpdateEntry4("calculo")%> </a></td>  
               <td class='td-r'><%=rsUpdateEntry4("Pareto")%></td>

        
        <%
        strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum((   cast(isnull('"& rsUpdateEntry("calculo2") &"',0) as float) +  cast(isnull('"& rsUpdateEntry2("calculo2") &"',0) as float) +   cast(isnull('"& rsUpdateEntry3("calculo2") &"',0) as float)   +   cast(isnull('"& rsUpdateEntry4("calculo2") &"',0) as float)      ))),1) as calculo "
        'response.write strSQL
        rsUpdateEntry6.Open strSQL, adoCon4  
        %>
          <td class='td-r'>
             <A href="Javascript:sendReq('DistCuentas.asp','total,asesor,mode','<%=replace(trim(rsUpdateEntry6("calculo")),",","")%>,<%=rsUpdateEntry7("asesor")%>,2','asesor')"> <%=rsUpdateEntry6("calculo")%> </a></td>  
        <%
        'response.write("<td class='td-r'>" & rsUpdateEntry4("calculo") &"</td></tr>")
        rsUpdateEntry6.close

        rsUpdateEntry.close
        rsUpdateEntry2.close
        rsUpdateEntry3.close
        rsUpdateEntry4.close

        rsUpdateEntry7.MoveNext        
    wend
    rsUpdateEntry7.close

    separador
    response.write("<tr><td class='subtitulo td-r'>TOTAL</td>")

    for i=1 to 4 step 1
        response.write("<td class='td-r'>"& Total_anio_(i) &"</td><td class='td-r'>&nbsp;</td>")
    next

    response.write("<td class='td-r'>&nbsp;</td></tr>")      
    response.write("<table>")
    response.write("<div id='asesor' name='asesor'>&nbsp;</div>")







    response.write("<P>&nbsp;</P>")
    Titulo="<H3>RENDIMIENTO POR CUENTA SIGNIFICATIVA (POR TRIMESTRE) <BR>" & nTotalGraficas &" cuentas</H3>"
    DoTitulo(Titulo)
   

    CreateTable(1200)

    response.write("<tr><td class='td-c' width='35%'>")

    CreateTable(450)
    rsUpdateEntry5.MoveFirst
    while not rsUpdateEntry5.EOF   
        response.write("<tr><td class='td-l fontsmall'><a href='distribucion.asp?cardname=" & rsUpdateEntry5("cardname") &"&asesor="& rsUpdateEntry5("asesor") &"'>"& mid(rsUpdateEntry5("cardname"),1,40) &"</a></td>")
        response.write("<td class='td-l fontsmall'>"& mid(rsUpdateEntry5("asesor"),1,30) &"</td></tr>")        
        rsUpdateEntry5.MoveNext
    wend
    rsUpdateEntry5.MoveFirst
    closeTable

    response.write("</td><td class='td-c' style='vertical-align:top;'>")
    response.write("<div id='grafica'>")

    if request("cardname")<>"" then
        response.write("<H3>"& request("cardname") &"<BR>2018-2021<BR>"& request("asesor")&"<BR>")    
    else
        response.write("<H3>"& rsUpdateEntry5("cardname") &"<BR>2018-2021<BR>"& rsUpdateEntry5("asesor")&"<BR>")
    end if

    %>
     <div id="chartdiv" style="width:660px; height:300px;"></div>

    <%
    response.write("</H3></div></td></tr>")


    rsUpdateEntry5.close
    closeTable
    


    strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_Sales' AND TABLE_SCHEMA = 'dbo')     drop table dbo._sales;"
    rsUpdateEntry2.Open strSQL, adoCon4  'SBO'


%>



<script type="text/javascript"> 
      var chartData=[];
      chartData.push({date: new Date(2017, 11),  val: 0 });      
      chartData.push({date: new Date(2018, 2),  val: document.getElementById('variable1').value });
      chartData.push({date: new Date(2018, 5),  val: document.getElementById('variable2').value });
      chartData.push({date: new Date(2018, 8),  val: document.getElementById('variable3').value });
      chartData.push({date: new Date(2018, 11), val: document.getElementById('variable4').value });
      chartData.push({date: new Date(2019, 2),  val: document.getElementById('variable5').value });
      chartData.push({date: new Date(2019, 5),  val: document.getElementById('variable6').value });
      chartData.push({date: new Date(2019, 8),  val: document.getElementById('variable7').value });
      chartData.push({date: new Date(2019, 11), val: document.getElementById('variable8').value });
      chartData.push({date: new Date(2020, 2),  val: document.getElementById('variable9').value });
      chartData.push({date: new Date(2020, 5),  val: document.getElementById('variable10').value });
      chartData.push({date: new Date(2020, 8),  val: document.getElementById('variable11').value });
      chartData.push({date: new Date(2020, 11), val: document.getElementById('variable12').value });

      chartData.push({date: new Date(2021, 2),  val: document.getElementById('variable13').value });
      chartData.push({date: new Date(2021, 5),  val: document.getElementById('variable14').value });
      chartData.push({date: new Date(2021, 8),  val: document.getElementById('variable15').value });
      chartData.push({date: new Date(2021, 11),  val: document.getElementById('variable16').value });
         
     
      
                     
      AmCharts.ready(function() {
          var chart = new AmCharts.AmStockChart();
          chart.pathToImages = "/amcharts/images/";

          var dataSet = new AmCharts.DataSet();
          dataSet.dataProvider = chartData;
          dataSet.fieldMappings = [{fromField:"val", toField:"value"}];   
          dataSet.categoryField = "date";          
          chart.dataSets = [dataSet];

          var stockPanel = new AmCharts.StockPanel();
          chart.panels = [stockPanel];

          var panelsSettings = new AmCharts.PanelsSettings();
          panelsSettings.startDuration = 1;
          chart.panelsSettings = panelsSettings;   

          var graph = new AmCharts.StockGraph();
          graph.valueField = "value";
          graph.type = "column";
          graph.fillAlphas = 1;
          graph.title = "MyGraph"; 
          stockPanel.addStockGraph(graph);

          chart.write("chartdiv");
      });        
</script>     


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





