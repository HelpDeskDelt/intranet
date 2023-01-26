<!--#include virtual="/config/conn.asp"-->


<%
   Dim trimestres(4)
    trimestres(1)="(1,2,3)"
    trimestres(2)="(4,5,6)"
    trimestres(3)="(7,8,9)"
    trimestres(4)="(10,11,12)"

   response.write("<H3>"& request("cardname") &"<BR>2018-2019<BR>"& request("asesor")&"</h3><BR>")

   i=1
    for anio=2018 to 2020
            for P=1 to 4
                strSQL="select sum(Subtotal) as calculo from HistoricoVentas"&anio&" where cardname='"& request("cardname")&"' and month(DocDate) in "& trimestres(P)
                'response.write(strSQL &"<BR>")
                rsUpdateEntry6.Open strSQL, adoCon4

                response.write("<input type='hidden' id='graph"&i&"' value='"& trim(rsUpdateEntry6("calculo")) &"'>")  
                rsUpdateEntry6.close               
                i=i+1
            next
    next

%>

<script type="text/javascript"> 
      var chartData2=[];
      chartData2.push({date: new Date(2017, 11),  val: 0 });      
      chartData2.push({date: new Date(2018, 2),  val: document.getElementById('graph1').value });
      chartData2.push({date: new Date(2018, 5),  val: document.getElementById('graph2').value });
      chartData2.push({date: new Date(2018, 8),  val: document.getElementById('graph3').value });
      chartData2.push({date: new Date(2018, 11), val: document.getElementById('graph4').value });
      chartData2.push({date: new Date(2019, 2),  val: document.getElementById('graph5').value });
      chartData2.push({date: new Date(2019, 5),  val: document.getElementById('graph6').value });
      chartData2.push({date: new Date(2019, 8),  val: document.getElementById('graph7').value });
      chartData2.push({date: new Date(2019, 11), val: document.getElementById('graph8').value });
      chartData2.push({date: new Date(2020, 2),  val: document.getElementById('graph9').value });
      chartData2.push({date: new Date(2020, 5),  val: document.getElementById('graph10').value });
      chartData2.push({date: new Date(2020, 8),  val: document.getElementById('graph11').value });
      chartData2.push({date: new Date(2020, 11), val: document.getElementById('graph12').value });
         
     
      
                     
      AmCharts2.ready(function() {
          var chart = new AmCharts.AmStockChart();
          chart.pathToImages = "/amcharts/images/";

          var dataSet = new AmCharts.DataSet();
          dataSet.dataProvider = chartData2;
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

          chart.write("chartdiv2");
      });        
</script>     



 <div id="chartdiv2" style="width:660px; height:300px;"></div>