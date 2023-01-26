<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->


   
<BR><H1><B>PRORRATEO PRECIOS DE ENTREGA EN DMX</B></H1>
<H5><B>(a menos que se indique lo contrario todas las cifras son pesos)</B></H1><BR>
</center>   
<%
  Response.Expires = -1
  ValorMercancia=""
  if request("fcontrol")<>"" then
      control=request("fcontrol")
  else
      control=0
  end if



  response.write ("<form action='PrecioEntrega.asp' method='post'>")
  

    
     if control=0 then   
       response.write("<input type='hidden' name='fcontrol' id='fcontrol' value='1'>")       
       response.write("<P align='center'>Indique el numero de Pedimento a 15 digitos sin espacios: <input type='text' name='fpedimento' id='fpedimento' maxlength='15' value='"&  request("fpedimento") &"' ></P>")   
       response.write("<P align='center'>")


      %>
     
     OTROS (Almac&eacute;n USA): <input type='text' name='fotros' id='fotros' value='' style="width: 70px; background-color: #D3FFD3; font-size: 12px"> &nbsp;&nbsp;&nbsp;
     DTA: <input type='text' name='fdta' id='fdta' value='' style="width: 70px; background-color: #EAE4E4; font-size: 12px "> &nbsp;&nbsp;&nbsp;   
     PRV: <input type='text' name='fprv' id='fprv' value='240' style="width: 70px; background-color: #FBFCC8; font-size: 12px"> &nbsp;&nbsp;&nbsp;
     CNT: <input type='text' name='fcnt' id='fcnt' value='62' style="width: 70px; background-color: #FBFCC8; font-size: 12px"> &nbsp;&nbsp;&nbsp;
     FLETES DE CRUCE: <input type='text' name='ffletes' id='ffletes' value='' style="width: 70px; font-size: 12px">      <BR>  <BR>
     HONORARIOS: <input type='text' name='fhonorarios' id='fhonorarios' value='' style="width: 70px; font-size: 12px; background-color:  #FFCFCF">
     COMPL HONORARIOS: <input type='text' name='fhonorarios2' id='fhonorarios2' value='' style="width: 70px; font-size: 12px; "> &nbsp;&nbsp;&nbsp;
     Costos Puerto: <input type='text' name='fadicional1' id='fadicional1' value='' style="width: 70px; font-size: 12px; "> &nbsp;&nbsp;&nbsp;
     Costos P. Contenedores : <input type='text' name='fadicional2' id='fadicional2' value='' style="width: 70px; font-size: 12px; "> &nbsp;&nbsp;&nbsp;
     Maniobras : <input type='text' name='fadicional3' id='fadicional3' value='' style="width: 70px; font-size: 12px; "> &nbsp;&nbsp;&nbsp;
     Revision Mercancia : <input type='text' name='fadicional4' id='fadicional4' value='' style="width: 70px; font-size: 12px; "> &nbsp;&nbsp;&nbsp;
  
  
     <%      
         response.write("<input type='submit' value='Mostras Partidas'></form></P>")
     end if






   if len(request("fpedimento"))=15 and control=1 then     
         response.write("<input type='hidden' name='fcontrol' id='fcontrol' value='2' > ")
       
         FormHeader
         response.write("</P>")

        strSQL="if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[import1]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[import1]"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="SELECT 0 as ID,T0.DocNum,substring(T0.Comments,CHARINDEX('FAC',T0.Comments)+3,300) as 'factura',T0.DocRate,T1.LineNum, T1.ItemCode,T0.CardName, B.Custom as 'IGI%',round(T1.Quantity*T1.Price*T0.DocRate*B.Custom/100,2) as 'IGI',T1.Dscription,A.SuppCatNum,T1.Quantity,T1.Price,"
        strSQL=strSQL &"0 as subtotal into import1 "
        strSQL=strSQL &"FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry  inner join OITM A on T1.itemcode=A.itemcode "
        strSQL=strSQL &"inner join OARG B on A.CstGrpCode=B.CstGrpCode "
        strSQL=strSQL &"left join OITL T2 on t1.docentry = T2.[ApplyEntry] and T1.[LineNum] = T2.[ApplyLine] and T2.[ApplyType] = 20 left JOIN ITL1 T3 ON T2.LogEntry = T3.LogEntry "
        strSQL=strSQL &"left join OBTN T4 on T4.[ItemCode] = T3.[ItemCode] and T3.[MdAbsEntry] = t4.[absentry] WHERE T4.[DistNumber] = "& request("fpedimento") &" and T0.CANCELED='N' order by T0.DocNum,T1.LineNum"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="SELECT * from import1 order by DocNum,LineNum"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX
        i=1

        'enumerar partidas'
        while not rsUpdateEntry.EOF
          strSQL="UPDATE import1 SET ID="&i &" where DocNum="& rsUpdateEntry("DocNum") & " and LineNum=" & rsUpdateEntry("LineNum")          
          rsUpdateEntry2.Open strSQL, adoCon2  'DMX
          rsUpdateEntry.MoveNext           
          i=i+1
        wend
        rsUpdateEntry.close
    
        strSQL="alter table [import1] alter column [IGI] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import1] alter column [Price] decimal(18,4)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import1] alter column [subtotal] decimal(18,4)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX


         strSQL="select * from import1 order by ID"      
         rsUpdateEntry.Open strSQL, adoCon2  'DMX

   
  
     %>   


     <P>&nbsp</P><center>
     <table width='1100px' cellspacing=1 cellpadding=3 border=0 aling="center"><tr>
     <td class=subtitulo class=td-c><B>ID</td>     
     <td class=subtitulo class=td-c><B>Entrada</td>
      <td class=subtitulo class=td-c><B>Factura</td>
     <td class=subtitulo class=td-c><B>Linea</td>
     <td class=subtitulo class=td-c><B>Codigo</td>
     <td class=subtitulo class=td-c><B>Descripcion</td>
     <td class=subtitulo class=td-c><B>Proveedor</td>  
     <td class=subtitulo class=td-c><B># Item</td>   
     <td class=subtitulo class=td-c><B>Cantidad</td>
     <td class=subtitulo class=td-c><B>Precio (USD)</td>    
     <td class=subtitulo class=td-c><B>IGI (MXN)</td>   
     <td class=subtitulo class=td-c><B>checkbox</td>  
     </tr>
    

    <%
    vEntrada=""
      while not rsUpdateEntry.EOF
          response.write("<tr>")
             vEntrada=rsUpdateEntry("DocNum") 
             response.write("<td class=td-r><B>" & rsUpdateEntry("ID")  &"</B></td>" )   
             response.write("<td class=td-c><B>" & rsUpdateEntry("DocNum")  &"</B></td>" )  
             response.write("<td class=td-r><B>" & rsUpdateEntry("Factura")  &"</B></td>" ) 
             response.write("<td class=td-c>" & int(trim(rsUpdateEntry("LineNum")))+1  &"</td>" )  
             response.write("<td class=td-r>" & rsUpdateEntry("ItemCode")  &"</td>" )
             response.write("<td class=td-r>" & mid(rsUpdateEntry("Dscription"),1,35)  &"</td>" )
             response.write("<td class=td-r>" & mid(rsUpdateEntry("CardName"),1,35)  &"</td>" )
             response.write("<td class=td-r>" & rsUpdateEntry("SuppCatNum")  &"</td>" )
             response.write("<td class=td-r>" & rsUpdateEntry("Quantity")  &"</td>" )

             response.write("<td class=td-r><input style='width: 70px;' type='text' id='price"& rsUpdateEntry("ID") &"' name='price"& rsUpdateEntry("ID") &"' value='"& rsUpdateEntry("Price")  &"'></td>" )
             response.write("<td class=td-r><input style='width: 70px;' type='text' id='IGI"& rsUpdateEntry("ID") &"' name='IGI"& rsUpdateEntry("ID") &"' value='0.00'></td>" )
             response.write("<td class=td-c><input type='checkbox' value=0 ></td>" )

      
    
             rsUpdateEntry.MoveNext 
             if not rsUpdateEntry.EOF then
                 if  vEntrada<>rsUpdateEntry("DocNum")  then
                     response.write("</tr><tr><td colspan=12 style='font-size: 3px; background-color: #86BBDB'>&nbsp;</td></tr>")
                 end if
             end if
                    
      wend
      rsUpdateEntry.close      

         response.write("</tr><tr><td colspan=12><HR></td></tr><tr><td colspan=10 class=td-c>")

         response.write("<input type='submit' value='cotejar precios a factura y montos de IGI, continuar'></form></table>")
     

end if






if len(request("fpedimento"))=15 and control=2 then
        response.write("<input type='hidden' name='fcontrol' id='fcontrol' value='3' > ")
        
        strSQL="SELECT * from import1 order by ID"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        while not rsUpdateEntry.EOF
            vstring1="price" & trim(rsUpdateEntry("ID"))
            vstring2="IGI"  & trim(rsUpdateEntry("ID"))
            strSQL="UPDATE import1 SET [Price]=" & request(vstring1) &",[IGI]=" & request(vstring2)   &" where ID=" & trim(rsUpdateEntry("ID"))
            'response.write(strSQL&"<BR>")
            rsUpdateEntry2.Open strSQL, adoCon2  'DMX
          rsUpdateEntry.MoveNext                     
        wend
        rsUpdateEntry.close


        strSQL="UPDATE import1 SET [subtotal]=round(DocRate*Price*Quantity,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="select sum(subtotal) as calculo from import1"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        FormHeader
        response.write("&nbsp;&nbsp;&nbsp;VALOR COMERCIAL ADUANA: <font color=red><B>" & FormatCurrency(rsUpdateEntry("calculo"))       & "</B></font></P>")
        monto=""
        monto=trim(rsUpdateEntry("calculo"))
        if monto="" then monto="1" end if

        rsUpdateEntry.close

        strSQL="if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[import2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[import2]"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX


        strSQL="select DocNum,CardName,Factura,SUM(subtotal) as subtotal into import2 from import1  group by DocNum,CardName,Factura"        
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [tasa] float(53)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="UPDATE [import2] set [tasa]=subtotal/"& monto
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon2  'DMX


        strSQL="alter table [import2] add [Otros] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [DTA] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [PRVCNT] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Fletes] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Honorarios] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Honorarios2] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Adicional1] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Adicional2] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Adicional3] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [Adicional4] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [IGI] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        strSQL="alter table [import2] add [shipper] decimal(18,2)"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

         strSQL="update [import2] set [shipper]=0.00"
         rsUpdateEntry.Open strSQL, adoCon2  'DMX
             

         votros=""
         vdta=""
         vprv=""
         vcnt=""
         vfletes=""        
         vhonorarios2=""
         vadicional1=""
         vadicional2=""
         vadicional3=""
         vadicional4=""
     
     
         if request("fotros") ="" then             
                 votros="0"
          else   
                 votros=trim(request("fotros"))
         end if
         if request("fdta") ="" then             
                 vdta="0"
         else   
                 vdta=trim(request("fdta"))
         end if

          if request("fprv") ="" then             
                 vprv="0"
          else   
                 vprv=trim(request("fprv"))
         end if        
       
         if request("fcnt") ="" then             
                 vcnt="0"
          else   
                 vcnt=trim(request("fcnt"))
         end if

         vPRVCNT=0
         vPRVCNT=int(vprv)+int(vcnt)
        
         if request("ffletes") ="" then             
                 vfletes="0"
          else   
                 vfletes=trim(request("ffletes"))
         end if
        
          if request("fhonorarios") ="" then             
                 vhonorarios="0"
          else   
                 vhonorarios=trim(request("fhonorarios"))
         end if
          if request("fhonorarios2") ="" then             
                 vhonorarios2="0"
          else   
                 vhonorarios2=trim(request("fhonorarios2"))
         end if
         if request("fadicional1") ="" then             
                 vadicional1="0"
          else   
                 vadicional1=trim(request("fadicional1"))
         end if
          if request("fadicional2") ="" then             
                 vadicional2="0"
          else   
                 vadicional2=trim(request("fadicional2"))
         end if
          if request("fadicional3") ="" then             
                 vadicional3="0"
          else   
                 vadicional3=trim(request("fadicional3"))
         end if
          if request("fadicional4") ="" then             
                 vadicional4="0"
          else   
                 vadicional4=trim(request("fadicional4"))
         end if


        strSQL="UPDATE [import2] SET Otros=round("& votros &"*tasa,2),DTA=round("& vDTA &"*tasa,2),PRVCNT=round("& vPRVCNT &"*tasa,2),Fletes=round("& vfletes &"*tasa,2),Honorarios=round("& vhonorarios &"*tasa,2),Honorarios2=round("& vhonorarios2 &"*tasa,2),Adicional1=round("& vadicional1 &"*tasa,2),Adicional2=round("& vadicional2 &"*tasa,2),Adicional3=round("& vadicional3 &"*tasa,2),Adicional4=round("& vadicional4 &"*tasa,2)"
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        vDocnum=0
        strSQL="select DocNum from [import2] where subtotal=(select MAX(subtotal) from [import2])"
        'response.write  strSQL
        rsUpdateEntry.Open strSQL, adoCon2  'DMX
        vDocnum=trim(rsUpdateEntry("DocNum"))
        rsUpdateEntry.close

        strSQL="UPDATE [import2] set OTROS=Otros + ( "& votros &"-(select sum(Otros) from [import2] )),"
        strSQL= strSQL & "DTA=DTA + ( "& vDTA &"-(select sum(DTA) from [import2] )),"
        strSQL= strSQL & "PRVCNT=PRVCNT + ( "& vPRVCNT &"-(select sum(PRVCNT) from [import2] )),"
        strSQL= strSQL & "FLETES=FLETES + ( "& vfletes &"-(select sum(FLETES) from [import2] )),"
        strSQL= strSQL & "HONORARIOS=HONORARIOS + ( "& vhonorarios &"-(select sum(HONORARIOS) from [import2] )),"
        strSQL= strSQL & "HONORARIOS2=HONORARIOS2 + ( "& vhonorarios2 &"-(select sum(HONORARIOS2) from [import2] )),"
        strSQL= strSQL & "ADICIONAL1=ADICIONAL1 + ( "& vadicional1 &"-(select sum(ADICIONAL1) from [import2] )),"
        strSQL= strSQL & "ADICIONAL2=ADICIONAL2 + ( "& vadicional2 &"-(select sum(ADICIONAL2) from [import2] )), "
        strSQL= strSQL & "ADICIONAL3=ADICIONAL3 + ( "& vadicional3 &"-(select sum(ADICIONAL3) from [import2] )), "
        strSQL= strSQL & "ADICIONAL4=ADICIONAL4 + ( "& vadicional4 &"-(select sum(ADICIONAL4) from [import2] )) "
        strSQL= strSQL & "where DocNum="& vDocnum
        rsUpdateEntry.Open strSQL, adoCon2  'DMX


         strSQL="UPDATE a set a.IGI=aux.IGI from [import2] a left join ( select DocNum,SUM(IGI) as IGI from [import1] group by DocNum ) as aux on a.docnum=aux.DocNum"
         rsUpdateEntry.Open strSQL, adoCon2  'DMX

        
        ShowTable



      response.write("<tr><td colspan=15 class=td-c><input type='submit' value='ingrese fletes de facturas, continuar'></tr></form></table>")
end if






if len(request("fpedimento"))=15 and control=3 then
        strSQL="select sum(subtotal) as calculo from import1"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        FormHeader
        response.write("&nbsp;&nbsp;&nbsp;VALOR COMERCIAL ADUANA: <font color=red><B>" & FormatCurrency(rsUpdateEntry("calculo")) & "</B></font></P>")
        rsUpdateEntry.close
     
        strSQL="SELECT * from import2 order by DocNum"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX

        while not rsUpdateEntry.EOF
            vstring1="flete" & trim(rsUpdateEntry("DocNum"))     
            if  request(vstring1)<>"" then
                    vflete=request(vstring1)
            else
                    vflete="0.00"
            end if 

            strSQL="UPDATE import2 SET [shipper]=" & vflete &"  where DocNum=" & trim(rsUpdateEntry("DocNum"))
            'response.write(strSQL&"<BR>")
            rsUpdateEntry2.Open strSQL, adoCon2  'DMX
          rsUpdateEntry.MoveNext                     
        wend
        rsUpdateEntry.close

        strSQL="UPDATE a SET a.shipper=round(a.shipper*(select top 1 DocRate from import1 where DocNum=a.Docnum),2)  from import2 a"
        rsUpdateEntry.Open strSQL, adoCon2  'DMX


        ShowTable        
        ShowTable2

end if







sub FormHeader

         %>

     <input type='hidden' name='fpedimento' id='fpedimento' value='<%=request("fpedimento")%>' > 

     <P style="text-align: center"> <font style="font-size: 20px; font-weight: 600px; color: red; "><%=request("fpedimento")%> </font><BR>

     <font style="background-color: white">OTROS (Almac&eacute;n USA): <font style="background-color: #D3FFD3; font-size: 14px" ><B> <%=request("fotros")%></B>  &nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='fotros' id='fotros' value='<%=request("fotros")%>' > 
     <font style="background-color: white">DTA: <font style="background-color: #EAE4E4; font-size: 14px "><B>  <%=request("fdta")%> </B>&nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='fdta' id='fdta' value='<%=request("fdta")%>' >   
     <font style="background-color: white">PRV:  <font style="background-color: #FBFCC8; font-size: 14px"> <B> <%=request("fprv")%></B> &nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='fprv' id='fprv' value='<%=request("fprv")%>' > 
     <font style="background-color: white">CNT:  <font style="background-color: #FBFCC8; font-size: 14px ""> <B> <%=request("fcnt")%> </B>&nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='fcnt' id='fcnt' value='<%=request("fcnt")%>' ">
     <font style="background-color: white">FLETES DE CRUCE:  <font style="font-size: 14px"> <B> <%=request("ffletes")%> </B>&nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='ffletes' id='ffletes' value='<%=request("ffletes")%>' >     <BR>  <BR>
     <font style="background-color: white">HONORARIOS: <font style="4font-size: 14px; background-color:  #FFCFCF"> <B> <%=request("fhonorarios")%> </B> &nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='fhonorarios' id='fhonorarios' value='<%=request("fhonorarios")%>' > 
     <font style="background-color: white">COMPL HONORARIOS:  <font  style="font-size: 14px; "><B> <%=request("fhonorarios2")%></B> &nbsp;&nbsp;&nbsp;&nbsp; </font> 
     <input type='hidden' name='fhonorarios2' id='fhonorarios2' value='<%=request("fhonorarios2")%>' > 

     <font style="background-color: white">
     Costos Puerto:  <B> <%=request("fadicional1")%></B> &nbsp;&nbsp;&nbsp;&nbsp;
     Costos P. Contenedores: <B> <%=request("fadicional2")%></B> &nbsp;&nbsp;&nbsp;&nbsp; 
     Maniobras: <B> <%=request("fadicional3")%></B> &nbsp;&nbsp;&nbsp;&nbsp; 
     Revision Mecanica: <B> <%=request("fadicional4")%></B> &nbsp;&nbsp;&nbsp;&nbsp; 
     

     <input type='hidden' name='fadicional1' id='fadicional1' value='<%=request("fadicional1")%>' >   
     <input type='hidden' name='fadicional2' id='fadicional2' value='<%=request("fadicional2")%>' > 
     <input type='hidden' name='fadicional3' id='fadicional3' value='<%=request("fadicional3")%>' > 
     <input type='hidden' name='fadicional4' id='fadicional4' value='<%=request("fadicional4")%>' > 

<%
end sub








sub ShowTable
    
        if control=2 then
         strSQL="select *,(Otros+DTA+PRVCNT+Fletes+Honorarios+Honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+IGI) as costos from [import2] order by DocNum"
        end if

        if control=3 then
         strSQL="select *,(Otros+DTA+PRVCNT+Fletes+Honorarios+Honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+IGI+shipper) as costos from [import2] order by CardName,IGI,shipper"
        end if

         rsUpdateEntry.Open strSQL, adoCon2  'DMX

 %>
     <P>&nbsp</P><center>

     <table width='1360px' cellspacing=1 cellpadding=3 border=0 aling="center"><tr>
     <td class=subtitulo class=td-c><B>Entrada</td>     
     <td class=subtitulo class=td-c><B>Proveedor</td>
     <td class=subtitulo class=td-c><B># Factura</td>   
     <td class=subtitulo class=td-c><B>Valor <BR>Aduana</td>   
     <td class=subtitulo class=td-c><B>Tasa <BR> Prorrateo</td>
      
     <td class=subtitulo class=td-c><B>Otros <BR> Gasto Almcn</td>
     <td class=subtitulo class=td-c><B>DTA</td>    
     <td class=subtitulo class=td-c><B>PRV-CNT</td>
     <td class=subtitulo class=td-c><B>Fletes <BR> Cruce</B></td>     
     <td class=subtitulo class=td-c><B>Hono- <BR> rarios </B></td>
     <td class=subtitulo class=td-c><B>Complemn <BR>Honorarios </B></td>
     <td class=subtitulo class=td-c><B>Costos <BR>Puerto </B></td>
     <td class=subtitulo class=td-c><B>Costos P.<BR>Contenedores  </B></td>
     <td class=subtitulo class=td-c><B>Mani-<BR>obras </B></td>
     <td class=subtitulo class=td-c><B>Costos P.<BR>Contenedores  </B></td>
     <td class=subtitulo class=td-c><B>Revision<br>mecanica  </B></td>
     <%
     if control=2 then
           response.write("<td class=subtitulo class=td-c><B>Flete <BR> Provdor (USD)</B></td>")
     end if
     if control=3 then
           response.write("<td class=subtitulo class=td-c><B>Flete <BR> Provdor </B></td>")
     end if
     %>
     <td class=subtitulo class=td-c><B>Costos <BR> Totales  </B></td>
     </tr>
    

    <%
      while not rsUpdateEntry.EOF
          response.write("<tr>")
           response.write("<td class=td-r><B>" & rsUpdateEntry("DocNum")  &"</B></td>" )   
             response.write("<td class=td-r>" & mid(rsUpdateEntry("CardName"),1,30)  &"</td>" )
               response.write("<td class=td-r>" & rsUpdateEntry("factura")  &"</td>" )
               response.write("<td class=td-r style='padding-right: 8px'>" & FormatCurrency(rsUpdateEntry("subtotal"))  &"</td>" )
                response.write("<td class=td-r style='padding-left: 8px'>" & rsUpdateEntry("tasa")  &"</td>" )               
                  
                   response.write("<td class=td-r style='background-color: #D3FFD3'>" & FormatCurrency(rsUpdateEntry("Otros")  )&"</td>" )
                    response.write("<td class=td-r style='background-color: #EAE4E4'>" & FormatCurrency(rsUpdateEntry("DTA")  )&"</td>" )                     
                      response.write("<td class=td-r style=' background-color: #FBFCC8 '>" & FormatCurrency(rsUpdateEntry("PRVCNT")  )&"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("Fletes") )  &"</td>" )                      
                      response.write("<td class=td-r style=' background-color:  #FFCFCF ' >" & FormatCurrency(rsUpdateEntry("Honorarios") )  &"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("Honorarios2") )  &"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("Adicional1") )  &"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("Adicional2") )  &"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("Adicional3") )  &"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("Adicional4") )  &"</td>" )
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("IGI") )  &"</td>" )
                      response.write("<td class=td-r>" )
                      if control=2  then
                          response.write("<input style='width: 70px;' type='text' id='flete"& rsUpdateEntry("DocNum") &"' name='flete"& rsUpdateEntry("DocNum") &"' value='0.00'> ")
                      end if
                      if control=3  then
                          response.write( FormatCurrency(rsUpdateEntry("shipper") )   )
                      end if
                      response.write("</td>")
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("costos") )  &"</td>" )

          rsUpdateEntry.MoveNext 
                    'response.write("</tr><tr><td colspan=12><HR></td></tr>")
      wend
      rsUpdateEntry.close      
      

      strSQL="select count(distinct(DocNum)) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX      
      response.write("</tr><tr><td colspan=16><HR></tr><tr><td colspan=3 class=td-c><B># Total de Entradas: " & trim(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(subtotal) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo") ) &"</td>")
      rsUpdateEntry.close


      strSQL="select round(sum(tasa),6) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & trim(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

     
      strSQL="select sum(Otros) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(DTA) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close
     

      strSQL="select sum(PRVCNT) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Fletes) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close
    
      strSQL="select sum(Honorarios) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Honorarios2) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional1) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional2) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional3) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional4) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(IGI) as calculo from [import2] "
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      if control=2 then
           response.write("<td>&nbsp;</td>")
      end if
 
      if control=3 then
              strSQL="select sum(shipper) as calculo from [import2] "
              rsUpdateEntry.Open strSQL, adoCon2  'DMX
              response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
              rsUpdateEntry.close
      end if


      strSQL="select (SUM(Otros)+SUM(DTA)+SUM(PRVCNT)+SUM(Fletes)+SUM(Honorarios)+SUM(Honorarios2)+SUM(Adicional1)+SUM(Adicional2)+ SUM(Adicional3)+ SUM(Adicional4) + SUM(IGI) + SUM(shipper)) as calculo  from   [import2] "
      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      response.write("<td class=td-c style='color: red'><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td></tr>")
      rsUpdateEntry.close


      response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")


end sub    
  

  






   


response.write("</table><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>")






sub ShowTable2


   response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")
   response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")
   response.write("<tr><td colspan=15 class=td-c style='font-size: 20px; font-weight: 600px; color: red; '> PRECIOS DE ENTREGA SUGERIDOS PARA CREAR </td></tr> ")
   response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")

            
  
      strSQL="select a.cardname,a.IGI as IGI2,a.shipper as shipper2,SUM(otros) as Otros,SUM(DTA) as DTA,SUM(PRVCNT) as PRVCNT,SUM(fletes) as fletes,SUM(honorarios) as Honorarios,"
      strSQL=strSQL &"SUM(honorarios2) as Honorarios2,SUM(Adicional1) as Adicional1,SUM(Adicional2) as Adicional2, SUM(Adicional3) as Adicional3, SUM(Adicional4) as Adicional4, sum(IGI) as IGI,SUM(shipper) as shipper from import2 a group by cardname,IGI,shipper order by CardName,IGI,shipper"
      'response.write  strSQL
      rsUpdateEntry.Open strSQL, adoCon2  'DMX

 %>
     <P>&nbsp</P><center>

     <table width='1360px' cellspacing=1 cellpadding=3 border=0 aling="center"><tr>
     <td class=subtitulo class=td-c><B>Concepto</td>    

 <%
     i=1
     while not rsUpdateEntry.EOF
        response.write("<td class=subtitulo class=td-c><B>" & i &"</td> ")
        i=i+1
        rsUpdateEntry.MoveNext 
     wend
     rsUpdateEntry.MoveFirst
     contador=i
 %>
         
     </tr>
    <%
    Dim Entradas(50)    

    for i=1 to 15
        response.write("<tr>")
        select case i
            case 1  response.write("<td class=td-r>Entradas</td>")
            case 2  response.write("<td class=td-r>DTA</td>")
            case 3  response.write("<td class=td-r>PRVCNT</td>")
            case 4  response.write("<td class=td-r>Flete Proveedor</td>")
            case 5  response.write("<td class=td-r>Flete Cruce</td>")
            case 6  response.write("<td class=td-r>Almac&eacute;n</td>")
            case 7  response.write("<td class=td-r>Honorarios</td>")
            case 8  response.write("<td class=td-r>Complemn Honor</td>")
            case 9  response.write("<td class=td-r>Costos Puerto</td>")
            case 10  response.write("<td class=td-r>Costos P. Conted</td>")
            case 11  response.write("<td class=td-r>Maniobras</td>")
            case 12  response.write("<td class=td-r>Revision Mecanica</td>")
            case 13  response.write("<td class=td-r>Subtotal s/IGI</td>")
            case 14  response.write("<td class=td-r>IGI</td>")
            case 15  response.write("<td class=td-r>Costo total</td>")
            end select

              x=1
              while not rsUpdateEntry.EOF and i=1   
                        response.write("<td class=td-c style='background-color: #EAE8E8'><B>" )
                                  strSQL="select DocNum  from import2 where CardName='"& rsUpdateEntry("CardName") &"' and IGI= " & rsUpdateEntry("IGI2")  &" and shipper= " & rsUpdateEntry("shipper2") 
                                  rsUpdateEntry2.Open strSQL, adoCon2  'DMX
                                  
                                  vstring=""
                                  while not rsUpdateEntry2.EOF
                                      vstring= vstring & trim(rsUpdateEntry2("Docnum") )
                                      rsUpdateEntry2.MoveNext 
                                      if not rsUpdateEntry2.EOF then
                                              vstring= vstring & "," 
                                      end if
                                  wend
                                  Entradas(x)=trim(vstring)
                                  response.write( vstring & "</B></td>"   )
                                  rsUpdateEntry2.close                                          
                        
                        rsUpdateEntry.MoveNext  
                        x=x+1                                      
              wend

              while not rsUpdateEntry.EOF and i=2     
                        response.write("<td class=td-r>"& rsUpdateEntry("DTA") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
              wend
 
             while not rsUpdateEntry.EOF and i=3     
                        response.write("<td class=td-r>"& rsUpdateEntry("PRVCNT") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=4     
                        response.write("<td class=td-r>"& rsUpdateEntry("shipper") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=5     
                        response.write("<td class=td-r>"& rsUpdateEntry("fletes") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=6     
                        response.write("<td class=td-r>"& rsUpdateEntry("otros") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=7     
                        response.write("<td class=td-r>"& rsUpdateEntry("Honorarios") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=8     
                        response.write("<td class=td-r>"& rsUpdateEntry("Honorarios2") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

              while not rsUpdateEntry.EOF and i=9     
                        response.write("<td class=td-r>"& rsUpdateEntry("Adicional1") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=10     
                        response.write("<td class=td-r>"& rsUpdateEntry("Adicional2") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=11     
                        response.write("<td class=td-r>"& rsUpdateEntry("Adicional3") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend
             while not rsUpdateEntry.EOF and i=12     
                        response.write("<td class=td-r>"& rsUpdateEntry("Adicional4") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             if i=13 then
                for n=1 to contador-1
                      strSQL="SELECT isnull(sum(Otros+DTA+PRVCNT+fletes+Honorarios+honorarios2+Adicional1+Adicional2+shipper),0) as calculo from import2 where Docnum in ("& Entradas(n) &")"                 
                      'response.write("<td class=td-r>"& strSQL &" MXP </td>" )   
                      rsUpdateEntry2.Open strSQL, adoCon2  'DMX       
                      response.write("<td class=td-r style='background-color: #C0C0C0'><B>"& trim(rsUpdateEntry2("calculo")) &"</B></td>" )      
                      'response.write("<td class=td-r>"& strSQL &"</td>" )
                      rsUpdateEntry2.close   
                next
             end if

             while not rsUpdateEntry.EOF and i=14   
                        if  trim(rsUpdateEntry("IGI"))<>"0" then
                             response.write("<td class=td-r style='color: white;background-color: red'>"& rsUpdateEntry("IGI") &" MXP </td>" )                   
                        else
                             response.write("<td class=td-r>"& rsUpdateEntry("IGI") &" MXP </td>" )   
                        end if 
                        rsUpdateEntry.MoveNext                     
             wend

            if i=15 then
                for n=1 to contador-1
                      strSQL="SELECT isnull(sum(Otros+DTA+PRVCNT+fletes+Honorarios+honorarios2+Adicional1+Adicional2+shipper+IGI),0) as calculo from import2 where Docnum in ("& Entradas(n) &")"                 
                      'response.write("<td class=td-r>"& strSQL &" MXP </td>" )   
                      rsUpdateEntry2.Open strSQL, adoCon2  'DMX       
                      response.write("<td class=td-r style='background-color: #86BBDB'><B>"& trim(rsUpdateEntry2("calculo")) &"</B></td>" )      
                      'response.write("<td class=td-r>"& strSQL &"</td>" )
                      rsUpdateEntry2.close   
                next
             end if



             rsUpdateEntry.MoveFirst
           response.write("</tr>")
    next


    rsUpdateEntry.close      
      


      response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")
      


end sub    
  

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

