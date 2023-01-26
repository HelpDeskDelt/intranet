<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->
<!--#include virtual="/modules/contenidos.asp"-->

<%

   titulo="PRORRATEO EN PRECIOS DE ENTREGA<BR><font style='font-size: 11px'>(a menos que se indique lo contrario todas las cifras son pesos)</font>"
   Dotitulo(titulo)
   

   if request("msg")<>"" then  
           response.write("<P class='alert td-c'>" & request("msg") &"</P>")
   end if


  Response.Expires = -1

  ValorMercancia=""
  if request("fcontrol")<>"" then
      control=request("fcontrol")
  else
      strSQL="if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[import1]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[import1]"
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      rsUpdateEntry.Open strSQL, adoCon3  'DFW
      rsUpdateEntry.Open strSQL, adoCon5  'MEIDE

      strSQL="if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[import2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[import2]"
      rsUpdateEntry.Open strSQL, adoCon2  'DMX
      rsUpdateEntry.Open strSQL, adoCon3  'DFW
      rsUpdateEntry.Open strSQL, adoCon5  'MEIDE

      control=0
  end if

  response.write ("<form action='PrecioEntrega.asp' method='post'>")
  

    
     if control=0 then   
       response.write("<input type='hidden' name='fcontrol' id='fcontrol' value='1'>") 
       CreateTable(600)     
       rowin
         response.write("<td class='td-r subtitulo'>Indique la razon social: </td><td class='td-l'><select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option><option value='MME'>MME</option></select></td>")
       RowOut
       rowin
       response.write("<td class='td-r subtitulo'>Indique el numero de <B>Pedimento</B> a 15 digitos sin espacios: </td>")
       response.write("<td class='td-l'><input type='text' class=anchoSmals name='fpedimento' id='fpedimento' maxlength='15' value='"&  request("fpedimento") &"'></td>")
       rowout
       rowin
       response.write("<td class='td-r subtitulo'> Indique las <B>entradas</B> que desea ignorar para el prorrateo:</td>")
       response.write("<td class='td-l'><input class=anchoSmall type='text' name='fentradas' id='fentradas' value='"&  request("fpedimento") &"'></td>")
       rowout
       closetable

       
       response.write("<BR> </P><P align='center'>&nbsp</P>")   

       response.write("<P align='center'>")


      %>
     <table width="900px" cellpadding="2" align="center" border=1>
     <tr>        
     <td width="25%" class='td-r'>Almacenaje: &nbsp;&nbsp;<input type='text' name='fotros' id='fotros' value='' style="width: 70px; background-color: #D3FFD3; font-size: 12px" placeholder="no comas"> </td>
     <td width="25%" class='td-r'>DTA: &nbsp;&nbsp;<input type='text' name='fdta' id='fdta' value='' style="width: 70px; background-color: #EAE4E4; font-size: 12px " placeholder="no signos"> </td>   
     <td width="25%" class='td-r'>PRV: &nbsp;&nbsp;<input type='text' name='fprv' id='fprv' value='240' style="width: 70px; background-color: #FBFCC8; font-size: 12px"> </td>
     <td width="25%" class='td-r'>CNT: &nbsp;&nbsp;<input type='text' name='fcnt' id='fcnt' value='62' style="width: 70px; background-color: #FBFCC8; font-size: 12px"> </td>
     </tr><tr>

     <td class='td-r'>Flete Nacional <B>#</B>: &nbsp;&nbsp; <input type='text' name='ffletes' id='ffletes' value='' style="width: 70px; font-size: 12px" placeholder="no comas">   </td> 
     <td class='td-r'>Honorarios: &nbsp;&nbsp;<input type='text' name='fhonorarios' id='fhonorarios' value='' style="width: 70px; font-size: 12px; background-color:  #FFCFCF">  </td>
     <td class='td-r'>Compl Honorarios: &nbsp;&nbsp;<input type='text' name='fhonorarios2' id='fhonorarios2' value='' style="width: 70px; font-size: 12px; " placeholder="no signos"> </td> 
     <td class='td-r'>Costos Puerto: &nbsp;&nbsp; <input type='text' name='fadicional1' id='fadicional1' value='' style="width: 70px; font-size: 12px; "> </td>
     </tr><tr>    
     <td class='td-r'>Cuota Compens: &nbsp;&nbsp;<input type='text' name='fadicional2' id='fadicional2' value='' style="width: 70px; font-size: 12px; "> </td>
     <td class='td-r'>Costos P. Contendors: &nbsp;&nbsp;<input type='text' name='fadicional3' id='fadicional3' value='' style="width: 70px; font-size: 12px; "> </td>
     <td class='td-r'>Maniobras: &nbsp;&nbsp;<input type='text' name='fadicional4' id='fadicional4' value='' style="width: 70px; font-size: 12px; "> </td>
     <td class='td-r'>Revision Mercancia: &nbsp;&nbsp;<input type='text' name='fadicional5' id='fadicional5' value='' style="width: 70px; font-size: 12px; "></td>
     </tr><tr> 
     <td class='td-r'>Flete extranj import <B>*</B>: &nbsp;&nbsp;<input type='text' name='fadicional6' id='fadicional6' value='' style="width: 70px; font-size: 12px; "></td> 
     <td class='td-r'>Otros: &nbsp;&nbsp;<input type='text' name='fadicional7' id='fadicional7' value='' style="width: 70px; font-size: 12px; "></td> 
     <td class='td-r'>Etiquetado: &nbsp;&nbsp;</td> 
     <td class='td-r'>Flete Maritimo: &nbsp;&nbsp;<input type='text' name='fadicional8' id='fadicional8' value='' style="width: 70px; font-size: 12px; "></td> 
     </tr>
     <tr><td class=td-c colspan=4><B># (emitido en MEX viene con IVA/retencion, capturar sin impuestos) <BR>* (emitido fuera de MEX no viene con IVA o retencion)  </B></td> </tr>
     <tr><td class=td-c colspan=4><input style="font-size:14px" type='submit' value='Continuar, mostras Partidas'></form></td>

 </table>
  
     <center>
     <P>&nbsp</P>
     <img src="/images/descrp_costos.png" border=0>   
  
     <%
     end if






   if len(request("fpedimento"))=15 and control=1 then     
         strSQL="SELECT T0.DocDate FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry inner join OITM A on T1.itemcode=A.itemcode  inner join OARG B on A.CstGrpCode=B.CstGrpCode  left join OITL T2 on t1.docentry = T2.[ApplyEntry] and T1.[LineNum] = T2.[ApplyLine] and T2.[ApplyType] = 20  left JOIN ITL1 T3 ON T2.LogEntry = T3.LogEntry left join OBTN T4 on T4.[ItemCode] = T3.[ItemCode] and T3.[MdAbsEntry] = t4.[absentry]  WHERE T4.[DistNumber] = '"&  request("fpedimento") &"' and T0.CANCELED='N'  group by T0.DocDate "

          select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

          i=0
          while not rsUpdateEntry.EOF
            rsUpdateEntry.MoveNext
            i=i+1
          wend
          rsUpdateEntry.close

          if i>1 then 
                response.redirect("PrecioEntrega.asp?msg=error: hay entradas de mercancia con diferente fecha de contabilizacion, avise a almacen")
          end if

         response.write("<input type='hidden' name='fcontrol' id='fcontrol' value='2' > ")
         response.write("<input type='hidden' name='fRS' value='"&request("fRS")&"' > ")
       
         FormHeader
         response.write("</P>")

        strSQL="if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[import1]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[import1]"

        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="SELECT 0 as ID,T0.DocNum,replace(T0.Comments,'Basado en Pedidos ','PO-') as 'factura',T0.DocRate,T1.LineNum, T1.ItemCode,T0.CardName, B.Custom as 'IGI%',round(T1.Quantity*T1.Price*T0.DocRate*B.Custom/100,2) as 'IGI',T1.Dscription,A.SuppCatNum,T1.Quantity,T1.Price,"
        strSQL=strSQL &"0 as subtotal,isnull(round(cast(T1.Quantity as float)*cast(A.U_CBM as float),6),0) as CBM,isnull(round(T1.Quantity*A.U_Peso_Kg,2),0) as PESO into import1 "
        strSQL=strSQL &"FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry  inner join OITM A on T1.itemcode=A.itemcode "
        strSQL=strSQL &"inner join OARG B on A.CstGrpCode=B.CstGrpCode "
        strSQL=strSQL &"left join OITL T2 on t1.docentry = T2.[ApplyEntry] and T1.[LineNum] = T2.[ApplyLine] and T2.[ApplyType] = 20 left JOIN ITL1 T3 ON T2.LogEntry = T3.LogEntry "
        strSQL=strSQL &"left join OBTN T4 on T4.[ItemCode] = T3.[ItemCode] and T3.[MdAbsEntry] = t4.[absentry] "
        strSQL=strSQL &"WHERE T4.[DistNumber] = '"& request("fpedimento") &"' and T0.CANCELED='N' "
        if request("fentradas")<>"" then
               strSQL=strSQL &"AND T0.DocNum not in ("& request("fentradas") &") "
        end if
        strSQL=strSQL &"order by T0.DocNum,T1.LineNum"

         'response.write strSQL
         select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="SELECT * from import1 order by DocNum,LineNum"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        i=1
        'enumerar partidas'
        while not rsUpdateEntry.EOF
          strSQL="UPDATE import1 SET ID="&i &" where DocNum="& rsUpdateEntry("DocNum") & " and LineNum=" & rsUpdateEntry("LineNum")          
          select case request("fRS")
              case "DMX" rsUpdateEntry2.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry2.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry2.Open strSQL, adoCon5  'MEIDE
          end select
          rsUpdateEntry.MoveNext           
          i=i+1
        wend
        rsUpdateEntry.close
    
        strSQL="alter table [import1] alter column [IGI] decimal(18,2)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="alter table [import1] alter column [Price] decimal(18,4)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="alter table [import1] alter column [subtotal] decimal(18,4)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="alter table [import1] add [tasa_CBM] decimal(18,4)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select      

        strSQL="alter table [import1] add [tasa_PESO] decimal(18,4)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select  

        strSQL="UPDATE [import1] set [tasa_CBM]=case when isnull((select sum(CBM) from import1),0)>0 then round(CBM/(select sum(CBM) from import1),4) else 0 end"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="UPDATE [import1] set [tasa_PESO]=case when isnull((select sum(PESO) from import1),0)>0 then round(PESO/(select sum(PESO) from import1),4) else 0 end"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="UPDATE TOP(1) [import1] set [tasa_CBM]=[tasa_CBM]+(1-(select sum(tasa_CBM) from import1))"        
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="UPDATE TOP(1) [import1] set [tasa_PESO]=[tasa_PESO]+(1-(select sum(tasa_PESO) from import1))"       
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="alter table [import1] add [tasa_valor] decimal(18,4)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="select * from import1 order by ID"      
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
    
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
             response.write("<td class=td-l><B>" & rsUpdateEntry("ID")  &"</B></td>" )   
             response.write("<td class=td-c><B>" & rsUpdateEntry("DocNum")  &"</B></td>" )  
             response.write("<td class=td-l><B>" & rsUpdateEntry("Factura")  &"</B></td>" ) 
             response.write("<td class=td-c>" & int(trim(rsUpdateEntry("LineNum")))+1  &"</td>" )  
             response.write("<td class=td-l>" & rsUpdateEntry("ItemCode")  &"</td>" )
             response.write("<td class=td-l>" & mid(rsUpdateEntry("Dscription"),1,35)  &"</td>" )
             response.write("<td class=td-l>" & mid(rsUpdateEntry("CardName"),1,35)  &"</td>" )
             response.write("<td class=td-l>" & rsUpdateEntry("SuppCatNum")  &"</td>" )
             response.write("<td class=td-l>" & rsUpdateEntry("Quantity")  &"</td>" )

             response.write("<td class=td-l><input style='width: 70px;' type='text' id='price"& rsUpdateEntry("ID") &"' name='price"& rsUpdateEntry("ID") &"' value='"& rsUpdateEntry("Price")  &"'></td>" )
             response.write("<td class=td-l><input style='width: 70px;' type='text' id='IGI"& rsUpdateEntry("ID") &"' name='IGI"& rsUpdateEntry("ID") &"' value='0.00'></td>" )
             response.write("<td class=td-c><input type='checkbox' value=0 ></td>" )

      
    
             rsUpdateEntry.MoveNext 
             if not rsUpdateEntry.EOF then
                 if  vEntrada<>rsUpdateEntry("DocNum")  then
                     response.write("</tr><tr><td colspan=12 style='font-size: 3px; background-color: #86BBDB'>&nbsp;</td></tr>")
                 end if
             end if
                    
      wend
      rsUpdateEntry.close      

         response.write("</tr><tr><td colspan=12><HR></td></tr><tr><td colspan=12 class=td-c><input type='submit' value='cotejar precios a factura y montos de IGI, continuar'></form></table>")

end if






if len(request("fpedimento"))=15 and control=2 then
        response.write("<input type='hidden' name='fcontrol' id='fcontrol' value='3' > ")
        response.write("<input type='hidden' name='fRS' value='"&request("fRS")&"' > ")
        
        strSQL="SELECT * from import1 order by ID"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        while not rsUpdateEntry.EOF
            vstring1="price" & trim(rsUpdateEntry("ID"))
            vstring2="IGI"  & trim(rsUpdateEntry("ID"))
         

            strSQL="UPDATE import1 SET [Price]=" & request(vstring1) &",[IGI]=" & request(vstring2)   &" where ID=" & trim(rsUpdateEntry("ID"))
            'response.write(strSQL&"<BR>")
            select case request("fRS")
              case "DMX" rsUpdateEntry2.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry2.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry2.Open strSQL, adoCon5  'MEIDE
          end select
          rsUpdateEntry.MoveNext                     
        wend
        rsUpdateEntry.close


        strSQL="UPDATE import1 SET [subtotal]=round(DocRate*Price*Quantity,2)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="UPDATE [import1] set [tasa_valor]=round(subtotal/(select sum(subtotal) from import1),4) "
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="UPDATE TOP(1) [import1] set [tasa_valor]=[tasa_valor]+(1-(select sum(tasa_valor) from import1))"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="select isnull(sum(subtotal),0) as calculo from import1"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        FormHeader
        response.write("<BR>VALOR COMERCIAL ADUANA: <font color=red><B>" & FormatCurrency(rsUpdateEntry("calculo"))       & "</B></font></P>")
        monto=""
        monto=trim(rsUpdateEntry("calculo"))
        if monto="" then monto="1" end if

        rsUpdateEntry.close

        strSQL="if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[import2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[import2]"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select


        strSQL="select DocNum,CardName,Factura,SUM(subtotal) as subtotal into import2 from import1  group by DocNum,CardName,Factura"        
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="alter table [import2] add [tasa] float(53)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="UPDATE [import2] set [tasa]=subtotal/"& monto
        'response.write strSQL
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select


        strSQL="alter table [import2] add [Otros] decimal(18,2),[DTA] decimal(18,2),[PRVCNT] decimal(18,2),[Fletes] decimal(18,2),[Honorarios] decimal(18,2),[Honorarios2] decimal(18,2),[Adicional1] decimal(18,2),[Adicional2] decimal(18,2),[Adicional3] decimal(18,2),[Adicional4] decimal(18,2),[Adicional5] decimal(18,2),[Adicional6] decimal(18,2),[Adicional7] decimal(18,2),[Adicional8] decimal(18,2),[Labeling] decimal(18,2),[IGI] decimal(18,2),[shipper] decimal(18,2)"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        strSQL="update [import2] set [shipper]=0.00,[Labeling]=0.00"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select


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
         vadicional5=""
         vadicional6=""
         vadicional7=""
         vadicional8=""
         vadicional9=""
     
     
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
           if request("fadicional5") ="" then             
                 vadicional5="0"
          else   
                 vadicional5=trim(request("fadicional5"))
         end if
           if request("fadicional6") ="" then             
                 vadicional6="0"
          else   
                 vadicional6=trim(request("fadicional6"))
         end if
           if request("fadicional7") ="" then             
                 vadicional7="0"
          else   
                 vadicional7=trim(request("fadicional7"))
         end if
          if request("fadicional8") ="" then             
                 vadicional8="0"
          else   
                 vadicional8=trim(request("fadicional8"))
         end if
         


        strSQL="UPDATE [import2] SET Otros=round("& votros &"*tasa,2),DTA=round("& vDTA &"*tasa,2),PRVCNT=round("& vPRVCNT &"*tasa,2),Fletes=round("& vfletes &"*tasa,2),Honorarios=round("& vhonorarios &"*tasa,2),Honorarios2=round("& vhonorarios2 &"*tasa,2),Adicional1=round("& vadicional1 &"*tasa,2),Adicional2=round("& vadicional2 &"*tasa,2),Adicional3=round("& vadicional3 &"*tasa,2),Adicional4=round("& vadicional4 &"*tasa,2),Adicional5=round("& vadicional5 &"*tasa,2),Adicional6=round("& vadicional6 &"*tasa,2),Adicional7=round("& vadicional7 &"*tasa,2),Adicional8=round("& vadicional8 &"*tasa,2)"
        'response.write strSQL
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        vDocnum=0
        strSQL="select DocNum from [import2] where subtotal=(select MAX(subtotal) from [import2])"
        'response.write  strSQL
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
        end select

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
        strSQL= strSQL & "ADICIONAL4=ADICIONAL4 + ( "& vadicional4 &"-(select sum(ADICIONAL4) from [import2] )), "
        strSQL= strSQL & "ADICIONAL5=ADICIONAL5 + ( "& vadicional5 &"-(select sum(ADICIONAL5) from [import2] )), "
        strSQL= strSQL & "ADICIONAL6=ADICIONAL6 + ( "& vadicional6 &"-(select sum(ADICIONAL6) from [import2] )), "
        strSQL= strSQL & "ADICIONAL7=ADICIONAL7 + ( "& vadicional7 &"-(select sum(ADICIONAL7) from [import2] )), "
        strSQL= strSQL & "ADICIONAL8=ADICIONAL8 + ( "& vadicional8 &"-(select sum(ADICIONAL8) from [import2] )) "
       
        strSQL= strSQL & "where DocNum="& vDocnum
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select


         strSQL="UPDATE a set a.IGI=aux.IGI from [import2] a left join ( select DocNum,SUM(IGI) as IGI from [import1] group by DocNum ) as aux on a.docnum=aux.DocNum"
         select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        
        ShowTable


      response.write("<tr><td colspan=22 class=td-c><input type='submit' value='ingrese fletes/etiquetado continuar'></tr></form></table>")
end if






if len(request("fpedimento"))=15 and control=3 then
        strSQL="select sum(subtotal) as calculo from import1"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        FormHeader
        response.write("&nbsp;&nbsp;&nbsp;VALOR COMERCIAL ADUANA: <font color=red><B>" & FormatCurrency(rsUpdateEntry("calculo")) & "</B></font></P>")
        rsUpdateEntry.close
     
        strSQL="SELECT * from import2 order by DocNum"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

        while not rsUpdateEntry.EOF
            vstring1="flete" & trim(rsUpdateEntry("DocNum"))     
            vstring2="label" & trim(rsUpdateEntry("DocNum"))   

            if  request(vstring1)<>"" then
                    vflete=request(vstring1)
            else
                    vflete="0.00"
            end if 
            if  request(vstring2)<>"" then
                    vLabel=request(vstring2)
            else
                    vLabel="0.00"
            end if

            strSQL="UPDATE import2 SET [shipper]=" & vflete &",[labeling]="& vLabel &"  where DocNum=" & trim(rsUpdateEntry("DocNum"))
            'response.write(strSQL&"<BR>")
            select case request("fRS")
              case "DMX" rsUpdateEntry2.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry2.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry2.Open strSQL, adoCon5  'MEIDE
          end select

          rsUpdateEntry.MoveNext                     
        wend
        rsUpdateEntry.close

        strSQL="UPDATE a SET a.shipper=round(a.shipper*(select top 1 DocRate from import1 where DocNum=a.Docnum),2),a.labeling=round(a.labeling*(select top 1 DocRate from import1 where DocNum=a.Docnum),2)  from import2 a"
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select


        ShowTable        
        ShowTable2

end if





sub FormHeader

         %>

     <input type='hidden' name='fpedimento' id='fpedimento' value='<%=request("fpedimento")%>' > 

     <P style="text-align: center"> <font style="font-size: 20px; font-weight: 600px; color: red; "><%=request("fpedimento")%> </font><BR>

     
     <% CreateTable(1000)  %> 
     <tr>
     <td width="12%" class='td-r subtitulo'>Almacenaje </td>
     <td width="12%" class='td-r'> <font style="background-color: #D3FFD3; font-size: 14px" ><B> <%=request("fotros")%></B></font></td>     
     <td width="12%" class='td-r subtitulo'>DTA  </td>
     <td width="12%" class='td-r'><font style="background-color: #EAE4E4; font-size: 14px "><B>  <%=request("fdta")%> </B> </font></td>  
     <td width="12%" class='td-r subtitulo'>PRV </td>
     <td width="12%" class='td-r'><font style="background-color: #FBFCC8; font-size: 14px"> <B> <%=request("fprv")%></B> </font></td> 
     <td width="12%" class='td-r subtitulo'>CNT</td>
     <td width="12%" class='td-r'><font style="background-color: #FBFCC8; font-size: 14px"> <B> <%=request("fcnt")%> </B></font></td>
        
     </tr>


     <tr>
     <td class='td-r subtitulo'>Flete Nacional</td>
     <td class='td-r'> <font style="font-size: 14px"><B> <%=request("ffletes")%> </B></font></td> 

     <td class='td-r subtitulo'>Honorarios </td>
     <td class='td-r'><font style="font-size: 14px; background-color:  #FFCFCF"> <B> <%=request("fhonorarios")%> </B>  </font> 
     </td> 
     <td class='td-r subtitulo'>Compl Honorarios </td>
     <td class='td-r'><font  style="font-size: 14px; "><B> <%=request("fhonorarios2")%></B> </font> 
     </td> 
     <td class='td-r subtitulo'> Costos Puerto</td>
     <td class='td-r'>  <B> <%=request("fadicional1")%></B>  </td>  

     </tr>  
     <tr> 

     <td class='td-r subtitulo'> Cuota Compens</td>
     <td class='td-r'> <B> <%=request("fadicional2")%></B> </td>  
     <td class='td-r subtitulo'> Costos P. Contendors</td>
     <td class='td-r'> <B> <%=request("fadicional3")%></B> </td>  
     <td class='td-r subtitulo'> Maniobras</td>
     <td class='td-r'> <B> <%=request("fadicional4")%></B> </td>    
     <td class='td-r subtitulo'> Revision Mercancia</td>
     <td class='td-r'> <B> <%=request("fadicional5")%></B>  </td> 

     </tr>  
     <tr>  

     <td class='td-r subtitulo'> Flete extranj import</td>
     <td class='td-r'> <B> <%=request("fadicional6")%></B>  </td>  
     <td class='td-r subtitulo'> Otros</td>
     <td class='td-r'> <B> <%=request("fadicional7")%></B>  </td>  
     
     <td class='td-r subtitulo'> Flete Maritimo</td>
     <td class='td-r'> <B> <%=request("fadicional8")%></B>  </td>  
     </tr>       
     </table> 

 
 <input type="hidden" name="fotros" id="fotros" value="<%=request("fotros")%>"> 
 <input type="hidden" name="fdta" id="fdta" value="<%=request("fdta")%>">
 <input type="hidden" name="fprv" id="fprv" value="<%=request("fprv")%>"> 
 <input type="hidden" name="fcnt" id="fcnt" value="<%=request("fcnt")%>">
 <input type="hidden" name="ffletes" id="ffletes" value="<%=request("ffletes")%>">
 <input type="hidden" name="fhonorarios" id="fhonorarios" value="<%=request("fhonorarios")%>">
 <input type="hidden" name="fhonorarios2" id="fhonorarios2" value="<%=request("fhonorarios2")%>">
 <input type="hidden" name="fadicional1" id="fadicional1" value="<%=request("fadicional1")%>">   
 <input type="hidden" name="fadicional2" id="fadicional2" value="<%=request("fadicional2")%>"> 
 <input type="hidden" name="fadicional3" id="fadicional3" value="<%=request("fadicional3")%>"> 
 <input type="hidden" name="fadicional4" id="fadicional4" value="<%=request("fadicional4")%>"> 
 <input type="hidden" name="fadicional5" id="fadicional5" value="<%=request("fadicional5")%>">
 <input type="hidden" name="fadicional6" id="fadicional6" value="<%=request("fadicional6")%>">
 <input type="hidden" name="fadicional7" id="fadicional7" value="<%=request("fadicional7")%>">
 <input type="hidden" name="fadicional8" id="fadicional8" value="<%=request("fadicional8")%>">
 
<%
end sub








sub ShowTable
    
        if control=2 then
         strSQL="select *,(Otros+DTA+PRVCNT+Fletes+Honorarios+Honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+Adicional5+Adicional6+Adicional7+IGI+Adicional8) as costos from [import2] order by DocNum"
        end if

        if control=3 then
         strSQL="select *,(Otros+DTA+PRVCNT+Fletes+Honorarios+Honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+Adicional5+Adicional6+Adicional7+IGI+shipper+Adicional8+Labeling) as costos from [import2] order by CardName,IGI,shipper"         
        end if
        'response.write(strSQL&"<BR>")
        select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select

 %>
     <P>&nbsp</P><center>

     <table width='1500px' cellspacing=1 cellpadding=3 border=0 aling="center"><tr>
     <td class=subtitulo class=td-c><B>Entrada</td>     
     <td class=subtitulo class=td-c><B>Proveedor</td>
     <td class=subtitulo class=td-c><B># Factura</td>   
     <td class=subtitulo class=td-c><B>Valor <BR>Aduana</td>   
     <td class=subtitulo class=td-c><B>Tasa <BR> Prorrateo</td>
      
     <td class=subtitulo class=td-c><B>Almcnaje </td>
     <td class=subtitulo class=td-c><B>DTA</td>    
     <td class=subtitulo class=td-c><B>PRVCNT</td>
     <td class=subtitulo class=td-c><B>Flete <BR> Ncional</B></td>     
     <td class=subtitulo class=td-c><B>Hono- <BR> rarios </B></td>
     <td class=subtitulo class=td-c><B>Complmn <BR>Honorar </B></td>
    

     <td class=subtitulo class=td-c><B>Costos <BR>Puerto </B></td>
     <td class=subtitulo class=td-c><B>Cuota <BR>cmpens </B></td>
     <td class=subtitulo class=td-c><B>Costos P.<BR>Contndrs  </B></td>
     <td class=subtitulo class=td-c><B>Mani-<BR>obras </B></td>     
     <td class=subtitulo class=td-c><B>Revsion<br>mrcncia  </B></td>
     <td class=subtitulo class=td-c><B>Flete Ext<BR>import  </B></td>
     <td class=subtitulo class=td-c><B>Otros  </B></td>    
     <td class=subtitulo class=td-c><B>F Maritm  </B></td>

     <td class=subtitulo class=td-c><B>IGI</B></td>

     <%
     if control=2 then
           response.write("<td class=subtitulo class=td-c><B>Flete <BR> Provdr (USD)</B></td>")
     end if
     if control=3 then
           response.write("<td class=subtitulo class=td-c><B>Flete <BR> Provdr </B></td>")
     end if
     %>
     <td class=subtitulo class=td-c><B>Etique <br>tado (USD) </B></td>
     <td class=subtitulo class=td-c><B>Costos <BR> Totales  </B></td>
     </tr>
    

    <%
      while not rsUpdateEntry.EOF
          response.write("<tr>")
          response.write("<td class=td-l><B>" & rsUpdateEntry("DocNum")  &"</B></td>" )   
          response.write("<td class=td-l>" & mid(rsUpdateEntry("CardName"),1,10)  &"</td>" )
          response.write("<td class=td-l>" & rsUpdateEntry("factura")  &"</td>" )

          response.write("<td class=td-r style='padding-right: 8px'>" & FormatCurrency(rsUpdateEntry("subtotal"))  &"</td>" )
          response.write("<td class=td-r style='padding-left: 8px'>" & rsUpdateEntry("tasa")  &"</td>" )               
                  
          response.write("<td class=td-r style='background-color: #D3FFD3'>" & rsUpdateEntry("Otros") &"</td>" )
          response.write("<td class=td-r style='background-color: #EAE4E4'>" & rsUpdateEntry("DTA")  &"</td>" )                     
          response.write("<td class=td-r style=' background-color: #FBFCC8 '>" & rsUpdateEntry("PRVCNT")  &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Fletes")   &"</td>" )                      
          response.write("<td class=td-r style=' background-color:  #FFCFCF ' >" & rsUpdateEntry("Honorarios")  &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Honorarios2")   &"</td>" )
          

          response.write("<td class=td-r >" & rsUpdateEntry("Adicional1")   &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional2")   &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional3")   &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional4")   &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional5")   &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional6")  &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional7")  &"</td>" )
          response.write("<td class=td-r >" & rsUpdateEntry("Adicional8")  &"</td>" )

                      response.write("<td class=td-r >" & rsUpdateEntry("IGI")   &"</td>" )

                      response.write("<td class=td-r>" )

                      if control=2  then
                          response.write("<input style='width: 60px;text-align:right;' type='text' id='flete"& rsUpdateEntry("DocNum") &"' name='flete"& rsUpdateEntry("DocNum") &"' value='0.00'> </td>")
                          response.write("<td class=td-l ><input style='width: 60px;text-align:right;' type='text' id='label"& rsUpdateEntry("DocNum") &"' name='label"& rsUpdateEntry("DocNum") &"' value='0.00'> </td>" )
                      end if

                      if control=3  then
                          response.write( FormatCurrency(rsUpdateEntry("shipper") ) &"</td><td class=td-r>")
                          response.write( FormatCurrency(rsUpdateEntry("labeling") ) &"</td>")
                      end if
                    
                      
                      response.write("<td class=td-r >" & FormatCurrency(rsUpdateEntry("costos") )  &"</td>" )

          rsUpdateEntry.MoveNext 
                    'response.write("</tr><tr><td colspan=12><HR></td></tr>")
      wend
      rsUpdateEntry.close      
      

      strSQL="select count(distinct(DocNum)) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select      
      response.write("</tr><tr><td colspan=22><HR></tr><tr><td colspan=3 class=td-c><B># Total de Entradas: " & trim(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close



      strSQL="select sum(subtotal) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-c><B>" & FormatCurrency(rsUpdateEntry("calculo") ) &"</td>")
      rsUpdateEntry.close


      strSQL="select round(sum(tasa),6) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-c><B>" & trim(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

     
      strSQL="select sum(Otros) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(DTA) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close
     

      strSQL="select sum(PRVCNT) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Fletes) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close
    
      strSQL="select sum(Honorarios) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Honorarios2) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional1) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional2) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional3) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional4) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional5) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional6) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional7) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      strSQL="select sum(Adicional8) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close




      strSQL="select sum(IGI) as calculo from [import2] "
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
      rsUpdateEntry.close

      if control=2 then
           response.write("<td>&nbsp;</td>")
           response.write("<td>&nbsp;</td>")
      end if
 
      if control=3 then
              strSQL="select sum(shipper) as calculo from [import2] "
              select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
              end select
              response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
              rsUpdateEntry.close

              strSQL="select sum(labeling) as calculo from [import2] "
              select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
              response.write("<td class=td-r><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
              rsUpdateEntry.close
      end if

      'response.write("<td class=td-c>&nbsp;</td>")

      strSQL="select (SUM(Otros)+SUM(DTA)+SUM(PRVCNT)+SUM(Fletes)+SUM(Honorarios)+SUM(Honorarios2)+SUM(Adicional1)+SUM(Adicional2)+ SUM(Adicional3)+ SUM(Adicional4)+ SUM(Adicional5)+ SUM(Adicional6) + SUM(Adicional7) + SUM(Adicional8)+ SUM(IGI) + SUM(shipper) + SUM(labeling)  ) as calculo  from   [import2] "
      'response.write(strSQL&"<BR>")
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
          end select
      response.write("<td class=td-r style='color: red'><B>" & FormatCurrency(rsUpdateEntry("calculo")) &"</td></tr>")
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
      strSQL=strSQL &"SUM(honorarios2) as Honorarios2,SUM(Adicional1) as Adicional1,SUM(Adicional2) as Adicional2, SUM(Adicional3) as Adicional3, SUM(Adicional4) as Adicional4, SUM(Adicional5) as Adicional5,SUM(Adicional6) as Adicional6, SUM(Adicional7) as Adicional7, SUM(Adicional8) as Adicional8, sum(IGI) as IGI,SUM(shipper) as shipper,SUM(labeling) as labeling from import2 a group by cardname,IGI,shipper order by CardName,IGI,shipper"
      'response.write(strSQL&"<BR>")
      select case request("fRS")
              case "DMX" rsUpdateEntry.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry.Open strSQL, adoCon5  'MEIDE
      end select

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

    for i=1 to 21
        response.write("<tr>")
        select case i
            case 1  response.write("<td class=td-l>Entradas</td>")
            case 2  response.write("<td class=td-l>01.DTA</td>")
            case 3  response.write("<td class=td-l>02.PRVCNT</td>")
            case 4  response.write("<td class=td-l>03.Flete Proveedor</td>")
            case 5  response.write("<td class=td-l>04.Flete Cruce</td>")
            case 6  response.write("<td class=td-l>05.Almacenaje</td>")
            case 7  response.write("<td class=td-l>06.Honorarios</td>")
            case 8  response.write("<td class=td-l>07.Complemn Honor</td>")

            case 9  response.write("<td class=td-l>08.Costos Puerto</td>")
            case 10  response.write("<td class=td-l>09.Cuota Compenst</td>")
            case 11  response.write("<td class=td-l>10.Costos P. Conted</td>")
            case 12  response.write("<td class=td-l>11.Maniobras</td>")

            case 13  response.write("<td class=td-l>12.Otros</td>")
            case 14  response.write("<td class=td-l>13.------</td>")

            case 15  response.write("<td class=td-l>14.Revision Mercancia</td>")
            case 16  response.write("<td class=td-l>15.Flete Ext import</td>")

            case 17  response.write("<td class=td-l>16.Etiquetado</td>")
            case 18  response.write("<td class=td-l>17.F Maritimo</td>")

            case 19  response.write("<td class=td-l>Subtotal sin IGI</td>")
            case 20  response.write("<td class=td-l>IGI</td>")
            case 21  response.write("<td class=td-l>Costo total</td>")
        end select


              x=1
              while not rsUpdateEntry.EOF and i=1   
                        response.write("<td class=td-c style='background-color: #EAE8E8'><B>" )
                                  strSQL="select DocNum  from import2 where CardName='"& rsUpdateEntry("CardName") &"' and IGI= " & rsUpdateEntry("IGI2")  &" and shipper= " & rsUpdateEntry("shipper2") 
                                  select case request("fRS")
                                      case "DMX" rsUpdateEntry2.Open strSQL, adoCon2  'DMX
                                      case "DFW" rsUpdateEntry2.Open strSQL, adoCon3  'DFW
                                      case "MME" rsUpdateEntry2.Open strSQL, adoCon5  'MEIDE
                                  end select
                                  
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
                        response.write("<td class=td-l>"& rsUpdateEntry("DTA") &" MXP </td>" )         
                        rsUpdateEntry.MoveNext                     
              wend
 
             while not rsUpdateEntry.EOF and i=3     
                        response.write("<td class=td-l>"& rsUpdateEntry("PRVCNT") &" MXP </td>") 
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=4     
                        response.write("<td class=td-l>"& rsUpdateEntry("shipper") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=5     
                        response.write("<td class=td-l>"& rsUpdateEntry("fletes") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=6     
                        response.write("<td class=td-l>"& rsUpdateEntry("otros") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=7     
                        response.write("<td class=td-l>"& rsUpdateEntry("Honorarios") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=8     
                        response.write("<td class=td-l>"& rsUpdateEntry("Honorarios2") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

              while not rsUpdateEntry.EOF and i=9     
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional1") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=10     
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional2") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=11     
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional3") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=12    
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional4") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend

             while not rsUpdateEntry.EOF and i=13    
                         response.write("<td class=td-l>"& rsUpdateEntry("Adicional7") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                        
             wend


              while not rsUpdateEntry.EOF and i=14    'NO SE USA EN ESTE CONTEXTO'
                        response.write("<td class=td-l>0 MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend



             while not rsUpdateEntry.EOF and i=15   
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional5") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend
             while not rsUpdateEntry.EOF and i=16     
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional6") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend
              while not rsUpdateEntry.EOF and i=17     
                        response.write("<td class=td-l>"& rsUpdateEntry("labeling") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend
              while not rsUpdateEntry.EOF and i=18     
                        response.write("<td class=td-l>"& rsUpdateEntry("Adicional8") &" MXP </td>" )                    
                        rsUpdateEntry.MoveNext                     
             wend


             if i=19 then
                for n=1 to contador-1
                      strSQL="SELECT isnull(sum(Otros+DTA+PRVCNT+fletes+Honorarios+honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+Adicional5+Adicional6+Adicional7+Adicional8+shipper+labeling),0) as calculo from import2 where Docnum in ("& Entradas(n) &")"                 
                      'response.write("<td class=td-l>"& strSQL &" MXP </td>" )   
                      select case request("fRS")
                          case "DMX" rsUpdateEntry2.Open strSQL, adoCon2  'DMX
                          case "DFW" rsUpdateEntry2.Open strSQL, adoCon3  'DFW
                          case "MME" rsUpdateEntry2.Open strSQL, adoCon5  'MEIDE
                      end select     
                      response.write("<td class=td-l style='background-color: #C0C0C0'><B>"& trim(rsUpdateEntry2("calculo")) &"</B></td>" )      
                      'response.write("<td class=td-l>"& strSQL &"</td>" )
                      rsUpdateEntry2.close   
                next
             end if

             while not rsUpdateEntry.EOF and i=20   
                        if  trim(rsUpdateEntry("IGI"))<>"0" then
                             response.write("<td class=td-l style='color: white;background-color: red'>"& rsUpdateEntry("IGI") &" MXP </td>" )                   
                        else
                             response.write("<td class=td-l>"& rsUpdateEntry("IGI") &" MXP </td>" )   
                        end if 
                        rsUpdateEntry.MoveNext                     
             wend

            if i=21 then
                for n=1 to contador-1
                      strSQL="SELECT isnull(sum(Otros+DTA+PRVCNT+fletes+Honorarios+honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+Adicional5+Adicional6+Adicional7+Adicional8+shipper+IGI+labeling),0) as calculo,isnull(sum(Otros+DTA+PRVCNT+fletes+Honorarios+honorarios2+Adicional1+Adicional2+Adicional3+Adicional4+Adicional5+Adicional6+Adicional7+shipper+labeling),0) as SinMaritimo,Sum(Adicional8) as Maritimo from import2 where Docnum in ("& Entradas(n) &")"                 
                      'response.write("<td class=td-l>"& strSQL &" MXP </td>" )   
                      select case request("fRS")
                          case "DMX" rsUpdateEntry2.Open strSQL, adoCon2  'DMX
                          case "DFW" rsUpdateEntry2.Open strSQL, adoCon3  'DFW
                          case "MME" rsUpdateEntry2.Open strSQL, adoCon5  'MEIDE
                      end select      
                      response.write("<td class=td-l style='background-color: #86BBDB'><B>"& trim(rsUpdateEntry2("calculo")) &"</B></td>" )      
                      vSinMaritimo=""
                      vMaritimo=""
                      vSinMaritimo=trim(rsUpdateEntry2("SinMaritimo"))
                      vMaritimo=trim(rsUpdateEntry2("Maritimo"))
                      rsUpdateEntry2.close   
                next
             end if



           rsUpdateEntry.MoveFirst
           response.write("</tr>")
    next

     rsUpdateEntry.MoveFirst
     if int(trim(rsUpdateEntry("Adicional8")))>0 then
         response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")
         response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")
         response.write("<tr><td colspan=15 class='td-c'>")

         strSQL="select ("&vSinMaritimo &"+"& vMaritimo &")-(select sum(round("&vMaritimo&"*tasa_CBM,2)) from import1)-(select sum(round("&vSinMaritimo&"*tasa_valor,2)) from import1) as calculo"
         
         select case request("fRS")
              case "DMX" rsUpdateEntry7.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry7.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry7.Open strSQL, adoCon5  'MEIDE
          end select

         vCalculo=""
         vCalculo=trim(rsUpdateEntry7("calculo"))
         'response.write(strSQL&"<br>")
         rsUpdateEntry7.close

         strSQL="select ID,DocNum as Entrada,ItemCode as Codigo,Dscription as Descripcion,Quantity as Cant,subtotal as Valor,CBM,PESO,tasa_Valor as TasaValor,tasa_CBM as TasaCBM,tasa_PESO as TasaPESO,case when ID=1 then round("& vMaritimo &"*tasa_CBM+"&vSinMaritimo&"*tasa_valor,2)+"& vCalculo &" else round("& vMaritimo &"*tasa_CBM+"&vSinMaritimo&"*tasa_valor,2) end as enBASE_CBM,case when ID=1 then round("& vMaritimo &"*tasa_PESO+"&vSinMaritimo&"*tasa_valor,2)+"& vCalculo &" else round("& vMaritimo &"*tasa_PESO+"&vSinMaritimo&"*tasa_valor,2) end as enBASE_PESO from import1 order by ID"
         'response.write strSQL         
         select case request("fRS")
              case "DMX" rsUpdateEntry5.Open strSQL, adoCon2  'DMX
                         rsUpdateEntry6.Open strSQL, adoCon2  'DMX
              case "DFW" rsUpdateEntry5.Open strSQL, adoCon3  'DFW
                         rsUpdateEntry6.Open strSQL, adoCon3  'DFW
              case "MME" rsUpdateEntry5.Open strSQL, adoCon5  'MEIDE
                         rsUpdateEntry6.Open strSQL, adoCon5  'MEIDE
          end select
         

         Dim Campos(13) 
         i=1
         
         CreateTable(1300)
         response.write("<tr><td colspan=3 class='td-c' style='font-size: 20px; font-weight: 600px; color: red;'>IMPUTACION DE COSTOS</td></tr>")
         response.write("<tr><td class='td-c' width='500px'>")
                  
         For Each rsUpdateEntry5 in rsUpdateEntry5.Fields                        
            campos(i)=rsUpdateEntry5.Name
            i=i+1           
         Next   
    
         CreateTable(1000)
         RowIn
         for i=1 to 11
             Response.Write "<td class='subtitulo td-c'>" & campos(i) & "</td>"     
         next
         RowOut
         while not rsUpdateEntry6.EOF
               RowIn
               for n=1 to 3
                   response.write("<td class='td-c'>"& rsUpdateEntry6(campos(n))   &"</td>")
               next
               response.write("<td class='td-c'>"& mid(rsUpdateEntry6(campos(4)),1,26)   &"</td>")
               for n=5 to 11
                   response.write("<td class='td-c'>"& rsUpdateEntry6(campos(n))   &"</td>")
               next
               rsUpdateEntry6.MoveNext
               RowOut
         wend         
         closetable

         response.write("<td class='td-c' width='20px'>&nbsp;</td><td class='td-c' width='300px'>")
         rsUpdateEntry6.MoveFirst
         CreateTable(280)
         RowIn
         Response.Write "<td class='subtitulo td-c'>" & campos(12) & "</td>"    
         Response.Write "<td class='subtitulo td-c'>" & campos(13) & "</td>"           
         RowOut

         while not rsUpdateEntry6.EOF
            RowIn
            response.write("<td class='td-c'>"& rsUpdateEntry6(campos(12)) &" MXP"   &"</td>")
            response.write("<td class='td-c'>"& rsUpdateEntry6(campos(13)) &" MXP"   &"</td>")
            rsUpdateEntry6.MoveNext
            RowOut
         wend         
         closetable
         rsUpdateEntry6.close

         response.write("</td></tr>")
         closetable

         response.write("<tr><td colspan=15 class=td-c>&nbsp; </td></tr> ")
     end if

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

