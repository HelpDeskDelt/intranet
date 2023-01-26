<!--#include virtual="/config/conn.asp"-->


<%
     
           Response.Expires = -1
                     
           response.write("<form id='vouchers' action='ShowContent.asp' method='POST'>")
           response.write("<input type='hidden' name='contenido' value='95'>")
           response.write("<input type='hidden' name='action' value='2'>")
           response.write("<input type='hidden' name='control' value='"& request("variable") &"'>")


           CreateTable(600)       

           if request("variable")="V" then   'viatico'
               response.write("<tr><td class='subtitulo td-r FontMed' width='35%'>ID_viatico:</td><td class='td-l FontMed' width='65%'><select name='fviatico'>")

               strSQL="select * from VIATICOS_R1 where Estatus!='cancelado' and DATEDIFF(DAY,DocDate,getdate())<100  order by ID desc"
               rsUpdateEntry.Open strSQL, adoCon 


               while not rsUpdateEntry.EOF
                   response.write("<option value='"& rsUpdateEntry("ID")&"'>"& rsUpdateEntry("ID")&"|"& rsUpdateEntry("almacen") &"|"& rsUpdateEntry("id_unidad") &"|"& rsUpdateEntry("motivo") &"</option>")
                   rsUpdateEntry.MoveNext
               wend
               response.write("</select></td></tr>")
               rsUpdateEntry.close
            
           response.write("<tr><td class='subtitulo td-r FontMed'>Tarjeta combustible:</td><td class='td-l FontMed'><select name='ftarjeta'>")

               strSQL="select * from flotilla where ID_unidad!='000' and tipo_activo<4"
               rsUpdateEntry2.Open strSQL, adoCon 

               while not rsUpdateEntry2.EOF
                   response.write("<option value='"& rsUpdateEntry2("tarjetaCombustible")&"'>"& rsUpdateEntry2("tarjetaCombustible")&"|"& rsUpdateEntry2("id_unidad") &"</option>")
                   rsUpdateEntry2.MoveNext
               wend
               response.write("</select></td></tr>")
               rsUpdateEntry2.close


           response.write("<tr><td class='subtitulo td-r FontMed'>Fecha de la carga:</td><td class='td-l FontMed'><input type='date' name='ffecha'></td></tr>")        
           response.write("<tr><td class='subtitulo td-r FontMed'>Estacion:</td><td class='td-l FontMed'><input type='text' name='festacion'></td></tr>")  
           response.write("<tr><td class='subtitulo td-r FontMed'>Cantidad en litros:</td><td class='td-l FontMed'><input type='text' name='flitros' placeholder='no utilice comas'></td></tr>")  


             response.write("<tr><td colspan=2 class='td-c FontMed'><input type='submit' value='continuar'></form></td></tr>")


          end if



          if request("variable")="O" then   'otros'
              response.write("<tr><td class='subtitulo td-r FontMed' width='35%'>Entidad Federativa Destino:</td><td class='td-l FontMed' width='65%'>")   %>

              <select id="entidad_destino" onchange="Javascript:PassValueChangeDiv('entidad_destino','causa','/modules/FormChangeStateVouchers.asp')">   <%  'elemento,divn,modulo  

           strSQL="select * from cat_entidades order by id_entidad"
           'response.write strSQL  
           rsUpdateEntry2.Open strSQL, adoCon 

          while not rsUpdateEntry2.EOF
              response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad")  &"</option>")             
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.Close
          response.write("</select></td></tr>")    

          end if

        
           response.write("<tr><td colspan=2 class='td-c FontMed'>&nbsp;</td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'><a href='ShowContent.asp?contenido=95'>regresar</a></td></tr>")


          CloseTable

%>
