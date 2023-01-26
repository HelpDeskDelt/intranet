<!--#include virtual="/config/conn.asp"-->


<%
'response.write  request("carrier")
if request("remision") = "" then  
   
         %>  <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%


else

         

            'var variableNames = 'RS,remision';'
            strSQL="select  * from envios where RazonSocial='"& request("RS") &"' and DocNum=" & request("remision") &" and tipo='remision'" 
            rsUpdateEntry2.Open strSQL, adoCon4
             
            strSQL="select  * from ODLN where docnum="&request("remision") 
            flag=0

            if  request("RS")="DMX" then
               rsUpdateEntry.Open strSQL, adoCon2

               if rsUpdateEntry.EOF then
                  response.write ("<font class='fontbig td-l alert'>n&uacute;mero de remisi&oacute;n no existe, revise!</font>")
               else
                   flag=1
                 if rsUpdateEntry("DocCur") ="USD" then
                    response.write ("<font class='fontbig subtitulo'>RS: DMX / SN: "& rsUpdateEntry("cardname")  &", <br>Total documento en SAP: $ "& rsUpdateEntry("DocTotalFC") &" USD <BR>") 
                 else
                    response.write ("<font class='fontbig subtitulo'>RS: DMX / SN: "& rsUpdateEntry("cardname")  &", <br>Total documento en SAP: $ "& rsUpdateEntry("DocTotal") &" MXN <BR>") 
                 end if
                   rsUpdateEntry.close
                   if rsUpdateEntry2.EOF then
                         %>  <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%
                   else  %>
                         </font><font class=alert> Ya existe un env&iacute;o con esta remisi&oacute;n realizado en fecha: <%=rsUpdateEntry2("docdate")%> <BR>
                                                   Marque la casilla de re-env&iacute;o para registrar un segundo gasto</BR><font> 
                                                     <br> <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%
                   end if 
                   rsUpdateEntry2.close
                   response.write("</font><br><br>")
                end if
             end if


             if  request("RS")="DFW" then
               rsUpdateEntry.Open strSQL, adoCon3

               if rsUpdateEntry.EOF then
                  response.write ("<font class='fontbig td-l alert'>n&uacute;mero de remisi&oacute;n no existe, revise!</font>")
               else
                   flag=1
                 if rsUpdateEntry("DocCur") ="USD" then
                    response.write ("<font class='fontbig subtitulo'>RS: DFW / SN: "& rsUpdateEntry("cardname")  &", <br>Total documento en SAP: $ "& rsUpdateEntry("DocTotalFC") &" USD <br>")
                 else
                    response.write ("<font class='fontbig subtitulo'>RS: DFW / SN: "& rsUpdateEntry("cardname")  &", <br>Total documento en SAP: $ "& rsUpdateEntry("DocTotal") &" MXN <br>")
                 end if
                   rsUpdateEntry.close
                   if rsUpdateEntry2.EOF then
                         %>  <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%
                   else  %>
                         </font><font class=alert> Ya existe un env&iacute;o con esta remisi&oacute;n realizado en fecha: <%=rsUpdateEntry2("docdate")%> <BR>
                                                   Marque la casilla de re-env&iacute;o para registrar un segundo gasto</BR><font> 
                                                     <br> <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%
                   end if 
                   rsUpdateEntry2.close
                   response.write("</font><br><br>")
                end if
             end if


              if  request("RS")="MME" then
               rsUpdateEntry.Open strSQL, adoCon5

               if rsUpdateEntry.EOF then
                  response.write ("<font class='fontbig td-l alert'>n&uacute;mero de remisi&oacute;n no existe, revise!</font>")
               else
                   flag=1
                 if rsUpdateEntry("DocCur") ="USD" then
                    response.write ("<font class='fontbig subtitulo'>RS: MEIDE / SN: "& rsUpdateEntry("cardname")  &", <br>Total documento en SAP: $ "& rsUpdateEntry("DocTotalFC") &" USD <br>")
                 else
                    response.write ("<font class='fontbig subtitulo'>RS: MEIDE / SN: "& rsUpdateEntry("cardname")  &", <br>Total documento en SAP: $ "& rsUpdateEntry("DocTotal") &" MXN <br>")
                 end if
                   rsUpdateEntry.close
                   if rsUpdateEntry2.EOF then
                         %>  <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%
                   else  %>
                         </font><font class=alert> Ya existe un env&iacute;o con esta remisi&oacute;n realizado en fecha: <%=rsUpdateEntry2("docdate")%> <BR>
                                                   Marque la casilla de re-env&iacute;o para registrar un segundo gasto</BR><font> 
                                                     <br> <input type="button" value="registrar" onclick="SubmitForm('envios')">   <%
                   end if 
                   rsUpdateEntry2.close
                   response.write("</font><br><br>")
                end if
             end if


end if
%>