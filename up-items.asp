<!--#include virtual="config/conn.asp"-->

<%
vcontrol=0
vcontrol= int(trim(request("control")))

select case vcontrol
case 1
	strSQL="UPDATE [SBOTemp].dbo.Notificaciones SET notas='" & request("fnotas")  &"',sentdate=null,MargenUtil=null, control='ACTUALIZACION: ' "
	strSQL=strSQL &"where tipo='" &  request("fItems")  &"' and RazonSocial='" &  request("fRS") &"' and DocNum=" &  request("fDocNum")  &" and modulo='ventas' "
	'response.write strSQL
	rsUpdateEntry.Open strSQL, adoCon4


case 2   'REMISIONES PARA FACTURAS'
   
    strSQL="SELECT CASE WHEN DATEPART(HOUR, GETDATE() )>17 AND DATEPART(MINUTE, GETDATE() )>31 THEN CASE WHEN DATEPART(dw,getdate())=6 THEN  convert(varchar, DATEADD(MINUTE,3660+(55-DATEPART(MINUTE, GETDATE() )),GETDATE()), 120)    ELSE   convert(varchar, DATEADD(MINUTE,780+(55-DATEPART(MINUTE, GETDATE() )),GETDATE()), 120)     END   ELSE  convert(varchar, getdate(), 120)     END as calculo"
    response.write strSQL
    rsUpdateEntry5.Open strSQL, adoCon4
    vstring=trim(rsUpdateEntry5("calculo"))
        
    rsUpdateEntry5.close

	strSQL="SET DATEFORMAT YMD;UPDATE [SBOTemp].dbo.Notificaciones SET notas='" & request("fnotas")  &"',RazonHold='"& request("fHold") &"',SlpEmail='"&request("fSlpemail")&"',RequestDate='"& vstring &"',sentdate=null,facturar='"&  trim(request("fautoriza") )   &"' "
	if request("fIncoTerm")<>"" then
            strSQL=strSQL &",IncoTerm='"&request("fIncoTerm")&"'  "
    end if
	strSQL=strSQL &"where tipo='" &  request("fItems")  &"' and RazonSocial='" &  request("fRS") &"' and DocNum=" &  request("fDocNum")   &" and modulo='ventas' "


	response.write(strSQL&"<BR>")
	response.write( request("msg") &"<BR>")

    rsUpdateEntry.Open strSQL, adoCon4

case 3
	strSQL="UPDATE [SBOTemp].dbo.Notificaciones SET notas='" & request("fnotas")  &"',RequestDate=getdate(),facturar='"&  trim(request("fautoriza") )   &"' "
	strSQL=strSQL &"where tipo='" &  request("fItems")  &"' and RazonSocial='" &  request("fRS") &"' and DocNum=" &  request("fDocNum")   &" and modulo='ventas' "
	response.write(strSQL&"<BR>")
	response.write( request("msg") &"<BR>")

    rsUpdateEntry.Open strSQL, adoCon4    

case 4
	strSQL="UPDATE [SBOTemp].dbo.Notificaciones SET notas='" & request("fnotas")  &"',sentdate=null, control='UPDATE: PO "&  request("fRS") &"-',actualizado=1,date_update=getdate()  where tipo='" &  request("fItems")  &"' and RazonSocial='" &  request("fRS") &"' and DocNum=" &  request("fDocNum")  &" and modulo='compras' "	
	rsUpdateEntry.Open strSQL, adoCon4
	      
end select

if vcontrol<>10 then 
     response.redirect "/modules/helpdesk.asp?items="& request("fItems")  &"&msg=" & request("msg")
end if     
%>