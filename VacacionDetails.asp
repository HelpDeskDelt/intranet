<!--#include virtual="/config/conn.asp"-->


<%
response.write ("<BR><BR>")
flag=1
strSQL="select a.*,dbo.fn_GetMonthName(a.fecha,'spanish') as Fecha2,b.Nombres+' '+b.Paterno+' '+b.Materno as colaborador from Incidencias a left join Empleados b on a.ID=b.ID where a.INCIDENCIA=4 and a.ID="& request("ID")  &" and a.PERIODO="& request("periodo") &" order by a.FECHA"
'response.write strSQL
rsUpdateEntry.Open strSQL, adoCon

if rsUpdateEntry.EOF then
     flag=0
     rsUpdateEntry.close

     strSQL="select Nombres+' '+Paterno+' '+Materno as colaborador from Empleados where ID=" &  request("ID")
     rsUpdateEntry.Open strSQL, adoCon
end if


select case request("accion")
   case "1" ShowDiasTomados   'Dias Tomados'
end select



sub ShowDiasTomados 

      response.write ("<H1>VACACIONES EJERCIDAS DEL PERIODO <font color=red>" & request("periodo") &"</font><BR>"& rsUpdateEntry("colaborador") &"</H1>")
      strSQL="select *,DiasDerecho-DiasTomados as pendientes from PeriodosVacacionales where ID="& request("ID")  &" and PERIODO="& request("periodo") 
      rsUpdateEntry2.Open strSQL, adoCon
      response.write ("<P style='font-size:2px'>&nbsp;</P>")

      response.write("<table width='400px' border=1 cellpadding=2 cellspacing=2 align=center>") 
      response.write("<tr><td class='td-r subtitulo' width='70%'> A&ntilde;os de antiguedad </td><td class='td-l' width='30%'> "& rsUpdateEntry2("AniosAntiguedad") &" </td></tr>")
      response.write("<tr><td class='td-r subtitulo'> Vacaciones para el per&iacute;odo </td><td class='td-l'> "& rsUpdateEntry2("DiasDerecho") &" </td></tr>")
      response.write("<tr><td class='td-r subtitulo'> D&iacute;as ejercidos </td><td class='td-l'> "& rsUpdateEntry2("DiasTomados") &" </td></tr>")
      response.write("<tr><td class='td-r subtitulo'> D&iacute;as pendientes  </td><td class='td-l'> "& rsUpdateEntry2("pendientes") &" </td></tr>")
      response.write("<tr><td class='td-r subtitulo'> D&iacute;as programados  </td><td class='td-l'> "& rsUpdateEntry2("DiasProgramados") &" </td></tr>")
      response.write("</table>")
      rsUpdateEntry2.close
      response.write ("<P style='font-size:5px'>&nbsp;</P>")
      if  request("periodo")=2019 then
                 response.write ("<H1>INFORMACION DISPONIBLE 2019</h1>")
      end if

      if flag=1 then
          response.write("<table width='460px' border=1 cellpadding=2 cellspacing=2 align=center>")       
          response.write("<tr><td class='td-c subtitulo'>#</td><td class='td-c subtitulo'>Fecha </td><td class='td-c subtitulo'>Horas</td><td class='td-c subtitulo'>Notas</td></tr>")
          i=1
          while not rsUpdateEntry.EOF
              response.write("<tr>")
              response.write("<td class='td-r'> "& i &" </td>")
              response.write("<td class='td-c'> "& rsUpdateEntry("Fecha2") &" </td>")
              response.write("<td class='td-r'> "& rsUpdateEntry("Horas") &" </td>")
              response.write("<td class='td-l'> "& rsUpdateEntry("Notas") &" </td>")
              response.write("</tr>")
              rsUpdateEntry.movenext
              i=i+1
          wend
           response.write("</table>")
       end if
       response.write ("<P>&nbsp;</P>")

end sub



rsUpdateEntry.close
%>
<input type="button" value="cerrar" onclick="hidediv('BoxRoundedDetailTop')">    