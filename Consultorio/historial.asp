<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Consulta de historial de ficha de paciente</title>
</head>

<body>

<%
paciente = request("paciente")
paciente = cint(paciente)

set oConn =  Server.CreateObject("ADODB.Connection")
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
Set RSpac = oConn.Execute("select * from fichas where paciente = " & paciente & "")

if not RSpac.EOF then


    fecha = now

Dia = Day(fecha)
Mes = Month(fecha)
Anio = Year(fecha)

fecha = Dia &"/"& Mes &"/"& Anio


%>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="538">
    <tr>
      <td height="134"><img border="0" src="../images/Sup_C.jpg"></td>
    </tr>
    <tr>
      <td height="357">
      <p align="center">
      <font face="Verdana" size="2" color="#000080"><span lang="es">
      Historial de la ficha de paciente</span></font></p>
      <div align="center">
        <center>
        <table border="1" cellpadding="10" cellspacing="10" style="border-collapse: collapse" bordercolor="#2D4773" width="90%" id="AutoNumber1">
          <tr>
            <td width="100%" height="60">
        <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="707" height="256">
        

          <tr>
            <td width="118" align="right" bgcolor="#2D4773" height="22">
            <span lang="es">
            <font size="2" face="Verdana" color="#FFFFFF">*Apellido/s 
            :</font></span></td>
            <td width="356" height="22">
            <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="apellido" size="49" value="<%=RSpac("apellido")%>"></td>
            </span>      
            <td width="183" rowspan="2" valign="top" height="59">
            <div align="center">
              <center>
              <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="65%" id="AutoNumber2">
                <tr>
                  <td width="100%" bgcolor="#2D4773">
                  <p align="center">
                  <font color="#FFFFFF" size="2" face="Verdana"><span lang="es">
                  Ficha Nº</span></font></p>
                 
                    
    
 <p align="center"><b><font color="#FFFFFF" size="5" face="Verdana"><span lang="es"><%=RSpac("paciente")%></span></font></b>
                  
                
                  
                  
                  </td>
                </tr>
              </table>
              </center>
            </div>
            </td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773" height="22">
            <span lang="es">
            <font size="2" face="Verdana" color="#FFFFFF">*Nombre/s :</font></span></td>
            <span lang="es">
            <td width="356" height="22">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="nombre" size="49" value="<%=RSpac("nombre")%>"></span></td>
          </tr>
          </span>
          </span></span>
          </span></span></span></span>
          <tr>
            </span></span></span></span></span></span></span></span>
          </tr>
     

          <tr>
            <td width="118" align="right" valign="top" bgcolor="#2D4773" height="133">
            <span lang="es"><font face="Verdana" size="2" color="#FFFFFF">
            Historial</font></span><font size="2" face="Verdana" color="#FFFFFF"> 
            :</font></td>
            <td width="554" colspan="2" height="133" valign="top">
            <div align="left">
            
 <%
set oConn2 =  Server.CreateObject("ADODB.Connection")
oConn2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/historial.mdb")
Set RSpac2 = oConn2.Execute("select * from historial where paciente = " & paciente & " order by id desc")

if not RSpac2.EOF then
%>
           
            
            
              <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
                <tr>
                  <td width="18%" bgcolor="#4065A2"><b>
                  <font color="#FFFFFF" size="2" face="Verdana"><span lang="es">
                  Fecha</span></font></b></td>
                  <td width="28%" bgcolor="#4065A2"><b>
                  <font color="#FFFFFF" size="2" face="Verdana"><span lang="es">
                  Responsable</span></font></b></td>
                  <td width="54%" bgcolor="#4065A2"><b>
                  <font color="#FFFFFF" size="2" face="Verdana"><span lang="es">
                  Acción</span></font></b></td>
                </tr>
                
 <%
 paso = 1
do while not RSpac2.eof
if paso = 1 then
%>               
    
                <tr>
                  <td width="18%"><font color="#000080" size="2" face="Verdana">
                  <span lang="es"><%=RSpac2("fecha")%></span></font></td>
                  <td width="28%"><font color="#000080" size="2" face="Verdana">
                  <span lang="es"><%=RSpac2("responsable")%></span></font></td>
                  <td width="54%"><font color="#000080" size="2" face="Verdana">
                  <span lang="es"><%=RSpac2("accion")%></span></font></td>
                </tr>
                
<%paso = 2
 else%>
                <tr>
                  <td width="18%" bgcolor="#CDD9EB">
                  <font color="#000080" size="2" face="Verdana"><span lang="es">
                  <%=RSpac2("fecha")%></span></font></td>
                  <td width="28%" bgcolor="#CDD9EB">
                  <font color="#000080" size="2" face="Verdana"><span lang="es">
                  <%=RSpac2("responsable")%></span></font></td>
                  <td width="54%" bgcolor="#CDD9EB">
                  <font color="#000080" size="2" face="Verdana"><span lang="es">
                  <%=RSpac2("accion")%></span></font></td>
                </tr>
                
 <%paso=1
end if
RSpac2.movenext
loop%>                
                
                
              </table>
            </div>
            </td>
          </tr>
          


<%end if
RSpac2.close
set RSpac2 = nothing
oConn2.Close
set oConn2 = nothing
%>         
          
          <tr>
            <td colspan="3" height="19">
            <p align="center">
            &nbsp;</td>
          </tr>
          

          </table>
        </center>
      </div></p>
          </tr>
        </table>
        </center>
      </div>
 <%end if
RSpac.close
set RSpac = nothing
oConn.Close
set oConn = nothing
%> </td>
    </tr>
    <tr>
      <td bgcolor="#2D4773" height="47">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <a href="menu.asp"><span lang="es"><font color="#FFFFFF">Volver al Menú</font></span><font color="#FFFFFF">
      </font></a>]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>