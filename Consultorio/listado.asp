<%@ Language=VBScript %>
<% Response.Buffer = True %>

<%
DIM UserName 
UserName = Session("usuario")
DIM Password 
Password = Session("password")
DIM uConn
DIM RSu
DIM yes
DIM error
DIM odontologo
DIM matricula

set uConn =  Server.CreateObject("ADODB.Connection")

uConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/login.mdb")
Set RSu = uConn.Execute("select * from registrados where usuario = '" & UserName & "'  and  password = '" & Password & "'  and estado = True ")

if not RSu.eof then

  odontologo = RSu("nombre")
  matricula = RSu("matricula")
  Session("allow_shopp") = True
  Session.Timeout = 600

Else
yes = "yes"
Response.Redirect "login.asp?error="&yes&""
End If

RSu.close
uConn.close

%>

<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Listado de pacientes</title>
<script type="text/javascript" language="javascript">
function is_loaded() { //DOM
if (document.getElementById){
document.getElementById('preloader').style.visibility='hidden';
}else{
if (document.layers){ //NS4
document.preloader.visibility = 'hidden';
}
else { //IE4
document.all.preloader.style.visibility = 'hidden';
}
}
}
//-->
</script> 

<SCRIPT language="JavaScript"> 
function enviar() 
{ 
var donde_ir= confirm("Elija la opción");
if (donde_ir== true)
{ 
window.location="Continuar consulta anterior";
} 
else 
{ 
window.location="Iniciar nueva consulta";
}
} 
//-->
</SCRIPT>


</head>

<body onload="is_loaded();">


<div id="preloader" style="position:absolute; left:220; top:90; width:475; height:57">
<center>
<div align="center">
  <center>
<table border="0" cellpadding="6" cellspacing="0" style="font-family:Arial, Verdana; border: 2px solid #B02A38;; border-collapse:collapse" width="460">
<tr>
<td style="text-align:center; background-color:#B02A38" width="14">
<font style="font-weight:bold; color:#FFCE6F" size="6">!</font><font color="#FFCE6F" size="6">
</font>
</td>
<td style="text-align:center; background-color:#FED072" width="418">
<b><span lang="es"><font face="Georgia" size="2" color="#B02A38">Cargando 
listado de pacientes&nbsp; 
de <%=odontologo%></span></font></b><b><font face="Georgia"><font style="color:#B02A38; text-align:center" size="2"> ...</font><font color="#B02A38" size="2">
</font></font></b>
</td>
</tr>
</table>
  </center>
</div>
</center>
</div> 


<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="490">
    <tr>
      <td height="134"><img border="0" src="../images/Sup_C.jpg"></td>
    </tr>
    <tr>
      <td height="23">
      &nbsp;</td>
    </tr>
    <tr>
      <td height="282">
      <div align="center">
        <center>
        <table border="1" cellpadding="10" cellspacing="10" style="border-collapse: collapse" bordercolor="#2D4773" width="752" id="AutoNumber1">
          <tr>
            <td width="710">
            <p align="center"><font face="Verdana" color="#000080">
            <span lang="es"><br>
            Listado de pacientes de <%=odontologo%></span></font></p>
            <div align="center">
              <center>
              <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="729" id="AutoNumber2" height="56">
              
 <%

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
Set RSArt = oConn.Execute("select * from fichas order by paciente desc") 

if not RSArt.EOF then 
%>             
              
              
                <tr>
                  <td width="67" align="center" bgcolor="#000080" height="16"><b>
                  <span lang="es"><font size="2" color="#FFFFFF" face="Verdana">
                  Ficha Nº</font></span></b></td>
                  <td width="278" align="center" bgcolor="#000080" height="16"><b>
                  <span lang="es"><font size="2" color="#FFFFFF" face="Verdana">
                  Nombre</font></span></b></td>
                  <td width="341" align="center" bgcolor="#000080" colspan="5" height="16">
                  <span lang="es"><b>
                  <font face="Verdana" size="2" color="#FFFFFF">Opciones</font></b></span></td>
                </tr>
                
  <%
 paso = 1
do while not rsart.eof
if paso = 1 then
%>               
                
                
                
                <tr>
                  <td width="67" height="12"><span lang="es">
                  <font face="Verdana" size="2" color="#000080"><%=RSArt("paciente")%></font></span></td>
                  <td width="278" height="12">
                  <font face="Verdana" size="2" color="#000080">
                  <span lang="es">
                  <%=RSArt("nombre")%>&nbsp;<%=RSArt("apellido")%></span></font></td>
                  
                  <%if RSArt("estado_consulta") = "En curso" then%>
                  
                  <td width="68" height="12">
                  <p align="center">
                  <a  onclick="return confirm('Con esta opción continúa trabajando sobre una orden ya iniciada. ¿Procede?')" href="diagnostico.asp?paciente=<%=RSArt("paciente")%>">
                  <span lang="es"><font face="Verdana" size="2" color="#2D4773">
                  Consulta</font></span></a></td>
                  
                  <%else%>
                  
                  <td width="91" height="12">
                  <p align="center">
                  <a  onclick="return confirm('Esta opción inicia una nueva orden. ¿Procede?')" href="ini_cons.asp?paciente=<%=RSArt("paciente")%>">
                  <span lang="es"><font face="Verdana" size="2" color="#2D4773">
                  Nueva Orden</font></span></a></td>
                  
                  <%end if%>
                  
                  <td width="79" height="12">
                  <p align="center"><span lang="es">
                  <font face="Verdana" size="2" color="#2D4773">
                  <a  onclick="return confirm('Editar y actualizar datos de la ficha del paciente. ¿Procede?')" href="editficha.asp?paciente=<%=RSArt("paciente")%>"><font color="#2D4773">Actualizar</font></a></font></span></td>
                  <td width="64" height="12">
                  <p align="center"><span lang="es">
                  <a href="historial.asp?paciente=<%=RSArt("paciente")%>">
                  <font face="Verdana" size="2" color="#2D4773">Historial</font></a></span></td>
                  <td width="48" height="12">
                  <p align="center"><span lang="es">
                  <font face="Verdana" size="2" color="#2D4773">
                  <a onclick="return confirm('Esta opción borra definitivamente datos y ordenes del paciente. ¿Procede?')" href="borra_ficha.asp?paciente=<%=RSArt("paciente")%>">
                  <font color="#2D4773">Borrar</font></a></font></span></td>
                </tr>
  <%paso = 2
 else%>               
                
                <tr>
                  <td width="67" bgcolor="#CAEEFF" height="16"><span lang="es">
                  <font face="Verdana" size="2" color="#000080"><%=RSArt("paciente")%></font></span></td>
                  <td width="278" bgcolor="#CAEEFF" height="16">
                  <font face="Verdana" size="2" color="#000080">
                  <span lang="es">
                  <%=RSArt("nombre")%>&nbsp;<%=RSArt("apellido")%></span></font></td>
                  
                  
                  <%if RSArt("estado_consulta") = "En curso" then%>
                  
                  <td width="68" bgcolor="#CAEEFF" height="16">
                  <p align="center">
                  <a  onclick="return confirm('Con esta opción continúa trabajando sobre una orden ya iniciada. ¿Procede?')" href="diagnostico.asp?paciente=<%=RSArt("paciente")%>">
                  <span lang="es"><font face="Verdana" size="2" color="#2D4773">
                  Consulta</font></span></a></td>
                  
                  <%else%>
                  
                  <td width="91" bgcolor="#CAEEFF" height="16">
                  <p align="center">
                  <a  onclick="return confirm('Esta opción inicia una nueva orden. ¿Procede?')" href="ini_cons.asp?paciente=<%=RSArt("paciente")%>">
                  <span lang="es"><font face="Verdana" size="2" color="#2D4773">
                  Nueva Orden</font></span></a></td>
                  
                  <%end if%>
                  
                  
                  <td width="79" bgcolor="#CAEEFF" height="16">
                  <p align="center"><span lang="es">
                  <font face="Verdana" size="2" color="#2D4773">
                  <a  onclick="return confirm('Editar y actualizar datos de la ficha del paciente. ¿Procede?')" href="editficha.asp?paciente=<%=RSArt("paciente")%>"><font color="#2D4773">Actualizar</font></a></font></span></td>
                  <td width="64" bgcolor="#CAEEFF" height="16">
                  <p align="center"><span lang="es">
                  <a href="historial.asp?paciente=<%=RSArt("paciente")%>">
                  <font face="Verdana" size="2" color="#2D4773">Historial</font></a></span></td>
                  <td width="48" bgcolor="#CAEEFF" height="16">
                  <p align="center"><span lang="es">
                  <font face="Verdana" size="2" color="#2D4773">
                  <a onclick="return confirm('Esta opción borra definitivamente datos y ordenes del paciente. ¿Procede?')" href="borra_ficha.asp?paciente=<%=RSArt("paciente")%>">
                  <font color="#2D4773">Borrar</font></a></font></span></td>
                </tr>
   <%paso=1
  end if
  rsart.movenext
loop%>         

<%else
end if
RsArt.close
set RsArt = nothing
oConn.Close
set oConn = nothing%>                     
                
              </table>
              </center>
            </div>
            </td>
          </tr>
        </table>
        </center>
      </div>
      

      <p>&nbsp;</td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[ </a>
      <a href="menu.asp"><span lang="es"><font color="#FFFFFF">Volver al Menú</font></span><font color="#FFFFFF">
      </font></a>]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>