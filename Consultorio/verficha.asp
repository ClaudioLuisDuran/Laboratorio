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
<title>Consulta de ficha de paciente</title>
</head>

<body>

<%

paciente = request("paciente")
paciente = cint(paciente)


set oConn =  Server.CreateObject("ADODB.Connection")
set RSpac =  Server.CreateObject("ADODB.Recordset")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")

'Set RSpac = oConn.Execute("select * from fichas where paciente = " & paciente & "")

SQL = "select * from fichas where paciente = " & paciente & ""



'set  RSpac=oconn.execute(SQL) 

RSpac.open SQL,oConn,1,3

if not RSpac.EOF then

%>

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
      <p align="center"><span lang="es">
      <font face="Verdana" size="2" color="#000080">Consulta de ficha de 
      paciente</font></span></p>
      <div align="center">
        <center>
        <table border="1" cellpadding="10" cellspacing="10" style="border-collapse: collapse" bordercolor="#2D4773" width="90%" id="AutoNumber1">
          <tr>
            <td width="100%">
        <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="707">
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Apellido/s</span><span lang="es"> 
            :</font></span></td>
            <td width="356" colspan="2">
            <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="apellido" size="49" value="<%=RSpac("apellido")%>"></td>
            </span>      
            <td width="183" rowspan="3" valign="top">
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
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Nombre/s :</font></span></td>
            <span lang="es">
            <td width="356" colspan="2">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="nombre" size="49" value="<%=RSpac("nombre")%>"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Obra Social :</font></td>
            <td width="356" colspan="2">
        <span lang="es">      
        <input type="text" name="obrasocial" size="49" value="<%=RSpac("obrasocial")%>"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Afiliado número :</font></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <input type="text" name="afiliadonro" size="49" value="<%=RSpac("afiliadonro")%>"></span></td>
          </tr>
          </span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Domicilio :</span></font><span lang="es"><span lang="es"></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="domicilio" size="61" value="<%=RSpac("domicilio")%>"></span></td>
          </tr>
          </span></span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Ciudad :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="ciudad" size="32" value="<%=RSpac("ciudad")%>"></span></td>
          </tr>
          </span></span></span></span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Provincia :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="257">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="provincia" size="25" value="<%=RSpac("provincia")%>"></span></td>
            </span></span></span></span></span></span></span></span>
            <td width="84" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">
            <span lang="es">*País :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="183">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="pais" size="25" value="<%=RSpac("pais")%>"></span></td>
          </tr>
          </span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Teléfono :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="telefono" size="43" value="<%=RSpac("telefono")%>"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Email :</font></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <input type="text" name="email" size="61" value="<%=RSpac("email")%>"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" valign="top" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Observaciones :</font></td>
            <td width="554" colspan="3">
            <textarea rows="9" name="observaciones" cols="57"><%=RSpac("observaciones")%></textarea></td>
          </tr>
          <tr>
            <td colspan="4">
            <p align="center"><b><span lang="es">
            <font face="Verdana" size="2" color="#000080">¿Que desea hacer 
            ahora?</font></span></b></p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080">
            <a href="editficha.asp?paciente=<%=paciente%>"><font color="#000080">
            Realizar cambios a esta ficha</font></a></font></span></p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080">
            <a href="diagnostico.asp?paciente=<%=paciente%>">
            <font color="#000080">Cargar odontograma del 
            paciente <%=RSpac("nombre")%>&nbsp;<%=RSpac("apellido")%></font></a></font></span></p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080"><a href="newpac.asp">
            <font color="#000080">Agregar fichas de nuevos pacientes</font></a></font></span></td>
          </tr>
          </table>
        </center>
      </div></p>
 </span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></td>
          </tr>
        </table>
        </center>
      </div>
 <%
 else
 
  
 paciente = paciente + 1
 response.redirect "verficha.asp?paciente="&paciente&""
 

 end if
RSpac.close
set RSpac = nothing
oConn.Close
set oConn = nothing
%> </td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <a href="menu.asp"><span lang="es"><font color="#FFFFFF">Volver al Menú</font></span><font color="#FFFFFF">
      </font></a>]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>