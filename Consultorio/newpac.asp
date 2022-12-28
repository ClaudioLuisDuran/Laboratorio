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
DIM consulta

consulta = request("consulta")

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
<title>Nuevo paciente</title>
</head>

<body>



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
      <p align="center"><font face="Verdana" size="2" color="#000080"><span lang="es">
      <b>Ingreso de nueva ficha de paciente</b> [Odontólogo: <%=odontologo%> / 
      Matrícula: <%=matricula%> ]</span></font></p>
      <div align="center">
        <center>
        <table border="1" cellpadding="10" cellspacing="10" style="border-collapse: collapse" bordercolor="#2D4773" width="90%" id="AutoNumber1">
          <tr>
            <td width="100%">
            <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.apellido.value == "")
  {
    alert("Escriba un valor para el campo \"apellido\".");
    theForm.apellido.focus();
    return (false);
  }

  if (theForm.nombre.value == "")
  {
    alert("Escriba un valor para el campo \"nombre\".");
    theForm.nombre.focus();
    return (false);
  }

  if (theForm.domicilio.value == "")
  {
    alert("Escriba un valor para el campo \"domicilio\".");
    theForm.domicilio.focus();
    return (false);
  }

  if (theForm.ciudad.value == "")
  {
    alert("Escriba un valor para el campo \"ciudad\".");
    theForm.ciudad.focus();
    return (false);
  }

  if (theForm.provincia.value == "")
  {
    alert("Escriba un valor para el campo \"provincia\".");
    theForm.provincia.focus();
    return (false);
  }

  if (theForm.pais.value == "")
  {
    alert("Escriba un valor para el campo \"pais\".");
    theForm.pais.focus();
    return (false);
  }

  if (theForm.telefono.value == "")
  {
    alert("Escriba un valor para el campo \"telefono\".");
    theForm.telefono.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="carganewpac.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="707">
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Apellido/s</span><span lang="es"> 
            :</font></span></td>
            <td width="356" colspan="2">
            <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="apellido" size="49"></td>
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
                 
                    <%  
' recepcion de paciente
' paciente tal
  paciente = request("paciente")


set oConn =  Server.CreateObject("ADODB.Connection")
  oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
  Set RSArt = oConn.Execute("SELECT TOP 1 * from fichas ORDER BY paciente DESC") 
  
  if not RSArt.EOF then 
  
  ultimaficha = RSArt("paciente")
  nuevaficha = ultimaficha + 1
%>    
 <p align="center"><b><font color="#FFFFFF" size="5" face="Verdana"><span lang="es"><%=nuevaficha%></span></font></b>
                  
 <%end if
RsArt.close
set RsArt = nothing
oConn.Close
set oConn = nothing
%>                 
                  
                  
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
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="nombre" size="49"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Obra Social :</font></td>
            <td width="356" colspan="2">
        <span lang="es">      <input type="text" name="obrasocial" size="49"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Afiliado número :</font></td>
            <td width="554" colspan="3">
        <span lang="es">      <input type="text" name="afiliadonro" size="49"></span></td>
          </tr>
          </span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Domicilio :</span></font><span lang="es"><span lang="es"></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="domicilio" size="61"></span></td>
          </tr>
          </span></span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Ciudad :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="ciudad" size="32"></span></td>
          </tr>
          </span></span></span></span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Provincia :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="257">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="provincia" size="25" value="Mendoza"></span></td>
            </span></span></span></span></span></span></span></span>
            <td width="84" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">
            <span lang="es">*País :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="183">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="pais" size="25" value="Argentina"></span></td>
          </tr>
          </span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF"><span lang="es">*Teléfono :</span></font><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"><span lang="es"></td>
            <td width="554" colspan="3">
        <span lang="es">      
        <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="telefono" size="43"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Email :</font></td>
            <td width="554" colspan="3">
        <span lang="es">      <input type="text" name="email" size="61"></span></td>
          </tr>
          <tr>
            <td width="118" align="right" valign="top" bgcolor="#2D4773">
            <font size="2" face="Verdana" color="#FFFFFF">Observaciones :</font></td>
            <td width="554" colspan="3">
            <textarea rows="9" name="observaciones" cols="57"></textarea></td>
          </tr>
          <tr>
            <td colspan="4">
            <p align="center"><font face="Verdana">
            <input type="submit" value="Ingresar datos de paciente" name="B1"></font></td>
          </tr>
          </table>
        </center>
      </div>
        <input type="hidden" name="matricula" value="<%=matricula%>">
        <input type="hidden" name="odontologo" value="<%=odontologo%>">
        <input type="hidden" name="paciente" value="<%=nuevaficha%>">
        <input type="hidden" name="consulta" value="<%=consulta%>">
      </form></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></span></td>
          </tr>
        </table>
        </center>
      </div>
      

      </td>
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