<%@ Language=VBScript %>
<% Response.Buffer = True %>

<%
if Session("usuario") = "" then
response.redirect "../error_.asp"
end if
%>

<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="351">
    <tr>
      <td height="134"><img border="0" src="../../images/Sup_L.jpg"></td>
    </tr>
    <tr>
      <td height="143">
      <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="333" height="48">
          <tr>
            <td width="298" height="2" align="center" bgcolor="#2D4773"><b>
            <font color="#FFFFFF" face="Verdana" size="2"><span lang="es">
            Administración Laboratorio</span></font></b></td>
            </tr>
          <tr>
            <td width="298" height="1" align="center" bgcolor="#5F84C0">
            <span lang="es"><b><font face="Verdana" size="2" color="#FFFFFF">
            Planillas de Ordenes OSEP</font></b></span></td>
            </tr>
          </table>
        </center>
      </div>
<%

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
Set RSArt = oConn.Execute("select * from planillas order by id desc") 

if not RSArt.EOF then 
%>
      <div align="center">
        <center>
        <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="699">
          <tr>
            <td width="249" align="center">&nbsp;</td>
            <td width="44" align="center">&nbsp;</td>
            <td width="278" align="center" colspan="2">&nbsp;</td>
            <td width="52" align="center">&nbsp;</td>
            <td width="56" align="center">&nbsp;</td>
          </tr>
          <tr>
            <td align="center" bgcolor="#819FCD" width="249"><span lang="es"><b>
            <font face="Verdana" size="2" color="#FFFFFF">Período</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="44">&nbsp;</td>
            <td align="center" bgcolor="#819FCD" width="195"><span lang="es"><b>
            <font face="Verdana" size="2" color="#FFFFFF">Fecha de presentación</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="79"><span lang="es"><b>
            <font face="Verdana" size="2" color="#FFFFFF">Pagada ?</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="52"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Editar</font></span></b></td>
            <td align="center" bgcolor="#819FCD" width="56"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Borrar</font></span></b></td>
          </tr>
          
 <%
 paso = 1
do while not rsart.eof
if paso = 1 then
%>
          <tr>
            <td align="center" width="249"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("periodo")%></font></span></td>
            <td align="center" width="44"><span lang="es"><b>
            <font face="Verdana" size="2" color="#2D4773">
            <a href="ver_plan.asp?id=<%=RSArt("periodo")%>"><font color="#2D4773">Ver</font></a></font></b></span></td>
            <td align="center" width="195">
            <p align="left"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("fecha")%></font></span></td>
            <td align="center" width="79">
            <b><span lang="es">
            
                <%if RSArt("pagada") = "Si" then%>
             <b>
            <font color="#2D4773" size="2" face="Verdana">Si</font>
            <%else%></b> <b>
            <font size="2" face="Verdana" color="#FF0000">No</font>
            <%end if%></b></td>
            <td align="center" width="52"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="edit_plan.asp?id=<%=RSArt("id")%>"><font color="#2D4773">Editar</font></a></font></span></td>
            <td align="center" width="56"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a onclick="return confirm('Si lo borras no podrás recuperarlo. Estás seguro?')" href="borra_plan.asp?id=<%=RSArt("id")%>">
            <font color="#2D4773">Borrar</font></a></font></font></span></td>
          </tr>
 <%paso = 2
 else%>         
          <tr>
            <td align="center" bgcolor="#DBE3F0" width="249"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("periodo")%></font></span></td>
            <td align="center" bgcolor="#DBE3F0" width="44"><span lang="es"><b>
            <font face="Verdana" size="2" color="#2D4773">
            <a href="ver_plan.asp?id=<%=RSArt("periodo")%>">
            <font color="#2D4773">Ver</font></a></font></b></span></td>
            <td align="center" bgcolor="#DBE3F0" width="195">
            <p align="left"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("fecha")%></font></span></td>
            <td align="center" bgcolor="#DBE3F0" width="79">
             <%if RSArt("pagada") = "Si" then%>
             <b>
            <font color="#2D4773" size="2" face="Verdana">Si</font>
            <%else%></b> <b>
            <font size="2" face="Verdana" color="#FF0000">No</font>
            <%end if%></b>
</td>
            <td align="center" bgcolor="#DBE3F0" width="52"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="edit_plan.asp?id=<%=RSArt("id")%>"><font color="#2D4773">Editar</font></a></font></span></td>
            <td align="center" bgcolor="#DBE3F0" width="56"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a onclick="return confirm('Si lo borras no podrás recuperarlo. Estás seguro?')" href="borra_plan.asp?id=<%=RSArt("id")%>">
            <font color="#2D4773">Borrar</font></a></font></font></span></td>
          </tr>
  <%paso=1
  end if
  rsart.movenext
loop%>         
          
        </table>
        </center>
      </div>
  <%else
end if
RsArt.close
set RsArt = nothing
'oConn.Close
'set oConn = nothing%> <hr color="#2D4773" width="400" size="1">         
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[ 1.- Comenzar a cargar nueva planilla ]</span></font></b></p>
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.periodo.value == "")
  {
    alert("Escriba un valor para el campo \"periodo\".");
    theForm.periodo.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="do_pre_carga.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="639" bgcolor="#F4F7FB">
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Período</font></b></span><b><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              :</span></font></b></td>
              <td width="542">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="periodo" size="44"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: Enero a Marzo 
              2009)</font></span></td>
            </tr>
            </table>
          </center>
        </div>
        <p align="center">
        <input type="submit" value="Comenzar carga de ordenes" name="B1"></p>
      </form>
      <hr color="#2D4773" width="400" size="1">         
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[ 2.- Continuar cargando ordenes de siguiente planilla ]</span></font></b></p>
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form2_Validator(theForm)
{

  if (theForm.periodo.selectedIndex < 0)
  {
    alert("Elija una de las opciones \"periodo\".");
    theForm.periodo.focus();
    return (false);
  }

  if (theForm.periodo.selectedIndex == 0)
  {
    alert("La primera opción \"periodo\" no es válida. Elija una de las otras opciones.");
    theForm.periodo.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="do_pre_carga_ok.asp" onsubmit="return FrontPage_Form2_Validator(this)" language="JavaScript" name="FrontPage_Form2">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="639" bgcolor="#F4F7FB">
            <tr>
            <%

			'set oConn =  Server.CreateObject("ADODB.Connection")

			'oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
			Set RSArt = oConn.Execute("select * from planillas order by id") 

			if not RSArt.EOF then 
			%>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Período</font></b></span><b><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              :</span></font></b></td>
              <td width="542"><span lang="es">
              <!--webbot bot="Validation" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="periodo">
              <option selected>elija</option>
                        
 				<%

				do while not rsart.eof

					%>
              <option value="<%=RSArt("periodo")%>"><%=RSArt("periodo")%></option>
              <%RSArt.movenext
              loop%>
              </select>&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(elija una)</span></font></td>
                <%else
				end if
				RsArt.close
				set RsArt = nothing
				oConn.Close
				set oConn = nothing%>
              
            </tr>
            </table>
          </center>
        </div>
        <p align="center">
        <input type="submit" value="Continuar carga de ordenes" name="B1"></p>
      </form>
      <hr color="#2D4773" width="400" size="1">
      <p align="center"><b><font face="Verdana" size="2" color="#FFFFFF">
      <font color="#2D4773"><span lang="es">[ </span></font><span lang="es">
      <a href="index_adm.asp"><font color="#2D4773">Volver al Menú principal</font></a></span></font><font color="#2D4773" face="Verdana" size="2"><span lang="es"> 
      ]</span></font></b></p>
      <p align="center">&nbsp;</td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <span lang="es"><a href="abandon.asp"><font color="#FFFFFF">Desconectarse</font></a></span> ]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>