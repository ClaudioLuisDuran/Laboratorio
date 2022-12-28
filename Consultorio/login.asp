<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Login de entrada</title>
</head>

<body>
<%
odontologo = "Adriana Alessandro"

%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="490">
    <tr>
      <td height="134"><img border="0" src="../images/Sup_C.jpg"></td>
    </tr>
    <tr>
      <td height="23">
      <p align="center">&nbsp;</td>
    </tr>
    <tr>
      <td height="282">
      <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="677">
          <tr>
            <td width="188">
            <img border="0" src="images/sillon.jpg" width="218" height="182"></td>
            <td width="454">
            <p align="center"><span lang="es">      <%
     
    error = Request("error")
    

    if error = "yes" then
    
    %><b><font face="Georgia" size="2" color="#B02A38">Hubo un error de acceso.</font></b></p>
                  
              <p align="center" style="line-height: 150%">
              
              
                  <b><font face="Georgia" size="2" color="#B02A38">Intente de 
                  nuevo.</font></b></p>
                  
         <%else
         end if%>         
</span></p>
            <p align="center"><font color="#2D4773" face="Verdana" size="2">[
            <span lang="es">Ingrese sus datos <%=odontologo%></span>] </font>
            </p>
            <form method="POST" action="verificacion.asp">
              <div align="center">
                <center>
                <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="282" id="AutoNumber1">
                  <tr>
                    <td width="134" align="right"><span lang="es">
                    <font face="Verdana" size="2" color="#2D4773">Usuario:
                    </font></span></td>
                    <td width="148"><input type="text" name="usuario" size="20"></td>
                  </tr>
                  <tr>
                    <td width="134" align="right"><span lang="es">
                    <font face="Verdana" size="2" color="#2D4773">Contraseña :
                    </font></span></td>
                    <td width="148">
                    <input type="password" name="password" size="20"></td>
                  </tr>
                </table>
                </center>
              </div>
              <p align="center"><input type="submit" value="Ingresar" name="B1"></p>
            </form>
            <p align="center"><span lang="es"><a href="index.asp">
            <font face="Verdana" size="2" color="#000080">Salir</font></a></span></p>
            <p>&nbsp;</td>
          </tr>
          <tr>
            <td width="642" colspan="2">
            &nbsp;</td>
          </tr>
          </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      &nbsp;</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>