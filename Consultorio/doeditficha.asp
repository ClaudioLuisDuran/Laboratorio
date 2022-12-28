<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nuevo paciente</title>
</head>

<body>

<%
nombre = request.form("nombre")
apellido = request.form("apellido")
paciente = request.form("paciente")

Set oConn = Server.CreateObject("ADODB.Connection")

' actualizo

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")

SQL = "select * from fichas where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3
Set oConn = Server.CreateObject("ADODB.Connection")

             
             RS("apellido") = request.form("apellido")
             RS("nombre") = request.form("nombre")
             RS("obrasocial") = request.form("obrasocial")
             RS("afiliadonro") = request.form("afiliadonro")
             RS("domicilio") = request.form("domicilio")
             RS("ciudad") = request.form("ciudad")
             RS("provincia") = request.form("provincia")
             RS("pais") = request.form("pais")                          
             RS("telefono") = request.form("telefono")     
             RS("email") = request.form("email")
             RS("odontologo") = request.form("odontologo")
             RS("matricula") = request.form("matricula")
             RS("observaciones") = request.form("observaciones")
             'RS("fecha") = now
             
             RS.Update
             RS.Close
             
set RS=nothing
'oConn.Close

'Actualiza Historial

accion = "Actualizacion de datos realizada"

set RS = Server.CreateObject("ADODB.Recordset")  
set oConn2 =  Server.CreateObject("ADODB.Connection")

oConn2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/historial.mdb")

             RS.Open "historial",oConn2,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("responsable") = request.form("odontologo")
             RS("accion") = accion
             RS("fecha") = request.form("fecha")                        
             RS.Update
             RS.Close
             
set RS=nothing
oConn2.Close




'Response.Redirect "carga.asp"

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
      <div align="center">
        <center>
        <table border="1" cellpadding="10" cellspacing="10" style="border-collapse: collapse" bordercolor="#2D4773" width="90%" id="AutoNumber1">
          <tr>
            <td width="100%">
            <p align="center"><font face="Verdana" color="#000080">
            <span lang="es"><br>
            [&nbsp;&nbsp;&nbsp; La ficha ha sido actualizada con éxito&nbsp;&nbsp;&nbsp; 
            ]</span></font></p>
            <p align="center">&nbsp;</p>
            <p align="center"><b><span lang="es">
            <font face="Verdana" size="2" color="#000080">¿Que desea hacer 
            ahora?</font></span></b></p>
            <p align="center">
            <font face="Verdana" size="2" color="#000080">
            <span lang="es">
            <a href="verficha.asp?paciente=<%=paciente%>"><font color="#000080">Ver ficha del paciente 
            <%=nombre%>&nbsp;<%=apellido%></font></a> </span></font></p>
            <p align="center">
            <font face="Verdana" size="2" color="#000080">
            <span lang="es">
            <a href="diagnostico.asp?paciente=<%=paciente%>">
            <font color="#000080">Cargar odontograma del 
            paciente <%=nombre%>&nbsp;<%=apellido%></font></a> </span></font></p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080"><a href="newpac.asp">
            <font color="#000080">Agregar fichas de nuevos pacientes</font></a></font></span></td>
          </tr>
        </table>
        </center>
      </div>
      

      <p>&nbsp;</td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <span lang="es">...</span> ]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>