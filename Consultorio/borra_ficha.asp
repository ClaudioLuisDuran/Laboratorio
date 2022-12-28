<%@ Language=VBScript %>
<% Response.Buffer = True %>


<% 
  
Set oConn = Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")


SQL = "DELETE FROM fichas WHERE paciente = " & request("paciente") & ""

oConn.Execute(SQL)

' procedo a borrar odontogramas

Set oConn2 = Server.CreateObject("ADODB.Connection")

oConn2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")


SQL2 = "DELETE FROM corona WHERE paciente = " & request("paciente") & ""

oConn2.Execute(SQL2)


Set oConn3 = Server.CreateObject("ADODB.Connection")

oConn3.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")


SQL3 = "DELETE FROM extraccion WHERE paciente = " & request("paciente") & ""

oConn3.Execute(SQL3)


Set oConn4 = Server.CreateObject("ADODB.Connection")

oConn4.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")


SQL4 = "DELETE FROM odontograma WHERE paciente = " & request("paciente") & ""

oConn4.Execute(SQL4)


Set oConn5 = Server.CreateObject("ADODB.Connection")

oConn5.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")


SQL5 = "DELETE FROM odontograma2 WHERE paciente = " & request("paciente") & ""

oConn5.Execute(SQL5)



Response.Redirect "listado.asp"
 
 


%>