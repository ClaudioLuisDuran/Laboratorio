<%@ Language=VBScript %>
<% Response.Buffer = True %>


<% 
  
Id = Request("Id")

Set oConn = Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")


SQL = "DELETE FROM planillas WHERE Id = " & request("Id") & ""

oConn.Execute(SQL)


Response.Redirect "precarga.asp"
 
 


%>