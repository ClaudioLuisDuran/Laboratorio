<%@ Language=VBScript %>
<% Response.Buffer = True %>

<%

DIM UserName 
UserName = Request.form("usuario")
DIM Password 
Password = Request.form("password")
DIM oConn
DIM RSArt
DIM yes
DIM error
DIM visitas
DIM id

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/login.mdb")
Set RSArt = oconn.Execute("select * from registrados where usuario = '" & UserName & "'  and  password = '" & Password & "'  and estado = True ")

if not rsart.eof then

visitas = RSart("visitas")
visitas = visitas + 1
id = RSart("id")

Set rs = Server.CreateObject("ADODB.Recordset")

 SQL = "select * from registrados where id = " & id & ""

 rs.Open SQL, oConn, 2,3,1

			 RS("visitas") = visitas
			 RS("ultima") = Now

 rs.Update

rs.Close
set rs = nothing


  Session("allow_shopp") = True
  Session("usuario") = UserName
  Session("password") = Password
  Session.Timeout = 600

'Response.Redirect "admin_lux.asp?usuario="&UserName&""
Response.Redirect "menu.asp"

Else
yes = "yes"
Response.Redirect "login.asp?error="&yes&""
End If

rsart.close
'oconn.close

%>