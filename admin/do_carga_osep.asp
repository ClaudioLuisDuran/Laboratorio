<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>

<%

Set oConn = Server.CreateObject("ADODB.Connection")

' grabo escrito

set RS = Server.CreateObject("ADODB.Recordset")  


oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")

             RS.Open "prestaciones",oConn,2,2
             
             RS.AddNew
             
             RS("periodo") = request.form("periodo")
             RS("mes") = request.form("mes")
             RS("anio") = request.form("anio")
             RS("cupon") = request.form("cupon")
             RS("profesional") = request.form("profesional")
             RS("afiliado") = request.form("afiliado")
             RS("nombre") = request.form("nombre")
             RS("codigo") = request.form("codigo")
             
                codigo_ok = request.form("codigo")
             	Set RSArt2 = oConn.Execute("select * from valores Where codigo LIKE '%" & codigo_ok & "%';") 
			 	if not RSArt2.EOF then 
             	descripcion_ok = RSArt2("descripcion")
             	end if
                RsArt2.close
			 	set RsArt2 = nothing
                          
             RS("descripcion") = descripcion_ok      
             RS("cantidad") = request.form("cantidad")
             
             	
             	cantidad_ok = request.form("cantidad")
             	
			 	Set RSArt = oConn.Execute("select * from valores Where codigo LIKE '%" & codigo_ok & "%';") 
			 	if not RSArt.EOF then 
             	valor_ok = RSArt("valor")
             	total_ok = valor_ok * cantidad_ok
             	end if
                RsArt.close
			 	set RsArt = nothing
			 
             RS("precio") = valor_ok
             RS("total") = total_ok
             RS("valorcupon") = request.form("valorcupon")
             RS("fechaosep") = request.form("fechaosep")
             RS("fechacarga") = now
             
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close

sigue = request("mas")

if sigue = "Si" then

Session("mes") = request.form("mes")
Session("anio") = request.form("anio")
Session("cupon") = request.form("cupon")
Session("profesional") = request.form("profesional")
Session("afiliado") = request.form("afiliado")
Session("nombre") = request.form("nombre")
Session("valorcupon") = request.form("valorcupon")
Session("fechaosep") = request.form("fechaosep")

else

Session("mes") = request.form("mes")
Session("anio") = request.form("anio")
Session("profesional") = ""
Session("cupon") = ""
Session("afiliado") = "" 
Session("nombre") = ""
Session("valorcupon") = ""
Session("fechaosep") = ""

end if

Response.Redirect "carga.asp"

%>

</body>

</html>