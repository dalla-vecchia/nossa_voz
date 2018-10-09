<! -------

Página criada por: Adilson - adilson@shopmedia.com.br
Webmaster Executivo - Shopmedia Networks Brasil

Se essa página lhe for útil, gostaria que entrasse em contato para trocarmos emails.

Abraços,,,,

=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Shopmedia Networks Brasil
http://www.shopmedia.com.br
Internet como você quer!
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Icq: 15765673

--------- >
<%
if Request("Action") = "Enviar" then
	nome = Request.Form("nome")
	email = Request.Form("email")
	icq = Request.Form("icq")
	site = Request.Form("site")
	assunto = Request.Form("assunto")
	mensagem = Request.Form("mensagem")

	Set JMail = Server.CreateObject("JMail.SMTPMail") 
	JMail.ServerAddress = "smtp.nomedoservidor.com.br"
	JMail.Sender = email
	JMail.Subject = "Shopmedia Networks Brasil"
	JMail.AddRecipient("shopmedia@shopmedia.com.br")
'	JMail.AddRecipient("outro@mail.com.br")
             JMail.silent = true

	corpo = "Shopmedia Networks Brasil" & chr(10)
	corpo = corpo & "Nome: " & nome & chr(10)
	corpo = corpo & "E-Mail: " & email & chr(10)
	corpo = corpo & "ICQ#: " & icq & chr(10)
	corpo = corpo & "Site: " & site & chr(10)
	corpo = corpo & "Assunto: " & assunto & chr(10)
	corpo = corpo & "Mensagem: " & mensagem & chr(10)
	JMail.Body = Corpo
	if JMail.Execute() then
		response.redirect "jmail_formulario.asp"
			
	else
		response.write "Seu nome é:" & nome & " e seu email é:" & email
		response.write "<html><head><title>Erro</title></head>"
		response.write "<body color=#FFFF00>"
		response.write "<p>Houve um erro ao enviar a sua mensagem.</p>"
		response.write "<p>A mensagem de erro é:" & JMail.ErrorMessage & "</p>"
		response.write "<p><a href='javascript:history.back()'>Clique aqui</a> para voltar e tente novamente.</p>"
		response.write "<p>&nbsp;</p></body></html>"
		response.end
	end if
	Set Jmail = Nothing
end if		
%>
<html>

<head>
<meta name="AUTOR" content="ADILSON">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<style TYPE="text/css">
<!--
	A {
       text-decoration: none; }
	A:hover { color:#CACA00;
}
-->
</style>
<title>Shopmedia Networks Brasil</title>
<style>
<a        { color: blue; text-decoration: none; }
   a:hover  { color: blue; text-decoration: underline; }
   select   { font-family: Verdana; font-weight: bold; }
   input    { font-family: Verdana; font-weight: bold; }
   textarea { font-family: Verdana; font-weight: bold; }>
</style>
</head>

<body bgcolor="#FF9224" link="#0000FF" vlink="#0000FF" alink="#0000FF" topmargin="0"
leftmargin="0">

<p><!--webbot bot="HTMLMarkup" startspan --><!--webbot bot="HTMLMarkup" endspan --></p>
<div align="center"><center>

<table border="0" cellpadding="2" width="100%">
  <tr>
    <td align="center" valign="top" width="50%"></td>
    <td align="center" valign="top" width="50%"></td>
  </tr>
</table>
</center></div><div align="center"><center>

<table border="0" cellpadding="10" cellspacing="0" width="600">
  <tr>
    <td align="center" valign="top" width="300" bgcolor="#D6D6AD"><p align="center"><font
    size="2" face="Verdana"><br>
    MODELO DE ENVIO DE EMAIL</font></p>
    <p align="center"><font size="2" face="Verdana">UTILIZANDO JMAIL</font></p>
    <p align="left">&nbsp;</p>
    <p align="center"><font color="#FF0000" size="2" face="Verdana">Obrigado pela visita!<br>
    <br>
    </font><font size="2" face="Verdana">=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=<br>
    Shopmedia Networks Brasil<br>
    Adilson <br>
    <a href="http://www.shopmedia.com.br">http://www.shopmedia.com.br</a><br>
    Internet como você quer!<br>
    =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=<br>
    Icq: 15765673</font></td>
    <td align="center" bgcolor="#D6D6AD"><form action="jmail_formulario.asp" method="POST">
      <input type="hidden" name="destination" value="spy@epn.com.br"><input type="hidden"
      name="feedback" value="http://www.epn.com.br/~spy/thanks.htm"><input type="hidden"
      name="subject" value="Visitante - Spy Home Page"><p><font size="2" face="Verdana"><br>
      <br>
      <br>
      Nome:<br>
      <input type="text" size="20" name="nome"><br>
      <br>
      E.mail:<br>
      <input type="text" size="20" name="email"><br>
      Icq::<br>
      <input type="text" size="15" maxlength="24" name="icq"><br>
      Site:<br>
      <input type="text" size="20" name="site"><br>
      Assunto: <br>
      <input type="text" size="20" maxlength="60" name="assunto"><br>
      Mensagem:<br>
      <textarea name="mensagem" rows="3" cols="25"></textarea><br>
      <br>
      <input type="submit" value="Enviar" name="Action"> <input type="reset" value="Cancelar"></font></p>
    </form>
    </td>
  </tr>
</table>
</center></div>
</body>
</html>
