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
<html>
<% 

if request.form("Submit") = "Enviar" Then

  function enviar_email(de_email,de_nome,para_email,para_nome,assunto,texto)
     Set objmail = Server.CreateObject("CDONTS.NewMail") 

     objmail.from = de_email
     objmail.to = para_email
     objmail.subject = assunto 
     objmail.body = texto
     objmail.send

     Response.Write "Email enviado para : " & para_nome & "  " & para_emailtime & "<BR>"
     set objmail = nothing
  end function

  set conn = server.createobject("adodb.connection") 
  conn.open "DSN=bancodedados;UID=admin;PWD=senha" 
  set rsquery = conn.execute("select nome, email from tabela") 

  nr_geral = 0 


  assunto = request.form("assunto")
  texto = request.form("texto")

  while not rsquery.eof 
     Bodytexto="Olá " & rsquery("nome") & "," & chr(13) & chr(13) & texto & chr(13) & chr(13) & "Shopmedia - www.shopmedia.com.br - Internet como você quer!!!"
     enviar_email "shopmedia@shopmedia.com.br","ShopMedia",rsquery("email"),rsquery("nome"),assunto,Bodytexto
     nr_geral = nr_geral + 1
     rsquery.movenext
  wend 

  Response.write "Numero Total de Emails: " & nr_geral 

  rsquery.close
  set rsquery = nothing 
  set conn = nothing

end if

%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>Shopmedia Networks Brasil</title>
</head>

<body>

<form METHOD="POST" ACTION="cdontsmail_bd.asp" name>
  <div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="100%" valign="top" align="center"><font face="Verdana" size="2"><br>
      </font><a href="http://www.shopmedia.com.br"><font face="Verdana" size="2" color="#000080"><strong>Shopmedia
      Networks Brasil</strong></font></a><p><font face="Verdana" size="2"><br>
      </font></td>
    </tr>
  </table>
  </center></div><div align="center"><center><table border="0" cellpadding="0"
  cellspacing="0" width="100%">
    <tr>
      <td width="100%" valign="middle" align="center"><strong><font face="Verdana" size="2">Assunto:</font></strong></td>
    </tr>
    <tr>
      <td width="100%" valign="middle" align="center"><font face="Verdana" size="2"><input
      TYPE="text" NAME="assunto" SIZE="30"><br>
      <br>
      </font></td>
    </tr>
    <tr>
      <td width="100%" valign="middle" align="center"><strong><font face="Verdana" size="2">Mensagem</font></strong></td>
    </tr>
    <tr>
      <td width="100%" valign="middle" align="center"><font face="Verdana" size="2"><textarea
      NAME="texto" cols="40" rows="6"></textarea></font></td>
    </tr>
    <tr>
      <td width="100%" valign="middle" align="center"><font face="Verdana" size="2"><br>
      <input type="submit" name="Submit" value="Enviar"></font></td>
    </tr>
    <tr>
      <td width="100%" valign="middle" align="center"><br>
      <div align="center"><center><p><font face="Verdana" size="2">Abraços,</font></p>
      </center></div><div align="center"><center><p><font face="Verdana" size="2">=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=<br>
      Shopmedia Networks Brasil<br>
      Adilson <br>
      <a href="http://www.shopmedia.com.br">http://www.shopmedia.com.br</a><br>
      Internet como você quer!<br>
      =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=<br>
      Icq: 15765673</font></td>
    </tr>
  </table>
  </center></div><div align="center"><center><p><font face="Verdana" size="2">&nbsp; </font></p>
  </center></div>
</form>
</body>
</html>
