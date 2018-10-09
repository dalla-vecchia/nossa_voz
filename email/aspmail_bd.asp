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

<head>
<title>Shopmedia Networks Brasil</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
</head>
<% 
ConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="& Server.MapPath("/banco de dados.mdb")"

sql = "select nome, email from tabela"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnString

set rsquery = conn.execute(sql) 
nr_emails_corretos = 0
nr_emails_errados = 0
nr_geral = 0
while not rsquery.eof

Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.RemoteHost = "smtp.nomedoservidor.com.br" 
Mailer.FromName = "Shopmedia Networks Brasil"
Mailer.FromAddress = "shopmedia@shopmedia.com.br"
Mailer.AddRecipient rsquery("nome"),rsquery("email") 
Mailer.Subject=request.form("assunto")

Mailer.Bodytext = "Caro " & rsquery("nome") & "," & vbCrLf & request.form("texto")

x = Mailer.SendMail 
if x = true then 
MSG = "E-MAIL ENVIADO COM SUCESSO!"
nr_emails_corretos = nr_emails_corretos + 1
Else
MSG = " O E-MAIL NÃO FOI ENVIADO COM SUCESSO!" 
nr_emails_errados = nr_emails_errados + 1 
end if 
nr_geral = nr_geral + 1 
Response.write nr_geral & " - " & MSG 
rsquery.movenext
wend
Response.write "Numero Total de Emails: " & nr_geral & "<br>"
Response.write "Numero de Emails enviados: " & nr_emails_corretos & "<br>"
Response.write "Numero de Emails não enviados: " & nr_emails_errados & "<br>" 
rsquery.close
set rsquery = nothing 
conn.close
set conn = nothing 
%>

<body>
</body>
</html>
