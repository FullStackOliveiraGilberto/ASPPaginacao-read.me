<!--#include file="conectando.asp"-->
<%

		'Selecionamos todos os Produtos da Tabela
Set rsMensagens = Server.CreateObject("ADODB.Recordset")
strMensagens = "SELECT * FROM mensagens ORDER BY msgid"
		rsMensagens.open strMensagens, conexao, 3, 3

'Definimos o Numero de Mensagens por Paginas com a propriedade "PageSize" do objeto Recordset
rsMensagens.PageSize = 10

'Criamos as Validações
if rsMensagens.eof then
   Mensagem = "Nenhuma mensagem foi escrita por enquanto."
   Response.End 
else
   'Definimos em qual pagina o visitante está
   if Request.QueryString("pg")="" then 
	  intpagina = 1
   else
	  if cint(Request.QueryString("pg"))<1 then
intpagina = 1
	  else
if cint(Request.QueryString("pg"))>rsMensagens.PageCount then  
	intpagina = rsMensagens.PageCount
		 else
	intpagina = Request.QueryString("pg")
end if
	  end if	
   end if   
		end if
%>

<%

data = date() %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script src="ieupdate.js"></script>
<title>Dom Vinicius</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-image: url(bg_geral.jpg);
}
.style1 {
	font-size: 12px;
	color: #000000;
	font-family: Arial, Helvetica, sans-serif;
}
a:link {
	color: #000000;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
	color: #97180F;
}
a:active {
	text-decoration: none;
}
.style3 {
	color: #EEDBBA;
	font-size: 16px;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
-->
</style></head>

<body>
<table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	<td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="900" height="247">
	  <param name="movie" value="topo_domvinicius.swf" />
	  <param name="quality" value="high" />
	  <param name="wmode" value="transparent">
	  <embed src="topo_domvinicius.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="900" height="247"></embed>
	</object></td>
  </tr>
  <tr>
	<td><table width="898" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
	  <tr>
		<td><table width="898" border="0" cellpadding="0" cellspacing="0">
		  <tr>
			<td width="882" valign="top" background="bg_mural.jpg"><table width="871" border="0" align="left" cellpadding="0" cellspacing="2">
			  
			  
			  <tr>
				<td width="206"> </td>
				<td width="201"> </td>
				<td width="456"> </td>
			  </tr>
			  <tr>
				<td><div align="center" class="style3">Mural de Recados </div></td>
				<td colspan="2"> </td>
				</tr>
			  <tr>
				<td> </td>
				<td colspan="2" class="style1">Deixe aqui o seu recado ou comentário. </td>
			  </tr>
			  <tr>
				<td> </td>
				<td colspan="2"><table width="615" border="0" cellpadding="0" cellspacing="2">
				  <form method="post" action="addscrap.asp"> <tr>
					<td width="65"> </td>
					<td width="544"> </td>
				  </tr>
				  <tr>
					<td valign="top" class="style1">Nome:</td>
					<td><input name="nome" type="text" class="style1" id="nome" size="50" /></td>
				  </tr>
				  <tr>
					<td valign="top" class="style1">Mensagem:</td>
					<td><textarea name="mensagem" cols="50" rows="5" class="style1" id="mensagem"></textarea>
					  <input name="data" type="hidden" id="data" value="<%=data%>" />					  </td>
				  </tr>
				  <tr>
					<td> </td>
					<td>
					  <label>
					  <input name="Submit" type="submit" class="style1" value="Enviar" /></form>
					  </label>
					</td>
				  </tr>
				  <tr>
					<td> </td>
					<td class="style1">Recados e Mensagens:</td>
				  </tr>
				  <tr>
					<td> </td>
					<td><%
   'Iniciamos o Loop
	rsMensagem.AbsolutePage = intpagina 
	intrec = 0
	While intrec<rsMensagem.PageSize and not rsMensagem.eof  %><table width="200" border="0">
  
<tr>
	<td>aaaaaaaaa</td>
  </tr>
  <tr>
	<td>aaaaaaaaaaaa</td>
  </tr>
</table><%
	rsMensagem.MoveNext
	intrec = intrec + 1
	if rsMensagem.eof then 
	   response.write " " 
	end if   
	Wend  
  %>
<% 
	'Criamos as Validações para a navegação "Anterior" e "Próximo"  
	if intpagina>1 then 
	%> 
	<a href="mural.asp?pg=<%=intpagina-1%>">Anterior</a> 
	<% 
	end if
	if StrComp(intpagina,rsMensagem.PageCount)<>0 then   
	%>
	<a href="mural.asp?pg=<%=intpagina + 1%>">Próximo</a>  
	<%
	end if
	rsMensagem.close
	Set rsMensagem = Nothing
	%>

					   </td>
				  </tr>
				</table></td>
				</tr>
			  <tr>
				<td> </td>
				<td colspan="2"> </td>
				</tr>
			  <tr>
				<td> </td>
				<td> </td>
				<td> </td>
			  </tr>
			</table></td>
			<td width="16" bgcolor="#F5EEDB"> </td>
			</tr>
		</table></td>
		</tr>
	</table></td>
  </tr>
  <tr>
	<td><div align="center"><span class="style1"><img src="barra_inferior.jpg" width="900" height="37" /></span></div></td>
  </tr>
  <tr>
	<td><div align="center"><span class="style1">Todos direitos reservados a DOM VINICIUS - 2008<br />
	Produzido por <a href="http://www.realinformaticarn.com.br" target="_blank">Real Informática</a></span></div></td>
  </tr>
  <tr>
	<td><div align="center" class="style1"></div></td>
  </tr>
</table>


<script>ieupdate()</script>
</body>
</html>


