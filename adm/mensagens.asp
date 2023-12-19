<!DOCTYPE html>
<html lang="pt-br">
<head>
<title>Apinotec - Prestadora de serviços</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta charset="utf-8">

<script type="application/x-javascript"> addEventListener("load", function() { setTimeout(hideURLbar, 0); }, false); function hideURLbar(){ window.scrollTo(0,1); } </script>
<!-- bootstrap-css -->
<link href="css/bootstrap.css" rel="stylesheet" type="text/css" media="all" />
<!--// bootstrap-css -->
<!-- css -->
<link rel="stylesheet" href="css/style.css" type="text/css" media="all" />
<link rel="stylesheet" href="css/tabela.css" type="text/css" media="all" />
<link rel="stylesheet" href="css/estilo2.css" type="text/css" media="all" />

<!--// css -->
<!-- font-awesome icons -->
<link href="css/font-awesome.css" rel="stylesheet"> 
<!-- //font-awesome icons -->
<!-- font -->
<link href="//fonts.googleapis.com/css?family=Roboto:100,100i,300,300i,400,400i,500,500i,700,700i,900,900i" rel="stylesheet">
<!-- //font -->
<script src="js/jquery-2.2.3.min.js"></script>
<script src="js/bootstrap.js"></script>
</head>
<body>
	
				<div class="w3-banner-1 jarallax">
		
			<div class="w3layouts-header-top">
				<div class="container">
					<div class="w3-header-top-grids">
						<div class="w3-header-top-left">
							<p><i class="fa fa-volume-control-phone" aria-hidden="true"></i> +55 11 99942-5085</p>
						</div>
						<div class="w3-header-top-right">
							<div class="agileinfo-social-grids">
								<ul>
									<li><a href="https://www.facebook.com/Apinotec-368486660285398/" target=_blank><i class="fa fa-facebook"></i></a></li>
									<li><a href="https://twitter.com/apinotecps" target=_blank><i class="fa fa-twitter" ></i></a></li>
								<li><a href="https://www.linkedin.com/in/apinotec-servi%C3%A7os-41253215a/" target=_blank><i class="fa fa-linkedin"></i></a></li>


								</ul>
							</div>
							<div class="clearfix"> </div>
						</div>
						<div class="clearfix"> </div>
					</div>
				</div>
			</div>

			<div class="head">
				<div class="container">
					<div class="navbar-top">
							<!-- Brand and toggle get grouped for better mobile display -->
							<div class="navbar-header">
							  <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
								<span class="sr-only">Toggle navigation</span>
								<span class="icon-bar"></span>
								<span class="icon-bar"></span>
								<span class="icon-bar"></span>
							  </button>
								 <div class="navbar-brand logo ">
									<h1 ><a href="indexadm.asp">Apinotec- Administração</a></h1>
									<h5>O ponto mais elevado da eficiência</h5>
								</div>

							</div>

							<!-- Collect the nav links, forms, and other content for toggling -->
							<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
							 <ul class="nav navbar-nav link-effect-4">
								<li ><a href="indexadm.asp">Home</a></li>
								<li class="active first-list"><a href="mensagens.asp">Mensagens</a></li>
								<li ><a href="avaliacao.asp">Avaliações</a></li>
								
									</ul>
								
							</div><!-- /.navbar-collapse -->
						</div>
				</div>
			</div>
		
		
	</div>
	<!-- //w3-banner -->
		<!-- about -->
	<!-- courses -->




	<div class="courses">
		<div class="container"> 
		<h2 class="agileits-title">Mensagens</h2> 
		<h4 align="center">Orçamentos e serviços solicitados</h4>
<hr style='width: 100%; height:5px; text-align:center; border-style: solid; border-width: 0.5px; border-color:#FF5722; background:#FF5722;' />
		<br>


  
		

	<%
nome= request.form("nome")
email= request.form("email")
mensagem= request.form("mensagem")
data=day(Now) & " / " & month(Now) & " / " & year(Now)  & " , " &  hour(Now) & " : " & minute(Now) & " : " & second(Now)

ConnString="Provider=Microsoft.jet.OLEDB.4.0;Data Source=c:\webpages\apinotecadm.mdb;Persist Security info=false;"
Set Conexao= Server.CreateObject("ADODB.Connection")
Conexao.Open ConnString
Set Registros= Server.CreateObject("ADODB.RecordSet")
Registros.Open "tabela1",Conexao
SQL = "select * from tabela1"

Set RS= conexao.execute(SQL)
%>
<div >
<%

response.write ("<table class='bordered striped centered'><tr><th>Nome<th>Email<th>Mensagem<th>Data/Hora</tr>")

While Not RS.EOF
response.write ("<tr>")
response.write ("<td>")
response.write (rs("nome"))

response.write ("<td>")
response.write (rs("email"))

response.write ("<td>")
response.write (rs("mensagem"))

response.write ("<td>")

response.write (right("0"&month(rs("data")),2)&"/"&Right("0"&day(rs("data")),2)&"/"&year(rs("data"))& " " & right("0"&hour(rs("data")),2)&":"&right("0"&minute(rs("data")),2))
response.write ("</tr>")
rs.movenext
wend

response.write("</table>")




		%>
			</div>
		</div>
		<div class="contact agileits">
		<div class="container" align="center">
			<form action="excluimsg.asp" method="post">
				<input  type="submit" value="Excluir registros">
			</form>
		</div>
		</div>
	</div>
	<!-- //courses -->
	<!-- //about -->
		<!-- footer -->
		<div class="copyright">
			<div class="container">
				<p>© 2018 Apinotec - Prestadora de serviços. Todos os direitos reservados | Design by <a href="http://w3layouts.com">W3layouts</a></p>
			</div>
		</div>
	<!-- //footer -->

	<script src="js/jarallax.js"></script>
	<script src="js/SmoothScroll.min.js"></script>
	<script type="text/javascript">
		/* init Jarallax */
		$('.jarallax').jarallax({
			speed: 0.5,
			imgWidth: 1366,
			imgHeight: 768
		})
	</script>
</body>	
</html>