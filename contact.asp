<!--A Design by W3layouts
Author: W3layout
Author URL: http://w3layouts.com
License: Creative Commons Attribution 3.0 Unported
License URL: http://creativecommons.org/licenses/by/3.0/
-->
<!DOCTYPE html>
<html lang="pt-br">
<head>
<title>Apinotec - Prestadora de serviços</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta charset="utf-8">
<meta name="keywords" content="Apinotec, serviços, manutenção, elétrica, hidráulica, páineis, alvenaria, páineis elétricos, instalação, instalações, reparo, conserto, construções, material para contruções, pedreiro" />
<script type="application/x-javascript"> addEventListener("load", function() { setTimeout(hideURLbar, 0); }, false); function hideURLbar(){ window.scrollTo(0,1); } </script>
<script language="VBScript" runat="Server">
   Sub Session_OnStart()
      'Session.LCID = 1046 é o formato padrão para Portugues/Brasil                                                                                      
       Session.LCID = 1046
   End Sub
 
   Sub Session_OnEnd
   End Sub
</script>
<!-- bootstrap-css -->
<link href="css/bootstrap.css" rel="stylesheet" type="text/css" media="all" />
<!--// bootstrap-css -->
<!-- css -->
<link rel="stylesheet" href="css/style.css" type="text/css" media="all" />
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
									<li><a href="https://twitter.com/apinotecps"><i class="fa fa-twitter" target=_blank></i></a></li>
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
									<h1 ><a href="index.html">Apinotec</a></h1>
									<h5>O ponto mais elevado da eficiência</h5>
								</div>

							</div>

							<!-- Collect the nav links, forms, and other content for toggling -->
							<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
							 <ul class="nav navbar-nav link-effect-4">
								<li ><a href="index.html">Home</a></li>
								<li ><a href="about.html">Sobre</a></li>
								<li><a href="gallery.html">Galeria</a></li>
								<li><a href="servicos.html">Serviços</a></li>
								<li><a href="avaliacao.html">Avalie</a></li>
						        <li class="active first-list"><a href="contact.html">Contato</a></li>
									</ul>
								
							</div><!-- /.navbar-collapse -->
						</div>
				</div>
			</div>
		
		
	</div>
	<!-- //w3-banner -->
	<!-- contact -->
	<div class="contact agileits">
		<div class="container"> 
			<h2 class="agileits-title">Contato</h2>
			<div class="contact-agileinfo">
				<%
nome= request.form("nome")
email= request.form("email")
mensagem= request.form("mensagem")


ConnString="Provider=Microsoft.jet.OLEDB.4.0;Data Source=c:\webpages\apinotecadm.mdb;Persist Security info=false;"
Set Conexao= Server.CreateObject("ADODB.Connection")
Conexao.Open ConnString
Set Registros= Server.CreateObject("ADODB.RecordSet")
Registros.Open "tabela1",Conexao
data=day(Now) & "/" & month(Now) & "/" & year(Now)  & " " &  hour(Now) & ":" & minute(Now) & ":" & second(Now)
SQL="Insert into tabela1 (nome,email,mensagem,data) values('"& nome &"','"& email &"','"& mensagem &"',#" & data & "#)"

Set RS= conexao.execute(SQL)

%>


					<div class="contact-text w3-agileits">
						<h4><%response.write ("Mensagem enviada com sucesso")%></h4>
						<p><i class="fa fa-map-marker"></i> Av. Carlos Veiga, 411, Eloy Chaves, Jundiaí, SP, Brasil </p>
						<p><i class="fa fa-phone"></i> Telefone: +55 11 99942-5085</p>
						<p><i class="fa fa-envelope-o"></i> Email: apinotec@gmail.com</p> 
					</div> 
				</div> 
				
				<div class="clearfix"> </div>	
			</div>
		</div>
	</div>
	<!-- //contact -->  
	<!-- map -->
	<div class="map w3layouts">  
		<iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3667.6410000041783!2d-46.965418384581405!3d-23.183299253625382!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x94cf253f80ce3dbf%3A0xa1a4c9e10214f222!2sAv.+Carlos+Veiga%2C+411+-+Parque+Res.+Eloy+Chaves%2C+Jundia%C3%AD+-+SP!5e0!3m2!1spt-BR!2sbr!4v1519050225669" width="600" height="450" frameborder="0" style="border:0" allowfullscreen></iframe>

		
	</div>
	<!-- //map -->  

	<!-- footer -->
			<div class="agileits-w3layouts-footer">
			<div class="container">
				<div class="col-md-4 w3-agile-grid">
					<h5>Sobre Nós</h5>
					<p>Conosco você pode encontrar atendimentos de elétrica, hidráulica, alvenaria, painéis e pintura.</p>
					<div class="footer-agileinfo-social">
						<ul>
							<li><a href="https://www.facebook.com/Apinotec-368486660285398/"><i class="fa fa-facebook"></i></a></li>
									<li><a href="https://twitter.com/apinotecps"><i class="fa fa-twitter"></i></a></li>
									<li><a href="https://www.linkedin.com/in/apinotec-servi%C3%A7os-41253215a/"><i class="fa fa-linkedin"></i></a></li>
							
						</ul>
					</div>
				</div>
			
				<div class="col-md-4 w3-agile-grid">
					<h5>Endereço</h5>
					<div class="w3-address">
						<div class="w3-address-grid">
							<div class="w3-address-left">
								<i class="fa fa-phone" aria-hidden="true"></i>
							</div>
							<div class="w3-address-right">
								<h6>Número de Telefone</h6>
								<p>+55 11 99942-5085</p>
							</div>
							<div class="clearfix"> </div>
						</div>
						<div class="w3-address-grid">
							<div class="w3-address-left">
								<i class="fa fa-envelope" aria-hidden="true"></i>
							</div>
							<div class="w3-address-right">
								<h6>Email</h6>
								<p>Email:<a href="mailto:apinotec@gmail.com"> apinotec@gmail.com</a></p>
							</div>
							<div class="clearfix"> </div>
						</div>
						<div class="w3-address-grid">
							<div class="w3-address-left">
								<i class="fa fa-map-marker" aria-hidden="true"></i>
							</div>
							<div class="w3-address-right">
								<h6>Localização</h6>
								<p>Av. Carlos Veiga, 411, Eloy Chaves
								<p> Jundiaí, SP, Brasil
							    </p></p>
							</div>
							<div class="clearfix"> </div>
						</div>
					</div>
				</div>
				<div class="clearfix"> </div>
			</div>
		</div>
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