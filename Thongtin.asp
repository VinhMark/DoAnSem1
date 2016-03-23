<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="DangNhap.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim thongtin__MMColParam
thongtin__MMColParam = "1"
If (Request.Form("MaSP") <> "") Then 
  thongtin__MMColParam = Request.Form("MaSP")
End If
%>
<%
Dim thongtin
Dim thongtin_cmd
Dim thongtin_numRows

Set thongtin_cmd = Server.CreateObject ("ADODB.Command")
thongtin_cmd.ActiveConnection = MM_connect_STRING
thongtin_cmd.CommandText = "SELECT * FROM dbo.Sach WHERE MaSach = ?" 
thongtin_cmd.Prepared = true
thongtin_cmd.Parameters.Append thongtin_cmd.CreateParameter("param1", 5, 1, -1, thongtin__MMColParam) ' adDouble

Set thongtin = thongtin_cmd.Execute
thongtin_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Free Leoshop Website Template | Single:: w3layouts</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href="css/form.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>
<link href="css/mystyle.css" rel="stylesheet" type="text/css" />
<script src="js/jquery1.min.js"></script>
<!-- start menu -->
<link href="css/megamenu.css" rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" src="js/megamenu.js"></script>
<script>$(document).ready(function(){$(".megamenu").megamenu();});</script>
<script type="text/javascript" src="js/jquery.jscrollpane.min.js"></script>
		<script type="text/javascript" id="sourcecode">
			$(function()
			{
				$('.scroll-pane').jScrollPane();
			});
		</script>
<!-- start details -->
<script src="js/slides.min.jquery.js"></script>
   <script>
		$(function(){
			$('#products').slides({
				preload: true,
				preloadImage: 'img/loading.gif',
				effect: 'slide, fade',
				crossfade: true,
				slideSpeed: 350,
				fadeSpeed: 500,
				generateNextPrev: true,
				generatePagination: false
			});
		});
	</script>
<link rel="stylesheet" href="css/etalage.css">
<script src="js/jquery.etalage.min.js"></script>
<script>
			jQuery(document).ready(function($){

				$('#etalage').etalage({
					thumb_image_width: 360,
					thumb_image_height: 360,
					source_image_width: 900,
					source_image_height: 900,
					show_hint: true,
					click_callback: function(image_anchor, instance_id){
						alert('Callback example:\nYou clicked on an image with the anchor: "'+image_anchor+'"\n(in Etalage instance: "'+instance_id+'")');
					}
				});

			});
		</script>	
</head>

<body>
<div class="header-top">
  <div class="wrap"> 
    <div class="cssmenu">
				<ul>
					<li class="active"></li> |
					<li><a href="login.html">Đăng Nhập</a></li> |
					<li><a href="register.html">Đăng Ký</a></li>
				</ul>
			</div>
			<div class="clear"></div>
  </div>
	</div>
	<div class="header-bottom">
	    <div class="wrap">
			<div class="header-bottom-left">
				<div class="logo">
					<a href="index.asp"><img src="images/logo.png" alt=""/></a>
				</div>
				<div class="menu">
	            <ul class="megamenu skyblue">
			<li class="active grid"><a href="index.asp">Trang chủ</a></li>
			<li><a class="color4" href="#">Thể loại</a>
				<div class="megapanel">
					<div class="row">
						<div class="col1">
							<div class="h_nav">
								<h4>Contact Lenses</h4>
								<ul>
									<li><a href="womens.html">Daily-wear soft lenses</a></li>
									<li><a href="womens.html">Extended-wear</a></li>
									<li><a href="womens.html">Lorem ipsum </a></li>
									<li><a href="womens.html">Planned replacement</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Sun Glasses</h4>
								<ul>
									<li><a href="womens.html">Heart-Shaped</a></li>
									<li><a href="womens.html">Square-Shaped</a></li>
									<li><a href="womens.html">Round-Shaped</a></li>
									<li><a href="womens.html">Oval-Shaped</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Eye Glasses</h4>
								<ul>
									<li><a href="womens.html">Anti Reflective</a></li>
									<li><a href="womens.html">Aspheric</a></li>
									<li><a href="womens.html">Bifocal</a></li>
									<li><a href="womens.html">Hi-index</a></li>
									<li><a href="womens.html">Progressive</a></li>
								</ul>	
							</div>												
						</div>
				  </div>
			  </div>
				</li>				
				<li><a class="color5" href="#">Hổ trợ</a>
				<div class="megapanel">
					<div class="col1">
							<div class="h_nav">
								<h4>Contact Lenses</h4>
								<ul>
									<li><a href="mens.html">Daily-wear soft lenses</a></li>
									<li><a href="mens.html">Extended-wear</a></li>
									<li><a href="mens.html">Lorem ipsum </a></li>
									<li><a href="mens.html">Planned replacement</a></li>
								</ul>	
							</div>							
				  </div>
						<div class="col1">
							<div class="h_nav">
								<h4>Sun Glasses</h4>
								<ul>
									<li><a href="mens.html">Heart-Shaped</a></li>
									<li><a href="mens.html">Square-Shaped</a></li>
									<li><a href="mens.html">Round-Shaped</a></li>
									<li><a href="mens.html">Oval-Shaped</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Eye Glasses</h4>
								<ul>
									<li><a href="mens.html">Anti Reflective</a></li>
									<li><a href="mens.html">Aspheric</a></li>
									<li><a href="mens.html">Bifocal</a></li>
									<li><a href="mens.html">Hi-index</a></li>
									<li><a href="mens.html">Progressive</a></li>
								</ul>	
							</div>												
						</div>
				  </div>
				</li>
			</ul>
			</div>
		</div>
		  <div class="clear"></div>
     </div>
	</div>
<div class="mens">    
  <div class="main">
     <div class="wrap">
<div class="cont span_2_of_3">
	  	  <div class="grid images_3_of_2">
						
						 <div class="clearfix"><img src="<%=(thongtin.Fields.Item("HinhAnh").Value)%>" alt="" name="" width="265" height="339" /></div>
          </div>
		         <div class="desc1 span_3_of_2">
		         	<h3 class="m_3"> Tên : <%=(thongtin.Fields.Item("TenSach").Value)%></h3>
		             <p class="m_5"> Giá : <%=(thongtin.Fields.Item("Gia").Value)%> VNĐ</p>
		         	 <div class="btn_form">
						<form action="Giohang.asp" method="post">
							<input type="submit" value="Mua" title="">
					      <input name="MaSP" type="hidden" id="MaSP" value="<%=(thongtin.Fields.Item("MaSach").Value)%>" />
						  <input name="TenSP" type="hidden" id="TenSP" value="<%=(thongtin.Fields.Item("TenSach").Value)%>" />
						  <input name="HinhAnhSP" type="hidden" id="HinhAnhSP" value="<%=(thongtin.Fields.Item("HinhAnh").Value)%>" />
						  <input name="GiaSP" type="hidden" id="GiaSP" value="<%=(thongtin.Fields.Item("Gia").Value)%>" />
						</form>
					 </div>
                     <div class="m_3">
                       <p>Tác giả:                     <%=(thongtin.Fields.Item("TacGia").Value)%></p>
                       <p>&nbsp;</p>
                     </div>
                   <div class="m_3">
                       <p>Thể loại : <%=(thongtin.Fields.Item("TheLoai").Value)%></p>
                       <p>&nbsp;</p>
                     </div>
					<span class="m_link">Nội dung : </span> <i class="m_text2"><%=(thongtin.Fields.Item("MoTa").Value)%></i>			     </div>
			   <div class="clear"></div>	
	    <div class="clients">
	    <h3 class="m_3">10 Other Products in the same category</h3>
		 
	
	
     </div>
     
     
      </div><div class="clear"></div>
	</div>
			 <div class="clear"></div>
  </div>
</div>
	<div class="footer">
		<div class="footer-top">
			<div class="wrap">
			  <div class="section group example">
				<div class="col_1_of_2 span_1_of_2">
					<ul class="f-list">
					  <li><img src="images/2.png"><span class="f-text">Giao hàng miễn phí</span><div class="clear"></div></li>
					</ul>
				</div>
				<div class="col_1_of_2 span_1_of_2">
					<ul class="f-list">
					  <li><img src="images/3.png"><span class="f-text">Điện Thoại 0908070605 </span><div class="clear"></div></li>
					</ul>
				</div>
				<div class="clear"></div>
		      </div>
			</div>
		</div>
		
		<div class="footer-bottom">
			<div class="wrap">
	                <div class="copy">
			           <p>© 2014 Template by <a href="http://w3layouts.com" target="_blank">w3layouts</a></p>
		            </div>
				<div class="clear"></div>
		    </div>
		</div>
		
</body>
</h
>
<%
thongtin.Close()
Set thongtin = Nothing
%>
tml>