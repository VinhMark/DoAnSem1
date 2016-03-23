<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' *** Logout the current user.
MM_logoutRedirectPage = "DangNhap.asp"
Session.Contents.Remove("MM_Username")
Session.Contents.Remove("MM_UserAuthorization")
If (MM_logoutRedirectPage <> "") Then Response.Redirect(MM_logoutRedirectPage)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Free Leoshop Website Template | Login :: w3layouts</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>
<script type="text/javascript" src="js/jquery1.min.js"></script>
<!-- start menu -->
<link href="css/megamenu.css" rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" src="js/megamenu.js"></script>
<script>$(document).ready(function(){
			$(".megamenu").megamenu();
		});
</script>
<!-- dropdown -->
<script src="js/jquery.easydropdown.js"></script>
</head>
<body>
    <div class="header-top">
			<div class="wrap"> 
			  <div class="header-top-left">
			  	   

   				    <div class="clear"></div>
   			 </div>
			 <div class="cssmenu">
				<ul>
					<li class="active"><a href="login.html">Tài Khoản</a></li> |
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
					<a href="index.html"><img src="images/logo.png" alt=""/></a>
				</div>
				<div class="menu">
	            
			</div>	
		</div>
	   
     <div class="clear"></div>
     </div>
	</div>
        <div class="login">
          	<div class="wrap">
          	  <div class="clear">
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp; </p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	    <p>&nbsp;</p>
          	  </div>
		  </div>
		</div>
     <div class="footer" >
		<div class="footer-top" >
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
			           <p>© 2016 Sử Dụng Template Tại <a href="http://w3layouts.com" target="_blank">w3layouts</a></p>
		        </div>
			 
				<div class="clear"></div>
	      </div>
	   </div>
</div>
</body>
</html>
