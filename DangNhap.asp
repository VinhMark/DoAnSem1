<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("txtTaikhoan"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "index.asp"
  MM_redirectLoginFailed = "ThongBaoDangNhapThatBai.asp"

  MM_loginSQL = "SELECT TaiKhoan, MatKhau"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.KhachHang WHERE TaiKhoan = ? AND MatKhau = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_connect_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 100, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("txtMatKhau")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>Leoshop</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>
<link href="css/mystyle.css" rel="stylesheet" type="text/css" />
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

<!-- kiểm tra login-->
<script type="text/javascript">
			function CheckValidInput(){
				/*alert('You have click botton Submit');*/
				
			var userId = document.getElementById('modlgn_username');
			var password = document.getElementById('modlgn_passwd')
			if(userId.value == '' || password.value == ''){
				alert('Bạn chưa nhập mật khẩu hoặc tài khoản!')
			
				}
			}
		</script>
 <style>
			input:required{
				outline:1px black solid;
				color:green;
			}
			input:required:valid{
				outline:1px yellow solid;
			}
			input:required:invalid{
				outline:1px red solid;
			}
</style>       
</head>

<body>
    <div class="header-top">
			<div class="wrap"> 
			  <div class="header-top-left">
			  	   

   				    <div class="clear"></div>
   			 </div>
			 <div class="cssmenu">
				<ul>
					<li><a href="DangNhap.asp">Đăng Nhập</a></li> |
					<li><a href="DangKy.asp">Đăng Ký</a></li>
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
	            
			</div>	
		</div>
	   
     <div class="clear"></div>
     </div>
	</div>
        <div class="login">
          	<div class="wrap">
				
				<div class="col_1_of_login span_1_of_login">
				<div class="login-title">
	           		<h4 class="title">ĐĂNG NHẬP</h4>
					<div id="loginbox" class="loginbox">
						<form action="<%=MM_LoginAction%>" method="POST" name="login" id="login-form">
						  <fieldset class="input">
						    <p id="login-form-username">
						      <label for="modlgn_username">Tài Khoản</label>
						      <input id="modlgn_username" type="text" name="txtTaikhoan" class="inputbox" size="18" required="True" placeholder="Tài Khản" onfocus="this.value = ''; autocomplete="off">
						    </p>
						    <p id="login-form-password">
						      <label for="modlgn_passwd">Mật Khẩu</label>
						      <input id="modlgn_passwd" type="password" name="txtMatKhau" class="inputbox" size="18" required="true" placeholder="Mật Khẩu" onfocus="this.value = ''; autocomplete="off">
						    </p>
						    <div class="remember">
							    <p id="login-form-remember">
							      <label for="modlgn_remember"><a href="#">Quên Mật Khẩu ? </a></label>
							   </p>
							    <input type="submit" name="Submit" class="button" value="Đăng Nhập" onclick="CheckValidInput()"><div class="clear"></div>
						    </div>
						  </fieldset>
					  </form>
					</div>
			    </div>
				</div>
				<div class="clear"></div>
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
