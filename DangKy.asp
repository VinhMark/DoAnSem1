<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connect_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.KhachHang (TenKH, TaiKhoan, MatKhau, Email, DiaChi, NgaySinh, SDT) VALUES (?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("txtTenKhachHang")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 100, Request.Form("txtTenTAiKhoan")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("txtMatKhau1")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("txtEmail1")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 500, Request.Form("txtDiachi1")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 135, 1, -1, MM_IIF(Request.Form("txtNgaysinh1"), Request.Form("txtNgaysinh1"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 201, 1, 12, Request.Form("txtSDT1")) ' adLongVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "index.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Free Leoshop Website Template | Register :: w3layouts</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href="css/mystyle.css" rel="stylesheet" type="text/css" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>
<script type="text/javascript" src="js/jquery1.min.js"></script>
<!-- start menu -->
<link href="css/megamenu.css" rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" src="js/megamenu.js"></script>
<script>$(document).ready(function(){$(".megamenu").megamenu();});</script>
<script src="js/jquery.easydropdown.js"></script>
<!--xác nhận mật khẩu-->
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
<script type="text/javascript">
        function Validate() {
            var password = document.getElementById("txtMatKhau1").value;
            var confirmPassword = document.getElementById("txtMatKhau2").value;
            if (password != confirmPassword) {
                alert("Mật khẩu không trùng.");
            }
			var kiemtra = document.getElementById("txtMatKhau1").length
			if(kiemtra < 6 && kiemtra > 12){
				alert("Mật khẩu phải từ 6-12 kí tự");
				}
        }
 </script>
</head>
<body> 
	<div class="header-top">
			<div class="wrap">
			 <div class="cssmenu">
				<ul>
					<li class="active"></li> 
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
	            <ul class="megamenu skyblue">
			<li class="active grid"><a href="index.asp">Trang Chủ</a></li>
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
				<li><a class="color5" href="#">Hổ Trợ</a>
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
	   <div class="header-bottom-right">
         <div class="search">	  
				<input type="text" name="s" class="textbox" value="Search" onfocus="this.value = '';" onblur="if (this.value == '') {this.value = 'Search';}">
				<input type="submit" value="Subscribe" id="submit" name="submit">
				<div id="response"> </div>
		 </div>
	  <div class="tag-list">
		<ul class="icon1 sub-icon1 profile_img">
			<li><a class="active-icon c2" href="#"> </a>
				<ul class="sub-icon1 list">
					<li><h3>No Products</h3><a href=""></a></li>
					<li><p>Lorem ipsum dolor sit amet, consectetuer  <a href="">adipiscing elit, sed diam</a></p></li>
				</ul>
			</li>
		</ul>
	  </div>
    </div>
     <div class="clear"></div>
     </div>
	</div>
          <div class="register_account">
          	<div class="wrap">
    	      <h4 class="title">Tạo Tài Khoản</h4>
    		   <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
    			 <div class="col_1_of_2 span_1_of_2">
		   			 <div><input name="txtTenKhachHang" type="text" id="txtTenKhachHang" placeholder="Tên" onfocus="this.value = ''; value="Tên"></div>
		    			<div><input name="txtTenTAiKhoan" type="Text" required="true" placeholder="Tên tài khoản" id="txtTenTAiKhoan"  onfocus="this.value = ''; value="Tên tài khoản">
		    			</div>
		    			<div><input name="txtMatKhau1" type="Password" required="true" placeholder="Mật Khẩu"  min="6" max="12"  id="txtMatKhau1" onfocus="this.value = ''; value="Mật khẩu"></div>
		    			<div><input name="txtMatKhau2" type="Password" id="txtMatKhau2" required="true" placeholder="Xác nhận mật khẩu" min="6" max="12" onfocus="this.value = '';   value="Xác nhận mật khẩu"></div>
		    	 </div>
		    	  <div class="col_1_of_2 span_1_of_2">	
		    		<div><input name="txtEmail1" type="Email"  required="true" placeholder="Email" multiple autocomplete="off" id="txtEmail1"  onfocus="this.value = ''; value="E-mail"></div>
		    		<div><input name="txtDiachi1"  type="Text" required="true" id="txtDiachi1" placeholder="Địa chỉ" onfocus="this.value = ''; value="Địa chỉ">
		            </div>		        
		          <div><input name="txtNgaysinh1" type="text" id="txtNgaysinh1" required="True" placeholder="Năm/Tháng/Ngày Sinh" onfocus="this.value = ''; value="Ngày Sinh"></div>
		           <div>
		          </div>
		          	<input name="txtSDT1" type="text" required="true" placeholder="SĐT" id="txtSDT1" onfocus="this.value = ''; value="SĐT">
		    	  </div>
		      <button class="grey" onclick="Validate()">Dang ky</button>
		    <p class="terms">&nbsp;</p>
		    <div class="clear"></div>
            <input type="hidden" name="MM_insert" value="form1" />
               </form>
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
		</div>
</body>
</html>
