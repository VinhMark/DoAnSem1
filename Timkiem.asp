<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
Dim timkiem__MMColParam
timkiem__MMColParam = "1"
If (Request.Form("txtTimkiem") <> "") Then 
  timkiem__MMColParam = Request.Form("txtTimkiem")
End If
%>
<%
Dim timkiem
Dim timkiem_cmd
Dim timkiem_numRows

Set timkiem_cmd = Server.CreateObject ("ADODB.Command")
timkiem_cmd.ActiveConnection = MM_connect_STRING
timkiem_cmd.CommandText = "SELECT * FROM dbo.Sach WHERE TenSach like ? and TinhTrang=1 and Hienthi=1" 
timkiem_cmd.Prepared = true
timkiem_cmd.Parameters.Append timkiem_cmd.CreateParameter("param1", 200, 1, 255, "%" + timkiem__MMColParam + "%") ' adVarChar

Set timkiem = timkiem_cmd.Execute
timkiem_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 9
Repeat1__index = 0
timkiem_numRows = timkiem_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Free Leoshop Website Template | Home :: w3layouts</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href="css/form.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>
<script type="text/javascript" src="js/jquery1.min.js"></script>
<!-- start menu -->
<link href="css/mystyle.css" rel="stylesheet" type="text/css">
<link href="css/megamenu.css" rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" src="js/megamenu.js"></script>
<script>$(document).ready(function(){$(".megamenu").megamenu();});</script>
<!--start slider -->
    <link rel="stylesheet" href="css/fwslider.css" media="all">
<script src="js/jquery-ui.min.js"></script>
<script src="js/css3-mediaqueries.js"></script>
<script src="js/fwslider.js"></script>
<!--end slider -->
<script src="js/jquery.easydropdown.js"></script>
</head>

<body>
     <div class="header-top">
	   <div class="wrap"> 
			  
			 <div class="cssmenu">
				<ul>
					<li class="active"></li> 
					<li><a href="DangXuat user.asp">Đăng xuất</a></li>
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
			<li class="active grid"><a href="index.asp">TRANG CHỦ</a></li>
			<li><a class="color4" href="#">THỂ LOẠI</a>
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
		  <div class="clear"></div>
     </div>
	</div>
  <!-- start slider --><!--/slider -->
<div class="main">
	<div class="wrap">
		<div class="section group">
		  <div class="cont span_2_of_3">
		  	<h2 class="head">Tìm kiếm</h2>
			<div class="top-box"><!-- TemplateBeginEditable name="EditRegion1" -->
            <% 
While ((Repeat1__numRows <> 0) AND (NOT timkiem.EOF)) 
%>
  <div class="item">
    
    <div class="hinh hinhanh"><img src="<%=(timkiem.Fields.Item("HinhAnh").Value)%>" alt="" name="" width="226" height="283" />
    	<form action="Thongtin.asp" method="post">
        	<div class="thean">
            <input type="image" name="imageField" id="imageField" src="images/xemthem2.png" />
        	<input name="MaSP" type="hidden" id="MaSP" value="<%=(timkiem.Fields.Item("MaSach").Value)%>" />
            </div>
        </form>
    </div>
    
    
    
    <div class="thongtin">
      <div class="the-trai">
        <p id="the-trai"><%=(timkiem.Fields.Item("TenSach").Value)%></p>
        <div class="giasach">Giá :<%=(timkiem.Fields.Item("Gia").Value)%> VNĐ</div>
        </div>
      <form action="Giohang.asp" method="post">  
      <div class="the-phai">
        <input name="MaSP" type="hidden" id="MaSP" value="<%=(timkiem.Fields.Item("MaSach").Value)%>" />
        <input name="TenSP" type="hidden" id="TenSP" value="<%=(timkiem.Fields.Item("TenSach").Value)%>" />
        <input name="GiaSP" type="hidden" id="GiaSP" value="<%=(timkiem.Fields.Item("Gia").Value)%>" />
        <input name="HinhAnhSP" type="hidden" id="HinhAnhSP" value="<%=(timkiem.Fields.Item("HinhAnh").Value)%>" />
        <input type="image" name="imageField" id="imageField" src="images/cart.png" />
      
      </div>
      </form>
      <div class="tay"></div>
      </div>
    
  </div>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  timkiem.MoveNext()
Wend
%>
            <!-- TemplateEndEditable -->
		       <div class="clear"></div>
			</div>
			<h2 class="head">&nbsp;</h2>
		  </div>
		  <div class="clear"></div>
	</div>
	</div>
	</div>
   <div class="footer">
		<div class="footer-top">
			<div class="wrap">
			  <div class="section group example">
				<div class="col_1_of_2 span_1_of_2">

					<ul class="f-list">
					  <li><img src="images/2.png"><span class="f-text">Miễn Phí Giao Hàng</span><div class="clear"></div></li>
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
			    
		      </div>
	     </div>
	</div>
</body>
</html>
<%
timkiem.Close()
Set timkiem = Nothing
%>
