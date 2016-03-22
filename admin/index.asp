<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="dangnhap.asp"
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
Dim thongtintaikhoan
Dim thongtintaikhoan_cmd
Dim thongtintaikhoan_numRows

Set thongtintaikhoan_cmd = Server.CreateObject ("ADODB.Command")
thongtintaikhoan_cmd.ActiveConnection = MM_connect_STRING
thongtintaikhoan_cmd.CommandText = "SELECT * FROM dbo.KhachHang ORDER BY MaKH DESC" 
thongtintaikhoan_cmd.Prepared = true

Set thongtintaikhoan = thongtintaikhoan_cmd.Execute
thongtintaikhoan_numRows = 0
%>
<%
Dim SoLuongKH
Dim SoLuongKH_cmd
Dim SoLuongKH_numRows

Set SoLuongKH_cmd = Server.CreateObject ("ADODB.Command")
SoLuongKH_cmd.ActiveConnection = MM_connect_STRING
SoLuongKH_cmd.CommandText = "SELECT COUNT( MaKH) FROM dbo.KhachHang " 
SoLuongKH_cmd.Prepared = true

Set SoLuongKH = SoLuongKH_cmd.Execute
SoLuongKH_numRows = 0
%>
<%
Dim SoLuongSanPham
Dim SoLuongSanPham_cmd
Dim SoLuongSanPham_numRows

Set SoLuongSanPham_cmd = Server.CreateObject ("ADODB.Command")
SoLuongSanPham_cmd.ActiveConnection = MM_connect_STRING
SoLuongSanPham_cmd.CommandText = "SELECT COUNT(MaSach) FROM dbo.Sach" 
SoLuongSanPham_cmd.Prepared = true

Set SoLuongSanPham = SoLuongSanPham_cmd.Execute
SoLuongSanPham_numRows = 0
%>
<%
Dim SoLuongAmdin
Dim SoLuongAmdin_cmd
Dim SoLuongAmdin_numRows

Set SoLuongAmdin_cmd = Server.CreateObject ("ADODB.Command")
SoLuongAmdin_cmd.ActiveConnection = MM_connect_STRING
SoLuongAmdin_cmd.CommandText = "SELECT COUNT(MaAdmin) FROM dbo.Admin" 
SoLuongAmdin_cmd.Prepared = true

Set SoLuongAmdin = SoLuongAmdin_cmd.Execute
SoLuongAmdin_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
thongtintaikhoan_numRows = thongtintaikhoan_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Creative - Bootstrap 3 Responsive Admin Template">
    <meta name="author" content="GeeksLabs">
    <meta name="keyword" content="Creative, Dashboard, Admin, Template, Theme, Bootstrap, Responsive, Retina, Minimal">
    <link rel="shortcut icon" href="img/favicon.png">

    <title>Admin</title>

    <!-- Bootstrap CSS -->    
    <link href="../css/bootstrap.min.css" rel="stylesheet">
    <!-- bootstrap theme -->
    <link href="../css/bootstrap-theme.css" rel="stylesheet">
    <!--external css-->
    <!-- font icon -->
    <link href="../css/stylecuatoi.css" rel="stylesheet" type="text/css" />
    <link href="../css/elegant-icons-style.css" rel="stylesheet" />
    <link href="../css/font-awesome.min.css" rel="stylesheet" />    
    <!-- full calendar css-->
    <link href="../assets/fullcalendar/fullcalendar/bootstrap-fullcalendar.css" rel="stylesheet" />
	<link href="../assets/fullcalendar/fullcalendar/fullcalendar.css" rel="stylesheet" />
    <!-- easy pie chart-->
    <link href="../assets/jquery-easy-pie-chart/jquery.easy-pie-chart.css" rel="stylesheet" type="text/css" media="screen"/>
    <!-- owl carousel -->
    <link rel="stylesheet" href="../css/owl.carousel.css" type="text/css">
	<link href="../css/jquery-jvectormap-1.2.2.css" rel="stylesheet">
    <!-- Custom styles -->
	<link rel="stylesheet" href="../css/fullcalendar.css">
	<link href="../css/widgets.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="../css/style-responsive.css" rel="stylesheet" />
	<link href="../css/xcharts.min.css" rel=" stylesheet">	
	<link href="../css/jquery-ui-1.10.4.min.css" rel="stylesheet">
    <!-- HTML5 shim and Respond.js IE8 support of HTML5 -->
    <!--[if lt IE 9]>
      <script src="../js/html5shiv.js"></script>
      <script src="../js/respond.min.js"></script>
      <script src="../js/lte-ie7.js"></script>
    <![endif]-->
  </head>

<body>
  <!-- container section start -->
  <section id="container" class="">
     
      
    <header class="header dark-bg">
            <div class="toggle-nav">
                <div class="icon-reorder tooltips" data-original-title="Toggle Navigation" data-placement="bottom"><i class="icon_menu"></i></div>
            </div>

            <!--logo start-->
            <a href="index.asp" class="logo"><span class="lite">Admin</span></a>
            <!--logo end-->

            <div class="nav search-row" id="top_menu">
                <!--  search form start -->
                <ul class="nav top-menu">                    
                    <li>
                        <!--<form class="navbar-form">
                            <input class="form-control" placeholder="Search" type="text">
                        </form>-->
                    </li>                    
                </ul>
                <!--  search form end -->                
            </div>

            <div class="top-nav notification-row">                
                <!-- notificatoin dropdown start-->
                <ul class="nav pull-right top-menu">
                    
                    <!-- task notificatoin start -->
                    <li id="task_notificatoin_bar" class="dropdown">
                       
                        <ul class="dropdown-menu extended tasks-bar">
                            <div class="notify-arrow notify-arrow-blue"></div>
                            <li>
                              
                            </li>
                            <li>
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                    <div class="task-info">
                                        
                                    </div>
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                    
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                    <div class="task-info">
                                      
                                    </div>
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                  
                                    <div class="progress progress-striped active">
                                      
                                    </div>

                                </a>
                            </li>
                            
                        </ul>
                    </li>
                    <!-- task notificatoin end -->
                    <!-- inbox notificatoin start-->
                    <li id="mail_notificatoin_bar" class="dropdown">
                       
                        <ul class="dropdown-menu extended inbox">
                            <div class="notify-arrow notify-arrow-blue"></div>
                            
                            <li>
                                
                               					
                            </li>
                            <li>
                               
                            </li>
                            <li>
                                
                            </li>
                            <li>
                                
                            </li>
                            <li>
                               
                            </li>
                        </ul>
                    </li>
                    <!-- inbox notificatoin end -->
                    <!-- alert notification start-->
                    <li id="alert_notificatoin_bar" class="dropdown">
                        
                        
                    </li>
                    <!-- alert notification end-->
                    <!-- user login dropdown start-->
                    <li class="dropdown">
                        
                            
                            <span class="username"><a href="../admin/dangnhap.asp">Đăng xuất</a></span>
                        </a>
                        
                    </li>
                    <!-- user login dropdown end -->
                </ul>
                <!-- notificatoin dropdown end-->
            </div>
      </header>      
      <!--header end-->

      <!--sidebar start-->
      <aside>
        <div id="sidebar"  class="nav-collapse ">
              <!-- sidebar menu start-->
              <ul class="sidebar-menu">                
                  <li class="active">
                      <a class="" href="index.asp">
                          <i class="icon_house_alt"></i>
                          <span>Trang chủ</span></a>
                  </li>
				  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_document_alt"></i>
                          <span>Sản Phẩm</span>
                          <span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">
                      	  <li><a class="" href="XemSach.asp">Sách</a></li> 
                          <li><a class="" href="ThemSach.asp">Thêm Sách</a></li>                          
                          <li><a class="" href="ThemTL.asp">Thêm Thể Loại</a></li>
                          <li><a class="" href="ThemTacGia.asp">Thêm Thể Loại</a></li>
                      </ul>
                  </li>       
                  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_desktop"></i>
                          <span>Tài khoản</span>
                          <span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">
                      	  <li><a class="" href="TaiKhoan.asp">Khách hàng</a></li>
                          <li><a class="" href="TaiKhoanAdmin.asp">Admin</a></li> 
                      </ul>
                      
                  </li>
                  <li>
                      <a class="" href="widgets.html">
                          <i class="icon_genius"></i>
                          <span>Widgets</span>
                      </a>
                  </li>
                  <li>                     
                      <a class="" href="chart-chartjs.html">
                          <i class="icon_piechart"></i>
                          <span>Charts</span>
                          
                      </a>
                                         
                  </li>
                             
                  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_table"></i>
                          <span>Tables</span>
                          <span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">
                          <li><a class="" href="basic_table.html">Basic Table</a></li>
                      </ul>
                  </li>
                  
                  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_documents_alt"></i>
                          <span>Pages</span>
                          <span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">                          
                          <li><a class="" href="profile.html">Profile</a></li>
                          <li><a class="" href="login.html"><span>Login Page</span></a></li>
                          <li><a class="" href="blank.html">Blank Page</a></li>
                          <li><a class="" href="404.html">404 Error</a></li>
                      </ul>
                  </li>
                  
              </ul>
              <!-- sidebar menu end-->
          </div>
      </aside>
      <!--sidebar end-->
      
      <!--main content start-->
      <section id="main-content">
          <section class="wrapper">            
              <!--overview start-->
			  <div class="row">
				<div class="col-lg-12">
					<h3 class="page-header"><i class="fa fa-laptop"></i> Dashboard</h3>
				  <ol class="breadcrumb">
						<li><i class="fa fa-home"></i><a href="index.html">Home</a></li>
						<li><i class="fa fa-laptop"></i>Dashboard</li>						  	
					</ol>
				</div>
			</div>
              
            <div class="row">
				<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
					<div class="info-box blue-bg">
						<i class="fa fa-cloud-download"></i>
						<div class="count">6.674</div>
						<div class="title">Download</div>						
					</div><!--/.info-box-->			
				</div><!--/.col-->
				
				<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
					<div class="info-box brown-bg">
						<i class="fa fa-shopping-cart"></i>
						<div class="count">7.538</div>
						<div class="title">Purchased</div>						
					</div><!--/.info-box-->			
				</div><!--/.col-->	
				
				<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
					<div class="info-box dark-bg">
						<i class="fa fa-thumbs-o-up"></i>
						<div class="count">4.362</div>
						<div class="title">Order</div>						
					</div><!--/.info-box-->			
				</div><!--/.col-->
				
				<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
					<div class="info-box green-bg">
						<i class="fa fa-cubes"></i>
						<div class="count">1.426</div>
						<div class="title">Stock</div>						
					</div><!--/.info-box-->			
				</div><!--/.col-->
				
			</div><!--/.row-->
		
			
           <div class="row">
		    <div class="col-lg-9 col-md-12">
					
					<div class="panel panel-default">
					  <div class="panel-heading">
							<h2><i class="fa fa-map-marker red"></i><strong>Bản đồ</strong></h2>
								
						</div>
						<div class="panel-body-map">
							<div id="map" style="height:380px;"></div>	
						</div>
	
					</div>
			 </div>
              <div class="col-md-3">
              <!-- List starts -->
				<ul class="today-datas">
                <!-- List #1 -->
				<li>
                  <!-- Graph -->
                  <div><span id="todayspark1" class="spark"></span></div>
                  <!-- Text -->
                  <div class="datas-text"><%=(SoLuongKH.Fields.Item("").Value)%> Khách Hàng   </div>
                </li>
                <li>
                  <div><span id="todayspark2" class="spark"></span></div>
                  <div class="datas-text"><%=(SoLuongSanPham.Fields.Item("").Value)%> Sản Phẩm</div>
                </li>
                <li>
                  <div><span id="todayspark3" class="spark"></span></div>
                  <div class="datas-text"><%=(SoLuongAmdin.Fields.Item("").Value)%> Quản Trị</div>
                </li>
                <li>
                  <div><span id="todayspark4" class="spark"></span></div>
                  <div class="datas-text">$16,00 Revenue/Day</div>
                </li> 
                <li>
                  <div><span id="todayspark5" class="spark"></span></div>
                  <div class="datas-text">12,000000 visitors every Month</div>
                </li>                                                                                                              
              </ul>
              </div>
              
			 
           </div>  
            
		  
		  <!-- Today status end -->
			
              
				
			<div class="row">
               	
				<div class="col-lg-9 col-md-12">	
					<div class="panel panel-default">
					  <div class="panel-heading">
							<h2><i class="fa fa-flag-o red"></i><strong>Thành Viên</strong></h2>
							
						</div>
						<div class="panel-body">
							<table class="table bootstrap-datatable countries">
								<thead>
									<tr>
										<th>Tên Tài khoản</th>
										<th>Tên thành viên</th>
										<th>Số điện thoại</th>
										<th>Địa chỉ</th>
									</tr>
								</thead>   
								<tbody>
                                  <% 
While ((Repeat1__numRows <> 0) AND (NOT thongtintaikhoan.EOF)) 
%>
  <tr>
    <td><%=(thongtintaikhoan.Fields.Item("TaiKhoan").Value)%></td>
    <td><%=(thongtintaikhoan.Fields.Item("TenKH").Value)%></td>
    <td><%=(thongtintaikhoan.Fields.Item("SDT").Value)%></td>
    <td><%=(thongtintaikhoan.Fields.Item("DiaChi").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  thongtintaikhoan.MoveNext()
Wend
%>
                                </tbody>
							</table>
						</div>
	
					</div>	

				</div><!--/col-->
				<div class="col-md-3">
					
					<div class="social-box facebook">
						<i class="fa fa-facebook"></i>
						<ul>
							<li>
								<strong>256k</strong>
								<span>friends</span>
							</li>
							<li>
								<strong>359</strong>
								<span>feeds</span>
							</li>
						</ul>
					</div><!--/social-box-->
				</div>
				<div class="col-md-3">
					
					<div class="social-box google-plus">
						<i class="fa fa-google-plus"></i>
						<ul>
							<li>
								<strong>962</strong>
								<span>followers</span>
							</li>
							<li>
								<strong>256</strong>
								<span>circles</span>
							</li>
						</ul>
					</div><!--/social-box-->			
					
				</div><!--/col--><!--/col-->
				
            </div>

                    
                   
                <!-- statics end -->
              
            
				

              <!-- project team & activity start -->
          <div class="row">
            <div class="col-lg-8">
                <!--Project Activity start-->
                <section class="panel">
                          <div class="panel-body progress-panel">
                            <div class="row">
                              <div class="col-lg-8 task-progress pull-left">
                                  <h1>Thống kê</h1>                                  
                              </div>
                              <div class="col-lg-4"></div>
                            </div>
                          </div>
                          <table class="table table-hover personal-task">
                              <tbody>
                              <tr>
                                  <td>Danh Mục</td>
                                  <td>Số Lượng</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>
                              <tr>
                                  <td>Sản Phẩm</td>
                                  <td><%=(SoLuongSanPham.Fields.Item("").Value)%></td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>
                              <tr>
                                  <td>Khách Hàng</td>
                                  <td><%=(SoLuongKH.Fields.Item("").Value)%></td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>                              
                              <tr>
                                  <td>Admin</td>
                                  <td><%=(SoLuongAmdin.Fields.Item("").Value)%></td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>
                              <tr>
                                  <td>Đơn Hàng</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>                              
                              <tr>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>
							  <tr>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>
							  <tr>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                              </tr>
                              </tbody>
                          </table>
                      </section>
                      <!--Project Activity end-->
                  </div>
            </div><br><br>
		
		<div class="row">
			<div class="col-md-6 portlets">
            <div class="panel panel-default">
			  <div class="panel-heading">
                  <h2><strong>Lịch</strong></h2>
				
                 
                </div><br><br><br>
                <div class="panel-body">
                  <!-- Widget content -->
                  
                    <!-- Below line produces calendar. I am using FullCalendar plugin. -->
                    <div id="calendar"></div>
                  
                </div>
              </div> 
               
            </div>
			
				 <div class="col-md-6 portlets">
              <div class="panel panel-default">
                <div class="panel-heading">
                  <div class="pull-left">Quick Post</div>
                  <div class="widget-icons pull-right">
                    <a href="#" class="wminimize"><i class="fa fa-chevron-up"></i></a> 
                    <a href="#" class="wclose"><i class="fa fa-times"></i></a>
                  </div>  
                  <div class="clearfix"></div>
                </div>
                <div class="panel-body">
                  <div class="padd">
                    
                      <div class="form quick-post">
                                      <!-- Edit profile form (not working)-->
                                      <form class="form-horizontal">
                                          <!-- Title -->
                                          <div class="form-group">
                                            <label class="control-label col-lg-2" for="title">Title</label>
                                            <div class="col-lg-10"> 
                                              <input type="text" class="form-control" id="title">
                                            </div>
                                          </div>   
                                          <!-- Content -->
                                          <div class="form-group">
                                            <label class="control-label col-lg-2" for="content">Content</label>
                                            <div class="col-lg-10">
                                              <textarea class="form-control" id="content"></textarea>
                                            </div>
                                          </div>                           
                                          <!-- Cateogry -->
                                          <div class="form-group">
                                            <label class="control-label col-lg-2">Category</label>
                                            <div class="col-lg-10">                               
                                                <select class="form-control">
                                                  <option value="">- Choose Cateogry -</option>
                                                  <option value="1">General</option>
                                                  <option value="2">News</option>
                                                  <option value="3">Media</option>
                                                  <option value="4">Funny</option>
                                                </select>  
                                            </div>
                                          </div>              
                                          <!-- Tags -->
                                          <div class="form-group">
                                            <label class="control-label col-lg-2" for="tags">Tags</label>
                                            <div class="col-lg-10">
                                              <input type="text" class="form-control" id="tags">
                                            </div>
                                          </div>
                                          
                                          <!-- Buttons -->
                                          <div class="form-group">
                                             <!-- Buttons -->
											 <div class="col-lg-offset-2 col-lg-9">
												<button type="submit" class="btn btn-primary">Publish</button>
												<button type="submit" class="btn btn-danger">Save Draft</button>
												<button type="reset" class="btn btn-default">Reset</button>
											 </div>
                                          </div>
                                      </form>
                    </div>
                  

                  </div>
                  <div class="widget-foot">
                    <!-- Footer goes here -->
                  </div>
                </div>
              </div>
              
            </div>
                        
          </div> 
              <!-- project team & activity end -->

          </section>
      </section>
      <!--main content end-->
  </section>
  <!-- container section start -->

    <!-- javascripts -->
<script src="../js/jquery.js"></script>
<script src="../js/jquery-ui-1.10.4.min.js"></script>
<script src="../js/jquery-1.8.3.min.js"></script>
<script type="text/javascript" src="../js/jquery-ui-1.9.2.custom.min.js"></script>
    <!-- bootstrap -->
<script src="../js/bootstrap.min.js"></script>
    <!-- nice scroll -->
<script src="../js/jquery.scrollTo.min.js"></script>
<script src="../js/jquery.nicescroll.js" type="text/javascript"></script>
    <!-- charts scripts -->
<script src="../assets/jquery-knob/js/jquery.knob.js"></script>
<script src="../js/jquery.sparkline.js" type="text/javascript"></script>
<script src="../assets/jquery-easy-pie-chart/jquery.easy-pie-chart.js"></script>
<script src="../js/owl.carousel.js" ></script>
    <!-- jQuery full calendar -->
    <<script src="../js/fullcalendar.min.js"></script> <!-- Full Google Calendar - Calendar -->
<script src="../assets/fullcalendar/fullcalendar/fullcalendar.js"></script>
    <!--script for this page only-->
<script src="../js/calendar-custom.js"></script>
<script src="../js/jquery.rateit.min.js"></script>
    <!-- custom select -->
<script src="../js/jquery.customSelect.min.js" ></script>
<script src="../assets/chart-master/Chart.js"></script>
   
    <!--custome script for all page-->
<script src="../js/scripts.js"></script>
    <!-- custom script for this page-->
<script src="../js/sparkline-chart.js"></script>
<script src="../js/easy-pie-chart.js"></script>
<script src="../js/jquery-jvectormap-1.2.2.min.js"></script>
<script src="../js/jquery-jvectormap-world-mill-en.js"></script>
<script src="../js/xcharts.min.js"></script>
<script src="../js/jquery.autosize.min.js"></script>
<script src="../js/jquery.placeholder.min.js"></script>
<script src="../js/gdp-data.js"></script>	
<script src="../js/morris.min.js"></script>
<script src="../js/sparklines.js"></script>	
<script src="../js/charts.js"></script>
<script src="../js/jquery.slimscroll.min.js"></script>
<script>

      //knob
      $(function() {
        $(".knob").knob({
          'draw' : function () { 
            $(this.i).val(this.cv + '%')
          }
        })
      });

      //carousel
      $(document).ready(function() {
          $("#owl-slider").owlCarousel({
              navigation : true,
              slideSpeed : 300,
              paginationSpeed : 400,
              singleItem : true

          });
      });

      //custom select box

      $(function(){
          $('select.styled').customSelect();
      });
	  
	  /* ---------- Map ---------- */
	$(function(){
	  $('#map').vectorMap({
	    map: 'world_mill_en',
	    series: {
	      regions: [{
	        values: gdpData,
	        scale: ['#000', '#000'],
	        normalizeFunction: 'polynomial'
	      }]
	    },
		backgroundColor: '#eef3f7',
	    onLabelShow: function(e, el, code){
	      el.html(el.html()+' (GDP - '+gdpData[code]+')');
	    }
	  });
	});



  </script>

</body>
</html>
<%
SoLuongKH.Close()
Set SoLuongKH = Nothing
%>
<%
SoLuongSanPham.Close()
Set SoLuongSanPham = Nothing
%>
<%
SoLuongAmdin.Close()
Set SoLuongAmdin = Nothing
%>
<%
thongtintaikhoan.Close()
Set thongtintaikhoan = Nothing
%>
