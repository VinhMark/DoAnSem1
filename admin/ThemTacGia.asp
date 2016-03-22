<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1"
MM_authFailedURL="XemSach.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connect_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.TacGia (TenTG) VALUES (?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("txtTacGia")) ' adVarWChar
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
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/template-admin.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Creative - Bootstrap 3 Responsive Admin Template">
    <meta name="author" content="GeeksLabs">
    <meta name="keyword" content="Creative, Dashboard, Admin, Template, Theme, Bootstrap, Responsive, Retina, Minimal">
    <link rel="shortcut icon" href="img/favicon.png">

    <title>Admin LeoShop</title>

    <!-- Bootstrap CSS -->  
    <link href="../css/stylecuatoi.css" rel="stylesheet" type="text/css" />  
    <link href="../css/bootstrap.min.css" rel="stylesheet">
    <!-- bootstrap theme -->
    <link href="../css/bootstrap-theme.css" rel="stylesheet">
    <!--external css-->
    <!-- font icon -->
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
                        
                            
                            <span class="username"><a href="dangnhap.asp">Đăng xuất</a></span>
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
			  <div class="row"><!-- InstanceBeginEditable name="nội dung" -->
			    <table class="theten" width="98%" height="30px"><td>Thêm Tác Giả</td></table>
			    <div class="noidung">
			    <form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
			      <table class="thethongtin" width="98%" border="0" cellspacing="0" cellpadding="0">
			        <tr>
			          <td>&nbsp;</td>
			          <td>Thêm Tác Giả</td>
			          <td><label for="txtTacGia"></label>
		              <input type="text" name="txtTacGia" id="txtTacGia" /></td>
		            </tr>
			        <tr>
			          <td>&nbsp;</td>
			          <td><input type="submit" name="Submit" id="button" value="Thêm" class="btn"/></td>
			          <td><input type="reset" name="button2" id="button2" value="Hủy" class="btn"/></td>
		            </tr>
		          </table>
                  <input type="hidden" name="MM_insert" value="form1" />
                </form>
              </div>
                <!--------------------------end noidung--------------------------->
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			  <!-- InstanceEndEditable --></div>
              
            <div class="row"><!--/.col--><!--/.col--><!--/.col--><!--/.col-->
				
			</div><!--/.row-->
		
			
           <div class="row"></div>  
            
		  
		  <!-- Today status end -->
			
              
				
			<div class="row"><!--/col--><!--/col--></div>

                    
                   
                <!-- statics end -->
              
            
				

              <!-- project team & activity start -->
          <div class="row"></div>
          <p>&nbsp;	</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p><br>
            <br>
            
          </p>
          <div class="row"></div> 
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
<!-- InstanceEnd --></html>
