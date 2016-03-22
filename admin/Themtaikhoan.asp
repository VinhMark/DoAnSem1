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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connect_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.[Admin] (TaiKhoan, MatKhau) VALUES (?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, -1, Request.Form("txtTaikhoan")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, -1, Request.Form("txtmk")) ' adLongVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "dangnhap.asp"
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
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Creative - Bootstrap 3 Responsive Admin Template">
    <meta name="author" content="GeeksLabs">
    <meta name="keyword" content="Creative, Dashboard, Admin, Template, Theme, Bootstrap, Responsive, Retina, Minimal">
    <link rel="shortcut icon" href="../img/favicon.png">

    <title>Login Page 2 | Creative - Bootstrap 3 Responsive Admin Template</title>

    <!-- Bootstrap CSS -->    
    <link href="../css/bootstrap.min.css" rel="stylesheet">
    <!-- bootstrap theme -->
    <link href="../css/bootstrap-theme.css" rel="stylesheet">
    <!--external css-->
    <!-- font icon -->
    <link href="../css/elegant-icons-style.css" rel="stylesheet" />
    <link href="../css/font-awesome.css" rel="stylesheet" />
    <!-- Custom styles -->
    <link href="../css/style.css" rel="stylesheet">
    <link href="../css/style-responsive.css" rel="stylesheet" />

    <!-- HTML5 shim and Respond.js IE8 support of HTML5 -->
    <!--[if lt IE 9]>
    <script src="js/html5shiv.js"></script>
    <script src="js/respond.min.js"></script>
    <![endif]-->
</head>

<body class="login-img3-body">

    <div class="container">

      <form name="form1" METHOD="POST" class="login-form" action="<%=MM_editAction%>">        
        <div class="login-wrap">
            <p class="login-img"><i class="icon_lock_alt"></i></p>
          <div class="input-group">
              <span class="input-group-addon"><i class="icon_profile"></i></span>
              <input name="txtTaikhoan" type="text" autofocus class="form-control" id="txtTaikhoan" placeholder="Username">
            </div>
          <div class="input-group">
                <span class="input-group-addon"><i class="icon_key_alt"></i></span>
                <input name="txtmk" type="password" class="form-control" id="txtmk" placeholder="Password">
            </div>
            <div class="input-group"> <span class="input-group-addon"><i class="icon_key_alt"></i></span>
              <input name="txtmk2" type="password" class="form-control" id="txtmk2" placeholder="Password" />
            </div>
            <button class="btn btn-primary btn-lg btn-block" type="submit">Đăng ký</button>
        </div>
        <input type="hidden" name="MM_insert" value="form1" />
      </form>

    </div>


</body>
</html>
