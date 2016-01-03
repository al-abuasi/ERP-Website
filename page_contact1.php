<!DOCTYPE html>
<!--[if IE 8]> <html lang="en" class="ie8"> <![endif]-->  
<!--[if IE 9]> <html lang="en" class="ie9"> <![endif]-->  
<!--[if !IE]><!--> <html lang="en"> <!--<![endif]-->  
<head>
    <title>ERP Safety | Contact US</title>

    <!-- Meta -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">

    <!-- Favicon -->
    <link rel="shortcut icon" href="favicon.ico">

    <!-- CSS Global Compulsory -->
    <link rel="stylesheet" href="assets/plugins/bootstrap/css/bootstrap.min.css">
    <link rel="stylesheet" href="assets/css/style.css">

    <!-- CSS Implementing Plugins -->
    <link rel="stylesheet" href="assets/plugins/line-icons/line-icons.css">
    <link rel="stylesheet" href="assets/plugins/font-awesome/css/font-awesome.min.css">
    <link rel="stylesheet" href="assets/plugins/flexslider/flexslider.css">     

    <!-- CSS Page Style -->    
    <link rel="stylesheet" href="assets/css/pages/page_contact.css">

    <!-- CSS Theme -->    
    <link rel="stylesheet" href="assets/css/themes/red.css" id="style_color">

    <!-- CSS Customization -->
    <link rel="stylesheet" href="assets/css/custom.css">
</head> 

<body>
    

<?php include ("navigation.php"); ?> 



    <!--=== Breadcrumbs ===-->
    <div class="breadcrumbs"
>
        <div class="container">
            <h1 class="pull-left">Contact Us</h1>
            <ul class="pull-right breadcrumb">
                <li><a href="index.php">Home</a></li>
               
                <li class="active">Contacts</li>
            </ul>
        </div>
    </div><!--/breadcrumbs-->
    <!--=== End Breadcrumbs ===-->

    <!-- Google Map -->
    
  <div id="map" class="map map-box map-box-space margin-bottom-20">
    </div>
 
<!---/map-->
    <!-- End Google Map -->

    <!--=== Content Part ===-->
    <div class="container content">		
    	<div class="row margin-bottom-30">
    		<div class="col-md-9 mb-margin-bottom-30">
                <div class="headline" id="head11"><h2>Contact Us</h2></div>
                <p>Please fill out this short form and we will get back to you as soon as possible with a reply. Thank you</p><br />
				<?php if(isset($_GET['failure']))
				{
				?>
				<div class="alert alert-danger">
				<strong>Unable to Submit Contact Form. Please Fill all the required fields and make sure you have a valid email address.</strong>
				</div>
				<?php
				}
				
				if(isset($_GET['success']))
				{
				?>
				<div class="alert alert-success">
				Your Message Has been Submitted Successfully. We will get back to you shortly.
				</div>
				<?php
				}
				else {
				?>
    			<form action="process.php" method="POST">
                    <label>Name <span class="color-red">*</span></label>
                    <div class="row margin-bottom-20">
                        <div class="col-md-7 col-md-offset-0">
                            <input type="text" name="name" class="form-control" required="">
                        </div>                
                    </div>
                    
                    <label>Email <span class="color-red">*</span></label>
                    <div class="row margin-bottom-20">
                        <div class="col-md-7 col-md-offset-0">
                            <input type="text" name="email" class="form-control" required="">
                        </div>                
                    </div>
                    
                     <label>Phone Number </label>
                    <div class="row margin-bottom-20">
                        <div class="col-md-7 col-md-offset-0">
                            <input type="text" name="phone" class="form-control">
                        </div>                
                    </div>
                    
                    
                    <label>Message</label>
                    <div class="row margin-bottom-20">
                        <div class="col-md-11 col-md-offset-0">
                            <textarea name="message" rows="8" class="form-control" required=""></textarea>
                        </div>                
                    </div>
                    
                    <p><button type="submit" class="btn-u">Send Message</button></p>
                </form>
				<?php
				}
				?>
            </div><!--/col-md-9-->
            
    		<div class="col-md-3">
            	<!-- Contacts -->
                <div class="headline"><h2>Our Information</h2></div>
                <ul class="list-unstyled who margin-bottom-30">
                    <li><a href="#"><i class="fa fa-home"></i>101 W. Ayre Street ,Newport, DE 19804</a></li>
                    <li><a href="#"><i class="fa fa-envelope"></i>mail@erpsafety.com</a></li>
                    <li><a href="#"><i class="fa fa-phone"></i>1(302)994-2600</a></li>
                    <li><a href="#"><i class="fa fa-globe"></i>http://www.erpsafety.com</a></li>
                </ul>

            	<!-- Business Hours -->
                <div class="headline"><h2>Business Hours</h2></div>
                <ul class="list-unstyled margin-bottom-30">
                	<li><strong>Monday-Friday:</strong> 8:30am to 5:30pm</li>
                	<li><strong>Saturday:</strong> Closed</li>
                	<li><strong>Sunday:</strong> Closed</li>
                </ul>

      
    </div><!--/container-->		
    <!--=== End Content Part ===-->
    </div>
    </div>
    <?php include ("footer.php"); ?> 



<script type="text/javascript" src="assets/plugins/jquery-1.10.2.min.js"></script>
<script type="text/javascript" src="assets/plugins/jquery-migrate-1.2.1.min.js"></script>
<script type="text/javascript" src="assets/plugins/bootstrap/js/bootstrap.min.js"></script> 
<!-- JS Implementing Plugins -->           
<script type="text/javascript" src="assets/plugins/back-to-top.js"></script>
<script type="text/javascript" src="assets/plugins/flexslider/jquery.flexslider-min.js"></script>
<script type="text/javascript" src="http://maps.google.com/maps/api/js?sensor=true"></script>
<script type="text/javascript" src="assets/plugins/gmap/gmap.js"></script>
<!-- JS Page Level -->           
<script type="text/javascript" src="assets/js/app.js"></script>
<script type="text/javascript" src="assets/js/pages/page_contacts.js"></script>

<script type="text/javascript">
    jQuery(document).ready(function() {
        App.init();
        App.initSliders(); 
ContactPage.initMap();  		
    });
</script>
<!--[if lt IE 9]>
    <script src="assets/plugins/respond.js"></script>
    <script src="assets/plugins/html5shiv.js"></script>
<![endif]-->

</body>
</html> 

<!-- DESIGNED AND DEVELOPED BY ALMUTHANNA ABUASI -->