<?php
error_reporting(E_ALL);
$to="al@erpsafety.com"; //<========Replace the email in quotes with the email address where you want contact emails to be sent to



if(isset($_POST['message']) and $_POST['message']!='' and strlen($_POST['message'])>10 and $_POST['message']!=null and filter_var($_POST['email'] , FILTER_VALIDATE_EMAIL))
{

$from=$_POST['email'];
$message=$_POST['message'];
$phone=" ";
$name=" ";
$name=((isset($_POST['name']) and $_POST['name']!='') ? $_POST['name'] : "Not Provided");
$phone=((isset($_POST['phone']) and $_POST['phone']!='') ? $_POST['phone'] : "Not Provided");
$subject = 'Contact Form Submission';
$email="<html>
<head>
  <title>Contact Form Submission</title>
</head>
<body>";
$email.="Hello admin, <br/> Someone used the contact us form at ERP website's contact page. <br/><br/>Here are the details: <br/><p>Message: <br/>$message </p><br/>Email : <strong>$from</strong><br/>Full Name : <strong>$name</strong><br/>Phone : <strong>$phone</strong><br/>";
$email.="</body></html>";

// To send HTML mail, the Content-type header must be set
$headers  = 'MIME-Version: 1.0' . "\r\n";
$headers .= 'Content-type: text/html; charset=iso-8859-1' . "\r\n";

// Additional headers
$headers .= 'To: Admin <'.$to.'>' . "\r\n";
$headers .= 'From: ERP User <'.$from.'>' . "\r\n";
mail($to, $subject, $email, $headers);
header('location:page_contact1.php?success=1#head11');

}
else {
header('location:page_contact1.php?failure=1#head11');
}


?>

<!-- DESIGNED AND DEVELOPED BY ALMUTHANNA ABUASI -->