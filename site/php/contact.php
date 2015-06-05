<?php
if(!isset($_POST['submit']))
{
    //This page should not be accessed directly. Need to submit the form.
    echo "Error; you need to submit the form!";
}
$user_name = $_POST['username'];
$user_email = $_POST['useremail'];
$user_subject = $_POST['subject'];
$user_message = $_POST['message'];

if(IsInjected($user_email))
{
    echo "Bad email value!";
    exit;
}

$email_from = "$user_email";//<== update the email address
$email_subject = "DebateHelper Message: $user_name";
$email_body = "Subject: $user_subject \nMessage: $user_message \n\n\nTo: ".

$to = "zachkirsch@gmail.com";//<== update the email address
$headers = "From: $email_from \r\n";

//Send the email!
mail($to,$email_subject,$email_body,$headers);
//done. redirect to thank-you page.
header('Location: http://zachkirsch.com/debatehelper');

// Function to validate against any email injection attempts
function IsInjected($str)
{
  $injections = array('(\n+)',
              '(\r+)',
              '(\t+)',
              '(%0A+)',
              '(%0D+)',
              '(%08+)',
              '(%09+)'
              );
  $inject = join('|', $injections);
  $inject = "/$inject/i";
  if(preg_match($inject,$str))
    {
    return true;
  }
  else
    {
    return false;
  }
}

?>
