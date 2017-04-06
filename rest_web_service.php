<? 
$ServiceUrl="https://aim-mcinstance-new-02.qwasi.com/REST/rest21.php";
$MCUsername="user123";
$MCPassword="Qwasi123";
$MCCustomerId="140";
$mode=$_POST['submit']; 
$curl = curl_init($ServiceUrl);

if ($mode=='Register'){
$phonenumber=$_POST['phonenumberreg']; 
$myusername=$_POST['myusername']; 
$email=$_POST['email']; 
$last_name=$_POST['lastname'];
$curl_post_data = array(
        "username" => $MCUsername,
        "password" => $MCPassword,
     "customer_id" => $MCCustomerId,
      #"level"      => '0',
       "operation" => 'Put',
        "method"   => 'member.create',
         "msisdn"  => $phonenumber,
         "first_name"=> $myusername,
         "last_name" => $last_name,
         "optin_status" => '1',
          "email" => $email
        );
}
else{
$phonenumber=$_POST['phonenumber']; 
   $curl_post_data = array(
        "username" => $MCUsername,
        "password" => $MCPassword,
        "customer_id" => $MCCustomerId,
        "level"      => '0',
        "operation" =>'Put',
        "method"   =>'misc.sendsms',
         "mobile"  =>$phonenumber,
         "message" =>'Hello World'
        );
}

   curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
   curl_setopt($curl, CURLOPT_POST, true);
   curl_setopt($curl, CURLOPT_POSTFIELDS, $curl_post_data);
   $curl_response = curl_exec($curl);

echo $curl_response;#TO PRINT THE RESPONSE
if ($mode!='Register')
echo" You will Reecive Message shortly";# TO SEND A CUSTOMISE MESSAGE TO USER.


curl_close($curl);

#username={your username}&password={your password}&customer_id={your customer id}&level=0&operation=Put&method=misc.sendsms&mobile={your phone number}&message=Hello+World;
#username={your username}&password={your password}&customer_id={your customer id}&level=0&operation=Put&method=misc.sendsms&mobile={your phone number}&message=Hello+World
#operation=Put&method=misc.sendsms&mobile={your phone number}&message=Hello+World
#curl -u "{your username}:{your password}" -H "qid:{your customer id}" -d "operation=Put&method=member.create&msisdn=4085551212&first_name=John&last_name=Smith&optin_status=1" 



?>
