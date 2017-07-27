<?php
// $Id: mrbs_auth.inc,v 1.8.2.1 2005/03/29 13:26:27 jberanek Exp $

// marco - user level check 

// include the authentification wrappers
include "auth_$auth[type].inc";
if (isset($auth['session'])) include "session_$auth[session].inc";

/* getAuthorised($user, $pass, $level)
 * 
 * Check to see if the user name/password is valid
 * 
 * $user  - The user name
 * $pass  - The users password
 * $level - The access level required
 * 
 * Returns:
 *   0        - The user does not have the required access
 *   non-zero - The user has the required access
 */
function getAuthorised($level)
{
    global $auth;

    $user = getUserName();
    if(isset($user) == FALSE) {
        authGet();
        return 0;
    }
    return authGetUserLevel($user, $auth["admin"]) >= $level;
}

function isLoggedIn()
{
    global $auth;

    $user = getUserName();
    if(isset($user) == FALSE) {
//        authGet();
        return 0;
    }
    return 1;
}

function isCourseAdmin()
{
    global $auth;

    $user = getUserName();
    if(isset($user) == FALSE) {
//        authGet();
        return 0;
    }
    return getUserLevel($user) == 1.5;
}

function isDeptHead()
{
    global $auth;

    $user = getUserName();
    if(isset($user) == FALSE) {
//        authGet();
        return 0;
    }
    return getUserLevel($user) == 1.7;
}

function isEB()
{
    global $auth;

    $user = getUserName();
    if(isset($user) == FALSE) {
//        authGet();
        return 0;
    }
    return getUserLevel($user) == 1.9;
}

function isAdmin()
{	
    global $auth;

    $user = getUserName();
    if(isset($user) == FALSE) {
//        authGet();
        return 0;
    }
    return getUserLevel($user) == 2;
}

/* getWritable($creator, $user)
 * 
 * Determines if a user is able to modify an entry
 *
 * $creator - The creator of the entry
 * $user    - Who wants to modify it
 *
 * Returns:
 *   0        - The user does not have the required access
 *   non-zero - The user has the required access
 */
function getWritable($creator, $user)
{
    global $auth;

    // Always allowed to modify your own stuff
    if(strcasecmp($creator, $user) == 0)
        return 1;

    if(authGetUserLevel($user, $auth["admin"]) >= 2)
        return 1;

    // Unathorised access
    return 0;
}

function getDeletable($creator, $user, $booktime )
{
    global $auth;
    $roleValid = false;
    $result = false;

    // Always allowed to modify your own stuff
    if ( (strcasecmp($creator, $user) == 0) || (authGetUserLevel($user, $auth["admin"]) >= 2) )
        $roleValid = true;
    else
        $roleValid = false;

    $start_day   = (int)strftime('%d', $booktime);
    $start_month = (int)strftime('%m', $booktime);
    $start_year  = (int)strftime('%Y', $booktime);
    $bookday = mktime(0,0,0,$start_month,$start_day,$start_year);

    $this_day   = date("d");
    $this_month = date("m");
    $this_year  = date("Y");
    $today = mktime(0,0,0,$this_month,$this_day,$this_year);

    $notifyday = ($bookday-$today)/60/60/24;

    if ( ($notifyday >= 28) && ($roleValid) ) //minimum notify time is 28 days
	$result = true;
    else
        $result = false;

//    return $result;
    return false;  // disabled the delete function

}


/* showAccessDenied()
 * 
 * Displays an appropate message when access has been denied
 * 
 * Retusns: Nothing
 */
function showAccessDenied($day, $month, $year, $area)
{
    global $HTTP_REFERER;

    print_header($day, $month, $year, $area);
?>
  <H1><?php echo get_vocab("accessdenied")?></H1>
  <P>
   <?php echo get_vocab("norights")?>
  </P>
  <P>
   <A HREF="<?php echo $HTTP_REFERER; ?>"><?php echo get_vocab("returnprev"); ?></A>
  </P>
 </BODY>
</HTML>
<?php
}

function showAccessDeniedForNotLogin($day, $month, $year, $area, $type)
{
    global $HTTP_REFERER;

    print_header($day, $month, $year, $area);
?>
  <H1><?php echo get_vocab("accessdenied")?></H1>
  <P>
   <?php 
       if ($type=="day") {
	   echo get_vocab("outofrange"); }
       else {
	   echo get_vocab("noweekmonth"); }   
   ?>
  </P>

<?php
$currentDay   = date("d");
$currentMonth = date("m");
$currentYear  = date("Y");
?>

  <P>
   <A HREF="<?php echo $url_base; ?>day.php?day=<?php echo $currentDay?>&month=<?php echo $currentMonth?>&year=<?php echo $currentYear?>&area=2&room=3"><?php echo get_vocab("returnmain"); ?></A>
  </P>

  <P><HR><B><?php echo get_vocab("viewday") ?></B>

<?php

for($i = 0; $i <= 13; $i++)
{
	$ctime = mktime(0, 0, 0, $currentMonth, $currentDay + $i, $currentYear);

	$str = utf8_strftime(empty($dateformat)? "%b %d" : "%d %b", $ctime);

	$cyear  = date("Y", $ctime);
	$cmonth = date("m", $ctime);
	$cday   = date("d", $ctime);
	if ($i != -6) echo " | ";
	echo "<a href=\"day.php?year=$cyear&month=$cmonth&day=$cday\">$str</a>\n";
}
?>

 </BODY>
</HTML>
<?php
}
?>
