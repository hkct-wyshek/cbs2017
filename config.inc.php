<?php
ini_set('MAX_EXECUTION_TIME', -1);
ini_set('display_errors', 'Off');
 date_default_timezone_set("Asia/Taipei");
# $Id: config.inc.php,v 1.39.2.5 2005/08/20 09:46:44 jberanek Exp $

###########################################################################
#   MRBS Configuration File
#   Configure this file for your site.
#   You shouldn't have to modify anything outside this file.
###########################################################################

###################
# Database settings
###################
# Which database system: "pgsql"=PostgreSQL, "mysql"=MySQL
$dbsys = "mysql"; 
# Hostname of database server. For pgsql, can use "" instead of localhost
# to use Unix Domain Sockets instead of TCP/IP.
$db_host = "localhost";
#$db_host = "127.0.0.1";
# Database name:
$db_database = "classrm2016";
# Database login user name:
$db_login = "root";
# Database login password:
$db_password = '';
# Prefix for table names.  This will allow multiple installations where only
# one database is available
$db_tbl_prefix = "mrbs_";
# Uncomment this to NOT use PHP persistent (pooled) database connections:
# $db_nopersist = 1;

################################
# Site identification information
#################################
$mrbs_admin = "Department of IT & AV Facilities";
$mrbs_admin_email = "helpdesk@it.hkct.edu.hk";

# This is the text displayed in the upper left corner of every page. Either
# type the name of your organization, or you can put your logo like this :
# $mrbs_company = "<a href=http://www.your_organisation.com/>
# <img src=your_logo.gif border=0></a>";
$mrbs_company = "香港專業進修學校";

# This is to fix URL problems when using a proxy in the environment.
# If links inside MRBS appear broken, then specify here the URL of
# your MRBS root directory, as seen by the users. For example:
# $url_base =  "http://webtools.uab.ericsson.se/oam";
# It is also recommended that you set this if you intend to use email
# notifications, to ensure that the correct URL is displayed in the
# notification.
$url_base = "http://ctapps2.intranet.hkct.edu.hk/rmbook/classrm/2016/";


###################
# Calendar settings
###################
# Note: Be careful to avoid specify options that displays blocks overlaping
# the next day, since it is not properly handled.

# This setting controls whether to use "clock" based intervals (FALSE and
# the default) or user defined periods (TRUE).  If user-defined periods
# are used then $resolution, $morningstarts, $eveningends,
# $eveningends_minutes and $twentyfourhour_format are ignored.
$enable_periods = FALSE;

# Resolution - what blocks can be booked, in seconds.
# Default is half an hour: 1800 seconds.
$resolution = 3600;

# Start and end of day, NOTE: These are integer hours only, 0-23, and
# morningstarts must be < eveningends. See also eveningends_minutes.
$morningstarts = 9;
$eveningends   = 21;

# Minutes to add to $morningstarts to get to the real start of the day.
# Be sure to consider the value of $eveningends_minutes if you change
# this, so that you do not cause a day to finish before the start of
# the last period.  For example if resolution=3600 (1 hour)
# morningstarts = 8 and morningstarts_minutes = 30 then for the last
# period to start at say 4:30pm you would need to set eveningends = 16
# and eveningends_minutes = 30
$morningstarts_minutes = 0;

# Minutes to add to $eveningends hours to get the real end of the day.
# Examples: To get the last slot on the calendar to be 16:30-17:00, set
# eveningends=16 and eveningends_minutes=30. To get a full 24 hour display
# with 15-minute steps, set morningstarts=0; eveningends=23;
# eveningends_minutes=45; and resolution=900.
$eveningends_minutes = 0;

# Define the name or description for your periods in chronological order
# For example:
# $periods[] = "Period&nbsp;1"
# $periods[] = "Period&nbsp;2"
# ...
# or
# $periods[] = "09:15&nbsp;-&nbsp;09:50"
# $periods[] = "09:55&nbsp;-&nbsp;10:35"
# ...
# &nbsp; is used to ensure that the name or description is not wrapped
# when the browser determines the column widths to use in day and week
# views

# $periods[] = "Period&nbsp;1";
# $periods[] = "Period&nbsp;2";

# $periods[] = "9AM..11AM";
# $periods[] = "11AM..1PM";
# $periods[] = "1PM..2PM";
# $periods[] = "2PM..4PM";
# $periods[] = "4PM..6PM";
# $periods[] = "6PM..7PM";
# $periods[] = "7PM..10PM";


# Start of week: 0 for Sunday, 1 for Monday, etc.
$weekstarts = 0;

# Trailer date format: 0 to show dates as "Jul 10", 1 for "10 Jul"
$dateformat = 0;

# Time format in pages. 0 to show dates in 12 hour format, 1 to show them
# in 24 hour format
$twentyfourhour_format = 0;

########################
# Miscellaneous settings
########################

# Maximum repeating entrys (max needed +1):
$max_rep_entrys = 365 + 1;

# Default report span in days:
$default_report_days = 60;

# Results per page for searching:
$search["count"] = 20;

# Page refresh time (in seconds). Set to 0 to disable
$refresh_rate = 0;

# should areas be shown as a list or a drop-down select box?
$area_list_format = "list";
#$area_list_format = "select";

# Entries in monthly view can be shown as start/end slot, brief description or
# both. Set to "description" for brief description, "slot" for time slot and
# "both" for both. Default is "both", but 6 entries per day are shown instead
# of 12.
$monthly_view_entries_details = "both";

# To view weeks in the bottom (trailer.inc) as week numbers (42) instead of
# 'first day of the week' (13 Oct), set this to TRUE
$view_week_number = FALSE;

# To display times on right side in day and week view, set to TRUE;
$times_right_side = FALSE;

# Control the active cursor in day/week/month views.
$javascript_cursor = true; # Change to false if clients have old browsers
                           # incompatible with JavaScript.
$show_plus_link = true; # Change to true to always show the (+) link as in
                        # MRBS 1.1.
$highlight_method = "hybrid"; # One of "bgcolor", "class", "hybrid".

# Define default starting view (month, week or day)
# Default is day
$default_view = "day";

# Define default room to start with (used by index.php)
# Room numbers can be determined by looking at the Edit or Delete URL for a
# room on the admin page.
# Default is 0
$default_room = 3;

###############################################
# Authentication settings - read AUTHENTICATION
###############################################
$auth["session"] = "php"; # How to get and keep the user ID. One of
			  # "http" "php" "cookie" "ip" "host" "nt" "omni".
$auth["type"] = "db"; # How to validate the user/password. One of "none"
                          # "config" "db" "db_ext" "pop3" "imap" "ldap" "nis"
                          # "nw" "ext".

# The list of administrators (can modify other peoples settings)
# $auth["admin"][] = "127.0.0.1";	# localhost IP address. Useful with IP sessions.
# $auth["admin"][] = "administrator";	# A user name from the user list. Useful 
                                    #with most other session schemes.
#$auth["admin"][] = "10.0.0.1";
#$auth["admin"][] = "10.0.0.2";
#$auth["admin"][] = "10.0.0.3";

# 'auth_config' user database
# Format: $auth["user"]["name"] = "password";

$auth["admin"][] = "administrator";
$auth["admin"][] = "DFM";


$cunit[0] = "AEC-YJDPT";
$cunit[1] = "CIE";
$cunit[2] = "DHSS";
$cunit[3] = "DSS";
$cunit[4] = "DLGS-CL";

$cunit[5] = "DLGS-YJD";
$cunit[6] = "DSF";
$cunit[7] = "DLGS-AL";
$cunit[8] = "DDTB-CIT";
$cunit[9] = "DDTB-TP";

$cunit[10] = "DDTB-CAD";
$cunit[11] = "SVT-ERB";
$cunit[12] = "DDTB-ALP";
$cunit[13] = "DDTB-BSP";
$cunit[14] = "DTH";

$cunit[15] = "DSD";
$cunit[16] = "SLP";
$cunit[17] = "SE";
$cunit[18] = "OTHER";


$auth["courseadmin"][] = "AEC-YJDPT";
$auth["courseadmin"][] = "CIE";
$auth["courseadmin"][] = "DHSS";
$auth["courseadmin"][] = "DSS";
$auth["courseadmin"][] = "DLGS-CL";

$auth["courseadmin"][] = "DLGS-YJD";
$auth["courseadmin"][] = "DSF";
$auth["courseadmin"][] = "DLGS-AL";
$auth["courseadmin"][] = "DDTB-CIT";
$auth["courseadmin"][] = "DDTB-TP";

$auth["courseadmin"][] = "DDTB-CAD";
$auth["courseadmin"][] = "SVT-ERB";
$auth["courseadmin"][] = "DDTB-ALP";
$auth["courseadmin"][] = "DDTB-BSP";
$auth["courseadmin"][] = "DTH";

$auth["courseadmin"][] = "DSD";
$auth["courseadmin"][] = "SLP";
$auth["courseadmin"][] = "SE";
$auth["courseadmin"][] = "OTHER";
$auth["courseadmin"][] = "DIT-VISITOR";


$auth["depthead"][] = "AEC-YJDPHEAD";
$auth["depthead"][] = "CIEHEAD";
$auth["depthead"][] = "DHSSHEAD";
$auth["depthead"][] = "DSDAHEAD";
$auth["depthead"][] = "DLGS-CLHEAD";

$auth["depthead"][] = "DDTB-CITHEAD";
$auth["depthead"][] = "DDTBT-TPHEAD";
$auth["depthead"][] = "DDTB-CADHEAD";
$auth["depthead"][] = "DLGS-YJDHEAD";
$auth["depthead"][] = "DSFHEAD";

$auth["depthead"][] = "DLGS-ALHEAD";
$auth["depthead"][] = "SVT-ERBHEAD";
$auth["depthead"][] = "DDTB-ALPHEAD";
$auth["depthead"][] = "DDTB-BSPHEAD";
$auth["depthead"][] = "DTH-THPHEAD";
$auth["depthead"][] = "DSSHEAD";

$auth["eb"][] = "EB";
$auth["eb"][] = "BH";

$auth["user"]["administrator"] = "91664095";
$auth["user"]["DA"] = "daLWMct6";

# $auth["user"]["alice"] = "a";
# $auth["user"]["bob"] = "b";

# 'session_http' configuration settings
$auth["realm"]  = "mrbs";

# 'auth_ext' configuration settings
$auth["prog"]   = "";
$auth["params"] = "";

# 'auth_db_ext' configuration settings
$auth['db_ext']['db_host'] = 'localhost';
$auth['db_ext']['db_username'] = 'authuser';
$auth['db_ext']['db_password'] = 'authpass';
$auth['db_ext']['db_name'] = 'authdb';
$auth['db_ext']['db_table'] = 'users';
$auth['db_ext']['column_name_username'] = 'name';
$auth['db_ext']['column_name_password'] = 'password';
# Either 'md5', 'sha1', 'crypt' or 'plaintext'
$auth['db_ext']['password_format'] = 'md5';

# 'auth_ldap' configuration settings
# Where is the LDAP server
#$ldap_host = "localhost";
# LDAP base distinguish name
# See AUTHENTICATION for details of how check against multiple base dn's
#$ldap_base_dn = "ou=organizationalunit,dc=my-domain,dc=com";
# Attribute within the base dn that contains the username
#$ldap_user_attrib = "uid";

# 'auth_ldap' extra configuration for ldap configuration of who can use
# the system
# If it's set, the $ldap_filter will be combined with the value of
# $ldap_user_attrib like this:
#   (&($ldap_user_attrib=username)($ldap_filter))
# After binding to check the password, this check is used to see that
# they are a valid user of mrbs.
#$ldap_user_filter = "mrbsuser=y";

# 'auth_imap' configuration settings
# See AUTHENTICATION for details of how check against multiple servers
# Where is the IMAP server
$imap_host = "imap-server-name";
# The IMAP server port
$imap_port = "143";

# 'auth_pop3' configuration settings
# See AUTHENTICATION for details of how check against multiple servers
# Where is the POP3 server
$pop3_host = "pop3-server-name";
# The POP3 server port
$pop3_port = "110";


###############################################
# Email settings
###############################################

# Set to TRUE if you want to be notified when entries are booked. Default is
# FALSE
define ("MAIL_ADMIN_ON_BOOKINGS", FALSE);

# Set to TRUE if you want AREA ADMIN to be notified when entries are booked.
# Default is FALSE. Area admin emails are set in room_area admin page.
define ("MAIL_AREA_ADMIN_ON_BOOKINGS", FALSE);

# Set to TRUE if you want ROOM ADMIN to be notified when entries are booked.
# Default is FALSE. Room admin emails are set in room_area admin page.
define ("MAIL_ROOM_ADMIN_ON_BOOKINGS", FALSE);

# Set to TRUE if you want ADMIN to be notified when entries are deleted. Email
# will be sent to mrbs admin, area admin and room admin as per above settings,
# as well as to booker if MAIL_BOOKER is TRUE (see below).
define ("MAIL_ADMIN_ON_DELETE", FALSE);

# Set to TRUE if you want to be notified on every change (i.e, on new entries)
# but also each time they are edited. Default is FALSE (only new entries)
define ("MAIL_ADMIN_ALL", FALSE);

# Set to TRUE is you want to show entry details in email, otherwise only a
# link to view_entry is provided. Irrelevant for deleted entries. Default is
# FALSE.
define ("MAIL_DETAILS", FALSE);

# Set to TRUE if you want BOOKER to receive a copy of his entries as well any
# changes (depends of MAIL_ADMIN_ALL, see below). Default is FALSE. To know
# how to set mrbs to send emails to users/bookers, see INSTALL.
define ("MAIL_BOOKER", FALSE);

# If MAIL_BOOKER is set to TRUE (see above) and you use an authentication
# scheme other than 'auth_db', you need to provide the mail domain that will
# be appended to the username to produce a valid email address (ie.
# "@domain.com").
define ("MAIL_DOMAIN", '');

# If you use MAIL_DOMAIN above and username returned by mrbs contains extra
# strings appended like domain name ('username.domain'), you need to provide
# this extra string here so that it will be removed from the username.
define ("MAIL_USERNAME_SUFFIX", '');

# Set the name of the Backend used to transport your mails. Either "mail",
# "smtp" or "sendmail". Default is 'mail'. See INSTALL for more details.
define ("MAIL_ADMIN_BACKEND", "mail");

#*******************
# Sendmail settings

# Set the path of the Sendmail program (only used with "sendmail" backend).
# Default is "/usr/bin/sendmail"
define ("SENDMAIL_PATH", "/usr/bin/sendmail");

# Set additional Sendmail parameters (only used with "sendmail" backend).
# (example "-t -i"). Default is ""
define ("SENDMAIL_ARGS", '');

#*******************
# SMTP settings

# Set smtp server to connect. Default is 'localhost' (only used with "smtp"
# backend).
define ("SMTP_HOST", "localhost");

# Set smtp port to connect. Default is '25' (only used with "smtp" backend).
define ("SMTP_PORT", 25);

# Set whether or not to use SMTP authentication. Default is 'FALSE'
define ("SMTP_AUTH", FALSE);

# Set the username to use for SMTP authentication. Default is ""
define ("SMTP_USERNAME", '');

# Set the password to use for SMTP authentication. Default is ""
define ("SMTP_PASSWORD", '');

#****************************
# Miscellaneous settings

# Set the language used for emails (choose an available lang.* file).
# Default is 'en'.
define ("MAIL_ADMIN_LANG", 'en');

# Set the email address of the From field. Default is $mrbs_admin_email
define ("MAIL_FROM", $mrbs_admin_email);

# Set the recipient email. Default is $mrbs_admin_email. You can define
# more than one recipient like this "john@doe.com,scott@tiger.com"
define ("MAIL_RECIPIENTS", $mrbs_admin_email);

# Set email address of the Carbon Copy field. Default is ''. You can define
# more than one recipient (see MAIL_RECIPIENTS)
define ("MAIL_CC", '');

# Set the content of the Subject field for added/changed entries.
$mail["subject"] = "Entry added/changed for $mrbs_company MRBS";

# Set the content of the Subject field for deleted fields.
$mail["subject_delete"] = "Entry deleted for $mrbs_company MRBS";

# Set the content of the message when a new entry is booked. What you type
# here will be added at the top of the message body.
$mail["new_entry"] = "A new entry has been booked, here are the details:";

# Set the content of the message when an entry is modified. What you type
# here will be added at the top of the message body.
$mail["changed_entry"] = "An entry has been modified, here are the details:";

# Set the content of the message when an entry is deleted. What you type
# here will be added at the top of the message body.
$mail["deleted_entry"] = "An entry has been deleted, here are the details:";

##########
# Language
##########

# Set this to 1 to use UTF-8 in all pages and in the database, otherwise
# text gets enterered in the database in different encodings, dependent
# on the users' language
$unicode_encoding = 1;

# Set this to a different language specifier to default to different
# language tokens. This must equate to a lang.* file in MRBS.
# e.g. use "fr" to use the translations in "lang.fr" as the default
# translations
$default_language_tokens = "utf8";

# Set this to 1 to disable the automatic language changing MRBS performs
# based on the user's browser language settings. It will ensure that
# the language displayed is always the value of $default_language_tokens,
# as specified above
$disable_automatic_language_changing = 1;

# Set this to a valid locale (for the OS you run the MRBS server on)
# if you want to override the automatic locale determination MRBS
# performs
$override_locale = "";

# faq file language selection. IF not set, use the default english file.
# IF your language faq file is available, set $faqfilelang to match the
# end of the file name, including the underscore (ie. for site_faq_fr.html
# use "_fr"
$faqfilelang = ""; 

# This next require must be done after the definitions above, as the definitions
# are used in the included file
require_once "language.inc";

#############
# Entry Types
#############
# This array maps entry type codes (letters A through J) into descriptions.
# Each type has a color (see TD.x classes in the style sheet mrbs.css).
#    A=Pink  B=Blue-green  C=Peach  D=Yellow      E=Light blue
#    F=Tan   G=Red         H=Aqua   I=Light green J=Gray
# The value for each type is a short (one word is best) description of the
# type. The values must be escaped for HTML output ("R&amp;D").
# Please leave I and E alone for compatibility.
# If a type's entry is unset or empty, that type is not defined; it will not
# be shown in the day view color-key, and not offered in the type selector
# for new or edited entries.

# $typel["A"] = "A";
# $typel["B"] = "B";
# $typel["C"] = "C";
# $typel["D"] = "D";
# $typel["E"] = get_vocab("external");
# $typel["F"] = "F";
# $typel["G"] = "G";
# $typel["H"] = "H";
# $typel["I"] = get_vocab("internal");
# $typel["J"] = "J";
/*
$cunit[0] = "AEC-YJDP";
$cunit[1] = "CIE-HKCC";
$cunit[2] = "CIE-PAS";
$cunit[3] = "CIE-UGS";
$cunit[4] = "DAT";
$cunit[5] = "DB";
$cunit[6] = "DHSS";
$cunit[7] = "DMC";
$cunit[8] = "DLC-CL";
$cunit[9] = "DMGS-TP";
$cunit[10] = "DMGS-YJDF";
$cunit[11] = "DMGS-SM";
$cunit[12] = "DMGS-AL";
$cunit[13] = "SVT-ERB";
$cunit[14] = "OTHER";
$cunit[15] = "_1";
$cunit[16] = "_2";
*/

$cunitname[0] = get_vocab($cunit[0]);
$cunitname[1] = get_vocab($cunit[1]);
$cunitname[2] = get_vocab($cunit[2]);
$cunitname[3] = get_vocab($cunit[3]);
$cunitname[4] = get_vocab($cunit[4]);
$cunitname[5] = get_vocab($cunit[5]);
$cunitname[6] = get_vocab($cunit[6]);
$cunitname[7] = get_vocab($cunit[7]);
$cunitname[8] = get_vocab($cunit[8]);
$cunitname[9] = get_vocab($cunit[9]);
$cunitname[10] = get_vocab($cunit[10]);
$cunitname[11] = get_vocab($cunit[11]);
$cunitname[12] = get_vocab($cunit[12]);
$cunitname[13] = get_vocab($cunit[13]);
$cunitname[14] = get_vocab($cunit[14]);
$cunitname[15] = get_vocab($cunit[15]);
$cunitname[16] = get_vocab($cunit[16]);
$cunitname[17] = get_vocab($cunit[17]);
$cunitname[18] = get_vocab($cunit[18]);
//$cunitname[19] = get_vocab($cunit[19]);

$cpname = array(
				0 => array('code'=>'HMTSB', 'name'=>'何文田校舍'),
				1 => array('code'=>'HMT',   'name'=>'何文田(勞校)'),
				2 => array('code'=>'MOS',   'name'=>'馬鞍山校舍(MOS)'),
				3 => array('code'=>'MUC',   'name'=>'馬鞍山本科校園(MUC)'),
				4 => array('code'=>'JDN',   'name'=>'佐敦培訓中心'),
				5 => array('code'=>'OPC',   'name'=>'開源道培訓中心'),
				6 => array('code'=>'PLS',   'name'=>'砵蘭街培訓中心'),
				7 => array('code'=>'TKO',   'name'=>'將軍澳培訓中心'),
				8 => array('code'=>'YL',    'name'=>'元朗教學中心'),
				9 => array('code'=>'CSW',   'name'=>'長沙灣培訓中心'),
				10=> array('code'=>'AUS',   'name'=>'柯士甸道教學中心'),
);

##########################################
# PHP System Configuration - internal use, do not change
##########################################
# Disable magic quoting on database returns:
# marco mark @ 2013-7-30
#set_magic_quotes_runtime(0);

# Make sure notice errors are not reported, they can break mrbs code:
error_reporting (E_ALL ^ E_NOTICE);

# These variables specify the names of the tables in the database
# These should not need to be changed.  Please change $db_tbl_prefix
# in the database section above.
$tbl_area   = $db_database.".".$db_tbl_prefix . "area";
$tbl_entry  = $db_database.".".$db_tbl_prefix . "entry";
$tbl_repeat = $db_database.".".$db_tbl_prefix . "repeat";
$tbl_room   = $db_database.".".$db_tbl_prefix . "room";
$tbl_users  = $db_database.".".$db_tbl_prefix . "users";
$tbl_remove = $db_database.".".$db_tbl_prefix . "remove_day";

# MRBS developers, make sure to update this string before each release:
$mrbs_version = "MRBS 1.2.3";

?>
