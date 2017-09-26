# DSC.EWSsendmail
small tool for sending customized mass emails via the MS Office365 / Exchange Online EWS managed API<br>
<ul>
	<li>Reads  recipient list from csv file</li>
	<li>can read up to 5 customized strings (per recipient) from CSV; for use in email text/body</li>
	<li>allows for HTML formatting of email text body</li>
	<li>successively sends single email to each recipient via EWS managed API (EWS API limits apply)</li>
	<li> relies on <a href=https://msdn.microsoft.com/de-de/library/office/dn528373(v=exchg.150).aspx>Microsoft.Exchange.WebServices.dll</a> (tested with 15.0.913.15)
	<li>tested with Office365 / Exchange Online</li>
</ul>
