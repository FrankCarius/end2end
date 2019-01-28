<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
<html>
<head><title>MailboxPerfmon Status</title></head>
<body>
<h1>MailboxPerfmon Status</h1> 
<h2>Parameters</h2> 
<table border="1" width="100%">
	<tr><th>Starttime</th><td><xsl:value-of select="MailboxPerfmon/starttime" /></td></tr>
	<tr><th>EndTime</th> <td><xsl:value-of select="MailboxPerfmon/endtime" /></td></tr>
	<tr><th>gelesene Objekte</th> <td><xsl:value-of select="MailboxPerfmon/totalread" /></td></tr>
</table>
<h2>Details</h2> 
<table border="1" width="100%">
<tr bgcolor="#808080">
	<th>Mailbox:</th> 
	<th>Timestamp:</th> 
	<th>Dauer:</th> 
  </tr>
<xsl:for-each select="MailboxPerfmon/object"> 
	<tr>
	<td>
		<xsl:value-of select="user" /> 
	</td>
	<td>
		<xsl:value-of select="timestamp2" /> 
	</td>
	<td>
		<xsl:value-of select="dauer" /> 
	</td>
	</tr>
</xsl:for-each>
</table>
</body>
</html>
</xsl:template>
</xsl:stylesheet> 
