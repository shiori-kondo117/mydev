<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:param name="struct_name"/>
<xsl:output method="html" indent="yes" />
<xsl:template match="/">

<html>
<head>
	<title>構造体一覧</title>
</head>

<body>
<table border="1" cellspacing="0" cellpadding="0">
<!--thead>
	<tr>
		<th>構造体名</th>
		<th>名称</th>
		<th>定義</th>
	</tr>
</thead-->

<tbody>
	<xsl:apply-templates select="struct-info/struct[st-name=$struct_name]" />
</tbody>

</table>
</body>
</html>
</xsl:template>

<xsl:template match="struct">
	<tr>
		<th valign="top" width="100">構造体名</th>
		<td valign="top">
			<xsl:value-of select="./st-name" />
		</td>
	</tr>
	<tr>
		<th valign="top">名称</th>
		<td valign="top">
			<xsl:value-of select="./name" />
		</td>
	</tr>
	<tr>
		<th valign="top">定義</th>
		<td valign="top" nowrap="true" width="500">
			<pre><xsl:value-of select="./define" /></pre>
		</td>
	</tr>
</xsl:template>
</xsl:stylesheet>
