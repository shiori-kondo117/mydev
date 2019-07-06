<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:param name="struct_name"></xsl:param>
<xsl:output method="html" indent="yes" />
<xsl:template match="/">
<html>
<head>
	<title>構造体一覧</title>
</head>

<body>
<h3>構造体一覧</h3>
key:<xsl:value-of select="$struct_name"/>
<table border="1" cellspacing="0" cellpadding="0">
<thead>
	<tr>
		<th>構造体名</th>
		<th>名称</th>
		<th>ファイル名</th>
		<th>パス</th>
	</tr>
</thead>

<tbody>
	<xsl:apply-templates select="struct-info/struct[st-name=$struct_name]" />
</tbody>

</table>
</body>
</html>
</xsl:template>

<xsl:template match="struct">
	<xsl:param name="struct_name">
		<xsl:value-of select="./st-name" />
	</xsl:param>
	<tr>
		<td>
			<span>
			<xsl:attribute name="onclick">
				cmdSelect_Click('<xsl:value-of select="./st-name" />');
			</xsl:attribute>
			<xsl:value-of select="./st-name" />
			</span>
			<br/>
		</td>
		<td nowrap="true">
			<xsl:value-of select="./name" />
			<br/>
		</td>
		<td>
			<xsl:value-of select="./@file" />
			<br/>
		</td>
		<td nowrap="true">
			<xsl:value-of select="./@path" />
			<br/>
		</td>
	</tr>
</xsl:template>
</xsl:stylesheet>
