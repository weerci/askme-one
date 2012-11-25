<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
    <xsl:output method="xml" indent="yes"/>
    <xsl:template match="/">
      <rows>
	      <xsl:apply-templates select="//row"/>
      </rows>
    </xsl:template>
    
    <xsl:template match="row">
      <row code="{./c[1]/v}" ob="{./c[2]/v}" name="{./c[3]/v}" fid="{./c[4]/v}" name2="{./c[5]/v}"/>
    </xsl:template>
  
</xsl:stylesheet>
