<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
    <xsl:output method="xml" indent="yes"/>

    <xsl:template match="@* | node()">
      <all>
        <xsl:apply-templates select="//measuringchannel"/>
      </all>
    </xsl:template>
  
  <xsl:template match="measuringchannel">
    <xsl:if test="count(./period)!=48">
      <cnl>
        <xsl:attribute name="code">
          <xsl:value-of select="../@code"/>
        </xsl:attribute>
        <xsl:attribute name="name">
          <xsl:value-of select="../@name"/>
        </xsl:attribute>
        <xsl:attribute name="desc">
          <xsl:value-of select="./@code"/>_<xsl:value-of select="./@desc"/>
        </xsl:attribute>
        <xsl:value-of select="count(./period)"/>
      </cnl>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
