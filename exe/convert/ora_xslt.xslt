<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" indent="yes"/>

  <xsl:template match="/">
    <DocumentElement>
      <xsl:apply-templates select="//oracleResult"/>
    </DocumentElement>
  </xsl:template >

  <xsl:template match="oracleResult">
    <oracleResult>
      <xsl:attribute name="N_OB"><xsl:value-of select="./N_OB"/></xsl:attribute>
      <xsl:attribute name="N_FID"><xsl:value-of select="./N_FID"/></xsl:attribute>
      <xsl:attribute name="N_GR_TY"><xsl:value-of select="./N_GR_TY"/></xsl:attribute>
      <xsl:attribute name="DD_MM_YYYY"><xsl:value-of select="./DD_MM_YYYY"/></xsl:attribute>
      <xsl:attribute name="N_INTER_RAS"><xsl:value-of select="./N_INTER_RAS"/></xsl:attribute>
      <xsl:attribute name="VAL"><xsl:value-of select="./VAL"/></xsl:attribute>
      <xsl:attribute name="MIN_0"><xsl:value-of select="./MIN_0"/></xsl:attribute>
      <xsl:attribute name="MIN_1"><xsl:value-of select="./MIN_1"/></xsl:attribute>
    </oracleResult>

  </xsl:template>
  
</xsl:stylesheet>
