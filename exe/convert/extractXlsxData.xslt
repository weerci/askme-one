<?xml version="1.0" encoding="utf-8"?>
<!-- File: extractXlsxData.xslt -->
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    exclude-result-prefixes="xsl r msxsl">

  <xsl:output version="1.0" method="xml"
    standalone="yes" encoding="utf-8"/>

  <xsl:param name="prmPathToSharedStringsFile">
    <xsl:text></xsl:text> 
  </xsl:param>
  <xsl:variable name="varDocSharedStrings">
    <xsl:copy-of select="document(concat($prmPathToSharedStringsFile, 'sharedStrings.xml'))" />
  </xsl:variable>

  <xsl:template match="/">
    <sheetData>
      <xsl:apply-templates select="/./*[local-name()='worksheet']/*[local-name()='sheetData']/*" />
    </sheetData>
  </xsl:template>

  <xsl:template match="@*|node()">
    <xsl:variable name="varLocalName">
      <xsl:value-of select="local-name()" />
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="string-length($varLocalName)>0 and $varLocalName != 'v'">
        <xsl:element name="{$varLocalName}">
          <xsl:for-each select="@*">
            <xsl:attribute name="{name()}">
              <xsl:value-of select="." />
            </xsl:attribute>
          </xsl:for-each>
          <xsl:apply-templates select="*" />
        </xsl:element>
      </xsl:when>
      <xsl:when test="$varLocalName='v'">
        <xsl:choose>
          <xsl:when test="parent::*/@t='s'">
            <xsl:variable name="varNumSs">
              <xsl:value-of select="." />
            </xsl:variable>
            <v>
              <xsl:value-of select="msxsl:node-set($varDocSharedStrings)/./*[local-name()='sst']/*[local-name()='si'][number($varNumSs) + 1]/*[local-name()='t']" />
            </v>
          </xsl:when>
          <xsl:otherwise>
            <v>
              <xsl:value-of select="." />
            </v>
          </xsl:otherwise> 
        </xsl:choose> 
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="@*|node()" />
      </xsl:otherwise>
    </xsl:choose> 
  </xsl:template>

</xsl:stylesheet> 

