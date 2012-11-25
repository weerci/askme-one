<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" encoding="windows-1251" indent="yes"/>

  <xsl:variable name="varDocResult">
    <xsl:copy-of select="document('xmlRow.xml')" />
  </xsl:variable>

  <xsl:key name="obj" match="oracleResult" use="concat(@N_OB, @N_FID)"/>

  <xsl:param name="class" select="80020"/>
  <xsl:param name="version" select="2"/>
  <xsl:param name="number" select="1100"/>
  <xsl:param name="timestamp" select="20091215015013"/>
  <xsl:param name="daylightsavingtime" select="0"/>
  <xsl:param name="day" select="20091214"/>
  <xsl:param name="name" select="'Русэнергосбыт'"/>
  <xsl:param name="inn" select="7706284124"/>
  <xsl:param name="name2" select="'Красноярская железная дорога; Граница с Красноярскэнерго'"/>
  <xsl:param name="inn3" select="7700000105"/>

  <xsl:template match="/">
    <message class="{$class}" version="{$version}" number="{$number}">
      <datetime>
        <timestamp>
          <xsl:value-of select="$timestamp"/>
        </timestamp>
        <daylightsavingtime>
          <xsl:value-of select="$daylightsavingtime"/>
        </daylightsavingtime>
        <day>
          <xsl:value-of select="$day"/>
        </day>
      </datetime>
      <sender>
        <name>
          <xsl:value-of select="$name"/>
        </name>
        <inn>
          <xsl:value-of select="$inn"/>
        </inn>
      </sender>
      <area>
        <name>
          <xsl:value-of select="$name2"/>
        </name>
        <inn>
          <xsl:value-of select="$inn3"/>
        </inn>
        <xsl:for-each select="//oracleResult[generate-id()=generate-id(key('obj',concat(@N_OB, @N_FID)))]">
          <xsl:sort select="@N_OB"/>
          <xsl:sort select="@N_FID"/>

          <xsl:variable name="ob" select="./@N_OB"/>
          <xsl:variable name="fid" select="./@N_FID"/>

          <xsl:variable name="pathOb" select="msxsl:node-set($varDocResult)//row[@ob=$ob][@fid=$fid]"/>
          <xsl:variable name="code" select="$pathOb/@code"/>
          <xsl:variable name="name" select="concat($pathOb/@name, $pathOb/@name2)"/>
          <xsl:if test="$pathOb">
            <measuringpoint code="{$code}" name="{$name}">
              <measuringchannel code="01" desc="Активная прием">
                <xsl:variable name="activeReception" select="//oracleResult[@N_OB=$ob][@N_FID=$fid][@N_GR_TY=1]"/>
                <xsl:choose>
                  <xsl:when test="$activeReception">
                    <xsl:apply-templates select="$activeReception"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="emptyMeasuringChannel">
                      <xsl:with-param name="n" select="number(47)"/>
                    </xsl:call-template>                    
                  </xsl:otherwise>
                </xsl:choose>
              </measuringchannel>
              <measuringchannel code="02" desc="Активная отдача">
                <xsl:variable name="activePass" select="//oracleResult[@N_OB=$ob][@N_FID=$fid][@N_GR_TY=2]"/>
                <xsl:choose>
                  <xsl:when test="$activePass">
                    <xsl:apply-templates select="$activePass"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="emptyMeasuringChannel">
                      <xsl:with-param name="n" select="number(47)"/>
                    </xsl:call-template>
                  </xsl:otherwise>
                </xsl:choose>
              </measuringchannel>
              <measuringchannel code="03" desc="Реактивная прием">
                <xsl:variable name="reactiveReception" select="//oracleResult[@N_OB=$ob][@N_FID=$fid][@N_GR_TY=3]"/>
                <xsl:choose>
                  <xsl:when test="$reactiveReception">
                    <xsl:apply-templates select="$reactiveReception"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="emptyMeasuringChannel">
                      <xsl:with-param name="n" select="number(47)"/>
                    </xsl:call-template>
                  </xsl:otherwise>
                </xsl:choose>
              </measuringchannel>
              <measuringchannel code="04" desc="Реактивная отдача">
                <xsl:variable name="reactivePass" select="//oracleResult[@N_OB=$ob][@N_FID=$fid][@N_GR_TY=4]"/>
                <xsl:choose>
                  <xsl:when test="$reactivePass">
                    <xsl:apply-templates select="$reactivePass"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="emptyMeasuringChannel">
                      <xsl:with-param name="n" select="number(47)"/>
                    </xsl:call-template>
                  </xsl:otherwise>
                </xsl:choose>
              </measuringchannel>
            </measuringpoint>
          </xsl:if>
        </xsl:for-each>
      </area>
    </message>

  </xsl:template>

  <xsl:template match="oracleResult">
    <xsl:variable name="i" select="./@N_INTER_RAS*1-1"/>
    <xsl:choose>
      <xsl:when test="($i mod 2)=0">
        <period>
          <xsl:attribute name="start">
            <xsl:choose>
              <xsl:when test="$i=0">0000</xsl:when>
              <xsl:otherwise>
                <xsl:number value="$i*50" format="0001"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="end">
            <xsl:number value="$i*50+30" format="0001"/>
          </xsl:attribute>
          <value status="0">
            <xsl:value-of select="./@VAL"/>
          </value>
        </period>
      </xsl:when>
      <xsl:otherwise>
        <period>
          <xsl:attribute name="start">
            <xsl:number value="($i*1-1)*50+30" format="0001"/>
          </xsl:attribute>
          <xsl:attribute name="end">
            <xsl:choose>
              <xsl:when test="$i=47">0000</xsl:when>
              <xsl:otherwise>
                <xsl:number value="($i+1)*50" format="0001"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <value status="0">
            <xsl:value-of select="./@VAL"/>
          </value>
        </period>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="emptyMeasuringChannel">
    <xsl:param name="n"/>
    <xsl:choose>
      <xsl:when test="$n=0">
        <period start="0000" end="0030">
          <value status="0">0</value>
        </period>
      </xsl:when>
      <xsl:otherwise>
          <xsl:call-template name="emptyMeasuringChannel">
            <xsl:with-param name="n" select="$n - 1"/>
          </xsl:call-template>
        <xsl:choose>
          <xsl:when test="($n mod 2)=0">
            <period>
              <xsl:attribute name="start">
                <xsl:choose>
                  <xsl:when test="$n=0">0000</xsl:when>
                  <xsl:otherwise>
                    <xsl:number value="$n*50" format="0001"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="end">
                <xsl:number value="$n*50+30" format="0001"/>
              </xsl:attribute>
              <value status="0">0</value>
            </period>
          </xsl:when>
          <xsl:otherwise>
            <period>
              <xsl:attribute name="start">
                <xsl:number value="($n*1-1)*50+30" format="0001"/>
              </xsl:attribute>
              <xsl:attribute name="end">
                <xsl:choose>
                  <xsl:when test="$n=47">0000</xsl:when>
                  <xsl:otherwise>
                    <xsl:number value="($n+1)*50" format="0001"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <value status="0">0</value>
            </period>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>
