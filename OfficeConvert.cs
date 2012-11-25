using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using System.IO.Packaging;

namespace Askme
{
    class OfficeFiles
    {
        // Constructors
        public OfficeFiles(string fileCodeBase)
        {
            this.fileCodeBase = fileCodeBase;
        }

        // Methods
        public List<string> GetObjectIds()
        {
            try
            {
                package = Package.Open(fileCodeBase);

                XmlDocument idObjects = GetPartFile(OfficeFilePart.Sheet, 1);
                XmlNamespaceManager xnm = new XmlNamespaceManager(idObjects.NameTable);
                xnm.AddNamespace("def", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                List<string> listString = new List<string>();
                XmlNodeList elemList = idObjects.SelectNodes("//def:row/def:c[starts-with(@r, 'C')]/def:v", xnm);

                for (int i = 2; i < elemList.Count; i++)
                    listString.Add(elemList[i].InnerXml);

                return listString.Distinct().ToList();
            }
            finally
            {
                package.Close();
            }
        }
        public void CreateCodeBaseXML()
        {
            try
            {
                package = Package.Open(fileCodeBase);

                XmlDocument sharedString = GetPartFile(OfficeFilePart.ShareString, 0);
                sharedString.Save(Common.SHARED_STRING);
                XmlDocument sheet = GetPartFile(OfficeFilePart.Sheet, 1);

                XsltArgumentList xsltArgListSheet = new XsltArgumentList();
                XsltSettings settings = new XsltSettings(true, true);
                XslCompiledTransform xslTransSheet = new XslCompiledTransform();
                xslTransSheet.Load(Common.XSLT_CODE_BASE, settings, new XmlUrlResolver());
                xsltArgListSheet.AddParam("prmDocSharedStrings", "", sharedString.CreateNavigator());
                string sOutXmlSheet = System.String.Empty;
                using (FileStream fs = new FileStream(Common.XML_CODE_BASE, FileMode.Create))
                {
                    xslTransSheet.Transform(sheet.CreateNavigator(), xsltArgListSheet, fs);
                }

                XslCompiledTransform xslRowSheet = new XslCompiledTransform();
                xslRowSheet.Load(Common.XSLT_TO_ROW);
                xslRowSheet.Transform(Common.XML_CODE_BASE, Common.XML_ROW);

            }
            finally
            {
                package.Close();
            }
        }

        // Private methods
        private XmlDocument GetPartFile(OfficeFilePart filePart, int index)
        {
            XmlDocument xmlDocSharedStrings = new XmlDocument();
            Uri uriPath = null;
            switch (filePart)
            {
                case OfficeFilePart.ShareString:
                    uriPath = new Uri(OfficeConst.sharedString, UriKind.Relative);
                    break;
                case OfficeFilePart.Sheet:
                    uriPath = new Uri(String.Format(OfficeConst.formatSheet, index), UriKind.Relative);
                    break;
                case OfficeFilePart.Table:
                    uriPath = new Uri(String.Format(OfficeConst.formatTable, index), UriKind.Relative);
                    break;
                default:
                    break;
            }
            xmlDocSharedStrings.Load(package.GetPart(uriPath).GetStream());
            return xmlDocSharedStrings;
        }

        // Fields
        private string fileCodeBase;
        private Package package;
    }

    static class OfficeConst
    {
        public const string sharedString = "/xl/sharedStrings.xml";
        public const string formatSheet = "/xl/worksheets/sheet{0}.xml";
        public const string formatTable = "/xl/tables/table{0}.xml";
        public const int numberSheet = 1;
        public const int numberTable = 1;
    }

    enum OfficeFilePart { ShareString, Sheet, Table};
}
