using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using Gems.Smev.Rpgu.Service.RecieveStatement.Facade.RV;

namespace DocxFromDotx
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var savePath = @"E:\temp\DocxFromDotx";
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(Resource1.RpguRv);

            var namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            namespaceManager.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
            namespaceManager.AddNamespace("guszkh", "http://xsd.smev.ru/ppu/guszkh");           
            var rvRequestXml = xmlDoc.SelectSingleNode("/soapenv:Envelope/soapenv:Body/guszkh:createSmvGusPodgRazreshVvodExplV2Request", namespaceManager).InnerXml;
            var serializerRs = new XmlSerializer(typeof(smvGusPodgRazreshVvodExplV2));
            var requestObj = (smvGusPodgRazreshVvodExplV2)serializerRs.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(rvRequestXml)));
            Console.WriteLine(requestObj.senderFio);

            var statementDocx = new RvStatementDocxGus("GusRv", requestObj).BuildReport();
            File.WriteAllBytes(Path.Combine(savePath,"rpguRv.docx"), statementDocx);
            Console.ReadKey();
        }
    }
}
