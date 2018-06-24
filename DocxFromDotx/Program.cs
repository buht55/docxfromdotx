using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using Gems.Smev.Rpgu.Service.RecieveStatement.Facade.RV;

namespace DocxFromDotx
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var rvRequestXml = Resource1.RpguRv;
            var serializerRs = new XmlSerializer(typeof(smvGusPodgRazreshVvodExplV2));
            var requestObj = (smvGusPodgRazreshVvodExplV2)serializerRs.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(rvRequestXml)));
            Console.WriteLine(requestObj.senderFio);
            Console.ReadKey();
        }
    }
}
