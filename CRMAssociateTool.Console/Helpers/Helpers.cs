using System.Xml;

namespace CRMAssociateTool.Console.Helpers
{
    public static class Helpers
    {
        public static bool ValidateXml(string xml)
        {
            try
            {
                new XmlDocument().LoadXml(xml);
                return true;
            }
            catch (XmlException e)
            {
                return false;
            }
        }
    }
}