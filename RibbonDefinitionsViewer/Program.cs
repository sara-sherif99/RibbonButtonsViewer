using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.ServiceModel.Description;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace RibbonButtonsViewer
{
    internal class Program
    {
        static void Main(string[] args)
        {            
            IOrganizationService service = CreateCrmConnection(
                ConfigurationManager.AppSettings["CrmUrl"],
                ConfigurationManager.AppSettings["CrmUsername"],
                ConfigurationManager.AppSettings["CrmPassword"]
                );

            Console.WriteLine("\nEnter the logical name of the entity: ");
            string entityName = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(entityName))
            {
                Console.WriteLine("Entity name cannot be empty.");
                return;
            }

            byte[] zippedRibbonData = GetEntityRibbonData(entityName,service);

            using (var memStream = new MemoryStream(zippedRibbonData))
            using (var zipArchive = new ZipArchive(memStream, ZipArchiveMode.Read))
            {
                ZipArchiveEntry ribbonEntry = zipArchive.GetEntry("RibbonXml.xml");
                if (ribbonEntry != null)
                {
                    using (var reader = new StreamReader(ribbonEntry.Open()))
                    {
                        string xmlContent = reader.ReadToEnd();
                        GeneratePdfFile(entityName, xmlContent);
                    }
                }
            }
        }
        static IOrganizationService CreateCrmConnection(string url, string username, string password)
        {
            ClientCredentials devCredentials = new ClientCredentials();
            devCredentials.UserName.UserName = username;
            devCredentials.UserName.Password = password;
            Uri devServiceUri = new Uri(url);


            OrganizationServiceProxy devProxy = new OrganizationServiceProxy(devServiceUri, null, devCredentials, null);
            devProxy.EnableProxyTypes();
            devProxy.Timeout = new TimeSpan(0, 5, 0);

            return devProxy;
        }
        static byte[] GetEntityRibbonData(string entityName, IOrganizationService service) {
            RetrieveEntityRibbonRequest request = new RetrieveEntityRibbonRequest
            {
                EntityName = entityName,
                RibbonLocationFilter = RibbonLocationFilters.All
            };

            RetrieveEntityRibbonResponse response = (RetrieveEntityRibbonResponse)service.Execute(request);
            return response.CompressedEntityXml;
        }
        static void GeneratePdfFile(string entityName, string xmlContent)
        {
            string exePath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            string fileName = "RibbonButtons.pdf";
            string outputPath = Path.Combine(exePath, $"{entityName}_{fileName}");

            int count = 1;
            while (File.Exists(outputPath))
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                string extension = Path.GetExtension(fileName);
                string newFileName = $"{entityName}_{fileNameWithoutExt}_{count}{extension}";
                outputPath = Path.Combine(exePath, newFileName);
                count++;
            }
            Document doc = new Document();
            PdfWriter.GetInstance(doc, new FileStream(outputPath, FileMode.Create));
            doc.Open();
            var headingFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.RED);
            Paragraph heading = new Paragraph($"{entityName} Ribbon Buttons", headingFont);
            heading.Alignment = Element.ALIGN_CENTER;
            doc.Add(heading);
            ParseRibbonXml(xmlContent, doc);
            doc.Close();

            Console.WriteLine($"PDF created at: {outputPath}");
        }
        static void ParseRibbonXml(string xmlContent, Document doc)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlContent);

            XmlNodeList buttons = xmlDoc.GetElementsByTagName("Button");
            foreach (XmlElement button in buttons)
            {
                Paragraph buttonParagraph = new Paragraph();

                string buttonId = button.GetAttribute("Id");
                string command = button.GetAttribute("Command");
                string image16x16 = button.GetAttribute("Image16by16");
                string image32x32 = button.GetAttribute("Image32by32");
                string modernImage = button.GetAttribute("ModernImage");

                var buttonFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, BaseColor.RED);
                buttonParagraph.Add(new Chunk($"\n\nButton: {buttonId}\n", buttonFont));
                buttonParagraph.Add(new Chunk($"   - Icon:\n"));
                buttonParagraph.Add(new Chunk("      - Image 16x16: {image16x16}\n"));
                buttonParagraph.Add(new Chunk($"      - Image 32x32: {image32x32}\n"));
                buttonParagraph.Add(new Chunk($"      - Modern Image: {modernImage}\n"));
                buttonParagraph.Add(new Chunk($"\n   - Command: {command}\n"));

                var commandNode = xmlDoc.SelectSingleNode($"//CommandDefinition[@Id='{command}']");
                if (commandNode != null)
                {
                    var jsFns = commandNode.SelectNodes("Actions/JavaScriptFunction");
                    foreach(XmlNode jsFn in jsFns)
                    {
                        var functionName = jsFn?.Attributes["FunctionName"]?.Value;
                        var library = jsFn?.Attributes["Library"]?.Value;
                        buttonParagraph.Add(new Chunk($"      - FunctionName: {functionName}\n"));
                        buttonParagraph.Add(new Chunk($"      - Library: {library}\n"));
                        buttonParagraph.Add(new Chunk($"      - Function Parameters:\n"));

                        foreach (XmlNode childNode in jsFn.ChildNodes)
                        {
                            buttonParagraph.Add(new Chunk($"          -- {childNode.Name}: {childNode.Attributes["Value"].Value}\n"));
                        }
                    }

                    buttonParagraph.Add(new Chunk($"\n   - Enable Rules:\n"));
                    foreach (XmlNode rule in commandNode.SelectNodes("EnableRules/EnableRule"))
                    {
                        string ruleId = rule.Attributes["Id"].Value;
                        var enableRuleNode = xmlDoc.SelectSingleNode($"//RuleDefinitions//EnableRules//EnableRule[@Id='{ruleId}']");

                        if (enableRuleNode != null)
                        {
                            buttonParagraph.Add(new Chunk($"      - Enable Rule Id: {ruleId}\n"));
                            foreach (XmlNode childNode in enableRuleNode.ChildNodes)
                            {
                                foreach (XmlAttribute attr in childNode.Attributes)
                                {
                                    buttonParagraph.Add(new Chunk($"          -- {attr.Name}: {attr.Value}\n"));
                                }
                            }
                        }
                    }

                    buttonParagraph.Add(new Chunk($"\n   - Display Rules:\n"));
                    foreach (XmlNode rule in commandNode.SelectNodes("DisplayRules/DisplayRule"))
                    {
                        string ruleId = rule.Attributes["Id"].Value;
                        var displayRuleNode = xmlDoc.SelectSingleNode($"//RuleDefinitions//DisplayRules//DisplayRule[@Id='{ruleId}']");

                        if (displayRuleNode != null)
                        {
                            buttonParagraph.Add(new Chunk($"      - Display Rule: {ruleId}\n"));
                            foreach (XmlNode childNode in displayRuleNode.ChildNodes)
                            {
                                foreach (XmlAttribute attr in childNode.Attributes)
                                {
                                    buttonParagraph.Add(new Chunk($"          -- {attr.Name}: {attr.Value}\n"));
                                }
                            }
                        }
                    }
                }

                doc.Add(buttonParagraph);
            }
        }
    }
}

