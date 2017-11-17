using System;
using System.Collections.Generic;
using log4net;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.SqlServer.Dts.Runtime;
using Excel = Microsoft.Office.Interop.Excel;

//using Excel = Microsoft.Office.Interop.Excel;

namespace Driver
{
	public partial class MainForm : Form
	{
		private readonly ILog _logger;

		public MainForm()
		{
			_logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
			InitializeComponent();
		}

		private void executeButton_Click(object sender, EventArgs e)
		{
			var pkgLocation = @"D:\Code\flow-test\flow-test\Package.dtsx";

			var app = new Microsoft.SqlServer.Dts.Runtime.Application();
			var pkg = app.LoadPackage(pkgLocation, null);
			pkg.Parameters["InputFile"].Value = @"C:\Users\shane.hopcroft\Desktop\FlowTest\input.txt";
			pkg.Parameters["OutputFile"].Value = @"C:\Users\shane.hopcroft\Desktop\FlowTest\output.txt";
			pkg.Parameters["ErrorsFile"].Value = @"C:\Users\shane.hopcroft\Desktop\FlowTest\errors.txt";

			var pkgResults = pkg.Execute();

			if (pkgResults == DTSExecResult.Success)
				Console.WriteLine(@"Package ran successfully");
			else
				Console.WriteLine(@"Package failed");
		}

		private void closeButton_Click(object sender, EventArgs e)
		{
			Close();
		}

		internal string GetWorkingFolder(string inputFile)
		{
			var parentDirectory = Directory.GetParent(inputFile).FullName;
			var workingFolder = Path.Combine(parentDirectory, "Output");

			return workingFolder;
		}

		private void splitFileButton_Click(object sender, EventArgs e)
		{
			//var inputFile = @"C:\Users\shane\Desktop\FlowTest\sch-2017-09-01-r1.xml";
			var inputFile = @"C:\Users\shane.hopcroft\Desktop\FlowTest\sch-2017-09-01-r1.xml";
			var outputFolder = GetWorkingFolder(inputFile);
			Directory.CreateDirectory(outputFolder);

			SplitOutInfoElement(inputFile, outputFolder);
			SplitOutProgramElements(inputFile, outputFolder);
			SplitOutValuesListElement(inputFile, outputFolder);
			SplitOutDrugsListElement(inputFile, outputFolder);
			SplitOutOrganisationsListElement(inputFile, outputFolder);

			MessageBox.Show(@"Done!");
		}

		internal void SplitOutInfoElement(string inputFileName, string outputFolder)
		{
			_logger.Info("Processing info element started");

			using (var reader = XmlReader.Create(inputFileName))
			{
				if (reader.ReadToFollowing("root"))
				{
					if (reader.ReadToDescendant("info"))
					{
						var element = reader.ReadSubtree();
						WriteXmlToFile(Path.Combine(outputFolder, reader.Name + ".xml"), element);
					}
				}
			}

			_logger.Info("Processing info element completed");
		}

		internal void SplitOutProgramElements(string inputFileName, string outputFolder)
		{
			_logger.Info("Processing program elements started");

			using (var reader = XmlReader.Create(inputFileName))
			{
				if (reader.ReadToFollowing("root"))
				{
					if (reader.ReadToDescendant("schedule"))
					{
						if (reader.ReadToDescendant("program"))
						{
							do
							{
								var element = reader.ReadSubtree();
								var programCode = GetProgramCode(ref element);

								// First of all work pull out all the prescribing rules into 
								var prescribingRules = ExtractPrescribingRules(ref element)
									.ToList();

								if (prescribingRules.Count() >= 700)
								{
									// There are two many prescribing-rules to write out to a single file

									// Create a version of the program without any prescribing-rule elements
									var programReader = RemovePrescribingRulesFromProgram(element);
									var strippedProgramDocument = new XmlDocument();
									strippedProgramDocument.Load(programReader);

									const int passSize = 700;
									for (var pass = 0; pass * passSize < prescribingRules.Count; pass++)
									{
										var programDocument = (XmlDocument)strippedProgramDocument.CloneNode(true);
										var root = programDocument.DocumentElement;

										var currentPass = prescribingRules
											.Skip(pass * passSize)
											.Take(passSize);

										foreach (var prescribingRule in currentPass)
										{
											var newElement = programDocument.CreateElement("prescribing-rule");
											newElement.InnerXml = prescribingRule.InnerXml;
											root?.InsertAfter(newElement, root.LastChild);
										}

										programDocument.Save(Path.Combine(outputFolder, "Program_" + programCode + "_" + (pass + 1) + ".xml"));
									}
								}
								else
								{
									// The file is small enought that we can write it out verbatim
									WriteXmlToFile(Path.Combine(outputFolder, reader.Name + "_" + programCode + ".xml"), element);
								}
							} while (reader.ReadToNextSibling("program"));
						}
					}
				}
			}

			_logger.Info("Processing program elements completed");
		}

		private string GetProgramCode(ref XmlReader programReader)
		{
			// https://pmouse.wordpress.com/2008/11/24/seeking-with-xmlreader-and-linq-to-xml/

			var xElement = XElement.Load(programReader);
			programReader.Close();
			programReader = xElement.CreateReader();

			string programCode = "-";

			// Here you can peek ahead
			if (programReader.ReadToFollowing("info"))
			{
				if (programReader.ReadToDescendant("code"))
				{
					programReader.Read();
					programCode = programReader.Value;
				}
			}

			// and here the reader gets reset to the beginning
			programReader.Close();
			programReader = xElement.CreateReader();

			return programCode;
		}

		internal IEnumerable<XmlDocument> ExtractPrescribingRules(ref XmlReader programReader)
		{
			var xElement = XElement.Load(programReader);
			programReader.Close();
			programReader = xElement.CreateReader();

			var prescribingRuleList = new List<XmlDocument>();

			if (programReader.ReadToFollowing("program"))
			{
				if (programReader.ReadToFollowing("prescribing-rule"))
				{
					do
					{
						var prescribingRule = programReader.ReadSubtree();

						var prescribingRuleDocument = new XmlDocument();
						prescribingRuleDocument.Load(prescribingRule);

						prescribingRuleList.Add(prescribingRuleDocument);

					} while (programReader.ReadToNextSibling("prescribing-rule"));
				}
			}

			// and here the reader gets reset to the beginning
			programReader.Close();
			programReader = xElement.CreateReader();

			return prescribingRuleList;
		}

		private XmlReader RemovePrescribingRulesFromProgram(XmlReader programReader)
		{
			var stream = new MemoryStream();
			using (var writer = XmlWriter.Create(stream))
			{
				while (programReader.Read())
				{
					if (programReader.Name == "prescribing-rule")
					{
						programReader.Skip();
					}
					else
					{
						WriteShallowNode(programReader, writer);
					}
				}
			}

			stream.Seek(0, SeekOrigin.Begin); // Reset stream position to read from the beginning

			var reader = XmlReader.Create(stream);
			return reader;
		}

		internal void SplitOutValuesListElement(string inputFileName, string outputFolder)
		{
			_logger.Info("Processing values-list element started");

			using (var reader = XmlReader.Create(inputFileName))
			{
				if (reader.ReadToFollowing("root"))
				{
					if (reader.ReadToDescendant("schedule"))
					{
						if (reader.ReadToDescendant("values-list"))
						{
							var element = reader.ReadSubtree();
							WriteXmlToFile(Path.Combine(outputFolder, reader.Name + ".xml"), element);
						}
					}
				}
			}

			_logger.Info("Processing values-list element completed");
		}

		internal void SplitOutDrugsListElement(string inputFileName, string outputFolder)
		{
			_logger.Info("Processing drugs-list element started");

			using (var reader = XmlReader.Create(inputFileName))
			{
				if (reader.ReadToFollowing("root"))
				{
					if (reader.ReadToDescendant("drugs-list"))
					{
						var element = reader.ReadSubtree();
						WriteXmlToFile(Path.Combine(outputFolder, reader.Name + ".xml"), element);
					}
				}
			}

			_logger.Info("Processing drugs-list element completed");
		}

		internal void SplitOutOrganisationsListElement(string inputFileName, string outputFolder)
		{
			_logger.Info("Processing organisations-list element started");

			using (var reader = XmlReader.Create(inputFileName))
			{
				if (reader.ReadToFollowing("root"))
				{
					if (reader.ReadToDescendant("organisations-list"))
					{
						var element = reader.ReadSubtree();
						WriteXmlToFile(Path.Combine(outputFolder, reader.Name + ".xml"), element);
					}
				}
			}

			_logger.Info("Processing organisations-list element completed");
		}

		internal void WriteXmlToFile(string fileName, XmlReader element)
		{
			using (var fileStream = File.Open(fileName, FileMode.Create))
			{
				using (var writer = XmlWriter.Create(fileStream))
				{
					writer.WriteNode(element, true);
				}
			}
		}

		private void xmlToCsvButton_Click(object sender, EventArgs e)
		{
			//var xmlFileName = @"C:\Users\shane\Desktop\FlowTest\Output\organisations-list.xml";
			var xmlFileName = @"C:\Users\shane\Desktop\FlowTest\Output\program_CA.xml";
			//var xmlFileName = @"C:\Users\shane\Desktop\FlowTest\Output\program_GE_1.xml";
			var targetFileName = Path.ChangeExtension(xmlFileName, "csv");

			var excelApplication = new Excel.Application
			{
				DisplayAlerts = false,
			};
			var excelWorkBook = excelApplication.Workbooks.OpenXML(xmlFileName, Type.Missing, Excel.XlXmlLoadOption.xlXmlLoadOpenXml);

			excelWorkBook.SaveAs(targetFileName, Excel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

			excelWorkBook.Close();
			excelApplication.Workbooks.Close();

			MessageBox.Show(@"Done!");
		}

		//This method is useful for streaming transformation with the
		//XmlReader and the XmlWriter. It pushes through single nodes in the stream
		static void WriteShallowNode(XmlReader reader, XmlWriter writer)
		{
			if (reader == null)
			{
				throw new ArgumentNullException(nameof(reader));
			}
			if (writer == null)
			{
				throw new ArgumentNullException(nameof(writer));
			}

			switch (reader.NodeType)
			{
				case XmlNodeType.Element:
					writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
					writer.WriteAttributes(reader, true);
					if (reader.IsEmptyElement)
					{
						writer.WriteEndElement();
					}
					break;
				case XmlNodeType.Text:
					writer.WriteString(reader.Value);
					break;
				case XmlNodeType.Whitespace:
				case XmlNodeType.SignificantWhitespace:
					writer.WriteWhitespace(reader.Value);
					break;
				case XmlNodeType.CDATA:
					writer.WriteCData(reader.Value);
					break;
				case XmlNodeType.EntityReference:
					writer.WriteEntityRef(reader.Name);
					break;
				case XmlNodeType.XmlDeclaration:
				case XmlNodeType.ProcessingInstruction:
					writer.WriteProcessingInstruction(reader.Name, reader.Value);
					break;
				case XmlNodeType.DocumentType:
					// ReSharper disable once AssignNullToNotNullAttribute
					writer.WriteDocType(reader.Name, reader.GetAttribute("PUBLIC"), reader.GetAttribute("SYSTEM"), reader.Value);
					break;
				case XmlNodeType.Comment:
					writer.WriteComment(reader.Value);
					break;
				case XmlNodeType.EndElement:
					writer.WriteFullEndElement();
					break;
			}
		}
	}
}
