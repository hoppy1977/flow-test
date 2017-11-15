using System;
using log4net;
using System.IO;
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
			var inputFile = @"C:\Users\shane\Desktop\FlowTest\sch-2017-09-01-r1.xml";
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

								WriteXmlToFile(Path.Combine(outputFolder, reader.Name + "_" + programCode + ".xml"), element);
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
			using (var fileStream = File.OpenWrite(fileName))
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
	}
}
