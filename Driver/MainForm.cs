using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Microsoft.SqlServer.Dts.Runtime;

namespace Driver
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
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

		private void splitFileButton_Click(object sender, EventArgs e)
		{
			var path = @"C:\Users\shane.hopcroft\Desktop\FlowTest"; // TODO:
			var outputPath = Path.Combine(path, "output");
			Directory.CreateDirectory(outputPath);
			var originalFileName = Path.Combine(path, "sch-2017-09-01-r1.xml");

			using (var reader = XmlReader.Create(originalFileName))
			{
				while (reader.Read())
				{
					if (reader.NodeType != XmlNodeType.Element)
					{
						continue;
					}

					Console.WriteLine(reader.Name); // TODO:

					if (reader.Name == "info")
					{
						var element = reader.ReadSubtree();

						using (var fileStream = File.OpenWrite(Path.Combine(outputPath, $"{reader.Name}.xml")))
						{
							using (var writer = XmlWriter.Create(fileStream))
							{
								writer.WriteNode(element, true);
							}
						}
					}
					else if (reader.Name == "schedule")
					{
						using (var scheduleReader = XmlReader.Create(originalFileName))
						{
							while (scheduleReader.Read())
							{
								Console.WriteLine("***" + scheduleReader.Name); // TODO:

								if (scheduleReader.NodeType != XmlNodeType.Element)
								{
									continue;
								}

								if (scheduleReader.Name == "program")
								{
									var element = scheduleReader.ReadSubtree();

									var id = scheduleReader.GetAttribute("xml:id");
									using (var fileStream = File.OpenWrite(Path.Combine(outputPath, $"{scheduleReader.Name}_{id}.xml")))
									{
										using (var writer = XmlWriter.Create(fileStream))
										{
											writer.WriteNode(element, true);
										}
									}
								}
								else if (scheduleReader.Name == "values-list")
								{
									var element = scheduleReader.ReadSubtree();

									using (var fileStream = File.OpenWrite(Path.Combine(outputPath, $"{scheduleReader.Name}.xml")))
									{
										using (var writer = XmlWriter.Create(fileStream))
										{
											writer.WriteNode(element, true);
										}
									}
								}
							}
						}
					}
					else if (reader.Name == "drugs-list")
					{
						var element = reader.ReadSubtree();

						using (var fileStream = File.OpenWrite(Path.Combine(outputPath, $"{reader.Name}.xml")))
						{
							using (var writer = XmlWriter.Create(fileStream))
							{
								writer.WriteNode(element, true);
							}
						}
					}
					else if (reader.Name == "organisations-list")
					{
						var element = reader.ReadSubtree();

						using (var fileStream = File.OpenWrite(Path.Combine(outputPath, $"{reader.Name}.xml")))
						{
							using (var writer = XmlWriter.Create(fileStream))
							{
								writer.WriteNode(element, true);
							}
						}
					}
				}
			}

			MessageBox.Show(@"Done!");
		}
	}
}
