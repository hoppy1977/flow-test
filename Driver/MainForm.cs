using System;
using System.Windows.Forms;
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
	}
}
