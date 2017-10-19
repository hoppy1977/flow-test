using System;
using System.Windows.Forms;
//using Microsoft.SqlServer;
using Microsoft.SqlServer.Dts.Runtime;
using Application = Microsoft.SqlServer.Dts.Runtime.Application;

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
			string pkgLocation;
			Package pkg;
			Application app;
			DTSExecResult pkgResults;

			pkgLocation = @"D:\Code\flow-test\flow-test\Package.dtsx";

			app = new Application();
			pkg = app.LoadPackage(pkgLocation, null);

			//Variables vars;
			//vars = pkg.Variables;
			//vars["A_Variable"].Value = "Some value";

			//pkgResults = pkg.Execute(null, vars, null, null, null);
			pkgResults = pkg.Execute();

			if (pkgResults == DTSExecResult.Success)
				Console.WriteLine("Package ran successfully");
			else
				Console.WriteLine("Package failed");
		}

		private void closeButton_Click(object sender, EventArgs e)
		{
			Close();
		}
	}
}
