using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PSChart;

namespace PSChartTests
{
	[TestClass]
	public class OutChartTests
	{
		[TestMethod]
		public void ExtensionCountLength()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir | Group-Object -NoElement Extension | Out-Chart -Property Name, Count");
				ps.Invoke();
			}
		}
	}
}
