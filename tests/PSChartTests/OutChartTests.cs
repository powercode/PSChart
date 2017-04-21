using System;
using System.Linq;
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
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -Property Name, Count -Path e:\\temp\\charts\\chart.png");
				ps.Invoke();
				Assert.IsFalse(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e=>e.Exception.Message)));
			}
		}

		[TestMethod]
		public void ExtensionCountLength3D()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -ChartSettings @{Enable3D = $true} -Path e:\\temp\\charts\\chart3D.png");
				ps.Invoke();
				Assert.IsFalse(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}

		[TestMethod]
		public void ExtensionCountLengthCustomColor()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -ChartSettings @{SeriesColor = 'Red'} -Path e:\\temp\\charts\\chartRed.png");
				ps.Invoke();
				Assert.IsFalse(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}

		[TestMethod]
		public void ExtensionCountLengthPie()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -ChartSettings @{ChartType = 'Pie'; ShowValueAsLabel=$false} -Path e:\\temp\\charts\\chartPie.png");
				ps.Invoke();
				Assert.IsFalse(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}


		[TestMethod]
		public void ExtensionCountLengthPie3D()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript(@"
$settings = @{ChartType = 'Pie'; ShowValueAsLabel=$false; Enable3d = $true} 
Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -ChartSettings $settings -Path e:\\temp\\charts\\chartPie3d.png");
				ps.Invoke();
				Assert.IsFalse(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}

		[TestMethod]
		public void MultiSeries()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir -File | Select -first 5 | Select-Object Name,Length, @{ 'Name'='Hour';'Expression'={$_.LastAccessTime.TimeOfDay.Hours * 1000}} | Out-Chart -property Name,Length,Hour  -Path e:\\temp\\charts\\multi.png");
				ps.Invoke();
				Assert.IsFalse(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}
	}
}
