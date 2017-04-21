using System;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Windows.Forms.DataVisualization.Charting;
using PSChart;
using Xunit;

namespace PSChartTests
{

	public class OutChartTests
	{
		[Fact]
		public void ExtensionCountLength()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -Property Name, Count -Path e:\\temp\\charts\\chart.png");
				ps.Invoke();
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e=>e.Exception.Message)));
			}
		}

		[Fact]
		public void ExtensionCountLength3D()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -Enable3d  -Path e:\\temp\\charts\\chart3D.png");
				ps.Invoke();
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}

		[Fact]
		public void ExtensionCountLengthCustomColor()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -ChartSettings @{SeriesColor = 'Red'} -Path e:\\temp\\charts\\chartRed.png");
				ps.Invoke();
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}

		[Fact]
		public void ExtensionCountLengthPie()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript("Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -ChartSettings @{ChartType = 'Pie'; ShowValueAsLabel=$false} -Path e:\\temp\\charts\\chartPie.png");
				ps.Invoke();
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}

		[Fact]
		public void ExtensionCountLengthPieBackColor()
		{
			var iss = InitialSessionState.CreateDefault2();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				ps.AddScript(@"
$s = @{BackColor = 'Aqua'}
Get-ChildItem $env:WinDir\\System32 | Group-Object -NoElement Extension | Out-Chart -charttype Column -chartsettings $s -Path e:\\temp\\charts\\chartColumnBackColor.png");
				ps.Invoke();
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}


		[Fact]
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
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}


		[Theory]
		[InlineData(SeriesChartType.Bar)]
		[InlineData(SeriesChartType.Column)]
		[InlineData(SeriesChartType.Line)]
		[InlineData(SeriesChartType.Area)]
		[InlineData(SeriesChartType.SplineArea)]
		[InlineData(SeriesChartType.StackedArea)]
		[InlineData(SeriesChartType.StackedArea100)]
		[InlineData(SeriesChartType.ErrorBar)]
		[InlineData(SeriesChartType.BoxPlot)]
		[InlineData(SeriesChartType.Candlestick)]
		[InlineData(SeriesChartType.ErrorBar)]
		[InlineData(SeriesChartType.Doughnut)]
		[InlineData(SeriesChartType.FastLine)]
		[InlineData(SeriesChartType.FastPoint)]
		[InlineData(SeriesChartType.RangeBar)]
		[InlineData(SeriesChartType.RangeColumn)]

		public void MultiSeriesColumn(SeriesChartType chartType)
		{
			var iss = InitialSessionState.CreateDefault();
			iss.Commands.Add(new SessionStateCmdletEntry("Out-Chart", typeof(OutChartCommand), null));
			using (var ps = PowerShell.Create(iss))
			{
				var script = $@"
$data = gcim Win32_LogicalDisk | where DriveType -eq 3 | select DeviceId, @{{n = 'SizeGB'; ex = {{[int]($_.Size / 1gb)}}}},@{{n = 'FreespaceGB'; ex = {{[int]($_.FreeSpace/ 1gb)}}}}
$data | Out-Chart -property DeviceId,SizeGB,FreeSpaceGB -ChartSettings @{{LabelFormatString ='N0'}} -Path e:\\temp\\charts\\multicolumn{chartType}.png -ChartType {chartType}";
				ps.AddScript(script );
				ps.Invoke();
				Assert.False(ps.HadErrors, string.Join(Environment.NewLine, ps.Streams.Error.ReadAll().Select(e => e.Exception.Message)));
			}
		}
	}
}
