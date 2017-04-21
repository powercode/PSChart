using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Management.Automation;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PSChart
{
	[Cmdlet(VerbsData.Out, "Chart", DefaultParameterSetName = "group")]
	[OutputType(typeof(FileInfo))]
	public class OutChartCommand : PSCmdlet
	{
		[Parameter(ValueFromPipeline = true, Mandatory = true)]
		public PSObject[] InputObject { get; set; }

		[Parameter(Mandatory = true, ParameterSetName = "Property")]
		[ValidateCount(2, 2)]
		public string[] Property
		{
			get => new[] {XLabel, YLabel};
			set
			{
				XLabel = value[0];
				YLabel = value[1];
			}
		}

		[Parameter]
		public string XLabel { get; set; } = "Name";

		[Parameter]
		public string YLabel { get; set; } = "Count";


		[Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "group")]
		public Microsoft.PowerShell.Commands.GroupInfo[] GroupInfo { get; set; }

		[Parameter]
		[ValidateRange(200, 2000)]
		public int Width { get; set; } = 1200;

		[Parameter]
		[ValidateRange(200, 2000)]
		public int Height { get; set; } = 800;

		[Parameter]
		public string Title { get; set; } = "PowerChart";


		[Parameter]
		public string OutPath { get; set; }

		[Parameter]
		public SwitchParameter Clipboard { get; set; }


		private readonly Dictionary<string, int> _data = new Dictionary<string, int>(1000);


		protected override void ProcessRecord()
		{
			if (ParameterSetName == "Property")
			{
				var name = Property[0];
				var val = Property[1];
				foreach (var input in InputObject)
				{
					var props = input.Properties;
					_data.Add(LanguagePrimitives.ConvertTo<string>( props[name].Value), LanguagePrimitives.ConvertTo<int>(props[val].Value));
				}
			}
			else
			{
				foreach (var gi in GroupInfo)
				{
					_data.Add(gi.Name, gi.Count);
				}
			}

		}

		protected override void EndProcessing()
		{
			Application.EnableVisualStyles();
			Chart chart = new Chart
			{
				Text = Title,
				Width = Width,
				Height = Height,
				Anchor =  AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
				BackColor = Color.WhiteSmoke
			};
			var chartArea = new ChartArea("Data")
			{
				AxisX =
				{
					Title = XLabel,
					Interval = 1,
					MajorGrid = {Enabled = false},
					MinorGrid = {Enabled = false}

				},
				AxisY = {Title = YLabel}
			};
			chart.ChartAreas.Add(chartArea);

			chart.Titles.Add(Title);

			var data = chart.Series.Add("Data");
			data.IsValueShownAsLabel = true;
			data.ChartType = SeriesChartType.Bar;

			data.Points.DataBindXY(_data.Keys, _data.Values);

			if (Clipboard)
			{
				var memoryStream = new MemoryStream(4 * 1024*1025);
				chart.SaveImage(memoryStream, ChartImageFormat.Bmp);
				var img = new Bitmap(memoryStream);
				System.Windows.Forms.Clipboard.SetImage(img);
			}

			if (string.IsNullOrEmpty(OutPath))
			{
				var form = new Form
				{
					Text = "PowerChart",
					Width = chart.Width + 100,
					Height = chart.Height + 100
				};
				using (form)
				{
					chart.BackColor = Color.Transparent;
					form.Controls.Add(chart);
					form.Shown += (sender, args) => ((Form) sender).Activate();
					var win = new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle);
					chart.Invalidate();

					form.ShowDialog(win);
				}
			}
			else
			{
				var outPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutPath);
				var dir = Path.GetDirectoryName(outPath);
				if (!Directory.Exists(dir))
				{
					Directory.CreateDirectory(dir);
				}
				chart.BackColor = Color.WhiteSmoke;
				chart.SaveImage(outPath, ChartImageFormat.Png);
				var fi = new FileInfo(outPath);
				WriteObject(fi);
			}
		}
	}
}