using System;
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
        [ValidateCount(2, 5)]
		public string[] Property
		{
			set
			{
				XLabel = value[0];
				YLabel = new string[value.Length - 1];
				Array.Copy(value, 1, YLabel,0, YLabel.Length);
			}
		}

		[Parameter]
		public ChartSettings ChartSettings { get; set; } = new ChartSettings();

		[Parameter]
		public string XLabel { get; set; } = "Name";

		[Parameter]
		public string[] YLabel { get; set; } = {"Count" };


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
        public string Path { get; set; }

		[Parameter]
        public SwitchParameter ForceForm { get; set; }

        [Parameter]
        public SwitchParameter Enable3D { get; set; }

        [Parameter]
        [ValidateCount(1, 5)]
        public KnownColor[] SeriesColor{ get; set; }

        [Parameter]
        public SeriesChartType ChartType{ get; set; }


        [Parameter]
		public SwitchParameter Clipboard { get; set; }

		private readonly Dictionary<string, Series> _data = new Dictionary<string, Series>(10);

		protected override void BeginProcessing()
		{

			foreach (var y in YLabel)
			{
				var s = new Series(y) {Legend = "Legend"};
				_data.Add(y, s);
			}
            if (MyInvocation.BoundParameters.ContainsKey(nameof(SeriesColor)))
            {
	            ChartSettings.SeriesColor = SeriesColor;
            }
            if (Enable3D)
            {
                ChartSettings.Enable3D = true;
            }
            if (MyInvocation.BoundParameters.ContainsKey(nameof(ChartType)))
            {
                ChartSettings.ChartType = ChartType;
            }

	}

		protected override void ProcessRecord()
		{
			if (ParameterSetName == "Property")
			{
				var name = XLabel;
				foreach (var input in InputObject)
				{
					var props = input.Properties;
					var x = LanguagePrimitives.ConvertTo<string>(props[XLabel].Value);
					foreach (var yVal in YLabel)
					{
						var serie = _data[yVal];
						switch (props[yVal].Value)
						{
							case double d: serie.Points.AddXY(x, d); break;
							case float f: serie.Points.AddXY(x, f); break;
							case ulong ul: serie.Points.AddXY(x, ul); break;
							case long l: serie.Points.AddXY(x, l); break;
							case uint ui: serie.Points.AddXY(x, ui); break;
							case int i: serie.Points.AddXY(x, i); break;
							default:
								var yValue = LanguagePrimitives.ConvertTo<double>(props[yVal].Value);
								serie.Points.AddXY(x, yValue);
								break;
						}

						//serie.Points[0].LegendText = yVal.Length == 0 ? "<none>" : yVal;
					}
				}
			}
			else
			{
				foreach (var gi in GroupInfo)
				{
					var serie = _data["Count"];
					serie.Points.AddXY(gi.Name, gi.Count);
					serie.Points[0].LegendText = gi.Name.Length == 0 ? "<none>" : gi.Name;
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
			var legend = new Legend
			{
				LegendStyle = LegendStyle.Column,
				Name = "Legend",
				Enabled = ChartSettings.ShowLegend
			};
			chart.Legends.Add(legend);

			var chartArea = new ChartArea("Data")
			{
				AxisX =
				{
					Title = XLabel,
					Interval = 1,
					MajorGrid = {Enabled = false},
					MinorGrid = {Enabled = false}

				},
				AxisY = {Title = YLabel.Length == 1 ? YLabel[0] : "Values"},

				Area3DStyle =  new ChartArea3DStyle{Enable3D = ChartSettings.Enable3D },
				BackColor = Color.FromKnownColor(ChartSettings.BackColor),
			};
			chart.ChartAreas.Add(chartArea);

			chart.Titles.Add(Title);
			int i = 0;
			foreach (var s in _data.Values)
			{
				chart.Series.Add(s);
				s.IsValueShownAsLabel = ChartSettings.ShowValueAsLabel;
				s.SmartLabelStyle.Enabled = true;
				s.SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Yes;
				var colorIndex = Math.Min(i, ChartSettings.SeriesColor.Length - 1);
				s.Color = Color.FromKnownColor(ChartSettings.SeriesColor[colorIndex]);
				s.ShadowColor = Color.FromArgb(0xFF, 0x20, 0x20, 0x20);
				s.ShadowOffset = ChartSettings.ShadowOffset;
				s.ChartType = ChartSettings.ChartType;
				s.LabelFormat = ChartSettings.LabelFormatString;

				i++;
			}


			if (Clipboard)
			{
				var memoryStream = new MemoryStream(5 * Width*Height);
				chart.SaveImage(memoryStream, ChartImageFormat.Bmp);
				var img = new Bitmap(memoryStream);
				System.Windows.Forms.Clipboard.SetImage(img);
			}

			var writeChartToDisk = !string.IsNullOrEmpty(Path);

			if (ForceForm || (!Clipboard && !writeChartToDisk))
			{
				var form = new Form
				{
					Text = "PowerChart",
					Width = chart.Width + 100,
					Height = chart.Height + 100
				};
				using (form)
				{
					form.Controls.Add(chart);
					form.Shown += (sender, args) => ((Form) sender).Activate();
					var win = new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle);
					chart.Invalidate();

					form.ShowDialog(win);
				}
			}
			if(writeChartToDisk)
			{
				var outPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
				var dir = System.IO.Path.GetDirectoryName(outPath);
				if (!Directory.Exists(dir))
				{
					Directory.CreateDirectory(dir);
				}
				chart.SaveImage(outPath, ChartImageFormat.Png);
				var fi = new FileInfo(outPath);
				WriteObject(fi);
			}
		}
	}
}