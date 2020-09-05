using System.Drawing;
using System.Management.Automation;
using System.Windows.Forms.DataVisualization.Charting;

namespace PSChart
{
	public class ChartSettings
	{

		public bool Enable3D { get; set; }

		public KnownColor[] SeriesColor { get; set; } = {KnownColor.Blue, KnownColor.DarkOrange, KnownColor.Aquamarine, KnownColor.Red};

		public KnownColor BackColor { get; set; } = KnownColor.WhiteSmoke;

		public SeriesChartType ChartType{ get; set; } = SeriesChartType.Bar;

		public int ShadowOffset { get; set; } = 2;

		public bool ShowValueAsLabel { get; set; } = true;

		public bool ShowLegend { get; set; } = true;

		public string LabelFormatString { get; set; }
	}

	[Cmdlet(VerbsCommon.New, "ChartSetting")]
	public class NewChartSetting : PSCmdlet
	{
		[Parameter] public SwitchParameter Enable3D { get; set; }

		[Parameter] public KnownColor[] SeriesColor { get; set; }

		[Parameter] public int ShadowOffset { get; set; }

		[Parameter] public bool NoValueAsLabel { get; set; }

		[Parameter] public bool NoLegend { get; set; }

		[Parameter] public SeriesChartType ChartType { get; set; }

		[Parameter] public KnownColor BackColor { get; set; }

		[Parameter] public string LabelFormatString { get; set; }

		protected override void EndProcessing()
		{
			var s = new ChartSettings();
			foreach (var n in MyInvocation.BoundParameters.Keys)
			{
				switch (n)
				{
					case nameof(Enable3D):
						s.Enable3D = Enable3D;
						continue;
					case nameof(SeriesColor):
						s.SeriesColor = SeriesColor;
						continue;
					case nameof(ShadowOffset):
						s.ShadowOffset = ShadowOffset;
						continue;
					case nameof(NoValueAsLabel):
						s.ShowValueAsLabel = !NoValueAsLabel;
						continue;
					case nameof(ChartType):
						s.ChartType = ChartType;
						continue;
					case nameof(BackColor):
						s.BackColor = BackColor;
						continue;
					case nameof(NoLegend):
						s.ShowLegend = !NoLegend;
						continue;
					case nameof(LabelFormatString):
						s.LabelFormatString = LabelFormatString;
						continue;

				}
			}

			WriteObject(s);
		}
	}
}