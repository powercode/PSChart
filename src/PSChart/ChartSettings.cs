using System.Drawing;
using System.Management.Automation;
using System.Windows.Forms.DataVisualization.Charting;

namespace PSChart
{
	public class ChartSettings
	{
		public bool Enable3D { get; set; }

		public Color[] SeriesColor { get; set; } = {Color.Blue, Color.DarkOrange, Color.Aquamarine, Color.Red};

		public SeriesChartType ChartType{ get; set; } = SeriesChartType.Bar;

		public int ShadowOffset { get; set; } = 2;

		public bool ShowValueAsLabel { get; set; } = true;

		public bool ShowLegend { get; set; } = true;
	}

	[Cmdlet(VerbsCommon.New, "ChartSetting")]
	public class NewChartSetting : Cmdlet
	{
		[Parameter] public SwitchParameter Enable3D { get; set; }

		[Parameter] public Color[] SeriesColor { get; set; }

		[Parameter] public int ShadowOffset { get; set; }

		[Parameter] public bool NoValueAsLabel { get; set; }

		[Parameter] public bool NoLegend { get; set; }

		[Parameter] public SeriesChartType ChartType { get; set; }

		protected override void EndProcessing()
		{
			var s = new ChartSettings
			{
				Enable3D = Enable3D,
				SeriesColor = SeriesColor,
				ShadowOffset = ShadowOffset,
				ShowValueAsLabel = !NoValueAsLabel,
				ShowLegend = !NoLegend,
				ChartType = ChartType
			};

			WriteObject(s);
		}
	}
}