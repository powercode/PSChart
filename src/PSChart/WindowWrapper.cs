using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PSChart
{
	public class WindowWrapper : IWin32Window
	{
		public WindowWrapper(IntPtr handle)
		{
			Handle = handle;
		}

		public IntPtr Handle { get; }
	}
}
