using System.Windows.Forms;

namespace System.Windows.Forms
{
    public class MRUComboBoxStripControlHost : ToolStripControlHost
    {
        public MRUComboBoxStripControlHost()
            : base(new Control())
        {
        }
        public MRUComboBoxStripControlHost(Control c)
            : base(c)
        {
        }
    }
}
