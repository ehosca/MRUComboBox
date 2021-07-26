//Declare a class that inherits from ToolStripControlHost.

using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace Hosca.Windows.Forms
{
    [ToolStripItemDesignerAvailability(ToolStripItemDesignerAvailability.ToolStrip | ToolStripItemDesignerAvailability.StatusStrip)]
    public class ToolStripMRUComboBox : MRUComboBoxStripControlHost
    {
        // Call the base constructor passing in a MaskedTextBox instance.
        public ToolStripMRUComboBox() : base(CreateControlInstance()) { }

        public MRUComboBox ComboBox => Control as MRUComboBox;

        public ComboBox.ObjectCollection Items => ComboBox.Items;


        private static Control CreateControlInstance()
        {
            MRUComboBox mtb = new MRUComboBox();
            return mtb;
        }
    }
}