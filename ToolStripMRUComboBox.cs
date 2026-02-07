using System.ComponentModel;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace Hosca.Windows.Forms
{
    [ToolStripItemDesignerAvailability(ToolStripItemDesignerAvailability.ToolStrip | ToolStripItemDesignerAvailability.StatusStrip)]
    public class ToolStripMRUComboBox : MRUComboBoxStripControlHost
    {
        // Call the base constructor passing in an MRUComboBox instance.
        public ToolStripMRUComboBox() : base(CreateControlInstance()) { }

        public MRUComboBox ComboBox => Control as MRUComboBox;

        public ComboBox.ObjectCollection Items => ComboBox?.Items;

        [Category("Behavior")]
        [Description("Maximum number of MRU items to retain.")]
        [DefaultValue(10)]
        public int MaxItems
        {
            get => ComboBox?.MaxItems ?? 10;
            set { if (ComboBox != null) ComboBox.MaxItems = value; }
        }

        [Category("Behavior")]
        [Description("Whether duplicate detection is case-sensitive.")]
        [DefaultValue(false)]
        public bool CaseSensitive
        {
            get => ComboBox?.CaseSensitive ?? false;
            set { if (ComboBox != null) ComboBox.CaseSensitive = value; }
        }

        public void AddMRUItem(string item)
        {
            ComboBox?.AddMRUItem(item);
        }

        private static Control CreateControlInstance()
        {
            return new MRUComboBox();
        }
    }
}
