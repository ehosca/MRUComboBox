using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Hosca.Windows.Forms.Properties;

namespace Hosca.Windows.Forms
{
    public class MRUItemEventArgs : EventArgs
    {
        public MRUItemEventArgs(string item)
        {
            Item = item;
        }

        public string Item { get; }
    }

    public class MRUComboBox : ComboBox
    {
        private const int DeleteIconVerticalPadding = 2;
        private const int DefaultMaxItems = 10;

        private readonly Dictionary<int, Rectangle> _deleteRectangles = new Dictionary<int, Rectangle>();
        private bool _suppressPromote;

        public MRUComboBox()
        {
            DropDownStyle = ComboBoxStyle.DropDown;
            DrawMode = DrawMode.OwnerDrawVariable;
            DrawItem += HandleMruComboBoxDrawItem;
            MaxItems = DefaultMaxItems;
            CaseSensitive = false;
        }

        [Category("Behavior")]
        [Description("Maximum number of MRU items to retain.")]
        [DefaultValue(DefaultMaxItems)]
        public int MaxItems { get; set; }

        [Category("Behavior")]
        [Description("Whether duplicate detection is case-sensitive.")]
        [DefaultValue(false)]
        public bool CaseSensitive { get; set; }

        public event EventHandler<MRUItemEventArgs> ItemDeleted;

        public event EventHandler<MRUItemEventArgs> ItemAdded;

        private DropdownListBoxWindow ListBoxWindow { get; set; }

        public void AddMRUItem(string item)
        {
            if (string.IsNullOrEmpty(item))
                return;

            var comparison = CaseSensitive
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            _suppressPromote = true;
            try
            {
                for (int i = Items.Count - 1; i >= 0; i--)
                {
                    if (string.Equals(Items[i]?.ToString(), item, comparison))
                    {
                        Items.RemoveAt(i);
                    }
                }

                Items.Insert(0, item);

                while (Items.Count > MaxItems && MaxItems > 0)
                {
                    Items.RemoveAt(Items.Count - 1);
                }

                Text = item;
            }
            finally
            {
                _suppressPromote = false;
            }

            ItemAdded?.Invoke(this, new MRUItemEventArgs(item));
        }

        protected override void OnSelectedIndexChanged(EventArgs e)
        {
            base.OnSelectedIndexChanged(e);

            if (_suppressPromote || SelectedIndex <= 0 || SelectedIndex >= Items.Count)
                return;

            var selectedItem = Items[SelectedIndex]?.ToString();
            if (selectedItem == null)
                return;

            _suppressPromote = true;
            try
            {
                Items.RemoveAt(SelectedIndex);
                Items.Insert(0, selectedItem);
                SelectedIndex = 0;
            }
            finally
            {
                _suppressPromote = false;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            var hWnd = GetComboBoxListInternal(Handle);
            if (hWnd == IntPtr.Zero)
                return;

            ListBoxWindow = new DropdownListBoxWindow(hWnd);
            ListBoxWindow.MouseUp += HandleListBoxWindowMouseUp;
        }

        private void HandleListBoxWindowMouseUp(object sender, MouseEventArgs e)
        {
            var indicesToRemove =
                (from kv in _deleteRectangles where kv.Value.Contains(e.Location) select kv.Key).ToList();

            indicesToRemove.Sort((a, b) => b.CompareTo(a));

            if (indicesToRemove.Any())
            {
                foreach (var i in indicesToRemove)
                {
                    var itemText = Items[i]?.ToString();
                    Items.RemoveAt(i);
                    _deleteRectangles.Remove(i);
                    if (itemText != null)
                        ItemDeleted?.Invoke(this, new MRUItemEventArgs(itemText));
                }

                Focus();
            }
        }

        internal static IntPtr GetComboBoxListInternal(IntPtr cboHandle)
        {
            var cbInfo = new COMBOBOXINFO {cbSize = Marshal.SizeOf<COMBOBOXINFO>()};
            if (!GetComboBoxInfo(cboHandle, ref cbInfo))
                return IntPtr.Zero;
            return cbInfo.hwndList;
        }

        private void HandleMruComboBoxDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count)
                return;

            e.DrawBackground();

            var itemText = Items[e.Index]?.ToString() ?? string.Empty;

            var itemBounds = e.Bounds;
            itemBounds.Width -= itemBounds.Height;

            using (var sb = new SolidBrush(e.ForeColor))
            {
                e.Graphics.DrawString(itemText, e.Font, sb, itemBounds);
            }

            var deleteRectangle = new Rectangle(itemBounds.Width,
                itemBounds.Y + DeleteIconVerticalPadding,
                Resources.DeleteIcon.Width,
                Resources.DeleteIcon.Height);

            e.Graphics.DrawIcon(Resources.DeleteIcon, deleteRectangle);

            _deleteRectangles[e.Index] = deleteRectangle;
        }

        private class DropdownListBoxWindow : NativeWindow
        {
            public DropdownListBoxWindow(IntPtr hWnd)
            {
                AssignHandle(hWnd);
            }

            public event MouseEventHandler MouseUp;

            protected override void WndProc(ref Message m)
            {
                switch (m.Msg)
                {
                    case WM_LBUTTONUP:
                    {
                        var lParam = m.LParam.ToInt64();
                        var p = new Point((int)(lParam & 0xFFFF), (int)((lParam >> 16) & 0xFFFF));
                        MouseUp?.Invoke(null, new MouseEventArgs(MouseButtons.Left, 1, p.X, p.Y, 0));
                        break;
                    }
                }

                base.WndProc(ref m);
            }
        }

        #region Win32 Definitions

        private const int WM_LBUTTONUP = 0x0202;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern bool GetComboBoxInfo(IntPtr hWnd, ref COMBOBOXINFO pcbi);

        [StructLayout(LayoutKind.Sequential)]
        internal struct COMBOBOXINFO
        {
            public int cbSize;
            public Rectangle rcItem;
            public Rectangle rcButton;
            public int buttonState;
            public IntPtr hwndCombo;
            public IntPtr hwndEdit;
            public IntPtr hwndList;
        }

        #endregion
    }
}
