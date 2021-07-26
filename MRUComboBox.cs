using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Hosca.Windows.Forms.Properties;

namespace Hosca.Windows.Forms
{
    public class MRUComboBox : ComboBox
    {
        private readonly Dictionary<int, Rectangle> _deleteRectangles = new Dictionary<int, Rectangle>();

        public MRUComboBox()
        {
            DropDownStyle = ComboBoxStyle.DropDown;
            DrawMode = DrawMode.OwnerDrawVariable;
            DrawItem += HandleMruComboBoxDrawItem;
        }

        private DropdownListBoxWindow ListBoxWindow { get; set; }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            var hWind = GetComboBoxListInternal(Handle);
            ListBoxWindow = new DropdownListBoxWindow(hWind);
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
                    Items.RemoveAt(i);
            }
            
            Focus();

        }

        internal static IntPtr GetComboBoxListInternal(IntPtr cboHandle)
        {
            var cbInfo = new COMBOBOXINFO {cbSize = Marshal.SizeOf<COMBOBOXINFO>()};
            GetComboBoxInfo(cboHandle, ref cbInfo);
            return cbInfo.hwndList;
        }

        private void HandleMruComboBoxDrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            var itemText = (string) Items[e.Index];

            var itemBounds = e.Bounds;
            itemBounds.Width -= itemBounds.Height;

            using (var sb = new SolidBrush(e.ForeColor))
            {
                e.Graphics.DrawString(itemText, e.Font, sb, itemBounds);
            }

            var deleteRectangle = new Rectangle(itemBounds.Width,
                itemBounds.Y + 2, 
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
                        var p = new Point((int) m.LParam & 0xFFFF, ((int) m.LParam >> 16) & 0xFFFF);
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