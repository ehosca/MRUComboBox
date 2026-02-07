using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Hosca.Windows.Forms;
using NUnit.Framework;

namespace MRUComboBox.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class MRUComboBoxTests
    {
        private Hosca.Windows.Forms.MRUComboBox _comboBox;

        [SetUp]
        public void SetUp()
        {
            _comboBox = new Hosca.Windows.Forms.MRUComboBox();
        }

        [TearDown]
        public void TearDown()
        {
            _comboBox?.Dispose();
        }

        #region Constructor Tests

        [Test]
        public void Constructor_SetsDropDownStyleToDropDown()
        {
            Assert.AreEqual(ComboBoxStyle.DropDown, _comboBox.DropDownStyle);
        }

        [Test]
        public void Constructor_SetsDrawModeToOwnerDrawVariable()
        {
            Assert.AreEqual(DrawMode.OwnerDrawVariable, _comboBox.DrawMode);
        }

        [Test]
        public void Constructor_InitializesWithNoItems()
        {
            Assert.AreEqual(0, _comboBox.Items.Count);
        }

        [Test]
        public void Constructor_SetsMaxItemsToDefault()
        {
            Assert.AreEqual(10, _comboBox.MaxItems);
        }

        [Test]
        public void Constructor_SetsCaseSensitiveToFalse()
        {
            Assert.IsFalse(_comboBox.CaseSensitive);
        }

        #endregion

        #region Item Management Tests

        [Test]
        public void AddItem_SingleItem_CountIsOne()
        {
            _comboBox.Items.Add("Item1");
            Assert.AreEqual(1, _comboBox.Items.Count);
        }

        [Test]
        public void AddItem_MultipleItems_CountMatches()
        {
            _comboBox.Items.Add("Item1");
            _comboBox.Items.Add("Item2");
            _comboBox.Items.Add("Item3");
            Assert.AreEqual(3, _comboBox.Items.Count);
        }

        [Test]
        public void AddItem_PreservesInsertionOrder()
        {
            _comboBox.Items.Add("Alpha");
            _comboBox.Items.Add("Beta");
            _comboBox.Items.Add("Gamma");

            Assert.AreEqual("Alpha", _comboBox.Items[0]);
            Assert.AreEqual("Beta", _comboBox.Items[1]);
            Assert.AreEqual("Gamma", _comboBox.Items[2]);
        }

        [Test]
        public void RemoveItem_DecreasesCount()
        {
            _comboBox.Items.Add("Item1");
            _comboBox.Items.Add("Item2");
            _comboBox.Items.RemoveAt(0);

            Assert.AreEqual(1, _comboBox.Items.Count);
        }

        [Test]
        public void RemoveItem_RemovesCorrectItem()
        {
            _comboBox.Items.Add("Item1");
            _comboBox.Items.Add("Item2");
            _comboBox.Items.Add("Item3");
            _comboBox.Items.RemoveAt(1);

            Assert.AreEqual("Item1", _comboBox.Items[0]);
            Assert.AreEqual("Item3", _comboBox.Items[1]);
        }

        [Test]
        public void RemoveItem_LastItem_ListBecomesEmpty()
        {
            _comboBox.Items.Add("OnlyItem");
            _comboBox.Items.RemoveAt(0);
            Assert.AreEqual(0, _comboBox.Items.Count);
        }

        [Test]
        public void ClearItems_RemovesAllItems()
        {
            _comboBox.Items.Add("Item1");
            _comboBox.Items.Add("Item2");
            _comboBox.Items.Clear();

            Assert.AreEqual(0, _comboBox.Items.Count);
        }

        [Test]
        public void AddItem_DuplicateItems_BothAdded()
        {
            _comboBox.Items.Add("Duplicate");
            _comboBox.Items.Add("Duplicate");

            Assert.AreEqual(2, _comboBox.Items.Count);
        }

        [Test]
        public void InsertItem_AtSpecificIndex_PlacedCorrectly()
        {
            _comboBox.Items.Add("First");
            _comboBox.Items.Add("Third");
            _comboBox.Items.Insert(1, "Second");

            Assert.AreEqual("First", _comboBox.Items[0]);
            Assert.AreEqual("Second", _comboBox.Items[1]);
            Assert.AreEqual("Third", _comboBox.Items[2]);
        }

        [Test]
        public void AddRange_AddsAllItems()
        {
            _comboBox.Items.AddRange(new object[] { "A", "B", "C", "D" });
            Assert.AreEqual(4, _comboBox.Items.Count);
        }

        [Test]
        public void Items_ContainsCheck_ReturnsTrue()
        {
            _comboBox.Items.Add("FindMe");
            Assert.IsTrue(_comboBox.Items.Contains("FindMe"));
        }

        [Test]
        public void Items_ContainsCheck_ReturnsFalse()
        {
            _comboBox.Items.Add("NotThis");
            Assert.IsFalse(_comboBox.Items.Contains("Other"));
        }

        [Test]
        public void Items_IndexOf_ReturnsCorrectIndex()
        {
            _comboBox.Items.Add("Zero");
            _comboBox.Items.Add("One");
            _comboBox.Items.Add("Two");

            Assert.AreEqual(1, _comboBox.Items.IndexOf("One"));
        }

        [Test]
        public void Items_IndexOf_NotFound_ReturnsNegativeOne()
        {
            _comboBox.Items.Add("Exists");
            Assert.AreEqual(-1, _comboBox.Items.IndexOf("Missing"));
        }

        #endregion

        #region Selection Tests

        [Test]
        public void SelectedIndex_DefaultIsNegativeOne()
        {
            Assert.AreEqual(-1, _comboBox.SelectedIndex);
        }

        [Test]
        public void SelectedIndex_SetValid_AutoPromotesToTop()
        {
            _comboBox.Items.Add("Item1");
            _comboBox.Items.Add("Item2");
            _comboBox.SelectedIndex = 1;

            // Auto-promote moves selected item to index 0
            Assert.AreEqual(0, _comboBox.SelectedIndex);
            Assert.AreEqual("Item2", _comboBox.Items[0]);
            Assert.AreEqual("Item1", _comboBox.Items[1]);
        }

        [Test]
        public void SelectedItem_ReturnsCorrectItem()
        {
            _comboBox.Items.Add("Item1");
            _comboBox.Items.Add("Item2");
            _comboBox.SelectedIndex = 0;

            Assert.AreEqual("Item1", _comboBox.SelectedItem);
        }

        [Test]
        public void Text_CanBeSetDirectly()
        {
            _comboBox.Text = "CustomText";
            Assert.AreEqual("CustomText", _comboBox.Text);
        }

        #endregion

        #region Delete Rectangle Tracking Tests

        [Test]
        public void DeleteRectangles_InitiallyEmpty()
        {
            var field = typeof(Hosca.Windows.Forms.MRUComboBox)
                .GetField("_deleteRectangles", BindingFlags.NonPublic | BindingFlags.Instance);
            var dict = (Dictionary<int, Rectangle>)field.GetValue(_comboBox);

            Assert.AreEqual(0, dict.Count);
        }

        #endregion

        #region Property Tests

        [Test]
        public void DropDownStyle_CannotBeChangedToDropDownList_StillAllowed()
        {
            _comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            Assert.AreEqual(ComboBoxStyle.DropDownList, _comboBox.DropDownStyle);
        }

        [Test]
        public void DrawMode_IsOwnerDrawVariable()
        {
            Assert.AreEqual(DrawMode.OwnerDrawVariable, _comboBox.DrawMode);
        }

        #endregion

        #region Edge Case Tests

        [Test]
        public void AddItem_NullItem_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => _comboBox.Items.Add(null));
        }

        [Test]
        public void AddItem_EmptyString_DoesNotThrow()
        {
            Assert.DoesNotThrow(() => _comboBox.Items.Add(""));
            Assert.AreEqual(1, _comboBox.Items.Count);
        }

        [Test]
        public void AddItem_NonStringObject_Allowed()
        {
            _comboBox.Items.Add(42);
            Assert.AreEqual(1, _comboBox.Items.Count);
        }

        [Test]
        public void RemoveAt_InvalidIndex_ThrowsArgumentOutOfRange()
        {
            _comboBox.Items.Add("Item");
            Assert.Throws<ArgumentOutOfRangeException>(() => _comboBox.Items.RemoveAt(5));
        }

        [Test]
        public void AddItem_VeryLongString_Accepted()
        {
            var longString = new string('X', 10000);
            _comboBox.Items.Add(longString);
            Assert.AreEqual(longString, _comboBox.Items[0]);
        }

        [Test]
        public void AddItem_SpecialCharacters_Accepted()
        {
            _comboBox.Items.Add("C:\\Path\\To\\File.txt");
            _comboBox.Items.Add("https://example.com?q=test&p=1");
            _comboBox.Items.Add("\u65E5\u672C\u8A9E\u30C6\u30B9\u30C8");

            Assert.AreEqual(3, _comboBox.Items.Count);
            Assert.AreEqual("\u65E5\u672C\u8A9E\u30C6\u30B9\u30C8", _comboBox.Items[2]);
        }

        #endregion

        #region Bulk Operations Tests

        [Test]
        public void RemoveMultipleItems_InReverseOrder_RemovesCorrectly()
        {
            _comboBox.Items.AddRange(new object[] { "A", "B", "C", "D", "E" });

            _comboBox.Items.RemoveAt(4);
            _comboBox.Items.RemoveAt(2);
            _comboBox.Items.RemoveAt(0);

            Assert.AreEqual(2, _comboBox.Items.Count);
            Assert.AreEqual("B", _comboBox.Items[0]);
            Assert.AreEqual("D", _comboBox.Items[1]);
        }

        [Test]
        public void RemoveMultipleItems_InForwardOrder_ShiftsIndices()
        {
            _comboBox.Items.AddRange(new object[] { "A", "B", "C", "D", "E" });

            _comboBox.Items.RemoveAt(0);
            _comboBox.Items.RemoveAt(0);

            Assert.AreEqual(3, _comboBox.Items.Count);
            Assert.AreEqual("C", _comboBox.Items[0]);
        }

        #endregion

        #region Internal P/Invoke Tests

        [Test]
        public void GetComboBoxListInternal_WithZeroHandle_ReturnsZero()
        {
            var method = typeof(Hosca.Windows.Forms.MRUComboBox)
                .GetMethod("GetComboBoxListInternal", BindingFlags.Static | BindingFlags.NonPublic);

            if (method != null)
            {
                var result = (IntPtr)method.Invoke(null, new object[] { IntPtr.Zero });
                Assert.AreEqual(IntPtr.Zero, result);
            }
            else
            {
                Assert.Inconclusive("GetComboBoxListInternal method not found via reflection.");
            }
        }

        #endregion

        #region MaxItems Tests

        [Test]
        public void MaxItems_DefaultIs10()
        {
            Assert.AreEqual(10, _comboBox.MaxItems);
        }

        [Test]
        public void MaxItems_CanBeSet()
        {
            _comboBox.MaxItems = 5;
            Assert.AreEqual(5, _comboBox.MaxItems);
        }

        [Test]
        public void AddMRUItem_ExceedingMaxItems_EvictsOldest()
        {
            _comboBox.MaxItems = 3;
            _comboBox.AddMRUItem("First");
            _comboBox.AddMRUItem("Second");
            _comboBox.AddMRUItem("Third");
            _comboBox.AddMRUItem("Fourth");

            Assert.AreEqual(3, _comboBox.Items.Count);
            Assert.AreEqual("Fourth", _comboBox.Items[0]);
            Assert.AreEqual("Third", _comboBox.Items[1]);
            Assert.AreEqual("Second", _comboBox.Items[2]);
        }

        [Test]
        public void AddMRUItem_MaxItemsOne_KeepsOnlyLatest()
        {
            _comboBox.MaxItems = 1;
            _comboBox.AddMRUItem("A");
            _comboBox.AddMRUItem("B");

            Assert.AreEqual(1, _comboBox.Items.Count);
            Assert.AreEqual("B", _comboBox.Items[0]);
        }

        #endregion

        #region AddMRUItem Tests

        [Test]
        public void AddMRUItem_InsertsAtTop()
        {
            _comboBox.AddMRUItem("First");
            _comboBox.AddMRUItem("Second");

            Assert.AreEqual("Second", _comboBox.Items[0]);
            Assert.AreEqual("First", _comboBox.Items[1]);
        }

        [Test]
        public void AddMRUItem_DuplicateMovesToTop()
        {
            _comboBox.AddMRUItem("A");
            _comboBox.AddMRUItem("B");
            _comboBox.AddMRUItem("C");
            _comboBox.AddMRUItem("A");

            Assert.AreEqual(3, _comboBox.Items.Count);
            Assert.AreEqual("A", _comboBox.Items[0]);
            Assert.AreEqual("C", _comboBox.Items[1]);
            Assert.AreEqual("B", _comboBox.Items[2]);
        }

        [Test]
        public void AddMRUItem_CaseInsensitive_DuplicateMovesToTop()
        {
            _comboBox.CaseSensitive = false;
            _comboBox.AddMRUItem("Hello");
            _comboBox.AddMRUItem("World");
            _comboBox.AddMRUItem("hello");

            Assert.AreEqual(2, _comboBox.Items.Count);
            Assert.AreEqual("hello", _comboBox.Items[0]);
            Assert.AreEqual("World", _comboBox.Items[1]);
        }

        [Test]
        public void AddMRUItem_CaseSensitive_DifferentCaseNotDuplicate()
        {
            _comboBox.CaseSensitive = true;
            _comboBox.AddMRUItem("Hello");
            _comboBox.AddMRUItem("hello");

            Assert.AreEqual(2, _comboBox.Items.Count);
            Assert.AreEqual("hello", _comboBox.Items[0]);
            Assert.AreEqual("Hello", _comboBox.Items[1]);
        }

        [Test]
        public void AddMRUItem_NullOrEmpty_Ignored()
        {
            _comboBox.AddMRUItem(null);
            _comboBox.AddMRUItem("");

            Assert.AreEqual(0, _comboBox.Items.Count);
        }

        [Test]
        public void AddMRUItem_SetsTextToItem()
        {
            _comboBox.AddMRUItem("TestItem");
            Assert.AreEqual("TestItem", _comboBox.Text);
        }

        [Test]
        public void AddMRUItem_MultipleDuplicates_AllRemoved()
        {
            _comboBox.Items.Add("A");
            _comboBox.Items.Add("A");
            _comboBox.Items.Add("B");
            _comboBox.AddMRUItem("A");

            Assert.AreEqual(2, _comboBox.Items.Count);
            Assert.AreEqual("A", _comboBox.Items[0]);
            Assert.AreEqual("B", _comboBox.Items[1]);
        }

        #endregion

        #region CaseSensitive Tests

        [Test]
        public void CaseSensitive_DefaultIsFalse()
        {
            Assert.IsFalse(_comboBox.CaseSensitive);
        }

        [Test]
        public void CaseSensitive_CanBeToggled()
        {
            _comboBox.CaseSensitive = true;
            Assert.IsTrue(_comboBox.CaseSensitive);

            _comboBox.CaseSensitive = false;
            Assert.IsFalse(_comboBox.CaseSensitive);
        }

        #endregion

        #region ItemDeleted Event Tests

        [Test]
        public void ItemDeleted_EventExists()
        {
            var eventInfo = typeof(Hosca.Windows.Forms.MRUComboBox).GetEvent("ItemDeleted");
            Assert.IsNotNull(eventInfo);
        }

        [Test]
        public void ItemDeleted_EventHandlerType()
        {
            var eventInfo = typeof(Hosca.Windows.Forms.MRUComboBox).GetEvent("ItemDeleted");
            Assert.AreEqual(typeof(EventHandler<MRUItemEventArgs>), eventInfo.EventHandlerType);
        }

        #endregion

        #region ItemAdded Event Tests

        [Test]
        public void ItemAdded_EventExists()
        {
            var eventInfo = typeof(Hosca.Windows.Forms.MRUComboBox).GetEvent("ItemAdded");
            Assert.IsNotNull(eventInfo);
        }

        [Test]
        public void ItemAdded_FiredOnAddMRUItem()
        {
            string addedItem = null;
            _comboBox.ItemAdded += (s, e) => addedItem = e.Item;

            _comboBox.AddMRUItem("Test");

            Assert.AreEqual("Test", addedItem);
        }

        [Test]
        public void ItemAdded_NotFiredForNullOrEmpty()
        {
            bool fired = false;
            _comboBox.ItemAdded += (s, e) => fired = true;

            _comboBox.AddMRUItem(null);
            _comboBox.AddMRUItem("");

            Assert.IsFalse(fired);
        }

        [Test]
        public void ItemAdded_FiredForDuplicateReAdd()
        {
            int fireCount = 0;
            _comboBox.ItemAdded += (s, e) => fireCount++;

            _comboBox.AddMRUItem("A");
            _comboBox.AddMRUItem("A");

            Assert.AreEqual(2, fireCount);
        }

        #endregion

        #region MRUItemEventArgs Tests

        [Test]
        public void MRUItemEventArgs_StoresItem()
        {
            var args = new MRUItemEventArgs("test");
            Assert.AreEqual("test", args.Item);
        }

        [Test]
        public void MRUItemEventArgs_InheritsFromEventArgs()
        {
            var args = new MRUItemEventArgs("test");
            Assert.IsInstanceOf<EventArgs>(args);
        }

        #endregion
    }

    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ToolStripMRUComboBoxTests
    {
        private ToolStripMRUComboBox _toolStripCombo;

        [SetUp]
        public void SetUp()
        {
            _toolStripCombo = new ToolStripMRUComboBox();
        }

        [TearDown]
        public void TearDown()
        {
            _toolStripCombo?.Dispose();
        }

        [Test]
        public void Constructor_CreatesInternalMRUComboBox()
        {
            Assert.IsNotNull(_toolStripCombo.ComboBox);
            Assert.IsInstanceOf<Hosca.Windows.Forms.MRUComboBox>(_toolStripCombo.ComboBox);
        }

        [Test]
        public void ComboBox_Property_ReturnsMRUComboBox()
        {
            var combo = _toolStripCombo.ComboBox;
            Assert.IsNotNull(combo);
        }

        [Test]
        public void Items_Property_ReturnsComboBoxItems()
        {
            var items = _toolStripCombo.Items;
            Assert.IsNotNull(items);
            Assert.AreEqual(0, items.Count);
        }

        [Test]
        public void Items_AddItem_ReflectedInComboBox()
        {
            _toolStripCombo.Items.Add("Test");
            Assert.AreEqual(1, _toolStripCombo.ComboBox.Items.Count);
            Assert.AreEqual("Test", _toolStripCombo.ComboBox.Items[0]);
        }

        [Test]
        public void Items_SameReferenceAsComboBoxItems()
        {
            Assert.AreSame(_toolStripCombo.Items, _toolStripCombo.ComboBox.Items);
        }

        [Test]
        public void ComboBox_HasCorrectDropDownStyle()
        {
            Assert.AreEqual(ComboBoxStyle.DropDown, _toolStripCombo.ComboBox.DropDownStyle);
        }

        [Test]
        public void ComboBox_HasCorrectDrawMode()
        {
            Assert.AreEqual(DrawMode.OwnerDrawVariable, _toolStripCombo.ComboBox.DrawMode);
        }

        [Test]
        public void MaxItems_DelegatesToComboBox()
        {
            _toolStripCombo.MaxItems = 5;
            Assert.AreEqual(5, _toolStripCombo.ComboBox.MaxItems);
            Assert.AreEqual(5, _toolStripCombo.MaxItems);
        }

        [Test]
        public void MaxItems_DefaultIs10()
        {
            Assert.AreEqual(10, _toolStripCombo.MaxItems);
        }

        [Test]
        public void CaseSensitive_DelegatesToComboBox()
        {
            _toolStripCombo.CaseSensitive = true;
            Assert.IsTrue(_toolStripCombo.ComboBox.CaseSensitive);
            Assert.IsTrue(_toolStripCombo.CaseSensitive);
        }

        [Test]
        public void CaseSensitive_DefaultIsFalse()
        {
            Assert.IsFalse(_toolStripCombo.CaseSensitive);
        }

        [Test]
        public void AddMRUItem_DelegatesToComboBox()
        {
            _toolStripCombo.AddMRUItem("Hello");
            Assert.AreEqual(1, _toolStripCombo.ComboBox.Items.Count);
            Assert.AreEqual("Hello", _toolStripCombo.ComboBox.Items[0]);
        }

        [Test]
        public void AddMRUItem_InsertsAtTop()
        {
            _toolStripCombo.AddMRUItem("First");
            _toolStripCombo.AddMRUItem("Second");

            Assert.AreEqual("Second", _toolStripCombo.Items[0]);
            Assert.AreEqual("First", _toolStripCombo.Items[1]);
        }
    }

    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class MRUComboBoxStripControlHostTests
    {
        [Test]
        public void DefaultConstructor_CreatesControl()
        {
            using (var host = new MRUComboBoxStripControlHost())
            {
                Assert.IsNotNull(host.Control);
            }
        }

        [Test]
        public void ParameterizedConstructor_UsesProvidedControl()
        {
            using (var button = new Button())
            using (var host = new MRUComboBoxStripControlHost(button))
            {
                Assert.AreSame(button, host.Control);
            }
        }

        [Test]
        public void ParameterizedConstructor_WithMRUComboBox_Works()
        {
            using (var combo = new Hosca.Windows.Forms.MRUComboBox())
            using (var host = new MRUComboBoxStripControlHost(combo))
            {
                Assert.AreSame(combo, host.Control);
                Assert.IsInstanceOf<Hosca.Windows.Forms.MRUComboBox>(host.Control);
            }
        }

        [Test]
        public void Namespace_IsHoscaWindowsForms()
        {
            var type = typeof(MRUComboBoxStripControlHost);
            Assert.AreEqual("Hosca.Windows.Forms", type.Namespace);
        }
    }
}
