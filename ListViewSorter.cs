using System.Collections;
using System.Windows.Forms;

namespace Service_Manager
{
    class ListViewSorter : IComparer
    {
        private CaseInsensitiveComparer ItemComparer;
        private int sortColumn { get; set; }
        private SortOrder sortOrder { get; set; }

        public int SortColumn
        {
            set { sortColumn = value; }
            get { return sortColumn; }
        }
        public SortOrder SortOrder
        {
            set { sortOrder = value; }
            get { return sortOrder; }
        }

        public ListViewSorter(int column = 0, SortOrder order = SortOrder.Ascending)
        {
            sortColumn = column;
            sortOrder = order;
            ItemComparer = new CaseInsensitiveComparer();
        }

        public int Compare(object x, object y)
        {
            ListViewItem itemX = (ListViewItem)x;
            ListViewItem itemY = (ListViewItem)y;
            int CompareResult = ItemComparer.Compare(itemX.SubItems[this.sortColumn].Text, itemY.SubItems[this.sortColumn].Text);
            if (this.sortOrder == SortOrder.Ascending)
                return CompareResult;
            else
            if (this.sortOrder == SortOrder.Descending)
                return -CompareResult;
            else
                return 0;
        }

    }
}
