using SAPFEWSELib;
//using SapROTWr;
using System;
using System.Collections.Generic;
using System.Linq;
using Serilog;

public class GuiTableWrapper
{

    private string id;
    private GuiSession s;
    public GuiTableControl table;
    public int rowCount;
    public int vRowCount;
    public int posMax;
    private int currentRowIndex;

    public GuiTableWrapper(GuiSession session, string tableId)
    {
        s = session;
        id = tableId;
        table = s.FindById<GuiTableControl>(tableId);
        rowCount = table.RowCount;
        vRowCount = table.VisibleRowCount;
        posMax = table.VerticalScrollbar.Maximum;
    }
    public string Name => table.Name;

    public string Status()
    {
        return $"{table.Name} Row Count:{rowCount},Visible Row Count:{vRowCount},Vertical Scrollbar Position:{table.VerticalScrollbar.Position}/{posMax}";
    }
    public void SelectRow(int rowNo)
    {
        var row = GetRow(rowNo);
        row.Select();
    }
    public void UnselectRow(int rowNo)
    {
        var row = GetRow(rowNo);
        row.Selected = false;
    }
    public GuiTableRow GetRow(int rowNo)
    {
        EnsureRowVisible(rowNo);
        var row = table.GetAbsoluteRow(rowNo);
        return row;
    }
    public void EnsureRowVisible(int i)
    {
        if (i >= posMax)
        {
            table.VerticalScrollbar.Position = posMax;
            RefreshTable();
        }
        if (i <= posMax && i >= table.VerticalScrollbar.Position + vRowCount)//往下滚动到指定行可见。
        {
            table.VerticalScrollbar.Position = Math.Min(i, posMax);
            RefreshTable();
            Log.Debug("{table.Name} VerticalScrollbar.Position updated to {i},refresh table...", table.Name, i);
        }
        if (i >= 0 && i < table.VerticalScrollbar.Position)// 往上滚到到指定行可见。
        {
            table.VerticalScrollbar.Position = i;
            RefreshTable();
        }
    }
    public void RefreshTable()
    {
        table = s.FindById<GuiTableControl>(id);
    }
    private GuiTableRow GetNextRow()
    {
        return table.GetAbsoluteRow(currentRowIndex);
    }
    public void Iter(Func<int, GuiTableRow, bool> func)
    {
        for (int i = 0; i < rowCount; i++) // i是绝对行索引
        {
            EnsureRowVisible(i);
            currentRowIndex = i;
            if (func(i, GetNextRow()))
            {
                break;
            }
        }
    }
    public void FillRows<T>(IEnumerable<T> dataCollection, Action<GuiTableRow, T> action)
    {
        /**
         * 这个方法的实现思考了挺长时间
         * 一开始纠结于定位第一个空行时一次滚动一行还是一次滚动一屏
         * 这个问题和填充动作会改变table的行数，很麻烦。
         * 最后方案是寻找第一个空行使用一次滚动一屏的方式，单独一个方法
         * 填充时则一次滚动一行，也是单独一个方法。
         * 
         * VS在调试功能上的强大程度不是一般小工具能比的，比如LinqPad.
         * 
         * 2025-3-17
         */
        var dataList = dataCollection.ToList();
        int absRow = GotoBlankRow();
        if (absRow == -1)
        {
            throw new Exception("没有找到空行!");
        }
        while (dataList.Count > 0)
        {
            T data = dataList[0];
            dataList.RemoveAt(0);

            var rowObj = table.GetAbsoluteRow(absRow);
            action(rowObj, data);

            s.PressEnter(); // 在空行填充数据之后，按下回车键，sap系统会检查数据，同时会添加更多的空行
            RefreshTable(); //按下回车后，sap可能会添加空行，这会导致变量失效，需要刷新table的变量。

            table.VerticalScrollbar.Position = ++absRow;
            RefreshTable();
        }
    }
    public int LastRealRow(int col = 0)
    {
        //最后一行，列Count > 0
        int result = 0;
        for (int i = rowCount - 1; i >= 0; i--)
        {
            EnsureRowVisible(i);
            var row = table.GetAbsoluteRow(i);
            if (row.Count > 0 && row.GetTableRowText(col) != string.Empty)
            {
                result = i;
                break;
            }

        }
        return result;
    }
    public int GotoBlankRow()
    {
        for (int i = 0; i < rowCount; i++)
        {
            if (i < table.VerticalScrollbar.Position) // 滚动条不在顶部
            {
                i = table.VerticalScrollbar.Position;
            }

            EnsureRowVisible(i);
            // 获取当前行
            currentRowIndex = i;
            GuiTableRow row = GetNextRow();
            // 检查首列是否为空
            //row.ColText(1).Dump($"{i}");
            if (row.Count > 0 && row.GetTableRowText(1) == string.Empty)
            {
                table.VerticalScrollbar.Position = i; // 把空行作为第一行
                RefreshTable();
                return i;
            }
        }

        // 如果没有找到空行，返回 -1
        return -1;
    }

}

public class GuiTabStripWrapper // 2025-1-20
{
    private GuiSession s;
    private string id;
    private GuiTabStrip tabStrip;

    public GuiTabStripWrapper(GuiSession session, string tabStripId)
    {
        id = tabStripId;
        s = session;
        tabStrip = s.findById2(id);
        s.WaitSap();
    }
    public void Select(string tabText)
    {
        if (tabStrip.SelectedTab.Text == tabText) return;
        foreach (GuiTab tab in tabStrip.Children)
        {
            if (tab.Text == tabText)
            {
                tab.Select();
                tabStrip = s.findById2(id);
                return;
            }
        }
    }
}
public class GuiGridWrapper
{
    /*
     * GuiGridView 并没有GuiGridViewRow,
     * 它只能通过行列坐标访问单元格
     * 2025-3-28
     */
    private string id;
    private GuiSession s;
    private GuiGridView grid;
    public int? RowCount = 0;
    public GuiGridWrapper(GuiSession session, string gridID)
    {
        id = gridID;
        s = session;
        grid = s.findById2(id);
        RowCount = grid?.RowCount;
    }
    public void SelectRow(int row)
    {
        grid.SelectedRows = row.ToString();
    }
    public string CellValue(int row, int col)
    {
        if (row >= grid.RowCount || col >= grid.ColumnCount)
        {
            return "";
        }
        if (grid.FirstVisibleRow + grid.VisibleRowCount <= row)
        {
            grid.FirstVisibleRow = row;
        }
        return grid.GetCellValue(row, grid.ColumnOrder.Item(col));
    }
}

public class MyGuiSession
{
    GuiSession _session;
    public GuiSession Session => _session;

    public MyGuiSession(GuiSession s)
    {
        _session = s;
    }
    public override string ToString()
    {
        var info = Session.Info.Transaction + "@";
        if (Environment.MachineName == "HOME")
            info += (Session.Parent as GuiConnection).Description.Split()[1];
        else
            info += (Session.Parent as GuiConnection).Description.Split()[0];
        info += " " + _session.ActiveWindow.Text.Trim();
        return info;
    }
}
