using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Dumpify;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

// Filters
public interface IGuiRowCondition
{
    bool IsSatisfiedBy(GuiTableRow row); // use by GuiTableControl
    bool IsSatisfiedBy(GuiGridView grid, int row); // use by GuiGridView
}
public interface IGuiGridViewRowCondition
{
    bool IsSatisfiedBy(int row);
}
// 示例1：列包含指定文本
public class ContainsCondition : IGuiRowCondition
{
    private readonly int _columnIndex;
    private readonly string _value;

    public ContainsCondition(int columnIndex, string value)
    {
        _columnIndex = columnIndex;
        _value = value;
    }

    public bool IsSatisfiedBy(GuiTableRow row)
    {
        return row.ColText(_columnIndex).Contains(_value);
    }

    public bool IsSatisfiedBy(GuiGridView grid, int row)
    {
        return grid.GetCellValue(row, _columnIndex).Contains(_value);
    }
}
public class NotContainsCondition : IGuiRowCondition
{
    private readonly int _columnIndex;
    private readonly string _value;

    public NotContainsCondition(int columnIndex, string value)
    {
        _columnIndex = columnIndex;
        _value = value;
    }

    public bool IsSatisfiedBy(GuiTableRow row)
    {
        return !row.ColText(_columnIndex).Contains(_value);
    }

    public bool IsSatisfiedBy(GuiGridView grid, int row)
    {
        return !grid.GetCellValue(row, _columnIndex).Contains(_value);
    }
}

// 示例2：列等于指定数值
public class EqualsNumericCondition<T> : IGuiRowCondition where T : struct
{
    private readonly int _columnIndex;
    private readonly T _expectedValue;

    public EqualsNumericCondition(int columnIndex, T expectedValue)
    {
        _columnIndex = columnIndex;
        _expectedValue = expectedValue;
    }

    public bool IsSatisfiedBy(GuiTableRow row)
    {
        return row.TryGetColumnValue(_columnIndex, out T actualValue)
               && actualValue.Equals(_expectedValue);
    }

    public bool IsSatisfiedBy(GuiGridView grid, int row)
    {
        return grid.TryGetColumnValue(row, _columnIndex, out T actualValue)
            && actualValue.Equals(_expectedValue);
    }
}
public class GuiRowFilterBuilder
{
    private readonly List<IGuiRowCondition> _conditions = new List<IGuiRowCondition>();

    // 添加包含条件
    public GuiRowFilterBuilder Contains(int columnIndex, string value)
    {
        _conditions.Add(new ContainsCondition(columnIndex, value));
        return this;
    }
    public GuiRowFilterBuilder NotContains(int columnIndex, string value)
    {
        _conditions.Add(new NotContainsCondition(columnIndex, value));
        return this;
    }
    // 添加等于数值条件
    public GuiRowFilterBuilder EqualsNumeric<T>(int columnIndex, T value) where T : struct
    {
        _conditions.Add(new EqualsNumericCondition<T>(columnIndex, value));
        return this;
    }

    // 组合条件（AND 逻辑）
    public Func<GuiTableRow, bool> BuildForTable()
    {
        return row => _conditions.All(condition => condition.IsSatisfiedBy(row));
    }
    public Func<GuiGridView, int, bool> BuildForGrid() // for GridView
    {
        return (grid, rowNo) => _conditions.All(condition => condition.IsSatisfiedBy(grid, rowNo));
    }
    // 扩展点：支持 OR 逻辑或其他组合
    public Func<GuiTableRow, bool> BuildAnyForTable()
    {
        return row => _conditions.Any(condition => condition.IsSatisfiedBy(row));
    }
    public Func<GuiGridView, int, bool> BuildAnyForGrid()
    {
        return (grid, rowNo) => _conditions.Any(condition => condition.IsSatisfiedBy(grid, rowNo));
    }
}
//// 构建筛选条件
//var filter = new SapRowFilterBuilder()
//    .Contains(0, "VAS")       // 列0包含 "VAS"
//    .EqualsNumeric(2, 100)    // 列2等于 100
//    .Build();

//// 应用筛选
//var filteredRows = sapRows.Where(filter);

//var dynamicFilter = new SapRowFilterBuilder();

//if (checkBoxVAS.Checked)
//    dynamicFilter.Contains(0, "VAS");

//if (numericUpDown.Value > 0)
//    dynamicFilter.EqualsNumeric(2, (int) numericUpDown.Value);

//var finalFilter = dynamicFilter.Build();
//var result = sapRows.Where(finalFilter);

//Filters end

public static class SapGuiScriptingExtensions
{
    //********* Get Wrapper Methods *******************************
    public static GuiGridWrapper GetGridWrapper(this GuiSession s, string gridId)
    {
        return new GuiGridWrapper(s, gridId);
    }
    public static GuiTabStripWrapper GetTabStripWrapper(this GuiSession s, string id)
    {
        return new GuiTabStripWrapper(s, id);
    }
    public static GuiTableWrapper GetTableWrapper(this GuiSession s, string id)
    {
        return new GuiTableWrapper(s, id);
    }
    //********* Get Wrapper Methods End*******************************

    public static void SelectTab(this GuiSession s, string parentId, string tabText)
    {
        s.GetTabStripWrapper(parentId).Select(tabText);
    }

    //*******************GuiTableRow Extension Methods************
    public static void SetFocus(this GuiTableRow rowObj, int colIndex)
    {
        (rowObj.ElementAt(colIndex) as GuiVComponent).SetFocus();
    }
    public static void DbClickCell(this GuiTableRow rowObj, int colIndex)
    {
        /*
         * GuiTableRow 的类型为GuiComponentCollection,不能转型为GuiComponent!!!!
         * 2025-4-23
         */
        rowObj.SetFocus(colIndex);
        //(rowObj is null).Dump("DbClickCell:rowObj is null?");
        GuiComponent c = rowObj.ElementAt(0);
        //(c is null).Dump("DbClickCell:c is null?");
        c.GetSession().ActiveWindow.SendVKey(2);
    }
    public static void Select(this GuiTableRow rowObj)
    {
        rowObj.Selected = true;
    }
    public static void UnSelect(this GuiTableRow rowObj)
    {
        rowObj.Selected = false;
    }
    //2024-3-18
    public static string GetTableRowText(this GuiTableRow rowObj, int col)
    {
        return (rowObj.ElementAt(col) as GuiVComponent).Text;
    }
    //2024-3-18
    public static void SetTableRowText(this GuiTableRow rowObj, int col, string txt)
    {
        (rowObj.ElementAt(col) as GuiVComponent).Text = txt;
    }
    public static string ColText(this GuiTableRow row, int colIndex)
    {
        return (row.ElementAt(colIndex) as GuiVComponent).Text;
    }
    // 支持数值类型转换（如需要比较数值）
    public static bool TryGetColumnValue<T>(this GuiTableRow row, int columnIndex, out T value)
    {
        value = default;
        string text = row.ColText(columnIndex);
        try
        {
            value = (T)Convert.ChangeType(text, typeof(T));
            return true;
        }
        catch
        {
            return false;
        }
    }
    public static bool TryGetColumnValue<T>(this GuiGridView grid, int row, int columnIndex, out T value)
    {
        value = default;
        string text = grid.GetCellValue(row, columnIndex);
        try
        {
            value = (T)Convert.ChangeType(text, typeof(T));
            return true;
        }
        catch
        {
            return false;
        }
    }
    public static void SetCol(this GuiTableRow row, int colIndex, string txt)
    {
        (row.ElementAt(colIndex) as GuiVComponent).Text = txt;
    }
    //*******************GuiTableRow Extension Methods End************

    public static void ForEach(this GuiVContainer vc, Func<GuiComponent, bool> Action)
    {/*2024-12-4*/
        foreach (GuiComponent c in vc.Children)
        {
            var r = Action(c);
            if (r == true)
            {
                return;
            }
        }
    }
    public static void ForEach(this GuiComponent guiComp, Func<GuiComponent, bool> Action)
    {
        if (guiComp is GuiVContainer vc)
        {
            vc.ForEach(c => Action(c));
        }
        else
        {
            $"{guiComp.Id} has no child!".Dump("Atention:");
        }
    }
    public static void FillText(this GuiComponent c, (string fieldName, string fieldValue)[] inputArray)
    {/*2024-12-6*/
        c.ForEach(v =>
        {
            if (v is GuiCTextField ctxt)
            {
                //$"{ctxt.Id}-{ctxt.Name}".Dump();
                var kv = inputArray.FirstOrDefault(a => a.fieldName == ctxt.Name);
                if (kv.fieldName != null)
                    ctxt.Text = kv.fieldValue ?? "";
            }
            if (v is GuiTextField txt && !txt.IsOField) // OField is Display only
            {
                //$"{txt.Id} | {txt.Name} | {txt.IsOField}".Dump();
                var kv = inputArray.FirstOrDefault(a => a.fieldName == txt.Name);
                if (kv.fieldName != null)
                    txt.Text = kv.fieldValue ?? "";
            }
            return false;
        });
    }
    #region common
    public static void WaitSap(this GuiSession s)
    {
        while (s.Busy) Thread.Sleep(1000);
    }
    public static async Task WaitSapAsync(this GuiSession s)
    {
        while (s.Busy)
        {
            await Task.Delay(200);
        }
    }
    public static dynamic findById2(this GuiSession session, string Id)
    { //找不到控件时返回null
        GuiComponent comp = session.FindById(Id, false);
        return comp;
    }
    public static void Set(this GuiSession s, string id, string text)//2025-2-19
    {
        var vc = s.FindById<GuiVComponent>(id);
        vc.Text = text;
    }
    public static async Task SetAsync(this GuiSession s, string id, string text)//2025-2-19
    {
        await Task.Run(() => Set(s, id, text));
    }
    public static string Text(this GuiSession s, string id)
    {
        return s.FindById<GuiVComponent>(id).Text;
    }
    public static void Click(this GuiSession s, string id)
    {
        s.FindById<GuiButton>(id).Press();
    }
    public static void Select(this GuiSession s, string id)
    {
        GuiMenu menu = s.FindById<GuiMenu>(id);
        if (menu.Changeable) menu.Select();
    }
    public static T FindById<T>(this GuiSession session, string Id)
    {
        GuiComponent comp = session.FindById(Id);
        return (T)comp;
    }
    //2024-1-24
    public static void F8(this GuiSession s)
    {
        s.Click("wnd[0]/tbar[1]/btn[8]");
    }
    public static void PressEnter(this GuiSession session)
    {
        session.Click("wnd[0]/tbar[0]/btn[0]");
    }
    public static void PressBackF3(this GuiSession session)
    {
        session.Click("wnd[0]/tbar[0]/btn[3]");
    }
    public static void PressSave(this GuiSession session)
    {
        session.Click("wnd[0]/tbar[0]/btn[11]");
    }
    public static GuiUserArea GetUserArea(this GuiSession session)
    {
        return session.FindById<GuiUserArea>("wnd[0]/usr");
    }
    public static GuiFrameWindow GetMainWindow(this GuiSession session)
    {
        return session.FindById<GuiFrameWindow>("wnd[0]");
    }
    public static string GetStatus(this GuiSession session)
    {
        return session.FindById<GuiStatusbar>("wnd[0]/sbar").Text;
    }
    public static dynamic findById2(this GuiFrameWindow win, string Id)
    {
        return win.FindById(Id, false);
    }
    public static dynamic findByName2(this GuiFrameWindow win, string name, string typeName)
    {
        try { return win.FindByName(name, typeName); }
        catch (COMException)
        {
            return null;
        }

    }
    public static T FindByName2<T>(this GuiFrameWindow win, string name)//2024-3-19
    {
        try
        {
            string typeName = typeof(T).Name;
            GuiComponent comp = win.FindByName(name, typeName);
            return (T)comp;
        }
        catch (COMException)
        {
            return default(T);
        }

    }
    public static void ConfirmWarning(this GuiSession s)
    {
        /*
        S	Success
        W	Warning
        E	Error
        A	Abort
        I	Information
        2024-10-12
        */
        while (s.Sbar().MessageType == "W" || s.Sbar().MessageType == "I")
        {
            s.PressEnter();
        }
    }
    public static GuiStatusbar Sbar(this GuiSession s)
    {
        return s.FindById<GuiStatusbar>("wnd[0]/sbar");
    }
    public static string Status(this GuiSession s)
    {
        return s.Sbar().Text;
    }

    public static bool HasPop(this GuiSession s, int winIndex)
    {
        return s.findById2($"wnd[{winIndex}]") != null;
    }
    public static GuiModalWindow Pop(this GuiSession s, int winIndex)
    {
        return s.findById2($"wnd[{winIndex}]");
    }
    public static int GetPopIndex(this GuiModalWindow pop)
    {
        // 定义正则表达式，匹配 wnd[] 中的数字
        string pattern = @"wnd\[(\d+)\]";
        var match = Regex.Match(pop.Id, pattern);

        // 如果匹配成功，提取数字
        if (match.Success)
        {
            return int.Parse(match.Groups[1].Value);
        }

        // 未找到时返回 -1
        return -1;
    }
    public static string PopText(this GuiSession s, int winIndex)
    {
        return s.Pop(winIndex).Text;
    }
    public static void CheckStatus(this GuiSession s,
        Action<GuiModalWindow, int> popAction = null,
        Action<string, string> statusAction = null)
    {
        /**
         * 检查是否跳出了对话框，后者状态栏有消息出现。
         * 一般在切换界面时或者检查输入数据后会出现。
         * 默认情况只是以错误消息的信息提醒，并列出出现的消息
         * 后续可以针对具体的消息，增加处理的额外代码。
         * 2025-3-18
         */

        for (int i = 1; i <= 4; i++) // 对话框一般不会超出4个
        {
            if (s.HasPop(i))
            {
                var pop = s.Pop(i);
                int popIndex = pop.GetPopIndex();
                bool hasToolbar = s.PopToolbar(popIndex) == null ? false : true;
                var buttons = CollectButtons(pop);
                new
                {
                    Name = pop.Name,
                    Id = pop.Id,
                    Title = pop.Text,
                    PopIndex = popIndex,
                    HasToolbar = hasToolbar,
                    Buttons = buttons,
                }.Dump("Unexpected Popup Window!");
                popAction?.Invoke(pop, popIndex);
            }
        }
        var mtype = s.Sbar().MessageType;
        if (mtype != "S" && mtype != "") // S - Success
        {
            new
            {
                MessgaeType = mtype,
                Message = s.Sbar().Text,
            }.Dump("Status Message!");

            statusAction?.Invoke(mtype, s.Sbar().Text);
        }

        string CollectButtons(GuiModalWindow pp)
        {
            int popIndex = pp.GetPopIndex();
            bool hasToolbar = s.PopToolbar(popIndex) == null ? false : true;
            var buttons = "";
            if (hasToolbar)
            {
                var toolbar = s.PopToolbar(popIndex) as GuiToolbar;

                for (int i = 0; i < toolbar.Children.Count; i++)
                {
                    var ctl = toolbar.Children.Item(i) as GuiVComponent;
                    buttons += $"index: {i + 1} - {ctl.Name} - {ctl.Tooltip}\n";
                }
            }
            return buttons;
        }
    }

    //*******************  点击对话框上面的按钮 ********************
    public static GuiToolbar PopToolbar(this GuiSession s, int winIndex)
    {
        return s.findById2($"wnd[{winIndex}]/tbar[0]");
    }
    //对话框的按钮作为控制栏按钮的形式存在，可以通过索引访问。
    public static void ClickPopButton(this GuiSession s, int winIndex, int btnIndex)
    {
        (s.PopToolbar(winIndex).Children.Item(btnIndex - 1) as GuiButton).Press();
    }
    //对话框按钮作为UserArea的子元素存在，不能通过索引访问。
    public static void YesPop(this GuiSession s, int winIndex)
    {
        s.findById2($"wnd[{winIndex}]/usr/btnSPOP-OPTION1").press();
    }
    public static void NoPop(this GuiSession s, int winIndex)
    {
        s.findById2($"wnd[{winIndex}]/usr/btnSPOP-OPTION2").press();
    }
    public static void CancelPop(this GuiSession s, int winIndex)
    {
        s.findById2($"wnd[{winIndex}]/usr/btnSPOP-OPTION_CAN").press();
    }
    //*******************  点击对话框上面的按钮 结束********************

    public static GuiSession GetSession(this GuiComponent c)
    {
        if (c is null)
        {
            throw new NullReferenceException("input parameter GuiComponent c is null!");
        }
        if (c is GuiSession)
        {
            return (GuiSession)c;
        }
        else
        {
            GuiComponent p = c.Parent;
            return GetSession(p);
        }
    }
    public static GuiConnection GetConnection(this GuiSession s)//2024-2-4
    {
        return s.Parent as GuiConnection;
    }
    public static GuiApplication GetApp(this GuiSession s)
    {
        return s.GetConnection().Parent as GuiApplication;
    }
    #endregion common
    #region GuiTable
    public static Dictionary<int, string> ColTitles(this GuiTableControl table)
    {
        Dictionary<int, string> dict = new Dictionary<int, string>();
        int colIndex = 0;
        foreach (GuiTableColumn col in table.Columns)
        {
            dict.Add(colIndex, col.Title);
            colIndex++;
        }
        return dict;
    }
    public static string GetCellValue(this GuiGridView gridView, int row, int col)//2024-1-15
    {
        return gridView.GetCellValue(row, gridView.ColumnOrder[col]);
    }
    public static string CV(this GuiGridView gridView, int row, int col)//2024-1-15
    {
        return gridView.GetCellValue(row, gridView.ColumnOrder[col]);
    }
    public static void GridViewGoThrough(this GuiGridView grid, Func<int, bool> ActionOnRow)
    { //这个更灵活，方便退出循环
        for (int i = 0; i < grid.RowCount; i++)
        {
            if (i > 0 && i % grid.VisibleRowCount == 0) grid.FirstVisibleRow = i;
            var r = ActionOnRow(i);
            if (r) return;
        }
    }
    public static void TableControlGoThrough(this GuiSession session, GuiTableControl table, Func<GuiTableRow, int, int, bool> FunctionOnRow)
    {
        /*
		* Func<GuiTableRow:行对象, int:绝对行号码, int:可见区域行号码, bool:是否结束循环> 参数
		* 
		* 2024-9-25
		**/
        string tableId = table.Id;
        int rowCount = table.RowCount;//全部行数
        int visibleRowCount = table.VisibleRowCount;//可见行数
        int colCount = table.Columns.Count;//列数
        /*
        下面一句代码对总行数进行分组
        若总行数是可见行数的整数倍，则可以直接算出组数。
        否则，使用浮点数除法计算，然后取整，然后再加1，最后加的一组是个零头组。
        2024-1-15
        */
        var vrBlockCount = (rowCount % visibleRowCount == 0) ?
                            rowCount / visibleRowCount :
                            Math.Round(rowCount * 1.0 / visibleRowCount) + 1;

        for (int vrBlock = 0; vrBlock < vrBlockCount; vrBlock++)
        {
            table.VerticalScrollbar.Position = vrBlock * visibleRowCount;//竖直滚动条的位置是使用总行数进行设置的。
            table = session.FindById(tableId) as GuiTableControl;//滚动之后，引用失效，需要从新获取

            int remainRowCount = rowCount - (vrBlock + 1) * visibleRowCount;//如果当前组是最后一组，可能得到一个领头组行数，此时它< visibleRowCount
            for (int row = 0; row < Math.Min(visibleRowCount, remainRowCount + 1); row++)
            {
                var rowObject = table.Rows.Item(row);
                bool break_ = FunctionOnRow(rowObject, row + (vrBlock * visibleRowCount), row);//调用用户提供的delegate.
                table = session.FindById(tableId) as GuiTableControl;//Func调用之后，引用可能失效，重新获取
                if (break_) return;
            }
        }
    }
    #endregion GuiTable
}
