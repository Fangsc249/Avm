
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using Serilog;
using SAPFEWSELib;
using System.Dynamic;
using System.Collections.Generic;
using AutomationVM.Core;
using Dumpify;
using static AutomationVM.Core.InstructionParser;
using System;
using System.Runtime.InteropServices;
/*
* 为了方便，这个直接引用AutomationVM项目，实际情况可能是引用其dll更符合常规。
* 2025-3-26
* 
* 什么SAP操作值得创建相应的指令确实不容易。
* 2025-4-22
*/
namespace AutomationVM.Module.SAPInstructions
{
    public static class RowFilterListExt
    { //2025-4-23
        public static Func<GuiTableRow, bool> BuildTableFilter(this List<RowFilter> filters)
        {
            var filterBuilder = new GuiRowFilterBuilder();
            //"Hello".Dump();
            foreach (var rf in filters)
            {
                if (rf.FilterRule == "Contains")
                    filterBuilder.Contains(rf.ColumnIndex, rf.Value);
                if (rf.FilterRule == "NotContains")
                    filterBuilder.NotContains(rf.ColumnIndex, rf.Value);
            }
            var filter = filterBuilder.BuildForTable();
            return filter;
        }
        public static Func<GuiGridView, int, bool> BuildGridFilter(this List<RowFilter> filters)
        {
            var filterBuilder = new GuiRowFilterBuilder();

            foreach (var rf in filters)
            {
                if (rf.FilterRule == "Contains")
                    filterBuilder.Contains(rf.ColumnIndex, rf.Value);
                if (rf.FilterRule == "NotContains")
                    filterBuilder.NotContains(rf.ColumnIndex, rf.Value);
            }
            var filter = filterBuilder.BuildForGrid();
            return filter;
        }
    }
    public class RowFilter // 表格行数据的筛选条件参数 2025-3-27
    {
        public int ColumnIndex { get; set; }
        public string FilterRule { get; set; }
        public string Value { get; set; }
    }
    public abstract class SapInstructionBase : Instruction
    {
        protected abstract Task PerformSapAction(GuiSession session);
        protected ExecutionContext Context { get; private set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Context = context;
            GuiSession session = context.Vars.session;
            // 子类实现具体操作
            await PerformSapAction(session);
            await Task.Run(() =>
            {
                // 公共的等待和状态处理
                session.WaitSap();
                UpdateStatusBar(context, session);
                context.Vars.Title = session.ActiveWindow.Text;
                context.Vars.HasPop1 = session.HasPop(1);
                context.Vars.HasPop2 = session.HasPop(2);
            });
        }

        protected void UpdateStatusBar(ExecutionContext context, GuiSession session)
        {
            GuiStatusbar sbar = session.Sbar();
            context.Vars.SapStatusType = sbar.MessageType;
            context.Vars.SapStatus = sbar.Text;
        }
    }
    public class SapConnectTo : Instruction
    {
        public string System { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            System = await EvaluateParameterAsync(System, context) as string;
            var s = context.GetService<SAPActive>().GetSession(System);
            context.Vars.session = s ?? throw new SapConnectionFailureException($"Can't logon to {System}", CallerInfo());
            Log.Information("Connected to: {System}", System);
            Log.Debug($"System Number : {s.Info.SystemNumber}");
            Log.Debug($"Session Id    : {s.Info.SystemSessionId}");
            Log.Debug($"Transaction   : {s.Info.Transaction}");
            Log.Information("User        : {User}", s.Info.User);
            Log.Debug($"UI Guidline   : {s.Info.UI_GUIDELINE}");
            Log.Debug($"Screen Number : {s.Info.ScreenNumber}");
            Log.Debug($"App Server    : {s.Info.ApplicationServer}");
            Log.Debug($"Client        : {s.Info.Client}");
            Log.Debug($"Code Page     : {s.Info.Codepage}");
            Log.Debug($"Group         : {s.Info.Group}");
            Log.Debug($"Gui CodePage  : {s.Info.GuiCodepage}");
            Log.Debug($"Language      : {s.Info.Language}");
            Log.Debug($"Message Server: {s.Info.MessageServer}");
            Log.Debug($"Program       : {s.Info.Program}");

            if (s.Info.Transaction == "S000")
            {
                context.Vars.NeedLogon = true;
            }
            else
            {
                context.Vars.NeedLogon = false;
            }

        }
    }
    public class SapLogon : SapInstructionBase
    {
        public string UserId { get; set; }
        public string User { get; set; }

        public string PasswordId { get; set; }
        public string Password { get; set; }

        public string LangCodeId { get; set; }
        public string LangCode { get; set; }

        //public override async Task ExecuteCoreAsync(ExecutionContext context)
        //{
        //    UserId = await EvaluateParameterAsync(UserId, context) as string;
        //    User = await EvaluateParameterAsync(User, context) as string;
        //    PasswordId = await EvaluateParameterAsync(PasswordId, context) as string;
        //    Password = await EvaluateParameterAsync(Password, context) as string;
        //    LangCodeId = await EvaluateParameterAsync(LangCodeId, context) as string;
        //    LangCode = await EvaluateParameterAsync(LangCode, context) as string;

        //    GuiSession s = context.Vars.session;
        //    await s.WaitSapAsync();
        //    await Try(CallerInfo(), () => s.SetAsync(UserId, User));
        //    await Try(CallerInfo(), () => s.SetAsync(PasswordId, Password));
        //    await Try(CallerInfo(), () => s.SetAsync(LangCodeId, LangCode));

        //    s.PressEnter();
        //    await s.WaitSapAsync();
        //    GuiStatusbar sbar = s.Sbar();
        //    context.Vars.SapStatusType = sbar.MessageType;
        //    context.Vars.SapStatus = sbar.Text;
        //    context.Vars.NeedLogon = false;
        //}

        protected override async Task PerformSapAction(GuiSession s)
        {
            UserId = await EvaluateParameterAsync(UserId, Context) as string;
            User = await EvaluateParameterAsync(User, Context) as string;
            PasswordId = await EvaluateParameterAsync(PasswordId, Context) as string;
            Password = await EvaluateParameterAsync(Password, Context) as string;
            LangCodeId = await EvaluateParameterAsync(LangCodeId, Context) as string;
            LangCode = await EvaluateParameterAsync(LangCode, Context) as string;

            await s.WaitSapAsync();
            await Try(CallerInfo(), () => s.SetAsync(UserId, User));
            await Try(CallerInfo(), () => s.SetAsync(PasswordId, Password));
            await Try(CallerInfo(), () => s.SetAsync(LangCodeId, LangCode));

            s.PressEnter();
            await s.WaitSapAsync();
        }
    }
    public class SapStartTransaction : SapInstructionBase
    {
        public string TCode { get; set; }
        protected override async Task PerformSapAction(GuiSession session)
        {
            session.StartTransaction(TCode);
            await Task.CompletedTask;
        }
    }
    //*******************************************************************
    public class SapEnter : SapInstructionBase
    {
        protected override async Task PerformSapAction(GuiSession session)
        {
            session.PressEnter();
            await Task.CompletedTask;
        }
    }
    public class SapBack : SapInstructionBase
    {
        protected override async Task PerformSapAction(GuiSession session)
        {
            session.PressBackF3();
            await Task.CompletedTask;
        }
    }
    public class SapExecute : SapInstructionBase
    {
        protected override async Task PerformSapAction(GuiSession session)
        {
            session.F8();
            await Task.CompletedTask;
        }
    }
    public class SapSave : SapInstructionBase
    {
        protected override async Task PerformSapAction(GuiSession session)
        {
            session.PressSave();
            await Task.CompletedTask;
        }
    }

    // ***************** text field *************************************
    public class SapSetText : SapInstructionBase
    {
        // "Id": "$id_01",
        // Json脚本中的上述这样会导致这里的Id 被赋值为$id_01
        // 其中的$符号提示我们，它是字典Variables中的一个Key
        // 所以我们需要去掉$,以id_01作为key从字典中获取它的值后再赋值给Id
        // 2025-3-22

        public string Id { get; set; }
        public string Value { get; set; }

        protected override async Task PerformSapAction(GuiSession session)
        {
            var evaluedId = await EvaluateParameterAsync(Id, Context) as string;
            var evaluedValue = await EvaluateParameterAsync(Value, Context) as string;
            try
            {
                session.Set(evaluedId, evaluedValue);
            }
            catch (COMException cex)
            {
                cex.Data.Add("File", "SAPInstructions.cs Line 236");
                cex.Data.Add("Operation", "SapSetText");
                cex.Data.Add("Id", evaluedId);
                cex.Data.Add("Value", evaluedValue);
                throw;
            }
        }
    }
    public class SapGetText : SapInstructionBase
    {
        public string Id { get; set; }
        public string TargetVariable { get; set; }

        protected override async Task PerformSapAction(GuiSession session)
        {
            Id = await EvaluateParameterAsync(Id, Context) as string;

            try
            {
                Context.VarsDict[TargetVariable] = session.Text(Id);
            }
            catch (COMException cex)
            {
                cex.Data.Add("File", "SAPInstructions.cs Line 258");
                cex.Data.Add("Method", "SapGetText");
                cex.Data.Add("Id", Id);
                cex.Data.Add("TargetVariable", TargetVariable);
                throw;
            }
        }
    }
    //*******************************************************************
    public class SapClickButton : SapInstructionBase
    {
        public string Id { get; set; }
        //public override async Task ExecuteCoreAsync(ExecutionContext context)
        //{
        //    Id = await EvaluateParameterAsync(Id, context) as string;
        //    GuiSession session = context.Vars.session;
        //    await Task.Run(() =>
        //    {
        //        session.Click(Id);
        //        session.WaitSap();
        //        GuiStatusbar sbar = session.Sbar();
        //        context.Vars.SapStatusType = sbar.MessageType;
        //        context.Vars.SapStatus = sbar.Text;
        //    });
        //}

        protected override async Task PerformSapAction(GuiSession session)
        {
            Id = await EvaluateParameterAsync(Id, Context) as string;
            try
            {
                session.Click(Id);
            }
            catch (COMException cex)
            {
                cex.Data.Add("File", "SAPInstructions.cs Line 294");
                cex.Data.Add("Method", "SapClickButton");
                cex.Data.Add("Id", Id);
                throw;
            }
            session.WaitSap();
        }
    }
    public class SapSelectMenu : SapInstructionBase
    {
        public string Id { get; set; }
        protected override async Task PerformSapAction(GuiSession session)
        {
            Id = await EvaluateParameterAsync(Id, Context) as string;

            try
            {
                GuiMenu menu = session.FindById<GuiMenu>(Id);
                if (menu.Changeable)
                {
                    menu.Select();
                    session.WaitSap();
                }
            }
            catch (COMException cex)
            {
                cex.Data.Add("File", "SAPInstructions.cs 333 Line");
                cex.Data.Add("Method", nameof(SapSelectMenu));
                cex.Data.Add("Id", Id);
                throw;
            }
        }
    }
    public class SapSelectTab : SapInstructionBase
    {
        public string Id { get; set; }
        public string TabText { get; set; }
        
        protected override async Task PerformSapAction(GuiSession session)
        {
            Id = await EvaluateParameterAsync(Id, Context) as string;
            TabText = await EvaluateParameterAsync(TabText, Context) as string;

            try
            {
                session.SelectTab(Id, TabText);
                session.WaitSap();
            }
            catch (COMException cex)
            {
                cex.Data.Add("Method", nameof(SapSelectTab));
                cex.Data.Add("Id", Id);
                cex.Data.Add("TabText", TabText);
                throw;
            }
        }
    }
    public class SapSelectComboBox : SapInstructionBase
    {
        //public override Task ExecuteCoreAsync(ExecutionContext context)
        //{
        //    //throw new Exception("");
        //    return Task.CompletedTask;
        //}

        protected override async Task PerformSapAction(GuiSession session)
        {
            await Task.CompletedTask;
        }
    }
    //*******************GuiTableControl GuiGridView*********************
    public class SapDoubleClickTableLine : SapInstructionBase
    {
        //双击GuiTableControl的某一行的某一列
        public string Id { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }

        protected override async Task PerformSapAction(GuiSession session)
        {
            var id = await EvaluateParameterAsync(Id, Context) as string;

            try
            {
                var wtable = session.GetTableWrapper(id);
                GuiTableRow row = wtable.GetRow(Row);
                row.DbClickCell(Column);
            }
            catch (Exception ex)
            {
                ex.Data.Add("Method", nameof(SapDoubleClickTableLine));
                ex.Data.Add("Id", id);
                ex.Data.Add("Row", Row);
                ex.Data.Add("Column", Column);
                throw;
            }
        }
    }
    public class SapGetTableData : Instruction
    {
        //获取全部行，全部列的数据
        //如果表格太大，考虑写单独的脚本，这里面对通用场景。
        //2025-3-22
        public string Id { get; set; }
        public string Columns { get; set; } //"0,2,5,7" etc.
        public string TargetVariable { get; set; }
        public List<RowFilter> Filters { get; set; } // 2025-3-28

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Id = await EvaluateParameterAsync(Id, context) as string;

            GuiSession s = context.Vars.session;
            var list = new List<ExpandoObject>();
            context.VarsDict[TargetVariable] = list;
            //var filterBuilder = new GuiRowFilterBuilder();
            var table = s.findById2(Id) as GuiTableControl;
            //Filters.Dump("Row Filters");
            //foreach (var rf in Filters)
            //{
            //    if (rf.FilterRule == "Contains")
            //        filterBuilder.Contains(rf.ColumnIndex, rf.Value);
            //    if (rf.FilterRule == "NotContains")
            //        filterBuilder.NotContains(rf.ColumnIndex, rf.Value);
            //}
            //var filter = filterBuilder.BuildForTable();
            var filter = Filters.BuildTableFilter();
            if (Columns != null)
            {
                await Task.Run(() =>
                {
                    var colList = Columns.Split(',');
                    var colNames = table.ColTitles();
                    s.TableControlGoThrough(table, (rowObj, rowTotal, rowVisible) =>
                    {
                        if (!filter(rowObj))
                        {
                            return false;
                        }
                        var epo = new ExpandoObject() as IDictionary<string, object>;
                        foreach (string indexStr in colList)
                        {
                            int i = int.Parse(indexStr);
                            if (i < colNames.Count)
                                epo[colNames[i]] = rowObj.ColText(i).Replace('\'', ' ');
                        }

                        list.Add(epo as ExpandoObject);
                        return false;
                    });
                });
            }
            else
            {
                await Task.Run(() =>
                {
                    var colNames = table.ColTitles();
                    s.TableControlGoThrough(table, (rowObj, rowTotal, rowVisible) =>
                    {
                        if (!filter(rowObj))
                        {
                            return false;
                        }
                        var epo = new ExpandoObject() as IDictionary<string, object>;
                        foreach (var kv in colNames)
                        {
                            if (kv.Key < rowObj.Count)
                                epo[kv.Value] = rowObj.ColText(kv.Key).Replace('\'', ' ');
                        }

                        list.Add(epo as ExpandoObject);
                        return false;
                    });
                });
            }

        }
    }
    public class SapGetGridData : Instruction
    {
        public string Id { get; set; }
        public string Columns { get; set; } //"0,2,5,7" etc.
        public string TargetVariable { get; set; }
        public List<RowFilter> Filters { get; set; } // 2025-3-28

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Id = await EvaluateParameterAsync(Id, context) as string;

            GuiSession s = context.Vars.session;
            var list = new List<ExpandoObject>();
            context.VarsDict[TargetVariable] = list;

            //var filterBuilder = new GuiRowFilterBuilder();
            //Filters.Dump("Row Filters");
            //foreach (var rf in Filters)
            //{
            //    if (rf.FilterRule == "Contains")
            //        filterBuilder.Contains(rf.ColumnIndex, rf.Value);
            //    if (rf.FilterRule == "NotContains")
            //        filterBuilder.NotContains(rf.ColumnIndex, rf.Value);
            //}
            //var filter = filterBuilder.BuildForGrid();
            var filter = Filters.BuildGridFilter();//2025-4-23
            var grid = s.findById2(Id) as GuiGridView;
            //grid.ColumnOrder 这里是dynamic，vs2022下则是object
            var colNames = (grid.ColumnOrder as GuiCollection);
            var fieldList = new List<object>();
            for (int i = 0; i < grid.ColumnCount; i++)
            {
                //每一个列都有多个版本的Title，所以是一个集合。
                //grid.GetColumnTitles vs2022 返回object!
                var titles = grid.GetColumnTitles(colNames.Item(i));
                var field = new
                {
                    Field = colNames.Item(i),
                    Desc = titles.Count > 0 ? titles[0] : "",
                    Desc2 = titles.Count > 1 ? titles[1] : "",
                    Desc3 = titles.Count > 2 ? titles[2] : "",
                    Desc4 = titles.Count > 3 ? titles[3] : ""
                };
                fieldList.Add(field);
            }
            var colIndexStrs = Columns.Split(',');
            grid.GridViewGoThrough(row =>
            {
                if (!filter(grid, row)) // 如果不符合筛选条件，直接跳过。
                {
                    return false;
                }

                var epo = new ExpandoObject() as IDictionary<string, object>;
                foreach (string indexStr in colIndexStrs)
                {
                    int col = int.Parse(indexStr);
                    epo[colNames.Item(col)] = grid.GetCellValue(row, col);
                }
                list.Add(epo as ExpandoObject);
                return false;
            });
        }
    }
    public static partial class SapInstructionParsers
    {
        public static SapConnectTo ParseSapConnectTo(JToken token)
        {
            return new SapConnectTo
            {
                System = token["System"]?.ToString()
            };
        }
        public static SapLogon ParseSapLogon(JToken token)
        {
            return new SapLogon
            {
                UserId = token["UserId"].ToString(),
                User = token["User"].ToString(),
                PasswordId = token["PasswordId"].ToString(),
                Password = token["Password"].ToString(),
                LangCodeId = token["LangCodeId"].ToString(),
                LangCode = token["LangCode"].ToString()
            };
        }
        public static SapStartTransaction ParseSapStartTransaction(JToken token) => new SapStartTransaction
        {
            TCode = token["TCode"].ToString()
        };
        public static SapEnter ParseSapEnter(JToken token) => new SapEnter();
        public static SapClickButton ParseSapClickButton(JToken token)
        {
            return new SapClickButton
            {
                Id = token["Id"].ToString(),
            };
        }
        public static SapSelectMenu ParseSapSelectMenu(JToken token)
        {
            return new SapSelectMenu
            {
                Id = token["Id"].ToString(),
            };
        }
        public static SapSelectTab ParseSapSelectTab(JToken token)
        {
            return new SapSelectTab
            {
                Id = token["Id"].ToString(),
                TabText = token["TabText"].ToString(),
            };
        }
        public static SapBack ParseSapBack(JToken token) => new SapBack();
        public static SapExecute ParseSapExecute(JToken token) => new SapExecute();
        public static SapSave ParseSapSave(JToken token) => new SapSave();
        public static SapSetText ParseSapSetText(JToken token)
        {
            if (AppConfig.LogLevel < 2)
            {
                $"ParseSapSetText: Id = {token["Id"].ToString()}, Value = {token["Value"].ToString()}".Dump();
            }
            return new SapSetText
            {
                Id = token["Id"].ToString(),
                Value = token["Value"].ToString()
            };
        }
        public static SapGetText ParseSapGetText(JToken token) => new SapGetText
        {
            Id = token["Id"].ToString(),
            TargetVariable = token["TargetVariable"].ToString()
        };
        public static SapDoubleClickTableLine ParseSapDoubleClickTableLine(JToken token)
        {
            return new SapDoubleClickTableLine
            {
                Id = GetTokenValue<string>(token, "Id"),
                Row = GetTokenValue<int>(token, "Row"),
                Column = GetTokenValue<int>(token, "Column"),
            };
        }
        public static SapGetTableData ParseSapGetTableData(JToken token)
        {
            return new SapGetTableData
            {
                Id = token["Id"].ToString(),
                TargetVariable = token["TargetVariable"].ToString(),
                Columns = token["Columns"].ToString(),
                Filters = token["Filters"].ToObject<List<RowFilter>>(),
            };
        }
        public static SapGetGridData ParseSapGetGridData(JToken token)
        {
            return new SapGetGridData
            {
                Id = token["Id"].ToString(),
                TargetVariable = token["TargetVariable"].ToString(),
                Columns = token["Columns"].ToString(),
                Filters = token["Filters"].ToObject<List<RowFilter>>(),
            };
        }
    }
}
