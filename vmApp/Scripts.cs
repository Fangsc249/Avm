using AutomationVM.Core;
using AutomationVM.Module.SAPInstructions;
using AutomationVM.Module.PlaywrightIns;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace vmApp
{
    class Scripts
    {
        /*
         * 在代码中写脚本,体会抽象与隔离，以及上下文。
         * 曾记得，当初在南京学习WWF时，对其中的上下文概念不能透彻理解。
         * 体会每个功能都用统一的模式进行封装带来的复用便利性，以及性能损失。
         * 2025-4-22
         */
        public static async Task CreateScriptAsync()
        {
            //以代码的形式创建脚本
            var instCode = new CodeInstruction() { Statements = $@"
                =
                Console.WriteLine(""Hello world!"");
                VarsDict[""Message2""] = ""嗯，Yaml确实是一个好用的工具！"";
            " };
            var instStringAssign = new StringAssignInstruction()
            {
                Variable = "Message",
                LiteralValue = "Yaml 这是一个好用的工具！"
            };
            var instExpressionAssign = new ExpressionAssignInstruction()
            {
                Variable = "x",
                Expression = "=2+3+5"
            };
            var instCode2 = new CodeInstruction()
            {
                Statements = "= Vars.x2 = 2+3+5;VarsDict.Dump();",
            };
            List<Instruction> instructions = new List<Instruction>
            {
                instCode,instStringAssign,instCode2
            };
            var ctx = new ExecutionContext();
            foreach (var inst in instructions)
            {
                await inst.ExecuteAsync(ctx);
            }
        }
        public static async Task SapScriptAsync()
        {
            List<Instruction> instructions = new List<Instruction>();
            instructions.Add(new SapConnectTo() { System = "AAA-EHP6" });
            instructions.Add(new IfInstruction()
            {
                Condition = "= Vars.NeedLogon",
                thenInstructions = new List<Instruction>
                {
                    new SapLogon()
                    {
                        UserId=SAPID.ID_USERID, User="LRP",
                        PasswordId=SAPID.ID_PASSWORD, Password="123456",
                        LangCodeId=SAPID.ID_LANGCODE, LangCode="EN",
                    }
                }
            });
            instructions.Add(new SapStartTransaction() { TCode = "CS03" });
            instructions.Add(new SapSetText() { Id = SAPID.ID_CS03_MATERIAL, Value = "DPC6" });
            instructions.Add(new SapSetText() { Id = SAPID.ID_CS03_PLANT, Value = "1200" });
            instructions.Add(new SapSetText() { Id = SAPID.ID_CS03_BOM_USAGE, Value = "1" });
            instructions.Add(new SapEnter());
            instructions.Add(new SapGetTableData()
            {
                Id = SAPID.ID_CS03_BOM_TABLE,
                TargetVariable = "tableData",
                Columns = "0,1,2,3,4,5,8",
                Filters = new List<RowFilter> {
                new RowFilter { ColumnIndex = 3, FilterRule = "Contains", Value = "" } },
            });
            instructions.Add(new CodeInstruction() { Statements = @"=VarsDict[""tableData""].Dump();" });

            var services = new ServiceCollection()
                .AddTransient<SAPActive, SAPActive>();

            var vm = new MultiplatformVM();
            vm.SetServiceProvider(services.BuildServiceProvider());
            await vm.ExecuteCore(instructions);
        }
        public static async Task PWScriptAsync()
        {
            List<Instruction> insts = new List<Instruction>();

            //
            insts.Add(new OpenBrowserInstruction { BrowserType = "Chromium" });
            insts.Add(new NavigateToInstruction { TimeoutS = 160, Url = @"https://service.danfoss.net/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dservice_offering%253D61bbf581c3c319507652176ce0013165%255EstateNOT%2520IN6%252C7%252C8%255Eassigned_toISEMPTY%26sysparm_first_row%3D1%26sysparm_view%3D" });
            insts.Add(new PageQuerySelectorInstruction { Selector = "macroponent-f51912f4c700201072b211d4d8c26010" });
            insts.Add(new ElementQuerySelectorInstruction { Selector = "#gsft_main" });
            insts.Add(new ElementContentFrameInstruction());
            insts.Add(new FrameQuerySelectorInstruction { Selector = "#incident_table" });
            insts.Add(new ElementQuerySelectorAllInstruction { Selector = "tbody > tr" });
            insts.Add(new CodeInstruction { Statements= @"=
                var headDict = new Dictionary<int,string>(){
                    [2] = ""IncidentId"",
                    [4] = ""Desc"",
                    [5] = ""Client"",
                    [6] = ""Priority"",
                    [7] = ""State"",
                    [8] = ""Category"",
                    [9] = ""Assignment Group"",
                };
                VarsDict[""headDict""] = headDict;
            " });
            insts.Add(new SaveTableContentInstruction() { CellType = "td",Columns="2,4,5,6" });
            insts.Add(new CodeInstruction { Statements = @"=
                VarsDict[""dataList""].Dump();
                //VarsDict[""headDict""].Dump();
            " });
            //
            var services = new ServiceCollection()
                .AddTransient<SAPActive, SAPActive>();
            var vm = new MultiplatformVM();
            vm.SetServiceProvider(services.BuildServiceProvider());
            await vm.ExecuteCore(insts);

        }
        public static async Task ScriptTemplateAsync()
        {
            List<Instruction> instructions = new List<Instruction>();

            var services = new ServiceCollection()
                .AddTransient<SAPActive, SAPActive>();

            var vm = new MultiplatformVM();
            vm.SetServiceProvider(services.BuildServiceProvider());
            await vm.ExecuteCore(instructions);
        }
    }
}
