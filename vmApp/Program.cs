
using System.Threading.Tasks;
using Serilog;
using Microsoft.Extensions.DependencyInjection;
using AutomationVM.Core;
using System;
using Dumpify;
using AutomationVM.Module.PlaywrightIns;
using System.Collections.Generic;
using System.IO;
using AutomationVM.Module.SAPInstructions;
using vmApp;
using System.Text.RegularExpressions;

namespace vmApp
{
    /* *
     * 
     * 多用途自动化虚拟机
     * 目前第一个用途：SAP GuiScripting
     * 
     * 2025-3-20
     * 
     * 阶段性成果：
     *  通用指令控制类型，If，While，ForEach，Code
     *  Roslyn表达式解析引擎，非常强大
     *  脚本格式支持.json,.yaml
     *  所有的指令代码在执行时会在基类的Try函数里捕获异常并记录，但不会终止程序。
     *  遗留的赋值语句做个纪念吧，有了表达式引擎，赋值语句没有存在的必要了。
     *  表达式解析分三种类型：1，字面量，而$前缀代表变量引用，会从VarsDict中寻找，=前缀表示需要Roslyn计算
     *  后两者可以合二为一，就是说$x = VarsDict["x"], 那么可以用 =Vars.x 代替。Vars是ExpandoObject的dynamic形式
     *  VarsDict是ExpandoObject 的字典形式，二者是同一个对象。
     *  
     *  ToDo：实现Continue及Break指令，
     * 2025-3-24
     * 
     * 
     */

    class Program
    {
        static async Task Main(string[] args)
        {
            Console.Clear();
            Log.Logger = LoggerConfig.CreateLogger();

            var script = "";
            if (args.Length > 0)
            {
                script = args[0];
            }
            else //无参数启动，一般是在开发中测试
            {
                await Scripts.SapScriptAsync();
                return;
            }
            /*
             * 不同的指令集需要执行上下文参数ExecutionContex提供不同的服务
             * 比如Sap指令需要SAPActive类提供的获取GuiSession的功能，而
             * Playwright则需要其它功能以启动浏览器，下面就对这些服务
             * 进行注册。以SapConnectToInstruction为例，它需要获取一个GuiSession
             * var s = context.GetService<SAPActive>().GetSession(System);
             * 
             * 2025-3-27
             */
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                //Log.Fatal(e.ExceptionObject as Exception, "全局未处理异常");
                var ex = e.ExceptionObject as Exception;
                //ExceptionFormatter.FormatException(ex).Dump();
                ex.Dump();
                //ex.ToCleanStackTrace().Dump();
                Environment.Exit(1);
            };
            //throw new IOException("Test");
            var services = new ServiceCollection()
                .AddTransient<SAPActive, SAPActive>();
            //根据此主程序需要用到的指令集（添加指令集引用）
            //下面是把指令集中提供的指令解析类中的指令解析方注册到一个静态字典中。
            InstructionParser.BuildParserCache(typeof(AutomationVM.Module.SAPInstructions.SapInstructionParsers));
            InstructionParser.BuildParserCache(typeof(AutomationVM.UtilInstructions.UtilsInstructionParsers));
            InstructionParser.BuildParserCache(typeof(PlaywrightInstructionParsers));
            var vm = new MultiplatformVM(); // 内部会注册CoreInstruction
            vm.SetServiceProvider(services.BuildServiceProvider());

            script = FindFile(script, new[] { ".yaml", });
            await vm.ExecuteScript(script);

        }
        static string FindFile(string fileName, IEnumerable<string> extList)
        { //2025-4-19
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("文件名为空！");
            }
            if (Path.HasExtension(fileName))
            {
                if (File.Exists(fileName))
                {
                    return fileName;
                }
                else if (File.Exists("test_" + fileName))
                {
                    return "test_" + fileName;
                }

            }
            foreach (var ext in extList)
            {
                var fullName = fileName + ext;
                if (File.Exists(fullName))
                {
                    return fullName;
                }
                else if (File.Exists("test_" + fullName))
                {
                    return "test_" + fullName;
                }
            }
            throw new FileNotFoundException();
        }

    }
    public static class LoggerConfig
    {
        public static ILogger CreateLogger()
        {
            var logBuilder = new LoggerConfiguration()
                .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File("logs/vm.log",
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 3)
                ;
            switch (AppConfig.LogLevel)
            {
                case 0: logBuilder.MinimumLevel.Verbose(); break;
                case 1: logBuilder.MinimumLevel.Debug(); break;
                case 2: logBuilder.MinimumLevel.Information(); break;
                case 3: logBuilder.MinimumLevel.Warning(); break;
                case 4: logBuilder.MinimumLevel.Error(); break;
                case 5: logBuilder.MinimumLevel.Fatal(); break;
            }

            return logBuilder.CreateLogger();
        }
    }

}
public static class ExceptionExtensions
{
    public static object ToCleanStackTrace(this Exception ex)
    {
        var cleanedStackTrace = Regex.Replace(
            ex.StackTrace ?? "",
            @"(in\s+)(?:.*[\\/])?([^\\/]+\.cs)(:line\s+\d+)",
            "$1$2$3"
        );

        return new
        {
            ex.HelpLink,
            ex.HResult,
            ex.Message,
            ex.Source,
            ex.TargetSite,
            ex.Data,
            CleanStackTrace = cleanedStackTrace,
            InnerException = ex.InnerException?.ToCleanStackTrace()
        };
    }
}
class Scripts
{
    /*
     * 在代码中写脚本,可以进一步体会抽象与隔离，以及上下文。
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
            Filters = new List<RowFilter> { new RowFilter { ColumnIndex = 3, FilterRule = "Contains", Value = "" } },
        });
        instructions.Add(new CodeInstruction() { Statements = @"=VarsDict[""tableData""].Dump();" });

        var services = new ServiceCollection()
            .AddTransient<SAPActive, SAPActive>();

        var vm = new MultiplatformVM();
        vm.SetServiceProvider(services.BuildServiceProvider());
        await vm.ExecuteCore(instructions);
    }
}