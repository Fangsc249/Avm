
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
            DumpConfig.Default.TypeNamingConfig.ShowTypeNames = false;
            DumpConfig.Default.TableConfig.ShowRowSeparators = true;
            DumpConfig.Default.TypeRenderingConfig.QuoteStringValues = false;
            DumpConfig.Default.ColorConfig.PropertyValueColor = DumpColor.FromHexString("#008b8b");
            var script = "";
            if (args.Length > 0)
            {
                script = args[0];
            }
            else //无参数启动，开发中测试
            {
                await Scripts.PWScriptAsync();
                return;
            }
            /*
             * 不同的指令集需要执行上下文参数ExecutionContex提供不同的服务
             * 比如Sap指令需要SAPActive类提供的获取GuiSession的功能，而
             * Playwright则需要其它功能以启动浏览器，下面对这些服务
             * 进行注册。以SapConnectToInstruction为例，它需要获取一个GuiSession
             * var s = context.GetService<SAPActive>().GetSession(System);
             * 
             * 2025-3-27
             */
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                //Log.Fatal(e.ExceptionObject as Exception, "全局未处理异常");
                var ex = e.ExceptionObject as Exception;
                ex.ToCleanStackTrace().Dump();

                if (System.Diagnostics.Debugger.IsAttached)
                {
                    //"调试器已附加",在vs中Debug时，不会退出
                }
                else
                {
                    Environment.Exit(1);//此处注释掉，CLR会再次显示异常信息!
                }
            };
            //throw new IOException("Test");
            var services = new ServiceCollection()
                .AddTransient<SAPActive, SAPActive>();
            //根据此主程序需要用到的指令集（添加指令集引用）
            //下面是把指令集中提供的指令解析类中的指令解析方注册到一个静态字典中。
            InstructionParser.BuildParserCache(typeof(SapInstructionParsers));
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
                .WriteTo.File($"{AppDomain.CurrentDomain.BaseDirectory}/logs/vm.log",
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
