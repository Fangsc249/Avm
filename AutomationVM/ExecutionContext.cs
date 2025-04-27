using Dumpify;
using Serilog;
using SAPFEWSELib;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;

using System;
using System.Dynamic;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.DependencyInjection;


namespace AutomationVM.Core
{
    public class ServiceNotFoundException : System.Exception
    {
        public ServiceNotFoundException(System.Type serviceType)
            : base($"服务 {serviceType.Name} 未注册") { }
    }
    public class Breakpoint
    {
        public string ScriptPath { get; set; }
        public int LineNumber { get; set; }
        public string Condition { get; set; }
    }
    public class DebuggerService // 逻辑有问题，单步执行要么死循环，要么和continue一样，待修复。
    {
        public bool IsStepping { get; set; }
        public void EnterDebugLoop(ExecutionContext context)
        {
            while (context.IsDebugging)
            {
                Console.Write("(AVM Debug)>>>");
                var command = Console.ReadLine().Trim();
                var parts = command.Split(' ');

                switch (parts[0].ToLower())
                {
                    case "s":
                    case "step":
                        IsStepping = true;
                        context.IsDebugging = false;
                        return;//返回后执行下一条指令。
                    case "c":
                    case "continue":
                        IsStepping = false;
                        context.IsDebugging = false;
                        break;
                    case "b":
                    case "break":
                        if (parts.Length >= 3 && int.TryParse(parts[2], out int line))
                        {
                            context.Debugger.AddBreakpoint(context,
                                new Breakpoint
                                {
                                    ScriptPath = parts[1],
                                    LineNumber = line,
                                }
                            );
                            Console.WriteLine($"断点已添加:{parts[1]}:{line}");
                        }
                        break;
                    case "db":
                    case "delete-break":
                        if (parts.Length >= 3 && int.TryParse(parts[2], out int lineToRemove))
                        {
                            context.Debugger.RemoveBreakpoint(context, parts[1], lineToRemove);
                            Console.WriteLine($"断点已删除：{parts[1]}:{lineToRemove}");
                        }
                        break;
                    case "lb":
                    case "list-breaks":
                        foreach (var bp in context.Breakpoints)
                        {
                            Console.WriteLine($"断点：{bp.ScriptPath}:{bp.LineNumber}");
                        }
                        break;
                    case "vars": LogVariables(context); break;

                }
            }
        }
        public void AddBreakpoint(ExecutionContext context, Breakpoint bp)
            => context.Breakpoints.Add(bp);

        public void RemoveBreakpoint(ExecutionContext context, string scriptPath, int lineNumber)
            => context.Breakpoints.RemoveAll(bp =>
                bp.ScriptPath == scriptPath && bp.LineNumber == lineNumber);
        private void LogVariables(ExecutionContext ctx)
        {
            ctx.VarsDict.Dump("当前变量状态：");
        }
    }
    public class ExecutionContext
    {

        public IServiceProvider ServiceProvider { get; set; }
        public T GetService<T>() where T : class
         => ServiceProvider.GetService<T>() ?? throw new ServiceNotFoundException(typeof(T));

        public bool IsDebugging { get; set; }//2025-4-1
        public DebuggerService Debugger { get; set; } = new DebuggerService();
        public List<Breakpoint> Breakpoints { get; } = new List<Breakpoint>();

        public MultiplatformVM VM { get; set; }
        public ScriptState<object> ScriptSate { get; set; } //2025-3-22 中午发现脚本引擎的威力！
        public string CurrentScriptPath { get; set; }
        /*
         * ！下面两个变量的使用要特别注意！
         * Vars 因为声明成了dynamic,不可以在表达式中引用，因为Roslyn无法在编译表达式时知道它的类型
         * 因此：Vars可在指令的执行上下文中进行赋值context.Vars.xxx = yyy;
         * 在表达式中可以这样 var other = VarsDict["xxx"];或者 var other = $xxx;
         * Instruction.cs 中解析表达式逻辑现已修改为 expr = Regex.Replace(expr, @"\$([a-zA-Z_]\w*)", "VarsDict[\"$1\"]");
         * 2025-3-31
         */
        public dynamic Vars { get; } = new ExpandoObject();
        public IDictionary<string, object> VarsDict => Vars;


        public Stack<int> LoopStack { get; } = new Stack<int>();

        public Stack<LoopState> _loopStates = new Stack<LoopState>();
        public LoopState CurrentLoopState => _loopStates.Count > 0 ? _loopStates.Peek() : null;

        public void PushLoopState(int index = 0)
        {
            _loopStates.Push(new LoopState() { CurrentIndex = index });
            Log.Debug("进入循环层级：{level}", _loopStates.Count);
        }
        public void PopLoopState()
        {
            if (_loopStates.Count > 0)
            {
                _loopStates.Pop();
                Log.Debug("退出循环层级:{level}", _loopStates.Count + 1);
            }
        }
        public ILogger Logger { get; set; } = Log.Logger; // 默认使用全局Logger
        public bool ShouldBreak { get; set; }
        public bool ShouldContinue { get; set; }
        
        public class LoopState
        {
            public bool ShouldBreak { get; set; }
            public bool ShouldContinue { get; set; }
            public int CurrentIndex { get; set; }
        }
        public ScriptOptions scriptOptions;
        public ExecutionContext()
        {
            var dumpifyAssembly = typeof(DumpConfig).Assembly;
            var serilogAssembly = typeof(Log).Assembly;
            var sapguiAssembly = typeof(GuiSession).Assembly;

            scriptOptions = ScriptOptions.Default
                .AddReferences(
                    typeof(System.Linq.Enumerable).Assembly,
                    typeof(Newtonsoft.Json.JsonConvert).Assembly,
                    typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo).Assembly,
                    typeof(SqlKata.AbstractClause).Assembly,
                    typeof(SqlKata.Execution.QueryFactory).Assembly,
                    typeof(Dapper.SqlMapper).Assembly,
                    typeof(ExecutionContext).Assembly
                )
                .AddReferences(
                    MetadataReference.CreateFromFile(dumpifyAssembly.Location),
                    MetadataReference.CreateFromFile(serilogAssembly.Location),
                    MetadataReference.CreateFromFile(sapguiAssembly.Location),
                    MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
                //MetadataReference.CreateFromFile(@"C:\Users\u324926\OneDrive - Danfoss\VS\Wpf\Tools\Lib\Interop.SAPFEWSELib.dll")
                ).AddImports(
                      "System", "System.IO", "System.Linq"
                    , "System.Collections.Generic", "System.Threading.Tasks"
                    , "Newtonsoft.Json", "Dumpify", "Serilog", "SAPFEWSELib"
                    , "SqlKata","SqlKata.Compilers", "SqlKata.Execution", "Dapper"
                    , "System.Data.SqlClient", "AutomationVM.Core");
        }
        public void PushLoopIndex(int index)
        {
            LoopStack.Push(index);
            Log.Debug("压入循环索引: {Index}, 当前堆栈深度: {Depth}", index, LoopStack.Count);
        }
        public void PopLoopIndex()
        {
            if (LoopStack.Count > 0)
            {
                var index = LoopStack.Pop();
                Log.Debug("弹出循环索引: {Index}, 剩余堆栈深度: {Depth}", index, LoopStack.Count);
            }
        }
        public int CurrentLoopIndex
        {
            get
            {
                int value = LoopStack.Count > 0 ? LoopStack.Peek() : -1;
                Log.Debug("访问 CurrentLoopIndex: {Value}", value);
                return value;
            }
        }
        public async Task<object> EvaluateAsync(string expression)
        {
            // return await EvaluateSmart(expression); // 2025-3-25
            // 用Code指令加载一个.cs文件，在其中定义一些类型，然后在后续的Code指令种使用这些类型
            // 依赖下面这种保存状态的表达式解析方式。
            return await EvaluateWithStateAsync(expression);
        }
        public async Task<object> EvaluateWithStateAsync(string expression)
        {
            // 两种方式：
            //globals: this
            //globalsType: typeof(ExecutionContext)
            var newState =
                ScriptSate?.ContinueWithAsync(expression, scriptOptions)
                ?? CSharpScript.RunAsync(expression, scriptOptions, globals: this);
            ScriptSate = await newState;
            return ScriptSate.ReturnValue;
        }
        // 在 ExecutionContext 中实现智能缓存
        public async Task<object> EvaluateSmart(string code)
        {
            // 高频表达式使用缓存
            if (IsHotExpression(code))
            {
                return await VM.EvaluateWithCache(code, this);
            }
            // 需要状态连续的场景
            else
            {
                return await EvaluateWithStateAsync(code);
            }
        }

        // 在 MultiplatformVM 中实现频率检测
        private bool IsHotExpression(string expr)
        {
            // 实现基于统计的检测逻辑（示例：长度小于50且不包含赋值）
            return expr.Length < 50 && !expr.Contains("=");
        }
    }
}
