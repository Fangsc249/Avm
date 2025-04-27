using System;
using System.Collections.Generic;
using System.IO;
using Serilog;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using System.Collections.Concurrent;
using System.Linq;
using AutomationVM.Core;
using Dumpify;
using static AutomationVM.Core.InstructionParser;

public class MultiplatformVM : IDisposable
{
    private readonly ExecutionContext _context;
    private readonly static ConcurrentDictionary<string, Script<object>> scriptCache
        = new ConcurrentDictionary<string, Script<object>>();
    public Stack<string> CallStack = new Stack<string>();//for call subflow

    public MultiplatformVM(ILogger logger = null)
    {
        _context = new ExecutionContext { VM = this }; // CallSubflowInstruction 需要这个变量访问虚拟机的执行子流程函数
        //_context.VarsDict["context"] = _context;
        BuildParserCache(typeof(CoreInstructionParsers));
    }
    public void SetServiceProvider(IServiceProvider provider)
    {
        _context.ServiceProvider = provider;
    }
    public async Task<object> EvaluateWithCache(string expr, ExecutionContext ctx)
    {
        if (!scriptCache.TryGetValue(expr, out var script))
        {
            script = CSharpScript.Create(expr, ctx.scriptOptions, typeof(ExecutionContext));

            scriptCache.TryAdd(expr, script);
            Log.Debug("{expr} was added to scriptCache", expr);
        }
        else
        {
            Log.Debug("{expr} was gotten from cache", expr);
        }
        return (await script.RunAsync(globals: ctx)).ReturnValue;
    }
    public void WarmUpExpressions(params string[] exprs)
    {
        Parallel.ForEach(exprs, expr =>
        {
            Log.Debug("Precompiling {expression} ", expr);
            var script = CSharpScript.Create(expr, _context.scriptOptions);
            script.Compile();
        });
    }
    public async Task ExecuteScript(string scriptPath)
    {
        if (!File.Exists(scriptPath))
        {
            Log.Error("{script} not exist!", scriptPath);
            return;
        }
        Log.Information("开始执行脚本: {ScriptPath}", scriptPath);
        _context.CurrentScriptPath = scriptPath;

        var parser = new ScriptParserFactory().CreateParser(scriptPath);

        using (var sr = File.OpenRead(scriptPath))
        {
            var instructions = parser.Parse(sr);

            //instructions.Dump($"{scriptPath}");

            var hotExpressions = ExtractHotExpressions(instructions); // 提取高频表达式
            WarmUpExpressions(hotExpressions);
            await ExecuteCore(instructions);
        }

        Log.Information("脚本执行完毕！");
    }
    public async Task ExecuteSubflow(string path, ExecutionContext parentContext)
    {
        if (CallStack.Count > 12)
        {
            throw new StackOverflowException($"调用深度超过限制:{CallStack.Count}");
        }
        if (!File.Exists(path))
        {
            Log.Error("{script} not exist!", path);
        }
        Log.Information("开始执行脚本: {ScriptPath}", path);
        CallStack.Push(path);
        parentContext.CurrentScriptPath = path;
        var parser = new ScriptParserFactory().CreateParser(path);
        using (var sr = File.OpenRead(path))
        {
            var instructions = parser.Parse(sr);
            foreach (var inst in instructions)
            {
                var para = inst.ExePara;
                para.Context = new { Script = path };
                await inst.ExecuteAsync(parentContext);
                
            }
            CallStack.Pop();
        }
    }
    private string[] ExtractHotExpressions(List<Instruction> instructions)
    {
        var exprs = new HashSet<string>();
        foreach (var instr in instructions)
        {
            if (instr is CodeInstruction codeInstr && codeInstr.Statements != null)
            {
                exprs.Add(codeInstr.Statements);
            }
        }
        return exprs.ToArray();
    }
    public async Task ExecuteCore(List<Instruction> instructions)
    {
        if (instructions.Count > 0)
            Log.Information("开始执行指令序列，总指令数: {Count}", instructions.Count);
        else
        {
            Log.Information("指令数为零.");
            return;
        }

        foreach (var instr in instructions)
        {
            await instr.ExecuteAsync(_context);
        }
        //下面是实现Debug的机制，目前没有成功。
        //int pc = 0;
        //while (pc < instructions.Count)// 逻辑有问题，单步执行要么死循环，要么和continue一样，待修复。
        //{
        //    bool shouldPause = false;
        //    // 1. 检查单步模式
        //    if (_context.Debugger.IsStepping)
        //    {
        //        _context.IsDebugging = true;
        //        _context.Debugger.EnterDebugLoop(_context);
        //        //_context.Debugger.IsStepping = false; // 执行后立即关闭单步模式
        //        shouldPause = true;
        //    }

        //    // 2. 断点检查逻辑（原有代码）
        //    if (_context.Breakpoints.Any(bp =>
        //        bp.ScriptPath == _context.CurrentScriptPath &&
        //        bp.LineNumber == pc))
        //    {
        //        _context.IsDebugging = true;
        //        _context.Debugger.EnterDebugLoop(_context);
        //        shouldPause = true;
        //    }

        //    // 3. 执行当前指令
        //    if (!shouldPause)
        //    {
        //        await instructions[pc].ExecuteAsync(_context);
        //        pc++;
        //    }

        //}
    }
    public void Dispose() => _context.VarsDict.Clear();
}

