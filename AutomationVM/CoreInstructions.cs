using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Serilog;
using System.Dynamic;
using static AutomationVM.Core.InstructionParser;
using System.Text.RegularExpressions;
/*
* 设计过程中想法多次反复。
* 这里对SAP界面操作的指令化封装要做到粒度最小化才能实现最大的复用度。
* 一个指令操作所需要输入的ID数量不能超过1， 否则又会回到老路上去。
* 2025-3-23
* 
* 周末又到了。小语上午艺东画画课时准备完善异常处理与错误检查
* 因为要解析脚本文件，本程序相当于一个解释器，排查问题的机制需要完善。
* Your app has entered a break state, 
* but there is no code to show because all threads were executing external code 
* (typically system or framework code).
* 上述这样的情况很难处理，这是异步模式导致的吧。
* 2025-3-29
* 
*/


namespace AutomationVM.Core
{
    public class BreakInstruction : Instruction
    {
        public override Task ExecuteCoreAsync(ExecutionContext context)
        {
            if (context.CurrentLoopState == null)
                throw new InvalidOperationException("Break指令只能在循环内部使用");

            context.CurrentLoopState.ShouldBreak = true;
            Log.Debug("触发循环中断");
            return Task.CompletedTask;
        }
    }
    public class ContinueInstruction : Instruction
    {
        public override Task ExecuteCoreAsync(ExecutionContext context)
        {
            if (context.CurrentLoopState == null)
                throw new InvalidOperationException("Continue指令只能在循环内部使用");

            context.CurrentLoopState.ShouldContinue = true;
            Log.Debug("触发继续下一轮迭代");
            return Task.CompletedTask;
        }
    }
    public class WhileInstruction : Instruction
    {
        public string Condition { get; set; }
        public List<Instruction> Instructions { get; set; } = new List<Instruction>();
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            context.PushLoopState(); // 进入循环作用域
            try //2025-3-30 
            {
                while (true)
                {
                    bool conditionResult = Convert.ToBoolean(
                        await EvaluateParameterAsync(Condition, context)
                    );

                    if (!conditionResult) break;

                    foreach (var instr in Instructions)
                    {
                        await instr.ExecuteAsync(context);

                        // 处理循环控制状态
                        if (context.CurrentLoopState.ShouldBreak)
                        {
                            context.PopLoopState();
                            return; // 完全退出循环
                        }
                        if (context.CurrentLoopState.ShouldContinue)
                            break;   // 跳过当前迭代剩余指令
                    }

                    // 重置迭代状态
                    context.CurrentLoopState.ShouldContinue = false;
                }
            }
            finally
            {
                context.PopLoopState(); // 退出循环作用域
            }

        }
    }
    public class ForEachInstruction : Instruction
    {
        public string CollectionExpression { get; set; }
        public string ItemVariable { get; set; }
        public List<Instruction> Instructions { get; set; } = new List<Instruction>();

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            var collection = await EvaluateParameterAsync(CollectionExpression, context) as IEnumerable;
            if (collection == null) throw new InvalidOperationException("集合表达式无效");

            context.PushLoopState();//2025-3-25
            try //2025-3-30
            {
                int index = 0;
                foreach (var item in collection) // 要遍历的对象集合
                {
                    context.VarsDict[ItemVariable] = item;
                    context.PushLoopIndex(index++);

                    foreach (var instr in Instructions)
                    {
                        await instr.ExecuteAsync(context);
                        // 确保每次迭代后清除可能的缓存
                        //context.VarsDict.Remove(ItemVariable);

                        if (context.CurrentLoopState.ShouldBreak)
                        {
                            context.PopLoopState();
                            return;
                        }
                        if (context.CurrentLoopState.ShouldContinue) break;
                    }

                    context.CurrentLoopState.ShouldContinue = false;
                }
            }
            finally
            {
                context.PopLoopState();
                Log.Debug("ForEach循环状态已清理");
            }

        }
    }
    public class LoadDataInstruction : Instruction
    {
        public string SourceType { get; set; }    // 数据源类型（JSON/CSV/DB）
        public string SourcePath { get; set; }     // 文件路径或连接字符串
        public string TargetVariable { get; set; } // 存储数据的变量名

        public override Task ExecuteCoreAsync(ExecutionContext context)
        {
            try
            {
                List<Dictionary<string, object>> data = LoadData(context);
                context.VarsDict[TargetVariable] = data; // 把加载的数据放入一个上下文变量
                context.Logger.Information("成功加载数据到变量: {Variable}, 共 {Count} 条", TargetVariable, data.Count);
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                context.Logger.Error(ex, "数据加载失败: {Source}", SourcePath);
                throw;
            }
        }

        private List<Dictionary<string, object>> LoadData(ExecutionContext context)
        {
            switch (SourceType.ToUpper())
            {
                case "JSON": return LoadJsonData();
                //case "CSV": return LoadCsvData();
                //case "DB": return LoadDatabaseData(context);
                default: throw new NotSupportedException($"不支持的数据源类型: {SourceType}");
            }
        }

        private List<Dictionary<string, object>> LoadJsonData()
        {
            string json = File.ReadAllText(SourcePath);
            return JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);
        }

        private List<Dictionary<string, object>> LoadCsvData()
        {
            var data = new List<Dictionary<string, object>>();
            using (var reader = new StreamReader(SourcePath))
                //using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                //{
                //    csv.Read();
                //    csv.ReadHeader();
                //    while (csv.Read())
                //    {
                //        var row = new Dictionary<string, object>();
                //        foreach (var header in csv.HeaderRecord)
                //        {
                //            row[header] = csv.GetField(header);
                //        }
                //        data.Add(row);
                //    }
                //}
                return data;
        }

        private List<Dictionary<string, object>> LoadDatabaseData(ExecutionContext context)
        {
            // 示例：从数据库加载（需扩展上下文支持连接池）
            using (var conn = new SqlConnection(SourcePath))
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM SourceTable", conn);
                var reader = cmd.ExecuteReader();
                var data = new List<Dictionary<string, object>>();
                while (reader.Read())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        row[reader.GetName(i)] = reader.GetValue(i);
                    }
                    data.Add(row);
                }
                return data;
            }
        }
    }
    public class IfInstruction : Instruction
    {
        public string Condition { get; set; }
        public JArray Then { get; set; }
        public JArray Else { get; set; }
        public List<Instruction> thenInstructions;
        public List<Instruction> elseInstructions;

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            if (thenInstructions == null)
            {
                thenInstructions = ParseBranch(Then, context);
            }
            if (elseInstructions == null && Else != null)
            {
                elseInstructions = ParseBranch(Else, context);
            }
            bool conditionResult = Convert.ToBoolean(
                await EvaluateParameterAsync(Condition, context)
                );
            Log.Debug("条件评估结果：{Condition} --> {Result}", Condition, conditionResult);
            var instrs = conditionResult ? thenInstructions : elseInstructions;
            if (instrs != null)
            {
                await ExecuteBranchAsync(instrs, context);
            }
        }

        /// <summary>
        /// 解析分支指令
        /// </summary>
        /// <param name="branchTokens"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        private List<Instruction> ParseBranch(JArray branchTokens, ExecutionContext ctx)
        {
            if (branchTokens == null) return null;

            return InstructionParser.ParseInstructions(branchTokens);
        }
        /// <summary>
        /// 执行分支指令
        /// </summary>
        /// <param name="instructions"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        private async Task ExecuteBranchAsync(List<Instruction> instructions, ExecutionContext ctx)
        {
            foreach (var instr in instructions)
            {
                Log.Debug("执行分支指令：{Type}", instr.GetType().Name);
                await instr.ExecuteAsync(ctx);
            }
        }
    }
    public class CodeInstruction : Instruction
    {
        /*
         * 特别注意Statemens / FileName 处理之间的差异！
         * 此处坑很多。
         * 2025-4-1
         */
        public string Statements { get; set; }
        public string FileName { get; set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            if (Statements != null && Statements != string.Empty)
            {
                await EvaluateParameterAsync(Statements, context); // 使用基类的方法
            }
            if (FileName != null)
            {
                string code = File.ReadAllText(FileName).Trim();
                if (code != string.Empty)
                {
                    //Log.Debug("code in file {code} {fileName}",code,FileName);
                    code = await EvaluateParameterAsync(code, context) as string; // 处理脚本中可能存在的$变量引用
                    //await context.EvaluateSmart(code);//这里不能使用EvaluateParameterAsync，因为脚本文件没有以=开头
                    await context.EvaluateAsync(code);
                }
                else
                {
                    Log.Error($"{FileName} does not contains code!");
                }
            }
        }
    }
    public abstract class AssignInstruction : Instruction
    {
        public string Variable { get; set; }

        protected void LogAssignment(object value)
        {
            Log.Debug("变量赋值： {Var} = {Value} (Type: {Type})",
                    Variable, value, value?.GetType().Name ?? "null");
        }
    }
    public class StringAssignInstruction : AssignInstruction
    {
        /// <summary>
        /// 直接赋值的字符串值
        /// 示例："Hello World"
        /// </summary>
        public string LiteralValue { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            // 直接使用原始字符串值
            var value = LiteralValue;
            context.VarsDict[Variable] = value;
            LogAssignment(value);

            await Task.CompletedTask; // 保持异步接口兼容性
        }
    }
    public class ExpressionAssignInstruction : AssignInstruction
    {
        /// <summary>
        /// 需要解析的表达式
        /// 示例："DateTime.Now.ToString("yyyy-MM-dd")"
        /// </summary>
        public string Expression { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            // 执行表达式解析
            var value = await EvaluateParameterAsync(Expression, context);

            context.VarsDict[Variable] = value;
            LogAssignment(value);
        }
    }
    public class CallSubflowInstruction : Instruction
    {
        public string SubflowPath { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            var fullpath = ResolePath(context, SubflowPath);
            if (!File.Exists(fullpath))
            {
                throw new FileNotFoundException($"子流程文件{fullpath}不存在。");
            }
            var origVars = new ExpandoObject();
            
            await context.VM.ExecuteSubflow(fullpath, context);

        }
        private string ResolePath(ExecutionContext context, string path)
        {
            if (Path.IsPathRooted(path))
            {
                return path;
            }
            return Path.Combine(Path.GetDirectoryName(context.CurrentScriptPath), path);
        }
    }
    public class DebugBreakInstruction : Instruction
    {
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            context.IsDebugging = true;
            context.Debugger.EnterDebugLoop(context);
            await Task.CompletedTask;
        }
    }
    public class LogVariablesInstruction : Instruction
    {
        public string Filter { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Filter = await EvaluateParameterAsync(Filter, context) as string;
            var variables = context.VarsDict
                .Where(kv => Regex.IsMatch(kv.Key, Filter))
                .ToDictionary(k => k.Key, v => v.Value);
            Log.Information("变量快照:{Variables}",
                JsonConvert.SerializeObject(variables, Formatting.Indented)
                );
            await Task.CompletedTask;
        }
    }
    public class BreakpointInstruction : Instruction
    {
        public string ScriptPath { get; set; }
        public int LineNumber { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            context.Debugger.AddBreakpoint(context, new Breakpoint
            {
                ScriptPath = ScriptPath,
                LineNumber = LineNumber
            });
            await Task.CompletedTask;
        }
    }

    public static class CoreInstructionParsers
    {
        public static IfInstruction ParseIf(JToken token)
        {
            try
            {
                return new IfInstruction
                {
                    Condition = GetTokenValue<string>(token, "Condition"),
                    Then = GetTokenValue<JArray>(token, "Then"),
                    Else = GetTokenValue<JArray>(token, "Else", false)//可选参数
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static CodeInstruction ParseCode(JToken token)
        {
            try
            {
                return new CodeInstruction
                {
                    Statements = GetTokenValue<string>(token, "Statements", false),//可选参数
                    FileName = GetTokenValue<string>(token, "FileName", false) //可选参数
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static LoadDataInstruction ParseLoadData(JToken token)
        {
            try
            {
                return new LoadDataInstruction
                {
                    SourceType = GetTokenValue<string>(token, "SourceType"),
                    SourcePath = GetTokenValue<string>(token, "SourcePath"),
                    TargetVariable = GetTokenValue<string>(token, "TargetVariable")
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static WhileInstruction ParseWhile(JToken token)
        {
            try
            {
                return new WhileInstruction
                {
                    Condition = GetTokenValue<string>(token, "Condition"),
                    Instructions = ParseInstructions(GetTokenValue<JArray>(token, "Instructions"))
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ForEachInstruction ParseForEach(JToken token)
        {

            try
            {
                return new ForEachInstruction
                {
                    CollectionExpression = GetTokenValue<string>(token, "Collection"),
                    ItemVariable = GetTokenValue<string>(token, "ItemVar"),
                    Instructions = ParseInstructionsFromToken(GetTokenValue<JArray>(token, "Instructions")) // 独立方法
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Debug("{caller}{message}", CallerInfo(), aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Debug("{caller}{message}", CallerInfo(), iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
            List<Instruction> ParseInstructionsFromToken(JArray jArray)
            {
                try
                {
                    return ParseInstructions(jArray);
                }
                catch (Exception ex) when (ex is ArgumentNullException aex)
                {
                    Log.Debug("{caller}{message}", CallerInfo(), aex.Message);
                    throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
                }
                catch (Exception ex) when (ex is InvalidCastException iex)
                {
                    Log.Debug("{caller}{message}", CallerInfo(), iex.Message);
                    throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
                }
            }
        }
        public static StringAssignInstruction ParseStringAssign(JToken token)
        {
            try
            {
                return new StringAssignInstruction
                {
                    Variable = GetTokenValue<string>(token, "Variable"),
                    LiteralValue = GetTokenValue<string>(token, "LiteralValue")
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ExpressionAssignInstruction ParseExpressionAssign(JToken token)
        {
            try
            {
                return new ExpressionAssignInstruction
                {
                    Variable = GetTokenValue<string>(token, "Variable"),
                    Expression = GetTokenValue<string>(token, "Expression")
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static BreakInstruction ParseBreak(JToken token)
        {
            return new BreakInstruction { };
        }
        public static ContinueInstruction ParseContinue(JToken token)
        {
            try
            {
                return new ContinueInstruction();
            }
            catch (Exception)
            {
                throw;
            }

        }
        public static CallSubflowInstruction ParseCallSubflow(JToken token)
        {
            try
            {
                return new CallSubflowInstruction
                {
                    SubflowPath = GetTokenValue<string>(token, "SubflowPath")
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static DebugBreakInstruction ParseDebugBreak(JToken token)
        {
            return new DebugBreakInstruction();
        }
        public static LogVariablesInstruction ParseLogVariables(JToken token)
        {
            try
            {
                return new LogVariablesInstruction
                {
                    Filter = GetTokenValue<string>(token, "Filter")
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static BreakpointInstruction ParseBreakpoint(JToken token)
        {
            try
            {
                return new BreakpointInstruction
                {
                    ScriptPath = GetTokenValue<string>(token, "ScriptPath"),
                    LineNumber = GetTokenValue<int>(token, "LineNumber"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
    }

}