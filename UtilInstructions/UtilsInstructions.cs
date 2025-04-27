using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using AutomationVM.Core;
using Dumpify;
using Newtonsoft.Json.Linq;
using Serilog;
using Dapper;

using static AutomationVM.Core.InstructionParser;
using static AutomationVM.Core.Encryptor;
using System.Linq;
using System.Data.SqlClient;
using SqlKata.Execution;

namespace AutomationVM.UtilInstructions
{
    public class CopyAsInstruction : Instruction
    {
        //！注意 ！
        //在循环中使用此指令，必须从上下文对象动态获取参数！
        //指令参数在脚本解析时已经固定下来
        public string SrcFile { get; set; } // 在指令解析阶段已经获得了值，后续不会自动更新。
        public string DesFile { get; set; }
        public bool Overwrite { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            // 下面两句用于解析指令带过来的参数，适合一次性使用的情况
            SrcFile = await EvaluateParameterAsync(SrcFile, context) as string;
            DesFile = await EvaluateParameterAsync(DesFile, context) as string;

            // 下面的覆盖用于循环使用的情况，一次性使用时，上下文参数和指令参数是一致的。可以覆盖。
            if (context.VarsDict.ContainsKey("SrcFile"))
            {
                SrcFile = context.Vars.SrcFile;
                DesFile = context.Vars.DesFile;
            }
            //Log.Debug("Source File {SrcFile}", SrcFile);
            //Log.Debug("Dest File {DesFile}", DesFile);
            if (!File.Exists(SrcFile))
                throw new FileNotFoundException("源文件不存在", SrcFile);
            string destDir = Path.GetDirectoryName(DesFile);
            if (!Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);
            File.Copy(SrcFile, DesFile, Overwrite);

        }
    }
    public class ForEachFileInstruction : Instruction
    {
        public string Folder { get; set; }
        public string Extension { get; set; }
        public string CurrentFile { get; set; }
        public bool Recursive { get; set; }
        public List<Instruction> Actions { get; set; } = new List<Instruction>();

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Folder = Folder.Replace('/', '\\');
            await ProcessFile(Folder, Extension, Actions, context);
        }
        private async Task ProcessFile(string folder, string extension, List<Instruction> actions, ExecutionContext ctx)
        {
            Log.Information("Processing {folder}", folder);
            foreach (var f in Directory.GetFiles(folder, extension))
            {
                Log.Debug("file {file}", f);
                ctx.VarsDict[CurrentFile] = f;
                foreach (var instr in actions)
                {
                    await instr.ExecuteAsync(ctx);
                }
            }
            if (!Recursive) return;
            foreach (var dir in Directory.GetDirectories(folder))
            {
                await ProcessFile(dir, extension, actions, ctx);//递归调用
            }

        }
    }
    //public class TestInstruction : Instruction
    //{
    //    public List<RowFilter> Filters { get; set; }
    //    public override Task ExecuteCoreAsync(ExecutionContext context)
    //    {
    //        Filters.Dump();
    //        return Task.CompletedTask;
    //    }
    //}
    public class SetConnectionStringInstruction : Instruction
    {
        public string Name { get; set; }
        public string ConnectionString { get; set; }
        public bool IsEncrypied { get; set; }
        public string TargetVariable { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            ConnectionString = await EvaluateParameterAsync(ConnectionString, context) as string;
            if (IsEncrypied)
            {
                ConnectionString = Decrypt(ConnectionString);
            }
            context.VarsDict[TargetVariable] = ConnectionString;
            
        }
    }
    public class SqlServerDBTransactionInstruction : Instruction
    {
        public string ConnectionString { get; set; }
        public string TargetConn { get; set; }
        public string TargetTrns { get; set; }
        public List<Instruction> Actions { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            ConnectionString = await EvaluateParameterAsync(ConnectionString, context) as string;
            TargetConn = await EvaluateParameterAsync(TargetConn, context) as string;
            TargetTrns = await EvaluateParameterAsync(TargetTrns, context) as string;

            using (var conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                context.VarsDict[TargetConn] = conn;
                using (var trans = conn.BeginTransaction())
                {
                    context.VarsDict[TargetTrns] = trans;
                    try
                    {
                        foreach (var inst in Actions)
                        {
                            var insName = inst.GetType().Name.Replace("Instruction", "");
                            Log.Information("执行指令类型: {Type}", insName);
                            try
                            {
                                await inst.ExecuteAsync(context);
                            }
                            catch (Exception ex)
                            {
                                Log.Error("{inst} {caller} {message}", insName, CallerInfo(), ex.Message);
                                throw; // 外围的try捕捉到异常才会回滚操作！
                            }

                        }
                        trans.Commit();
                        Log.Information("Database transaction successed.");
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        Log.Error("{caller} {error}", CallerInfo(), ex.Message);

                    }
                }

            }

        }
    }
    public class ExecuteRawSqlInstruction : Instruction
    {
        public string ConnectionString { get; set; }
        public string SqlStatement { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public string TargetVariable { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            ConnectionString = await EvaluateParameterAsync(ConnectionString, context) as string;

            using (var connection = new SqlConnection(ConnectionString))
            {
                await connection.OpenAsync();
                // 动态解析参数值
                var dynamicParams = new DynamicParameters();
                foreach (var param in Parameters ?? new Dictionary<string, object>())
                {
                    dynamicParams.Add(
                        param.Key,
                        await EvaluateParameterAsync(param.Value?.ToString(), context)
                    );
                }

                try
                {
                    var result = await connection.QueryAsync(
                        await EvaluateParameterAsync(SqlStatement, context) as string,
                        dynamicParams
                    );

                    context.VarsDict[TargetVariable] = result
                        .Select(r => (IDictionary<string, object>)r)
                        .ToList();

                    Log.Information("执行SQL成功: {rows} 行受影响", result.Count());
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "SQL执行失败: {sql}", SqlStatement);
                    throw new SqlExecutionException(SqlStatement, ex);
                }
            }

        }
    }

    // 自定义异常类
    public class SqlExecutionException : Exception
    {
        public string Sql { get; }

        public SqlExecutionException(string sql, Exception inner)
            : base($"SQL执行失败: {sql}", inner)
        {
            Sql = sql;
        }
    }
    //****************************************************************
    //public class RowFilter // 表格行数据的筛选条件参数 2025-3-27
    //{
    //    public int ColumnIndex { get; set; }
    //    public string FilterRule { get; set; }
    //    public string Value { get; set; }
    //}

    public static class UtilsInstructionParsers
    {
        //public static TestInstruction ParseTest(JToken token)
        //{
        //    return new TestInstruction
        //    {
        //        Filters = token["Filters"].ToObject<List<RowFilter>>()
        //    };
        //}
        public static CopyAsInstruction ParseCopyAs(JToken token)
        {
            Log.Debug("Parsing CopyAs {SrcFile} {DesFile}", token["SrcFile"].ToString(), token["DesFile"].ToString());
            return new CopyAsInstruction
            {
                SrcFile = token["SrcFile"].ToString(),
                DesFile = token["DesFile"].ToString(),
                Overwrite = token["Overwrite"].ToObject<bool>(),
            };
        }
        public static ForEachFileInstruction ParseForEachFile(JToken token)
        {
            try
            {
                return new ForEachFileInstruction
                {
                    Folder = token["Folder"].ToString(),
                    Extension = token["Extension"].ToString(),
                    CurrentFile = token["CurrentFile"].ToString(),
                    Recursive = token["Recursive"].ToObject<bool>(),
                    Actions = ParseInstructions(GetTokenValue<JArray>(token, "Actions"))
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), iex);
            }
        }
        public static SetConnectionStringInstruction ParseSetConnectionString(JToken token)
        {
            try
            {
                return new SetConnectionStringInstruction
                {
                    Name = GetTokenValue<string>(token, "Name"),
                    ConnectionString = GetTokenValue<string>(token, "ConnectionString"),
                    TargetVariable = GetTokenValue<string>(token, "TargetVariable"),
                    IsEncrypied = GetTokenValue<bool>(token, "IsEncrypied")
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ExecuteRawSqlInstruction ParseExecuteRawSql(JToken token)
        {
            try
            {
                return new ExecuteRawSqlInstruction
                {
                    ConnectionString = GetTokenValue<string>(token, "ConnectionString"),
                    SqlStatement = GetTokenValue<string>(token, "SqlStatement"),
                    TargetVariable = GetTokenValue<string>(token, "TargetVariable"),
                    Parameters = GetTokenValue<Dictionary<string, object>>(token, "Parameters", false)
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), iex);
            }
        }
        public static SqlServerDBTransactionInstruction ParseSqlServerDBTransaction(JToken token)
        {
            try
            {
                return new SqlServerDBTransactionInstruction
                {
                    ConnectionString = GetTokenValue<string>(token, "ConnectionString"),
                    TargetConn = GetTokenValue<string>(token, "TargetConn"),
                    TargetTrns = GetTokenValue<string>(token, "TargetTrns"),
                    Actions = ParseInstructions(GetTokenValue<JArray>(token, "Actions"))
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("Exception when Get Token Values", CallerInfo(), iex);
            }
        }
    }

}
