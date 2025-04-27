using Dumpify;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace AutomationVM.Core
{
    public static class Encryptor
    {
        /*
         * internal 类只能在本程序集内部使用
         * 2025-2-21
         */
        static string password = "noshare";

        public static string Encrypt(string toEncrypt)
        {
            return Convert.ToBase64String(EncryptString(toEncrypt));
        }
        public static string Decrypt(string base64Str)
        {
            return DecryptToString(Convert.FromBase64String(base64Str));
        }

        //******************************************************************
        public static byte[] EncryptString(string toEncrypt)
        {
            var key = GetKey(password);

            using (var aes = Aes.Create())
            using (var encryptor = aes.CreateEncryptor(key, key))
            {
                var plainText = Encoding.UTF8.GetBytes(toEncrypt);
                return encryptor
                    .TransformFinalBlock(plainText, 0, plainText.Length);
            }
        }

        public static string DecryptToString(byte[] encryptedData)
        {
            var key = GetKey(password);

            using (var aes = Aes.Create())
            using (var encryptor = aes.CreateDecryptor(key, key))
            {
                var decryptedBytes =
                    encryptor
                        .TransformFinalBlock(encryptedData,
                        0,
                        encryptedData.Length);
                return Encoding.UTF8.GetString(decryptedBytes);
            }
        }

        // converts password to 128 bit hash
        private static byte[] GetKey(string password)
        {
            var keyBytes = Encoding.UTF8.GetBytes(password);
            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(keyBytes);
            }
        }
    }
    // InstructionParser.cs 
    // 可视为CPU的指令解析部件
    //2025-3-27
    public static class InstructionParser
    {
        public static readonly Dictionary<string, Func<JToken, Instruction>> parserDict
            = new Dictionary<string, Func<JToken, Instruction>>();

        public static T GetTokenValue<T>(JToken token, string key, bool required = true)
        {
            //原来的解析方法工作的好好的，新增的yaml格式会转为JToken进行解析 2025-3-24 
            var valueToken = token[key];

            // 空值检查（required=true时强制校验）
            if (valueToken == null)
            {
                if (required)
                {
                    var errmsg = $"{CallerInfo()},指令缺少必要参数: {key}，" +
                          $"完整Token结构: {token.ToString(Formatting.None)}";
                    throw new ArgumentNullException(errmsg);
                }

                return default;
            }

            // 自动处理JArray/JObject类型
            try
            {
                if (typeof(T) == typeof(JArray) && valueToken is JArray jarr)
                    return (T)(object)jarr;

                if (typeof(T) == typeof(JObject) && valueToken is JObject jobj)
                    return (T)(object)jobj;

                return valueToken.ToObject<T>();
            }
            catch (Exception ex)
            {
                var errmsg = $"指令参数类型转换失败 [Key={key}, Expected={typeof(T).Name}, Actual={valueToken.Type}]，" +
                    $"原始值: {valueToken.ToString(Formatting.None)}";
                throw new InvalidCastException(
                    errmsg,
                    ex
                );
            }
        }
        public static Dictionary<string, Func<JToken, Instruction>> BuildParserCache(Type type)
        {
            var methods = type.GetMethods(BindingFlags.Static | BindingFlags.Public);

            foreach (var method in methods)
            {
                //Log.Debug("发现方法: {MethodName} (静态: {IsStatic})", method.Name, method.IsStatic);
                if (method.Name.StartsWith("Parse") &&
                    method.ReturnType.IsSubclassOf(typeof(Instruction)))
                {

                    var key = method.Name.Substring("Parse".Length);
                    var del = (Func<JToken, Instruction>)Delegate.CreateDelegate(
                        typeof(Func<JToken, Instruction>), method);
                    if (!parserDict.ContainsKey(key))
                    {
                        parserDict.Add(key, del);
                    }
                }
            }
            return parserDict;
        }
        public static List<Instruction> ParseInstructions(JArray jArray)
        {
            if (jArray is null)
            {
                throw new NullReferenceException("JArray is null.", CallerInfo()) { ObjectName = nameof(jArray) };
            }
            var instructions = new List<Instruction>();
            foreach (var token in jArray)
            {
                try
                {
                    Instruction instr = ParseInstruction(token);
                    instructions.Add(instr);
                }
                catch (Exception ex) when (ex is ParseInstructionException pie)
                {
                    Log.Error("{0}{1}", pie.CallerInfo, pie.Message);
                }
                catch (Exception ex) when (ex is NotSupportedException nse)
                {
                    Log.Error("{caller} {msg}", CallerInfo(), nse.Message);
                }
            }
            return instructions;
        }
        public static Instruction ParseInstruction(JToken token)
        {
            var typeKey = token["Type"]?.ToString();
            if (typeKey is null)
            {
                throw new NotSupportedException($"指令格式错误,没有Type字段！");
            }
            if (parserDict.TryGetValue(typeKey, out var parser))
            {
                try
                {
                    return parser(token);
                }
                catch (Exception ex) when (ex is JTokenParseException jtp)
                {
                    Log.Error("{caller} {message}", jtp.CallerInfo, jtp.Message);
                    throw new ParseInstructionException("JTokenParseException when create instruction", CallerInfo()) { };
                }
            }
            throw new NotSupportedException($"不支持的指令类型: {typeKey}");
        }
        public static async Task Try(string callerInfo, Func<Task> asyncAction, bool throwOnError = true)
        {
            try
            {
                await asyncAction();
            }
            catch (Exception ex) when (!throwOnError)
            {
                Log.Error($"{callerInfo}: {ex.Message}");
            }
            catch (Exception ex)
            {
                await HandleExceptionAsync(ex, callerInfo);
                //throw; // 重新抛出以中断流程
            }
        }
        public static Task HandleExceptionAsync(Exception ex, string context)
        {
            Log.Error(ex, "指令执行失败: {Context}", context);
            return Task.CompletedTask;
        }
        public static string CallerInfo(string info = "",
            [CallerFilePath]string file = null,
            [CallerLineNumber] int line = 0,
            [CallerMemberName]string method = null)
        {
            var fileName = System.IO.Path.GetFileName(file);
            return $"{fileName} : line {line} {method} {info}";
        }
    }
    public static class ExceptionFormatter
    {
        public static string FormatException(Exception ex)
        {
            var sb = new StringBuilder();
            FormatExceptionRecursive(ex, sb, 0);
            return sb.ToString().TrimEnd();
        }

        private static void FormatExceptionRecursive(Exception ex, StringBuilder sb, int indentLevel)
        {
            if (ex == null) return;

            string indent = new string(' ', indentLevel * 4);

            // 异常消息处理
            sb.AppendLine($"{indent}{ex.GetType().Name}: {ex.Message}");

            // 堆栈跟踪处理
            if (!string.IsNullOrEmpty(ex.StackTrace))
            {
                sb.AppendLine($"{indent}StackTrace:");
                foreach (var line in ProcessStackTrace(ex.StackTrace).Split('\n'))
                {
                    sb.AppendLine($"{indent}    {line.Trim()}");
                }
            }

            // 内部异常处理
            if (ex.InnerException != null)
            {
                sb.AppendLine($"{indent}Inner Exception:");
                FormatExceptionRecursive(ex.InnerException, sb, indentLevel + 1);
            }
        }

        private static string ProcessStackTrace(string stackTrace)
        {
            var result = new StringBuilder();
            foreach (var line in stackTrace.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                result.AppendLine(ProcessStackTraceLine(line));
            }
            return result.ToString().TrimEnd();
        }

        private static string ProcessStackTraceLine(string line)
        {
            var match = Regex.Match(line,
                @"at\s+(?<method>.+?)\s+in\s+(?<file>.+):line (?<line>\d+)");

            if (!match.Success) return line;

            return $"在 {ProcessMethod(match.Groups["method"].Value)} 在 " +
                   $"{ProcessFilePath(match.Groups["file"].Value)}:line {match.Groups["line"].Value}";
        }

        private static string ProcessMethod(string method)
        {
            var methodParts = method.Split(new[] { '(' }, 2);
            var fullMethod = methodParts[0];
            var parameters = methodParts.Length > 1 ? $"({methodParts[1]}" : "";

            var segments = fullMethod.Split('.');
            if (segments.Length < 2) return method;

            return $"{segments[segments.Length - 2]}.{segments[segments.Length - 1]}{parameters}";
        }

        private static string ProcessFilePath(string path)
        {
            try
            {
                return $"{Path.GetFileName(path)}";
            }
            catch
            {
                return path;
            }
        }
    }
    public static class SafeExecutor
    {
        /*
        * 这几个函数特别烧脑,很难理解哦
        * 2025-4-3
        */
        public class Paras
        {
            public string OpName { get; set; }
            public object Context { get; set; }
            public int maxRetry { get; set; } = 1;
            public Type[] IgnorableExes { get; set; }
        }
        //忽略泛型委托返回值,
        public static bool Try<T>(Func<T> operation, Paras para)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    var temp = operation(); // 这里operation的返回值可以赋值给Try外面的变量,但不能作为Try的返回值使用.
                    return true;
                }
                catch (Exception ex) when (IsIgnorable(ex, para.IgnorableExes))
                {
                    LogSuppressedError(ex, para.OpName, para.Context, attempts);

                    if (++attempts >= para.maxRetry)
                        return false;

                    Thread.Sleep(CalculateBackoff(attempts)); // 指数退避
                }
            }
        }
        //委托无返回值
        public static bool Try(Action operation, Paras para)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    operation(); // 这里operation的返回值可以赋值给Try外面的变量,但不能作为Try的返回值使用.
                    return true;
                }
                catch (Exception ex) when (IsIgnorable(ex, para.IgnorableExes))
                {
                    LogSuppressedError(ex, para.OpName, para.Context, attempts);

                    if (++attempts >= para.maxRetry)
                        return false;

                    Thread.Sleep(CalculateBackoff(attempts)); // 指数退避
                }
            }
        }
        //忽略泛型委托返回值,异步模式
        public static async Task<bool> TryAsync<T>(Func<Task<T>> operation, Paras para)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    var temp = await operation();// 这里operation的返回值可以赋值给Try外面的变量,不能作为返回值使用.
                    return true;
                }
                catch (Exception ex) when (IsIgnorable(ex, para.IgnorableExes))
                {
                    LogSuppressedError(ex, para.OpName, para.Context, attempts);

                    if (++attempts >= para.maxRetry)
                        return false;

                    await Task.Delay(CalculateBackoff(attempts)); // 指数退避
                }
            }
        }

        //
        // 推荐这个形式!泛型委托无返回值,异步模式 Func<Task> 代表委托是异步的,但诶呦返回值
        // 2025-4-5
        //
        public static async Task<bool> TryAsync(Func<Task> operation, Paras para)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    await operation();
                    return true; //成功则退出while循环
                }
                catch (Exception ex) when (IsIgnorable(ex, para.IgnorableExes))
                {
                    LogSuppressedError(ex, para.OpName, para.Context, attempts);

                    if (++attempts >= para.maxRetry)
                        return false;//超出上限,退出while循环

                    await Task.Delay(CalculateBackoff(attempts)); // 指数退避,稍后再试
                }
            }
        }
        // 有返回值的同步版本 - 复杂
        public static (T Result, bool Success) Try2<T>(Func<T> operation, Paras para)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    var result = operation();
                    return (result, true);
                }
                catch (Exception ex) when (IsIgnorable(ex, para.IgnorableExes))
                {
                    LogSuppressedError(ex, para.OpName, para.Context, attempts);

                    if (++attempts >= para.maxRetry)
                        return (default, false);

                    Thread.Sleep(CalculateBackoff(attempts)); // 指数退避
                }
            }
        }
        // 有返回值的异步版本 - 复杂
        public static async Task<(T Result, bool Success)> Try2Async<T>(Func<Task<T>> operation, Paras para)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    var result = await operation();
                    return (result, true);
                }
                catch (Exception ex) when (IsIgnorable(ex, para.IgnorableExes))
                {
                    LogSuppressedError(ex, para.OpName, para.Context, attempts);

                    if (++attempts >= para.maxRetry)
                        return (default, false);

                    await Task.Delay(CalculateBackoff(attempts)); // 指数退避
                }
            }
        }

        // 指数退避算法
        private static int CalculateBackoff(int attempt) =>
            Math.Min(1000, (int)Math.Pow(2, attempt) * 100);
        private static bool IsIgnorable(Exception ex, Type[] allowedTypes)
        {
            if (allowedTypes == null) return false; //不提供白名单则不会忽略任何异常
            return allowedTypes.Any(t => t.IsInstanceOfType(ex));//忽略指定的异常
        }

        private static void LogSuppressedError(Exception ex, string operation, object ctx, int attempt)
        {
            var logEntry = new
            {
                Operation = operation,
                Context = ctx,
                Attempt = attempt,
                Error = ExceptionFormatter.FormatException(ex),
                Environment.MachineName,
                Environment.UserName,
                Environment.OSVersion,
                Environment.CurrentDirectory,
                ThreadId = Environment.CurrentManagedThreadId
            };

            //Log.Debug($"SUPPRESSED_ERROR: {JsonConvert.SerializeObject(logEntry)}");
            $"SUPPRESSED_ERROR: {logEntry}".Dump();
        }
    }
}
