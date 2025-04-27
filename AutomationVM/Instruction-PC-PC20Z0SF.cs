using Microsoft.CodeAnalysis.Scripting;
using Serilog;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AutomationVM.Core
{
    public abstract class Instruction
    {
        public abstract Task ExecuteCoreAsync(ExecutionContext context); // 子类实现这个。
        public async Task ExecuteAsync(ExecutionContext context)//子类对象中调用这个。
        {
            await ExecuteCoreAsync(context);
        }
        // 统一参数求值方法
        //protected async Task<object> EvaluateParameterAsync(string rawValue, ExecutionContext context)
        //{
        //    if (rawValue == null) return null;

        //    // 1. 变量引用（$开头）
        //    if (rawValue.Trim().StartsWith("$"))
        //    {
        //        string varName = rawValue.Substring(1);
        //        if (context.VarsDict.TryGetValue(varName, out object value))
        //            return value;
        //        else
        //            throw new KeyNotFoundException($"变量未定义: {varName}");
        //    }

        //    // 2. 表达式（=开头）
        //    if (rawValue.Trim().StartsWith("="))
        //    {
        //        string expr = rawValue.Substring(1);
        //        try
        //        {
        //            return await context.EvaluateAsync(expr);
        //        }
        //        catch (CompilationErrorException ex)
        //        {
        //            Log.Error("表达式编译错误！{Expression} {Errors}", expr, string.Join("\n", ex.Diagnostics));
        //            throw;
        //        }
        //    }

        //    // 3. 字面量（直接返回）
        //    return rawValue;
        //}
        protected async Task<object> EvaluateParameterAsync(string rawValue, ExecutionContext context)
        {
            /*
            表达式中的动态变量访问：

            原逻辑：仅支持整个字符串为$var的变量引用。
            新逻辑：使用正则表达式\$([a - zA - Z_]\w *)匹配所有$开头的变量名（如$city）。
            表达式处理：将匹配到的变量转换为Vars.varName（例如$city变为Vars.city），允许在表达式中动态访问变量。
            普通字符串的变量替换：

            对于非表达式字符串（不以 = 开头），直接替换所有$var为变量的当前值。
            若变量不存在则抛出KeyNotFoundException。
            正则表达式增强：

            支持变量名包含字母、数字和下划线（如$user_name）。
            避免匹配到非变量字符，确保精准替换。
            2025-3-30
            */
            if (rawValue == null) return null;

            // 处理表达式（以=开头）
            if (rawValue.Trim().StartsWith("="))
            {
                string expr = rawValue.Trim().Substring(1);
                // 替换表达式中的$变量为Vars.变量名
                expr = Regex.Replace(expr, @"\$([a-zA-Z_]\w*)", "Vars.$1");//
                //expr = Regex.Replace(expr, @"\$([a-zA-Z_]\w*)", "VarsDict[\"$1\"]");
                try
                {
                    return await context.EvaluateAsync(expr);
                }
                catch (CompilationErrorException ex)
                {
                    Log.Error("表达式编译错误！{Expression} {Errors}", expr, string.Join("\n", ex.Diagnostics));
                    throw;
                }
            }

            // 处理普通字符串中的变量引用
            string processedValue = Regex.Replace(rawValue, @"\$([a-zA-Z_]\w*)", match =>
            {
                string varName = match.Groups[1].Value;
                if (context.VarsDict.TryGetValue(varName, out object value))
                {
                    return value?.ToString() ?? string.Empty;
                }
                throw new KeyNotFoundException($"变量未定义: {varName}");
            });

            return processedValue;
        }
    }
}
