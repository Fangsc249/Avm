using Newtonsoft.Json.Linq;
using System;

namespace AutomationVM.Core
{
    /*
     * 捕获下级异常，记录日志
     * 抛出本级异常，附加本级消息
     * 捕获时记录日志，抛出不记录。
     * 抛出时提供callerinfo
     * 2025-3-29
    */
    public class AVMException : Exception
    {
        public string CallerInfo { get; set; }
        public AVMException(string message, string caller, Exception inner = null) : base(message, inner)
        {
            CallerInfo = caller;
        }
    }
    public class ParseInstructionException : AVMException
    {
        public string InstructionName { get; set; }
        public ParseInstructionException(string msg, string caller, Exception inner = null) : base(msg, caller, inner) { }
    }

    public class JTokenParseException : AVMException
    {
        public JToken FaultyToken { get; set; }
        public string FaultyKey { get; set; }
        public JTokenParseException(string message, string caller, Exception inner = null) : base(message, caller, inner) { }
    }

    public class NullReferenceException : AVMException
    {
        public string ObjectName { get; set; }
        public NullReferenceException(string message, string caller, Exception inner = null) : base(message, caller, inner) { }
    }

    public class SapConnectionFailureException : AVMException
    {
        public SapConnectionFailureException(string message, string caller, Exception inner = null)
            : base(message, caller, inner)
        { }
    }
    public class ExecuteSubflowException : AVMException
    {
        public ExecuteSubflowException(string message, string caller, Exception inner = null)
            : base(message, caller, inner)
        { }
    }
}