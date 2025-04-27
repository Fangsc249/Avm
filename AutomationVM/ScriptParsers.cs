
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using System.Linq;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using AutomationVM.Core;
using static AutomationVM.Core.InstructionParser;
using Dumpify;

public interface IScriptParser
{
    List<Instruction> Parse(Stream stream);
}
public class JsonScriptParser : IScriptParser
{ /* 原有JSON解析逻辑重构至此 */
    public List<Instruction> Parse(Stream stream)
    {
        using (var reader = new StreamReader(stream))
        {
            var jArray = JArray.Parse(reader.ReadToEnd());
            return ParseInstructions(jArray);
        }
    }
}
public class YamlScriptParser : IScriptParser
{
    public List<Instruction> Parse(Stream stream)
    {
        using (var reader = new StreamReader(stream))
        {
            var deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();

            // 将YAML转换为中间字典结构
            var yamlObject = deserializer.Deserialize<object>(reader);
            var normalized = NormalizeYamlObject(yamlObject);

            // 转换为JToken兼容结构
            var jToken = JToken.FromObject(normalized);
            return ParseInstructions(jToken as JArray);
        }
    }

    // 处理YAML特有的数据类型（如嵌套块）
    private object NormalizeYamlObject(object yamlObj)
    {
        if (yamlObj is IDictionary<object, object> dict)
        {
            var normalized = new Dictionary<string, object>();
            foreach (var kv in dict)
            {
                var key = kv.Key.ToString();
                normalized[key] = NormalizeYamlObject(kv.Value);
            }
            return normalized;
        }
        if (yamlObj is IList list)
        {
            var normalized = new List<object>();
            foreach (var item in list)
            {
                normalized.Add(NormalizeYamlObject(item));
            }
            return normalized;
        }
        return yamlObj;
    }
}


public class XmlScriptParser : IScriptParser //2025-3-30,未完成，不能使用
{
    public List<Instruction> Parse(Stream stream)
    {
        var doc = XDocument.Load(stream);
        var root = doc.Root;

        // 将XML转换为与JSON/YAML兼容的JArray结构
        var normalized = NormalizeXmlElement(root);
        normalized.Dump();
        return new List<Instruction>();
        //return ParseInstructions(JArray.FromObject(normalized));
    }

    private object NormalizeXmlElement(XElement element)
    {
        //if(element.Name.LocalName == "Script")
        //{
        //    return NormalizeXmlElement(element.Elements().First());
        //}
        // 处理指令节点（如<While>、<SapConnectTo>）
        if (element.Elements().Any())
        {
            var dict = new Dictionary<string, object>();
            foreach (var child in element.Elements())
            {
                // 处理重复节点（如多个<Action>）
                if (dict.ContainsKey(child.Name.LocalName))
                {
                    var existing = dict[child.Name.LocalName];
                    if (existing is List<object> list)
                        list.Add(NormalizeXmlElement(child));
                    else
                        dict[child.Name.LocalName] = new List<object> { existing, NormalizeXmlElement(child) };
                }
                else
                {
                    dict[child.Name.LocalName] = NormalizeXmlElement(child);
                }
            }
            return dict;
        }
        return element.Value; // 叶子节点直接返回值
    }
}

public class ScriptParserFactory
{
    public IScriptParser CreateParser(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLower();

        // 优先处理原始文件存在的情况
        if (File.Exists(filePath))
        {
            switch (ext)
            {
                case ".json": return new JsonScriptParser();
                case ".yaml": return new YamlScriptParser();
                case ".yml": return new YamlScriptParser();
                default:
                    throw new NotSupportedException("不支持的脚本格式");
            }
        }

        // 处理文件不存在的情况
        List<string> candidateExtensions = new List<string>();

        if (string.IsNullOrEmpty(ext))
        {
            // 没有扩展名时尝试 .yaml 和 .yml
            candidateExtensions.Add(".yaml");
            candidateExtensions.Add(".yml");
        }
        else if (ext == ".yaml")
        {
            // 原始为 .yaml 时尝试 .yml
            candidateExtensions.Add(".yml");
        }
        else if (ext == ".yml")
        {
            // 原始为 .yml 时尝试 .yaml
            candidateExtensions.Add(".yaml");
        }
        else
        {
            // 其他扩展名直接抛出异常
            throw new NotSupportedException("不支持的脚本格式");
        }

        // 遍历候选扩展名查找文件
        foreach (var candidate in candidateExtensions)
        {
            var newPath = Path.ChangeExtension(filePath, candidate);
            if (File.Exists(newPath))
            {
                if(candidate==".json")
                {
                    return new JsonScriptParser();
                }
                else
                {
                    return new YamlScriptParser();
                }
            }
        }

        // 所有候选扩展名都不存在时抛出异常
        throw new FileNotFoundException(
            $"找不到支持的脚本文件: {Path.GetFileNameWithoutExtension(filePath)}[.yaml/.yml]"
        );
    }
}