//using System;
//using YamlDotNet.Serialization;
//using Newtonsoft.Json;
//using YamlDotNet.Serialization.NamingConventions;

//public static class YamlJsonConverter
//{
//    // JSON → YAML
//    public static string JsonToYaml(string json)
//    {
//        var obj = JsonConvert.DeserializeObject(json);
//        return new SerializerBuilder()
//            .DisableAliases() // 禁用锚点
//            .WithIndentedSequences() // 美化数组格式
//            .Build()
//            .Serialize(obj);
//    }

//    // YAML → JSON
//    public static string YamlToJson(string yaml)
//    {
//        var deserializer = new DeserializerBuilder()
//            .WithNamingConvention(CamelCaseNamingConvention.Instance)
//            .Build();

//        var yamlObj = deserializer.Deserialize(yaml);
//        return JsonConvert.SerializeObject(yamlObj, new JsonSerializerSettings
//        {
//            NullValueHandling = NullValueHandling.Ignore,
//            Formatting = Formatting.Indented
//        });
//    }
//}