用c#设计的一个虚拟机，它通过加载特定格式(Json,Yaml)的指令序列实现一些任务的自动化。
它的第一个用途将会是通过SAP GuiScripting 实现sap操作自动化但不会局限于此，
比如类似的,通过playwright实现网页app操作自动化，实现windows 文件系统批量操作自动化，等等。

Usage:
用yaml格式写一个脚本 test_get_tabledata.yaml:
```yaml
############# SAP Test ##################################################

- Type: Code
  FileName: ../csscripts/SAPID.cs

- Type: Code
  FileName: ../csscripts/TypeDefines.cs

- Type: Code
  Statements: |
    =
    Vars.bomQList = new List<BOMQuery>()
    {
      new BOMQuery {Material="DPC6",Plant="1200",Usage="1"},
      new BOMQuery {Material="40-100C",Plant="1200",Usage="1"},
      new BOMQuery {Material="200-200",Plant="1100",Usage="1"},
      new BOMQuery {Material="AS-101",Plant="1000",Usage="3"},
      //new BOMQuery {Material="135X2327",Plant="G501",Usage="1"},
    };

- Type: SapConnectTo
  System: =SAPID.ERP

- Type: If
  Condition: =Vars.NeedLogon
  Then:
    - Type: SapLogon
      UserId: =SAPID.ID_USERID # sap登录界面上输入UserId的那个输入框，它的控件id，据此搜索控件
      User: LRP
      PasswordId: =SAPID.ID_PASSWORD
      Password: 123456
      LangCodeId: =SAPID.ID_LANGCODE
      LangCode: EN

- Type: ForEach
  Collection: = (Vars.bomQList as List<BOMQuery>)
  ItemVar: bomQ
  Instructions:
    - Type: SapStartTransaction
      TCode: CS03

    - Type: SapSetText
      Id: =SAPID.ID_CS03_MATERIAL
      Value: =$bomQ.Material

    - Type: SapSetText
      Id: =SAPID.ID_CS03_PLANT
      Value: =$bomQ.Plant

    - Type: SapSetText
      Id: =SAPID.ID_CS03_BOM_USAGE #=Vars.id_03
      Value: =$bomQ.Usage

    - Type: SapEnter

    - Type: Code
      Statements: |
        =
        VarsDict.Dump();

    - Type: If
      Condition: = $Title.Contains("Alternative") # 如果存在多个Alternative BOM,此处跳出一个界面，让用户双击选择
      Then:
        - Type: Code
          Statements: |
            =
            Log.Information("There are more then one Alt BOM exists, we select the first.");
        
        - Type: SapDoubleClickTableLine
          Id: =SAPID.ID_CS03_ALT_BOM_TABLE
          Row: 0
          Column: 0

    - Type: If
      Condition: = $SapStatusType !="E" && $SapStatusType != "W"
      Then:
        - Type: SapGetTableData
          Id: =SAPID.ID_CS03_BOM_TABLE
          TargetVariable: tableData
          Columns: "0,1,2,3,4,5,8"
          Filters:
            - ColumnIndex: 3
              FilterRule: Contains
              Value: ""
        - Type: Code
          Statements: |
            =
            //VarsDict["tableData"].Dump();
      Else:
        - Type: Code
          Statements: |
            =
            Console.WriteLine(Vars.SapStatus);

```
运行效果
![2025-04-27_13-40](https://github.com/user-attachments/assets/1ed47926-6431-4193-95d1-0b705f802158)

