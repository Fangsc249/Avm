
using Microsoft.Playwright;
using System.Threading.Tasks;
using Serilog;
using Newtonsoft.Json.Linq;
using System;

using AutomationVM.Core;
using static AutomationVM.Core.InstructionParser;
using System.Collections.Generic;

namespace AutomationVM.Module.PlaywrightIns
{
    public abstract class PlaywrightBaseInstruction : Instruction
    {
        protected static IPlaywright _playwright;
        protected static IBrowserContext _browserCtx;
        protected static IBrowser _browser;
        protected static IPage _page;
        public PlaywrightBaseInstruction()
        {
            //_page.Request += (object o, IRequest requst) => { Console.WriteLine($"Event request "); };
        }
        // 初始化 Playwright 实例
        protected async Task EnsurePlaywrightInitializedAsync()
        {
            if (_playwright == null)
            {
                _playwright = await Playwright.CreateAsync();
                //_browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                //{
                //    Headless = false,
                //    Channel = "msedge",
                //    SlowMo = 650,
                //});

                var userFolder = @"C:\Users\u324926\AppData\Local\Microsoft\Edge\User Data\Default";
                _browserCtx = await _playwright.Chromium.LaunchPersistentContextAsync(userFolder
                    ,
                    new BrowserTypeLaunchPersistentContextOptions
                    {
                        Headless = false,
                        Channel = "msedge",
                        SlowMo = 650,
                    });
                _page = await _browserCtx.NewPageAsync();
            }
        }

    }
    public class OpenBrowserInstruction : PlaywrightBaseInstruction
    {
        public string BrowserType { get; set; } // Chromium/Firefox/WebKit
        public bool Headless { get; set; } = true;

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            Log.Information("已启动 {BrowserType} 浏览器", BrowserType);
        }
    }
    public class NavigateToInstruction : PlaywrightBaseInstruction
    {
        public string Url { get; set; }
        public int TimeoutS { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Url = await EvaluateParameterAsync(Url, context) as string;

            await EnsurePlaywrightInitializedAsync();
            await _page.GotoAsync(Url, new PageGotoOptions { WaitUntil = WaitUntilState.NetworkIdle, Timeout = 1000 * TimeoutS });
            Log.Information("已导航至: {Url}", Url);
        }
    }
    public class TypeTextInstruction : PlaywrightBaseInstruction
    {
        public string Selector { get; set; }
        public string Text { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            await _page.FillAsync(Selector, Text);
            Log.Debug("在元素 {Selector} 输入文本: {Text}", Selector, Text);
        }
    }
    public class ClickElementInstruction : PlaywrightBaseInstruction
    {
        public string Selector { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            await _page.ClickAsync(Selector);
            Log.Debug("点击元素: {Selector}", Selector);
        }
    }
    public class ScreenshotInstruction : PlaywrightBaseInstruction
    {
        public string FilePath { get; set; }

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            await _page.ScreenshotAsync(new PageScreenshotOptions
            {
                Path = FilePath,
                FullPage = true
            });
            Log.Information("截图已保存至: {Path}", FilePath);
        }
    }
    public class CloseBrowserInstruction : PlaywrightBaseInstruction
    {
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            if (_browser != null)
            {
                await _browser.CloseAsync();
                _playwright.Dispose();
                Log.Information("浏览器已关闭");
            }
        }
    }
    public class PageQuerySelectorInstruction : PlaywrightBaseInstruction
    {
        public string Selector { get; set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            Selector = await EvaluateParameterAsync(Selector, context) as string;
            Log.Debug("Selector:{0}", Selector);
            var element = await _page.QuerySelectorAsync(Selector);
            if (element is null)
            {
                Log.Debug("{selector} was not found!", Selector);
                return;
            }
            else
            {
                Log.Debug("{selector} was found!", Selector);
                context.VarsDict["element"] = element;
            }

        }
    }
    public class ElementQuerySelectorInstruction : PlaywrightBaseInstruction
    {
        public string Selector { get; set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            Selector = await EvaluateParameterAsync(Selector, context) as string;
            var parent = context.VarsDict["element"] as IElementHandle;
            var element = await parent.QuerySelectorAsync(Selector);
            if (element is null)
            {
                Log.Debug("{selector} was not found!", Selector);
                return;
            }
            else
            {
                Log.Debug("{selector} was found!", Selector);
                context.VarsDict["element"] = element;
            }

        }
    }
    public class ElementQuerySelectorAllInstruction : PlaywrightBaseInstruction
    {
        public string Selector { get; set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            Selector = await EvaluateParameterAsync(Selector, context) as string;
            var parent = context.VarsDict["element"] as IElementHandle;
            var elements = await parent.QuerySelectorAllAsync(Selector);
            if (elements is null)
            {
                Log.Debug("{selector} was not found!", Selector);
                return;
            }
            else
            {
                Log.Debug("{selector} was found!", Selector);
                context.VarsDict["elements"] = elements;
            }

        }
    }
    public class FrameQuerySelectorInstruction : PlaywrightBaseInstruction
    {
        public string Selector { get; set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            Selector = await EvaluateParameterAsync(Selector, context) as string;
            var parent = context.VarsDict["frame"] as IFrame;
            var element = await parent.QuerySelectorAsync(Selector);
            if (element is null)
            {
                Log.Debug("{selector} was not found!", Selector);
                return;
            }
            else
            {
                Log.Debug("{selector} was found!", Selector);
                context.VarsDict["element"] = element;
            }

        }
    }
    public class ElementContentFrameInstruction : PlaywrightBaseInstruction
    {
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            var parent = context.VarsDict["element"] as IElementHandle;
            var frame = await parent.ContentFrameAsync();
            if (frame is null)
            {
                Log.Debug("Frame was not found!");
                return;
            }
            else
            {
                Log.Debug("Frame was found!");
                context.VarsDict["frame"] = frame;
            }
        }
    }
    public class ElementContentTextInstruction : PlaywrightBaseInstruction
    {
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            await EnsurePlaywrightInitializedAsync();
            var element = context.VarsDict["element"] as IElementHandle;
            var txt = await element.TextContentAsync();
            if (txt is null)
            {
                Log.Debug("text was not found!");
                return;
            }
            else
            {
                Log.Debug("text was found!");
                context.VarsDict["text"] = txt;
            }
        }
    }
    public class SaveTableContentInstruction : PlaywrightBaseInstruction
    {
        public string CellType { get; set; }
        public string Columns { get; set; }
        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            var rows = context.VarsDict["elements"] as IReadOnlyList<IElementHandle>;
            var headDict = context.VarsDict["headDict"] as Dictionary<int, string>;
            int[] intArray = Array.ConvertAll(Columns.Split(','), int.Parse);

            var dataList = new List<Dictionary<string, string>>();
            foreach (var row in rows)
            {
                var dataDict = new Dictionary<string, string>();
                var cells = await row.QuerySelectorAllAsync(CellType);
                for (int i = 0; i < intArray.Length; i++)
                {
                    int colIndex = intArray[i];
                    dataDict.Add(headDict[colIndex], await cells[colIndex].TextContentAsync());
                }
                dataList.Add(dataDict);
            }
            context.VarsDict["dataList"] = dataList;
        }
    }
    //************************************** Instruction Parsers ************************************

    public static class PlaywrightInstructionParsers
    {
        public static OpenBrowserInstruction ParseOpenBrowser(JToken token)
        {
            try
            {
                return new OpenBrowserInstruction
                {
                    BrowserType = GetTokenValue<string>(token, "BrowserType"),
                    Headless = GetTokenValue<bool>(token, "Headless", false),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static NavigateToInstruction ParseNavigateTo(JToken token)
        {
            try
            {
                return new NavigateToInstruction
                {
                    Url = GetTokenValue<string>(token, "Url"),
                    TimeoutS = GetTokenValue<int>(token, "TimeOutS"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ScreenshotInstruction ParseScreenshot(JToken token)
        {
            try
            {
                return new ScreenshotInstruction
                {
                    FilePath = GetTokenValue<string>(token, "FilePath"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static TypeTextInstruction ParseTypeText(JToken token)
        {
            try
            {
                return new TypeTextInstruction
                {
                    Selector = GetTokenValue<string>(token, "Selector"),
                    Text = GetTokenValue<string>(token, "Text"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ClickElementInstruction ParseClickElement(JToken token)
        {
            try
            {
                return new ClickElementInstruction
                {
                    Selector = GetTokenValue<string>(token, "Selector"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static PageQuerySelectorInstruction ParsePageQuerySelector(JToken token)
        {
            try
            {
                return new PageQuerySelectorInstruction
                {
                    Selector = GetTokenValue<string>(token, "Selector"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ElementQuerySelectorInstruction ParseElementQuerySelector(JToken token)
        {
            try
            {
                return new ElementQuerySelectorInstruction
                {
                    Selector = GetTokenValue<string>(token, "Selector"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ElementQuerySelectorAllInstruction ParseElementQuerySelectorAll(JToken token)
        {
            try
            {
                return new ElementQuerySelectorAllInstruction
                {
                    Selector = GetTokenValue<string>(token, "Selector"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static FrameQuerySelectorInstruction ParseFrameQuerySelector(JToken token)
        {
            try
            {
                return new FrameQuerySelectorInstruction
                {
                    Selector = GetTokenValue<string>(token, "Selector"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ElementContentFrameInstruction ParseElementContentFrame(JToken token)
        {
            try
            {
                return new ElementContentFrameInstruction
                {

                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static ElementContentTextInstruction ParseElementContentText(JToken token)
        {
            try
            {
                return new ElementContentTextInstruction
                {

                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
        public static SaveTableContentInstruction ParseSaveTableContent(JToken token)
        {
            try
            {
                return new SaveTableContentInstruction
                {
                    CellType = GetTokenValue<string>(token, "CellType"),
                    Columns = GetTokenValue<string>(token, "Columns"),
                };
            }
            catch (Exception ex) when (ex is ArgumentNullException aex)
            {
                Log.Error("{message}", aex.Message);
                throw new JTokenParseException("ArgumentNullException when Get Token Values", CallerInfo(), aex);
            }
            catch (Exception ex) when (ex is InvalidCastException iex)
            {
                Log.Error("{message}", iex.Message);
                throw new JTokenParseException("InvalidCastException when Get Token Values", CallerInfo(), iex);
            }
        }
    }
}