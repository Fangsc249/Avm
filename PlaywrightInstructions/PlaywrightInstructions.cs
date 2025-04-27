using AutomationVM.Core;
using Microsoft.Playwright;
using System.Threading.Tasks;
using Serilog;
using Newtonsoft.Json.Linq;
using System;
using static AutomationVM.Core.InstructionParser;

namespace AutomationVM.Module.PlaywrightIns
{
    public abstract class PlaywrightBaseInstruction : Instruction
    {
        protected static IPlaywright _playwright;
        protected static IBrowserContext _browserCtx;
        protected static IBrowser _browser;
        protected static IPage _page;

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

        public override async Task ExecuteCoreAsync(ExecutionContext context)
        {
            Url = await EvaluateParameterAsync(Url, context) as string;

            await EnsurePlaywrightInitializedAsync();
            await _page.GotoAsync(Url);
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