using Microsoft.VisualBasic;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.IO;
using SapROTWr;

public interface ISAPActive
{
    GuiSession GetSession(string sid);
}
public class SAPActive : ISAPActive
{
    public static GuiSession session = null;
    public static GuiApplication GetRunningSapApp()
    {
        //Interaction类来自如下这个dll
        //C:\Program Files\dotnet\shared\Microsoft.NETCore.App\8.0.8\Microsoft.VisualBasic.Core.dll
        object sap = Interaction.GetObject("SAPGUI", "");
        return sap.GetType().InvokeMember("GetScriptingEngine",
            System.Reflection.BindingFlags.InvokeMethod, null, sap, null) as GuiApplication;
    }
    public static GuiApplication GetSapApp()
    {
        try { return GetRunningSapApp(); }
        catch (Exception e)
        {
            try
            {
                Console.WriteLine(@"failed by: object sap = Interaction.GetObject(""SAPGUI"", "")");
                Console.WriteLine(e.Message);
                Console.WriteLine("Starting SAP Gui Logon Pad...");
                string SapPath = @"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe";
                if (File.Exists(SapPath))
                {
                    System.Diagnostics.Process.Start(SapPath);
                    DateTime StartTime = DateTime.Now;
                    object SapGuilRot = new CSapROTWrapper().GetROTEntry("SAPGUI");
                    while (SapGuilRot == null && 30 >= (DateTime.Now - StartTime).TotalSeconds)//等待启动过程
                    {
                        SapGuilRot = new CSapROTWrapper().GetROTEntry("SAPGUI");
                    }
                    return SapGuilRot.GetType().InvokeMember("GetScriptingEngine",
                        System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null) as GuiApplication;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
    public static GuiConnection GetConnection(string sid)
    {
        GuiApplication guiApp = GetSapApp();
        foreach (GuiConnection conn in guiApp.Connections)
        {
            if (conn.Description.Trim().Contains(sid.Trim()))//可能SID指定的系统已经打开了.
            {
                return conn;
            }
        }
        return guiApp.OpenConnection(sid, true);//打开并返回对应的系统.
    }
    // <summary>2024-1-13</summary>
    public static List<GuiSession> GetSessionList()
    {
        List<GuiSession> sessionList = new List<SAPFEWSELib.GuiSession>();
        GuiApplication app = GetSapApp();
        foreach (GuiConnection conn in app.Connections)
        {
            foreach (GuiSession s in conn.Children)
            {
                sessionList.Add(s);
            }
        }
        return sessionList;
    }
    public static List<MyGuiSession> GetMySessionList()//2024-1-13
    {
        List<MyGuiSession> sessionList = new List<MyGuiSession>();
        GuiApplication app = GetSapApp();
        foreach (GuiConnection conn in app.Connections)
        {
            foreach (GuiSession s in conn.Children)
            {
                sessionList.Add(new MyGuiSession(s));
            }
        }
        return sessionList;
    }
    // <summary>2023-7-19</summary>
    public static GuiSession GetSession(string sid, int sessionIndex)
    {
        var conn = GetConnection(sid);
        var cnt = conn.Sessions.Count;
        if (sessionIndex < cnt)
        {
            session = (GuiSession)conn.Sessions.Item(sessionIndex);
            return session;
        }
        return null;
    }
    public GuiSession GetSession(string sid)
    {
        var conn = GetConnection(sid);
        return conn.Sessions.Item(0) as GuiSession;
    }
    public static GuiSession GetSession(string sid, string tcode)
    { //
        var conn = GetConnection(sid);
        foreach (GuiSession s in conn.Sessions)
        {
            if (s.ActiveWindow.Text == "System Messages") //2022-1-21 notification message dialog
            {
                ((GuiButton)s.ActiveWindow.FindById("tbar[0]/btn[12]")).Press();
                System.Threading.Thread.Sleep(2000);
            }
        }

        foreach (GuiSession s in conn.Sessions)
        {
            if (s.Info.Transaction == tcode)
            {
                //s.SendCommand("/N" + tcode);//回到初始界面
                session = s;
                return s; // There is one session with the target TCode
            }
        }
        foreach (GuiSession s in conn.Sessions)
        {
            if (s.Info.Transaction == "SESSION_MANAGER")
            {
                //s.SendCommand("/N" + tcode);
                session = s;
                return s;// There is one session is SESSION_MANAGER
            }
        }
        if (conn.Sessions.Count < 6)
        {
            (conn.Sessions.Item(conn.Sessions.Count - 1) as GuiSession).SendCommand("/o" + tcode);
            session = conn.Sessions.Item(conn.Sessions.Count - 1) as GuiSession;
            return session;
        }
        else
        {
            GuiSession s = conn.Sessions.Item(5) as GuiSession; // return the last session
            s.SendCommand("/N" + tcode);
            session = s;
            return s;
        }

    }
}

