using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;

namespace ExcelHighlighterRibbon
{
    // COM add-in that delivers the Highlighter ribbon on Excel builds where
    // VBA-only ribbon delivery (customUI package part / IRibbonExtensibility
    // on ThisWorkbook) is never queried - empirically established on Excel
    // 2024 (build 16.0.20228): the .xlam loads and runs fine, but Excel
    // silently never asks the VBA project for ribbon XML, so the tab never
    // appears. Excel DOES query this object for IRibbonExtensibility
    // directly - a guaranteed mechanism (proven: the ribbon renders).
    //
    // Every ribbon callback is implemented here as a method on this class and
    // delegates to the VBA add-in via Application.Run. Value-returning
    // callbacks (getLabel/getPressed/getEnabled/...) call the *Value wrapper
    // functions in RibbonCallbacks.bas (Application.Run passes arguments by
    // value, so the Sub + ByRef pattern cannot return values through it).
    //
    // Build + register with scripts/build-comribbon.ps1 (called from
    // install.ps1). Unregister with uninstall.ps1.
    [ComVisible(true)]
    [Guid("6E2F4A11-83C5-4B9D-9A07-2D51C8E4F0B6")]
    [ProgId("ExcelHighlighter.Ribbon")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _application;
        private string _addinFolder;

        // The ribbon this add-in delivered. STATIC on purpose: Excel hands
        // the live IRibbonUI to OnLoad of the CONNECTED instance, but the
        // connected instance is unreachable from VBA (Application.COMAddIns(
        // "...").Object returns Nothing for this add-in on Excel 2024, and
        // Application.Run cannot marshal an IRibbonUI argument into VBA -
        // both observed empirically). VBA instead creates a throwaway
        // instance (CreateObject "ExcelHighlighter.Ribbon") and calls
        // InvalidateRibbon on it; the static makes that throwaway instance
        // invalidate the REAL ribbon the connected instance was handed.
        private static Microsoft.Office.Core.IRibbonUI _ribbon;

        // Swatch cache: getItemImage/getImage are polled by Excel far more
        // often than the underlying colour changes (every repaint, hover,
        // invalidate). Generating a fresh bitmap each time is wasteful, so we
        // build each colour's swatch once and reuse it. Keyed by RGB value.
        private readonly Dictionary<int, object> _swatchCache = new Dictionary<int, object>();

        private void Log(string file, string text)
        {
            try { File.AppendAllText(Path.Combine(Path.GetTempPath(), file), text); }
            catch { }
        }

        private object Run(string vbaMacro, params object[] args)
        {
            if (_application == null) return null;
            try
            {
                object[] callArgs = new object[args.Length + 1];
                callArgs[0] = vbaMacro;
                Array.Copy(args, 0, callArgs, 1, args.Length);
                return _application.GetType().InvokeMember("Run",
                    System.Reflection.BindingFlags.InvokeMethod, null, _application, callArgs);
            }
            catch (Exception ex)
            {
                Log("eh-com-callback-error.txt", vbaMacro + ": " + ex.Message + "\r\n");
                return null;
            }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _application = Application;
            _addinFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Microsoft", "AddIns", "ExcelHighlighter");
            Log("eh-com-connection.txt", "OnConnection ConnectMode=" + ConnectMode + " at " + DateTime.Now + "\r\n");
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _application = null;
            Log("eh-com-connection.txt", "OnDisconnection RemoveMode=" + RemoveMode + " at " + DateTime.Now + "\r\n");
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            Log("eh-com-getui.txt", "GetCustomUI called RibbonID=" + RibbonID + " at " + DateTime.Now + "\r\n");
            string xmlPath = Path.Combine(_addinFolder, "customUI14.xml");
            if (File.Exists(xmlPath))
            {
                try
                {
                    string xml = File.ReadAllText(xmlPath);
                    // This object owns the callbacks now: the onLoad name must
                    // resolve to a method on THIS object (OnLoad below), never
                    // to the VBA module. Use a regex rather than an exact
                    // string match so a future whitespace/casing edit to the
                    // XML cannot silently re-point onLoad at a name Excel
                    // cannot resolve on the COM object (which would reproduce
                    // the "Custom UI Runtime Error" dialog).
                    xml = System.Text.RegularExpressions.Regex.Replace(
                        xml,
                        @"onLoad\s*=\s*\""[^\""]*\""",
                        "onLoad=\"OnLoad\"",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    Log("eh-com-getui.txt", "  read " + xml.Length + " chars from " + xmlPath + "\r\n");
                    return xml;
                }
                catch { }
            }
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">" +
                "<ribbon><tabs><tab id=\"tabHighlighter\" label=\"Highlighter\">" +
                "<group id=\"grpFallback\" label=\"Highlighter\">" +
                "<button id=\"btnAboutFallback\" label=\"About\" imageMso=\"Info\" onAction=\"OnAbout_Action\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }

        // ================= onLoad ==========================================
        public void OnLoad(IRibbonUI ribbon)
        {
            // Keep the ribbon in the static field. NOTE: we deliberately do
            // NOT forward this to VBA's RibbonCallbacks.onLoad - Application
            // .Run cannot marshal an IRibbonUI argument on Excel 2024 (the
            // call always throws, polluting eh-com-callback-error.txt on
            // every launch, and VBA's mRibbon never gets set anyway). The
            // static _ribbon below is the authoritative invalidation route,
            // and VBA-side one-time init (Profiles.Init) runs from
            // AddinHost.StartUp instead.
            _ribbon = ribbon;
            Log("eh-com-invalidate.txt", "OnLoad: ribbon captured at " + DateTime.Now + "\r\n");
        }

        // ================= InvalidateRibbon ================================
        // Called by VBA (RibbonCallbacks.InvalidateRibbon) after any state
        // change that happened outside the ribbon itself - the context-menu
        // toggle and the Ctrl+Shift+H hotkey both route here. Invalidate()
        // makes Excel re-query every get* callback, so the toggle button's
        // pressed state and label refresh to match the workbook the user
        // actually toggled. Works on a throwaway CreateObject'd instance
        // because _ribbon is static and shared across instances.
        public void InvalidateRibbon()
        {
            // _ribbon is static, so this works even when called on a
            // throwaway instance created from VBA - it invalidates the
            // ribbon of the connected instance in this process.
            // Log the outcome either way so a stale toggle can be
            // distinguished from a working invalidate: if the static is
            // null (ribbon never handed to OnLoad, or COM add-in disabled)
            // the VBA-side "CreateObject returned Nothing" branch never
            // fires, so this is the only trace that the invalidation
            // silently no-oped.
            try
            {
                if (_ribbon != null)
                {
                    _ribbon.Invalidate();
                    Log("eh-com-invalidate.txt", "InvalidateRibbon OK at " + DateTime.Now + "\r\n");
                }
                else
                {
                    Log("eh-com-invalidate.txt", "InvalidateRibbon: static _ribbon is NULL at " + DateTime.Now + "\r\n");
                }
            }
            catch (Exception ex)
            {
                Log("eh-com-invalidate.txt", "InvalidateRibbon threw: " + ex.Message + " at " + DateTime.Now + "\r\n");
            }
        }

        // ================= Controls group ===================================
        public bool GetToggle_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetToggle_PressedValue", control)); }
        public string GetToggle_Label(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetToggle_LabelValue", control)); }
        public void OnToggle_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnToggle_Action", control, pressed); }

        public bool GetMode_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetMode_PressedValue", control)); }
        public bool GetMode_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetMode_EnabledValue", control)); }
        public void OnMode_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnMode_Action", control, pressed); }

        // ================= Colour gallery ===================================
        public bool GetGallery_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetGallery_EnabledValue", control)); }
        public string GetGallery_SelectedItemID(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetGallery_SelectedItemIDValue", control)); }
        public int GetGallery_ItemCount(IRibbonControl control)
        { return Convert.ToInt32(Run("RibbonCallbacks.GetGallery_ItemCountValue", control)); }
        public string GetGallery_ItemID(IRibbonControl control, int index)
        { return Convert.ToString(Run("RibbonCallbacks.GetGallery_ItemIDValue", control, index)); }
        public string GetGallery_ItemLabel(IRibbonControl control, int index)
        { return Convert.ToString(Run("RibbonCallbacks.GetGallery_ItemLabelValue", control, index)); }
        public string GetGallery_ItemScreentip(IRibbonControl control, int index)
        { return Convert.ToString(Run("RibbonCallbacks.GetGallery_ItemScreentipValue", control, index)); }
        public object GetGallery_ItemImage(IRibbonControl control, int index)
        {
            // Swatch generated in C# (safe on 64-bit Excel - no VBA GDI /
            // OleCreatePictureIndirect), colour from VBA so presets and
            // recent custom colours both render.
            object rgb = Run("RibbonCallbacks.GetGallery_ItemRGBValue", control, index);
            return MakeSwatch(rgb == null ? 65535 : Convert.ToInt32(rgb));
        }
        public void OnGallery_Action(IRibbonControl control, string selectedId, int index)
        { Run("RibbonCallbacks.OnGallery_Action", control, selectedId, index); }

        public void OnCustomColour_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnCustomColour_Action", control); }

        // Custom-colour button icon: the currently selected highlight colour.
        public object GetSwatchImage(IRibbonControl control)
        {
            object rgb = Run("RibbonCallbacks.GetSwatchRGBValue", control);
            return MakeSwatch(rgb == null ? 65535 : Convert.ToInt32(rgb));
        }

        // ================= Scope (per-workbook vs all) ======================
        public bool GetScopeAll_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetScopeAll_PressedValue", control)); }
        public string GetScopeAll_Label(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetScopeAll_LabelValue", control)); }
        public void OnScopeAll_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnScopeAll_Action", control, pressed); }

        public bool GetPerModeColours_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetPerModeColours_PressedValue", control)); }
        public bool GetPerModeColours_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetPerModeColours_EnabledValue", control)); }
        public void OnPerModeColours_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnPerModeColours_Action", control, pressed); }

        public bool GetRowColourGallery_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetRowColourGallery_EnabledValue", control)); }
        public string GetRowColour_SelectedItemID(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetRowColour_SelectedItemIDValue", control)); }
        public void OnRowColour_Action(IRibbonControl control, string selectedId, int index)
        { Run("RibbonCallbacks.OnRowColour_Action", control, selectedId, index); }

        public string GetColColour_SelectedItemID(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetColColour_SelectedItemIDValue", control)); }
        public void OnColColour_Action(IRibbonControl control, string selectedId, int index)
        { Run("RibbonCallbacks.OnColColour_Action", control, selectedId, index); }

        // ================= Options group ====================================
        public bool GetExclude_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetExclude_PressedValue", control)); }
        public bool GetExclude_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetExclude_EnabledValue", control)); }
        public string GetExclude_Label(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetExclude_LabelValue", control)); }
        public void OnExclude_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnExclude_Action", control, pressed); }

        public bool GetExcludeSheet_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetExcludeSheet_PressedValue", control)); }
        public bool GetExcludeSheet_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetExcludeSheet_EnabledValue", control)); }
        public string GetExcludeSheet_Label(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetExcludeSheet_LabelValue", control)); }
        public void OnExcludeSheet_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnExcludeSheet_Action", control, pressed); }

        public bool GetAllowProtected_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetAllowProtected_PressedValue", control)); }
        public bool GetAllowProtected_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetAllowProtected_EnabledValue", control)); }
        public void OnAllowProtected_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnAllowProtected_Action", control, pressed); }

        public bool GetDarkMode_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetDarkMode_PressedValue", control)); }
        public bool GetDarkMode_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetDarkMode_EnabledValue", control)); }
        public void OnDarkMode_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnDarkMode_Action", control, pressed); }

        public void OnReset_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnReset_Action", control); }

        // ================= Style group ======================================
        public bool GetStyle_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetStyle_PressedValue", control)); }
        public bool GetStyle_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetStyle_EnabledValue", control)); }
        public void OnStyle_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnStyle_Action", control, pressed); }

        public bool GetIntersection_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetIntersection_PressedValue", control)); }
        public bool GetIntersection_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetIntersection_EnabledValue", control)); }
        public void OnIntersection_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnIntersection_Action", control, pressed); }

        public bool GetAnimated_Pressed(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetAnimated_PressedValue", control)); }
        public bool GetAnimated_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetAnimated_EnabledValue", control)); }
        public void OnAnimated_Action(IRibbonControl control, bool pressed)
        { Run("RibbonCallbacks.OnAnimated_Action", control, pressed); }

        // ================= History group ====================================
        public bool GetHistoryBack_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetHistoryBack_EnabledValue", control)); }
        public void OnHistoryBack_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnHistoryBack_Action", control); }

        public bool GetHistoryForward_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetHistoryForward_EnabledValue", control)); }
        public void OnHistoryForward_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnHistoryForward_Action", control); }

        // ================= Profiles group ===================================
        public bool GetProfile_Enabled(IRibbonControl control)
        { return Convert.ToBoolean(Run("RibbonCallbacks.GetProfile_EnabledValue", control)); }
        public string GetProfilesContent(IRibbonControl control)
        { return Convert.ToString(Run("RibbonCallbacks.GetProfilesContentValue", control)); }
        public void OnProfileMenu_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnProfileMenu_Action", control); }
        public void OnSaveProfile_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnSaveProfile_Action", control); }

        // ================= Help group =======================================
        public void OnAbout_Action(IRibbonControl control)
        { Run("RibbonCallbacks.OnAbout_Action", control); }

        // ================= Swatch generation =================================

        // AxHost.GetIPictureDispFromPicture is protected - this tiny subclass
        // (the standard trick) exposes it so we can hand Excel an IPictureDisp
        // built from a System.Drawing.Bitmap.
        private sealed class PictureHelper : System.Windows.Forms.AxHost
        {
            private PictureHelper() : base("") { }
            public static object ToPictureDisp(System.Drawing.Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }

        private object MakeSwatch(int rgb)
        {
            try
            {
                object cached;
                if (_swatchCache.TryGetValue(rgb, out cached)) return cached;

                // rgb arrives as VBA's BGR Long (0x00BBGGRR).
                int r = rgb & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = (rgb >> 16) & 0xFF;

                using (var bmp = new System.Drawing.Bitmap(32, 32))
                using (var gfx = System.Drawing.Graphics.FromImage(bmp))
                {
                    gfx.Clear(System.Drawing.Color.FromArgb(r, g, b));
                    // 1px inner border so pale swatches stay visible on white.
                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(110, 0, 0, 0)))
                    {
                        gfx.DrawRectangle(pen, 0, 0, 31, 31);
                    }
                    object pic = PictureHelper.ToPictureDisp(bmp);
                    _swatchCache[rgb] = pic;
                    return pic;
                }
            }
            catch (Exception ex)
            {
                Log("eh-com-callback-error.txt", "MakeSwatch(" + rgb + "): " + ex.Message + "\r\n");
                return null;
            }
        }
    }
}
