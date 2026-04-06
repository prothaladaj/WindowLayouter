// WindowLayouter.Native — native Win32 window layout manager.
//
// Architecture overview:
//   - Single-file Win32 application; all logic lives in an anonymous namespace.
//   - Three top-level windows: main window (g_mainWindow), preset editor
//     (g_presetEditorWindow), and about dialog (g_aboutWindow).
//   - Persistent settings are stored in settings.ini next to the executable.
//   - UI language is loaded at runtime from lang/<code>.ini (UTF-8 encoded).
//   - The app registers as DPI-aware (per-monitor v2); all pixel sizes must go
//     through MulDiv(value, dpi, 96) before use.
//   - A system-tray icon provides quick access without a taskbar button.
//   - Global hotkeys are registered via RegisterHotKey for each preset.

#include <windows.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shellscalingapi.h>
#include <dwmapi.h>
#include <string>
#include <vector>
#include <sstream>
#include <algorithm>
#include <cwctype>
#include <map>
#include <set>
#include <fstream>
#include <filesystem>
#include <uxtheme.h>
#include "resource.h"
#include "build-version.h"

namespace
{
    // -------------------------------------------------------------------------
    // Control IDs — main window
    // -------------------------------------------------------------------------
    constexpr int IDC_REFRESH = 1001;
    constexpr int IDC_CLEAR_SELECTION = 1002;
    constexpr int IDC_PRESET_GRID_2X2 = 1004;
    constexpr int IDC_PRESET_SPLIT_2X2 = 1005;
    constexpr int IDC_PRESET_SPLIT_2X3 = 1006;
    constexpr int IDC_PRESET_SPLIT_3X2 = 1007;
    constexpr int IDC_PRESET_SPLIT_3X3 = 1008;
    constexpr int IDC_WINDOW_LIST = 1009;
    constexpr int IDC_STATUS = 1010;
    constexpr int IDC_TAB = 1011;
    constexpr int IDC_TAB_MONITORS = 1012;
    constexpr int IDC_TAB_PRESETS = 1013;
    constexpr int IDC_TAB_RULES = 1014;
    constexpr int IDC_FILTER_TITLE = 1015;
    constexpr int IDC_FILTER_PROCESS = 1016;
    constexpr int IDC_FILTER_TITLE_LABEL = 1017;
    constexpr int IDC_FILTER_PROCESS_LABEL = 1018;
    constexpr int IDC_PRESET_EDITOR_LIST = 1019;
    constexpr int IDC_PRESET_NAME = 1020;
    constexpr int IDC_PRESET_HOTKEY = 1021;
    constexpr int IDC_PRESET_KIND = 1022;
    constexpr int IDC_PRESET_COLUMNS = 1023;
    constexpr int IDC_PRESET_ROWS = 1024;
    constexpr int IDC_PRESET_WIDTH = 1025;
    constexpr int IDC_PRESET_LEFT_TITLE = 1026;
    constexpr int IDC_PRESET_LEFT_PROCESS = 1027;
    constexpr int IDC_PRESET_RIGHT_TITLE = 1028;
    constexpr int IDC_PRESET_RIGHT_PROCESS = 1029;
    constexpr int IDC_PRESET_SAVE = 1030;
    constexpr int IDC_PRESET_APPLY_FORM = 1031;
    constexpr int IDC_PRESET_STATUS = 1032;
    constexpr int IDC_PRESET_NAME_LABEL = 1033;
    constexpr int IDC_PRESET_HOTKEY_LABEL = 1034;
    constexpr int IDC_PRESET_KIND_LABEL = 1035;
    constexpr int IDC_PRESET_COLUMNS_LABEL = 1036;
    constexpr int IDC_PRESET_ROWS_LABEL = 1037;
    constexpr int IDC_PRESET_WIDTH_LABEL = 1038;
    constexpr int IDC_PRESET_LEFT_TITLE_LABEL = 1039;
    constexpr int IDC_PRESET_LEFT_PROCESS_LABEL = 1040;
    constexpr int IDC_PRESET_RIGHT_TITLE_LABEL = 1041;
    constexpr int IDC_PRESET_RIGHT_PROCESS_LABEL = 1042;
    constexpr int IDC_APPLY_SELECTED_PRESET = 1043;
    constexpr int IDC_PRESET_TRAY_LABEL = 1044;
    constexpr int IDC_PRESET_BUTTON_LABEL = 1045;
    constexpr int IDC_PRESET_SHOW_TRAY = 1046;
    constexpr int IDC_PRESET_SHOW_BUTTON = 1047;
    constexpr int IDC_PRESET_TRAY_LABEL_TEXT = 1048;
    constexpr int IDC_PRESET_BUTTON_LABEL_TEXT = 1049;
    constexpr int IDC_PRESET_USE_SELECTION = 1050;
    constexpr int IDC_APP_FORCE_SELECTION = 1051;
    constexpr int IDC_PRESET_REQUIRE_EXTERNAL = 1052;
    constexpr int IDC_PRESET_RESOLUTION_CLASS = 1053;
    constexpr int IDC_PRESET_RESOLUTION_LABEL = 1054;
    constexpr int IDC_PRESET_SCOPE = 1055;
    constexpr int IDC_PRESET_SCOPE_LABEL = 1056;
    constexpr int IDC_APP_SCOPE_OVERRIDE = 1057;
    constexpr int IDC_APP_SCOPE_OVERRIDE_LABEL = 1058;
    constexpr int IDC_OPEN_PRESET_EDITOR = 1059;
    constexpr int IDC_PRESET_DEVICE = 1060;
    constexpr int IDC_PRESET_DEVICE_LABEL = 1061;
    constexpr int IDC_PRESET_SUMMARY = 1062;
    constexpr int IDC_PRESET_PREVIEW = 1063;
    constexpr int IDC_LANGUAGE_LABEL = 1064;
    constexpr int IDC_LANGUAGE = 1065;
    constexpr int IDC_PRESETS_DASHBOARD_TITLE = 1066;
    constexpr int IDC_PRESET_EDITOR_TITLE = 1067;
    constexpr int IDC_OPEN_ABOUT = 1068;
    constexpr int IDC_ABOUT_TITLE = 1069;
    constexpr int IDC_ABOUT_TEXT = 1070;
    constexpr int IDC_ABOUT_CLOSE = 1071;
    constexpr UINT_PTR TIMER_LANG_RELOAD = 1;

    // -------------------------------------------------------------------------
    // Control IDs — tray context menu
    // -------------------------------------------------------------------------
    constexpr int ID_TRAY_SHOW = 3001;
    constexpr int ID_TRAY_HIDE = 3002;
    constexpr int ID_TRAY_EXIT = 3003;
    constexpr int ID_TRAY_PRESET_EDITOR = 3004;
    constexpr int ID_TRAY_ABOUT = 3005;

    // Custom WM_APP message sent by the tray icon; LOWORD(lParam) = mouse event.
    constexpr UINT WMAPP_TRAYICON = WM_APP + 1;
    constexpr UINT TRAY_ICON_UID = 1;
    // Hotkey IDs start here; each preset uses HOTKEY_BASE_ID + preset index.
    constexpr int HOTKEY_BASE_ID = 4001;

    // -------------------------------------------------------------------------
    // Data structures
    // -------------------------------------------------------------------------

    // Integer rectangle — avoids RECT/POINT dependencies in model code.
    struct RectInt
    {
        int left{};
        int top{};
        int width{};
        int height{};
    };

    // Physical monitor metadata collected via EnumDisplayMonitors.
    struct MonitorInfo
    {
        std::wstring deviceName;  // e.g. \\.\DISPLAY1
        bool isPrimary{};
        RectInt bounds{};         // full monitor area in virtual-screen coords
        RectInt workArea{};       // area excluding taskbar
        UINT dpi{96};
    };

    // A single top-level window discovered by EnumWindows.
    struct WindowInfo
    {
        HWND handle{};
        int zOrder{};             // enumeration order, used for stable sort
        std::wstring title;
        std::wstring processName; // basename, e.g. "chrome.exe"
        std::wstring className;
        std::wstring monitorName; // deviceName of the monitor it's on
        RectInt bounds{};
        UINT dpi{96};
        bool minimized{};
    };

    // Preset layout strategy.
    // Grid  — tiles all selected windows into an NxM grid.
    // Split — routes "left" windows (matched by pattern) to a grid on the left
    //         portion and "right" windows (e.g. Chrome) to a stack on the right.
    enum class PresetKind
    {
        Grid,
        Split
    };

    // Broad monitor resolution categories used for per-preset activation guards.
    enum class ResolutionClass
    {
        Any,
        R2K,   // ~2560×1440
        R25K,  // ~2560×1600
        R4K    // ~3840×2160
    };

    // Which windows a preset acts on when applied.
    enum class SelectionScope
    {
        SelectedRows,         // only rows selected in the list view
        AllVisibleRows,       // all rows currently shown (respects filters)
        AllDiscoveredWindows  // every visible top-level window
    };

    // Global override that overrides every preset's SelectionScope.
    enum class AppSelectionOverride
    {
        FollowPreset,         // each preset uses its own scope
        SelectedRows,
        AllVisibleRows,
        AllDiscoveredWindows
    };

    // Parsed representation of a hotkey string such as "Ctrl+Alt+1".
    struct HotkeySpec
    {
        UINT modifiers{};    // MOD_CONTROL | MOD_ALT | …
        UINT virtualKey{};   // VK_* code
        std::wstring text;   // original user-readable string
        bool valid{};
    };

    // Complete description of a layout preset.  Instances live in g_presets.
    struct PresetDefinition
    {
        int commandId{};      // WM_COMMAND ID used for toolbar buttons and tray
        int hotkeyId{};       // ID passed to RegisterHotKey
        std::wstring iniKey;  // section name in settings.ini

        std::wstring name;
        std::wstring trayLabel;
        std::wstring toolbarLabel;
        bool showInTray{true};
        bool showAsToolbarButton{true};

        SelectionScope selectionScope{SelectionScope::SelectedRows};
        PresetKind kind{PresetKind::Grid};

        int columns{};
        int rows{};
        // Split mode: percentage of monitor width given to the right stack.
        // Clamped to [20, 80] so neither side becomes unusably narrow.
        int chromeWidthPercent{};

        // Split mode patterns: wildcards matched against window title / process name.
        std::wstring leftTitlePattern;
        std::wstring leftProcessPattern;
        std::wstring rightTitlePattern;
        std::wstring rightProcessPattern;

        // Activation guards — preset is blocked if the condition is not met.
        bool requireExternalMonitor{};
        ResolutionClass requiredResolutionClass{ResolutionClass::Any};
        std::wstring requiredMonitorDevice; // empty = any device

        HotkeySpec hotkey;
    };

    // section → (key → value) map for both settings.ini and language files.
    using IniData = std::map<std::wstring, std::map<std::wstring, std::wstring>>;

    // -------------------------------------------------------------------------
    // Global window handles — main window and its child controls
    // -------------------------------------------------------------------------
    HWND g_mainWindow{};
    HWND g_listView{};      // list of discovered top-level windows
    HWND g_status{};        // status bar at the bottom
    HWND g_tab{};           // tab control (Monitors / Presets / Rules)
    HWND g_monitorsPage{};
    HWND g_presetsPage{};
    HWND g_rulesPage{};
    HWND g_monitorsEdit{};  // read-only text dump of detected monitors
    HWND g_rulesEdit{};     // read-only help text for the Rules tab
    HWND g_presetsSummaryEdit{};
    HWND g_filterTitleEdit{};
    HWND g_filterProcessEdit{};
    HWND g_appScopeOverrideCombo{};
    HWND g_languageCombo{};
    HWND g_mainTooltips{}; // TTM_ADDTOOL tooltip window for the main window

    // -------------------------------------------------------------------------
    // Preset editor window and its child controls
    // -------------------------------------------------------------------------
    HWND g_presetEditorWindow{};
    HWND g_presetList{};        // listbox — one entry per preset
    HWND g_presetStatus{};      // description / validation feedback
    HWND g_presetNameEdit{};
    HWND g_presetHotkeyEdit{};
    HWND g_presetKindCombo{};   // Grid / Split
    HWND g_presetColumnsEdit{};
    HWND g_presetRowsEdit{};
    HWND g_presetWidthEdit{};   // right stack width %, enabled only for Split
    HWND g_presetTrayLabelEdit{};
    HWND g_presetButtonLabelEdit{};
    HWND g_presetLeftTitleEdit{};
    HWND g_presetLeftProcessEdit{};
    HWND g_presetRightTitleEdit{};
    HWND g_presetRightProcessEdit{};
    HWND g_presetScopeCombo{};
    HWND g_presetShowTrayCheck{};
    HWND g_presetShowButtonCheck{};
    HWND g_presetRequireExternalCheck{};
    HWND g_presetResolutionCombo{};
    HWND g_presetDeviceCombo{};
    HWND g_presetPreview{};     // custom-drawn tile layout preview
    HWND g_presetEditorTooltips{};

    // -------------------------------------------------------------------------
    // About dialog
    // -------------------------------------------------------------------------
    HWND g_aboutWindow{};
    HWND g_aboutText{};

    // -------------------------------------------------------------------------
    // GDI font handles — main window uses g_uiFont*; preset editor and about
    // use separate g_auxUiFont* so DPI changes on different monitors work
    // independently.
    // -------------------------------------------------------------------------
    HFONT g_uiFont{};
    HFONT g_uiFontBold{};
    HFONT g_uiFontTitle{};
    HFONT g_auxUiFont{};
    HFONT g_auxUiFontBold{};
    HFONT g_auxUiFontTitle{};

    // -------------------------------------------------------------------------
    // Application state
    // -------------------------------------------------------------------------
    NOTIFYICONDATAW g_trayIcon{};
    UINT g_taskbarCreatedMessage{}; // WM registered by the shell; re-adds icon on taskbar restart
    bool g_allowExit{};             // set to true before DestroyWindow to allow real exit
    bool g_trayAdded{};
    std::wstring g_iniPath;
    std::wstring g_languageCode;
    IniData g_languageData;
    long long g_languageFilesStamp{};  // mtime sum used to detect live language file changes
    int g_windowGapPx{};              // gap between tiled windows, read from settings.ini
    AppSelectionOverride g_appSelectionOverride{AppSelectionOverride::FollowPreset};
    int g_sortColumn{};
    bool g_sortAscending{true};
    int g_selectedPresetIndex{};
    UINT g_currentDpi{96}; // DPI of the monitor the main window is on; updated on WM_DPICHANGED

    std::vector<WindowInfo> g_windows;       // all enumerated windows (full set)
    std::vector<WindowInfo*> g_visibleWindows; // subset matching current filters
    std::vector<MonitorInfo> g_monitors;
    std::vector<PresetDefinition> g_presets;
    std::map<HWND, std::wstring> g_tooltipTexts;
    std::set<HWND> g_registeredTooltipControls;

    std::wstring FormatHotkeyText(const HotkeySpec& hotkey);
    std::wstring BuildPresetDescription(const PresetDefinition& preset);
    std::wstring GetControlText(HWND control);
    void RegisterPresetHotkeys();
    void SaveSettings();
    void UpdateTabContent();
    bool MatchesOptionalPattern(const std::wstring& pattern, const std::wstring& value);
    ResolutionClass DetectResolutionClass(const MonitorInfo& monitor);
    std::wstring FormatResolutionClass(ResolutionClass value);
    ResolutionClass ParseResolutionClass(const std::wstring& value);
    const MonitorInfo* GetPrimaryMonitor();
    std::wstring FormatSelectionScope(SelectionScope value);
    SelectionScope ParseSelectionScope(const std::wstring& value);
    std::wstring FormatAppSelectionOverride(AppSelectionOverride value);
    AppSelectionOverride ParseAppSelectionOverride(const std::wstring& value);
    void RefreshMonitorDeviceOptions(HWND combo, const std::wstring& selectedValue);
    void ShowPresetEditorWindow();
    void ResizePresetEditorControls();
    std::wstring BuildPresetSummaryText();
    LRESULT CALLBACK PresetEditorProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
    LRESULT CALLBACK PresetPreviewProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
    RectInt InsetRect(RectInt rect, int padding);
    std::vector<RectInt> BuildGrid(RectInt frame, int columns, int rows, int count);
    IniData ReadIniDataFromPath(const std::wstring& path);
    std::wstring GetExecutableDirectory();
    std::wstring DetectDefaultLanguageCode();
    void LoadLanguageResources(const std::wstring& languageCode);
    std::wstring Lang(const std::wstring& key, const std::wstring& fallback);
    std::wstring Tip(const std::wstring& key, const std::wstring& fallback);
    void ApplyMainWindowLocalization();
    void ApplyPresetEditorLocalization();
    void UpdateStatus(const std::wstring& text);
    HWND CreateTooltipWindow(HWND owner);
    void SetToolTipText(HWND tooltipWindow, HWND owner, HWND control, const std::wstring& text);
    void ApplyAboutLocalization();
    void ShowAboutWindow();
    void CreateAboutWindow(HINSTANCE instance);
    std::wstring BuildAboutText();
    LRESULT CALLBACK AboutProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
    void ResizeAboutControls();
    long long GetLanguageFilesStamp();
    void ReloadLanguageIfChanged();
    UINT GetWindowDpiSafe(HWND window);
    PresetDefinition BuildPreviewPresetFromEditor();
    void RefreshPresetEditorPreviewState();
    void RefreshWidthFieldState();

    HICON LoadAppIcon(int width, int height)
    {
        return static_cast<HICON>(LoadImageW(
            GetModuleHandleW(nullptr),
            MAKEINTRESOURCEW(IDI_APP_ICON),
            IMAGE_ICON,
            width,
            height,
            LR_DEFAULTCOLOR));
    }

    BOOL CALLBACK ApplyFontToChildren(HWND window, LPARAM)
    {
        SendMessageW(window, WM_SETFONT, reinterpret_cast<WPARAM>(g_uiFont), TRUE);
        SetWindowTheme(window, L"Explorer", nullptr);
        return TRUE;
    }

    void CreateUiFonts()
    {
        if (g_uiFont != nullptr) { DeleteObject(g_uiFont); g_uiFont = nullptr; }
        if (g_uiFontBold != nullptr) { DeleteObject(g_uiFontBold); g_uiFontBold = nullptr; }
        if (g_uiFontTitle != nullptr) { DeleteObject(g_uiFontTitle); g_uiFontTitle = nullptr; }

        const int bodySize = -MulDiv(13, static_cast<int>(g_currentDpi), 96);
        const int titleSize = -MulDiv(16, static_cast<int>(g_currentDpi), 96);

        g_uiFont = CreateFontW(bodySize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_uiFontBold = CreateFontW(bodySize, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_uiFontTitle = CreateFontW(titleSize, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    }

    void CreateAuxUiFonts(UINT dpi)
    {
        if (g_auxUiFont != nullptr) { DeleteObject(g_auxUiFont); g_auxUiFont = nullptr; }
        if (g_auxUiFontBold != nullptr) { DeleteObject(g_auxUiFontBold); g_auxUiFontBold = nullptr; }
        if (g_auxUiFontTitle != nullptr) { DeleteObject(g_auxUiFontTitle); g_auxUiFontTitle = nullptr; }

        const int bodySize = -MulDiv(13, static_cast<int>(dpi), 96);
        const int titleSize = -MulDiv(16, static_cast<int>(dpi), 96);

        g_auxUiFont = CreateFontW(bodySize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_auxUiFontBold = CreateFontW(bodySize, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_auxUiFontTitle = CreateFontW(titleSize, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    }

    void ApplyUiFonts(HWND window)
    {
        if (g_uiFont == nullptr)
        {
            CreateUiFonts();
        }

        SendMessageW(window, WM_SETFONT, reinterpret_cast<WPARAM>(g_uiFont), TRUE);
        EnumChildWindows(window, ApplyFontToChildren, 0);

        const int boldIds[] = {
            IDC_APPLY_SELECTED_PRESET,
            IDC_PRESET_GRID_2X2,
            IDC_PRESET_SPLIT_2X2,
            IDC_PRESET_SPLIT_2X3,
            IDC_PRESET_SPLIT_3X2,
            IDC_PRESET_SPLIT_3X3,
            IDC_PRESET_SAVE,
            IDC_PRESET_APPLY_FORM,
            IDC_OPEN_PRESET_EDITOR
        };

        for (int controlId : boldIds)
        {
            if (auto* control = GetDlgItem(window, controlId); control != nullptr)
            {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(g_uiFontBold), TRUE);
            }
            if (auto* control = GetDlgItem(g_presetsPage, controlId); control != nullptr)
            {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(g_uiFontBold), TRUE);
            }
        }

        const int titleIds[] = {
            IDC_PRESETS_DASHBOARD_TITLE,
            IDC_PRESET_EDITOR_TITLE,
            IDC_ABOUT_TITLE
        };

        for (int controlId : titleIds)
        {
            if (auto* control = GetDlgItem(window, controlId); control != nullptr)
            {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(g_uiFontTitle), TRUE);
            }
        }
    }

    void ApplyAuxUiFonts(HWND window)
    {
        if (window == nullptr)
        {
            return;
        }

        const UINT dpi = GetWindowDpiSafe(window);
        if (g_auxUiFont == nullptr)
        {
            CreateAuxUiFonts(dpi);
        }

        SendMessageW(window, WM_SETFONT, reinterpret_cast<WPARAM>(g_auxUiFont), TRUE);
        EnumChildWindows(window, [](HWND child, LPARAM) -> BOOL
        {
            SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(g_auxUiFont), TRUE);
            SetWindowTheme(child, L"Explorer", nullptr);
            return TRUE;
        }, 0);

        const int boldIds[] = {
            IDC_PRESET_SAVE,
            IDC_PRESET_APPLY_FORM,
            IDC_ABOUT_CLOSE
        };

        for (int controlId : boldIds)
        {
            if (auto* control = GetDlgItem(window, controlId); control != nullptr)
            {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(g_auxUiFontBold), TRUE);
            }
        }

        const int titleIds[] = {
            IDC_PRESET_EDITOR_TITLE,
            IDC_ABOUT_TITLE
        };

        for (int controlId : titleIds)
        {
            if (auto* control = GetDlgItem(window, controlId); control != nullptr)
            {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(g_auxUiFontTitle), TRUE);
            }
        }
    }

    std::wstring Trim(const std::wstring& value)
    {
        const auto begin = std::find_if_not(value.begin(), value.end(), [](wchar_t character) { return std::iswspace(character) != 0; });
        if (begin == value.end())
        {
            return {};
        }

        const auto end = std::find_if_not(value.rbegin(), value.rend(), [](wchar_t character) { return std::iswspace(character) != 0; }).base();
        return std::wstring(begin, end);
    }

    std::vector<std::wstring> Split(const std::wstring& value, wchar_t separator)
    {
        std::vector<std::wstring> parts;
        size_t start = 0;
        while (start <= value.size())
        {
            const auto next = value.find(separator, start);
            if (next == std::wstring::npos)
            {
                parts.push_back(value.substr(start));
                break;
            }

            parts.push_back(value.substr(start, next - start));
            start = next + 1;
        }

        return parts;
    }

    std::wstring ToUpper(std::wstring value)
    {
        std::transform(value.begin(), value.end(), value.begin(), [](wchar_t character)
        {
            return static_cast<wchar_t>(std::towupper(character));
        });
        return value;
    }

    std::wstring FormatResolutionClass(ResolutionClass value)
    {
        switch (value)
        {
        case ResolutionClass::R2K:
            return L"2K";
        case ResolutionClass::R25K:
            return L"2.5K";
        case ResolutionClass::R4K:
            return L"4K";
        case ResolutionClass::Any:
        default:
            return Lang(L"resolution.any", L"Any");
        }
    }

    ResolutionClass ParseResolutionClass(const std::wstring& value)
    {
        const auto upper = ToUpper(Trim(value));
        if (upper == L"2K")
        {
            return ResolutionClass::R2K;
        }
        if (upper == L"2.5K" || upper == L"2,5K" || upper == L"25K")
        {
            return ResolutionClass::R25K;
        }
        if (upper == L"4K")
        {
            return ResolutionClass::R4K;
        }
        return ResolutionClass::Any;
    }

    ResolutionClass DetectResolutionClass(const MonitorInfo& monitor)
    {
        const int width = std::max(monitor.bounds.width, monitor.bounds.height);
        const int height = std::min(monitor.bounds.width, monitor.bounds.height);
        if (width >= 3800 && height >= 2100)
        {
            return ResolutionClass::R4K;
        }
        if (width >= 2500 && height >= 1550)
        {
            return ResolutionClass::R25K;
        }
        if (width >= 2500 && height >= 1400)
        {
            return ResolutionClass::R2K;
        }
        return ResolutionClass::Any;
    }

    std::wstring FormatSelectionScope(SelectionScope value)
    {
        switch (value)
        {
        case SelectionScope::AllVisibleRows:
            return Lang(L"scope.all_visible_rows", L"All visible rows");
        case SelectionScope::AllDiscoveredWindows:
            return Lang(L"scope.all_discovered_windows", L"All discovered windows");
        case SelectionScope::SelectedRows:
        default:
            return Lang(L"scope.selected_rows", L"Selected rows");
        }
    }

    SelectionScope ParseSelectionScope(const std::wstring& value)
    {
        const auto upper = ToUpper(Trim(value));
        if (upper == L"ALL VISIBLE ROWS")
        {
            return SelectionScope::AllVisibleRows;
        }
        if (upper == L"ALL DISCOVERED WINDOWS")
        {
            return SelectionScope::AllDiscoveredWindows;
        }
        return SelectionScope::SelectedRows;
    }

    std::wstring FormatAppSelectionOverride(AppSelectionOverride value)
    {
        switch (value)
        {
        case AppSelectionOverride::SelectedRows:
            return Lang(L"scope.selected_rows", L"Selected rows");
        case AppSelectionOverride::AllVisibleRows:
            return Lang(L"scope.all_visible_rows", L"All visible rows");
        case AppSelectionOverride::AllDiscoveredWindows:
            return Lang(L"scope.all_discovered_windows", L"All discovered windows");
        case AppSelectionOverride::FollowPreset:
        default:
            return Lang(L"scope.follow_preset", L"Follow preset");
        }
    }

    AppSelectionOverride ParseAppSelectionOverride(const std::wstring& value)
    {
        const auto upper = ToUpper(Trim(value));
        if (upper == L"SELECTED ROWS")
        {
            return AppSelectionOverride::SelectedRows;
        }
        if (upper == L"ALL VISIBLE ROWS")
        {
            return AppSelectionOverride::AllVisibleRows;
        }
        if (upper == L"ALL DISCOVERED WINDOWS")
        {
            return AppSelectionOverride::AllDiscoveredWindows;
        }
        return AppSelectionOverride::FollowPreset;
    }

    std::wstring ReadWindowText(HWND window)
    {
        const int length = GetWindowTextLengthW(window);
        if (length <= 0)
        {
            return {};
        }

        std::wstring text(static_cast<size_t>(length) + 1, L'\0');
        GetWindowTextW(window, text.data(), length + 1);
        text.resize(static_cast<size_t>(length));
        return text;
    }

    std::wstring ReadClassName(HWND window)
    {
        wchar_t buffer[256]{};
        GetClassNameW(window, buffer, static_cast<int>(std::size(buffer)));
        return buffer;
    }

    bool IsWindowCloaked(HWND window)
    {
        DWORD cloaked = 0;
        return DwmGetWindowAttribute(window, DWMWA_CLOAKED, &cloaked, sizeof(cloaked)) == S_OK && cloaked != 0;
    }

    // Recursive wildcard match (? = any char, * = any substring).
    // Called only through WildcardMatch() which limits patterns to max 2 asterisks
    // to prevent exponential backtracking on malicious or accidental input.
    bool WildcardMatchImpl(const wchar_t* pattern, const wchar_t* value)
    {
        while (*pattern != L'\0')
        {
            if (*pattern == L'*')
            {
                ++pattern;
                if (*pattern == L'\0')
                {
                    return true;
                }

                while (*value != L'\0')
                {
                    if (WildcardMatchImpl(pattern, value))
                    {
                        return true;
                    }
                    ++value;
                }

                return false;
            }

            if (*value == L'\0' || std::towlower(*pattern) != std::towlower(*value))
            {
                return false;
            }

            ++pattern;
            ++value;
        }

        return *value == L'\0';
    }

    bool WildcardMatch(const std::wstring& pattern, const std::wstring& value)
    {
        int starCount = 0;
        for (auto ch : pattern)
        {
            if (ch == L'*' && ++starCount > 2)
            {
                return false;
            }
        }
        return WildcardMatchImpl(pattern.c_str(), value.c_str());
    }

    std::wstring GetAppDataDirectory()
    {
        wchar_t path[MAX_PATH]{};
        SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, path);
        return std::wstring(path) + L"\\WindowLayouter.Native";
    }

    std::wstring GetExecutableDirectory()
    {
        wchar_t path[MAX_PATH]{};
        GetModuleFileNameW(nullptr, path, static_cast<DWORD>(std::size(path)));
        std::wstring fullPath(path);
        const auto slash = fullPath.find_last_of(L"\\/");
        return slash == std::wstring::npos ? L"." : fullPath.substr(0, slash);
    }

    UINT GetWindowDpiSafe(HWND window)
    {
        if (window == nullptr)
        {
            return g_currentDpi == 0 ? 96 : g_currentDpi;
        }

        const UINT dpi = GetDpiForWindow(window);
        return dpi == 0 ? 96 : dpi;
    }

    void EnsureConfigDirectory()
    {
        CreateDirectoryW(GetAppDataDirectory().c_str(), nullptr);
    }

    std::wstring StripInlineComment(const std::wstring& value)
    {
        const auto commentIndex = value.find(L"//");
        if (commentIndex == std::wstring::npos)
        {
            return Trim(value);
        }

        return Trim(value.substr(0, commentIndex));
    }

    // =========================================================================
    // INI / settings persistence
    // =========================================================================

    IniData ReadIniDataFromPath(const std::wstring& path)
    {
        IniData data;

        // Read raw bytes so we can handle UTF-8 encoding correctly.
        // std::wifstream with the default "C" locale cannot decode UTF-8.
        std::ifstream rawInput{std::filesystem::path(path), std::ios::binary};
        if (!rawInput.is_open())
        {
            return data;
        }
        std::string utf8{std::istreambuf_iterator<char>(rawInput), std::istreambuf_iterator<char>()};

        // Strip UTF-8 BOM (EF BB BF) if present.
        if (utf8.size() >= 3 &&
            static_cast<unsigned char>(utf8[0]) == 0xEF &&
            static_cast<unsigned char>(utf8[1]) == 0xBB &&
            static_cast<unsigned char>(utf8[2]) == 0xBF)
        {
            utf8.erase(0, 3);
        }

        // Convert UTF-8 → UTF-16 using the Windows API.
        int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), nullptr, 0);
        std::wstring wContent(static_cast<size_t>(wlen), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), wContent.data(), wlen);

        // Parse the wide-character content line by line.
        std::wstring currentSection;
        std::wistringstream stream{wContent};
        std::wstring line;
        while (std::getline(stream, line))
        {
            // Strip UTF-16 LE BOM if somehow present at the very first character.
            if (!line.empty() && line.front() == 0xFEFF)
            {
                line.erase(line.begin());
            }
            // Strip Windows-style carriage return left by getline.
            if (!line.empty() && line.back() == L'\r')
            {
                line.pop_back();
            }

            auto trimmed = Trim(line);
            if (trimmed.empty() || trimmed.rfind(L"//", 0) == 0)
            {
                continue;
            }

            if (trimmed.front() == L'[' && trimmed.back() == L']')
            {
                currentSection = Trim(trimmed.substr(1, trimmed.size() - 2));
                continue;
            }

            const auto equalsIndex = trimmed.find(L'=');
            if (equalsIndex == std::wstring::npos || currentSection.empty())
            {
                continue;
            }

            const auto key = Trim(trimmed.substr(0, equalsIndex));
            const auto value = StripInlineComment(trimmed.substr(equalsIndex + 1));
            if (!key.empty())
            {
                data[currentSection][key] = value;
            }
        }

        return data;
    }

    IniData ReadIniData()
    {
        return ReadIniDataFromPath(g_iniPath);
    }

    std::wstring ReadIniString(const IniData& data, const std::wstring& section, const std::wstring& key, const std::wstring& fallback)
    {
        const auto sectionIt = data.find(section);
        if (sectionIt == data.end())
        {
            return fallback;
        }

        const auto keyIt = sectionIt->second.find(key);
        return keyIt == sectionIt->second.end() ? fallback : keyIt->second;
    }

    int ReadIniInt(const IniData& data, const std::wstring& section, const std::wstring& key, int fallback)
    {
        const auto text = ReadIniString(data, section, key, std::to_wstring(fallback));
        try
        {
            return std::stoi(text);
        }
        catch (...)
        {
            return fallback;
        }
    }

    std::wstring Lang(const std::wstring& key, const std::wstring& fallback)
    {
        return ReadIniString(g_languageData, L"Strings", key, fallback);
    }

    std::wstring Tip(const std::wstring& key, const std::wstring& fallback)
    {
        return ReadIniString(g_languageData, L"Tooltips", key, fallback);
    }

    std::wstring DetectDefaultLanguageCode()
    {
        const auto languageId = GetUserDefaultUILanguage();
        const auto primary = PRIMARYLANGID(languageId);
        return primary == LANG_POLISH ? L"pl" : L"en";
    }

    void LoadLanguageResources(const std::wstring& languageCode)
    {
        auto sanitized = Trim(languageCode);
        if (sanitized.empty())
        {
            sanitized = DetectDefaultLanguageCode();
        }
        for (auto ch : sanitized)
        {
            if (!((ch >= L'a' && ch <= L'z') || (ch >= L'A' && ch <= L'Z') || (ch >= L'0' && ch <= L'9') || ch == L'-' || ch == L'_'))
            {
                sanitized = L"en";
                break;
            }
        }
        g_languageCode = sanitized;
        auto languagePath = GetExecutableDirectory() + L"\\lang\\" + g_languageCode + L".ini";
        g_languageData = ReadIniDataFromPath(languagePath);

        if (g_languageData.empty() && _wcsicmp(g_languageCode.c_str(), L"en") != 0)
        {
            g_languageCode = L"en";
            languagePath = GetExecutableDirectory() + L"\\lang\\en.ini";
            g_languageData = ReadIniDataFromPath(languagePath);
        }
        g_languageFilesStamp = GetLanguageFilesStamp();
    }

    long long GetLanguageFilesStamp()
    {
        long long stamp = 0;
        const auto langDir = std::filesystem::path(GetExecutableDirectory()) / "lang";
        if (!std::filesystem::exists(langDir))
        {
            return stamp;
        }

        for (const auto& entry : std::filesystem::directory_iterator(langDir))
        {
            if (!entry.is_regular_file() || entry.path().extension() != ".ini")
            {
                continue;
            }

            const auto timeCount = entry.last_write_time().time_since_epoch().count();
            stamp ^= static_cast<long long>(timeCount);
        }

        return stamp;
    }

    HWND CreateTooltipWindow(HWND owner)
    {
        return CreateWindowExW(
            WS_EX_TOPMOST,
            TOOLTIPS_CLASSW,
            nullptr,
            WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            owner,
            nullptr,
            GetModuleHandleW(nullptr),
            nullptr);
    }

    void SetToolTipText(HWND tooltipWindow, HWND owner, HWND control, const std::wstring& text)
    {
        if (tooltipWindow == nullptr || control == nullptr)
        {
            return;
        }

        g_tooltipTexts[control] = text;

        TOOLINFOW toolInfo{};
        toolInfo.cbSize = sizeof(toolInfo);
        toolInfo.uFlags = TTF_SUBCLASS | TTF_IDISHWND;
        toolInfo.hwnd = owner;
        toolInfo.uId = reinterpret_cast<UINT_PTR>(control);
        toolInfo.lpszText = const_cast<LPWSTR>(g_tooltipTexts[control].c_str());

        if (g_registeredTooltipControls.insert(control).second)
        {
            SendMessageW(tooltipWindow, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&toolInfo));
        }
        else
        {
            SendMessageW(tooltipWindow, TTM_UPDATETIPTEXTW, 0, reinterpret_cast<LPARAM>(&toolInfo));
        }
    }

    void SetControlTextById(HWND parent, int controlId, const std::wstring& text)
    {
        if (auto* control = GetDlgItem(parent, controlId); control != nullptr)
        {
            SetWindowTextW(control, text.c_str());
        }
    }

    void SetListViewColumnText(int columnIndex, const std::wstring& text)
    {
        if (g_listView == nullptr)
        {
            return;
        }

        LVCOLUMNW column{};
        column.mask = LVCF_TEXT;
        column.pszText = const_cast<LPWSTR>(text.c_str());
        ListView_SetColumn(g_listView, columnIndex, &column);
    }

    // =========================================================================
    // Localization — apply translated strings from g_languageData to all controls.
    // Called once at startup and again whenever the user switches language.
    // =========================================================================

    void ApplyMainWindowLocalization()
    {
        if (g_mainWindow == nullptr)
        {
            return;
        }

        SetWindowTextW(g_mainWindow, Lang(L"app.title", L"WindowLayouter.Native").c_str());
        SetControlTextById(g_mainWindow, IDC_REFRESH, Lang(L"main.refresh", L"Refresh"));
        SetControlTextById(g_mainWindow, IDC_CLEAR_SELECTION, Lang(L"main.clear_selection", L"Clear Selection"));
        SetControlTextById(g_mainWindow, IDC_APPLY_SELECTED_PRESET, Lang(L"main.apply_selected_preset", L"Apply Selected Preset"));
        SetControlTextById(g_mainWindow, IDC_OPEN_PRESET_EDITOR, Lang(L"main.open_preset_editor", L"Preset Editor"));
        SetControlTextById(g_mainWindow, IDC_FILTER_TITLE_LABEL, Lang(L"main.filter_title", L"Title"));
        SetControlTextById(g_mainWindow, IDC_FILTER_PROCESS_LABEL, Lang(L"main.filter_process", L"Process"));
        SetControlTextById(g_rulesPage, IDC_APP_SCOPE_OVERRIDE_LABEL, Lang(L"rules.scope_override", L"App Selection Override"));
        SetControlTextById(g_presetsPage, IDC_PRESETS_DASHBOARD_TITLE, Lang(L"presets.dashboard_title", L"Preset Dashboard"));
        SetControlTextById(g_rulesPage, IDC_OPEN_ABOUT, Lang(L"about.open_button", L"About"));

        if (g_tab != nullptr)
        {
            TCITEMW item{};
            item.mask = TCIF_TEXT;
            auto tab0 = Lang(L"tabs.monitors", L"Monitors");
            item.pszText = const_cast<LPWSTR>(tab0.c_str());
            TabCtrl_SetItem(g_tab, 0, &item);
            auto tab1 = Lang(L"tabs.presets", L"Presets");
            item.pszText = const_cast<LPWSTR>(tab1.c_str());
            TabCtrl_SetItem(g_tab, 1, &item);
            auto tab2 = Lang(L"tabs.rules", L"Rules");
            item.pszText = const_cast<LPWSTR>(tab2.c_str());
            TabCtrl_SetItem(g_tab, 2, &item);
        }

        SetListViewColumnText(0, Lang(L"list.title", L"Title"));
        SetListViewColumnText(1, Lang(L"list.process", L"Process"));
        SetListViewColumnText(2, Lang(L"list.class", L"Class"));
        SetListViewColumnText(3, Lang(L"list.monitor", L"Monitor"));
        SetListViewColumnText(4, Lang(L"list.hwnd", L"HWND"));

        if (g_languageCombo != nullptr)
        {
            SetControlTextById(g_rulesPage, IDC_LANGUAGE_LABEL, Lang(L"rules.language", L"Language"));
            SendMessageW(g_languageCombo, CB_RESETCONTENT, 0, 0);
            SendMessageW(g_languageCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"language.english", L"English").c_str()));
            SendMessageW(g_languageCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"language.polish", L"Polski").c_str()));
            SendMessageW(g_languageCombo, CB_SETCURSEL, _wcsicmp(g_languageCode.c_str(), L"pl") == 0 ? 1 : 0, 0);
        }

        if (g_appScopeOverrideCombo != nullptr)
        {
            const auto currentSelection = static_cast<int>(g_appSelectionOverride);
            SendMessageW(g_appScopeOverrideCombo, CB_RESETCONTENT, 0, 0);
            SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.follow_preset", L"Follow preset").c_str()));
            SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.selected_rows", L"Selected rows").c_str()));
            SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.all_visible_rows", L"All visible rows").c_str()));
            SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.all_discovered_windows", L"All discovered windows").c_str()));
            SendMessageW(g_appScopeOverrideCombo, CB_SETCURSEL, currentSelection, 0);
        }

        if (g_mainTooltips == nullptr)
        {
            g_mainTooltips = CreateTooltipWindow(g_mainWindow);
        }

        if (g_trayAdded)
        {
            g_trayIcon.uFlags = NIF_TIP;
            const auto trayTip = Lang(L"app.title", L"WindowLayouter.Native");
            lstrcpynW(g_trayIcon.szTip, trayTip.c_str(), static_cast<int>(std::size(g_trayIcon.szTip)));
            Shell_NotifyIconW(NIM_MODIFY, &g_trayIcon);
            g_trayIcon.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
        }

        SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_mainWindow, IDC_REFRESH), Tip(L"main.refresh", L"Reload the detected window and monitor list."));
        SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_mainWindow, IDC_CLEAR_SELECTION), Tip(L"main.clear_selection", L"Clear all selected rows on the window list."));
        SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_mainWindow, IDC_APPLY_SELECTED_PRESET), Tip(L"main.apply_selected_preset", L"Apply the preset currently selected in the preset dashboard or editor."));
        SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_mainWindow, IDC_OPEN_PRESET_EDITOR), Tip(L"main.open_preset_editor", L"Open the separate preset editor window."));
        SetToolTipText(g_mainTooltips, g_mainWindow, g_filterTitleEdit, Tip(L"main.filter_title", L"Filter discovered windows by title wildcard, for example *Messenger*."));
        SetToolTipText(g_mainTooltips, g_mainWindow, g_filterProcessEdit, Tip(L"main.filter_process", L"Filter discovered windows by process wildcard, for example *chrome*."));
        SetToolTipText(g_mainTooltips, g_mainWindow, g_listView, Tip(L"main.window_list", L"Window list with multi-select, sorting and filter support."));
        SetToolTipText(g_mainTooltips, g_mainWindow, g_appScopeOverrideCombo, Tip(L"rules.scope_override", L"Override every preset scope with one global selection mode."));
        SetToolTipText(g_mainTooltips, g_mainWindow, g_languageCombo, Tip(L"rules.language", L"Choose UI language loaded from an external language file."));
        SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_rulesPage, IDC_OPEN_ABOUT), Tip(L"about.open", L"Open application information, version and support details."));
        for (const auto& preset : g_presets)
        {
            SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_mainWindow, preset.commandId), BuildPresetDescription(preset));
        }

        UpdateStatus(Lang(L"status.ready", L"Ready."));
    }

    void ApplyPresetEditorLocalization()
    {
        if (g_presetEditorWindow == nullptr)
        {
            return;
        }

        SetWindowTextW(g_presetEditorWindow, Lang(L"preset_editor.window_title", L"WindowLayouter Preset Editor").c_str());
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_EDITOR_TITLE, Lang(L"preset_editor.title", L"Preset Editor"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_NAME_LABEL, Lang(L"preset_editor.name", L"Preset Name"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_HOTKEY_LABEL, Lang(L"preset_editor.hotkey", L"Hotkey"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_TRAY_LABEL_TEXT, Lang(L"preset_editor.tray_label", L"Tray Label"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_BUTTON_LABEL_TEXT, Lang(L"preset_editor.toolbar_label", L"Toolbar Label"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_SCOPE_LABEL, Lang(L"preset_editor.scope", L"Selection Scope"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_KIND_LABEL, Lang(L"preset_editor.mode", L"Mode"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_RESOLUTION_LABEL, Lang(L"preset_editor.resolution", L"Resolution Class"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_DEVICE_LABEL, Lang(L"preset_editor.device", L"Monitor Device"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_COLUMNS_LABEL, Lang(L"preset_editor.columns", L"Grid Columns"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_ROWS_LABEL, Lang(L"preset_editor.rows", L"Grid Rows"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_WIDTH_LABEL, Lang(L"preset_editor.right_stack_width", L"Right Stack Width %"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_LEFT_TITLE_LABEL, Lang(L"preset_editor.left_title", L"Left Title"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_LEFT_PROCESS_LABEL, Lang(L"preset_editor.left_process", L"Left Process"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_RIGHT_TITLE_LABEL, Lang(L"preset_editor.right_title", L"Right Title"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_RIGHT_PROCESS_LABEL, Lang(L"preset_editor.right_process", L"Right Process"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_REQUIRE_EXTERNAL, Lang(L"preset_editor.require_external", L"Require external monitor"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_SHOW_TRAY, Lang(L"preset_editor.show_in_tray", L"Show in tray"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_SHOW_BUTTON, Lang(L"preset_editor.show_as_toolbar", L"Show as toolbar button"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_SAVE, Lang(L"preset_editor.save", L"Save Preset"));
        SetControlTextById(g_presetEditorWindow, IDC_PRESET_APPLY_FORM, Lang(L"preset_editor.apply", L"Apply Preset"));

        if (g_presetScopeCombo != nullptr)
        {
            const auto currentSelection = g_selectedPresetIndex >= 0 && g_selectedPresetIndex < static_cast<int>(g_presets.size())
                ? static_cast<int>(g_presets[static_cast<size_t>(g_selectedPresetIndex)].selectionScope)
                : 0;
            SendMessageW(g_presetScopeCombo, CB_RESETCONTENT, 0, 0);
            SendMessageW(g_presetScopeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.selected_rows", L"Selected rows").c_str()));
            SendMessageW(g_presetScopeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.all_visible_rows", L"All visible rows").c_str()));
            SendMessageW(g_presetScopeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"scope.all_discovered_windows", L"All discovered windows").c_str()));
            SendMessageW(g_presetScopeCombo, CB_SETCURSEL, currentSelection, 0);
        }

        if (g_presetKindCombo != nullptr)
        {
            const auto currentSelection = SendMessageW(g_presetKindCombo, CB_GETCURSEL, 0, 0);
            SendMessageW(g_presetKindCombo, CB_RESETCONTENT, 0, 0);
            SendMessageW(g_presetKindCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"kind.grid", L"Grid").c_str()));
            SendMessageW(g_presetKindCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"kind.split", L"Split").c_str()));
            SendMessageW(g_presetKindCombo, CB_SETCURSEL, currentSelection == CB_ERR ? 0 : currentSelection, 0);
        }

        if (g_presetResolutionCombo != nullptr)
        {
            const auto currentSelection = SendMessageW(g_presetResolutionCombo, CB_GETCURSEL, 0, 0);
            SendMessageW(g_presetResolutionCombo, CB_RESETCONTENT, 0, 0);
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(Lang(L"resolution.any", L"Any").c_str()));
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"2K"));
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"2.5K"));
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"4K"));
            SendMessageW(g_presetResolutionCombo, CB_SETCURSEL, currentSelection == CB_ERR ? 0 : currentSelection, 0);
        }

        if (g_presetEditorTooltips == nullptr)
        {
            g_presetEditorTooltips = CreateTooltipWindow(g_presetEditorWindow);
        }

        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetList, Tip(L"preset_editor.list", L"Choose which preset you want to edit."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetNameEdit, Tip(L"preset_editor.name", L"User-visible preset name."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetHotkeyEdit, Tip(L"preset_editor.hotkey", L"Global hotkey, for example Ctrl+Alt+1."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetScopeCombo, Tip(L"preset_editor.scope", L"Selected rows, all visible rows, or all discovered windows."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetResolutionCombo, Tip(L"preset_editor.resolution", L"Activate preset only when a monitor of this class is connected."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetDeviceCombo, Tip(L"preset_editor.device", L"Optional monitor deviceName match such as \\\\.\\DISPLAY2."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, g_presetPreview, Tip(L"preset_editor.preview", L"Visual preview of the current tile structure."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, GetDlgItem(g_presetEditorWindow, IDC_PRESET_SAVE), Tip(L"preset_editor.save", L"Save changes to settings.ini."));
        SetToolTipText(g_presetEditorTooltips, g_presetEditorWindow, GetDlgItem(g_presetEditorWindow, IDC_PRESET_APPLY_FORM), Tip(L"preset_editor.apply", L"Save and apply the edited preset immediately."));
    }

    void ApplyAboutLocalization()
    {
        if (g_aboutWindow == nullptr)
        {
            return;
        }

        SetWindowTextW(g_aboutWindow, Lang(L"about.window_title", L"About WindowLayouter").c_str());
        SetControlTextById(g_aboutWindow, IDC_ABOUT_TITLE, Lang(L"about.title", L"About"));
        SetControlTextById(g_aboutWindow, IDC_ABOUT_CLOSE, Lang(L"about.close", L"Close"));
        if (g_aboutText != nullptr)
        {
            SetWindowTextW(g_aboutText, BuildAboutText().c_str());
        }

        if (g_mainTooltips == nullptr)
        {
            g_mainTooltips = CreateTooltipWindow(g_mainWindow);
        }
        SetToolTipText(g_mainTooltips, g_mainWindow, GetDlgItem(g_rulesPage, IDC_OPEN_ABOUT), Tip(L"about.open", L"Open application information, version and support details."));
    }

    void ReloadLanguageIfChanged()
    {
        const auto currentStamp = GetLanguageFilesStamp();
        if (currentStamp == g_languageFilesStamp)
        {
            return;
        }

        LoadLanguageResources(g_languageCode);
        ApplyMainWindowLocalization();
        ApplyPresetEditorLocalization();
        ApplyAboutLocalization();
        UpdateTabContent();
        RefreshPresetEditorPreviewState();
    }

    void SaveIniData()
    {
        std::wofstream output{std::filesystem::path(g_iniPath), std::ios::trunc};
        output << L"// WindowLayouter.Native settings\n";
        output << L"// Inline comments after // are ignored by the app parser.\n\n";

        RECT windowRect{CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT + 1840, CW_USEDEFAULT + 1040};
        if (g_mainWindow != nullptr)
        {
            RECT actualRect{};
            if (GetWindowRect(g_mainWindow, &actualRect))
            {
                windowRect = actualRect;
            }
        }

        const int tabIndex = g_tab == nullptr ? 0 : TabCtrl_GetCurSel(g_tab);

        output << L"[Window]\n";
        output << L"left=" << windowRect.left << L" // Main window left position in screen coordinates.\n";
        output << L"top=" << windowRect.top << L" // Main window top position in screen coordinates.\n";
        output << L"width=" << (windowRect.right - windowRect.left) << L" // Main window width in pixels.\n";
        output << L"height=" << (windowRect.bottom - windowRect.top) << L" // Main window height in pixels.\n";
        output << L"tab=" << tabIndex << L" // Active tab index: 0=Monitors, 1=Presets, 2=Rules.\n";
        output << L"window_gap_px=" << g_windowGapPx << L" // Gap in pixels between arranged windows. Use 0 for no gap.\n\n";

        output << L"[App]\n";
        output << L"language=" << g_languageCode << L" // UI language loaded from external file in lang\\<code>.ini.\n";
        output << L"selection_scope_override=" << FormatAppSelectionOverride(g_appSelectionOverride) << L" // Follow preset, Selected rows, All visible rows, All discovered windows.\n\n";

        output << L"[Hotkeys]\n";
        for (const auto& preset : g_presets)
        {
            output << preset.iniKey << L"=" << FormatHotkeyText(preset.hotkey)
                   << L" // Global hotkey for preset: " << preset.name << L".\n";
        }

        output << L"\n[Presets]\n";
        output << L"count=" << static_cast<int>(g_presets.size()) << L" // Number of editable built-in presets.\n\n";

        for (size_t index = 0; index < g_presets.size(); ++index)
        {
            const auto& preset = g_presets[index];
            output << L"[Preset" << index << L"]\n";
            output << L"name=" << preset.name << L" // Preset label used in UI and tray.\n";
            output << L"tray_label=" << preset.trayLabel << L" // Label used in tray menu.\n";
            output << L"toolbar_label=" << preset.toolbarLabel << L" // Label used on main window button.\n";
            output << L"show_in_tray=" << (preset.showInTray ? 1 : 0) << L" // 1 shows preset in tray menu.\n";
            output << L"show_as_toolbar_button=" << (preset.showAsToolbarButton ? 1 : 0) << L" // 1 shows preset button in main UI.\n";
            output << L"selection_scope=" << FormatSelectionScope(preset.selectionScope) << L" // Selected rows, All visible rows, All discovered windows.\n";
            output << L"require_external_monitor=" << (preset.requireExternalMonitor ? 1 : 0) << L" // 1 requires a non-primary monitor to be connected.\n";
            output << L"required_resolution_class=" << FormatResolutionClass(preset.requiredResolutionClass) << L" // Any, 2K, 2.5K or 4K.\n";
            output << L"required_monitor_device=" << preset.requiredMonitorDevice << L" // Device name match, for example \\\\.\\DISPLAY2, or empty for Any.\n";
            output << L"kind=" << (preset.kind == PresetKind::Grid ? L"Grid" : L"Split") << L" // Grid or Split.\n";
            output << L"columns=" << preset.columns << L" // Grid columns for AG or full-grid layout.\n";
            output << L"rows=" << preset.rows << L" // Grid rows for AG or full-grid layout.\n";
            output << L"chrome_width_percent=" << preset.chromeWidthPercent << L" // Width of the right stacked area in split layout.\n";
            output << L"left_title=" << preset.leftTitlePattern << L" // Title wildcard for left/AG windows.\n";
            output << L"left_process=" << preset.leftProcessPattern << L" // Process wildcard for left/AG windows.\n";
            output << L"right_title=" << preset.rightTitlePattern << L" // Title wildcard for right/Chrome windows.\n";
            output << L"right_process=" << preset.rightProcessPattern << L" // Process wildcard for right/Chrome windows.\n\n";
        }
    }

    void EnsureIniFileExists()
    {
        if (std::filesystem::exists(std::filesystem::path(g_iniPath)))
        {
            return;
        }

        SaveIniData();
    }

    HotkeySpec ParseHotkeyText(const std::wstring& rawText)
    {
        HotkeySpec spec{};
        spec.text = Trim(rawText);
        if (spec.text.empty())
        {
            return spec;
        }

        auto parts = Split(spec.text, L'+');
        if (parts.empty())
        {
            return spec;
        }

        std::wstring keyPart;
        for (auto& part : parts)
        {
            auto upper = ToUpper(Trim(part));
            if (upper == L"CTRL" || upper == L"CONTROL")
            {
                spec.modifiers |= MOD_CONTROL;
            }
            else if (upper == L"ALT")
            {
                spec.modifiers |= MOD_ALT;
            }
            else if (upper == L"SHIFT")
            {
                spec.modifiers |= MOD_SHIFT;
            }
            else if (upper == L"WIN" || upper == L"WINDOWS")
            {
                spec.modifiers |= MOD_WIN;
            }
            else
            {
                keyPart = upper;
            }
        }

        if (keyPart.empty())
        {
            return spec;
        }

        if (keyPart.size() == 1)
        {
            const auto character = keyPart[0];
            if ((character >= L'A' && character <= L'Z') || (character >= L'0' && character <= L'9'))
            {
                spec.virtualKey = static_cast<UINT>(character);
                spec.valid = true;
                return spec;
            }
        }

        if (keyPart[0] == L'F')
        {
            const auto functionNumber = _wtoi(keyPart.c_str() + 1);
            if (functionNumber >= 1 && functionNumber <= 24)
            {
                spec.virtualKey = static_cast<UINT>(VK_F1 + functionNumber - 1);
                spec.valid = true;
                return spec;
            }
        }

        return spec;
    }

    std::wstring FormatHotkeyText(const HotkeySpec& hotkey)
    {
        if (!hotkey.valid)
        {
            return L"(disabled)";
        }

        std::wstringstream text;
        bool hasModifier = false;
        if ((hotkey.modifiers & MOD_CONTROL) != 0)
        {
            text << L"Ctrl";
            hasModifier = true;
        }
        if ((hotkey.modifiers & MOD_ALT) != 0)
        {
            if (hasModifier)
            {
                text << L"+";
            }
            text << L"Alt";
            hasModifier = true;
        }
        if ((hotkey.modifiers & MOD_SHIFT) != 0)
        {
            if (hasModifier)
            {
                text << L"+";
            }
            text << L"Shift";
            hasModifier = true;
        }
        if ((hotkey.modifiers & MOD_WIN) != 0)
        {
            if (hasModifier)
            {
                text << L"+";
            }
            text << L"Win";
            hasModifier = true;
        }

        if (hasModifier)
        {
            text << L"+";
        }

        if (hotkey.virtualKey >= VK_F1 && hotkey.virtualKey <= VK_F24)
        {
            text << L"F" << (hotkey.virtualKey - VK_F1 + 1);
        }
        else
        {
            text << static_cast<wchar_t>(hotkey.virtualKey);
        }

        return text.str();
    }

    std::wstring BuildPresetDescription(const PresetDefinition& preset)
    {
        std::wstringstream text;
        if (preset.kind == PresetKind::Grid)
        {
            text << Lang(L"status.preset_desc_grid_prefix", L"Selected windows on target monitor in ")
                 << preset.columns << L"x" << preset.rows << Lang(L"status.preset_desc_grid_suffix", L" grid.");
        }
        else
        {
            text << Lang(L"status.preset_desc_split_prefix", L"Split layout: AG on left ")
                 << (100 - preset.chromeWidthPercent) << L"% "
                 << Lang(L"status.preset_desc_split_middle", L"as ")
                 << preset.columns << L"x" << preset.rows
                 << Lang(L"status.preset_desc_split_suffix", L", Chrome stacked on right ")
                 << preset.chromeWidthPercent << L"%.";
        }
        text << L" " << Lang(L"status.selection_scope_label", L"Selection scope") << L": ";
        if (g_appSelectionOverride == AppSelectionOverride::FollowPreset)
        {
            text << FormatSelectionScope(preset.selectionScope);
        }
        else
        {
            text << FormatAppSelectionOverride(g_appSelectionOverride) << L" (" << Lang(L"rules.app_override_short", L"app override") << L")";
        }
        text << L" " << Lang(L"status.activation_label", L"Activation") << L": ";
        if (!preset.requireExternalMonitor && preset.requiredResolutionClass == ResolutionClass::Any)
        {
            if (Trim(preset.requiredMonitorDevice).empty())
            {
                text << Lang(L"status.activation_always", L"always active.");
            }
            else
            {
                text << Lang(L"status.activation_requires_device", L"requires device ") << preset.requiredMonitorDevice << L".";
            }
        }
        else
        {
            if (preset.requireExternalMonitor)
            {
                text << Lang(L"status.activation_requires_external", L"requires external monitor");
                if (preset.requiredResolutionClass != ResolutionClass::Any)
                {
                    text << L" " << Lang(L"status.activation_with", L"with") << L" " << FormatResolutionClass(preset.requiredResolutionClass);
                }
                if (!Trim(preset.requiredMonitorDevice).empty())
                {
                    text << L" " << Lang(L"status.activation_named", L"named") << L" " << preset.requiredMonitorDevice;
                }
                text << L".";
            }
            else
            {
                if (!Trim(preset.requiredMonitorDevice).empty() && preset.requiredResolutionClass != ResolutionClass::Any)
                {
                    text << Lang(L"status.activation_requires", L"requires ") << FormatResolutionClass(preset.requiredResolutionClass) << L" " << Lang(L"status.activation_monitor_named", L"monitor named ") << preset.requiredMonitorDevice << L".";
                }
                else if (!Trim(preset.requiredMonitorDevice).empty())
                {
                    text << Lang(L"status.activation_monitor_named_prefix", L"requires monitor named ") << preset.requiredMonitorDevice << L".";
                }
                else
                {
                    text << Lang(L"status.activation_monitor_class", L"requires monitor class ") << FormatResolutionClass(preset.requiredResolutionClass) << L".";
                }
            }
        }
        return text.str();
    }

    bool MonitorMatchesPresetActivation(const MonitorInfo& monitor, const PresetDefinition& preset)
    {
        if (preset.requireExternalMonitor && monitor.isPrimary)
        {
            return false;
        }

        if (preset.requiredResolutionClass != ResolutionClass::Any &&
            DetectResolutionClass(monitor) != preset.requiredResolutionClass)
        {
            return false;
        }

        if (!Trim(preset.requiredMonitorDevice).empty() &&
            !WildcardMatch(preset.requiredMonitorDevice, monitor.deviceName))
        {
            return false;
        }

        return true;
    }

    const MonitorInfo* GetPresetTargetMonitor(const PresetDefinition& preset)
    {
        for (const auto& monitor : g_monitors)
        {
            if (MonitorMatchesPresetActivation(monitor, preset))
            {
                return &monitor;
            }
        }

        if (!preset.requireExternalMonitor && preset.requiredResolutionClass == ResolutionClass::Any && Trim(preset.requiredMonitorDevice).empty())
        {
            return GetPrimaryMonitor();
        }

        return nullptr;
    }

    std::wstring BuildPresetInactiveReason(const PresetDefinition& preset)
    {
        std::wstringstream text;
        text << Lang(L"status.preset_inactive_prefix", L"Preset inactive: ");
        if (preset.requireExternalMonitor && preset.requiredResolutionClass != ResolutionClass::Any && !Trim(preset.requiredMonitorDevice).empty())
        {
            text << Lang(L"status.connect_external_monitor", L"connect external monitor ") << preset.requiredMonitorDevice << L" " << Lang(L"status.in_class", L"in class ") << FormatResolutionClass(preset.requiredResolutionClass) << L".";
        }
        else if (preset.requireExternalMonitor && preset.requiredResolutionClass != ResolutionClass::Any)
        {
            text << Lang(L"status.connect_external_class_prefix", L"connect an external ") << FormatResolutionClass(preset.requiredResolutionClass) << L" " << Lang(L"status.connect_monitor_suffix", L"monitor.") ;
        }
        else if (preset.requireExternalMonitor && !Trim(preset.requiredMonitorDevice).empty())
        {
            text << Lang(L"status.connect_external_monitor", L"connect external monitor ") << preset.requiredMonitorDevice << L".";
        }
        else if (preset.requireExternalMonitor)
        {
            text << Lang(L"status.connect_any_external", L"connect any external monitor.");
        }
        else if (!Trim(preset.requiredMonitorDevice).empty() && preset.requiredResolutionClass != ResolutionClass::Any)
        {
            text << Lang(L"status.connect_class_monitor_prefix", L"connect a ") << FormatResolutionClass(preset.requiredResolutionClass) << L" " << Lang(L"status.connect_monitor_named_mid", L"monitor ") << preset.requiredMonitorDevice << L".";
        }
        else if (!Trim(preset.requiredMonitorDevice).empty())
        {
            text << Lang(L"status.connect_monitor_named_prefix", L"connect monitor ") << preset.requiredMonitorDevice << L".";
        }
        else if (preset.requiredResolutionClass != ResolutionClass::Any)
        {
            text << Lang(L"status.connect_class_monitor_prefix", L"connect a ") << FormatResolutionClass(preset.requiredResolutionClass) << L" " << Lang(L"status.connect_monitor_suffix", L"monitor.");
        }
        else
        {
            text << Lang(L"status.monitor_condition_not_satisfied", L"matching monitor condition not satisfied.");
        }

        return text.str();
    }

    void RefreshMonitorDeviceOptions(HWND combo, const std::wstring& selectedValue)
    {
        if (combo == nullptr)
        {
            return;
        }

        SendMessageW(combo, CB_RESETCONTENT, 0, 0);
        const auto anyLabel = Lang(L"resolution.any", L"Any");
        SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(anyLabel.c_str()));
        int selectedIndex = 0;

        for (size_t index = 0; index < g_monitors.size(); ++index)
        {
            const auto& deviceName = g_monitors[index].deviceName;
            const int itemIndex = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(deviceName.c_str())));
            if (_wcsicmp(selectedValue.c_str(), deviceName.c_str()) == 0)
            {
                selectedIndex = itemIndex;
            }
        }

        SendMessageW(combo, CB_SETCURSEL, selectedIndex, 0);
    }

    std::vector<PresetDefinition> BuildDefaultPresets()
    {
        return
        {
            {IDC_PRESET_GRID_2X2, HOTKEY_BASE_ID + 0, L"PresetGrid2x2", Lang(L"preset.grid2x2.name", L"4K > 2x2"), Lang(L"preset.grid2x2.tray", L"4K > 2x2"), Lang(L"preset.grid2x2.button", L"4K > 2x2"), true, true, SelectionScope::SelectedRows, PresetKind::Grid, 2, 2, 0, L"*Antigravity*", L"", L"*Chrome*", L"*chrome*", false, ResolutionClass::R4K, L"", ParseHotkeyText(L"Ctrl+Alt+1")},
            {IDC_PRESET_SPLIT_2X2, HOTKEY_BASE_ID + 1, L"PresetSplit2x2", Lang(L"preset.split2x2.name", L"4K > AG 2x2 + Chrome"), Lang(L"preset.split2x2.tray", L"AG 2x2 + Chrome"), Lang(L"preset.split2x2.button", L"AG 2x2 + Chrome"), true, true, SelectionScope::SelectedRows, PresetKind::Split, 2, 2, 50, L"*Antigravity*", L"", L"*Chrome*", L"*chrome*", false, ResolutionClass::R4K, L"", ParseHotkeyText(L"Ctrl+Alt+2")},
            {IDC_PRESET_SPLIT_2X3, HOTKEY_BASE_ID + 2, L"PresetSplit2x3", Lang(L"preset.split2x3.name", L"4K > AG 2x3 + Chrome"), Lang(L"preset.split2x3.tray", L"AG 2x3 + Chrome"), Lang(L"preset.split2x3.button", L"AG 2x3 + Chrome"), true, true, SelectionScope::SelectedRows, PresetKind::Split, 2, 3, 50, L"*Antigravity*", L"", L"*Chrome*", L"*chrome*", false, ResolutionClass::R4K, L"", ParseHotkeyText(L"Ctrl+Alt+3")},
            {IDC_PRESET_SPLIT_3X2, HOTKEY_BASE_ID + 3, L"PresetSplit3x2", Lang(L"preset.split3x2.name", L"4K > AG 3x2 + Chrome 1/3"), Lang(L"preset.split3x2.tray", L"AG 3x2 + Chrome 1/3"), Lang(L"preset.split3x2.button", L"AG 3x2 + Chrome 1/3"), true, true, SelectionScope::SelectedRows, PresetKind::Split, 3, 2, 33, L"*Antigravity*", L"", L"*Chrome*", L"*chrome*", false, ResolutionClass::R4K, L"", ParseHotkeyText(L"Ctrl+Alt+4")},
            {IDC_PRESET_SPLIT_3X3, HOTKEY_BASE_ID + 4, L"PresetSplit3x3", Lang(L"preset.split3x3.name", L"4K > AG 3x3 + Chrome 1/3"), Lang(L"preset.split3x3.tray", L"AG 3x3 + Chrome 1/3"), Lang(L"preset.split3x3.button", L"AG 3x3 + Chrome 1/3"), true, true, SelectionScope::SelectedRows, PresetKind::Split, 3, 3, 33, L"*Antigravity*", L"", L"*Chrome*", L"*chrome*", false, ResolutionClass::R4K, L"", ParseHotkeyText(L"Ctrl+Alt+5")}
        };
    }

    std::wstring BuildMonitorText()
    {
        std::wstringstream text;
        text << Lang(L"monitors.header", L"Detected monitors") << L"\n";
        text << L"INI: " << g_iniPath << L"\n\n";
        for (const auto& monitor : g_monitors)
        {
            text << (monitor.isPrimary ? Lang(L"monitors.primary_prefix", L"[Primary] ") : L"          ");
            text << monitor.deviceName << L"\n";
            text << Lang(L"monitors.bounds", L"Bounds") << L":   " << monitor.bounds.left << L"," << monitor.bounds.top << L"  "
                 << monitor.bounds.width << L"x" << monitor.bounds.height << L"\n";
            text << Lang(L"monitors.workarea", L"WorkArea") << L": " << monitor.workArea.left << L"," << monitor.workArea.top << L"  "
                 << monitor.workArea.width << L"x" << monitor.workArea.height << L"\n";
            text << Lang(L"monitors.class", L"Class") << L":    " << FormatResolutionClass(DetectResolutionClass(monitor))
                 << (monitor.isPrimary ? Lang(L"monitors.primary_suffix", L" / built-in or primary") : Lang(L"monitors.external_suffix", L" / external")) << L"\n";
            text << L"DPI:      " << monitor.dpi << L" (" << static_cast<int>(monitor.dpi / 96.0 * 100.0 + 0.5) << L"%)\n\n";
        }

        if (g_monitors.empty())
        {
            text << Lang(L"monitors.none", L"No monitors detected.") << L"\n";
        }

        return text.str();
    }

    std::wstring BuildAboutText()
    {
        std::wstringstream text;
        text << Lang(L"about.product_name", L"WindowLayouter.Native") << L"\r\n";
        text << Lang(L"about.version_label", L"Version") << L": " << WL_VERSION_DOT << L"\r\n\r\n";
        text << Lang(L"about.description", L"Desktop tool for monitor-aware window layouts and preset-driven tiling.") << L"\r\n\r\n";
        text << Lang(L"about.publisher_label", L"Publisher") << L": " << Lang(L"about.publisher_value", L"TODO: Publisher") << L"\r\n";
        text << Lang(L"about.website_label", L"Website") << L": " << Lang(L"about.website_value", L"TODO: Website") << L"\r\n";
        text << Lang(L"about.support_label", L"Support") << L": " << Lang(L"about.support_value", L"TODO: Support contact") << L"\r\n";
        text << Lang(L"about.license_label", L"License") << L": " << Lang(L"about.license_value", L"TODO: License") << L"\r\n";
        text << Lang(L"about.build_label", L"Build") << L": " << WL_VERSION_DOT << L"\r\n\r\n";
        text << Lang(L"about.notes", L"Placeholder metadata. Fill publisher, support, website and license details before release.") << L"\r\n";
        return text.str();
    }

    std::wstring BuildRulesText()
    {
        std::wstringstream text;
        text << Lang(L"rules.header", L"RULES") << L"\n\n";
        text << Lang(L"rules.selection_header", L"Selection") << L"\n";
        text << L"  " << Lang(L"rules.selection_line1", L"Use row selection on the window list. There are no checkboxes.") << L"\n";
        text << L"  " << Lang(L"rules.selection_line2", L"App selection scope override") << L": " << FormatAppSelectionOverride(g_appSelectionOverride) << L"\n";
        text << L"  " << Lang(L"rules.selection_line3", L"Presets can use Selected rows, All visible rows, or All discovered windows.") << L"\n\n";
        text << Lang(L"rules.filters_header", L"Filters") << L"\n";
        text << L"  " << Lang(L"rules.filters_line1", L"Top filters support wildcard matching for title and process.") << L"\n";
        text << L"  " << Lang(L"rules.filters_line2", L"Example: process *chrome* and title *messenger*.") << L"\n\n";
        text << Lang(L"rules.layouts_header", L"Layouts") << L"\n";
        text << L"  " << Lang(L"rules.layouts_line1", L"Grid presets use the selected scope and fill a tile grid.") << L"\n";
        text << L"  " << Lang(L"rules.layouts_line2", L"Split presets route AG windows to the left grid and Chrome windows to the right stack.") << L"\n";
        text << L"  " << Lang(L"rules.layouts_line3", L"Left grid supports nested tile sizes such as 3x3 with a 2/3 split.") << L"\n\n";
        text << Lang(L"rules.activation_header", L"Activation") << L"\n";
        text << L"  " << Lang(L"rules.activation_line1", L"Each preset can require an external monitor, a resolution class and a specific deviceName.") << L"\n";
        text << L"  " << Lang(L"rules.activation_line2", L"If the required monitor is not connected, the preset is disabled in toolbar and tray.") << L"\n";
        text << L"  " << Lang(L"rules.activation_line3", L"When active, the preset uses the first matching monitor as its target.") << L"\n\n";
        text << Lang(L"rules.spacing_header", L"Spacing") << L"\n";
        text << L"  " << Lang(L"rules.spacing_line1", L"window_gap_px controls the gap between arranged windows and column edges.") << L"\n\n";
        text << Lang(L"rules.ordering_header", L"Ordering") << L"\n";
        text << L"  " << Lang(L"rules.ordering_line1", L"Matching order follows system z-order after filtering.") << L"\n";
        return text.str();
    }

    void SetControlText(HWND control, const std::wstring& text)
    {
        SetWindowTextW(control, text.c_str());
    }

    void RefreshToolbarPresetButtons()
    {
        for (size_t index = 0; index < g_presets.size(); ++index)
        {
            auto* button = GetDlgItem(g_mainWindow, g_presets[index].commandId);
            if (button == nullptr)
            {
                continue;
            }

            SetWindowTextW(button, g_presets[index].toolbarLabel.c_str());
            EnableWindow(button, GetPresetTargetMonitor(g_presets[index]) != nullptr);
            ShowWindow(button, g_presets[index].showAsToolbarButton ? SW_SHOW : SW_HIDE);
        }
    }

    void RefreshPresetListBox()
    {
        if (g_presetList == nullptr)
        {
            return;
        }

        SendMessageW(g_presetList, LB_RESETCONTENT, 0, 0);
        for (const auto& preset : g_presets)
        {
            SendMessageW(g_presetList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(preset.name.c_str()));
        }
        SendMessageW(g_presetList, LB_SETCURSEL, g_selectedPresetIndex, 0);
    }

    void LoadPresetIntoEditor(int index)
    {
        if (index < 0 || index >= static_cast<int>(g_presets.size()))
        {
            return;
        }

        g_selectedPresetIndex = index;
        const auto& preset = g_presets[static_cast<size_t>(index)];
        SetControlText(g_presetNameEdit, preset.name);
        SetControlText(g_presetHotkeyEdit, preset.hotkey.text);
        SetControlText(g_presetTrayLabelEdit, preset.trayLabel);
        SetControlText(g_presetButtonLabelEdit, preset.toolbarLabel);
        SendMessageW(g_presetKindCombo, CB_SETCURSEL, preset.kind == PresetKind::Grid ? 0 : 1, 0);
        RefreshWidthFieldState();
        SetControlText(g_presetColumnsEdit, std::to_wstring(preset.columns));
        SetControlText(g_presetRowsEdit, std::to_wstring(preset.rows));
        SetControlText(g_presetWidthEdit, std::to_wstring(preset.chromeWidthPercent));
        SetControlText(g_presetLeftTitleEdit, preset.leftTitlePattern);
        SetControlText(g_presetLeftProcessEdit, preset.leftProcessPattern);
        SetControlText(g_presetRightTitleEdit, preset.rightTitlePattern);
        SetControlText(g_presetRightProcessEdit, preset.rightProcessPattern);
        SendMessageW(g_presetShowTrayCheck, BM_SETCHECK, preset.showInTray ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(g_presetShowButtonCheck, BM_SETCHECK, preset.showAsToolbarButton ? BST_CHECKED : BST_UNCHECKED, 0);
        if (g_presetScopeCombo != nullptr)
        {
            SendMessageW(g_presetScopeCombo, CB_SETCURSEL, static_cast<WPARAM>(preset.selectionScope), 0);
        }
        SendMessageW(g_presetRequireExternalCheck, BM_SETCHECK, preset.requireExternalMonitor ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(g_presetResolutionCombo, CB_SETCURSEL, static_cast<WPARAM>(preset.requiredResolutionClass), 0);
        RefreshMonitorDeviceOptions(g_presetDeviceCombo, preset.requiredMonitorDevice);
        SetControlText(g_presetStatus, GetPresetTargetMonitor(preset) != nullptr ? BuildPresetDescription(preset) : BuildPresetInactiveReason(preset));
        SendMessageW(g_presetList, LB_SETCURSEL, index, 0);
        if (g_presetPreview != nullptr)
        {
            InvalidateRect(g_presetPreview, nullptr, TRUE);
        }
    }

    void RefreshPresetEditorPreviewState()
    {
        if (g_presetEditorWindow == nullptr || g_presetStatus == nullptr)
        {
            return;
        }

        const auto previewPreset = BuildPreviewPresetFromEditor();
        const bool isActive = GetPresetTargetMonitor(previewPreset) != nullptr;
        SetControlText(g_presetStatus, isActive ? BuildPresetDescription(previewPreset) : BuildPresetInactiveReason(previewPreset));
        if (g_presetPreview != nullptr)
        {
            InvalidateRect(g_presetPreview, nullptr, TRUE);
        }
    }

    // Enable/disable the Right Stack Width % field based on the selected mode.
    // The field is only meaningful for Split mode; Grid ignores it.
    void RefreshWidthFieldState()
    {
        if (g_presetWidthEdit == nullptr || g_presetKindCombo == nullptr)
        {
            return;
        }
        const bool isSplit = SendMessageW(g_presetKindCombo, CB_GETCURSEL, 0, 0) == 1;
        EnableWindow(g_presetWidthEdit, isSplit ? TRUE : FALSE);
    }

    PresetDefinition BuildPreviewPresetFromEditor()
    {
        PresetDefinition preset = g_selectedPresetIndex >= 0 && g_selectedPresetIndex < static_cast<int>(g_presets.size())
            ? g_presets[static_cast<size_t>(g_selectedPresetIndex)]
            : BuildDefaultPresets().front();

        const auto name = Trim(GetControlText(g_presetNameEdit));
        if (!name.empty())
        {
            preset.name = name;
        }

        const auto trayLabel = Trim(GetControlText(g_presetTrayLabelEdit));
        const auto buttonLabel = Trim(GetControlText(g_presetButtonLabelEdit));
        if (!trayLabel.empty())
        {
            preset.trayLabel = trayLabel;
        }
        if (!buttonLabel.empty())
        {
            preset.toolbarLabel = buttonLabel;
        }

        const auto kindSelection = SendMessageW(g_presetKindCombo, CB_GETCURSEL, 0, 0);
        preset.kind = kindSelection == 1 ? PresetKind::Split : PresetKind::Grid;
        const auto scopeSelection = SendMessageW(g_presetScopeCombo, CB_GETCURSEL, 0, 0);
        preset.selectionScope = scopeSelection == CB_ERR ? preset.selectionScope : static_cast<SelectionScope>(scopeSelection);
        const auto resolutionSelection = SendMessageW(g_presetResolutionCombo, CB_GETCURSEL, 0, 0);
        preset.requiredResolutionClass = resolutionSelection == CB_ERR ? preset.requiredResolutionClass : static_cast<ResolutionClass>(resolutionSelection);
        preset.requireExternalMonitor = SendMessageW(g_presetRequireExternalCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
        preset.showInTray = SendMessageW(g_presetShowTrayCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
        preset.showAsToolbarButton = SendMessageW(g_presetShowButtonCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
        preset.columns = std::max(1, _wtoi(GetControlText(g_presetColumnsEdit).c_str()));
        preset.rows = std::max(1, _wtoi(GetControlText(g_presetRowsEdit).c_str()));
        preset.chromeWidthPercent = std::clamp(_wtoi(GetControlText(g_presetWidthEdit).c_str()), 20, 80);
        preset.leftTitlePattern = Trim(GetControlText(g_presetLeftTitleEdit));
        preset.leftProcessPattern = Trim(GetControlText(g_presetLeftProcessEdit));
        preset.rightTitlePattern = Trim(GetControlText(g_presetRightTitleEdit));
        preset.rightProcessPattern = Trim(GetControlText(g_presetRightProcessEdit));
        preset.requiredMonitorDevice = Trim(GetControlText(g_presetDeviceCombo));
        if (_wcsicmp(preset.requiredMonitorDevice.c_str(), Lang(L"resolution.any", L"Any").c_str()) == 0 ||
            _wcsicmp(preset.requiredMonitorDevice.c_str(), L"Any") == 0 ||
            _wcsicmp(preset.requiredMonitorDevice.c_str(), L"Dowolna") == 0)
        {
            preset.requiredMonitorDevice.clear();
        }

        return preset;
    }

    // =========================================================================
    // Preset editor — read/write between editor controls and g_presets
    // =========================================================================

    bool SavePresetFromEditor()
    {
        if (g_selectedPresetIndex < 0 || g_selectedPresetIndex >= static_cast<int>(g_presets.size()))
        {
            return false;
        }

        auto& preset = g_presets[static_cast<size_t>(g_selectedPresetIndex)];
        const auto name = Trim(GetControlText(g_presetNameEdit));
        if (name.empty())
        {
            SetControlText(g_presetStatus, Lang(L"status.preset_name_required", L"Preset name is required."));
            return false;
        }

        auto hotkey = ParseHotkeyText(GetControlText(g_presetHotkeyEdit));
        if (!Trim(GetControlText(g_presetHotkeyEdit)).empty() && !hotkey.valid)
        {
            SetControlText(g_presetStatus, Lang(L"status.hotkey_invalid", L"Hotkey format is invalid."));
            return false;
        }

        preset.name = name;
        preset.trayLabel = Trim(GetControlText(g_presetTrayLabelEdit));
        preset.toolbarLabel = Trim(GetControlText(g_presetButtonLabelEdit));
        if (preset.trayLabel.empty())
        {
            preset.trayLabel = preset.name;
        }
        if (preset.toolbarLabel.empty())
        {
            preset.toolbarLabel = preset.name;
        }
        preset.showInTray = SendMessageW(g_presetShowTrayCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
        preset.showAsToolbarButton = SendMessageW(g_presetShowButtonCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
        const auto scopeSelection = SendMessageW(g_presetScopeCombo, CB_GETCURSEL, 0, 0);
        preset.selectionScope = scopeSelection == CB_ERR ? SelectionScope::SelectedRows : static_cast<SelectionScope>(scopeSelection);
        preset.requireExternalMonitor = SendMessageW(g_presetRequireExternalCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
        const auto resolutionSelection = SendMessageW(g_presetResolutionCombo, CB_GETCURSEL, 0, 0);
        preset.requiredResolutionClass = resolutionSelection == CB_ERR
            ? ResolutionClass::Any
            : static_cast<ResolutionClass>(resolutionSelection);
        preset.requiredMonitorDevice = Trim(GetControlText(g_presetDeviceCombo));
        if (_wcsicmp(preset.requiredMonitorDevice.c_str(), L"Any") == 0 ||
            _wcsicmp(preset.requiredMonitorDevice.c_str(), Lang(L"resolution.any", L"Any").c_str()) == 0 ||
            _wcsicmp(preset.requiredMonitorDevice.c_str(), L"Dowolna") == 0)
        {
            preset.requiredMonitorDevice.clear();
        }
        preset.kind = SendMessageW(g_presetKindCombo, CB_GETCURSEL, 0, 0) == 1 ? PresetKind::Split : PresetKind::Grid;
        preset.columns = std::max(1, _wtoi(GetControlText(g_presetColumnsEdit).c_str()));
        preset.rows = std::max(1, _wtoi(GetControlText(g_presetRowsEdit).c_str()));
        preset.chromeWidthPercent = std::clamp(_wtoi(GetControlText(g_presetWidthEdit).c_str()), 20, 80);
        preset.leftTitlePattern = Trim(GetControlText(g_presetLeftTitleEdit));
        preset.leftProcessPattern = Trim(GetControlText(g_presetLeftProcessEdit));
        preset.rightTitlePattern = Trim(GetControlText(g_presetRightTitleEdit));
        preset.rightProcessPattern = Trim(GetControlText(g_presetRightProcessEdit));
        preset.hotkey = hotkey.valid ? hotkey : HotkeySpec{};
        preset.hotkey.text = Trim(GetControlText(g_presetHotkeyEdit));

        RefreshPresetListBox();
        RefreshToolbarPresetButtons();
        RegisterPresetHotkeys();
        SaveSettings();
        UpdateTabContent();
        SetControlText(g_presetStatus, Lang(L"status.preset_saved", L"Preset saved."));
        return true;
    }

    void UpdateTabContent()
    {
        SetWindowTextW(g_monitorsEdit, BuildMonitorText().c_str());
        if (g_presetsSummaryEdit != nullptr)
        {
            SetWindowTextW(g_presetsSummaryEdit, BuildPresetSummaryText().c_str());
        }
        SetWindowTextW(g_rulesEdit, BuildRulesText().c_str());
        RefreshPresetListBox();
        RefreshToolbarPresetButtons();
        if (g_selectedPresetIndex >= 0 && g_selectedPresetIndex < static_cast<int>(g_presets.size()) && g_presetStatus != nullptr)
        {
            const auto& preset = g_presets[static_cast<size_t>(g_selectedPresetIndex)];
            const bool isActive = GetPresetTargetMonitor(preset) != nullptr;
            RefreshMonitorDeviceOptions(g_presetDeviceCombo, preset.requiredMonitorDevice);
            SetControlText(g_presetStatus, isActive ? BuildPresetDescription(preset) : BuildPresetInactiveReason(preset));
            EnableWindow(GetDlgItem(g_mainWindow, IDC_APPLY_SELECTED_PRESET), isActive);
            if (g_presetEditorWindow != nullptr)
            {
                EnableWindow(GetDlgItem(g_presetEditorWindow, IDC_PRESET_APPLY_FORM), isActive);
            }
        }
        ApplyAboutLocalization();
    }

    void UpdateStatus(const std::wstring& text)
    {
        SetWindowTextW(g_status, text.c_str());
    }

    std::wstring BuildPresetSummaryText()
    {
        std::wstringstream text;
        text << Lang(L"presets.summary_header", L"PRESETS") << L"\n\n";
        text << Lang(L"presets.summary_intro", L"Open the separate Preset Editor window to change hotkeys, activation conditions and layout structure.") << L"\n\n";
        for (const auto& preset : g_presets)
        {
            text << L"- " << preset.name << L"\n";
            text << L"  " << Lang(L"presets.summary_tray", L"Tray") << L": " << (preset.showInTray ? preset.trayLabel : Lang(L"presets.hidden", L"(hidden)")) << L"\n";
            text << L"  " << Lang(L"presets.summary_button", L"Button") << L": " << (preset.showAsToolbarButton ? preset.toolbarLabel : Lang(L"presets.hidden", L"(hidden)")) << L"\n";
            text << L"  " << Lang(L"presets.summary_scope", L"Scope") << L": " << (g_appSelectionOverride == AppSelectionOverride::FollowPreset ? FormatSelectionScope(preset.selectionScope) : FormatAppSelectionOverride(g_appSelectionOverride) + L" (" + Lang(L"rules.app_override_short", L"app override") + L")") << L"\n";
            text << L"  " << Lang(L"presets.summary_activation", L"Activation") << L": " << (GetPresetTargetMonitor(preset) != nullptr ? Lang(L"presets.active", L"active") : BuildPresetInactiveReason(preset)) << L"\n";
            text << L"  " << Lang(L"presets.summary_layout", L"Layout") << L": " << BuildPresetDescription(preset) << L"\n\n";
        }
        return text.str();
    }

    void ShowPresetEditorWindow()
    {
        if (g_presetEditorWindow == nullptr)
        {
            return;
        }

        ShowWindow(g_presetEditorWindow, SW_SHOW);
        ShowWindow(g_presetEditorWindow, SW_RESTORE);
        SetForegroundWindow(g_presetEditorWindow);
        LoadPresetIntoEditor(g_selectedPresetIndex);
    }

    void ShowAboutWindow()
    {
        if (g_aboutWindow == nullptr)
        {
            return;
        }

        ShowWindow(g_aboutWindow, SW_SHOW);
        ShowWindow(g_aboutWindow, SW_RESTORE);
        SetForegroundWindow(g_aboutWindow);
    }

    void ResizePresetEditorControls()
    {
        if (g_presetEditorWindow == nullptr)
        {
            return;
        }

        RECT client{};
        GetClientRect(g_presetEditorWindow, &client);
        const UINT dpi = GetWindowDpiSafe(g_presetEditorWindow);
        auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };

        const int margin = s(16);
        const int listWidth = s(220);
        const int previewWidth = s(260);
        const int rowHeight = s(26);
        const int comboHeight = s(240);
        const int rowGap = s(8);
        const int buttonHeight = s(34);
        const int formLeft = margin + listWidth + s(16);
        const int previewLeft = client.right - previewWidth - margin;
        const int labelWidth = s(118);
        const int editWidth = std::max(s(180), previewLeft - formLeft - labelWidth - s(16));
        int rowTop = margin + s(38);

        MoveWindow(GetDlgItem(g_presetEditorWindow, IDC_PRESET_EDITOR_TITLE), margin, margin, previewLeft - margin - s(8), s(28), TRUE);
        MoveWindow(g_presetList, margin, margin + s(38), listWidth, client.bottom - margin * 2 - s(38), TRUE);
        MoveWindow(g_presetPreview, previewLeft, margin + s(38), previewWidth, s(210), TRUE);

        const int labelIds[] = {
            IDC_PRESET_NAME_LABEL,
            IDC_PRESET_HOTKEY_LABEL,
            IDC_PRESET_TRAY_LABEL_TEXT,
            IDC_PRESET_BUTTON_LABEL_TEXT,
            IDC_PRESET_SCOPE_LABEL,
            IDC_PRESET_KIND_LABEL,
            IDC_PRESET_RESOLUTION_LABEL,
            IDC_PRESET_DEVICE_LABEL,
            IDC_PRESET_COLUMNS_LABEL,
            IDC_PRESET_ROWS_LABEL,
            IDC_PRESET_WIDTH_LABEL,
            IDC_PRESET_LEFT_TITLE_LABEL,
            IDC_PRESET_LEFT_PROCESS_LABEL,
            IDC_PRESET_RIGHT_TITLE_LABEL,
            IDC_PRESET_RIGHT_PROCESS_LABEL
        };
        HWND fields[] = {
            g_presetNameEdit,
            g_presetHotkeyEdit,
            g_presetTrayLabelEdit,
            g_presetButtonLabelEdit,
            g_presetScopeCombo,
            g_presetKindCombo,
            g_presetResolutionCombo,
            g_presetDeviceCombo,
            g_presetColumnsEdit,
            g_presetRowsEdit,
            g_presetWidthEdit,
            g_presetLeftTitleEdit,
            g_presetLeftProcessEdit,
            g_presetRightTitleEdit,
            g_presetRightProcessEdit
        };

        for (int index = 0; index < 15; ++index)
        {
            MoveWindow(GetDlgItem(g_presetEditorWindow, labelIds[index]), formLeft, rowTop + s(4), labelWidth, s(20), TRUE);
            const bool isCombo = fields[index] == g_presetScopeCombo ||
                fields[index] == g_presetKindCombo ||
                fields[index] == g_presetResolutionCombo ||
                fields[index] == g_presetDeviceCombo;
            MoveWindow(fields[index], formLeft + labelWidth, rowTop, editWidth, isCombo ? comboHeight : rowHeight, TRUE);
            rowTop += rowHeight + rowGap;
        }

        MoveWindow(g_presetRequireExternalCheck, formLeft, rowTop, s(210), rowHeight, TRUE);
        rowTop += rowHeight + rowGap;
        MoveWindow(g_presetShowTrayCheck, formLeft, rowTop, s(160), rowHeight, TRUE);
        MoveWindow(g_presetShowButtonCheck, formLeft + s(170), rowTop, s(190), rowHeight, TRUE);
        rowTop += rowHeight + rowGap;
        MoveWindow(GetDlgItem(g_presetEditorWindow, IDC_PRESET_SAVE), formLeft, rowTop + s(2), s(120), buttonHeight, TRUE);
        MoveWindow(GetDlgItem(g_presetEditorWindow, IDC_PRESET_APPLY_FORM), formLeft + s(132), rowTop + s(2), s(120), buttonHeight, TRUE);
        MoveWindow(g_presetStatus, formLeft, rowTop + buttonHeight + s(12), previewLeft - formLeft - s(8), s(90), TRUE);
    }

    LRESULT CALLBACK PresetPreviewProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
    {
        if (message == WM_PAINT)
        {
            PAINTSTRUCT paint{};
            HDC dc = BeginPaint(window, &paint);
            RECT client{};
            GetClientRect(window, &client);
            FillRect(dc, &client, reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1));

            HPEN borderPen = CreatePen(PS_SOLID, 1, RGB(208, 208, 208));
            HGDIOBJ oldPen = SelectObject(dc, borderPen);
            HBRUSH frameBrush = CreateSolidBrush(RGB(245, 247, 250));
            FillRect(dc, &client, frameBrush);
            Rectangle(dc, client.left, client.top, client.right, client.bottom);

            if (g_selectedPresetIndex >= 0 && g_selectedPresetIndex < static_cast<int>(g_presets.size()))
            {
                const auto preset = BuildPreviewPresetFromEditor();
                RectInt frame{12, 12, (client.right - client.left) - 24, (client.bottom - client.top) - 24};
                HBRUSH agBrush = CreateSolidBrush(RGB(255, 191, 87));
                HBRUSH chromeBrush = CreateSolidBrush(RGB(95, 122, 201));
                SetBkMode(dc, TRANSPARENT);

                auto drawTile = [&](const RectInt& rect, HBRUSH brush, const std::wstring& label)
                {
                    RECT tile{rect.left, rect.top, rect.left + rect.width, rect.top + rect.height};
                    FillRect(dc, &tile, brush);
                    Rectangle(dc, tile.left, tile.top, tile.right, tile.bottom);
                    SetTextColor(dc, RGB(40, 40, 40));
                    DrawTextW(dc, label.c_str(), -1, &tile, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                };

                if (preset.kind == PresetKind::Grid)
                {
                    auto rects = BuildGrid(frame, preset.columns, preset.rows, preset.columns * preset.rows);
                    for (size_t index = 0; index < rects.size(); ++index)
                    {
                        drawTile(rects[index], agBrush, L"Tile");
                    }
                }
                else
                {
                    // Derive short labels from the user's own patterns, stripping wildcards.
                    auto patternLabel = [](const std::wstring& a, const std::wstring& b) -> std::wstring {
                        for (const auto& src : {a, b}) {
                            std::wstring t;
                            for (auto ch : src) { if (ch != L'*' && ch != L'?') t += ch; }
                            if (!t.empty()) return t.size() > 6 ? t.substr(0, 6) : t;
                        }
                        return L"Left";
                    };
                    const std::wstring leftLabel  = patternLabel(preset.leftProcessPattern,  preset.leftTitlePattern);
                    const std::wstring rightLabel = patternLabel(preset.rightProcessPattern, preset.rightTitlePattern);

                    const int chromeWidth = static_cast<int>(frame.width * (preset.chromeWidthPercent / 100.0) + 0.5);
                    RectInt leftColumn{frame.left, frame.top, frame.width - chromeWidth, frame.height};
                    RectInt rightColumn{frame.left + leftColumn.width, frame.top, frame.width - leftColumn.width, frame.height};
                    auto leftRects = BuildGrid(leftColumn, preset.columns, preset.rows, preset.columns * preset.rows);
                    for (size_t index = 0; index < leftRects.size(); ++index)
                    {
                        drawTile(leftRects[index], agBrush, leftLabel);
                    }
                    drawTile(InsetRect(rightColumn, g_windowGapPx), chromeBrush, rightLabel);
                }

                DeleteObject(agBrush);
                DeleteObject(chromeBrush);
            }

            SelectObject(dc, oldPen);
            DeleteObject(frameBrush);
            DeleteObject(borderPen);
            EndPaint(window, &paint);
            return 0;
        }

        return DefWindowProcW(window, message, wParam, lParam);
    }

    void CreateAboutWindow(HINSTANCE instance)
    {
        const UINT dpi = GetWindowDpiSafe(g_mainWindow);
        auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
        g_aboutWindow = CreateWindowExW(
            WS_EX_TOOLWINDOW,
            L"WindowLayouterAboutClass",
            L"About WindowLayouter",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            s(560),
            s(420),
            g_mainWindow,
            nullptr,
            instance,
            nullptr);

        ShowWindow(g_aboutWindow, SW_HIDE);
    }

    void ResizeAboutControls()
    {
        if (g_aboutWindow == nullptr)
        {
            return;
        }

        RECT client{};
        GetClientRect(g_aboutWindow, &client);
        const UINT dpi = GetWindowDpiSafe(g_aboutWindow);
        auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };

        const int margin = s(16);
        const int buttonWidth = s(110);
        const int buttonHeight = s(32);
        const int titleHeight = s(28);

        MoveWindow(GetDlgItem(g_aboutWindow, IDC_ABOUT_TITLE), margin, margin - s(2), client.right - margin * 2 - buttonWidth - s(12), titleHeight, TRUE);
        MoveWindow(g_aboutText, margin, margin + s(36), client.right - margin * 2, client.bottom - margin * 3 - buttonHeight - s(36), TRUE);
        MoveWindow(GetDlgItem(g_aboutWindow, IDC_ABOUT_CLOSE), client.right - margin - buttonWidth, client.bottom - margin - buttonHeight, buttonWidth, buttonHeight, TRUE);
    }

    LRESULT CALLBACK AboutProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
    {
        switch (message)
        {
        case WM_CREATE:
        {
            const UINT dpi = GetWindowDpiSafe(window);
            auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
            CreateWindowW(L"STATIC", L"About", WS_VISIBLE | WS_CHILD, s(16), s(14), s(240), s(28), window, reinterpret_cast<HMENU>(IDC_ABOUT_TITLE), nullptr, nullptr);
            g_aboutText = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_READONLY, s(16), s(52), s(512), s(286), window, reinterpret_cast<HMENU>(IDC_ABOUT_TEXT), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"Close", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, s(418), s(348), s(110), s(32), window, reinterpret_cast<HMENU>(IDC_ABOUT_CLOSE), nullptr, nullptr);
            ApplyAuxUiFonts(window);
            ApplyAboutLocalization();
            ResizeAboutControls();
            return 0;
        }
        case WM_SIZE:
            ResizeAboutControls();
            return 0;
        case WM_DPICHANGED:
        {
            CreateAuxUiFonts(HIWORD(wParam) == 0 ? 96 : HIWORD(wParam));
            ApplyAuxUiFonts(window);
            auto* suggested = reinterpret_cast<const RECT*>(lParam);
            SetWindowPos(window, nullptr, suggested->left, suggested->top,
                suggested->right - suggested->left, suggested->bottom - suggested->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
            ResizeAboutControls();
            return 0;
        }
        case WM_COMMAND:
            if (LOWORD(wParam) == IDC_ABOUT_CLOSE)
            {
                ShowWindow(window, SW_HIDE);
                return 0;
            }
            break;
        case WM_CLOSE:
            ShowWindow(window, SW_HIDE);
            return 0;
        default:
            break;
        }

        return DefWindowProcW(window, message, wParam, lParam);
    }

    void ShowActiveTabPage()
    {
        const auto tabIndex = TabCtrl_GetCurSel(g_tab);
        ShowWindow(g_monitorsPage, tabIndex == 0 ? SW_SHOW : SW_HIDE);
        ShowWindow(g_presetsPage, tabIndex == 1 ? SW_SHOW : SW_HIDE);
        ShowWindow(g_rulesPage, tabIndex == 2 ? SW_SHOW : SW_HIDE);
    }

    BOOL CALLBACK EnumMonitorsProc(HMONITOR monitor, HDC, LPRECT, LPARAM data)
    {
        auto* monitors = reinterpret_cast<std::vector<MonitorInfo>*>(data);
        MONITORINFOEXW info{};
        info.cbSize = sizeof(info);
        if (!GetMonitorInfoW(monitor, &info))
        {
            return TRUE;
        }

        UINT dpiX = 96;
        UINT dpiY = 96;
        GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY);

        monitors->push_back({
            info.szDevice,
            (info.dwFlags & MONITORINFOF_PRIMARY) != 0,
            {info.rcMonitor.left, info.rcMonitor.top, info.rcMonitor.right - info.rcMonitor.left, info.rcMonitor.bottom - info.rcMonitor.top},
            {info.rcWork.left, info.rcWork.top, info.rcWork.right - info.rcWork.left, info.rcWork.bottom - info.rcWork.top},
            dpiX
        });

        return TRUE;
    }

    std::vector<MonitorInfo> GetMonitors()
    {
        std::vector<MonitorInfo> monitors;
        EnumDisplayMonitors(nullptr, nullptr, EnumMonitorsProc, reinterpret_cast<LPARAM>(&monitors));
        return monitors;
    }

    const MonitorInfo* GetPrimaryMonitor()
    {
        auto it = std::find_if(g_monitors.begin(), g_monitors.end(), [](const MonitorInfo& item) { return item.isPrimary; });
        if (it != g_monitors.end())
        {
            return &*it;
        }

        return g_monitors.empty() ? nullptr : &g_monitors.front();
    }

    WindowInfo BuildWindowInfo(HWND window, int zOrder)
    {
        RECT rect{};
        GetWindowRect(window, &rect);

        DWORD processId = 0;
        GetWindowThreadProcessId(window, &processId);

        std::wstring processName = L"unknown";
        HANDLE process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
        if (process != nullptr)
        {
            wchar_t path[MAX_PATH]{};
            DWORD length = MAX_PATH;
            if (QueryFullProcessImageNameW(process, 0, path, &length))
            {
                std::wstring fullPath(path, length);
                const auto slash = fullPath.find_last_of(L"\\/");
                processName = slash == std::wstring::npos ? fullPath : fullPath.substr(slash + 1);
            }
            CloseHandle(process);
        }

        HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONEAREST);
        MONITORINFOEXW info{};
        info.cbSize = sizeof(info);
        GetMonitorInfoW(monitor, &info);

        UINT dpiX = 96;
        UINT dpiY = 96;
        GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY);

        return {
            window,
            zOrder,
            ReadWindowText(window),
            processName,
            ReadClassName(window),
            info.szDevice,
            {rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top},
            dpiX,
            IsIconic(window) != FALSE
        };
    }

    // =========================================================================
    // Window enumeration and list view management
    // =========================================================================

    // EnumWindows callback: filters out tool windows, cloaked windows, owner-
    // less backgrounds, and tiny (<60×40) windows, then appends qualifying ones
    // to g_windows.
    BOOL CALLBACK EnumWindowsProc(HWND window, LPARAM)
    {
        if (!IsWindowVisible(window) || GetWindow(window, GW_OWNER) != nullptr)
        {
            return TRUE;
        }

        LONG_PTR exStyle = GetWindowLongPtrW(window, GWL_EXSTYLE);
        if ((exStyle & WS_EX_TOOLWINDOW) != 0 || (exStyle & WS_EX_NOACTIVATE) != 0)
        {
            return TRUE;
        }

        const auto title = ReadWindowText(window);
        if (title.empty() || IsWindowCloaked(window))
        {
            return TRUE;
        }

        RECT rect{};
        if (!GetWindowRect(window, &rect))
        {
            return TRUE;
        }

        if ((rect.right - rect.left) < 60 || (rect.bottom - rect.top) < 40)
        {
            return TRUE;
        }

        const auto className = ReadClassName(window);
        if (className == L"Progman" || className == L"WorkerW" || className == L"Shell_TrayWnd")
        {
            return TRUE;
        }

        g_windows.push_back(BuildWindowInfo(window, static_cast<int>(g_windows.size())));
        return TRUE;
    }

    std::wstring ToHandleHex(HWND handle)
    {
        std::wstringstream stream;
        stream << L"0x" << std::hex << reinterpret_cast<UINT_PTR>(handle);
        return stream.str();
    }

    std::wstring GetControlText(HWND control)
    {
        const int length = GetWindowTextLengthW(control);
        std::wstring text(static_cast<size_t>(length), L'\0');
        if (length > 0)
        {
            GetWindowTextW(control, text.data(), length + 1);
        }
        return text;
    }

    bool MatchesWindowListFilters(const WindowInfo& window)
    {
        const auto titlePattern = Trim(GetControlText(g_filterTitleEdit));
        const auto processPattern = Trim(GetControlText(g_filterProcessEdit));
        return MatchesOptionalPattern(titlePattern, window.title) && MatchesOptionalPattern(processPattern, window.processName);
    }

    int CompareWindowInfo(const WindowInfo& left, const WindowInfo& right, int column)
    {
        auto compareText = [](const std::wstring& a, const std::wstring& b)
        {
            return _wcsicmp(a.c_str(), b.c_str());
        };

        switch (column)
        {
        case 0:
            return compareText(left.title, right.title);
        case 1:
            return compareText(left.processName, right.processName);
        case 2:
            return compareText(left.className, right.className);
        case 3:
            return compareText(left.monitorName, right.monitorName);
        case 4:
        {
            const auto leftValue = reinterpret_cast<UINT_PTR>(left.handle);
            const auto rightValue = reinterpret_cast<UINT_PTR>(right.handle);
            return leftValue < rightValue ? -1 : (leftValue > rightValue ? 1 : 0);
        }
        default:
            return left.zOrder - right.zOrder;
        }
    }

    void RefreshWindowList()
    {
        std::vector<HWND> selectedHandles;
        const auto itemCount = ListView_GetItemCount(g_listView);
        for (int index = 0; index < itemCount; ++index)
        {
            if ((ListView_GetItemState(g_listView, index, LVIS_SELECTED) & LVIS_SELECTED) != 0 &&
                index < static_cast<int>(g_visibleWindows.size()))
            {
                selectedHandles.push_back(g_visibleWindows[static_cast<size_t>(index)]->handle);
            }
        }

        g_windows.clear();
        g_visibleWindows.clear();
        g_monitors = GetMonitors();
        ListView_DeleteAllItems(g_listView);
        EnumWindows(EnumWindowsProc, 0);

        std::sort(g_windows.begin(), g_windows.end(), [](const WindowInfo& left, const WindowInfo& right)
        {
            const int comparison = CompareWindowInfo(left, right, g_sortColumn);
            return g_sortAscending ? comparison < 0 : comparison > 0;
        });

        int listIndex = 0;
        for (int index = 0; index < static_cast<int>(g_windows.size()); ++index)
        {
            const auto& window = g_windows[static_cast<size_t>(index)];
            if (!MatchesWindowListFilters(window))
            {
                continue;
            }

            g_visibleWindows.push_back(&g_windows[static_cast<size_t>(index)]);

            LVITEMW item{};
            item.mask = LVIF_TEXT;
            item.iItem = listIndex;
            item.iSubItem = 0;
            item.pszText = const_cast<LPWSTR>(window.title.c_str());
            ListView_InsertItem(g_listView, &item);
            ListView_SetItemText(g_listView, listIndex, 1, const_cast<LPWSTR>(window.processName.c_str()));
            ListView_SetItemText(g_listView, listIndex, 2, const_cast<LPWSTR>(window.className.c_str()));
            ListView_SetItemText(g_listView, listIndex, 3, const_cast<LPWSTR>(window.monitorName.c_str()));

            const auto handleHex = ToHandleHex(window.handle);
            ListView_SetItemText(g_listView, listIndex, 4, const_cast<LPWSTR>(handleHex.c_str()));

            const auto wasSelected = std::find(selectedHandles.begin(), selectedHandles.end(), window.handle) != selectedHandles.end();
            if (wasSelected)
            {
                ListView_SetItemState(g_listView, listIndex, LVIS_SELECTED, LVIS_SELECTED);
            }
            ++listIndex;
        }

        UpdateTabContent();

        std::wstringstream status;
        status << Lang(L"status.loaded_prefix", L"Loaded ")
               << ListView_GetItemCount(g_listView)
               << Lang(L"status.loaded_middle", L" windows, ")
               << g_monitors.size()
               << Lang(L"status.loaded_suffix", L" monitors.");
        UpdateStatus(status.str());
    }

    void RefilterWindowList()
    {
        std::vector<HWND> selectedHandles;
        const auto itemCount = ListView_GetItemCount(g_listView);
        for (int index = 0; index < itemCount; ++index)
        {
            if ((ListView_GetItemState(g_listView, index, LVIS_SELECTED) & LVIS_SELECTED) != 0 &&
                index < static_cast<int>(g_visibleWindows.size()))
            {
                selectedHandles.push_back(g_visibleWindows[static_cast<size_t>(index)]->handle);
            }
        }

        g_visibleWindows.clear();
        ListView_DeleteAllItems(g_listView);

        int listIndex = 0;
        for (size_t index = 0; index < g_windows.size(); ++index)
        {
            const auto& window = g_windows[index];
            if (!MatchesWindowListFilters(window))
            {
                continue;
            }

            g_visibleWindows.push_back(&g_windows[index]);

            LVITEMW item{};
            item.mask = LVIF_TEXT;
            item.iItem = listIndex;
            item.iSubItem = 0;
            item.pszText = const_cast<LPWSTR>(window.title.c_str());
            ListView_InsertItem(g_listView, &item);
            ListView_SetItemText(g_listView, listIndex, 1, const_cast<LPWSTR>(window.processName.c_str()));
            ListView_SetItemText(g_listView, listIndex, 2, const_cast<LPWSTR>(window.className.c_str()));
            ListView_SetItemText(g_listView, listIndex, 3, const_cast<LPWSTR>(window.monitorName.c_str()));

            const auto handleHex = ToHandleHex(window.handle);
            ListView_SetItemText(g_listView, listIndex, 4, const_cast<LPWSTR>(handleHex.c_str()));

            if (std::find(selectedHandles.begin(), selectedHandles.end(), window.handle) != selectedHandles.end())
            {
                ListView_SetItemState(g_listView, listIndex, LVIS_SELECTED, LVIS_SELECTED);
            }
            ++listIndex;
        }

        std::wstringstream status;
        status << ListView_GetItemCount(g_listView) << L" / " << g_windows.size()
               << Lang(L"status.loaded_suffix", L" monitors.");
        UpdateStatus(status.str());
    }

    std::vector<WindowInfo*> GetSelectedWindows()
    {
        std::vector<WindowInfo*> windows;
        const auto itemCount = ListView_GetItemCount(g_listView);
        for (int index = 0; index < itemCount && index < static_cast<int>(g_visibleWindows.size()); ++index)
        {
            if ((ListView_GetItemState(g_listView, index, LVIS_SELECTED) & LVIS_SELECTED) != 0)
            {
                windows.push_back(g_visibleWindows[static_cast<size_t>(index)]);
            }
        }

        return windows;
    }

    std::vector<WindowInfo*> GetVisibleWindows()
    {
        return g_visibleWindows;
    }

    std::vector<WindowInfo*> GetAllDiscoveredWindows()
    {
        std::vector<WindowInfo*> windows;
        windows.reserve(g_windows.size());
        for (auto& window : g_windows)
        {
            windows.push_back(&window);
        }
        return windows;
    }

    SelectionScope GetEffectiveSelectionScope(const PresetDefinition& preset)
    {
        switch (g_appSelectionOverride)
        {
        case AppSelectionOverride::SelectedRows:
            return SelectionScope::SelectedRows;
        case AppSelectionOverride::AllVisibleRows:
            return SelectionScope::AllVisibleRows;
        case AppSelectionOverride::AllDiscoveredWindows:
            return SelectionScope::AllDiscoveredWindows;
        case AppSelectionOverride::FollowPreset:
        default:
            return preset.selectionScope;
        }
    }

    bool MatchesOptionalPattern(const std::wstring& pattern, const std::wstring& value)
    {
        if (Trim(pattern).empty())
        {
            return true;
        }

        return WildcardMatch(pattern, value);
    }

    std::vector<WindowInfo*> FilterWindows(const std::vector<WindowInfo*>& windows, const std::wstring& titlePattern, const std::wstring& processPattern)
    {
        std::vector<WindowInfo*> filtered;
        for (auto* window : windows)
        {
            if (MatchesOptionalPattern(titlePattern, window->title) && MatchesOptionalPattern(processPattern, window->processName))
            {
                filtered.push_back(window);
            }
        }

        std::sort(filtered.begin(), filtered.end(), [](const WindowInfo* left, const WindowInfo* right)
        {
            return left->zOrder < right->zOrder;
        });

        return filtered;
    }

    void ClearSelection()
    {
        const auto itemCount = ListView_GetItemCount(g_listView);
        for (int index = 0; index < itemCount; ++index)
        {
            ListView_SetItemState(g_listView, index, 0, LVIS_SELECTED);
        }
        UpdateStatus(Lang(L"status.selection_cleared", L"Selection cleared."));
    }

    RectInt InsetRect(RectInt rect, int padding)
    {
        rect.left += padding;
        rect.top += padding;
        rect.width = std::max(50, rect.width - padding * 2);
        rect.height = std::max(40, rect.height - padding * 2);
        return rect;
    }

    std::vector<RectInt> BuildGrid(RectInt frame, int columns, int rows, int count)
    {
        std::vector<RectInt> rects;
        if (columns <= 0 || rows <= 0 || count <= 0)
        {
            return rects;
        }

        const auto cellWidth = frame.width / columns;
        const auto cellHeight = frame.height / rows;
        for (int index = 0; index < count && index < columns * rows; ++index)
        {
            const auto column = index % columns;
            const auto row = index / columns;
            RectInt cell
            {
                frame.left + column * cellWidth,
                frame.top + row * cellHeight,
                column == columns - 1 ? frame.width - column * cellWidth : cellWidth,
                row == rows - 1 ? frame.height - row * cellHeight : cellHeight
            };
            rects.push_back(InsetRect(cell, g_windowGapPx));
        }

        return rects;
    }

    bool MoveWindowToRect(WindowInfo& window, const RectInt& rect)
    {
        if (window.minimized)
        {
            ShowWindow(window.handle, SW_RESTORE);
        }

        return SetWindowPos(
            window.handle,
            nullptr,
            rect.left,
            rect.top,
            rect.width,
            rect.height,
            SWP_NOACTIVATE | SWP_NOZORDER) != FALSE;
    }

    void ApplyPresetByDefinition(const PresetDefinition& preset)
    {
        RefreshWindowList();

        const auto* monitor = GetPresetTargetMonitor(preset);
        if (monitor == nullptr)
        {
            UpdateStatus(BuildPresetInactiveReason(preset));
            return;
        }

        const auto selectionScope = GetEffectiveSelectionScope(preset);
        std::vector<WindowInfo*> candidateWindows;
        switch (selectionScope)
        {
        case SelectionScope::SelectedRows:
            candidateWindows = GetSelectedWindows();
            break;
        case SelectionScope::AllVisibleRows:
            candidateWindows = GetVisibleWindows();
            break;
        case SelectionScope::AllDiscoveredWindows:
            candidateWindows = GetAllDiscoveredWindows();
            break;
        }
        if (candidateWindows.empty())
        {
            switch (selectionScope)
            {
            case SelectionScope::SelectedRows:
                UpdateStatus(Lang(L"status.select_windows_before_apply", L"Select windows on the list before applying a preset."));
                break;
            case SelectionScope::AllVisibleRows:
                UpdateStatus(Lang(L"status.no_visible_windows", L"No visible windows match current filters for this preset."));
                break;
            case SelectionScope::AllDiscoveredWindows:
                UpdateStatus(Lang(L"status.no_discovered_windows", L"No discovered windows are available for this preset."));
                break;
            }
            return;
        }

        int applied = 0;
        int failed = 0;
        int skipped = 0;

        if (preset.kind == PresetKind::Grid)
        {
            auto rects = BuildGrid(monitor->workArea, preset.columns, preset.rows, static_cast<int>(candidateWindows.size()));
            for (size_t index = 0; index < rects.size(); ++index)
            {
                if (MoveWindowToRect(*candidateWindows[index], rects[index]))
                {
                    ++applied;
                }
                else
                {
                    ++failed;
                }
            }
            skipped = static_cast<int>(candidateWindows.size()) - static_cast<int>(rects.size());
        }
        else
        {
            auto antigravity = FilterWindows(candidateWindows, preset.leftTitlePattern, preset.leftProcessPattern);
            auto chrome = FilterWindows(candidateWindows, preset.rightTitlePattern, preset.rightProcessPattern);

            const auto chromeWidth = static_cast<int>(monitor->workArea.width * (preset.chromeWidthPercent / 100.0) + 0.5);
            RectInt leftColumn
            {
                monitor->workArea.left,
                monitor->workArea.top,
                monitor->workArea.width - chromeWidth,
                monitor->workArea.height
            };
            RectInt rightColumn
            {
                monitor->workArea.left + leftColumn.width,
                monitor->workArea.top,
                monitor->workArea.width - leftColumn.width,
                monitor->workArea.height
            };

            auto agRects = BuildGrid(leftColumn, preset.columns, preset.rows, static_cast<int>(antigravity.size()));
            auto chromeRect = InsetRect(rightColumn, g_windowGapPx);

            for (size_t index = 0; index < agRects.size(); ++index)
            {
                if (MoveWindowToRect(*antigravity[index], agRects[index]))
                {
                    ++applied;
                }
                else
                {
                    ++failed;
                }
            }
            skipped += static_cast<int>(antigravity.size()) - static_cast<int>(agRects.size());

            for (auto* window : chrome)
            {
                if (MoveWindowToRect(*window, chromeRect))
                {
                    ++applied;
                }
                else
                {
                    ++failed;
                }
            }
        }

        std::wstringstream status;
        status << preset.name
               << Lang(L"status.applied_on", L" on ")
               << monitor->deviceName
               << Lang(L"status.applied_counts_prefix", L": applied ")
               << applied
               << Lang(L"status.applied_failed", L", failed ")
               << failed;
        if (skipped > 0)
        {
            status << Lang(L"status.applied_skipped", L", skipped ") << skipped;
        }
        UpdateStatus(status.str());
    }

    const PresetDefinition* FindPresetByCommand(int commandId)
    {
        auto it = std::find_if(g_presets.begin(), g_presets.end(), [&](const PresetDefinition& preset)
        {
            return preset.commandId == commandId;
        });

        return it == g_presets.end() ? nullptr : &*it;
    }

    const PresetDefinition* FindPresetByHotkeyId(int hotkeyId)
    {
        auto it = std::find_if(g_presets.begin(), g_presets.end(), [&](const PresetDefinition& preset)
        {
            return preset.hotkeyId == hotkeyId;
        });

        return it == g_presets.end() ? nullptr : &*it;
    }

    // =========================================================================
    // Window visibility and tray icon management
    // =========================================================================

    void ShowMainWindow()
    {
        ShowWindow(g_mainWindow, SW_SHOW);
        ShowWindow(g_mainWindow, SW_RESTORE);

        RECT rect{};
        GetWindowRect(g_mainWindow, &rect);
        HMONITOR monitor = MonitorFromRect(&rect, MONITOR_DEFAULTTONULL);
        if (monitor == nullptr)
        {
            HMONITOR primary = MonitorFromPoint({0, 0}, MONITOR_DEFAULTTOPRIMARY);
            MONITORINFO mi{};
            mi.cbSize = sizeof(mi);
            GetMonitorInfoW(primary, &mi);
            int w = rect.right - rect.left;
            int h = rect.bottom - rect.top;
            SetWindowPos(g_mainWindow, nullptr,
                mi.rcWork.left + (mi.rcWork.right - mi.rcWork.left - w) / 2,
                mi.rcWork.top + (mi.rcWork.bottom - mi.rcWork.top - h) / 2,
                0, 0, SWP_NOSIZE | SWP_NOZORDER);
        }

        // SetForegroundWindow() silently fails when the calling process does not
        // own the foreground.  Briefly making the window TOPMOST forces Windows
        // to grant it foreground focus regardless of process ownership.
        SetWindowPos(g_mainWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        SetWindowPos(g_mainWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        SetForegroundWindow(g_mainWindow);
    }

    void HideMainWindow()
    {
        ShowWindow(g_mainWindow, SW_HIDE);
    }

    void AddTrayIcon()
    {
        ZeroMemory(&g_trayIcon, sizeof(g_trayIcon));
        g_trayIcon.cbSize = sizeof(g_trayIcon);
        g_trayIcon.hWnd = g_mainWindow;
        g_trayIcon.uID = TRAY_ICON_UID;
        g_trayIcon.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
        g_trayIcon.uCallbackMessage = WMAPP_TRAYICON;
        g_trayIcon.hIcon = LoadAppIcon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
        const auto trayTip = Lang(L"app.title", L"WindowLayouter.Native");
        lstrcpynW(g_trayIcon.szTip, trayTip.c_str(), static_cast<int>(std::size(g_trayIcon.szTip)));

        g_trayAdded = Shell_NotifyIconW(NIM_ADD, &g_trayIcon) != FALSE;
    }

    void RemoveTrayIcon()
    {
        if (g_trayAdded)
        {
            Shell_NotifyIconW(NIM_DELETE, &g_trayIcon);
            g_trayAdded = false;
        }
    }

    void ShowTrayMenu()
    {
        HMENU menu = CreatePopupMenu();
        const auto trayShow = Lang(L"tray.show", L"Show");
        const auto trayEditor = Lang(L"tray.preset_editor", L"Preset Editor");
        const auto trayAbout = Lang(L"tray.about", L"About");
        const auto trayHide = Lang(L"tray.hide", L"Hide to Tray");
        const auto trayExit = Lang(L"tray.exit", L"Exit");
        AppendMenuW(menu, MF_STRING, ID_TRAY_SHOW, trayShow.c_str());
        AppendMenuW(menu, MF_STRING, ID_TRAY_PRESET_EDITOR, trayEditor.c_str());
        AppendMenuW(menu, MF_STRING, ID_TRAY_ABOUT, trayAbout.c_str());
        AppendMenuW(menu, MF_STRING, ID_TRAY_HIDE, trayHide.c_str());
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        for (const auto& preset : g_presets)
        {
            if (!preset.showInTray)
            {
                continue;
            }
            std::wstring label = preset.trayLabel + L"  [" + FormatHotkeyText(preset.hotkey) + L"]";
            const UINT flags = MF_STRING | (GetPresetTargetMonitor(preset) == nullptr ? MF_GRAYED : 0);
            AppendMenuW(menu, flags, preset.commandId, label.c_str());
        }
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT, trayExit.c_str());

        POINT cursor{};
        GetCursorPos(&cursor);
        SetForegroundWindow(g_mainWindow);
        TrackPopupMenu(menu, TPM_RIGHTBUTTON, cursor.x, cursor.y, 0, g_mainWindow, nullptr);
        DestroyMenu(menu);
    }

    void UnregisterPresetHotkeys()
    {
        for (const auto& preset : g_presets)
        {
            UnregisterHotKey(g_mainWindow, preset.hotkeyId);
        }
    }

    void RegisterPresetHotkeys()
    {
        UnregisterPresetHotkeys();
        for (auto& preset : g_presets)
        {
            if (!preset.hotkey.valid)
            {
                continue;
            }

            if (!RegisterHotKey(g_mainWindow, preset.hotkeyId, preset.hotkey.modifiers, preset.hotkey.virtualKey))
            {
                preset.hotkey.valid = false;
            }
        }
    }

    IniData LoadSettings()
    {
        EnsureConfigDirectory();
        g_iniPath = GetAppDataDirectory() + L"\\settings.ini";
        LoadLanguageResources(DetectDefaultLanguageCode());
        g_presets = BuildDefaultPresets();
        EnsureIniFileExists();

        const auto iniData = ReadIniData();
        LoadLanguageResources(ReadIniString(iniData, L"App", L"language", g_languageCode));
        g_windowGapPx = std::max(0, ReadIniInt(iniData, L"Window", L"window_gap_px", 0));
        g_appSelectionOverride = ParseAppSelectionOverride(ReadIniString(iniData, L"App", L"selection_scope_override", L"Follow preset"));

        for (auto& preset : g_presets)
        {
            const auto savedHotkey = ReadIniString(iniData, L"Hotkeys", preset.iniKey, preset.hotkey.text);
            auto parsed = ParseHotkeyText(savedHotkey);
            if (parsed.valid)
            {
                preset.hotkey = parsed;
            }
        }

        const auto presetCount = std::min<int>(ReadIniInt(iniData, L"Presets", L"count", static_cast<int>(g_presets.size())), static_cast<int>(g_presets.size()));
        for (int index = 0; index < presetCount; ++index)
        {
            auto& preset = g_presets[static_cast<size_t>(index)];
            const auto section = L"Preset" + std::to_wstring(index);

            preset.name = ReadIniString(iniData, section, L"name", preset.name);
            preset.trayLabel = ReadIniString(iniData, section, L"tray_label", preset.trayLabel);
            preset.toolbarLabel = ReadIniString(iniData, section, L"toolbar_label", preset.toolbarLabel);
            preset.showInTray = ReadIniInt(iniData, section, L"show_in_tray", preset.showInTray ? 1 : 0) != 0;
            preset.showAsToolbarButton = ReadIniInt(iniData, section, L"show_as_toolbar_button", preset.showAsToolbarButton ? 1 : 0) != 0;
            preset.selectionScope = ParseSelectionScope(ReadIniString(iniData, section, L"selection_scope", FormatSelectionScope(preset.selectionScope)));
            preset.requireExternalMonitor = ReadIniInt(iniData, section, L"require_external_monitor", preset.requireExternalMonitor ? 1 : 0) != 0;
            preset.requiredResolutionClass = ParseResolutionClass(ReadIniString(iniData, section, L"required_resolution_class", FormatResolutionClass(preset.requiredResolutionClass)));
            preset.requiredMonitorDevice = ReadIniString(iniData, section, L"required_monitor_device", preset.requiredMonitorDevice);
            const auto kindText = ToUpper(ReadIniString(iniData, section, L"kind", preset.kind == PresetKind::Grid ? L"Grid" : L"Split"));
            preset.kind = kindText == L"SPLIT" ? PresetKind::Split : PresetKind::Grid;
            preset.columns = std::max(1, ReadIniInt(iniData, section, L"columns", preset.columns));
            preset.rows = std::max(1, ReadIniInt(iniData, section, L"rows", preset.rows));
            preset.chromeWidthPercent = std::clamp(ReadIniInt(iniData, section, L"chrome_width_percent", preset.chromeWidthPercent), 20, 80);
            preset.leftTitlePattern = ReadIniString(iniData, section, L"left_title", preset.leftTitlePattern);
            preset.leftProcessPattern = ReadIniString(iniData, section, L"left_process", preset.leftProcessPattern);
            preset.rightTitlePattern = ReadIniString(iniData, section, L"right_title", preset.rightTitlePattern);
            preset.rightProcessPattern = ReadIniString(iniData, section, L"right_process", preset.rightProcessPattern);
        }

        return iniData;
    }

    void SaveSettings()
    {
        SaveIniData();
    }

    void ApplySavedWindowPlacement(const IniData& iniData)
    {
        const int left = ReadIniInt(iniData, L"Window", L"left", CW_USEDEFAULT);
        const int top = ReadIniInt(iniData, L"Window", L"top", CW_USEDEFAULT);
        const int width = ReadIniInt(iniData, L"Window", L"width", 1760);
        const int height = ReadIniInt(iniData, L"Window", L"height", 1020);

        SetWindowPos(g_mainWindow, nullptr, left, top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);

        const auto tabIndex = ReadIniInt(iniData, L"Window", L"tab", 0);
        TabCtrl_SetCurSel(g_tab, tabIndex);
        ShowActiveTabPage();
    }

    void ResizeControls(HWND window)
    {
        RECT client{};
        GetClientRect(window, &client);
        auto s = [](int v) { return MulDiv(v, static_cast<int>(g_currentDpi), 96); };

        const int margin = s(12);
        const int buttonHeight = s(32);
        const int buttonGap = s(8);
        const int statusHeight = s(24);
        const int rightWidth = s(430);
        const int toolbarTop = margin;
        const int widths[] = {s(80), s(120), s(160), s(170), s(170), s(180), s(180), s(140), s(150)};
        const int buttonIds[] = {
            IDC_REFRESH,
            IDC_CLEAR_SELECTION,
            IDC_APPLY_SELECTED_PRESET,
            IDC_PRESET_GRID_2X2,
            IDC_PRESET_SPLIT_2X2,
            IDC_PRESET_SPLIT_2X3,
            IDC_PRESET_SPLIT_3X2,
            IDC_PRESET_SPLIT_3X3,
            IDC_OPEN_PRESET_EDITOR
        };

        int currentLeft = margin;
        for (int index = 0; index < 9; ++index)
        {
            auto* button = GetDlgItem(window, buttonIds[index]);
            if (button == nullptr)
            {
                continue;
            }
            if (buttonIds[index] >= IDC_PRESET_GRID_2X2 && buttonIds[index] <= IDC_PRESET_SPLIT_3X3 && !IsWindowVisible(button))
            {
                continue;
            }
            MoveWindow(button, currentLeft, toolbarTop, widths[index], buttonHeight, TRUE);
            currentLeft += widths[index] + buttonGap;
        }

        const int filterTop = toolbarTop + buttonHeight + s(10);
        MoveWindow(GetDlgItem(window, IDC_FILTER_TITLE_LABEL), margin, filterTop + s(5), s(64), s(20), TRUE);
        MoveWindow(g_filterTitleEdit, margin + s(68), filterTop, s(250), s(26), TRUE);
        MoveWindow(GetDlgItem(window, IDC_FILTER_PROCESS_LABEL), margin + s(330), filterTop + s(5), s(74), s(20), TRUE);
        MoveWindow(g_filterProcessEdit, margin + s(408), filterTop, s(220), s(26), TRUE);

        const int contentTop = filterTop + s(34);
        const int contentHeight = client.bottom - contentTop - statusHeight - margin * 2;
        MoveWindow(g_listView, margin, contentTop, client.right - rightWidth - margin * 3, contentHeight, TRUE);
        MoveWindow(g_tab, client.right - rightWidth - margin, contentTop, rightWidth, contentHeight, TRUE);
        MoveWindow(g_status, margin, client.bottom - statusHeight - s(6), client.right - margin * 2, statusHeight, TRUE);

        RECT tabRect{};
        GetClientRect(g_tab, &tabRect);
        TabCtrl_AdjustRect(g_tab, FALSE, &tabRect);
        const int pageWidth = tabRect.right - tabRect.left;
        const int pageHeight = tabRect.bottom - tabRect.top;
        MoveWindow(g_monitorsPage, tabRect.left, tabRect.top, pageWidth, pageHeight, TRUE);
        MoveWindow(g_presetsPage, tabRect.left, tabRect.top, pageWidth, pageHeight, TRUE);
        MoveWindow(g_rulesPage, tabRect.left, tabRect.top, pageWidth, pageHeight, TRUE);
        MoveWindow(g_monitorsEdit, 0, 0, pageWidth, pageHeight, TRUE);
        MoveWindow(g_rulesEdit, 0, 0, pageWidth, pageHeight, TRUE);

        auto* openEditorButton = GetDlgItem(g_presetsPage, IDC_OPEN_PRESET_EDITOR);
        MoveWindow(openEditorButton, 0, 0, s(180), s(32), TRUE);
        MoveWindow(GetDlgItem(g_presetsPage, IDC_PRESETS_DASHBOARD_TITLE), 0, s(2), s(220), s(24), TRUE);
        MoveWindow(g_presetsSummaryEdit, 0, s(40), pageWidth, pageHeight - s(40), TRUE);

        MoveWindow(GetDlgItem(g_rulesPage, IDC_APP_SCOPE_OVERRIDE_LABEL), 0, 0, s(180), s(20), TRUE);
        MoveWindow(g_appScopeOverrideCombo, s(186), 0, s(220), s(240), TRUE);
        MoveWindow(GetDlgItem(g_rulesPage, IDC_LANGUAGE_LABEL), s(420), 0, s(90), s(20), TRUE);
        MoveWindow(g_languageCombo, s(514), 0, s(160), s(240), TRUE);
        MoveWindow(GetDlgItem(g_rulesPage, IDC_OPEN_ABOUT), s(688), 0, s(110), s(28), TRUE);
        MoveWindow(g_rulesEdit, 0, s(36), pageWidth, pageHeight - s(36), TRUE);
        ShowActiveTabPage();
    }

    void CreateTabs(HWND parent)
    {
        g_tab = CreateWindowW(
            WC_TABCONTROLW,
            L"",
            WS_VISIBLE | WS_CHILD | WS_CLIPSIBLINGS,
            0,
            0,
            0,
            0,
            parent,
            reinterpret_cast<HMENU>(IDC_TAB),
            nullptr,
            nullptr);

        TCITEMW tabItem{};
        tabItem.mask = TCIF_TEXT;

        tabItem.pszText = const_cast<LPWSTR>(L"Monitors");
        TabCtrl_InsertItem(g_tab, 0, &tabItem);
        tabItem.pszText = const_cast<LPWSTR>(L"Preset Editor");
        TabCtrl_InsertItem(g_tab, 1, &tabItem);
        tabItem.pszText = const_cast<LPWSTR>(L"Rules");
        TabCtrl_InsertItem(g_tab, 2, &tabItem);

        g_monitorsPage = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, g_tab, reinterpret_cast<HMENU>(IDC_TAB_MONITORS), nullptr, nullptr);
        g_presetsPage = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, g_tab, reinterpret_cast<HMENU>(IDC_TAB_PRESETS), nullptr, nullptr);
        g_rulesPage = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, g_tab, reinterpret_cast<HMENU>(IDC_TAB_RULES), nullptr, nullptr);

        const auto editStyle = WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_READONLY;
        g_monitorsEdit = CreateWindowW(L"EDIT", L"", editStyle, 0, 0, 0, 0, g_monitorsPage, nullptr, nullptr, nullptr);
        g_rulesEdit = CreateWindowW(L"EDIT", L"", editStyle, 0, 0, 0, 0, g_rulesPage, nullptr, nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Preset Dashboard", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, g_presetsPage, reinterpret_cast<HMENU>(IDC_PRESETS_DASHBOARD_TITLE), nullptr, nullptr);
        g_presetsSummaryEdit = CreateWindowW(L"EDIT", L"", editStyle, 0, 0, 0, 0, g_presetsPage, reinterpret_cast<HMENU>(IDC_PRESET_SUMMARY), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Open Preset Editor", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 0, 0, 0, 0, g_presetsPage, reinterpret_cast<HMENU>(IDC_OPEN_PRESET_EDITOR), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"App Selection Override", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, g_rulesPage, reinterpret_cast<HMENU>(IDC_APP_SCOPE_OVERRIDE_LABEL), nullptr, nullptr);
        g_appScopeOverrideCombo = CreateWindowW(WC_COMBOBOXW, L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST, 0, 0, 0, 0, g_rulesPage, reinterpret_cast<HMENU>(IDC_APP_SCOPE_OVERRIDE), nullptr, nullptr);
        SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Follow preset"));
        SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Selected rows"));
        SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"All visible rows"));
        SendMessageW(g_appScopeOverrideCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"All discovered windows"));
        CreateWindowW(L"STATIC", L"Language", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, g_rulesPage, reinterpret_cast<HMENU>(IDC_LANGUAGE_LABEL), nullptr, nullptr);
        g_languageCombo = CreateWindowW(WC_COMBOBOXW, L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST, 0, 0, 0, 0, g_rulesPage, reinterpret_cast<HMENU>(IDC_LANGUAGE), nullptr, nullptr);
        SendMessageW(g_languageCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"English"));
        SendMessageW(g_languageCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Polski"));
        CreateWindowW(L"BUTTON", L"About", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, g_rulesPage, reinterpret_cast<HMENU>(IDC_OPEN_ABOUT), nullptr, nullptr);
    }

    void CreatePresetEditorWindow(HINSTANCE instance)
    {
        g_presetEditorWindow = CreateWindowExW(
            WS_EX_APPWINDOW,
            L"WindowLayouterPresetEditorClass",
            L"WindowLayouter Preset Editor",
            WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            1180,
            860,
            nullptr,
            nullptr,
            instance,
            nullptr);

        ShowWindow(g_presetEditorWindow, SW_HIDE);
    }

    LRESULT CALLBACK PresetEditorProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
    {
        switch (message)
        {
        case WM_CREATE:
        {
            // Set the global handle early so ApplyPresetEditorLocalization() can find it.
            // CreateWindowExW sends WM_CREATE before returning, so g_presetEditorWindow
            // would still be null without this explicit assignment.
            g_presetEditorWindow = window;
            CreateWindowW(L"STATIC", L"Preset Editor", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_EDITOR_TITLE), nullptr, nullptr);
            g_presetList = CreateWindowW(L"LISTBOX", L"", WS_VISIBLE | WS_CHILD | WS_VSCROLL | LBS_NOTIFY | WS_BORDER, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_EDITOR_LIST), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Preset Name", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_NAME_LABEL), nullptr, nullptr);
            g_presetNameEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_NAME), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Hotkey", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_HOTKEY_LABEL), nullptr, nullptr);
            g_presetHotkeyEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_HOTKEY), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Tray Label", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_TRAY_LABEL_TEXT), nullptr, nullptr);
            g_presetTrayLabelEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_TRAY_LABEL), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Toolbar Label", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_BUTTON_LABEL_TEXT), nullptr, nullptr);
            g_presetButtonLabelEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_BUTTON_LABEL), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Selection Scope", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SCOPE_LABEL), nullptr, nullptr);
            g_presetScopeCombo = CreateWindowW(WC_COMBOBOXW, L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | CBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SCOPE), nullptr, nullptr);
            SendMessageW(g_presetScopeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Selected rows"));
            SendMessageW(g_presetScopeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"All visible rows"));
            SendMessageW(g_presetScopeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"All discovered windows"));
            CreateWindowW(L"STATIC", L"Kind", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_KIND_LABEL), nullptr, nullptr);
            g_presetKindCombo = CreateWindowW(WC_COMBOBOXW, L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | CBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_KIND), nullptr, nullptr);
            SendMessageW(g_presetKindCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Grid"));
            SendMessageW(g_presetKindCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Split"));
            CreateWindowW(L"STATIC", L"Resolution Class", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_RESOLUTION_LABEL), nullptr, nullptr);
            g_presetResolutionCombo = CreateWindowW(WC_COMBOBOXW, L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | CBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_RESOLUTION_CLASS), nullptr, nullptr);
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Any"));
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"2K"));
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"2.5K"));
            SendMessageW(g_presetResolutionCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"4K"));
            CreateWindowW(L"STATIC", L"Monitor Device", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_DEVICE_LABEL), nullptr, nullptr);
            g_presetDeviceCombo = CreateWindowW(WC_COMBOBOXW, L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | CBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_DEVICE), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Grid Columns", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_COLUMNS_LABEL), nullptr, nullptr);
            g_presetColumnsEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_COLUMNS), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Grid Rows", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_ROWS_LABEL), nullptr, nullptr);
            g_presetRowsEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_ROWS), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Right Stack Width %", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_WIDTH_LABEL), nullptr, nullptr);
            g_presetWidthEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_WIDTH), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Left Title", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_LEFT_TITLE_LABEL), nullptr, nullptr);
            g_presetLeftTitleEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_LEFT_TITLE), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Left Process", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_LEFT_PROCESS_LABEL), nullptr, nullptr);
            g_presetLeftProcessEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_LEFT_PROCESS), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Right Title", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_RIGHT_TITLE_LABEL), nullptr, nullptr);
            g_presetRightTitleEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_RIGHT_TITLE), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Right Process", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_RIGHT_PROCESS_LABEL), nullptr, nullptr);
            g_presetRightProcessEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_RIGHT_PROCESS), nullptr, nullptr);
            g_presetRequireExternalCheck = CreateWindowW(L"BUTTON", L"Require external monitor", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_REQUIRE_EXTERNAL), nullptr, nullptr);
            g_presetShowTrayCheck = CreateWindowW(L"BUTTON", L"Show in tray", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SHOW_TRAY), nullptr, nullptr);
            g_presetShowButtonCheck = CreateWindowW(L"BUTTON", L"Show as toolbar button", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SHOW_BUTTON), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"Save Preset", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SAVE), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"Apply Preset", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_APPLY_FORM), nullptr, nullptr);
            g_presetStatus = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_STATUS), nullptr, nullptr);
            g_presetPreview = CreateWindowW(L"WindowLayouterPresetPreviewClass", L"", WS_VISIBLE | WS_CHILD | WS_BORDER, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_PREVIEW), nullptr, nullptr);
            for (HWND edit : {g_presetNameEdit, g_presetHotkeyEdit, g_presetTrayLabelEdit, g_presetButtonLabelEdit,
                              g_presetLeftTitleEdit, g_presetLeftProcessEdit, g_presetRightTitleEdit, g_presetRightProcessEdit})
            {
                SendMessageW(edit, EM_SETLIMITTEXT, 256, 0);
            }
            for (HWND edit : {g_presetColumnsEdit, g_presetRowsEdit, g_presetWidthEdit})
            {
                SendMessageW(edit, EM_SETLIMITTEXT, 8, 0);
            }
            ApplyAuxUiFonts(window);
            ApplyPresetEditorLocalization();
            ResizePresetEditorControls();
            LoadPresetIntoEditor(g_selectedPresetIndex);
            return 0;
        }
        case WM_SIZE:
            ResizePresetEditorControls();
            return 0;
        case WM_DPICHANGED:
        {
            CreateAuxUiFonts(HIWORD(wParam) == 0 ? 96 : HIWORD(wParam));
            ApplyAuxUiFonts(window);
            auto* suggested = reinterpret_cast<const RECT*>(lParam);
            SetWindowPos(window, nullptr, suggested->left, suggested->top,
                suggested->right - suggested->left, suggested->bottom - suggested->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
            ResizePresetEditorControls();
            return 0;
        }
        case WM_COMMAND:
            switch (LOWORD(wParam))
            {
            case IDC_PRESET_SAVE:
                SavePresetFromEditor();
                return 0;
            case IDC_PRESET_APPLY_FORM:
                if (SavePresetFromEditor())
                {
                    ApplyPresetByDefinition(g_presets[static_cast<size_t>(g_selectedPresetIndex)]);
                }
                return 0;
            default:
                if (LOWORD(wParam) == IDC_PRESET_EDITOR_LIST && HIWORD(wParam) == LBN_SELCHANGE)
                {
                    LoadPresetIntoEditor(static_cast<int>(SendMessageW(g_presetList, LB_GETCURSEL, 0, 0)));
                    return 0;
                }
                if (HIWORD(wParam) == EN_CHANGE || HIWORD(wParam) == CBN_SELCHANGE)
                {
                    if (LOWORD(wParam) == IDC_PRESET_KIND && HIWORD(wParam) == CBN_SELCHANGE)
                    {
                        RefreshWidthFieldState();
                    }
                    RefreshPresetEditorPreviewState();
                }
                break;
            }
            break;
        case WM_CLOSE:
            ShowWindow(window, SW_HIDE);
            return 0;
        default:
            break;
        }

        return DefWindowProcW(window, message, wParam, lParam);
    }

    // =========================================================================
    // Main window message loop (WindowProc)
    // Handles: WM_CREATE (control creation), WM_SIZE (layout), WM_DPICHANGED,
    // WM_HOTKEY (global shortcuts), WMAPP_TRAYICON, WM_COMMAND, WM_CLOSE,
    // WM_DESTROY (cleanup).
    // =========================================================================
    LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
    {
        if (message == g_taskbarCreatedMessage)
        {
            AddTrayIcon();
            return 0;
        }

        switch (message)
        {
        case WM_CREATE:
        {
            g_mainWindow = window;
            g_currentDpi = GetDpiForWindow(window);
            if (g_currentDpi == 0) g_currentDpi = 96;

            INITCOMMONCONTROLSEX controls{};
            controls.dwSize = sizeof(controls);
            controls.dwICC = ICC_LISTVIEW_CLASSES | ICC_TAB_CLASSES;
            InitCommonControlsEx(&controls);

            CreateWindowW(L"BUTTON", L"Refresh", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_REFRESH), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"Clear Selection", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_CLEAR_SELECTION), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"Apply Selected Preset", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_APPLY_SELECTED_PRESET), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"4K > 2x2", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_GRID_2X2), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"4K > AG 2x2 + Chrome", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SPLIT_2X2), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"4K > AG 2x3 + Chrome", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SPLIT_2X3), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"4K > AG 3x2 + Chrome 1/3", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SPLIT_3X2), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"4K > AG 3x3 + Chrome 1/3", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_PRESET_SPLIT_3X3), nullptr, nullptr);
            CreateWindowW(L"BUTTON", L"Preset Editor", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_OPEN_PRESET_EDITOR), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Title", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_FILTER_TITLE_LABEL), nullptr, nullptr);
            g_filterTitleEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_FILTER_TITLE), nullptr, nullptr);
            CreateWindowW(L"STATIC", L"Process", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_FILTER_PROCESS_LABEL), nullptr, nullptr);
            g_filterProcessEdit = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_FILTER_PROCESS), nullptr, nullptr);
            SendMessageW(g_filterTitleEdit, EM_SETLIMITTEXT, 256, 0);
            SendMessageW(g_filterProcessEdit, EM_SETLIMITTEXT, 256, 0);

            g_listView = CreateWindowW(
                WC_LISTVIEWW,
                L"",
                WS_VISIBLE | WS_CHILD | LVS_REPORT | LVS_SHOWSELALWAYS,
                0,
                0,
                0,
                0,
                window,
                reinterpret_cast<HMENU>(IDC_WINDOW_LIST),
                nullptr,
                nullptr);

            ListView_SetExtendedListViewStyle(
                g_listView,
                LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_GRIDLINES);

            const std::wstring columns[] = {L"Title", L"Process", L"Class", L"Monitor", L"HWND"};
            const int widths[] = {360, 130, 140, 120, 140};
            for (int index = 0; index < 5; ++index)
            {
                LVCOLUMNW column{};
                column.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
                column.pszText = const_cast<LPWSTR>(columns[index].c_str());
                column.cx = widths[index];
                column.iSubItem = index;
                ListView_InsertColumn(g_listView, index, &column);
            }

            CreateTabs(window);
            g_status = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_STATUS), nullptr, nullptr);
            ApplyUiFonts(window);

            const auto iniData = LoadSettings();
            SendMessageW(g_languageCombo, CB_SETCURSEL, _wcsicmp(g_languageCode.c_str(), L"pl") == 0 ? 1 : 0, 0);
            SendMessageW(g_appScopeOverrideCombo, CB_SETCURSEL, static_cast<WPARAM>(g_appSelectionOverride), 0);
            CreatePresetEditorWindow(reinterpret_cast<LPCREATESTRUCTW>(lParam)->hInstance);
            CreateAboutWindow(reinterpret_cast<LPCREATESTRUCTW>(lParam)->hInstance);
            ApplyMainWindowLocalization();
            RefreshToolbarPresetButtons();
            AddTrayIcon();
            RegisterPresetHotkeys();
            SetTimer(window, TIMER_LANG_RELOAD, 1500, nullptr);
            RefreshWindowList();
            ApplySavedWindowPlacement(iniData);
            RefreshPresetListBox();
            LoadPresetIntoEditor(0);
            SaveSettings();
            return 0;
        }
        case WM_SIZE:
            ResizeControls(window);
            return 0;
        case WM_TIMER:
            if (wParam == TIMER_LANG_RELOAD)
            {
                ReloadLanguageIfChanged();
                return 0;
            }
            break;
        case WM_DPICHANGED:
        {
            g_currentDpi = HIWORD(wParam);
            if (g_currentDpi == 0) g_currentDpi = 96;
            CreateUiFonts();
            ApplyUiFonts(window);
            auto* suggested = reinterpret_cast<const RECT*>(lParam);
            SetWindowPos(window, nullptr, suggested->left, suggested->top,
                suggested->right - suggested->left, suggested->bottom - suggested->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
            return 0;
        }
        case WM_NOTIFY:
        {
            auto* notify = reinterpret_cast<LPNMHDR>(lParam);
            if (notify->idFrom == IDC_TAB && notify->code == TCN_SELCHANGE)
            {
                ShowActiveTabPage();
                return 0;
            }
            if (notify->idFrom == IDC_WINDOW_LIST && notify->code == LVN_COLUMNCLICK)
            {
                auto* columnClick = reinterpret_cast<LPNMLISTVIEW>(lParam);
                if (g_sortColumn == columnClick->iSubItem)
                {
                    g_sortAscending = !g_sortAscending;
                }
                else
                {
                    g_sortColumn = columnClick->iSubItem;
                    g_sortAscending = true;
                }
                RefreshWindowList();
                return 0;
            }
            break;
        }
        case WM_HOTKEY:
            if (const auto* preset = FindPresetByHotkeyId(static_cast<int>(wParam)); preset != nullptr)
            {
                ApplyPresetByDefinition(*preset);
                return 0;
            }
            break;
        case WMAPP_TRAYICON:
            switch (LOWORD(lParam))
            {
            case WM_LBUTTONDBLCLK:
                ShowMainWindow();
                return 0;
            case WM_RBUTTONUP:
            case WM_CONTEXTMENU:
                ShowTrayMenu();
                return 0;
            default:
                break;
            }
            break;
        case WM_COMMAND:
            switch (LOWORD(wParam))
            {
            case IDC_REFRESH:
                RefreshWindowList();
                return 0;
            case IDC_CLEAR_SELECTION:
                ClearSelection();
                return 0;
            case IDC_OPEN_PRESET_EDITOR:
                ShowPresetEditorWindow();
                return 0;
            case IDC_OPEN_ABOUT:
                ShowAboutWindow();
                return 0;
            case IDC_APPLY_SELECTED_PRESET:
                if (g_selectedPresetIndex >= 0 && g_selectedPresetIndex < static_cast<int>(g_presets.size()))
                {
                    ApplyPresetByDefinition(g_presets[static_cast<size_t>(g_selectedPresetIndex)]);
                }
                return 0;
            case IDC_APP_SCOPE_OVERRIDE:
                g_appSelectionOverride = static_cast<AppSelectionOverride>(SendMessageW(g_appScopeOverrideCombo, CB_GETCURSEL, 0, 0));
                SaveSettings();
                UpdateTabContent();
                return 0;
            case IDC_LANGUAGE:
                if (HIWORD(wParam) == CBN_SELCHANGE)
                {
                    g_languageCode = SendMessageW(g_languageCombo, CB_GETCURSEL, 0, 0) == 1 ? L"pl" : L"en";
                    LoadLanguageResources(g_languageCode);
                    ApplyMainWindowLocalization();
                    ApplyPresetEditorLocalization();
                    SaveSettings();
                    RefreshToolbarPresetButtons();
                    UpdateTabContent();
                }
                return 0;
            case ID_TRAY_SHOW:
                ShowMainWindow();
                return 0;
            case ID_TRAY_PRESET_EDITOR:
                ShowPresetEditorWindow();
                return 0;
            case ID_TRAY_ABOUT:
                ShowAboutWindow();
                return 0;
            case ID_TRAY_HIDE:
                HideMainWindow();
                return 0;
            case ID_TRAY_EXIT:
                g_allowExit = true;
                DestroyWindow(g_mainWindow);
                return 0;
            default:
                if (LOWORD(wParam) == IDC_PRESET_EDITOR_LIST && HIWORD(wParam) == LBN_SELCHANGE)
                {
                    LoadPresetIntoEditor(static_cast<int>(SendMessageW(g_presetList, LB_GETCURSEL, 0, 0)));
                    return 0;
                }
                if ((LOWORD(wParam) == IDC_FILTER_TITLE || LOWORD(wParam) == IDC_FILTER_PROCESS) && HIWORD(wParam) == EN_CHANGE)
                {
                    RefilterWindowList();
                    return 0;
                }
                if (const auto* preset = FindPresetByCommand(LOWORD(wParam)); preset != nullptr)
                {
                    ApplyPresetByDefinition(*preset);
                    return 0;
                }
                break;
            }
            break;
        case WM_CLOSE:
            if (!g_allowExit)
            {
                HideMainWindow();
                UpdateStatus(Lang(L"status.window_hidden_to_tray", L"Window hidden to tray."));
                return 0;
            }
            break;
        case WM_DESTROY:
            KillTimer(window, TIMER_LANG_RELOAD);
            SaveSettings();
            UnregisterPresetHotkeys();
            RemoveTrayIcon();
            if (g_trayIcon.hIcon != nullptr) { DestroyIcon(g_trayIcon.hIcon); g_trayIcon.hIcon = nullptr; }
            if (g_uiFont != nullptr) { DeleteObject(g_uiFont); g_uiFont = nullptr; }
            if (g_uiFontBold != nullptr) { DeleteObject(g_uiFontBold); g_uiFontBold = nullptr; }
            if (g_uiFontTitle != nullptr) { DeleteObject(g_uiFontTitle); g_uiFontTitle = nullptr; }
            if (g_auxUiFont != nullptr) { DeleteObject(g_auxUiFont); g_auxUiFont = nullptr; }
            if (g_auxUiFontBold != nullptr) { DeleteObject(g_auxUiFontBold); g_auxUiFontBold = nullptr; }
            if (g_auxUiFontTitle != nullptr) { DeleteObject(g_auxUiFontTitle); g_auxUiFontTitle = nullptr; }
            PostQuitMessage(0);
            return 0;
        default:
            break;
        }

        return DefWindowProcW(window, message, wParam, lParam);
    }
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int showCommand)
{
    // Must be called before any window is created; cannot be changed afterwards.
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

    // Single-instance guard: if another copy is running, bring it to the front
    // and exit.  The GUID suffix makes the mutex name globally unique.
    HANDLE instanceMutex = CreateMutexW(nullptr, TRUE, L"WindowLayouterNative_SingleInstance_8E729B17");
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        HWND existing = FindWindowW(L"WindowLayouterNativeClass", nullptr);
        if (existing != nullptr)
        {
            ShowWindow(existing, SW_SHOW);
            SetForegroundWindow(existing);
        }
        if (instanceMutex != nullptr)
        {
            CloseHandle(instanceMutex);
        }
        return 0;
    }
    g_taskbarCreatedMessage = RegisterWindowMessageW(L"TaskbarCreated");

    WNDCLASSEXW windowClass{};
    windowClass.cbSize = sizeof(windowClass);
    windowClass.lpfnWndProc = WindowProc;
    windowClass.hInstance = instance;
    windowClass.lpszClassName = L"WindowLayouterNativeClass";
    windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    windowClass.hIcon = LoadAppIcon(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
    windowClass.hIconSm = LoadAppIcon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));

    if (!RegisterClassExW(&windowClass))
    {
        return 1;
    }

    WNDCLASSEXW presetEditorClass{};
    presetEditorClass.cbSize = sizeof(presetEditorClass);
    presetEditorClass.lpfnWndProc = PresetEditorProc;
    presetEditorClass.hInstance = instance;
    presetEditorClass.lpszClassName = L"WindowLayouterPresetEditorClass";
    presetEditorClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    presetEditorClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    presetEditorClass.hIcon = LoadAppIcon(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
    presetEditorClass.hIconSm = LoadAppIcon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
    if (!RegisterClassExW(&presetEditorClass))
    {
        return 1;
    }

    WNDCLASSEXW previewClass{};
    previewClass.cbSize = sizeof(previewClass);
    previewClass.lpfnWndProc = PresetPreviewProc;
    previewClass.hInstance = instance;
    previewClass.lpszClassName = L"WindowLayouterPresetPreviewClass";
    previewClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    previewClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    if (!RegisterClassExW(&previewClass))
    {
        return 1;
    }

    WNDCLASSEXW aboutClass{};
    aboutClass.cbSize = sizeof(aboutClass);
    aboutClass.lpfnWndProc = AboutProc;
    aboutClass.hInstance = instance;
    aboutClass.lpszClassName = L"WindowLayouterAboutClass";
    aboutClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    aboutClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    aboutClass.hIcon = LoadAppIcon(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
    aboutClass.hIconSm = LoadAppIcon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
    if (!RegisterClassExW(&aboutClass))
    {
        return 1;
    }

    g_mainWindow = CreateWindowExW(
        0,
        windowClass.lpszClassName,
        L"WindowLayouter.Native",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        1840,
        1040,
        nullptr,
        nullptr,
        instance,
        nullptr);

    ShowWindow(g_mainWindow, showCommand);

    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0))
    {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }

    if (instanceMutex != nullptr)
    {
        CloseHandle(instanceMutex);
    }

    return static_cast<int>(message.wParam);
}
