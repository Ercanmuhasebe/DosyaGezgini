using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Xml;
using System.Xml.Xsl;
using System.Text;
using System.IO.Compression;
using System.Collections.Specialized;
using System.Collections.Concurrent;
using System.Net;
using System.Security.Cryptography;
using System.Reflection;
using System.Web.Script.Serialization;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

internal sealed class FileExplorerStarter : Form
{
    [DllImport("user32.dll", SetLastError = true)] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint key);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "FindFirstFileW")]
    private static extern IntPtr FindFirstFileNative(string fileName, out Win32FindData findData);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool FindClose(IntPtr findHandle);
    private const int ShowAppHotKeyId = 0x4D31;
    private const uint ModWin = 0x0008, ModAlt = 0x0001;
    private const uint IoReparseTagFilePlaceholder = 0x80000015;
    private const uint IoReparseTagCloudBase = 0x9000001A;
    private const uint IoReparseTagCloudMask = 0xFFFF0FFF;
    private const uint CfPlaceholderStatePlaceholder = 0x00000001;
    private const uint CfPlaceholderStatePartial = 0x00000010;
    private const uint CfPlaceholderStatePartiallyOnDisk = 0x00000020;
    private const uint FileAttributeOffline = 0x00001000;
    private const uint FileAttributeRecallOnOpen = 0x00040000;
    private const uint FileAttributeRecallOnDataAccess = 0x00400000;
    private bool appHotKeyRegistered;
    private readonly TextBox address = new TextBox();
    private readonly FlowLayoutPanel breadcrumbBar = new FlowLayoutPanel();
    private readonly ListBox folders = new ListBox();
    private readonly TreeView folderTree = new TreeView();
    private FileSystemWatcher currentFolderWatcher;
    private System.Windows.Forms.Timer directoryRefreshTimer;
    private System.Windows.Forms.Timer oneDriveStatusTimer;
    private BackgroundWorker oneDriveStatusWorker;
    private BackgroundWorker pdfOcrWorker;
    private bool refreshingDirectoryFromWatcher;
    private readonly List<string> pendingCutPaths = new List<string>();
    private string pendingFolderClipboardPath;
    private bool pendingFolderClipboardIsCut;
    private MemoryStream clipboardDropEffectStream;
    private readonly Dictionary<string, string> oneDriveStatusCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> oneDriveStatusLastChecked = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
    private string[] oneDriveRoots;
    private object shellApplication;
    private object windowsShareManager;
    private EventInfo windowsShareDataRequestedEvent;
    private Delegate windowsShareEventHandler;
    private Delegate windowsShareAddEventHandler;
    private Delegate windowsShareRemoveEventHandler;
    private string windowsSharePendingPath;
    private object windowsSharePendingStorageItem;
    private bool windowsShareHandlerRegistered;
    private readonly ListView files = new ListView();
    private OutlookAsyncDropTarget filesAsyncDropTarget;
    private readonly TextBox folderSearch = new TextBox();
    private readonly TextBox fileSearch = new TextBox();
    private readonly Label status = new Label();
    private readonly Label folderHeading = new Label();
    private readonly RadioButton classicFolderViewOption = new RadioButton();
    private readonly RadioButton treeFolderViewOption = new RadioButton();
    private readonly Label fileHeading = new Label();
    private readonly Button go = new Button();
    private readonly Button back = new Button();
    private readonly Button forward = new Button();
    private readonly ContextMenuStrip backHistoryMenu = new ContextMenuStrip();
    private readonly ContextMenuStrip forwardHistoryMenu = new ContextMenuStrip();
    private readonly Button up = new Button();
    private readonly Button refresh = new Button();
    private readonly Button quickLocations = new Button();
    private readonly ContextMenuStrip quickLocationsMenu = new ContextMenuStrip();
    private readonly Button addShortcut = new Button();
    private readonly Button togglePreview = new Button();
    private readonly Button openExplorer = new Button();
    private readonly ComboBox themeSelector = new ComboBox();
    private readonly TextBox edefterTemplateFolderBox = new TextBox();
    private readonly ComboBox historyRetentionCountSelector = new ComboBox();
    private readonly ComboBox historyRetentionUnitSelector = new ComboBox();
    private readonly CheckBox preserveHistoryRetentionOption = new CheckBox();
    private readonly CheckBox keepOnTopOption = new CheckBox();
    private readonly CheckBox rememberFolderSortOption = new CheckBox();
    private readonly CheckBox startMaximizedOption = new CheckBox();
    private readonly CheckBox minimizeAfterOpeningFileOption = new CheckBox();
    private readonly RadioButton closeToTrayOption = new RadioButton();
    private readonly RadioButton normalCloseOption = new RadioButton();
    private readonly ToolTip actionTooltips = new ToolTip();
    private readonly TabControl tabs = new TabControl();
    private TabPage transferPage;
    private TabPage syncPage;
    private Panel syncPageContent;
    private readonly string syncProfilesFile = PortableDataPath("sync-profiles.xml");
    private readonly List<SyncWorkspaceProfile> syncProfiles = new List<SyncWorkspaceProfile>();
    private readonly ComboBox syncProfileSelector = new ComboBox();
    private readonly TextBox syncLeftPathBox = new TextBox();
    private readonly TextBox syncRightPathBox = new TextBox();
    private readonly ComboBox syncComparisonSelector = new ComboBox();
    private readonly CheckBox syncSubfoldersOption = new CheckBox();
    private readonly CheckBox syncShowIdenticalOption = new CheckBox();
    private readonly ListView syncComparisonList = new ListView();
    private readonly Label syncStatus = new Label();
    private readonly Button syncCompareButton = new Button();
    private readonly Button syncCancelButton = new Button();
    private readonly Button syncApplyButton = new Button();
    private readonly Button syncFullScreenButton = new Button();
    private readonly Button syncLookupButton = new Button();
    private readonly Button syncExportCsvButton = new Button();
    private BackgroundWorker syncComparisonWorker;
    private BackgroundWorker syncLookupWorker;
    private List<SyncComparisonRow> syncCurrentRows = new List<SyncComparisonRow>();
    private bool bindingSyncProfile;
    private bool syncProfilePendingSave;
    private SyncWorkspaceProfile activeSyncProfile;
    private bool syncFullScreen;
    private Panel mainHeader;
    private FormBorderStyle syncPreviousBorderStyle;
    private FormWindowState syncPreviousWindowState;
    private Rectangle syncPreviousBounds;
    private TransferPaneState transferLeftPane;
    private TransferPaneState transferRightPane;
    private readonly List<string> transferClipboardPaths = new List<string>();
    private bool transferClipboardIsCut;
    private readonly RadioButton transferCopyMode = new RadioButton();
    private readonly RadioButton transferMoveMode = new RadioButton();
    private readonly Label transferStatus = new Label();
    private readonly string transferPanePathsFile = PortableDataPath("transfer-pane-paths.txt");
    private readonly string deepSearchResultsFile = PortableDataPath("search-results.xml");
    private readonly TabControl folderTabs = new TabControl();
    private readonly Panel explorerWorkspace = new Panel();
    private readonly Button newFolderTabButton = new Button();
    private readonly Stack<TabPage> recentlyClosedFolderTabs = new Stack<TabPage>();
    private readonly string folderTabsFile = PortableDataPath("folder-tabs.xml");
    private TabPage activeFolderTab;
    private bool suppressFolderTabEvents;
    private TabPage contextFolderTab;
    private TabPage workspacesPage;
    private TabPage developmentPage;
    private TabPage mediaPage;
    private readonly TextBox mediaApiKeyBox = new TextBox();
    private readonly RichTextBox mediaResearchOutput = new RichTextBox();
    private readonly Label mediaResearchStatus = new Label();
    private readonly Button mediaResearchButton = new Button();
    private readonly ListView mediaNewsList = new ListView();
    private readonly Button mediaNewsRefreshButton = new Button();
    private readonly Label mediaNewsStatus = new Label();
    private readonly ComboBox mediaNewsCategoryCombo = new ComboBox();
    private readonly RichTextBox mediaNewsDetails = new RichTextBox();
    private readonly Label mediaNewsDetailTitle = new Label();
    private readonly Button mediaNewsOpenButton = new Button();
    private List<MediaHeadline> mediaNewsCache = new List<MediaHeadline>();
    private readonly string mediaNewsCacheFile = PortableDataPath("media-news-cache.json");
    private readonly ListView mediaWorldNewsList = new ListView();
    private readonly ComboBox mediaWorldCategoryCombo = new ComboBox();
    private readonly Button mediaWorldRefreshButton = new Button();
    private readonly Button mediaWorldTranslateButton = new Button();
    private readonly Label mediaWorldStatus = new Label();
    private readonly RichTextBox mediaWorldDetails = new RichTextBox();
    private readonly Label mediaWorldDetailTitle = new Label();
    private readonly Button mediaWorldOpenButton = new Button();
    private List<MediaHeadline> mediaWorldNewsCache = new List<MediaHeadline>();
    private List<MediaHeadline> mediaWorldNewsTranslationBaseline = new List<MediaHeadline>();
    private bool mediaNewsInitialRefreshStarted;
    private bool mediaWorldNewsInitialRefreshStarted;
    private BackgroundWorker mediaResearchWorker;
    private BackgroundWorker mediaNewsWorker;
    private BackgroundWorker mediaWorldNewsWorker;
    private BackgroundWorker mediaWorldTranslateWorker;
    private DateTime mediaNewsLastRefresh = DateTime.MinValue;
    private DateTime mediaWorldNewsLastRefresh = DateTime.MinValue;
    private int mediaWorldDisplayedProgressCount = -1;
    private string mediaApiKey = String.Empty;
    private readonly string mediaApiKeyFile = ComputerSettingPath("openai-api-key.bin");
    private readonly string mediaResearchFile = PortableDataPath("media-research.txt");
    private readonly string argosRuntimeRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DosyaGezgini", "ArgosRuntime");
    private string discoveredSystemPythonPath;
    private bool systemPythonDiscoveryComplete;
    private const string ArgosPythonInstallerUrl = "https://www.python.org/ftp/python/3.13.17/python-3.13.17-amd64.exe";
    private const string ArgosPythonInstallerSha256 = "5da012101ba2e17113a9a1cd901e9ce1441086f2b7ecb5de06ec94163899a190";
    private const string ArgosModelUrl = "https://argos-net.com/v1/translate-en_tr-1_5.argosmodel";
    private const string ArgosModelFileName = "translate-en_tr-1_5.argosmodel";
    private const string ArgosPackageVersion = "1.11.0";
    private readonly string sharedDevelopmentFile = PortableDataPath("development-board.xml");
    private readonly List<DevelopmentTask> developmentTasks = new List<DevelopmentTask>();
    private readonly HashSet<Guid> deletedDevelopmentTaskIds = new HashSet<Guid>();
    private readonly List<DevelopmentVersion> developmentVersions = new List<DevelopmentVersion>();
    private readonly ListView developmentTaskList = new ListView();
    private readonly ListView developmentVersionList = new ListView();
    private readonly TextBox developmentTaskDetails = new TextBox();
    private readonly TextBox developmentVersionDetails = new TextBox();
    private readonly TextBox developmentTaskBox = new TextBox();
    private readonly ComboBox developmentChangeScaleSelector = new ComboBox();
    private readonly Button addDevelopmentTaskButton = new Button();
    private readonly Button completeDevelopmentTaskButton = new Button();
    private readonly Button editDevelopmentTaskButton = new Button();
    private readonly Button deleteDevelopmentTaskButton = new Button();
    private Font activeDevelopmentTaskFont;
    private readonly List<WorkProfile> workProfiles = new List<WorkProfile>();
    private readonly List<Tuple<Guid, string, DateTime>> portableShortcutRefs = new List<Tuple<Guid, string, DateTime>>();
    private readonly ListView workProfileList = new ListView();
    private Font workProfileActiveFont;
    private readonly ListView workProfileFolders = new ListView();
    private readonly Label workProfileTitle = new Label();
    private readonly Label workProfileRoot = new Label();
    private readonly TextBox workProfileNotes = new TextBox();
    private readonly Label workProfileLastOpened = new Label();
    private readonly Button newWorkProfileButton = new Button();
    private readonly Button editWorkProfileButton = new Button();
    private readonly Button removeWorkProfileButton = new Button();
    private readonly Button openWorkProfileButton = new Button();
    private readonly Button addWorkProfileFolderButton = new Button();
    private readonly Button removeWorkProfileFolderButton = new Button();
    private readonly Button addWorkProfileFileButton = new Button();
    private WorkProfile activeWorkProfile;
    private bool refreshingWorkProfileList;
    private readonly ListView shortcutList = new ListView();
    private readonly ListView historyList = new ListView();
    private ColumnHeader oneDriveStatusColumn;
    private readonly ComboBox fileFilter = new ComboBox();
    private readonly ComboBox shortcutFilter = new ComboBox();
    private readonly ComboBox historyFilter = new ComboBox();
    private readonly TextBox shortcutSearch = new TextBox();
    private readonly TextBox historySearch = new TextBox();
    private readonly TextBox deepSearchBox = new TextBox();
    private readonly Button deepSearchButton = new Button();
    private readonly Button cancelDeepSearchButton = new Button();
    private readonly Button clearDeepSearchButton = new Button();
    private readonly ListView deepSearchResults = new ListView();
    private Font deepSearchBoldFont;
    private readonly ProgressBar searchProgress = new ProgressBar();
    private readonly CheckBox searchSubfoldersOption = new CheckBox();
    private SplitContainer explorerLayout;
    private SplitContainer folderFilesSplit;
    private Panel previewPane;
    private Panel previewContent;
    private FilePreviewForm embeddedPreview;
    private bool previewPaneOpen;
    private int folderWidthBeforePreview;
    private readonly List<Entry> shortcuts = new List<Entry>();
    private readonly List<HistoryEntry> history = new List<HistoryEntry>();
    private readonly List<string> navigationHistory = new List<string>();
    private readonly List<MoveRecord> lastMoveRecords = new List<MoveRecord>();
    private readonly List<string> lastRecycleBinDeletedPaths = new List<string>();
    private string lastMoveOriginView;
    private readonly List<ListViewItem> fileRows = new List<ListViewItem>();
    private readonly Dictionary<string, string> fileTypeCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> fileIconKeyCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly ImageList fileIconImageList = new ImageList();
    private string folderIconKey;
    private readonly Dictionary<string, FileSortPreference> folderSortPreferences = new Dictionary<string, FileSortPreference>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, FileSortPreference> listSortPreferences = new Dictionary<string, FileSortPreference>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ColumnLayout> columnLayouts = new Dictionary<string, ColumnLayout>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ListView> columnLayoutLists = new Dictionary<string, ListView>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, PaneWidthPreference> paneWidthPreferences = new Dictionary<string, PaneWidthPreference>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, SplitContainer> linkedPaneSplits = new Dictionary<string, SplitContainer>(StringComparer.OrdinalIgnoreCase);
    private readonly BackgroundWorker searchWorker = new BackgroundWorker();
    private SearchRequest pendingSearch;
    private int shortcutSortColumn = -1;
    private bool shortcutSortAscending = true;
    private int historySortColumn = -1;
    private bool historySortAscending = true;
    private int fileSortColumn = -1;
    private bool fileSortAscending = true;
    private int fileGroupMode;
    private int searchSortColumn = -1;
    private bool searchSortAscending = true;
    private Entry pendingFolder;
    private MiniRestoreIcon miniRestoreIcon;
    private NotifyIcon trayIcon;
    private bool exitRequested;
    private const int ActivateExistingMessage = 0x8000 + 0x351;
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHFileOperation(ref ShellFileOperation operation);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern bool SHObjectProperties(IntPtr hwnd, uint shopObjectType, string objectName, string printerName);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int SHParseDisplayName(string name, IntPtr bindContext, out IntPtr itemIdList, uint attributesIn, out uint attributesOut);

    [DllImport("shell32.dll", PreserveSig = true)]
    private static extern int SHBindToParent(IntPtr itemIdList, ref Guid interfaceId, out IntPtr parentFolder, out IntPtr childItemIdList);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr CreatePopupMenu();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyMenu(IntPtr menu);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern ushort RegisterClipboardFormat(string format);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref System.Runtime.InteropServices.ComTypes.STGMEDIUM medium);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint DragQueryFile(IntPtr drop, uint fileIndex, StringBuilder fileName, uint fileNameLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern bool GlobalUnlock(IntPtr memory);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ShellFileOperation
    {
        public IntPtr hwnd;
        public uint wFunc;
        public string pFrom;
        public string pTo;
        public ushort fFlags;
        [MarshalAs(UnmanagedType.Bool)] public bool fAnyOperationsAborted;
        public IntPtr hNameMappings;
        public string lpszProgressTitle;
    }

    [ComImport, Guid("000214E6-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellFolderNative
    {
        [PreserveSig] int ParseDisplayName(IntPtr hwnd, IntPtr bindContext, [MarshalAs(UnmanagedType.LPWStr)] string displayName, out uint eaten, out IntPtr itemIdList, ref uint attributes);
        [PreserveSig] int EnumObjects(IntPtr hwnd, uint flags, out IntPtr enumerator);
        [PreserveSig] int BindToObject(IntPtr itemIdList, IntPtr bindContext, ref Guid interfaceId, out IntPtr result);
        [PreserveSig] int BindToStorage(IntPtr itemIdList, IntPtr bindContext, ref Guid interfaceId, out IntPtr result);
        [PreserveSig] int CompareIDs(IntPtr lParam, IntPtr firstItemIdList, IntPtr secondItemIdList);
        [PreserveSig] int CreateViewObject(IntPtr hwnd, ref Guid interfaceId, out IntPtr result);
        [PreserveSig] int GetAttributesOf(uint count, IntPtr itemIdLists, ref uint attributes);
        [PreserveSig] int GetUIObjectOf(IntPtr hwnd, uint count, IntPtr itemIdLists, ref Guid interfaceId, IntPtr reserved, out IntPtr result);
        [PreserveSig] int GetDisplayNameOf(IntPtr itemIdList, uint flags, out IntPtr name);
        [PreserveSig] int SetNameOf(IntPtr hwnd, IntPtr itemIdList, [MarshalAs(UnmanagedType.LPWStr)] string name, uint flags, out IntPtr newItemIdList);
    }

    [ComImport, Guid("000214E4-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellContextMenuNative
    {
        [PreserveSig] int QueryContextMenu(IntPtr menu, uint indexMenu, uint firstCommand, uint lastCommand, uint flags);
        [PreserveSig] int InvokeCommand(ref ShellInvokeCommandInfo commandInfo);
        [PreserveSig] int GetCommandString(IntPtr commandOffset, uint flags, IntPtr reserved, IntPtr commandName, uint commandNameLength);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ShellInvokeCommandInfo
    {
        public uint Size;
        public uint Mask;
        public IntPtr Window;
        public IntPtr Verb;
        public IntPtr Parameters;
        public IntPtr Directory;
        public int Show;
        public uint HotKey;
        public IntPtr Icon;
    }

    private const ushort FileOperationAllowUndo = 0x0040;
    private const ushort FileOperationSilent = 0x0004;
    private const ushort FileOperationNoConfirmation = 0x0010;
    private Rectangle savedWindowBounds;
    private FormWindowState savedWindowState;
    private FormWindowState lastNonMinimizedState = FormWindowState.Normal;
    private bool savedTopMost;
    private string[] fileSelectionToRestore;
    private Point fileDragOrigin;
    private bool fileDragPending;
    private string lastFileLabelClickPath;
    private DateTime lastFileLabelClickAt = DateTime.MinValue;
    private bool allowInlineRenameStart;
    private string currentPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    private bool showingComputerDrives;
    private string computerViewReturnPath;
    private const string ThisComputerPath = "shell:thispc";
    private readonly string settingsFile = PortableDataPath("lastpath.txt");
    private readonly string navigationHistoryFile = PortableDataPath("navigation-history.txt");
    private readonly string workProfilesFile = PortableDataPath("workflows.xml");
    private readonly string workProfileLocationsFile = PortableDataPath("workspace-locations.xml");
    private readonly string shortcutsFile = PortableDataPath("shortcuts.txt");
    private readonly string historyFile = PortableDataPath("history.txt");
    private readonly string splitterFile = ComputerSettingPath("splitter.txt");
    private readonly string windowStateFile = ComputerSettingPath("maximized.txt");
    private readonly string windowBoundsFile = ComputerSettingPath("window-bounds.txt");
    private readonly string themeFile = ComputerSettingPath("theme.txt");
    private readonly string preferencesFile = ComputerSettingPath("preferences.txt");
    private readonly string folderSortFile = ComputerSettingPath("folder-sorts.txt");
    private readonly string listSortFile = ComputerSettingPath("list-sorts.txt");
    private readonly string columnLayoutsFile = ComputerSettingPath("column-layouts.txt");
    private readonly string linkedPaneWidthsFile = ComputerSettingPath("linked-pane-widths.txt");
    private readonly string tabPaneWidthsFile = ComputerSettingPath("tab-pane-widths.txt");
    private bool restoreMaximized;
    private bool closeToTrayOnExit = true;
    private bool keepWindowOnTop;
    private bool rememberFolderSort = true;
    private bool preserveHistoryRetention = true;
    private bool startMaximizedOnLaunch;
    private bool minimizeAfterOpeningFile = true;
    private bool folderTreeToggleMouseDown;
    private bool classicFolderView;
    private string oneDriveRoot = String.Empty;
    private string savedActiveTabName = String.Empty;
    private bool oneDriveRootEnabled;
    private readonly CheckBox oneDriveRootOption = new CheckBox();
    private readonly CheckBox automaticUpdateCheckOption = new CheckBox();
    private readonly TextBox updateManifestUrlBox = new TextBox();
    private readonly BackgroundWorker updateCheckWorker = new BackgroundWorker();
    private readonly BackgroundWorker updateInstallWorker = new BackgroundWorker();
    private string updateManifestUrl = DefaultUpdateManifestUrl;
    private bool automaticUpdateCheck = true;
    private bool manualUpdateCheck;
    private UpdateRelease pendingUpdateRelease;
    private const string CurrentAppVersion = "1.0.0";
    private const string DefaultUpdateManifestUrl = "";
    private const string UpdaterExecutableName = "DosyaGezgini-Updater.exe";
    private readonly TextBox oneDriveRootBox = new TextBox();
    private int historyRetentionCount = 30;
    private string historyRetentionUnit = "Gün";
    private string edefterTemplateFolder = String.Empty;
    private bool addressEditMode;
    private bool synchronizingFolderTree;
    private bool synchronizingFileSelection;
    private bool formHasBeenShown;
    private int navigationHistoryIndex = -1;
    private bool restoringColumnLayout;
    private bool deepSearchColumnsCustomized;
    private bool applyingPaneWidths;
    private bool panePersistenceReady;
    private string activePaneKey = "Gezgin";
    private bool searchSubfolders = true;
    private bool transferLeftPaneHidden;
    private bool transferRightPaneHidden;

    private static string PortableDataPath(string fileName)
    {
        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Dosya", fileName);
    }

    private static string ComputerSettingPath(string fileName)
    {
        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Dosya",
            "Bilgisayar-" + Environment.MachineName, fileName);
    }

    static FileExplorerStarter()
    {
        // Set .NET's long-path switches before any file-system APIs run, so the EXE needs no sidecar .config file.
        try
        {
            Type appContext = Type.GetType("System.AppContext");
            MethodInfo setSwitch = appContext == null ? null : appContext.GetMethod("SetSwitch",
                BindingFlags.Public | BindingFlags.Static, null, new Type[] { typeof(string), typeof(bool) }, null);
            if (setSwitch != null)
            {
                setSwitch.Invoke(null, new object[] { "Switch.System.IO.BlockLongPaths", false });
                setSwitch.Invoke(null, new object[] { "Switch.System.IO.UseLegacyPathHandling", false });
            }
        }
        catch { }
    }

    [STAThread]
    private static void Main()
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveEmbeddedPdfSharpAssembly;
        try { RtlSetProcessPlaceholderCompatibilityMode(2); } catch { }
        IntPtr existingWindow = FindMainWindow();
        if (existingWindow != IntPtr.Zero)
        {
            PostMessage(existingWindow, ActivateExistingMessage, IntPtr.Zero, IntPtr.Zero);
            return;
        }
        bool firstInstance;
        using (var singleInstance = new Mutex(true, @"Local\DosyaGezgini_SingleInstance", out firstInstance))
        {
            if (!firstInstance)
            {
                // The main form may be hidden while its tiny floating icon remains visible.
                // Give its message loop time to expose the hidden main window after startup.
                for (int attempt = 0; attempt < 10 && existingWindow == IntPtr.Zero; attempt++)
                {
                    Thread.Sleep(100);
                    existingWindow = FindMainWindow();
                }
                if (existingWindow != IntPtr.Zero)
                {
                    PostMessage(existingWindow, ActivateExistingMessage, IntPtr.Zero, IntPtr.Zero);
                }
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FileExplorerStarter());
        }
    }

    private static Assembly ResolveEmbeddedPdfSharpAssembly(object sender, ResolveEventArgs args)
    {
        if (args == null || args.Name == null) return null;
        string assemblyName = args.Name.Split(',')[0];
        string resourceName = null;
        if (String.Equals(assemblyName, "PdfSharp-gdi", StringComparison.OrdinalIgnoreCase)) resourceName = "PdfSharp-gdi.dll";
        else if (String.Equals(assemblyName, "System.Runtime.WindowsRuntime", StringComparison.OrdinalIgnoreCase)) resourceName = "System.Runtime.WindowsRuntime.dll";
        else if (String.Equals(assemblyName, "System.Runtime.InteropServices.WindowsRuntime", StringComparison.OrdinalIgnoreCase)) resourceName = "System.Runtime.InteropServices.WindowsRuntime.dll";
        if (resourceName == null) return null;
        Assembly application = Assembly.GetExecutingAssembly();
        using (Stream resource = application.GetManifestResourceStream(resourceName))
        {
            if (resource == null) return null;
            using (var bytes = new MemoryStream())
            {
                resource.CopyTo(bytes);
                return Assembly.Load(bytes.ToArray());
            }
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

    [DllImport("ntdll.dll", EntryPoint = "RtlSetProcessPlaceholderCompatibilityMode")]
    private static extern sbyte RtlSetProcessPlaceholderCompatibilityMode(byte mode);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateFile(string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes,
        uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandleEx(IntPtr fileHandle, int informationClass,
        out FileAttributeTagInformation information, uint bufferSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr handle);

    [DllImport("CldApi.dll", EntryPoint = "CfGetPlaceholderStateFromAttributeTag")]
    private static extern uint CfGetPlaceholderStateFromAttributeTag(uint fileAttributes, uint reparseTag);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetWindowText(IntPtr window, StringBuilder text, int maxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetWindowRect(IntPtr window, out NativeRect rect);

    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern int StrCmpLogicalW(string first, string second);

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SHGetFileInfo(string path, uint attributes, ref ShellFileInfo info, uint infoSize, uint flags);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct ShellFileInfo
    {
        public IntPtr Icon;
        public int IconIndex;
        public uint Attributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string DisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string TypeName;
    }

    private delegate bool EnumWindowsProc(IntPtr window, IntPtr lParam);

    private static IntPtr FindMainWindow()
    {
        IntPtr bestWindow = IntPtr.Zero;
        long largestArea = 0;
        EnumWindows((window, lParam) =>
        {
            var title = new StringBuilder(256);
            GetWindowText(window, title, title.Capacity);
            if (!title.ToString().StartsWith("Dosya Gezgini", StringComparison.Ordinal)) return true;
            NativeRect rect;
            if (!GetWindowRect(window, out rect)) return true;
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            long area = (long)width * height;
            // Ignore the 37x37 floating restore icon, which has the same window title.
            if (width >= 200 && height >= 150 && area > largestArea)
            {
                bestWindow = window;
                largestArea = area;
            }
            return true;
        }, IntPtr.Zero);
        return bestWindow;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == 0x0312 && m.WParam.ToInt32() == ShowAppHotKeyId)
        {
            RestoreFromFloatingIcon(); Activate(); SetForegroundWindow(Handle); return;
        }
        if (m.Msg == ActivateExistingMessage)
        {
            RestoreFromFloatingIcon();
            Activate();
            SetForegroundWindow(Handle);
            return;
        }
        base.WndProc(ref m);
    }

    private FileExplorerStarter()
    {
        LoadPreferences();
        InitializeUpdateWorkers();
        PromptForPortableOneDriveRootIfNeeded();
        LoadFolderSortPreferences();
        LoadListSortPreferences();
        LoadColumnLayouts();
        LoadTabPaneWidths();
        LoadLinkedPaneWidths();
        Text = "Dosya Gezgini";
        Icon = CreateExplorerIcon();
        InitializeTrayIcon();
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(760, 480);
        Size = new Size(1000, 650);
        RestoreWindowBounds();
        BackColor = Color.FromArgb(246, 248, 251);
        Font = new Font("Segoe UI", 9F);
        TopMost = keepWindowOnTop;
        Resize += (s, e) =>
        {
            if (WindowState == FormWindowState.Minimized)
            {
                if (formHasBeenShown) MinimizeToFloatingIcon();
            }
            else if (miniRestoreIcon == null)
                lastNonMinimizedState = WindowState;
        };

        var top = new Panel { Dock = DockStyle.Top, Height = 72, BackColor = Color.White, Padding = new Padding(18, 12, 18, 10) };
        mainHeader = top;
        StyleIconButton(back, String.Empty, "Önceki klasöre git", new Font("Segoe UI Symbol", 9F));
        back.Image = CreateHistoryArrowIcon(false);
        back.ImageAlign = ContentAlignment.MiddleCenter;
        back.Padding = Padding.Empty;
        back.SetBounds(18, 28, 28, 26);
        back.Click += (s, e) => NavigateHistory(-1);
        back.MouseUp += (s, e) => { if (e.Button == MouseButtons.Right) ShowNavigationHistoryMenu(back, backHistoryMenu, -1); };
        StyleIconButton(forward, String.Empty, "Sonraki klasöre git", new Font("Segoe UI Symbol", 9F));
        forward.Image = CreateHistoryArrowIcon(true);
        forward.ImageAlign = ContentAlignment.MiddleCenter;
        forward.Padding = Padding.Empty;
        forward.SetBounds(50, 28, 28, 26);
        forward.Click += (s, e) => NavigateHistory(1);
        forward.MouseUp += (s, e) => { if (e.Button == MouseButtons.Right) ShowNavigationHistoryMenu(forward, forwardHistoryMenu, 1); };
        StyleIconButton(up, String.Empty, "Bir üst klasöre çık", new Font("Segoe UI Symbol", 9F));
        up.Image = CreateUpButtonIcon();
        up.ImageAlign = ContentAlignment.MiddleCenter;
        up.Padding = Padding.Empty;
        up.SetBounds(82, 28, 28, 26);
        up.Click += (s, e) =>
        {
            if (showingComputerDrives)
            {
                Navigate(computerViewReturnPath, null, null, false);
                return;
            }
            DirectoryInfo parent = Directory.GetParent(currentPath);
            if (parent != null) Navigate(parent.FullName);
            else ShowComputerDrives();
        };
        StyleIconButton(refresh, String.Empty, "Geçerli klasörün içeriğini yenile", new Font("Segoe UI Symbol", 9F));
        refresh.Image = CreateRefreshButtonIcon();
        refresh.ImageAlign = ContentAlignment.MiddleCenter;
        refresh.Padding = Padding.Empty;
        refresh.SetBounds(114, 28, 28, 26);
        refresh.Click += (s, e) => { if (showingComputerDrives) ShowComputerDrives(); else Navigate(currentPath, null, null, false); };
        StyleIconButton(quickLocations, "⌄", "Windows klasör kısayollarını aç", new Font("Segoe UI Symbol", 10F));
        quickLocations.SetBounds(146, 28, 26, 26);
        quickLocations.Padding = Padding.Empty;
        quickLocations.Click += (s, e) => ShowQuickLocationsMenu();
        address.SetBounds(178, 28, 706, 30);
        address.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
        address.Font = new Font("Segoe UI", 10F);
        address.BorderStyle = BorderStyle.FixedSingle;
        address.Visible = false;
        address.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                NavigateFromAddress();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                SetAddressEditMode(false);
                e.SuppressKeyPress = true;
            }
        };
        address.LostFocus += (s, e) =>
        {
            if (addressEditMode) SetAddressEditMode(false);
        };
        breadcrumbBar.SetBounds(178, 28, 706, 30);
        breadcrumbBar.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
        breadcrumbBar.FlowDirection = FlowDirection.LeftToRight;
        breadcrumbBar.WrapContents = false;
        // Keep the breadcrumb from adding a visible horizontal scrollbar when
        // the window is narrow. RefreshBreadcrumbs trims old path segments instead.
        breadcrumbBar.AutoScroll = false;
        breadcrumbBar.Padding = new Padding(3, 1, 3, 0);
        breadcrumbBar.Margin = Padding.Empty;
        breadcrumbBar.BackColor = Color.White;
        breadcrumbBar.BorderStyle = BorderStyle.FixedSingle;
        breadcrumbBar.DoubleClick += (s, e) => SetAddressEditMode(true);
        actionTooltips.SetToolTip(breadcrumbBar, "Klasör basamağına tıklayın. Adres yazmak için Ctrl+L.");
        StyleIconButton(go, "➜", "Adres çubuğundaki konuma git", new Font("Segoe UI Symbol", 9F, FontStyle.Bold));
        go.SetBounds(0, 28, 38, 26);
        go.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        go.Click += (s, e) => NavigateFromAddress();
        StyleIconButton(addShortcut, String.Empty, "Seçili dosya veya klasörü Kısayollarım'a ekle", new Font("Segoe UI Symbol", 9F));
        addShortcut.Image = CreateShortcutButtonIcon();
        addShortcut.ImageAlign = ContentAlignment.MiddleCenter;
        addShortcut.Padding = Padding.Empty;
        addShortcut.SetBounds(0, 28, 38, 26);
        addShortcut.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        addShortcut.Click += (s, e) => AddSelectedShortcuts();
        StyleIconButton(togglePreview, String.Empty, "Seçili dosyanın ön izlemesini aç veya kapat", new Font("Segoe UI Symbol", 9F));
        togglePreview.Image = CreatePreviewButtonIcon();
        togglePreview.ImageAlign = ContentAlignment.MiddleCenter;
        togglePreview.Padding = Padding.Empty;
        togglePreview.SetBounds(0, 28, 38, 26);
        togglePreview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        togglePreview.Click += (s, e) => ToggleEmbeddedPreview();
        StyleIconButton(openExplorer, String.Empty, "Geçerli klasörü Windows Explorer'da aç", new Font("Segoe UI Symbol", 9F));
        openExplorer.Image = CreateToolbarFolderIcon();
        openExplorer.ImageAlign = ContentAlignment.MiddleCenter;
        openExplorer.Padding = Padding.Empty;
        openExplorer.SetBounds(0, 28, 38, 26);
        openExplorer.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        openExplorer.Click += (s, e) => OpenInWindowsExplorer(currentPath);
        top.Controls.Add(back);
        top.Controls.Add(forward);
        top.Controls.Add(up);
        top.Controls.Add(refresh);
        top.Controls.Add(quickLocations);
        top.Controls.Add(address);
        top.Controls.Add(breadcrumbBar);
        top.Controls.Add(go);
        top.Controls.Add(addShortcut);
        top.Controls.Add(togglePreview);
        top.Controls.Add(openExplorer);
        Action layoutAddressButtons = () =>
        {
            addShortcut.Left = top.ClientSize.Width - 56;
            togglePreview.Left = addShortcut.Left - 42;
            go.Left = togglePreview.Left - 42;
            openExplorer.Left = go.Left - 42;
            address.Width = Math.Max(80, openExplorer.Left - address.Left - 12);
            breadcrumbBar.SetBounds(address.Left, address.Top, address.Width, address.Height);
            if (breadcrumbBar.Visible) RefreshBreadcrumbs();
        };
        top.Resize += (s, e) => layoutAddressButtons();
        layoutAddressButtons();

        var explorerPage = new TabPage("Gezgin");
        var shortcutsPage = new TabPage("Kısayollarım");
        var historyPage = new TabPage("Geçmiş");
        var searchPage = new TabPage("Arama");
        mediaPage = new TabPage("Medya");
        transferPage = new TabPage("Çiftli İşlem");
        syncPage = new TabPage("Senkronize");
        workspacesPage = new TabPage("İş Klasörleri");
        developmentPage = new TabPage("Geliştirme");
        var settingsPage = new TabPage("Ayarlar");
        tabs.Dock = DockStyle.Fill;
        tabs.TabPages.Add(explorerPage);
        tabs.TabPages.Add(shortcutsPage);
        tabs.TabPages.Add(historyPage);
        tabs.TabPages.Add(searchPage);
        tabs.TabPages.Add(mediaPage);
        tabs.TabPages.Add(transferPage);
        tabs.TabPages.Add(syncPage);
        tabs.TabPages.Add(workspacesPage);
        tabs.TabPages.Add(developmentPage);
        tabs.TabPages.Add(settingsPage);
        if (!String.IsNullOrWhiteSpace(savedActiveTabName))
        {
            TabPage savedTab = tabs.TabPages.Cast<TabPage>().FirstOrDefault(tab =>
                String.Equals(tab.Text, savedActiveTabName, StringComparison.OrdinalIgnoreCase));
            if (savedTab != null) tabs.SelectedTab = savedTab;
        }
        tabs.SelectedIndexChanged += (s, e) =>
        {
            savedActiveTabName = tabs.SelectedTab == null ? String.Empty : tabs.SelectedTab.Text;
            SavePreferences();
        };
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 140, BackColor = Color.FromArgb(226, 231, 238), BorderStyle = BorderStyle.FixedSingle };
        folderFilesSplit = split;
        explorerLayout = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, BackColor = Color.FromArgb(226, 231, 238), BorderStyle = BorderStyle.FixedSingle };
        explorerLayout.Panel2Collapsed = true;
        previewPane = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(10) };
        var previewTitle = new Label { Text = "Ön izleme", Dock = DockStyle.Top, Height = 30, ForeColor = Color.FromArgb(38, 51, 70), Font = new Font("Segoe UI", 10F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        previewContent = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
        previewPane.Controls.Add(previewContent);
        previewPane.Controls.Add(previewTitle);
        explorerLayout.Panel2.Controls.Add(previewPane);
        explorerLayout.Panel1.Controls.Add(split);
        explorerLayout.SizeChanged += (s, e) =>
        {
            if (!previewPaneOpen || explorerLayout.ClientSize.Width <= 0) return;
            int desired = (int)(explorerLayout.ClientSize.Width * 0.58);
            int max = explorerLayout.ClientSize.Width - explorerLayout.Panel2MinSize - explorerLayout.SplitterWidth;
            explorerLayout.SplitterDistance = Math.Max(explorerLayout.Panel1MinSize, Math.Min(max, desired));
        };
        files.SelectedIndexChanged += (s, e) =>
        {
            UpdateSelectedFileCountStatus();
            if (previewPaneOpen && files.SelectedItems.Count > 0)
            {
                Entry selected = files.SelectedItems[0].Tag as Entry;
                if (selected != null && File.Exists(selected.FullPath)) ShowEmbeddedPreview(selected.FullPath);
            }
        };
        split.SplitterMoved += (s, e) =>
        {
            if (applyingPaneWidths) return;
            SaveSplitterDistance(split.SplitterDistance);
            SaveActivePaneWidths();
        };
        split.Panel1.Padding = new Padding(14, 14, 7, 14);
        split.Panel2.Padding = new Padding(7, 14, 14, 14);
        folderHeading.Text = "Klasörler";
        fileHeading.Text = "Dosyalar";
        StyleHeading(folderHeading);
        StyleHeading(fileHeading);
        folderHeading.Dock = DockStyle.Top;
        folderHeading.Height = 30;
        folderHeading.TextAlign = ContentAlignment.MiddleLeft;
        classicFolderViewOption.Text = "Klasik";
        classicFolderViewOption.AutoSize = true;
        classicFolderViewOption.Checked = classicFolderView;
        treeFolderViewOption.Text = "Ağaç";
        treeFolderViewOption.AutoSize = true;
        treeFolderViewOption.Checked = !classicFolderView;
        fileHeading.Dock = DockStyle.Top;
        StyleList(folders);
        folders.Dock = DockStyle.Fill;
        folders.DrawMode = DrawMode.OwnerDrawFixed;
        folders.DrawItem += (s, e) => DrawFolderListItem(folders, e);
        folders.AllowDrop = true;
        folders.DragEnter += (s, e) => e.Effect = GetFileDropEffect(e);
        folders.DragOver += (s, e) => e.Effect = GetFileDropEffect(e);
        folders.DragDrop += (s, e) =>
        {
            if (e.Effect == DragDropEffects.None || e.Data == null) return;
            string destination = currentPath;
            int index = folders.IndexFromPoint(folders.PointToClient(new Point(e.X, e.Y)));
            Entry folder = index >= 0 ? folders.Items[index] as Entry : null;
            if (folder != null && !folder.IsParent && Directory.Exists(folder.FullPath)) destination = folder.FullPath;
            ProcessExplorerDrop(e.Data, destination, e.Effect != DragDropEffects.Copy);
        };
        StyleDetailsList(files, "Ad", "Değiştirme tarihi");
        InitializeFileIconList();
        files.SmallImageList = fileIconImageList;
        files.View = View.Details;
        files.Columns.Add("Tür", 165, HorizontalAlignment.Left);
        oneDriveStatusColumn = new ColumnHeader { Text = "OneDrive", Width = 105, TextAlign = HorizontalAlignment.Center };
        files.Columns.Add(oneDriveStatusColumn);
        files.ShowItemToolTips = true;
        files.OwnerDraw = true;
        files.DrawColumnHeader += (s, e) => e.DrawDefault = true;
        files.DrawItem += (s, e) => e.DrawDefault = false;
        files.DrawSubItem += DrawFileSubItem;
        files.CheckBoxes = true;
        files.LabelEdit = true;
        files.BeforeLabelEdit += (s, e) =>
        {
            if (!allowInlineRenameStart || showingComputerDrives || e.Item < 0 || e.Item >= files.Items.Count) { e.CancelEdit = true; return; }
            Entry entry = files.Items[e.Item].Tag as Entry;
            if (entry == null || (!File.Exists(entry.FullPath) && !Directory.Exists(entry.FullPath))) e.CancelEdit = true;
        };
        files.AfterLabelEdit += (s, e) =>
        {
            if (e.Label == null) return;
            e.CancelEdit = true;
            if (showingComputerDrives || e.Item < 0 || e.Item >= files.Items.Count) return;
            Entry entry = files.Items[e.Item].Tag as Entry;
            if (entry == null) return;
            string newName = e.Label.Trim();
            if (newName.Length == 0 || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                !String.Equals(Path.GetFileName(newName), newName, StringComparison.Ordinal))
            {
                MessageBox.Show(this, "Geçerli bir dosya veya klasör adı girin.", "Yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResumeInlineRename(e.Item, newName);
                return;
            }
            string sourcePath = entry.FullPath;
            string destinationPath = Path.Combine(Path.GetDirectoryName(sourcePath), newName);
            if (String.Equals(sourcePath, destinationPath, StringComparison.Ordinal)) return;
            try
            {
                if (File.Exists(destinationPath) || Directory.Exists(destinationPath))
                {
                    MessageBox.Show(this, "Bu adla bir dosya veya klasör zaten var.", "Yeniden adlandırılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ResumeInlineRename(e.Item, newName);
                    return;
                }
                if (Directory.Exists(sourcePath)) Directory.Move(sourcePath, destinationPath);
                else File.Move(sourcePath, destinationPath);
                BeginInvoke((MethodInvoker)(() =>
                {
                    RefreshFolderTreeChildrenAtPath(currentPath);
                    Navigate(currentPath, null, destinationPath, false);
                }));
            }
            catch (Exception ex)
            {
                if (!IsSharingViolation(ex))
                    MessageBox.Show(this, ex.Message, "Yeniden adlandırılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResumeInlineRename(e.Item, newName);
            }
        };
        folders.SelectionMode = SelectionMode.MultiExtended;
        files.MultiSelect = true;
        folders.MouseClick += (s, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            if ((ModifierKeys & (Keys.Control | Keys.Shift)) != Keys.None) return;
            int index = folders.IndexFromPoint(e.Location);
            Entry selected = index >= 0 ? folders.Items[index] as Entry : null;
            if (selected != null && String.Equals(selected.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) ShowComputerDrives();
            else if (selected != null && selected.IsParent && showingComputerDrives) Navigate(selected.FullPath, null, null, false);
            else if (selected != null && Directory.Exists(selected.FullPath)) Navigate(selected.FullPath, selected);
        };
        folders.KeyDown += (s, e) =>
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                for (int i = 0; i < folders.Items.Count; i++) folders.SetSelected(i, true);
            }
            else if (e.Control && (e.KeyCode == Keys.C || e.KeyCode == Keys.X))
            {
                string[] selectedPaths = GetSelectedFolderPaths(folders);
                if (selectedPaths.Length > 0)
                {
                    SetFileClipboard(selectedPaths, e.KeyCode == Keys.X);
                    pendingFolderClipboardPath = null;
                    pendingCutPaths.Clear();
                    pendingCutPaths.AddRange(e.KeyCode == Keys.X ? selectedPaths : new string[0]);
                    RefreshCutFileAppearance();
                    folders.Invalidate();
                    status.Text = selectedPaths.Length + " klasör " + (e.KeyCode == Keys.X ? "kesildi." : "kopyalandı.");
                }
            }
            else if (e.Control && e.KeyCode == Keys.V) PasteFiles();
            else if (e.KeyCode == Keys.Enter)
            {
                Entry selected = folders.SelectedItem as Entry;
                if (selected != null && String.Equals(selected.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) ShowComputerDrives();
                else if (selected != null && selected.IsParent && showingComputerDrives) Navigate(selected.FullPath, null, null, false);
                else if (selected != null && Directory.Exists(selected.FullPath)) Navigate(selected.FullPath, selected);
            }
            else if (e.KeyCode == Keys.F2 && folders.SelectedItem is Entry) RenameSelectedFolder(folders.SelectedItem as Entry);
            else return;
            e.Handled = e.SuppressKeyPress = true;
        };
        ConfigureMailContextMenu(folders);
        actionTooltips.SetToolTip(files, "Dosyaları tek tek işaretleyin; işaretli dosyalar kopyalama, taşıma ve ZIP işlemlerinde kullanılır.");
        files.ItemCheck += (s, e) =>
        {
            if (synchronizingFileSelection) return;
            int itemIndex = e.Index;
            bool selectItem = e.NewValue == CheckState.Checked;
            if (!files.IsHandleCreated) return;
            files.BeginInvoke((MethodInvoker)(() =>
            {
                if (itemIndex < 0 || itemIndex >= files.Items.Count) return;
                synchronizingFileSelection = true;
                try
                {
                    ListViewItem item = files.Items[itemIndex];
                    item.Selected = selectItem;
                    if (selectItem) files.FocusedItem = item;
                }
                finally { synchronizingFileSelection = false; }
            }));
        };
        files.ItemSelectionChanged += (s, e) =>
        {
            if (synchronizingFileSelection) return;
            synchronizingFileSelection = true;
            try { e.Item.Checked = e.IsSelected; }
            finally { synchronizingFileSelection = false; }
            UpdateSelectedFileCountStatus();
        };
        files.ItemCheck += (s, e) => BeginInvoke((MethodInvoker)UpdateSelectedFileCountStatus);
        ConfigureMailContextMenu(files);
        folderTree.Dock = DockStyle.Fill;
        folderTree.BorderStyle = BorderStyle.None;
        folderTree.BackColor = Color.White;
        folderTree.ForeColor = Color.FromArgb(38, 51, 70);
        folderTree.Font = new Font("Segoe UI", 10F);
        folderTree.HideSelection = false;
        folderTree.FullRowSelect = true;
        folderTree.DrawMode = TreeViewDrawMode.OwnerDrawText;
        folderTree.DrawNode += (s, e) =>
        {
            bool selected = e.Node == folderTree.SelectedNode;
            if (selected)
            {
                int selectionLeft = Math.Max(0, e.Bounds.Left - 2);
                int selectionRight = Math.Min(folderTree.ClientSize.Width, e.Bounds.Right + 3);
                using (var brush = new SolidBrush(Color.FromArgb(0, 120, 215)))
                    e.Graphics.FillRectangle(brush, new Rectangle(selectionLeft, e.Bounds.Top, Math.Max(0, selectionRight - selectionLeft), e.Bounds.Height));
            }
            TextRenderer.DrawText(e.Graphics, e.Node.Text, folderTree.Font, e.Bounds,
                selected ? Color.White : folderTree.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        };
        folderTree.LostFocus += (s, e) => folderTree.Invalidate();
        folderTree.GotFocus += (s, e) => folderTree.Invalidate();
        folderTree.ShowLines = true;
        folderTree.ShowPlusMinus = true;
        folderTree.BeforeExpand += (s, e) => LoadFolderTreeChildren(e.Node);
        folderTree.MouseDown += (s, e) =>
        {
            TreeViewHitTestInfo hit = folderTree.HitTest(e.Location);
            folderTreeToggleMouseDown = e.Button == MouseButtons.Left &&
                hit != null && (hit.Location & TreeViewHitTestLocations.PlusMinus) != 0;
        };
        folderTree.MouseUp += (s, e) => folderTreeToggleMouseDown = false;
        folderTree.AfterSelect += (s, e) =>
        {
            if (synchronizingFolderTree || folderTreeToggleMouseDown) return;
            string selectedPath = e.Node == null ? null : e.Node.Tag as string;
            if (!String.IsNullOrWhiteSpace(selectedPath) && Directory.Exists(selectedPath) &&
                !String.Equals(selectedPath, currentPath, StringComparison.OrdinalIgnoreCase))
                Navigate(selectedPath, new Entry(DisplayName(selectedPath), selectedPath));
        };
        Entry folderTreeContextTarget = null;
        var folderTreeMenu = new ContextMenuStrip();
        var openFolderInExplorer = new ToolStripMenuItem("Windows Explorer ile Aç");
        var renameFolderInTree = new ToolStripMenuItem("Değiştir...");
        var cutFolderInTree = new ToolStripMenuItem("Kes");
        var copyFolderInTree = new ToolStripMenuItem("Kopyala");
        var pasteFolderInTree = new ToolStripMenuItem("Yapıştır");
        var deleteFolderInTree = new ToolStripMenuItem("Sil");
        var mailFolder = new ToolStripMenuItem("Mail Gönder");
        var createTreeFolder = new ToolStripMenuItem("Klasör ekle");
        folderTreeMenu.Items.Add(openFolderInExplorer);
        folderTreeMenu.Items.Add(renameFolderInTree);
        folderTreeMenu.Items.Add(cutFolderInTree);
        folderTreeMenu.Items.Add(copyFolderInTree);
        folderTreeMenu.Items.Add(pasteFolderInTree);
        folderTreeMenu.Items.Add(deleteFolderInTree);
        folderTreeMenu.Items.Add(mailFolder);
        folderTreeMenu.Items.Add(new ToolStripSeparator());
        folderTreeMenu.Items.Add(createTreeFolder);
        folderTreeMenu.Opening += (s, e) =>
        {
            openFolderInExplorer.Enabled = folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath);
            renameFolderInTree.Enabled = folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath);
            cutFolderInTree.Enabled = copyFolderInTree.Enabled = folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath);
            pasteFolderInTree.Enabled = HasFileDropClipboard() || (!String.IsNullOrWhiteSpace(pendingFolderClipboardPath) && Directory.Exists(pendingFolderClipboardPath));
            deleteFolderInTree.Enabled = folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath) &&
                !String.Equals(folderTreeContextTarget.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase);
            mailFolder.Enabled = folderTreeContextTarget != null && File.Exists(folderTreeContextTarget.FullPath);
            createTreeFolder.Enabled = folderTreeContextTarget == null || Directory.Exists(folderTreeContextTarget.FullPath);
        };
        openFolderInExplorer.Click += (s, e) =>
        {
            if (folderTreeContextTarget != null) OpenInWindowsExplorer(folderTreeContextTarget.FullPath);
        };
        renameFolderInTree.Click += (s, e) => RenameSelectedFolder(folderTreeContextTarget);
        cutFolderInTree.Click += (s, e) =>
        {
            if (folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath)) SetFileClipboard(new[] { folderTreeContextTarget.FullPath }, true);
        };
        copyFolderInTree.Click += (s, e) =>
        {
            if (folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath)) SetFileClipboard(new[] { folderTreeContextTarget.FullPath }, false);
        };
        pasteFolderInTree.Click += (s, e) => PasteClipboardIntoFolder(folderTreeContextTarget != null ? folderTreeContextTarget.FullPath : currentPath);
        deleteFolderInTree.Click += (s, e) => DeleteFolderTreeSelection(folderTreeContextTarget);
        mailFolder.Click += (s, e) =>
        {
            if (folderTreeContextTarget != null && File.Exists(folderTreeContextTarget.FullPath))
                SendByMail(folderTreeContextTarget.FullPath);
        };
        createTreeFolder.Click += (s, e) => CreateFolderIn(folderTreeContextTarget != null && Directory.Exists(folderTreeContextTarget.FullPath)
            ? folderTreeContextTarget.FullPath : currentPath);
        folderTree.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            TreeNode hitNode = folderTree.GetNodeAt(e.Location);
            folderTreeContextTarget = hitNode == null ? null : new Entry(hitNode.Text, hitNode.Tag as string);
        };
        folderTree.ContextMenuStrip = folderTreeMenu;
        folderTree.KeyDown += (s, e) =>
        {
            string path = folderTree.SelectedNode == null ? null : folderTree.SelectedNode.Tag as string;
            if (e.Control && (e.KeyCode == Keys.C || e.KeyCode == Keys.X) && !String.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                SetFileClipboard(new[] { path }, e.KeyCode == Keys.X);
            else if (e.Control && e.KeyCode == Keys.V && !String.IsNullOrWhiteSpace(path) && Directory.Exists(path)) PasteClipboardIntoFolder(path);
            else if (e.KeyCode == Keys.F2 && !String.IsNullOrWhiteSpace(path) && Directory.Exists(path)) RenameSelectedFolder(new Entry(folderTree.SelectedNode.Text, path));
            else if (e.KeyCode == Keys.Enter && !String.IsNullOrWhiteSpace(path) && Directory.Exists(path)) Navigate(path, new Entry(folderTree.SelectedNode.Text, path));
            else return;
            e.Handled = e.SuppressKeyPress = true;
        };
        folderTree.AllowDrop = true;
        folderTree.DragEnter += (s, e) =>
        {
            e.Effect = GetFileDropEffect(e);
        };
        folderTree.DragOver += (s, e) =>
        {
            e.Effect = GetFileDropEffect(e);
            if (e.Effect == DragDropEffects.None) return;
            Point point = folderTree.PointToClient(new Point(e.X, e.Y));
            TreeNode target = folderTree.GetNodeAt(point);
            string targetPath = target == null ? null : target.Tag as string;
            if (String.IsNullOrWhiteSpace(targetPath) || !Directory.Exists(targetPath))
            {
                e.Effect = DragDropEffects.None;
                return;
            }
            synchronizingFolderTree = true;
            folderTree.SelectedNode = target;
            synchronizingFolderTree = false;
        };
        folderTree.DragDrop += (s, e) =>
        {
            if (e.Effect == DragDropEffects.None || e.Data == null) return;
            Point point = folderTree.PointToClient(new Point(e.X, e.Y));
            TreeNode target = folderTree.GetNodeAt(point);
            string targetPath = target == null ? null : target.Tag as string;
            if (!String.IsNullOrWhiteSpace(targetPath) && Directory.Exists(targetPath))
                ProcessExplorerDrop(e.Data, targetPath, e.Effect != DragDropEffects.Copy);
        };
        folderTree.NodeMouseClick += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            synchronizingFolderTree = true;
            folderTree.SelectedNode = e.Node;
            synchronizingFolderTree = false;
            folderTreeContextTarget = new Entry(e.Node.Text, e.Node.Tag as string);
        };
        PopulateFolderTreeRoots();
        files.Dock = DockStyle.Fill;
        files.AllowDrop = false;
        filesAsyncDropTarget = new OutlookAsyncDropTarget(files, (paths, move) =>
        {
            if (paths != null && paths.Length > 0)
                TransferExplorerDrop(paths, currentPath, move);
        }, message =>
        {
            if (!IsDisposed && status.IsHandleCreated) status.BeginInvoke((MethodInvoker)(() => status.Text = message));
        });
        folderSearch.Dock = DockStyle.Top;
        folderSearch.Height = 30;
        folderSearch.Font = new Font("Segoe UI", 10F);
        folderSearch.BorderStyle = BorderStyle.FixedSingle;
        folderSearch.TabStop = true;
        folderSearch.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Tab) e.IsInputKey = true; };
        folderSearch.TextChanged += (s, e) => SelectNearestFolder(folderSearch.Text);
        folderSearch.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Tab)
            {
                CycleSearchFocus(folderSearch, !e.Shift);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                if (String.Equals(folderSearch.Text.Trim(), ".", StringComparison.Ordinal))
                {
                    if (showingComputerDrives) Navigate(computerViewReturnPath, null, null, false);
                    else
                    {
                        DirectoryInfo parent = Directory.GetParent(currentPath);
                        if (parent != null) Navigate(parent.FullName);
                        else ShowComputerDrives();
                    }
                }
                else
                {
                    Entry selected = folders.SelectedItem as Entry;
                    if (selected != null && String.Equals(selected.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) ShowComputerDrives();
                    else if (selected != null && selected.IsParent && showingComputerDrives) Navigate(selected.FullPath, null, null, false);
                    else if (selected != null && Directory.Exists(selected.FullPath)) Navigate(selected.FullPath, selected);
                }
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                MoveFolderSelection(e.KeyCode == Keys.Up ? -1 : 1);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        };
        var folderSearchTip = new ToolTip();
        folderSearchTip.SetToolTip(folderSearch, "Klasör ara: oklarla seç, Enter ile aç");
        InitializeFilter(fileFilter, () => { RefreshFileRows(); SelectNearestFile(fileSearch.Text); });
        fileSearch.Dock = DockStyle.Top;
        fileSearch.Height = 30;
        fileSearch.Font = new Font("Segoe UI", 10F);
        fileSearch.BorderStyle = BorderStyle.FixedSingle;
        fileSearch.TabStop = true;
        fileSearch.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Tab) e.IsInputKey = true; };
        fileSearch.TextChanged += (s, e) =>
        {
            RefreshFileRows();
            SelectNearestFile(fileSearch.Text);
        };
        fileSearch.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Tab)
            {
                CycleSearchFocus(fileSearch, !e.Shift);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                Entry selected = files.SelectedItems.Count > 0 ? files.SelectedItems[0].Tag as Entry : null;
                if (selected != null) OpenFile(selected.FullPath);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                MoveFileSelection(e.KeyCode == Keys.Up ? -1 : 1);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        };
        var fileSearchTip = new ToolTip();
        fileSearchTip.SetToolTip(fileSearch, "Dosya ara: oklarla seç, Enter ile aç");
        split.Panel1.Controls.Add(folderTree);
        split.Panel1.Controls.Add(folders);
        split.Panel1.Controls.Add(folderHeading);
        split.Panel1.Controls.Add(folderSearch);
        ApplyFolderViewMode();
        split.Panel2.Controls.Add(files);
        split.Panel2.Controls.Add(fileHeading);
        split.Panel2.Controls.Add(fileSearch);
        var fileFilterHost = new Panel { Dock = DockStyle.Bottom, Height = 34, Padding = new Padding(0, 2, 0, 0), BackColor = Color.White };
        fileFilter.Dock = DockStyle.Fill;
        fileFilterHost.Controls.Add(fileFilter);
        split.Panel2.Controls.Add(fileFilterHost);
        InitializeFolderTabBar(explorerLayout, explorerPage);
        StyleDetailsList(shortcutList, "Ad", "Eklenme zamanı");
        SetSortIndicators(shortcutList, "Ad", "Eklenme zamanı", shortcutSortColumn, shortcutSortAscending);
        shortcutList.Dock = DockStyle.Fill;
        shortcutList.SmallImageList = fileIconImageList;
        shortcutList.MultiSelect = true;
        shortcutList.MouseClick += (s, e) => RevealClickedItem(shortcutList, e);
        shortcutList.MouseDoubleClick += (s, e) => OpenListItemInExplorer(shortcutList, e.Location);
        shortcutList.ColumnClick += (s, e) => SortShortcuts(e.Column, true);
        ConfigureItemContextMenu(shortcutList, false);
        shortcutsPage.Controls.Add(shortcutList);
        InitializeFilter(shortcutFilter, () => { RefreshShortcutList(); SelectNearestShortcut(shortcutSearch.Text); });
        shortcutsPage.Controls.Add(shortcutFilter);
        InitializeSearchBox(shortcutSearch, "Kısayollarımda ara", SelectNearestShortcut, () => RevealSelectedItem(shortcutList), shortcutList);
        shortcutsPage.Controls.Add(shortcutSearch);
        StyleDetailsList(historyList, "Ad", "Eklenme zamanı");
        SetSortIndicators(historyList, "Ad", "Eklenme zamanı", historySortColumn, historySortAscending);
        historyList.Dock = DockStyle.Fill;
        historyList.SmallImageList = fileIconImageList;
        historyList.MultiSelect = true;
        historyList.MouseClick += (s, e) => RevealClickedItem(historyList, e);
        historyList.MouseDoubleClick += (s, e) => OpenListItemInExplorer(historyList, e.Location);
        historyList.ColumnClick += (s, e) => SortHistory(e.Column, true);
        ConfigureItemContextMenu(historyList, true);
        historyPage.Controls.Add(historyList);
        InitializeFilter(historyFilter, () => { RefreshHistoryList(); SelectNearestHistory(historySearch.Text); });
        historyPage.Controls.Add(historyFilter);
        InitializeSearchBox(historySearch, "Geçmişte ara", SelectNearestHistory, () => RevealSelectedItem(historyList), historyList);
        historyPage.Controls.Add(historySearch);
        InitializeDeepSearchPage(searchPage);
        InitializeMediaPage(mediaPage);
        AttachColumnLayout(files, "Gezgin");
        AttachColumnLayout(shortcutList, "Kısayollarım");
        AttachColumnLayout(historyList, "Geçmiş");
        AttachColumnLayout(deepSearchResults, "Arama");
        InitializeWorkspacesPage(workspacesPage);
        InitializeDevelopmentPage(developmentPage);
        InitializeSettingsPage(settingsPage);
        InitializeTransferPage(transferPage);
        InitializeSyncPage(syncPage);
        SplitContainer shortcutsExplorerSplit = CreateLinkedExplorerSplit(shortcutsPage, "Kısayollarım");
        SplitContainer historyExplorerSplit = CreateLinkedExplorerSplit(historyPage, "Geçmiş");
        SplitContainer searchExplorerSplit = CreateLinkedExplorerSplit(searchPage, "Arama");
        tabs.SelectedIndexChanged += (s, e) =>
        {
            SaveActivePaneWidths();
            applyingPaneWidths = true;
            if (tabs.SelectedTab == shortcutsPage)
            {
                activePaneKey = "Kısayollarım";
                shortcutsExplorerSplit.Panel2.Controls.Add(explorerWorkspace);
                ScheduleActivePaneWidthRestore(activePaneKey);
            }
            else if (tabs.SelectedTab == historyPage)
            {
                activePaneKey = "Geçmiş";
                historyExplorerSplit.Panel2.Controls.Add(explorerWorkspace);
                ScheduleActivePaneWidthRestore(activePaneKey);
            }
            else if (tabs.SelectedTab == searchPage)
            {
                activePaneKey = "Arama";
                searchExplorerSplit.Panel2.Controls.Add(explorerWorkspace);
                ScheduleActivePaneWidthRestore(activePaneKey);
            }
            else if (tabs.SelectedTab == transferPage)
            {
                activePaneKey = "Çiftli İşlem";
                explorerPage.Controls.Add(explorerWorkspace);
                SyncTransferRightPaneToExplorer();
            }
            else if (tabs.SelectedTab == mediaPage)
            {
                activePaneKey = "Medya";
                explorerPage.Controls.Add(explorerWorkspace);
            }
            else
            {
                activePaneKey = tabs.SelectedTab == null ? "Gezgin" : tabs.SelectedTab.Text;
                explorerPage.Controls.Add(explorerWorkspace);
                if (activePaneKey == "Gezgin") ScheduleActivePaneWidthRestore(activePaneKey);
            }
            applyingPaneWidths = false;
        };
        workspacesPage.Enter += (s, e) => RefreshWorkProfileDetails();
        developmentPage.Enter += (s, e) => { MergeDevelopmentBoardFromDisk(); RefreshDevelopmentBoard(); };
        files.ColumnClick += (s, e) => SortFiles(e.Column, true);
        files.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter && files.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = files.SelectedItems[0];
                Entry selectedEntry = selectedItem.Tag as Entry;
                if (selectedEntry != null && Directory.Exists(selectedEntry.FullPath))
                    Navigate(selectedEntry.FullPath, selectedEntry);
                else if (!showingComputerDrives && selectedEntry != null && File.Exists(selectedEntry.FullPath))
                    OpenFile(selectedEntry.FullPath);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Space && files.SelectedItems.Count > 0)
            {
                ToggleEmbeddedPreview();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.A)
            {
                files.BeginUpdate();
                foreach (ListViewItem item in files.Items)
                {
                    item.Checked = true;
                    item.Selected = true;
                }
                if (files.Items.Count > 0) files.FocusedItem = files.Items[files.Items.Count - 1];
                files.EndUpdate();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.C) { if (!showingComputerDrives) CopySelectedExplorerItems(false); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.X) { if (!showingComputerDrives) CopySelectedExplorerItems(true); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.V) { if (!showingComputerDrives) PasteFiles(); e.Handled = true; }
            else if (e.KeyCode == Keys.Delete)
            {
                if (!showingComputerDrives) DeleteSelectedExplorerFiles();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.F2 && files.SelectedItems.Count > 0)
            {
                if (!showingComputerDrives) BeginInlineRename(files.SelectedItems[0]);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        };
        files.MouseDoubleClick += (s, e) => OpenFileAt(e.Location);
        files.MouseClick += (s, e) =>
        {
            if (e.Button != MouseButtons.Left || showingComputerDrives) return;
            if (e.Clicks != 1)
            {
                lastFileLabelClickPath = null;
                return;
            }
            ListViewHitTestInfo hitInfo = files.HitTest(e.Location);
            ListViewItem hit = hitInfo.Item;
            if (hit == null || hitInfo.SubItem == null || hit.SubItems.IndexOf(hitInfo.SubItem) != 0 ||
                (hitInfo.Location & ListViewHitTestLocations.StateImage) != 0)
            {
                lastFileLabelClickPath = null;
                return;
            }
            Entry entry = hit.Tag as Entry;
            if (entry == null || !hit.Selected)
            {
                lastFileLabelClickPath = null;
                return;
            }
            DateTime now = DateTime.Now;
            double elapsedMs = (now - lastFileLabelClickAt).TotalMilliseconds;
            bool secondSlowClick = String.Equals(lastFileLabelClickPath, entry.FullPath, StringComparison.OrdinalIgnoreCase) &&
                elapsedMs >= SystemInformation.DoubleClickTime && elapsedMs <= SystemInformation.DoubleClickTime * 4;
            lastFileLabelClickPath = entry.FullPath;
            lastFileLabelClickAt = now;
            if (!secondSlowClick) return;
            lastFileLabelClickPath = null;
            BeginInlineRename(hit);
        };
        files.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            ListViewHitTestInfo hitInfo = files.HitTest(e.Location);
            ListViewItem hit = hitInfo.Item;
            bool checkboxClick = (hitInfo.Location & ListViewHitTestLocations.StateImage) != 0;
            fileDragPending = hit != null && !checkboxClick;
            fileDragOrigin = e.Location;
            bool extendSelection = (Control.ModifierKeys & (Keys.Control | Keys.Shift)) != Keys.None;
            if (hit != null && !hit.Selected && !extendSelection && !checkboxClick)
            {
                foreach (ListViewItem selected in files.SelectedItems) selected.Selected = false;
                hit.Selected = true;
                files.FocusedItem = hit;
            }
        };
        files.MouseMove += (s, e) =>
        {
            if (!fileDragPending || (Control.MouseButtons & MouseButtons.Left) == 0) return;
            var dragBounds = new Rectangle(
                fileDragOrigin.X - SystemInformation.DragSize.Width / 2,
                fileDragOrigin.Y - SystemInformation.DragSize.Height / 2,
                SystemInformation.DragSize.Width,
                SystemInformation.DragSize.Height);
            if (dragBounds.Contains(e.Location)) return;
            fileDragPending = false;
            string[] paths = GetSelectedOrCheckedFileAndFolderPaths();
            if (paths.Length == 0) return;
            bool containsFolder = paths.Any(Directory.Exists);
            bool copyAtStart = containsFolder || (Control.ModifierKeys & Keys.Control) != 0;
            var data = new DataObject();
            data.SetData(DataFormats.FileDrop, true, paths);
            data.SetData("DosyaGezgini.InternalFileMove", true);
            using (var preferredEffect = new MemoryStream(BitConverter.GetBytes(copyAtStart ? 1 : 2)))
                data.SetData(DataFormats.GetFormat("Preferred DropEffect").Name, false, preferredEffect.ToArray());
            var feedback = new TransferDragFeedbackForm(paths.Length);
            GiveFeedbackEventHandler updateFeedback = (sender, args) =>
            {
                bool copying = args.Effect == DragDropEffects.Copy;
                feedback.SetDragInfo(copying ? "Kopyalanıyor" : "Taşınıyor", Cursor.Position);
                args.UseDefaultCursors = true;
            };
            files.GiveFeedback += updateFeedback;
            try
            {
                feedback.Show(this);
                feedback.SetDragInfo(copyAtStart ? "Kopyalanıyor" : "Taşınıyor", Cursor.Position);
                files.DoDragDrop(data, DragDropEffects.Copy | DragDropEffects.Move);
            }
            finally
            {
                files.GiveFeedback -= updateFeedback;
                feedback.Close();
                feedback.Dispose();
            }
        };
        files.MouseUp += (s, e) => fileDragPending = false;
        files.AllowDrop = true;
        files.DragEnter += (s, e) => e.Effect = GetFileDropEffect(e);
        files.DragOver += (s, e) =>
        {
            DragDropEffects effect = GetFileDropEffect(e);
            if (effect == DragDropEffects.None) { e.Effect = effect; return; }
            e.Effect = effect;
        };
        files.DragDrop += (s, e) =>
        {
            if (e.Effect == DragDropEffects.None || e.Data == null) return;
            string destination = currentPath;
            ListViewHitTestInfo hit = files.HitTest(files.PointToClient(new Point(e.X, e.Y)));
            Entry target = hit.Item == null ? null : hit.Item.Tag as Entry;
            if (target != null && Directory.Exists(target.FullPath)) destination = target.FullPath;
            ProcessExplorerDrop(e.Data, destination, e.Effect != DragDropEffects.Copy);
        };

        status.Dock = DockStyle.Bottom;
        status.Height = 30;
        status.Padding = new Padding(18, 6, 0, 0);
        status.BackColor = Color.White;
        status.ForeColor = Color.FromArgb(82, 94, 112);
        Controls.Add(tabs);
        Controls.Add(status);
        Controls.Add(top);
        var versionLabel = new Label { AutoSize = true, Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
            ForeColor = Color.RoyalBlue, BackColor = Color.White, Font = new Font("Segoe UI", 8F, FontStyle.Bold),
            Text = "v" + Assembly.GetExecutingAssembly().GetName().Version.ToString() };
        status.Controls.Add(versionLabel);
        Action placeVersion = () => versionLabel.Location = new Point(Math.Max(0, status.ClientSize.Width - versionLabel.Width - 12), 6);
        status.Resize += (s, e) => placeVersion();
        placeVersion();
        KeyPreview = true;
        KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Escape && syncFullScreen)
            {
                ToggleSyncFullScreen(false);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.F11 && (syncFullScreen || tabs.SelectedTab == syncPage))
            {
                ToggleSyncFullScreen(!syncFullScreen);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.Shift && e.KeyCode == Keys.T)
            {
                ReopenClosedFolderTab();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.T)
            {
                CreateFolderTab(currentPath);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.W)
            {
                CloseFolderTab(folderTabs.SelectedTab);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.Tab)
            {
                if (folderTabs.TabPages.Count > 1)
                {
                    int delta = e.Shift ? -1 : 1;
                    int next = (folderTabs.SelectedIndex + delta + folderTabs.TabPages.Count) % folderTabs.TabPages.Count;
                    folderTabs.SelectedIndex = next;
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.L)
            {
                SetAddressEditMode(true);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape && addressEditMode)
            {
                SetAddressEditMode(false);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        };
        try
        {
            if (File.Exists(settingsFile))
            {
                string savedPath = ResolvePortablePath(File.ReadAllText(settingsFile).Trim());
                if (Directory.Exists(savedPath)) currentPath = savedPath;
            }
        }
        catch { }
        LoadNavigationHistory();
        LoadFolderTabSessions();
        restoreMaximized = LoadMaximizedState();
        Shown += (s, e) =>
        {
            appHotKeyRegistered = RegisterHotKey(Handle, ShowAppHotKeyId, ModWin | ModAlt, (uint)Keys.E);
            formHasBeenShown = true;
            if (restoreMaximized || startMaximizedOnLaunch)
                WindowState = FormWindowState.Maximized;
            panePersistenceReady = true;
            BeginInvoke((MethodInvoker)(() =>
            {
                RestoreAllColumnLayouts();
                ScheduleActivePaneWidthRestore("Gezgin");
                if (automaticUpdateCheck) CheckForUpdates(false);
                var updateTimer = new System.Windows.Forms.Timer { Interval = 300000 };
                updateTimer.Tick += (sender, args) => { if (automaticUpdateCheck) CheckForUpdates(false); };
                FormClosed += (sender, args) => updateTimer.Dispose();
                updateTimer.Start();
                var activeJobs = workProfiles.Where(profile => !profile.IsCompleted).Select(profile => profile.Name).ToArray();
                if (activeJobs.Length > 0)
                    MessageBox.Show(this, String.Join(Environment.NewLine, activeJobs), "Aktif işler", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }));
        };
        FormClosing += (s, e) =>
        {
            if (appHotKeyRegistered) { UnregisterHotKey(Handle, ShowAppHotKeyId); appHotKeyRegistered = false; }
            SaveSelectedSyncProfile();
            SaveActivePaneWidths();
            SaveFolderTabSessions();
            if (e.CloseReason == CloseReason.WindowsShutDown || e.CloseReason == CloseReason.TaskManagerClosing) exitRequested = true;
            if (!exitRequested && closeToTrayOnExit && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                MinimizeToFloatingIcon();
                return;
            }
            if (!exitRequested) exitRequested = true;
            if (currentFolderWatcher != null) { currentFolderWatcher.EnableRaisingEvents = false; currentFolderWatcher.Dispose(); currentFolderWatcher = null; }
            DisposeTransferPaneWatcher(transferLeftPane);
            DisposeTransferPaneWatcher(transferRightPane);
            if (directoryRefreshTimer != null) { directoryRefreshTimer.Stop(); directoryRefreshTimer.Dispose(); directoryRefreshTimer = null; }
            if (oneDriveStatusTimer != null) { oneDriveStatusTimer.Stop(); oneDriveStatusTimer.Dispose(); oneDriveStatusTimer = null; }
            if (oneDriveStatusWorker != null && oneDriveStatusWorker.IsBusy) oneDriveStatusWorker.CancelAsync();
            if ((oneDriveStatusWorker == null || !oneDriveStatusWorker.IsBusy) && shellApplication != null && Marshal.IsComObject(shellApplication))
            { try { Marshal.FinalReleaseComObject(shellApplication); } catch { } shellApplication = null; }
            if (deepSearchBoldFont != null) { deepSearchBoldFont.Dispose(); deepSearchBoldFont = null; }
            if (workProfileActiveFont != null) { workProfileActiveFont.Dispose(); workProfileActiveFont = null; }
            if (activeDevelopmentTaskFont != null) { activeDevelopmentTaskFont.Dispose(); activeDevelopmentTaskFont = null; }
            if (trayIcon != null) { trayIcon.Visible = false; trayIcon.Dispose(); trayIcon = null; }
            SavePreferences();
            SaveWindowBounds();
            SaveLastPath();
            SaveNavigationHistory();
            SaveTransferPanePaths();
            SaveSplitterDistance(split.SplitterDistance);
            SaveMaximizedState(WindowState == FormWindowState.Maximized);
            SaveColumnLayouts();
        };
        LoadShortcuts();
        LoadHistory();
        try { if (File.Exists(deepSearchResultsFile)) File.Delete(deepSearchResultsFile); } catch { }
        LoadWorkProfiles();
        RefreshWorkProfileList();
        LoadDevelopmentBoard();
        StartDevelopmentSession();
        Navigate(currentPath, null, null, false);
        StartOneDriveStatusPolling();
        UpdateNavigationButtons();
        ApplyTheme();
    }

    private void BeginInlineRename(ListViewItem item)
    {
        if (item == null || showingComputerDrives || !files.LabelEdit) return;
        int itemIndex = item.Index;
        BeginInvoke((MethodInvoker)(() =>
        {
            if (showingComputerDrives || itemIndex < 0 || itemIndex >= files.Items.Count) return;
            ListViewItem editedItem = files.Items[itemIndex];
            Entry entry = editedItem.Tag as Entry;
            if (entry == null || (!File.Exists(entry.FullPath) && !Directory.Exists(entry.FullPath))) return;
            allowInlineRenameStart = true;
            try
            {
                editedItem.BeginEdit();
                string oldName = Path.GetFileName(entry.FullPath);
                int extensionLength = Directory.Exists(entry.FullPath) ? 0 : Path.GetExtension(oldName).Length;
                int baseNameLength = Math.Max(0, oldName.Length - extensionLength);
                if (baseNameLength > 0)
                {
                    IntPtr editHandle = SendMessage(files.Handle, 0x1018, IntPtr.Zero, IntPtr.Zero);
                    if (editHandle != IntPtr.Zero) SendMessage(editHandle, 0x00B1, IntPtr.Zero, new IntPtr(baseNameLength));
                }
            }
            finally { allowInlineRenameStart = false; }
        }));
    }

    private void ResumeInlineRename(int itemIndex, string editedName)
    {
        if (itemIndex < 0) return;
        BeginInvoke((MethodInvoker)(() =>
        {
            if (IsDisposed || files == null || files.IsDisposed || showingComputerDrives || itemIndex >= files.Items.Count) return;
            ListViewItem item = files.Items[itemIndex];
            item.Selected = true;
            files.FocusedItem = item;
            allowInlineRenameStart = true;
            try { item.BeginEdit(); }
            finally { allowInlineRenameStart = false; }

            IntPtr editHandle = SendMessage(files.Handle, 0x1018, IntPtr.Zero, IntPtr.Zero);
            if (editHandle == IntPtr.Zero) return;
            IntPtr textPointer = IntPtr.Zero;
            string nameToRestore = editedName ?? String.Empty;
            try
            {
                textPointer = Marshal.StringToHGlobalUni(nameToRestore);
                SendMessage(editHandle, 0x000C, IntPtr.Zero, textPointer);
                int caret = nameToRestore.Length;
                SendMessage(editHandle, 0x00B1, new IntPtr(caret), new IntPtr(caret));
            }
            finally
            {
                if (textPointer != IntPtr.Zero) Marshal.FreeHGlobal(textPointer);
            }
        }));
    }

    private static bool IsSharingViolation(Exception exception)
    {
        if (!(exception is IOException)) return false;
        int win32Error = Marshal.GetHRForException(exception) & 0xFFFF;
        return win32Error == 32 || win32Error == 33;
    }

    private void InitializeTransferPage(TabPage page)
    {
        string leftDefault = Directory.Exists(currentPath) ? currentPath : Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        string rightDefault = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (!Directory.Exists(rightDefault)) rightDefault = leftDefault;
        try
        {
            if (File.Exists(transferPanePathsFile))
            {
                string[] saved = File.ReadAllLines(transferPanePathsFile).Select(ResolvePortablePath).ToArray();
                if (saved.Length > 0 && Directory.Exists(saved[0])) leftDefault = saved[0];
                if (saved.Length > 1 && Directory.Exists(saved[1])) rightDefault = saved[1];
                transferLeftPaneHidden = saved.Length > 2 && saved[2] == "1";
                transferRightPaneHidden = saved.Length > 3 && saved[3] == "1";
            }
        }
        catch { }

        transferLeftPane = new TransferPaneState(0, "Sol panel", leftDefault);
        transferRightPane = new TransferPaneState(1, "Sağ panel", rightDefault);
        ApplySavedSortPreference("Çiftli İşlem Sol", transferLeftPane);
        ApplySavedSortPreference("Çiftli İşlem Sağ", transferRightPane);
        transferLeftPane.FolderPaneHidden = transferLeftPaneHidden;
        transferRightPane.FolderPaneHidden = transferRightPaneHidden;
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(12, 8, 12, 6) };
        var modeLabel = new Label { Text = "Sürükle bırak:", AutoSize = true, Location = new Point(14, 15) };
        transferCopyMode.Text = "Kopyala";
        transferCopyMode.AutoSize = true;
        transferCopyMode.Checked = false;
        transferCopyMode.Location = new Point(112, 13);
        transferMoveMode.Text = "Taşı";
        transferMoveMode.AutoSize = true;
        transferMoveMode.Checked = true;
        transferMoveMode.Location = new Point(195, 13);
        actionTooltips.SetToolTip(transferCopyMode, "Paneller arasında sürüklenen dosyalar kopyalanır.");
        actionTooltips.SetToolTip(transferMoveMode, "Paneller arasında sürüklenen dosyalar taşınır.");
        var help = new Label
        {
            Text = "Birden fazla dosya seçebilirsiniz. Klasöre tek tıkla girin; klasöre sağ tıklayıp yapıştırın veya sürükleyip bırakın.",
            AutoSize = true,
            Location = new Point(270, 15),
            ForeColor = Color.FromArgb(82, 94, 112)
        };
        toolbar.Controls.Add(modeLabel);
        toolbar.Controls.Add(transferCopyMode);
        toolbar.Controls.Add(transferMoveMode);
        toolbar.Controls.Add(help);

        var sides = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Size = new Size(1200, 700),
            Orientation = Orientation.Vertical,
            SplitterDistance = 600,
            Panel1MinSize = 260,
            Panel2MinSize = 260,
            BackColor = Color.FromArgb(226, 231, 238),
            BorderStyle = BorderStyle.FixedSingle
        };
        sides.Panel1.Controls.Add(CreateTransferPane(transferLeftPane));
        sides.Panel2.Controls.Add(CreateTransferPane(transferRightPane));
        transferStatus.Text = "Sol ve sağ panel için klasörleri ayrı ayrı seçin. Dosyaları paneller arasında sürükleyerek aktarın.";
        transferStatus.Dock = DockStyle.Bottom;
        transferStatus.Height = 28;
        transferStatus.Padding = new Padding(12, 5, 0, 0);
        transferStatus.BackColor = Color.White;
        page.Controls.Add(sides);
        page.Controls.Add(transferStatus);
        page.Controls.Add(toolbar);
        TransferNavigate(transferLeftPane, leftDefault);
        TransferNavigate(transferRightPane, rightDefault);
    }

    private void InitializeSyncPage(TabPage page)
    {
        page.BackColor = Color.FromArgb(245, 248, 252);
        syncPageContent = new Panel { Dock = DockStyle.Fill, BackColor = page.BackColor };
        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 48,
            Padding = new Padding(12, 8, 8, 4),
            WrapContents = false,
            AutoScroll = true,
            BackColor = Color.White
        };
        var title = new Label { Text = "Senkronize", AutoSize = false, Width = 118, Height = 30, Font = new Font("Segoe UI", 11F, FontStyle.Bold), ForeColor = Color.FromArgb(35, 53, 75), TextAlign = ContentAlignment.MiddleLeft };
        var profileLabel = new Label { Text = "Profil:", AutoSize = false, Width = 42, Height = 30, TextAlign = ContentAlignment.MiddleRight };
        syncProfileSelector.DropDownStyle = ComboBoxStyle.DropDownList;
        syncProfileSelector.Width = 220;
        syncProfileSelector.Height = 28;
        syncProfileSelector.SelectedIndexChanged += (s, e) => { if (!bindingSyncProfile) LoadSelectedSyncProfile(); };
        var newProfile = new Button { Text = "Yeni profil", Width = 86, Height = 28, Margin = new Padding(6, 1, 2, 0) };
        var saveProfile = new Button { Text = "Kaydet", Width = 70, Height = 28, Margin = new Padding(2, 1, 2, 0) };
        var deleteProfile = new Button { Text = "Profili sil", Width = 78, Height = 28, Margin = new Padding(2, 1, 2, 0) };
        syncFullScreenButton.Text = "Tam ekran";
        syncFullScreenButton.Width = 86;
        syncFullScreenButton.Height = 28;
        syncFullScreenButton.Margin = new Padding(8, 1, 2, 0);
        syncFullScreenButton.Click += (s, e) => ToggleSyncFullScreen(!syncFullScreen);
        newProfile.Click += (s, e) => CreateSyncProfile();
        saveProfile.Click += (s, e) => SaveSelectedSyncProfile();
        deleteProfile.Click += (s, e) => DeleteSelectedSyncProfile();
        toolbar.Controls.Add(title);
        toolbar.Controls.Add(profileLabel);
        toolbar.Controls.Add(syncProfileSelector);
        toolbar.Controls.Add(newProfile);
        toolbar.Controls.Add(saveProfile);
        toolbar.Controls.Add(deleteProfile);
        toolbar.Controls.Add(syncFullScreenButton);

        var pairPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 92,
            Padding = new Padding(12, 8, 12, 4),
            ColumnCount = 2,
            RowCount = 2,
            BackColor = Color.FromArgb(248, 250, 253)
        };
        pairPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        pairPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        pairPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        pairPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var leftLabel = new Label { Text = "SOL KLASÖR", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(39, 92, 146), TextAlign = ContentAlignment.MiddleLeft };
        var rightLabel = new Label { Text = "SAĞ KLASÖR", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(39, 92, 146), TextAlign = ContentAlignment.MiddleLeft };
        pairPanel.Controls.Add(leftLabel, 0, 0);
        pairPanel.Controls.Add(rightLabel, 1, 0);
        pairPanel.Controls.Add(CreateSyncPathEditor(syncLeftPathBox, true), 0, 1);
        pairPanel.Controls.Add(CreateSyncPathEditor(syncRightPathBox, false), 1, 1);

        var settingsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 54,
            Padding = new Padding(12, 7, 8, 2),
            WrapContents = false,
            AutoScroll = true,
            BackColor = Color.White
        };
        var comparisonLabel = new Label { Text = "Karşılaştır:", AutoSize = false, Width = 76, Height = 30, TextAlign = ContentAlignment.MiddleLeft };
        syncComparisonSelector.DropDownStyle = ComboBoxStyle.DropDownList;
        syncComparisonSelector.Width = 235;
        syncComparisonSelector.Items.Add("Dosya boyutu ve değiştirme zamanı");
        syncComparisonSelector.Items.Add("Dosya içeriği (SHA-256)");
        syncComparisonSelector.SelectedIndexChanged += (s, e) => { if (!bindingSyncProfile) syncProfilePendingSave = true; };
        syncSubfoldersOption.Text = "Alt klasörleri tara";
        syncSubfoldersOption.AutoSize = true;
        syncSubfoldersOption.Checked = true;
        syncSubfoldersOption.Margin = new Padding(12, 6, 4, 0);
        syncSubfoldersOption.CheckedChanged += (s, e) => { if (!bindingSyncProfile) syncProfilePendingSave = true; };
        syncShowIdenticalOption.Text = "Aynı dosyaları da göster";
        syncShowIdenticalOption.AutoSize = true;
        syncShowIdenticalOption.Margin = new Padding(8, 6, 4, 0);
        syncShowIdenticalOption.CheckedChanged += (s, e) => { if (!bindingSyncProfile) syncProfilePendingSave = true; };
        settingsPanel.Controls.Add(comparisonLabel);
        settingsPanel.Controls.Add(syncComparisonSelector);
        settingsPanel.Controls.Add(syncSubfoldersOption);
        settingsPanel.Controls.Add(syncShowIdenticalOption);

        var actionPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 48,
            Padding = new Padding(12, 6, 8, 4),
            WrapContents = false,
            AutoScroll = true,
            BackColor = Color.FromArgb(242, 246, 251)
        };
        syncCompareButton.Text = "⟳  Karşılaştır";
        syncCompareButton.Width = 126;
        syncCompareButton.Height = 30;
        syncCompareButton.Click += (s, e) => StartSyncComparison();
        syncCancelButton.Text = "Durdur";
        syncCancelButton.Width = 80;
        syncCancelButton.Height = 30;
        syncCancelButton.Enabled = false;
        syncCancelButton.Click += (s, e) =>
        {
            if (syncComparisonWorker != null && syncComparisonWorker.IsBusy) syncComparisonWorker.CancelAsync();
            if (syncLookupWorker != null && syncLookupWorker.IsBusy) syncLookupWorker.CancelAsync();
        };
        syncApplyButton.Text = "Eşitle";
        syncApplyButton.Width = 94;
        syncApplyButton.Height = 30;
        syncApplyButton.Enabled = false;
        syncApplyButton.Click += (s, e) => MessageBox.Show(this,
            "Bu ilk aşamada klasör profilleri ve karşılaştırma ön izlemesi hazır. Dosya kopyalama, silme, çakışma çözümü ve geri alma kurallarını birlikte belirledikten sonra eşitleme motorunu etkinleştireceğiz.",
            "Senkronizasyon", MessageBoxButtons.OK, MessageBoxIcon.Information);
        syncLookupButton.Text = "Karşı tarafta ara";
        syncLookupButton.Width = 145;
        syncLookupButton.Height = 30;
        syncLookupButton.Click += (s, e) => StartSyncLookup();
        actionTooltips.SetToolTip(syncLookupButton, "Yalnızca sol/sağ durumundaki dosya adlarını karşı taraftaki ana klasör ve tüm alt klasörlerde arar. Ek klasör seçmeniz gerekmez.");
        syncExportCsvButton.Text = "CSV dışa aktar";
        syncExportCsvButton.Width = 120;
        syncExportCsvButton.Height = 30;
        syncExportCsvButton.Click += (s, e) => ExportSyncComparisonCsv();
        var safetyNote = new Label
        {
            Text = "Ön izleme güvenlidir; bulut dosyaları da taranır.",
            AutoSize = true,
            Height = 30,
            Margin = new Padding(12, 7, 0, 0),
            ForeColor = Color.FromArgb(78, 91, 109)
        };
        actionPanel.Controls.Add(syncCompareButton);
        actionPanel.Controls.Add(syncCancelButton);
        actionPanel.Controls.Add(syncApplyButton);
        actionPanel.Controls.Add(syncLookupButton);
        actionPanel.Controls.Add(syncExportCsvButton);
        actionPanel.Controls.Add(safetyNote);
        actionTooltips.SetToolTip(safetyNote, "OneDrive çevrimiçi yer tutucuları karşılaştırmaya dahil edilir. Dosya içeriği indirilmeden boyut ve tarih bilgileri okunur. SHA-256 seçilirse, çevrimiçi dosya indirilmeden bırakılır ve içerik kontrolü bekliyor olarak gösterilir.");

        syncComparisonList.View = View.Details;
        syncComparisonList.FullRowSelect = true;
        syncComparisonList.GridLines = true;
        syncComparisonList.HideSelection = false;
        syncComparisonList.MultiSelect = false;
        syncComparisonList.Columns.Add("Durum", 145);
        syncComparisonList.Columns.Add("Göreli yol", 360);
        syncComparisonList.Columns.Add("Sol taraftaki dosya", 300);
        syncComparisonList.Columns.Add("Sağ taraftaki dosya", 300);
        syncComparisonList.Columns.Add("Boyut", 95);
        syncComparisonList.Columns.Add("Bulut durumu", 142);
        syncComparisonList.Columns.Add("Karşı klasörde", 95);
        syncComparisonList.Columns.Add("Bulunan dosyanın yolu", 240);
        syncComparisonList.Columns.Add("Bulunan dosya boyutu", 135);
        syncComparisonList.Dock = DockStyle.Fill;
        syncComparisonList.Resize += (s, e) => ResizeSyncComparisonColumns();
        actionTooltips.SetToolTip(syncComparisonSelector, "İlk yöntem hızlıdır; aynı boyut ve yakın değiştirme zamanı olanları aynı sayar. SHA-256 yöntemi dosya içeriğini karşılaştırır ve büyük klasörlerde daha uzun sürer.");
        actionTooltips.SetToolTip(syncApplyButton, "İlk aşamada kapalıdır. Önce karşılaştırma, çakışma ve geri alma kurallarını belirleyeceğiz.");

        syncStatus.Text = "İki klasör seçip Karşılaştır'a basın. Bu aşamada hiçbir dosya değiştirilmez veya silinmez.";
        syncStatus.Dock = DockStyle.Bottom;
        syncStatus.Height = 30;
        syncStatus.Padding = new Padding(14, 6, 0, 0);
        syncStatus.BackColor = Color.White;
        syncStatus.ForeColor = Color.FromArgb(77, 91, 109);

        syncPageContent.Controls.Add(syncComparisonList);
        syncPageContent.Controls.Add(syncStatus);
        syncPageContent.Controls.Add(actionPanel);
        syncPageContent.Controls.Add(settingsPanel);
        syncPageContent.Controls.Add(pairPanel);
        syncPageContent.Controls.Add(toolbar);
        page.Controls.Add(syncPageContent);

        syncComparisonWorker = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
        syncComparisonWorker.DoWork += (s, e) =>
        {
            var worker = (BackgroundWorker)s;
            e.Result = CompareSyncFolders((SyncComparisonRequest)e.Argument, worker);
            if (worker.CancellationPending) e.Cancel = true;
        };
        syncComparisonWorker.ProgressChanged += (s, e) => syncStatus.Text = Convert.ToString(e.UserState, CultureInfo.CurrentCulture);
        syncComparisonWorker.RunWorkerCompleted += (s, e) =>
        {
            syncCompareButton.Enabled = true;
            syncCancelButton.Enabled = false;
            if (e.Error != null)
            {
                syncStatus.Text = "Karşılaştırma tamamlanamadı: " + e.Error.GetBaseException().Message;
                MessageBox.Show(this, e.Error.ToString(), "Senkronizasyon karşılaştırma hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (e.Cancelled) { syncStatus.Text = "Karşılaştırma durduruldu."; return; }
            ShowSyncComparison((SyncComparisonResult)e.Result);
        };

        syncLookupWorker = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
        syncLookupWorker.DoWork += (s, e) =>
        {
            var worker = (BackgroundWorker)s;
            var request = (SyncLookupRequest)e.Argument;
            var leftScan = new SyncComparisonResult();
            var rightScan = new SyncComparisonResult();
            var leftFiles = new Dictionary<string, SyncFileMetadata>(StringComparer.OrdinalIgnoreCase);
            var rightFiles = new Dictionary<string, SyncFileMetadata>(StringComparer.OrdinalIgnoreCase);
            if (request.Rows.Any(row => row.Category == "Yalnızca sağ"))
                leftFiles = IndexSyncFolder(request.LeftRoot, true, "Sol karşı taraf", leftScan, worker);
            if (worker.CancellationPending) { e.Cancel = true; return; }
            if (request.Rows.Any(row => row.Category == "Yalnızca sol"))
                rightFiles = IndexSyncFolder(request.RightRoot, true, "Sağ karşı taraf", rightScan, worker);
            e.Result = new SyncLookupResult
            {
                Rows = request.Rows,
                LeftRoot = request.LeftRoot,
                RightRoot = request.RightRoot,
                LeftFilesByName = BuildSyncFileNameIndex(leftFiles),
                RightFilesByName = BuildSyncFileNameIndex(rightFiles),
                Errors = leftScan.Errors + rightScan.Errors,
                ErrorDetails = leftScan.ErrorDetails.Concat(rightScan.ErrorDetails).ToList()
            };
            if (worker.CancellationPending) e.Cancel = true;
        };
        syncLookupWorker.ProgressChanged += (s, e) => syncStatus.Text = Convert.ToString(e.UserState, CultureInfo.CurrentCulture);
        syncLookupWorker.RunWorkerCompleted += (s, e) =>
        {
            syncLookupButton.Enabled = true;
            syncCancelButton.Enabled = false;
            syncCompareButton.Enabled = true;
            if (e.Error != null)
            {
                syncStatus.Text = "Karşı klasörlerde arama tamamlanamadı: " + e.Error.GetBaseException().Message;
                MessageBox.Show(this, e.Error.ToString(), "Karşı klasör arama hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (e.Cancelled)
            {
                foreach (SyncComparisonRow row in syncCurrentRows.Where(item => item.LookupStatus == "Aranıyor…")) row.LookupStatus = "Durduruldu";
                RefreshSyncLookupCells(syncCurrentRows);
                syncStatus.Text = "Karşı klasörlerdeki arama durduruldu.";
                return;
            }
            ShowSyncLookupResults((SyncLookupResult)e.Result);
        };

        tabs.SelectedIndexChanged += (s, e) =>
        {
            if (syncProfilePendingSave && tabs.SelectedTab != syncPage) SaveSelectedSyncProfile();
        };

        LoadSyncProfiles();
        if (syncComparisonSelector.Items.Count > 0) syncComparisonSelector.SelectedIndex = 0;
        ResizeSyncComparisonColumns();
    }

    private void ResizeSyncComparisonColumns()
    {
        if (syncComparisonList.Columns.Count < 9) return;
        int width = Math.Max(480, syncComparisonList.ClientSize.Width - 8);
        int fixedWidth = 125 + 75 + 130 + 95 + 240 + 135;
        int flexible = Math.Max(270, width - fixedWidth);
        int relative = (int)(flexible * 0.38);
        int eachSide = Math.Max(95, (flexible - relative) / 2);
        syncComparisonList.Columns[0].Width = 125;
        syncComparisonList.Columns[1].Width = relative;
        syncComparisonList.Columns[2].Width = eachSide;
        syncComparisonList.Columns[3].Width = eachSide;
        syncComparisonList.Columns[4].Width = 75;
        syncComparisonList.Columns[5].Width = 130;
        syncComparisonList.Columns[6].Width = 95;
        syncComparisonList.Columns[7].Width = 240;
        syncComparisonList.Columns[8].Width = 135;
    }

    private Control CreateSyncPathEditor(TextBox pathBox, bool left)
    {
        var row = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 2, 8, 2) };
        pathBox.Dock = DockStyle.Fill;
        pathBox.ReadOnly = false;
        pathBox.BorderStyle = BorderStyle.FixedSingle;
        pathBox.TextChanged += (s, e) => { if (!bindingSyncProfile) syncProfilePendingSave = true; };
        pathBox.ContextMenuStrip = CreateSyncPathContextMenu(pathBox);
        var browse = new Button { Text = "Göz at…", Dock = DockStyle.Right, Width = 88 };
        browse.Click += (s, e) =>
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = left ? "Senkronizasyonun sol klasörünü seçin" : "Senkronizasyonun sağ klasörünü seçin";
                string current = pathBox.Text;
                if (Directory.Exists(current)) dialog.SelectedPath = current;
                if (dialog.ShowDialog(this) == DialogResult.OK && Directory.Exists(dialog.SelectedPath)) pathBox.Text = dialog.SelectedPath;
            }
        };
        row.Controls.Add(pathBox);
        row.Controls.Add(browse);
        return row;
    }

    private ContextMenuStrip CreateSyncPathContextMenu(TextBox box)
    {
        var menu = new ContextMenuStrip();
        var cut = new ToolStripMenuItem("Kes");
        var copy = new ToolStripMenuItem("Kopyala");
        var paste = new ToolStripMenuItem("Yapıştır");
        var selectAll = new ToolStripMenuItem("Tümünü seç");
        cut.Click += (s, e) => { if (box.SelectionLength > 0) box.Cut(); };
        copy.Click += (s, e) => { if (box.SelectionLength > 0) box.Copy(); };
        paste.Click += (s, e) => box.Paste();
        selectAll.Click += (s, e) => box.SelectAll();
        menu.Items.Add(cut);
        menu.Items.Add(copy);
        menu.Items.Add(paste);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(selectAll);
        menu.Opening += (s, e) =>
        {
            bool selected = box.SelectionLength > 0;
            cut.Enabled = selected;
            copy.Enabled = selected;
            try { paste.Enabled = Clipboard.ContainsText(); } catch { paste.Enabled = false; }
        };
        return menu;
    }

    private void LoadSyncProfiles()
    {
        syncProfiles.Clear();
        string selectedProfileName = String.Empty;
        try
        {
            if (File.Exists(syncProfilesFile))
            {
                var document = new XmlDocument();
                document.Load(syncProfilesFile);
                if (document.DocumentElement != null) selectedProfileName = document.DocumentElement.GetAttribute("selected");
                foreach (XmlNode node in document.SelectNodes("/syncProfiles/profile"))
                {
                    int comparisonMode = 0;
                    if (node.Attributes["comparison"] != null) Int32.TryParse(node.Attributes["comparison"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out comparisonMode);
                    var profile = new SyncWorkspaceProfile
                    {
                        Name = node.Attributes["name"] == null ? "Senkron profili" : node.Attributes["name"].Value,
                        LeftPath = ResolvePortablePath(node.Attributes["left"] == null ? String.Empty : node.Attributes["left"].Value),
                        RightPath = ResolvePortablePath(node.Attributes["right"] == null ? String.Empty : node.Attributes["right"].Value),
                        ComparisonMode = comparisonMode,
                        IncludeSubfolders = node.Attributes["subfolders"] == null || node.Attributes["subfolders"].Value != "0",
                        ShowIdentical = node.Attributes["showIdentical"] != null && node.Attributes["showIdentical"].Value == "1"
                    };
                    if (!String.IsNullOrWhiteSpace(profile.Name)) syncProfiles.Add(profile);
                }
            }
        }
        catch (Exception ex) { syncStatus.Text = "Senkron profilleri okunamadı: " + ex.Message; }
        if (syncProfiles.Count == 0)
        {
            syncProfiles.Add(CreateDefaultSyncProfile());
        }
        bindingSyncProfile = true;
        syncProfileSelector.Items.Clear();
        foreach (SyncWorkspaceProfile profile in syncProfiles) syncProfileSelector.Items.Add(profile);
        if (syncProfileSelector.Items.Count > 0)
        {
            int selectedIndex = syncProfiles.FindIndex(profile => String.Equals(profile.Name, selectedProfileName, StringComparison.CurrentCultureIgnoreCase));
            syncProfileSelector.SelectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
        }
        bindingSyncProfile = false;
        LoadSelectedSyncProfile();
    }

    private void LoadSelectedSyncProfile()
    {
        SyncWorkspaceProfile profile = syncProfileSelector.SelectedItem as SyncWorkspaceProfile;
        if (profile == null) return;
        if (!bindingSyncProfile && activeSyncProfile != null && !Object.ReferenceEquals(activeSyncProfile, profile))
        {
            UpdateSyncProfileFromControls(activeSyncProfile);
            SaveSyncProfiles();
        }
        activeSyncProfile = profile;
        syncLeftPathBox.Text = profile.LeftPath ?? String.Empty;
        syncRightPathBox.Text = profile.RightPath ?? String.Empty;
        syncSubfoldersOption.Checked = profile.IncludeSubfolders;
        syncShowIdenticalOption.Checked = profile.ShowIdentical;
        if (profile.ComparisonMode >= 0 && profile.ComparisonMode < syncComparisonSelector.Items.Count)
            syncComparisonSelector.SelectedIndex = profile.ComparisonMode;
        syncComparisonList.Items.Clear();
        syncCurrentRows.Clear();
        syncProfilePendingSave = false;
    }

    private void CreateSyncProfile()
    {
        UpdateSyncProfileFromControls(syncProfileSelector.SelectedItem as SyncWorkspaceProfile);
        string name = PromptForSyncProfileName("Yeni senkron profili", "Yeni profil");
        if (String.IsNullOrWhiteSpace(name)) return;
        if (syncProfiles.Any(existingProfile => String.Equals(existingProfile.Name, name.Trim(), StringComparison.CurrentCultureIgnoreCase)))
        {
            MessageBox.Show(this, "Bu adla bir profil zaten var.", "Senkronize", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var profile = new SyncWorkspaceProfile
        {
            Name = name.Trim(), LeftPath = syncLeftPathBox.Text, RightPath = syncRightPathBox.Text,
            ComparisonMode = Math.Max(0, syncComparisonSelector.SelectedIndex),
            IncludeSubfolders = syncSubfoldersOption.Checked, ShowIdentical = syncShowIdenticalOption.Checked
        };
        syncProfiles.Add(profile);
        bindingSyncProfile = true;
        syncProfileSelector.Items.Add(profile);
        syncProfileSelector.SelectedItem = profile;
        bindingSyncProfile = false;
        activeSyncProfile = profile;
        syncProfilePendingSave = false;
        SaveSyncProfiles();
    }

    private string PromptForSyncProfileName(string title, string initialValue)
    {
        using (var dialog = new Form { Text = title, StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false, ClientSize = new Size(390, 116) })
        {
            var box = new TextBox { Text = initialValue, Left = 14, Top = 18, Width = 360, Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right };
            var ok = new Button { Text = "Tamam", DialogResult = DialogResult.OK, Left = 202, Top = 62, Width = 80 };
            var cancel = new Button { Text = "İptal", DialogResult = DialogResult.Cancel, Left = 294, Top = 62, Width = 80 };
            dialog.Controls.Add(box); dialog.Controls.Add(ok); dialog.Controls.Add(cancel);
            dialog.AcceptButton = ok; dialog.CancelButton = cancel;
            dialog.Shown += (s, e) => { box.SelectAll(); box.Focus(); };
            return dialog.ShowDialog(this) == DialogResult.OK ? box.Text.Trim() : String.Empty;
        }
    }

    private void SaveSelectedSyncProfile()
    {
        SyncWorkspaceProfile profile = syncProfileSelector.SelectedItem as SyncWorkspaceProfile;
        if (profile == null) return;
        UpdateSyncProfileFromControls(profile);
        SaveSyncProfiles();
        syncProfilePendingSave = false;
        syncStatus.Text = "Profil kaydedildi: " + profile.Name;
    }

    private void UpdateSyncProfileFromControls(SyncWorkspaceProfile profile)
    {
        if (profile == null || syncLeftPathBox == null || syncRightPathBox == null) return;
        profile.LeftPath = syncLeftPathBox.Text.Trim().Trim('"');
        profile.RightPath = syncRightPathBox.Text.Trim().Trim('"');
        profile.ComparisonMode = Math.Max(0, syncComparisonSelector.SelectedIndex);
        profile.IncludeSubfolders = syncSubfoldersOption.Checked;
        profile.ShowIdentical = syncShowIdenticalOption.Checked;
    }

    private void DeleteSelectedSyncProfile()
    {
        SyncWorkspaceProfile profile = syncProfileSelector.SelectedItem as SyncWorkspaceProfile;
        if (profile == null) return;
        if (MessageBox.Show(this, "'" + profile.Name + "' profili silinsin mi?", "Profili sil", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        syncProfiles.Remove(profile);
        if (Object.ReferenceEquals(activeSyncProfile, profile)) activeSyncProfile = null;
        if (syncProfiles.Count == 0) syncProfiles.Add(CreateDefaultSyncProfile());
        bindingSyncProfile = true;
        syncProfileSelector.Items.Clear();
        foreach (SyncWorkspaceProfile item in syncProfiles) syncProfileSelector.Items.Add(item);
        syncProfileSelector.SelectedIndex = 0;
        bindingSyncProfile = false;
        LoadSelectedSyncProfile();
        SaveSyncProfiles();
    }

    private SyncWorkspaceProfile CreateDefaultSyncProfile()
    {
        string left = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string right = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (!Directory.Exists(left)) left = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!Directory.Exists(right) || SyncPathsOverlap(left, right))
        {
            string candidate = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            if (Directory.Exists(candidate) && !SyncPathsOverlap(left, candidate)) right = candidate;
            else right = String.Empty;
        }
        return new SyncWorkspaceProfile { Name = "Yeni profil", LeftPath = left, RightPath = right, IncludeSubfolders = true };
    }

    private static bool SyncPathsOverlap(string left, string right)
    {
        if (String.IsNullOrWhiteSpace(left) || String.IsNullOrWhiteSpace(right)) return false;
        try
        {
            string first = Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string second = Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return String.Equals(first, second, StringComparison.OrdinalIgnoreCase) || IsPathWithin(first, second) || IsPathWithin(second, first);
        }
        catch { return String.Equals(left, right, StringComparison.OrdinalIgnoreCase); }
    }

    private void SaveSyncProfiles()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(syncProfilesFile));
            var document = new XmlDocument();
            XmlElement root = document.CreateElement("syncProfiles");
            SyncWorkspaceProfile selected = syncProfileSelector.SelectedItem as SyncWorkspaceProfile;
            if (selected != null) root.SetAttribute("selected", selected.Name ?? String.Empty);
            document.AppendChild(root);
            foreach (SyncWorkspaceProfile profile in syncProfiles)
            {
                XmlElement node = document.CreateElement("profile");
                node.SetAttribute("name", profile.Name ?? String.Empty);
                node.SetAttribute("left", StorePortablePath(profile.LeftPath ?? String.Empty));
                node.SetAttribute("right", StorePortablePath(profile.RightPath ?? String.Empty));
                node.SetAttribute("comparison", profile.ComparisonMode.ToString(CultureInfo.InvariantCulture));
                node.SetAttribute("subfolders", profile.IncludeSubfolders ? "1" : "0");
                node.SetAttribute("showIdentical", profile.ShowIdentical ? "1" : "0");
                root.AppendChild(node);
            }
            document.Save(syncProfilesFile);
        }
        catch (Exception ex) { MessageBox.Show(this, "Senkron profilleri kaydedilemedi: " + ex.Message, "Senkronize", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }

    private void StartSyncComparison()
    {
        if (syncComparisonWorker == null || syncComparisonWorker.IsBusy || (syncLookupWorker != null && syncLookupWorker.IsBusy)) return;
        string left = syncLeftPathBox.Text.Trim().Trim('"');
        string right = syncRightPathBox.Text.Trim().Trim('"');
        if (!Directory.Exists(left) || !Directory.Exists(right))
        {
            MessageBox.Show(this, "Karşılaştırmak için sol ve sağ klasörlerin ikisi de mevcut olmalı.", "Senkronize", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        string leftFull = Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string rightFull = Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (SyncPathsOverlap(leftFull, rightFull))
        {
            MessageBox.Show(this, "Sol ve sağ klasör aynı olamaz; ayrıca biri diğerinin alt klasörü de olmamalı.", "Senkronize", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var request = new SyncComparisonRequest
        {
            LeftPath = leftFull, RightPath = rightFull,
            IncludeSubfolders = syncSubfoldersOption.Checked,
            ShowIdentical = syncShowIdenticalOption.Checked,
            CompareContent = syncComparisonSelector.SelectedIndex == 1
        };
        syncComparisonList.Items.Clear();
        syncCurrentRows = new List<SyncComparisonRow>();
        syncCompareButton.Enabled = false;
        syncCancelButton.Enabled = true;
        syncStatus.Text = "Klasörlerdeki dosyalar listeleniyor...";
        syncComparisonWorker.RunWorkerAsync(request);
    }

    private SyncComparisonResult CompareSyncFolders(SyncComparisonRequest request, BackgroundWorker worker)
    {
        var result = new SyncComparisonResult();
        Dictionary<string, SyncFileMetadata> left = IndexSyncFolder(request.LeftPath, request.IncludeSubfolders, "Sol", result, worker);
        if (worker.CancellationPending) return result;
        Dictionary<string, SyncFileMetadata> right = IndexSyncFolder(request.RightPath, request.IncludeSubfolders, "Sağ", result, worker);
        if (worker.CancellationPending) return result;
        var allKeys = new HashSet<string>(left.Keys, StringComparer.OrdinalIgnoreCase);
        allKeys.UnionWith(right.Keys);
        int index = 0;
        foreach (string relative in allKeys.OrderBy(path => path, StringComparer.CurrentCultureIgnoreCase))
        {
            if (worker.CancellationPending) return result;
            SyncFileMetadata leftFile, rightFile;
            bool hasLeft = left.TryGetValue(relative, out leftFile);
            bool hasRight = right.TryGetValue(relative, out rightFile);
            string category;
            string cloudStatus = GetSyncCloudStatus(leftFile, hasLeft, rightFile, hasRight);
            if (hasLeft && !hasRight) category = "Yalnızca sol";
            else if (!hasLeft && hasRight) category = "Yalnızca sağ";
            else
            {
                bool metadataSame = leftFile.Length == rightFile.Length && Math.Abs((leftFile.LastWriteTimeUtc - rightFile.LastWriteTimeUtc).TotalSeconds) <= 2;
                bool eitherOnlineOnly = leftFile.IsOnlineOnly || rightFile.IsOnlineOnly;
                bool same;
                category = String.Empty;
                try
                {
                    if (request.CompareContent && eitherOnlineOnly)
                    {
                        category = "Bulut içeriği bekliyor";
                        result.CloudContentDeferredCount++;
                        if (!metadataSame) cloudStatus += " · meta farklı";
                        same = false;
                    }
                    else same = request.CompareContent ? leftFile.Length == rightFile.Length && FilesHaveSameHash(leftFile.FullPath, rightFile.FullPath) : metadataSame;
                }
                catch (Exception ex)
                {
                    result.Errors++;
                    result.ErrorDetails.Add(relative + " — " + ex.Message);
                    category = "Okuma hatası";
                    result.Rows.Add(new SyncComparisonRow { Category = category, RelativePath = relative, LeftPath = leftFile.FullPath, RightPath = rightFile.FullPath, SizeText = FormatFileSize(leftFile.Length), CloudStatus = cloudStatus });
                    continue;
                }
                if (category != "Bulut içeriği bekliyor") category = same ? "Aynı" : "İki tarafta farklı";
                if (same) result.IdenticalCount++;
            }
            if (category == "Yalnızca sol") result.LeftOnlyCount++;
            else if (category == "Yalnızca sağ") result.RightOnlyCount++;
            else if (category == "İki tarafta farklı") result.DifferentCount++;
            if (category != "Aynı" || request.ShowIdentical)
                result.Rows.Add(new SyncComparisonRow { Category = category, RelativePath = relative, LeftPath = hasLeft ? leftFile.FullPath : String.Empty, RightPath = hasRight ? rightFile.FullPath : String.Empty, SizeText = FormatFileSize(hasLeft ? leftFile.Length : rightFile.Length), CloudStatus = cloudStatus });
            index++;
            if (index % 500 == 0) worker.ReportProgress(0, "Karşılaştırılıyor: " + index.ToString("N0", CultureInfo.CurrentCulture) + " göreli dosya yolu...");
        }
        return result;
    }

    private Dictionary<string, SyncFileMetadata> IndexSyncFolder(string root, bool recursive, string side, SyncComparisonResult result, BackgroundWorker worker)
    {
        var filesByRelativePath = new Dictionary<string, SyncFileMetadata>(StringComparer.OrdinalIgnoreCase);
        var pending = new Stack<string>();
        pending.Push(root);
        int scanned = 0;
        while (pending.Count > 0)
        {
            if (worker.CancellationPending) return filesByRelativePath;
            string directory = pending.Pop();
            try
            {
                foreach (string file in Directory.GetFiles(directory))
                {
                    try
                    {
                        SyncFileMetadata info = ReadSyncFileMetadata(file);
                        string relative = file.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                        if (info != null)
                        {
                            filesByRelativePath[relative] = info;
                            if (info.IsOnlineOnly) result.CloudPlaceholderCount++;
                        }
                        scanned++;
                        if (scanned % 1000 == 0) worker.ReportProgress(0, side + " klasörü taranıyor: " + scanned.ToString("N0", CultureInfo.CurrentCulture) + " dosya...");
                    }
                    catch (Exception ex) { result.Errors++; result.ErrorDetails.Add(file + " — " + ex.Message); }
                }
                if (recursive)
                    foreach (string child in Directory.GetDirectories(directory))
                    {
                        try
                        {
                            FileAttributes attributes = File.GetAttributes(child);
                            if ((attributes & FileAttributes.ReparsePoint) == 0 || IsCloudPlaceholder(child, attributes)) pending.Push(child);
                        }
                        catch (Exception ex) { result.Errors++; result.ErrorDetails.Add(child + " — " + ex.Message); }
                    }
            }
            catch (Exception ex) { result.Errors++; result.ErrorDetails.Add(directory + " — " + ex.Message); }
        }
        return filesByRelativePath;
    }

    private bool FilesHaveSameHash(string leftPath, string rightPath)
    {
        using (SHA256 sha = SHA256.Create())
        using (FileStream left = new FileStream(leftPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (FileStream right = new FileStream(rightPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            return String.Equals(BitConverter.ToString(sha.ComputeHash(left)), BitConverter.ToString(sha.ComputeHash(right)), StringComparison.Ordinal);
    }

    private SyncFileMetadata ReadSyncFileMetadata(string path)
    {
        Win32FindData data;
        IntPtr findHandle = FindFirstFileNative(path, out data);
        if (findHandle.ToInt64() != -1)
        {
            try
            {
                uint attributes = unchecked((uint)data.FileAttributes);
                uint tag = data.Reserved0;
                bool cloudTag = IsCloudReparseTag(tag);
                uint placeholderState = 0;
                bool stateKnown = false;
                if (cloudTag)
                {
                    try { placeholderState = CfGetPlaceholderStateFromAttributeTag(attributes, tag); }
                    catch { placeholderState = 0; }
                    stateKnown = placeholderState != 0;
                }
                bool cloud = cloudTag || (placeholderState & CfPlaceholderStatePlaceholder) != 0;
                bool onDemand = cloud && (((placeholderState & (CfPlaceholderStatePartial | CfPlaceholderStatePartiallyOnDisk)) != 0) ||
                    (attributes & (FileAttributeOffline | FileAttributeRecallOnOpen | FileAttributeRecallOnDataAccess)) != 0 || !stateKnown);
                long length = ((long)data.FileSizeHigh << 32) | data.FileSizeLow;
                long lastWrite = ((long)data.LastWriteTime.High << 32) | unchecked((uint)data.LastWriteTime.Low);
                DateTime modified;
                try { modified = DateTime.FromFileTimeUtc(lastWrite); }
                catch { modified = DateTime.MinValue; }
                return new SyncFileMetadata
                {
                    FullPath = path,
                    Length = length,
                    LastWriteTimeUtc = modified,
                    IsCloudPlaceholder = cloud,
                    IsOnlineOnly = onDemand
                };
            }
            finally { FindClose(findHandle); }
        }
        try
        {
            var info = new FileInfo(path);
            FileAttributes attributes = info.Attributes;
            return new SyncFileMetadata
            {
                FullPath = path,
                Length = info.Length,
                LastWriteTimeUtc = info.LastWriteTimeUtc,
                IsCloudPlaceholder = (attributes & FileAttributes.ReparsePoint) != 0,
                IsOnlineOnly = (attributes & FileAttributes.Offline) != 0
            };
        }
        catch { return null; }
    }

    private bool IsCloudPlaceholder(string path, FileAttributes attributes)
    {
        if ((attributes & FileAttributes.ReparsePoint) == 0) return false;
        SyncFileMetadata metadata = ReadSyncFileMetadata(path);
        return metadata != null && metadata.IsCloudPlaceholder;
    }

    private static bool IsCloudReparseTag(uint tag)
    {
        return tag == IoReparseTagFilePlaceholder || (tag & IoReparseTagCloudMask) == IoReparseTagCloudBase;
    }

    private static string GetSyncCloudStatus(SyncFileMetadata left, bool hasLeft, SyncFileMetadata right, bool hasRight)
    {
        bool leftOnline = hasLeft && left.IsOnlineOnly;
        bool rightOnline = hasRight && right.IsOnlineOnly;
        if (leftOnline && rightOnline) return "İki taraf çevrimiçi";
        if (leftOnline) return "Sol bulutta · indirilmedi";
        if (rightOnline) return "Sağ bulutta · indirilmedi";
        bool leftCloud = hasLeft && left.IsCloudPlaceholder;
        bool rightCloud = hasRight && right.IsCloudPlaceholder;
        if (leftCloud && rightCloud) return "Bulut klasörü · yerel";
        if (leftCloud) return "Sol bulut · yerel";
        if (rightCloud) return "Sağ bulut · yerel";
        return "Yerel";
    }

    private void ShowSyncComparison(SyncComparisonResult result)
    {
        syncCurrentRows = result.Rows;
        syncComparisonList.BeginUpdate();
        try
        {
            syncComparisonList.Items.Clear();
            foreach (SyncComparisonRow row in result.Rows)
            {
                row.LookupStatus = String.Empty;
                row.LookupRootPath = String.Empty;
                row.LookupSizeText = String.Empty;
                var item = new ListViewItem(row.Category);
                item.SubItems.Add(row.RelativePath);
                item.SubItems.Add(row.LeftPath);
                item.SubItems.Add(row.RightPath);
                item.SubItems.Add(row.SizeText);
                item.SubItems.Add(row.CloudStatus);
                item.SubItems.Add(row.LookupStatus);
                item.SubItems.Add(row.LookupRootPath);
                item.SubItems.Add(row.LookupSizeText);
                item.Tag = row;
                if (row.Category == "İki tarafta farklı" || row.Category == "Okuma hatası") item.ForeColor = Color.Firebrick;
                else if (row.Category == "Bulut içeriği bekliyor") item.ForeColor = Color.FromArgb(162, 105, 20);
                else if (row.Category == "Yalnızca sol" || row.Category == "Yalnızca sağ") item.ForeColor = Color.FromArgb(31, 100, 151);
                syncComparisonList.Items.Add(item);
            }
        }
        finally { syncComparisonList.EndUpdate(); }
        string mode = syncComparisonSelector.SelectedIndex == 1 ? "SHA-256 içerik" : "boyut + değiştirme zamanı";
        syncStatus.Text = result.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) + " fark • Sol " + result.LeftOnlyCount.ToString("N0", CultureInfo.CurrentCulture) +
            " • Sağ " + result.RightOnlyCount.ToString("N0", CultureInfo.CurrentCulture) + " • İki taraf farklı " + result.DifferentCount.ToString("N0", CultureInfo.CurrentCulture) +
            " • Bulut dosyası " + result.CloudPlaceholderCount.ToString("N0", CultureInfo.CurrentCulture) + " • SHA bekliyor " + result.CloudContentDeferredCount.ToString("N0", CultureInfo.CurrentCulture) + " • Uyarı " + result.Errors.ToString("N0", CultureInfo.CurrentCulture);
        actionTooltips.SetToolTip(syncStatus, "Karşılaştırma: " + mode + " | Aynı dosya: " + result.IdenticalCount.ToString("N0", CultureInfo.CurrentCulture) +
            " | Bulut yer tutucusu: " + result.CloudPlaceholderCount.ToString("N0", CultureInfo.CurrentCulture) +
            " | Çevrimiçi dosyaların içeriği indirilmeden bırakıldı.");
        if (result.Errors > 0)
        {
            var detail = new StringBuilder("Bazı dosya veya klasörler okunamadı:").AppendLine();
            foreach (string error in result.ErrorDetails.Take(12)) detail.AppendLine(error);
            if (result.ErrorDetails.Count > 12) detail.AppendLine("… ve " + (result.ErrorDetails.Count - 12).ToString(CultureInfo.CurrentCulture) + " başka uyarı.");
            MessageBox.Show(this, detail.ToString(), "Karşılaştırma uyarıları", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void StartSyncLookup()
    {
        if (syncLookupWorker == null || syncLookupWorker.IsBusy || (syncComparisonWorker != null && syncComparisonWorker.IsBusy)) return;
        if (syncCurrentRows == null || !syncCurrentRows.Any(row => row.Category == "Yalnızca sol" || row.Category == "Yalnızca sağ"))
        {
            MessageBox.Show(this, "Önce karşılaştırma yapın. Karşı tarafta aranacak yalnızca sol veya yalnızca sağ dosyası bulunamadı.", "Karşı tarafta ara", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        string leftRoot = syncLeftPathBox.Text.Trim().Trim('"');
        string rightRoot = syncRightPathBox.Text.Trim().Trim('"');
        if (!Directory.Exists(leftRoot) || !Directory.Exists(rightRoot))
        {
            MessageBox.Show(this, "Karşı klasörlerde arama yapmak için sol ve sağ klasörlerin ikisi de erişilebilir olmalı.", "Karşı tarafta ara", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        List<SyncComparisonRow> missingRows = syncCurrentRows.Where(row => row.Category == "Yalnızca sol" || row.Category == "Yalnızca sağ").ToList();
        foreach (SyncComparisonRow row in missingRows)
        {
            row.LookupStatus = "Aranıyor…";
            row.LookupRootPath = String.Empty;
            row.LookupSizeText = String.Empty;
        }
        RefreshSyncLookupCells(missingRows);
        syncLookupButton.Enabled = false;
        syncCancelButton.Enabled = true;
        syncCompareButton.Enabled = false;
        syncStatus.Text = "Eksik dosya adları karşı klasörlerin tüm alt klasörlerinde aranıyor…";
        syncLookupWorker.RunWorkerAsync(new SyncLookupRequest
        {
            LeftRoot = Path.GetFullPath(leftRoot),
            RightRoot = Path.GetFullPath(rightRoot),
            Rows = missingRows
        });
    }

    private static Dictionary<string, List<SyncFileMetadata>> BuildSyncFileNameIndex(Dictionary<string, SyncFileMetadata> files)
    {
        var byName = new Dictionary<string, List<SyncFileMetadata>>(StringComparer.OrdinalIgnoreCase);
        foreach (SyncFileMetadata file in files.Values)
        {
            string name = Path.GetFileName(file.FullPath);
            List<SyncFileMetadata> matches;
            if (!byName.TryGetValue(name, out matches)) byName[name] = matches = new List<SyncFileMetadata>();
            matches.Add(file);
        }
        foreach (List<SyncFileMetadata> matches in byName.Values)
            matches.Sort((first, second) => StringComparer.CurrentCultureIgnoreCase.Compare(first.FullPath, second.FullPath));
        return byName;
    }

    private void ShowSyncLookupResults(SyncLookupResult result)
    {
        int found = 0, missing = 0;
        foreach (SyncComparisonRow row in result.Rows)
        {
            Dictionary<string, List<SyncFileMetadata>> index = row.Category == "Yalnızca sol" ? result.RightFilesByName : result.LeftFilesByName;
            List<SyncFileMetadata> matches;
            bool exists = index.TryGetValue(Path.GetFileName(row.RelativePath), out matches) && matches.Count > 0;
            row.LookupStatus = exists ? (matches.Count == 1 ? "Var" : "Var (" + matches.Count.ToString("N0", CultureInfo.CurrentCulture) + ")") : "Yok";
            row.LookupRootPath = exists ? String.Join(" | ", matches.Select(file => StorePortablePath(file.FullPath))) : String.Empty;
            row.LookupSizeText = exists ? String.Join(" | ", matches.Select(file => FormatFileSize(file.Length))) : String.Empty;
            if (exists) found++; else missing++;
        }
        RefreshSyncLookupCells(result.Rows);
        syncStatus.Text = "Karşı klasörlerde ad araması tamamlandı • Var: " + found.ToString("N0", CultureInfo.CurrentCulture) + " • Yok: " + missing.ToString("N0", CultureInfo.CurrentCulture) +
            " • Alt klasörler dahil tarandı" + (result.Errors > 0 ? " • Okuma uyarısı: " + result.Errors.ToString("N0", CultureInfo.CurrentCulture) : String.Empty);
        if (result.Errors > 0)
        {
            var detail = new StringBuilder("Karşı klasörlerdeki bazı dosya veya klasörler okunamadı:").AppendLine();
            foreach (string error in result.ErrorDetails.Take(12)) detail.AppendLine(error);
            if (result.ErrorDetails.Count > 12) detail.AppendLine("… ve " + (result.ErrorDetails.Count - 12).ToString(CultureInfo.CurrentCulture) + " başka uyarı.");
            MessageBox.Show(this, detail.ToString(), "Karşı klasör arama uyarıları", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void RefreshSyncLookupCells(IEnumerable<SyncComparisonRow> rows)
    {
        foreach (SyncComparisonRow row in rows)
            foreach (ListViewItem item in syncComparisonList.Items)
                if (Object.ReferenceEquals(item.Tag, row))
                {
                    item.SubItems[6].Text = row.LookupStatus ?? String.Empty;
                    item.SubItems[7].Text = row.LookupRootPath ?? String.Empty;
                    item.SubItems[8].Text = row.LookupSizeText ?? String.Empty;
                    break;
                }
    }

    private void ExportSyncComparisonCsv()
    {
        if (syncComparisonList.Items.Count == 0)
        {
            MessageBox.Show(this, "Dışa aktarılacak karşılaştırma satırı yok.", "CSV dışa aktar", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using (var dialog = new SaveFileDialog())
        {
            dialog.Title = "Senkron karşılaştırmasını CSV olarak kaydet";
            dialog.Filter = "CSV dosyası (*.csv)|*.csv|Tüm dosyalar (*.*)|*.*";
            dialog.DefaultExt = "csv";
            dialog.AddExtension = true;
            dialog.FileName = "Senkron-Karsilastirma-" + DateTime.Now.ToString("yyyyMMdd-HHmm", CultureInfo.InvariantCulture) + ".csv";
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                var csv = new StringBuilder();
                csv.AppendLine(String.Join(";", new string[] { "Durum", "Göreli yol", "Sol taraftaki dosya", "Sağ taraftaki dosya", "Boyut", "Bulut durumu", "Karşı klasörde", "Bulunan dosyanın yolu", "Bulunan dosya boyutu" }.Select(CsvQuote)));
                foreach (ListViewItem item in syncComparisonList.Items)
                    csv.AppendLine(String.Join(";", item.SubItems.Cast<ListViewItem.ListViewSubItem>().Select(subItem => CsvQuote(subItem.Text))));
                File.WriteAllText(dialog.FileName, csv.ToString(), new UTF8Encoding(true));
                syncStatus.Text = "CSV kaydedildi: " + dialog.FileName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "CSV dosyası kaydedilemedi.\r\n\r\n" + ex.Message, "CSV dışa aktar", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private static string CsvQuote(string value)
    {
        return "\"" + (value ?? String.Empty).Replace("\"", "\"\"") + "\"";
    }

    private void ToggleSyncFullScreen(bool enter)
    {
        if (enter == syncFullScreen || syncPage == null || syncPageContent == null) return;
        try
        {
            if (enter)
            {
                syncPreviousWindowState = WindowState;
                syncPreviousBounds = WindowState == FormWindowState.Maximized ? RestoreBounds : Bounds;
                syncPreviousBorderStyle = FormBorderStyle;
                tabs.Visible = false;
                if (mainHeader != null) mainHeader.Visible = false;
                status.Visible = false;
                WindowState = FormWindowState.Normal;
                syncPage.Controls.Remove(syncPageContent);
                syncPageContent.Dock = DockStyle.Fill;
                Controls.Add(syncPageContent);
                FormBorderStyle = FormBorderStyle.None;
                Bounds = Screen.FromControl(this).Bounds;
                syncPageContent.BringToFront();
                syncFullScreen = true;
                syncFullScreenButton.Text = "Tam ekrandan çık";
                syncStatus.Text = "Tam ekran • Çıkmak için Esc veya F11.";
            }
            else
            {
                Controls.Remove(syncPageContent);
                syncPageContent.Dock = DockStyle.Fill;
                syncPage.Controls.Add(syncPageContent);
                tabs.Visible = true;
                if (mainHeader != null) mainHeader.Visible = true;
                status.Visible = true;
                FormBorderStyle = syncPreviousBorderStyle;
                Bounds = syncPreviousBounds;
                WindowState = syncPreviousWindowState;
                tabs.SelectedTab = syncPage;
                syncFullScreen = false;
                syncFullScreenButton.Text = "Tam ekran";
            }
        }
        catch (Exception ex)
        {
            try
            {
                if (syncPageContent.Parent == this) Controls.Remove(syncPageContent);
                if (syncPageContent.Parent != syncPage) syncPage.Controls.Add(syncPageContent);
                tabs.Visible = true;
                if (mainHeader != null) mainHeader.Visible = true;
                status.Visible = true;
                FormBorderStyle = syncPreviousBorderStyle;
                Bounds = syncPreviousBounds;
                WindowState = syncPreviousWindowState;
                tabs.SelectedTab = syncPage;
            }
            catch { }
            syncFullScreen = false;
            syncFullScreenButton.Text = "Tam ekran";
            MessageBox.Show(this, "Tam ekran geçişi tamamlanamadı; önceki görünüm geri yüklendi.\r\n\r\n" + ex,
                "Senkronize tam ekran", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SyncTransferRightPaneToExplorer()
    {
        if (transferRightPane == null || String.IsNullOrWhiteSpace(currentPath) || !Directory.Exists(currentPath)) return;
        Entry selected = files.SelectedItems.Count > 0 ? files.SelectedItems[0].Tag as Entry : null;
        string selectedFile = selected != null && File.Exists(selected.FullPath) ? selected.FullPath : null;
        string targetFolder = selectedFile == null ? currentPath : Path.GetDirectoryName(selectedFile);
        TransferNavigate(transferRightPane, targetFolder);
        if (selectedFile != null) SelectTransferFiles(transferRightPane, new[] { selectedFile });
        SaveTransferPanePaths();
    }

    private Control CreateTransferPane(TransferPaneState pane)
    {
        var root = new Panel { Dock = DockStyle.Fill, Padding = new Padding(6) };
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 36 };
        pane.LocationButton = new Button { Text = "⌄", Width = 30, Height = 28, Location = new Point(34, 3) };
        pane.LocationButton.Click += (s, e) => ShowTransferQuickLocationsMenu(pane);
        actionTooltips.SetToolTip(pane.LocationButton, "Windows klasör kısayollarını aç");
        pane.UpButton = new Button { Text = "↑", Width = 32, Height = 28, Location = new Point(0, 3) };
        pane.UpButton.Click += (s, e) =>
        {
            if (pane.ShowingComputerDrives)
            {
                TransferNavigate(pane, pane.ComputerViewReturnPath);
                return;
            }
            DirectoryInfo parent = Directory.GetParent(pane.CurrentPath);
            if (parent != null) TransferNavigate(pane, parent.FullName);
            else ShowTransferComputerDrives(pane);
        };
        pane.Address = new TextBox { Height = 27, Location = new Point(70, 4), Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right };
        pane.GoButton = new Button { Text = "Git", Width = 48, Height = 28, Anchor = AnchorStyles.Top | AnchorStyles.Right };
        pane.FolderWidthButton = new Button { Text = "◂", Width = 30, Height = 28, Anchor = AnchorStyles.Top | AnchorStyles.Right };
        actionTooltips.SetToolTip(pane.FolderWidthButton, "Klasör listesini daralt veya genişlet");
        Action layoutToolbar = () =>
        {
            pane.FolderWidthButton.Left = toolbar.ClientSize.Width - 32;
            pane.GoButton.Left = pane.FolderWidthButton.Left - 50;
            pane.Address.Width = Math.Max(80, pane.GoButton.Left - pane.Address.Left - 4);
        };
        toolbar.Resize += (s, e) => layoutToolbar();
        pane.FolderWidthButton.Click += (s, e) =>
        {
            if (!pane.FolderSplit.Panel1Collapsed)
            {
                pane.LastFolderPaneWidth = pane.FolderSplit.SplitterDistance;
                pane.FolderSplit.Panel1Collapsed = true;
                pane.FolderWidthButton.Text = "▸";
            }
            else
            {
                pane.FolderSplit.Panel1Collapsed = false;
                SetTransferFolderWidth(pane, pane.LastFolderPaneWidth > pane.FolderSplit.Panel1MinSize
                    ? pane.LastFolderPaneWidth : 180);
                pane.FolderWidthButton.Text = "◂";
            }
            pane.FolderPaneHidden = pane.FolderSplit.Panel1Collapsed;
            if (pane.Side == 0) transferLeftPaneHidden = pane.FolderPaneHidden; else transferRightPaneHidden = pane.FolderPaneHidden;
            SaveTransferPanePaths();
        };
        pane.GoButton.Click += (s, e) => TransferNavigateFromAddress(pane);
        pane.Address.KeyDown += (s, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            TransferNavigateFromAddress(pane);
            e.SuppressKeyPress = true;
        };
        toolbar.Controls.Add(pane.LocationButton);
        toolbar.Controls.Add(pane.UpButton);
        toolbar.Controls.Add(pane.Address);
        toolbar.Controls.Add(pane.GoButton);
        toolbar.Controls.Add(pane.FolderWidthButton);
        layoutToolbar();

        var body = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Size = new Size(700, 600),
            Orientation = Orientation.Vertical,
            SplitterDistance = 175,
            Panel1MinSize = 28,
            Panel2MinSize = 120,
            BackColor = Color.FromArgb(226, 231, 238),
            BorderStyle = BorderStyle.FixedSingle
        };
        pane.FolderSplit = body;
        pane.FolderList = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, SelectionMode = SelectionMode.MultiExtended };
        pane.FolderList.DrawMode = DrawMode.OwnerDrawFixed;
        pane.FolderList.ItemHeight = 26;
        pane.FolderList.DrawItem += (s, e) => DrawFolderListItem(pane.FolderList, e);
        pane.FileList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            MultiSelect = true,
            CheckBoxes = true,
            HideSelection = false,
            AllowDrop = true
        };
        pane.FileList.SmallImageList = fileIconImageList;
        pane.FileList.Columns.Add("Dosya adı", 250, HorizontalAlignment.Left);
        pane.FileList.Columns.Add("Boyut", 85, HorizontalAlignment.Right);
        pane.FileList.Columns.Add("Tür", 115, HorizontalAlignment.Left);
        pane.FileList.Columns.Add("Değiştirme tarihi", 130, HorizontalAlignment.Left);
        pane.FileList.ColumnClick += (s, e) => SortTransferPane(pane, e.Column);

        var folderArea = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4) };
        var folderHeading = new Label { Text = "Klasörler · tek tıkla aç", Dock = DockStyle.Top, Height = 27, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
        folderArea.Controls.Add(pane.FolderList);
        folderArea.Controls.Add(folderHeading);
        var fileArea = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4) };
        var fileHeading = new Label { Text = "Dosyalar", Dock = DockStyle.Top, Height = 27, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
        fileArea.Controls.Add(pane.FileList);
        fileArea.Controls.Add(fileHeading);
        body.Panel1.Controls.Add(folderArea);
        body.Panel2.Controls.Add(fileArea);
        root.Controls.Add(body);
        root.Controls.Add(toolbar);

        pane.FolderList.MouseClick += (s, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            if ((ModifierKeys & (Keys.Control | Keys.Shift)) != Keys.None) return;
            int index = pane.FolderList.IndexFromPoint(e.Location);
            Entry selected = index >= 0 ? pane.FolderList.Items[index] as Entry : null;
            if (selected != null && String.Equals(selected.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) ShowTransferComputerDrives(pane);
            else if (selected != null && selected.IsParent && pane.ShowingComputerDrives) TransferNavigate(pane, selected.FullPath);
            else if (selected != null && Directory.Exists(selected.FullPath)) TransferNavigate(pane, selected.FullPath);
        };
        pane.FolderList.KeyDown += (s, e) =>
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                for (int i = 0; i < pane.FolderList.Items.Count; i++) pane.FolderList.SetSelected(i, true);
            }
            else if (e.Control && (e.KeyCode == Keys.C || e.KeyCode == Keys.X)) SetTransferFolderClipboard(pane, e.KeyCode == Keys.X);
            else if (e.Control && e.KeyCode == Keys.V) PasteTransferClipboard(pane, GetTransferDestination(pane));
            else if (e.KeyCode == Keys.Enter)
            {
                Entry selected = pane.FolderList.SelectedItem as Entry;
                if (selected != null && String.Equals(selected.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) ShowTransferComputerDrives(pane);
                else if (selected != null && selected.IsParent && pane.ShowingComputerDrives) TransferNavigate(pane, selected.FullPath);
                else if (selected != null && Directory.Exists(selected.FullPath)) TransferNavigate(pane, selected.FullPath);
            }
            else if (e.KeyCode == Keys.F2) RenameTransferFolderSelection(pane);
            else return;
            e.Handled = e.SuppressKeyPress = true;
        };
        pane.FolderList.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            int index = pane.FolderList.IndexFromPoint(e.Location);
            if (index >= 0 && !pane.FolderList.GetSelected(index))
            {
                pane.FolderList.ClearSelected();
                pane.FolderList.SetSelected(index, true);
            }
        };
        pane.FileList.ItemCheck += (s, e) =>
        {
            if (pane.SyncingSelection) return;
            int index = e.Index;
            bool selected = e.NewValue == CheckState.Checked;
            try
            {
                pane.FileList.BeginInvoke((MethodInvoker)(() =>
                {
                    if (pane.FileList.IsDisposed || index >= pane.FileList.Items.Count) return;
                    pane.SyncingSelection = true;
                    try { pane.FileList.Items[index].Selected = selected; }
                    finally { pane.SyncingSelection = false; }
                }));
            }
            catch { }
        };
        pane.FileList.ItemSelectionChanged += (s, e) =>
        {
            if (pane.SyncingSelection) return;
            pane.SyncingSelection = true;
            try { e.Item.Checked = e.IsSelected; }
            finally { pane.SyncingSelection = false; }
        };
        pane.FileList.MouseDoubleClick += (s, e) =>
        {
            ListViewHitTestInfo hit = pane.FileList.HitTest(e.Location);
            Entry selected = hit.Item == null ? null : hit.Item.Tag as Entry;
            if (selected != null && Directory.Exists(selected.FullPath)) TransferNavigate(pane, selected.FullPath);
            else if (selected != null && File.Exists(selected.FullPath)) OpenFile(selected.FullPath);
        };
        pane.FileList.ItemDrag += (s, e) => StartTransferDrag(pane);
        pane.FileList.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                string selected = GetTransferSelectedPaths(pane).FirstOrDefault();
                if (!String.IsNullOrWhiteSpace(selected))
                {
                    if (Directory.Exists(selected)) TransferNavigate(pane, selected);
                    else if (File.Exists(selected)) OpenFile(selected);
                }
            }
            else if (e.KeyCode == Keys.Delete) DeleteTransferFiles(pane);
            else if (e.KeyCode == Keys.F2)
            {
                Entry selected = pane.FileList.SelectedItems.Count > 0 ? pane.FileList.SelectedItems[0].Tag as Entry : null;
                RenameTransferEntry(pane, selected);
            }
            else if (e.KeyCode == Keys.Space)
            {
                Entry selected = pane.FileList.SelectedItems.Count > 0 ? pane.FileList.SelectedItems[0].Tag as Entry : null;
                if (selected != null && File.Exists(selected.FullPath)) new FilePreviewForm(selected.FullPath, true).Show(this);
            }
            else if (e.Control && e.KeyCode == Keys.C) SetTransferClipboard(pane, false);
            else if (e.Control && e.KeyCode == Keys.X) SetTransferClipboard(pane, true);
            else if (e.Control && e.KeyCode == Keys.V) PasteTransferClipboard(pane, GetTransferDestination(pane));
            else if (e.Control && e.KeyCode == Keys.A)
            {
                foreach (ListViewItem item in pane.FileList.Items) { item.Selected = true; item.Checked = true; }
            }
            else return;
            e.Handled = e.SuppressKeyPress = true;
        };
        SetupTransferDropTargets(pane, pane.FolderList, pane.FolderList, true);
        SetupTransferDropTargets(pane, pane.FileList, pane.FileList, false);
        ConfigureTransferMenus(pane);
        pane.FileList.Enter += (s, e) => { activeTransferPane = pane; };
        pane.FolderList.Enter += (s, e) => { activeTransferPane = pane; };
        if (pane.FolderPaneHidden)
        {
            pane.FolderSplit.Panel1Collapsed = true;
            pane.FolderWidthButton.Text = "▸";
        }
        return root;
    }

    private TransferPaneState activeTransferPane;

    private void TransferNavigateFromAddress(TransferPaneState pane)
    {
        TransferNavigate(pane, pane.Address.Text);
    }

    private void ShowTransferQuickLocationsMenu(TransferPaneState pane)
    {
        if (pane.QuickLocationsMenu == null) pane.QuickLocationsMenu = new ContextMenuStrip();
        PopulateQuickLocationsMenu(pane.QuickLocationsMenu, path => TransferNavigate(pane, path), false);
        pane.QuickLocationsMenu.Show(pane.LocationButton, new Point(0, pane.LocationButton.Height));
    }

    private void TransferNavigate(TransferPaneState pane, string path)
    {
        try
        {
            string full = Path.GetFullPath(Environment.ExpandEnvironmentVariables((path ?? String.Empty).Trim().Trim('"')));
            if (!Directory.Exists(full)) throw new DirectoryNotFoundException("Bu klasör bulunamadı: " + full);
            int request = ++pane.LoadGeneration;
            pane.CurrentPath = full;
            pane.ShowingComputerDrives = false;
            pane.Address.Text = full;
            pane.Address.ReadOnly = false;
            pane.GoButton.Enabled = true;
            pane.UpButton.Enabled = true;
            pane.Loading = true;
            pane.FolderList.BeginUpdate();
            pane.FileList.BeginUpdate();
            pane.FolderList.Items.Clear();
            pane.FolderList.Items.Add(new Entry("Bu Bilgisayar", ThisComputerPath));
            DirectoryInfo parent = Directory.GetParent(full);
            if (parent != null) pane.FolderList.Items.Add(new Entry("↰  ..", parent.FullName, true));
            pane.FileList.Items.Clear();
            pane.FolderList.EndUpdate();
            pane.FileList.EndUpdate();
            transferStatus.Text = pane.Title + " · Klasör okunuyor...";
            ConfigureTransferPaneWatcher(pane);
            SaveTransferPanePaths();
            ThreadPool.QueueUserWorkItem(_ =>
            {
                var folders = new List<Entry>();
                var rows = new List<TransferFileRow>();
                Exception failure = null;
                try
                {
                    string[] directoryPaths = Directory.GetDirectories(full);
                    string[] filePaths = Directory.GetFiles(full);
                    foreach (string directory in directoryPaths.OrderBy(item => Path.GetFileName(item), NaturalNameComparer.Instance))
                        folders.Add(new Entry(Path.GetFileName(directory), directory));
                    foreach (string directory in directoryPaths.OrderBy(item => Path.GetFileName(item), NaturalNameComparer.Instance))
                    {
                        var info = new DirectoryInfo(directory);
                        rows.Add(new TransferFileRow(new Entry(info.Name, directory), info.Name, String.Empty, "Dosya klasörü", String.Empty, true));
                    }
                    foreach (string file in filePaths.OrderBy(item => Path.GetFileName(item), NaturalNameComparer.Instance))
                    {
                        var info = new FileInfo(file);
                        string size = FormatFileSize(info.Length);
                        string type;
                        try { type = GetFileTypeDescription(file); } catch { type = Path.GetExtension(file); }
                        rows.Add(new TransferFileRow(new Entry(info.Name, file), info.Name, size, type, info.LastWriteTime.ToString("dd.MM.yyyy HH:mm"), false));
                    }
                }
                catch (Exception ex) { failure = ex; }
                Action applyResult = () =>
                {
                    if (IsDisposed || Disposing || pane.FileList == null || pane.FileList.IsDisposed || request != pane.LoadGeneration) return;
                    string genericFileIcon = EnsureFileIcon(String.Empty, false);
                    pane.FolderList.BeginUpdate(); pane.FileList.BeginUpdate();
                    try
                    {
                        foreach (Entry folder in folders) pane.FolderList.Items.Add(folder);
                        foreach (TransferFileRow data in rows)
                        {
                            var item = new ListViewItem(new string[] { data.Name, data.Size, data.Type, data.Modified });
                            item.Tag = data.Entry;
                            item.ImageKey = data.IsDirectory ? folderIconKey : genericFileIcon;
                            pane.FileList.Items.Add(item);
                        }
                    }
                    finally { pane.FileList.EndUpdate(); pane.FolderList.EndUpdate(); }
                    pane.Loading = false; pane.Refreshing = false;
                    if (failure != null) transferStatus.Text = pane.Title + " · Klasör okunamadı: " + failure.Message;
                    else transferStatus.Text = pane.Title + " · " + full + " · " + folders.Count + " klasör, " + rows.Count + " öğe";
                    ApplyTransferSort(pane); RefreshTransferCutAppearance();
                };
                PostToUiWhenReady(applyResult);
            });
            return;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Çiftli İşlem", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void PostToUiWhenReady(Action action)
    {
        if (action == null || IsDisposed || Disposing) return;
        if (IsHandleCreated)
        {
            try { BeginInvoke((MethodInvoker)(() => action())); } catch { }
            return;
        }
        EventHandler created = null;
        created = (s, e) =>
        {
            HandleCreated -= created;
            if (IsDisposed || Disposing) return;
            try { BeginInvoke((MethodInvoker)(() => action())); } catch { }
        };
        HandleCreated += created;
        if (IsHandleCreated)
        {
            HandleCreated -= created;
            try { BeginInvoke((MethodInvoker)(() => action())); } catch { }
        }
    }

    private void ShowTransferComputerDrives(TransferPaneState pane)
    {
        if (pane == null) return;
        pane.ComputerViewReturnPath = pane.CurrentPath;
        pane.ShowingComputerDrives = true;
        pane.Address.Text = "Bu Bilgisayar";
        pane.Address.ReadOnly = true;
        pane.GoButton.Enabled = false;
        pane.UpButton.Enabled = true;
        pane.FolderList.BeginUpdate();
        pane.FileList.BeginUpdate();
        try
        {
            pane.FolderList.Items.Clear();
            pane.FileList.Items.Clear();
            pane.FolderList.Items.Add(new Entry("↰  Geri dön", pane.ComputerViewReturnPath, true));
            foreach (Entry drive in GetComputerDriveEntries())
            {
                pane.FolderList.Items.Add(drive);
                long capacity = 0;
                try
                {
                    var driveInfo = new DriveInfo(drive.FullPath);
                    if (driveInfo.IsReady) capacity = driveInfo.TotalSize;
                }
                catch { }
                string capacityText = capacity <= 0 ? String.Empty : (capacity / (1024.0 * 1024.0 * 1024.0)).ToString("0.##", CultureInfo.CurrentCulture) + " GB";
                var driveRow = new ListViewItem(new string[] { drive.ToString(), capacityText, "Sürücü", String.Empty });
                driveRow.Tag = drive;
                SetFileRowIcon(driveRow, drive.FullPath);
                pane.FileList.Items.Add(driveRow);
            }
        }
        finally
        {
            pane.FileList.EndUpdate();
            pane.FolderList.EndUpdate();
        }
        ConfigureTransferPaneWatcher(pane);
        transferStatus.Text = pane.Title + " · Bu Bilgisayar · " + Math.Max(0, pane.FolderList.Items.Count - 1) + " sürücü";
    }

    private List<Entry> GetComputerDriveEntries()
    {
        var entries = new List<Entry>();
        try
        {
            foreach (DriveInfo drive in DriveInfo.GetDrives().OrderBy(item => item.Name, NaturalNameComparer.Instance))
            {
                string path = drive.Name;
                string label = path;
                try
                {
                    if (drive.IsReady && !String.IsNullOrWhiteSpace(drive.VolumeLabel))
                        label = drive.VolumeLabel + " (" + path.TrimEnd(Path.DirectorySeparatorChar) + ")";
                }
                catch { }
                entries.Add(new Entry(label, path));
            }
        }
        catch { }
        return entries;
    }

    private void ShowComputerDrives()
    {
        if (!showingComputerDrives) computerViewReturnPath = currentPath;
        showingComputerDrives = true;
        files.LabelEdit = false;
        address.Text = "Bu Bilgisayar";
        address.ReadOnly = true;
        breadcrumbBar.Visible = false;
        up.Enabled = true;
        folderSearch.Text = String.Empty;
        fileSearch.Text = String.Empty;
        folders.BeginUpdate();
        files.BeginUpdate();
        try
        {
            folders.Items.Clear();
            files.Items.Clear();
            fileRows.Clear();
            folders.Items.Add(new Entry("↰  Geri dön", computerViewReturnPath, true));
            foreach (Entry drive in GetComputerDriveEntries())
            {
                folders.Items.Add(drive);
                long capacity = 0;
                try
                {
                    var driveInfo = new DriveInfo(drive.FullPath);
                    if (driveInfo.IsReady) capacity = driveInfo.TotalSize;
                }
                catch { }
                string size = capacity <= 0 ? String.Empty : (capacity / (1024.0 * 1024.0 * 1024.0)).ToString("0.##", CultureInfo.CurrentCulture) + " GB";
                var driveRow = new ListViewItem(new string[] { drive.ToString(), String.Empty, size, "Sürücü", String.Empty });
                driveRow.Tag = drive;
                SetFileRowIcon(driveRow, drive.FullPath);
                fileRows.Add(driveRow);
            }
        }
        finally
        {
            files.EndUpdate();
            folders.EndUpdate();
        }
        RefreshFileRows();
        status.Text = "Bu Bilgisayar   ·   " + Math.Max(0, folders.Items.Count - 1) + " sürücü";
        ConfigureCurrentFolderWatcher();
    }

    private void ConfigureTransferPaneWatcher(TransferPaneState pane)
    {
        if (pane == null) return;
        try
        {
            string desiredPath = !pane.ShowingComputerDrives && Directory.Exists(pane.CurrentPath) ? Path.GetFullPath(pane.CurrentPath) : String.Empty;
            if (String.Equals(pane.WatchedPath, desiredPath, StringComparison.OrdinalIgnoreCase) && pane.Watcher != null) return;
            if (pane.Watcher != null)
            {
                pane.Watcher.EnableRaisingEvents = false;
                pane.Watcher.Dispose();
                pane.Watcher = null;
            }
            pane.WatchedPath = desiredPath;
            if (String.IsNullOrWhiteSpace(desiredPath)) return;
            var watcher = new FileSystemWatcher(desiredPath)
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.Size
            };
            watcher.Created += (s, e) => ScheduleTransferPaneRefresh(pane);
            watcher.Deleted += (s, e) => ScheduleTransferPaneRefresh(pane);
            // OneDrive emits repeated Changed events while syncing/hydrating cloud files.
            // Those events do not change the directory listing and caused refresh loops.
            if (!IsOneDrivePath(desiredPath)) watcher.Changed += (s, e) => ScheduleTransferPaneRefresh(pane);
            watcher.Renamed += (s, e) => ScheduleTransferPaneRefresh(pane);
            watcher.Error += (s, e) => ScheduleTransferPaneRefresh(pane);
            pane.Watcher = watcher;
            watcher.EnableRaisingEvents = true;
        }
        catch { }
    }

    private void ScheduleTransferPaneRefresh(TransferPaneState pane)
    {
        if (pane == null || IsDisposed || Disposing) return;
        try
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (pane.FileList == null || pane.FileList.IsDisposed || pane.Refreshing || pane.Loading) return;
                if (pane.RefreshTimer == null)
                {
                    pane.RefreshTimer = new System.Windows.Forms.Timer { Interval = 500 };
                    pane.RefreshTimer.Tick += (s, e) =>
                    {
                        pane.RefreshTimer.Stop();
                        RefreshTransferPaneFromWatcher(pane);
                    };
                }
                if (pane.RefreshTimer.Enabled) return;
                pane.RefreshTimer.Stop();
                pane.RefreshTimer.Start();
            }));
        }
        catch { }
    }

    private void RefreshTransferPaneFromWatcher(TransferPaneState pane)
    {
        if (pane == null || pane.ShowingComputerDrives || pane.Refreshing || String.IsNullOrWhiteSpace(pane.CurrentPath) || !Directory.Exists(pane.CurrentPath)) return;
        if ((DateTime.UtcNow - pane.LastWatcherRefreshUtc).TotalMilliseconds < 1200) return;
        pane.LastWatcherRefreshUtc = DateTime.UtcNow;
        pane.Refreshing = true;
        string path = pane.CurrentPath;
        var selectedPaths = new HashSet<string>(pane.FileList.Items.Cast<ListViewItem>()
            .Where(item => item.Selected).Select(item => item.Tag as Entry).Where(entry => entry != null)
            .Select(entry => entry.FullPath), StringComparer.OrdinalIgnoreCase);
        var checkedPaths = new HashSet<string>(pane.FileList.Items.Cast<ListViewItem>()
            .Where(item => item.Checked).Select(item => item.Tag as Entry).Where(entry => entry != null)
            .Select(entry => entry.FullPath), StringComparer.OrdinalIgnoreCase);
        try
        {
            TransferNavigate(pane, path);
            pane.SyncingSelection = true;
            try
            {
                foreach (ListViewItem item in pane.FileList.Items)
                {
                    Entry entry = item.Tag as Entry;
                    if (entry == null) continue;
                    item.Selected = selectedPaths.Contains(entry.FullPath);
                    item.Checked = checkedPaths.Contains(entry.FullPath);
                }
            }
            finally { pane.SyncingSelection = false; }
        }
        finally { pane.Refreshing = false; }
    }

    private void DisposeTransferPaneWatcher(TransferPaneState pane)
    {
        if (pane == null) return;
        try
        {
            if (pane.RefreshTimer != null) { pane.RefreshTimer.Stop(); pane.RefreshTimer.Dispose(); pane.RefreshTimer = null; }
            if (pane.Watcher != null) { pane.Watcher.EnableRaisingEvents = false; pane.Watcher.Dispose(); pane.Watcher = null; }
        }
        catch { }
    }

    private void SetTransferFolderWidth(TransferPaneState pane, int requestedWidth)
    {
        SplitContainer split = pane.FolderSplit;
        if (split == null || split.IsDisposed || split.ClientSize.Width <= 0) return;
        int maximum = split.ClientSize.Width - split.Panel2MinSize - split.SplitterWidth;
        if (maximum < split.Panel1MinSize) return;
        try { split.SplitterDistance = Math.Max(split.Panel1MinSize, Math.Min(maximum, requestedWidth)); }
        catch { }
    }

    private void SortTransferPane(TransferPaneState pane, int column)
    {
        if (pane.SortColumn == column) pane.SortAscending = !pane.SortAscending;
        else
        {
            pane.SortColumn = column;
            pane.SortAscending = true;
        }
        ApplyTransferSort(pane);
        SaveListSortPreference(pane.Side == 0 ? "Çiftli İşlem Sol" : "Çiftli İşlem Sağ", pane.SortColumn, pane.SortAscending);
    }

    private void ApplySavedSortPreference(string key, TransferPaneState pane)
    {
        FileSortPreference preference;
        if (pane != null && listSortPreferences.TryGetValue(key, out preference))
        {
            pane.SortColumn = preference.Column;
            pane.SortAscending = preference.Ascending;
        }
    }

    private void ApplyTransferSort(TransferPaneState pane)
    {
        if (pane == null || pane.FileList == null || pane.FileList.Columns.Count < 4) return;
        string[] headers = { "Dosya adı", "Boyut", "Tür", "Değiştirme tarihi" };
        for (int i = 0; i < headers.Length; i++)
            pane.FileList.Columns[i].Text = headers[i] + (pane.SortColumn == i ? (pane.SortAscending ? "  ▲" : "  ▼") : String.Empty);
        if (pane.SortColumn < 0) return;
        pane.FileList.ListViewItemSorter = new TransferListViewComparer(pane.SortColumn, pane.SortAscending);
        pane.FileList.Sort();
    }

    private void ConfigureTransferMenus(TransferPaneState pane)
    {
        var fileMenu = new ContextMenuStrip();
        var openFile = new ToolStripMenuItem("Dosya Aç");
        var openExplorer = new ToolStripMenuItem("Windows Explorer ile Aç");
        var preview = new ToolStripMenuItem("Ön izleme");
        var rename = new ToolStripMenuItem("Yeniden adlandır");
        var delete = new ToolStripMenuItem("Sil");
        var undoDelete = new ToolStripMenuItem("Geri Al Silme");
        var cut = new ToolStripMenuItem("Kes");
        var copy = new ToolStripMenuItem("Kopyala");
        var paste = new ToolStripMenuItem("Yapıştır");
        var undoMove = new ToolStripMenuItem("Geri Al Taşı");
        var mail = new ToolStripMenuItem("Mail Gönder");
        var whatsapp = new ToolStripMenuItem("WhatsApp için karakter uyumlu paylaş");
        var zip = new ToolStripMenuItem("ZIP ile sıkıştır");
        var extractHere = new ToolStripMenuItem("Arşivi buraya ayıkla");
        var extractNamed = new ToolStripMenuItem("Arşivi kendi klasörüne ayıkla");
        var extractChoose = new ToolStripMenuItem("Arşivi seçilen klasöre ayıkla...");
        var properties = new ToolStripMenuItem("Özellikler");
        var copyPath = new ToolStripMenuItem("Yolu kopyala");
        var actionBar = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = false,
            Size = new Size(244, 38), Padding = new Padding(2, 1, 2, 1), Margin = Padding.Empty, BackColor = Color.White };
        Button cutButton = CreateExplorerMenuActionButton("Kes", "\uE8C6", 38);
        Button copyButton = CreateExplorerMenuActionButton("Kopyala", "\uE8C8", 54);
        Button pasteButton = CreateExplorerMenuActionButton("Yapıştır", "\uE77F", 45);
        Button shareButton = CreateExplorerMenuActionButton("Paylaş", "\uE72D", 40);
        Button deleteButton = CreateExplorerMenuActionButton("Sil", "\uE74D", 38);
        actionBar.Controls.AddRange(new Control[] { cutButton, CreateExplorerMenuSeparator(), copyButton,
            CreateExplorerMenuSeparator(), pasteButton, CreateExplorerMenuSeparator(), shareButton, CreateExplorerMenuSeparator(), deleteButton });
        fileMenu.Items.Add(new ToolStripControlHost(actionBar) { AutoSize = false, Size = new Size(244, 40), Margin = Padding.Empty, Padding = Padding.Empty, BackColor = Color.White });
        fileMenu.Items.Add(new ToolStripSeparator());
        var clipboardMenu = new ToolStripMenuItem("Pano ve taşıma");
        clipboardMenu.DropDownItems.AddRange(new ToolStripItem[] { cut, copy, paste });
        var viewMenu = new ToolStripMenuItem("Görünüm");
        var details = new ToolStripMenuItem("Ayrıntılar") { Checked = true, Enabled = false };
        viewMenu.DropDownItems.Add(details);
        var sortMenu = new ToolStripMenuItem("Sıralama ölçütü");
        foreach (int column in Enumerable.Range(0, pane.FileList.Columns.Count))
        {
            int captured = column;
            var option = new ToolStripMenuItem(pane.FileList.Columns[column].Text);
            option.Click += (s, e) => SortTransferPane(pane, captured);
            sortMenu.DropDownItems.Add(option);
        }
        var groupMenu = new ToolStripMenuItem("Gruplandırma ölçütü");
        groupMenu.DropDownItems.Add("Gruplandırma yok").Enabled = false;
        var undoMenu = new ToolStripMenuItem("Geri Al Taşı");
        var newMenu = new ToolStripMenuItem("Yeni");
        var newFolder = new ToolStripMenuItem("Klasör");
        newFolder.Click += (s, e) => CreateFolderIn(pane.CurrentPath);
        newMenu.DropDownItems.Add(newFolder);
        var manageMenu = new ToolStripMenuItem("Düzenle ve yönet");
        manageMenu.DropDownItems.AddRange(new ToolStripItem[] { rename, delete, undoDelete });
        var fileTasks = new ToolStripMenuItem("Dosya işlemleri");
        fileTasks.DropDownItems.Add(mail);
        var compression = new ToolStripMenuItem("Sıkıştırma");
        compression.DropDownItems.Add(zip);
        compression.DropDownItems.Add(new ToolStripSeparator());
        compression.DropDownItems.Add(extractHere);
        compression.DropDownItems.Add(extractNamed);
        compression.DropDownItems.Add(extractChoose);
        fileMenu.Items.Add(clipboardMenu);
        fileMenu.Items.Add(openFile);
        fileMenu.Items.Add(preview);
        fileMenu.Items.Add(viewMenu);
        fileMenu.Items.Add(sortMenu);
        fileMenu.Items.Add(groupMenu);
        fileMenu.Items.Add(new ToolStripSeparator());
        fileMenu.Items.Add(undoMenu);
        fileMenu.Items.Add(newMenu);
        fileMenu.Items.Add(manageMenu);
        fileMenu.Items.Add(new ToolStripSeparator());
        fileMenu.Items.Add(properties);
        fileMenu.Items.Add("Terminalde Aç").Click += (s, e) => OpenTerminalAt(pane.CurrentPath);
        fileMenu.Items.Add("Code ile Aç").Click += (s, e) => OpenCodeEditor(pane.CurrentPath);
        fileMenu.Items.Add(fileTasks);
        fileMenu.Items.Add(whatsapp);
        fileMenu.Items.Add(compression);
        var moreMenu = new ToolStripMenuItem("Diğer seçenekler");
        moreMenu.DropDownItems.AddRange(new ToolStripItem[] { openExplorer, copyPath });
        fileMenu.Items.Add(moreMenu);
        SetMenuGlyph(clipboardMenu, "\uE8C8"); SetMenuGlyph(cut, "\uE8C6"); SetMenuGlyph(copy, "\uE8C8"); SetMenuGlyph(paste, "\uE77F");
        SetMenuGlyph(openFile, "\uE8A5"); SetMenuGlyph(preview, "\uE890"); SetMenuGlyph(viewMenu, "\uE7F4");
        SetMenuGlyph(sortMenu, "\uE8CB"); SetMenuGlyph(groupMenu, "\uF168"); SetMenuGlyph(undoMenu, "\uE7A7");
        SetMenuGlyph(newMenu, "\uE710"); SetMenuGlyph(manageMenu, "\uE70F"); SetMenuGlyph(rename, "\uE8AC");
        SetMenuGlyph(delete, "\uE74D"); SetMenuGlyph(undoDelete, "\uE7A7"); SetMenuGlyph(properties, "\uE90F");
        SetMenuGlyph(fileTasks, "\uE8A5"); SetMenuGlyph(mail, "\uE715"); SetMenuGlyph(whatsapp, "\uE72D");
        SetMenuGlyph(compression, "\uE7B8"); SetMenuGlyph(zip, "\uE7B8"); SetMenuGlyph(extractHere, "\uE8B7");
        SetMenuGlyph(extractNamed, "\uE8B7"); SetMenuGlyph(extractChoose, "\uE8B7"); SetMenuGlyph(moreMenu, "\uE712");
        SetMenuGlyph(openExplorer, "\uE8B7"); SetMenuGlyph(copyPath, "\uE8C8");
        fileMenu.ShowItemToolTips = true;
        SetContextMenuTooltips(fileMenu.Items);
        pane.FileList.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            ListViewItem item = pane.FileList.HitTest(e.Location).Item;
            if (item != null && !item.Selected)
            {
                pane.FileList.SelectedItems.Clear();
                foreach (ListViewItem existing in pane.FileList.Items) existing.Checked = false;
                item.Selected = true;
                item.Checked = true;
                pane.FileList.FocusedItem = item;
            }
            activeTransferPane = pane;
        };
        fileMenu.Opening += (s, e) =>
        {
            if (pane.ShowingComputerDrives) { e.Cancel = true; return; }
            string[] selected = GetTransferSelectedPaths(pane);
            bool hasSelection = selected.Length > 0;
            openFile.Enabled = selected.Any(File.Exists);
            openExplorer.Enabled = selected.Length > 0;
            Entry firstEntry = pane.FileList.SelectedItems.Count > 0 ? pane.FileList.SelectedItems[0].Tag as Entry : null;
            preview.Enabled = firstEntry != null && File.Exists(firstEntry.FullPath);
            rename.Enabled = firstEntry != null && (File.Exists(firstEntry.FullPath) || Directory.Exists(firstEntry.FullPath));
            delete.Enabled = hasSelection;
            cut.Enabled = copy.Enabled = hasSelection;
            paste.Enabled = transferClipboardPaths.Count > 0;
            undoMove.Enabled = lastMoveRecords.Count > 0;
            undoDelete.Enabled = lastRecycleBinDeletedPaths.Count > 0;
            Entry contextEntry = firstEntry;
            mail.Enabled = whatsapp.Enabled = contextEntry != null && File.Exists(contextEntry.FullPath);
            string extension = contextEntry == null ? String.Empty : Path.GetExtension(contextEntry.FullPath);
            bool archive = contextEntry != null && File.Exists(contextEntry.FullPath) &&
                (String.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase) || String.Equals(extension, ".rar", StringComparison.OrdinalIgnoreCase));
            zip.Enabled = selected.Any(File.Exists);
            extractHere.Enabled = extractNamed.Enabled = extractChoose.Enabled = archive;
            compression.Enabled = zip.Enabled || archive;
            properties.Enabled = selected.Any(path => File.Exists(path) || Directory.Exists(path));
            copyPath.Enabled = contextEntry != null;
        };
        cut.Click += (s, e) => SetTransferClipboard(pane, true);
        copy.Click += (s, e) => SetTransferClipboard(pane, false);
        paste.Click += (s, e) => PasteTransferClipboard(pane, GetTransferDestination(pane));
        undoMove.Click += (s, e) => UndoLastMove();
        undoMenu.Click += (s, e) => UndoLastMove();
        undoDelete.Click += (s, e) => UndoLastRecycleBinDelete();
        mail.Click += (s, e) => { Entry item = GetSelectedTransferEntry(pane); if (item != null && File.Exists(item.FullPath)) SendByMail(item.FullPath); };
        whatsapp.Click += (s, e) => { Entry item = GetSelectedTransferEntry(pane); if (item != null && File.Exists(item.FullPath)) ShareWithWhatsAppSafeFileName(item.FullPath); };
        zip.Click += (s, e) => ZipTransferFiles(pane);
        extractHere.Click += (s, e) => ExtractSelectedArchive(GetSelectedTransferEntry(pane), ArchiveExtractionTarget.Here);
        extractNamed.Click += (s, e) => ExtractSelectedArchive(GetSelectedTransferEntry(pane), ArchiveExtractionTarget.ArchiveFolder);
        extractChoose.Click += (s, e) => ExtractSelectedArchive(GetSelectedTransferEntry(pane), ArchiveExtractionTarget.ChooseFolder);
        properties.Click += (s, e) => ShowSelectedItemsProperties(GetTransferSelectedPaths(pane));
        copyPath.Click += (s, e) => { Entry item = GetSelectedTransferEntry(pane); if (item != null) try { Clipboard.SetText(item.FullPath); } catch { } };
        openFile.Click += (s, e) =>
        {
            string path = GetTransferSelectedPaths(pane).FirstOrDefault();
            if (String.IsNullOrEmpty(path)) return;
            if (Directory.Exists(path)) TransferNavigate(pane, path); else if (File.Exists(path)) OpenFile(path);
        };
        openExplorer.Click += (s, e) =>
        {
            string path = GetTransferSelectedPaths(pane).FirstOrDefault();
            if (!String.IsNullOrEmpty(path)) OpenInWindowsExplorer(path);
        };
        preview.Click += (s, e) =>
        {
            Entry selected = pane.FileList.SelectedItems.Count > 0 ? pane.FileList.SelectedItems[0].Tag as Entry : null;
            if (selected != null && File.Exists(selected.FullPath)) new FilePreviewForm(selected.FullPath, true).Show(this);
        };
        rename.Click += (s, e) =>
        {
            Entry selected = pane.FileList.SelectedItems.Count > 0 ? pane.FileList.SelectedItems[0].Tag as Entry : null;
            RenameTransferEntry(pane, selected);
        };
        delete.Click += (s, e) => DeleteTransferFiles(pane);
        cutButton.Click += (s, e) => { fileMenu.Close(); SetTransferClipboard(pane, true); };
        copyButton.Click += (s, e) => { fileMenu.Close(); SetTransferClipboard(pane, false); };
        pasteButton.Click += (s, e) => { fileMenu.Close(); PasteTransferClipboard(pane, GetTransferDestination(pane)); };
        shareButton.Click += (s, e) => { fileMenu.Close(); Entry item = GetSelectedTransferEntry(pane); if (item != null) ShareWithWindows(item.FullPath); };
        deleteButton.Click += (s, e) => { fileMenu.Close(); DeleteTransferFiles(pane); };
        // The transfer pane is hosted inside a nested SplitContainer; show the menu
        // explicitly on mouse-up so it is reliable regardless of native ListView routing.
        pane.FileList.ContextMenuStrip = null;
        pane.FileList.MouseUp += (s, e) =>
        {
            if (e.Button != MouseButtons.Right || pane.ShowingComputerDrives) return;
            fileMenu.Show(pane.FileList, e.Location);
        };

        var folderMenu = new ContextMenuStrip();
        var enter = new ToolStripMenuItem("Klasöre gir");
        var openFolderExplorer = new ToolStripMenuItem("Windows Explorer ile Aç");
        var renameFolder = new ToolStripMenuItem("Yeniden adlandır");
        var cutFolder = new ToolStripMenuItem("Kes");
        var copyFolder = new ToolStripMenuItem("Kopyala");
        var pasteInto = new ToolStripMenuItem("Buraya yapıştır");
        var deleteFolder = new ToolStripMenuItem("Sil");
        folderMenu.Items.Add(cutFolder);
        folderMenu.Items.Add(copyFolder);
        folderMenu.Items.Add(pasteInto);
        folderMenu.Items.Add(renameFolder);
        folderMenu.Items.Add(deleteFolder);
        folderMenu.Items.Add(new ToolStripSeparator());
        folderMenu.Items.Add(openFolderExplorer);
        folderMenu.Items.Add(enter);
        folderMenu.Opening += (s, e) =>
        {
            if (pane.ShowingComputerDrives) { e.Cancel = true; return; }
            Entry selected = pane.FolderList.SelectedItem as Entry;
            enter.Enabled = selected != null && Directory.Exists(selected.FullPath);
            openFolderExplorer.Enabled = enter.Enabled;
            renameFolder.Enabled = selected != null && Directory.Exists(selected.FullPath) && !selected.IsParent;
            deleteFolder.Enabled = selected != null && Directory.Exists(selected.FullPath) && !selected.IsParent;
            pasteInto.Enabled = selected != null && Directory.Exists(selected.FullPath) && transferClipboardPaths.Count > 0;
            cutFolder.Enabled = copyFolder.Enabled = GetSelectedFolderPaths(pane.FolderList).Length > 0;
        };
        enter.Click += (s, e) =>
        {
            Entry selected = pane.FolderList.SelectedItem as Entry;
            if (selected != null && Directory.Exists(selected.FullPath)) TransferNavigate(pane, selected.FullPath);
        };
        openFolderExplorer.Click += (s, e) =>
        {
            Entry selected = pane.FolderList.SelectedItem as Entry;
            if (selected != null) OpenInWindowsExplorer(selected.FullPath);
        };
        renameFolder.Click += (s, e) => RenameTransferFolderSelection(pane);
        deleteFolder.Click += (s, e) => DeleteTransferFolderSelection(pane);
        pasteInto.Click += (s, e) =>
        {
            Entry selected = pane.FolderList.SelectedItem as Entry;
            if (selected != null) PasteTransferClipboard(pane, selected.FullPath);
        };
        cutFolder.Click += (s, e) => SetTransferFolderClipboard(pane, true);
        copyFolder.Click += (s, e) => SetTransferFolderClipboard(pane, false);
        pane.FolderList.ContextMenuStrip = folderMenu;
    }

    private void SetTransferClipboard(TransferPaneState pane, bool cut)
    {
        string[] selected = GetTransferSelectedPaths(pane);
        if (selected.Length == 0) return;
        try { SetFileClipboard(selected, cut); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, cut ? "Kes" : "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        transferClipboardPaths.Clear();
        transferClipboardPaths.AddRange(selected);
        transferClipboardIsCut = cut;
        pendingFolderClipboardPath = null;
        pendingCutPaths.Clear();
        if (cut) pendingCutPaths.AddRange(selected);
        RefreshTransferCutAppearance();
        RefreshCutFileAppearance();
        transferStatus.Text = selected.Length + " dosya/klasör " + (cut ? "taşınmak üzere kesildi." : "kopyalanmak üzere seçildi.");
    }

    private Entry GetSelectedTransferEntry(TransferPaneState pane)
    {
        ListViewItem item = pane == null || pane.FileList.SelectedItems.Count == 0 ? null : pane.FileList.SelectedItems[0];
        return item == null ? null : item.Tag as Entry;
    }

    private void ZipTransferFiles(TransferPaneState pane)
    {
        string[] paths = GetTransferSelectedPaths(pane).Where(File.Exists).ToArray();
        if (paths.Length == 0) return;
        string zipName = paths.Length == 1
            ? Path.GetFileNameWithoutExtension(paths[0]) + ".zip"
            : "SecilenDosyalar-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".zip";
        string zipPath = GetUnusedFilePath(Path.Combine(pane.CurrentPath, zipName));
        try
        {
            using (var output = new FileStream(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create))
                foreach (string path in paths)
                {
                    ZipArchiveEntry entry = archive.CreateEntry(Path.GetFileName(path), CompressionLevel.Optimal);
                    using (Stream input = File.OpenRead(path))
                    using (Stream target = entry.Open()) input.CopyTo(target);
                }
            TransferNavigate(pane, pane.CurrentPath);
            SelectTransferFiles(pane, new[] { zipPath });
        }
        catch (Exception ex)
        {
            try { if (File.Exists(zipPath)) File.Delete(zipPath); } catch { }
            MessageBox.Show(this, "ZIP arşivi oluşturulamadı.\r\n\r\n" + ex.Message, "ZIP oluştur", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void SetTransferFolderClipboard(TransferPaneState pane, bool cut)
    {
        string[] selected = GetSelectedFolderPaths(pane.FolderList);
        if (selected.Length == 0) return;
        transferClipboardPaths.Clear();
        transferClipboardPaths.AddRange(selected);
        transferClipboardIsCut = cut;
        pane.FolderList.Invalidate();
        transferStatus.Text = selected.Length + " klasör " + (cut ? "kesildi; hedefte Yapıştır seçin." : "kopyalanmak üzere seçildi.");
    }

    private void RenameTransferFolderSelection(TransferPaneState pane)
    {
        Entry selected = pane.FolderList.SelectedItem as Entry;
        if (selected == null || !Directory.Exists(selected.FullPath)) return;
        RenameTransferEntry(pane, selected);
    }

    private void DeleteTransferFolderSelection(TransferPaneState pane)
    {
        string[] paths = GetSelectedFolderPaths(pane == null ? null : pane.FolderList);
        if (paths.Length == 0) return;
        if (MessageBox.Show(this, paths.Length + " seçili klasör Geri Dönüşüm Kutusu'na taşınsın mı?", "Silme onayı",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        try
        {
            MovePathsToRecycleBin(paths);
            TransferNavigate(pane, pane.CurrentPath);
            foreach (string parent in paths.Select(Path.GetDirectoryName).Where(parent => !String.IsNullOrWhiteSpace(parent)).Distinct(StringComparer.OrdinalIgnoreCase))
                if (Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void RenameTransferEntry(TransferPaneState pane, Entry selected)
    {
        if (pane == null || selected == null || (!File.Exists(selected.FullPath) && !Directory.Exists(selected.FullPath))) return;
        bool isDirectory = Directory.Exists(selected.FullPath);
        string oldName = Path.GetFileName(selected.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        using (var dialog = new RenameFileDialog(oldName, isDirectory))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string newName = dialog.NewName.Trim();
            if (String.IsNullOrWhiteSpace(newName) || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || newName == "." || newName == "..")
            {
                MessageBox.Show(this, "Geçerli bir ad girin.", "Yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string parent = Path.GetDirectoryName(selected.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            string destination = Path.Combine(parent, newName);
            if (File.Exists(destination) || Directory.Exists(destination))
            {
                MessageBox.Show(this, "Bu adla başka bir öğe zaten var.", "Yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                if (isDirectory) Directory.Move(selected.FullPath, destination); else File.Move(selected.FullPath, destination);
                if (Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
                TransferNavigate(pane, pane.CurrentPath);
                if (String.Equals(currentPath, parent, StringComparison.OrdinalIgnoreCase)) Navigate(currentPath, null, destination, false);
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Yeniden adlandırılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }

    private void DeleteTransferFiles(TransferPaneState pane)
    {
        string[] paths = GetTransferSelectedPaths(pane);
        if (paths.Length == 0) return;
        DialogResult answer = MessageBox.Show(this, paths.Length + " seçili dosya/klasör Geri Dönüşüm Kutusu'na taşınsın mı?", "Silme onayı",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes) return;
        foreach (string path in paths)
            try { MovePathsToRecycleBin(new string[] { path }); }
            catch (Exception ex) { MessageBox.Show(this, path + "\r\n" + ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        TransferNavigate(pane, pane.CurrentPath);
    }

    private static void CopyDirectoryTree(string source, string destination)
    {
        Directory.CreateDirectory(destination);
        foreach (string file in Directory.GetFiles(source)) File.Copy(file, Path.Combine(destination, Path.GetFileName(file)), false);
        foreach (string directory in Directory.GetDirectories(source)) CopyDirectoryTree(directory, Path.Combine(destination, Path.GetFileName(directory)));
    }

    private static string[] GetTransferSelectedPaths(TransferPaneState pane)
    {
        return pane.FileList.Items.Cast<ListViewItem>()
            .Where(item => item.Selected || item.Checked)
            .Select(item => item.Tag as Entry)
            .Where(entry => entry != null && (File.Exists(entry.FullPath) || Directory.Exists(entry.FullPath)))
            .Select(entry => entry.FullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private string GetTransferDestination(TransferPaneState pane)
    {
        Entry folder = pane.FolderList.SelectedItem as Entry;
        return folder != null && Directory.Exists(folder.FullPath) ? folder.FullPath : pane.CurrentPath;
    }

    private void PasteTransferClipboard(TransferPaneState pane, string destination)
    {
        if (transferClipboardPaths.Count == 0 || !Directory.Exists(destination)) return;
        TransferFilesToFolder(pane, transferClipboardPaths.ToArray(), destination, transferClipboardIsCut, true);
    }

    private void TransferFilesToFolder(TransferPaneState destinationPane, string[] sources, string destinationFolder, bool move, bool updateCutClipboard)
    {
        if (sources == null || !Directory.Exists(destinationFolder)) return;
        if (move)
        {
            string destination = Path.GetFullPath(destinationFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string[] foldersToMove = sources.Where(Directory.Exists)
                .Where(source =>
                {
                    string fullSource = Path.GetFullPath(source).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    return !String.Equals(fullSource, destination, StringComparison.OrdinalIgnoreCase) && !IsPathWithin(destination, fullSource);
                }).ToArray();
            if (foldersToMove.Length > 0 && !ConfirmFolderMove(foldersToMove, destinationFolder)) return;
        }
        bool applyToAll = false;
        FileConflictChoice applyChoice = FileConflictChoice.Skip;
        var remainingCut = new List<string>();
        var completedDestinations = new List<string>();
        var completedMoves = new List<MoveRecord>();
        int completed = 0;
        foreach (string source in sources.Where(path => File.Exists(path) || Directory.Exists(path)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            string destination = Path.Combine(destinationFolder, Path.GetFileName(source));
            bool sourceIsDirectory = Directory.Exists(source);
            if (String.Equals(source, destination, StringComparison.OrdinalIgnoreCase))
            {
                if (move) { remainingCut.Add(source); continue; }
                destination = GetUnusedFilePath(destination);
            }
            if (sourceIsDirectory && (destination.StartsWith(source.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)))
            { if (move) remainingCut.Add(source); continue; }
            string backup = null;
            bool existingRemoved = false;
            try
            {
                if (File.Exists(destination) || Directory.Exists(destination))
                {
                    FileConflictChoice choice;
                    if (applyToAll) choice = applyChoice;
                    else
                    {
                        choice = FileConflictDialog.Ask(this, source, destination, out applyToAll);
                        if (applyToAll) applyChoice = choice;
                    }
                    if (choice == FileConflictChoice.Skip) { if (move) remainingCut.Add(source); continue; }
                    if (choice == FileConflictChoice.KeepBoth) destination = GetUnusedFilePath(destination);
                    else if (choice == FileConflictChoice.Replace)
                    {
                        if (sourceIsDirectory || Directory.Exists(destination))
                            MovePathsToRecycleBin(new string[] { destination });
                        else
                        {
                            string backupDirectory = Path.Combine(Path.GetDirectoryName(transferPanePathsFile), "TransferReplaceBackup");
                            Directory.CreateDirectory(backupDirectory);
                            backup = Path.Combine(backupDirectory, Guid.NewGuid().ToString("N") + ".bak");
                            File.Copy(destination, backup, false);
                            File.Delete(destination);
                            existingRemoved = true;
                        }
                    }
                }
                if (sourceIsDirectory)
                {
                    if (move) MoveDirectoryAcrossVolumes(source, destination);
                    else CopyDirectoryTree(source, destination);
                }
                else if (move) MoveFileAcrossVolumes(source, destination);
                else File.Copy(source, destination, false);
                completed++;
                completedDestinations.Add(destination);
                if (move) completedMoves.Add(new MoveRecord(source, destination, backup));
                else if (!String.IsNullOrEmpty(backup)) { try { File.Delete(backup); } catch { } }
            }
            catch (Exception ex)
            {
                if (move && (File.Exists(source) || Directory.Exists(source))) remainingCut.Add(source);
                if (!String.IsNullOrEmpty(backup))
                {
                    try
                    {
                        if (existingRemoved && !File.Exists(destination) && File.Exists(backup)) File.Move(backup, destination);
                        else File.Delete(backup);
                    }
                    catch { }
                }
                MessageBox.Show(this, "Dosya aktarılamadı:\r\n" + source + "\r\n\r\n" + ex.Message,
                    move ? "Taşı" : "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        if (move && updateCutClipboard)
        {
            transferClipboardPaths.Clear();
            transferClipboardPaths.AddRange(remainingCut);
            if (remainingCut.Count == 0) transferClipboardIsCut = false;
        }
        if (move && completedMoves.Count > 0)
        {
            ClearLastMoveUndo();
            lastMoveRecords.AddRange(completedMoves);
            lastMoveOriginView = Path.GetDirectoryName(completedMoves[0].SourcePath);
        }
        if (destinationPane != null) TransferNavigate(destinationPane, destinationFolder);
        else if (completedDestinations.Count > 0) Navigate(destinationFolder, null, completedDestinations[completedDestinations.Count - 1]);
        if (move)
        {
            foreach (MoveRecord record in completedMoves.Where(item => Directory.Exists(item.DestinationPath)))
            {
                string sourceParent = Path.GetDirectoryName(record.SourcePath);
                if (!String.IsNullOrWhiteSpace(sourceParent) && Directory.Exists(sourceParent))
                    RefreshFolderTreeChildrenAtPath(sourceParent);
            }
            RefreshFolderTreeChildrenAtPath(destinationFolder);
        }
        foreach (TransferPaneState sourcePane in new TransferPaneState[] { transferLeftPane, transferRightPane })
            if (sourcePane != null && sources.Any(path => String.Equals(Path.GetDirectoryName(path), sourcePane.CurrentPath, StringComparison.OrdinalIgnoreCase)))
                TransferNavigate(sourcePane, sourcePane.CurrentPath);
        if (destinationPane != null) SelectTransferFiles(destinationPane, completedDestinations);
        RefreshTransferCutAppearance();
        string message = completed + " öğe " + (move ? "taşındı." : "kopyalandı.") +
            (move && completedMoves.Count > 0 ? " Sağ tık menüsünden Geri Al Taşı ile geri alabilirsiniz." : String.Empty);
        if (destinationPane != null) transferStatus.Text = message;
        else status.Text = message;
    }

    private void SelectTransferFiles(TransferPaneState pane, IEnumerable<string> paths)
    {
        if (pane == null || pane.FileList == null) return;
        var destinationPaths = new HashSet<string>(paths ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        pane.SyncingSelection = true;
        try
        {
            foreach (ListViewItem item in pane.FileList.Items)
            {
                item.Selected = false;
                item.Checked = false;
            }
            foreach (ListViewItem item in pane.FileList.Items)
            {
                Entry entry = item.Tag as Entry;
                if (entry == null || !destinationPaths.Contains(entry.FullPath)) continue;
                item.Selected = true;
                item.Checked = true;
                pane.FileList.FocusedItem = item;
                pane.FileList.EnsureVisible(item.Index);
            }
        }
        finally { pane.SyncingSelection = false; }
    }

    private void StartTransferDrag(TransferPaneState pane)
    {
        string[] paths = GetTransferSelectedPaths(pane);
        if (paths.Length == 0) return;
        bool containsFolder = paths.Any(Directory.Exists);
        bool copyAtStart = containsFolder || transferCopyMode.Checked || (Control.ModifierKeys & Keys.Control) != 0;
        var data = new DataObject();
        data.SetData(DataFormats.FileDrop, paths);
        data.SetData("DosyaGezgini.TransferSourceSide", pane.Side);
        var feedback = new TransferDragFeedbackForm(paths.Length);
        GiveFeedbackEventHandler updateFeedback = (s, e) =>
        {
            bool copying = e.Effect == DragDropEffects.Copy ||
                (e.Effect != DragDropEffects.Move && ((Control.ModifierKeys & Keys.Control) != 0 || transferCopyMode.Checked));
            feedback.SetDragInfo(copying ? "Kopyalanıyor" : "Taşınıyor", Cursor.Position);
            e.UseDefaultCursors = true;
        };
        using (var preferredEffect = new MemoryStream(BitConverter.GetBytes(copyAtStart ? 1 : 2)))
        {
            data.SetData(DataFormats.GetFormat("Preferred DropEffect").Name, false, preferredEffect);
            pane.FileList.GiveFeedback += updateFeedback;
            try
            {
                feedback.Show(this);
                feedback.SetDragInfo(copyAtStart ? "Kopyalanıyor" : "Taşınıyor", Cursor.Position);
                pane.FileList.DoDragDrop(data, DragDropEffects.Copy | DragDropEffects.Move);
            }
            finally
            {
                pane.FileList.GiveFeedback -= updateFeedback;
                feedback.Close();
                feedback.Dispose();
            }
        }
    }

    private static void DrawTransferDocumentGlyph(Graphics graphics, Rectangle bounds)
    {
        Point[] paper =
        {
            new Point(bounds.Left + 7, bounds.Top + 2),
            new Point(bounds.Right - 11, bounds.Top + 2),
            new Point(bounds.Right - 3, bounds.Top + 10),
            new Point(bounds.Right - 3, bounds.Bottom - 2),
            new Point(bounds.Left + 7, bounds.Bottom - 2)
        };
        using (var fill = new SolidBrush(Color.White))
        using (var outline = new Pen(Color.FromArgb(58, 91, 128), 1.5F))
        using (var line = new Pen(Color.FromArgb(113, 139, 168), 1.2F))
        {
            graphics.FillPolygon(fill, paper);
            graphics.DrawPolygon(outline, paper);
            graphics.DrawLine(outline, bounds.Right - 11, bounds.Top + 2, bounds.Right - 11, bounds.Top + 10);
            graphics.DrawLine(outline, bounds.Right - 11, bounds.Top + 10, bounds.Right - 3, bounds.Top + 10);
            graphics.DrawLine(line, bounds.Left + 11, bounds.Top + 17, bounds.Right - 8, bounds.Top + 17);
            graphics.DrawLine(line, bounds.Left + 11, bounds.Top + 23, bounds.Right - 8, bounds.Top + 23);
        }
    }

    private void SetupTransferDropTargets(TransferPaneState pane, Control target, Control hitTarget, bool folderTarget)
    {
        target.AllowDrop = true;
        target.DragEnter += (s, e) => SetTransferDragEffect(pane, e);
        target.DragOver += (s, e) =>
        {
            SetTransferDragEffect(pane, e);
            if (e.Effect == DragDropEffects.None) return;
            pane.DropTargetPath = pane.CurrentPath;
            if (folderTarget && hitTarget is ListBox)
            {
                var list = (ListBox)hitTarget;
                Point point = list.PointToClient(new Point(e.X, e.Y));
                int index = list.IndexFromPoint(point);
                if (index >= 0)
                {
                    list.SelectedIndex = index;
                    Entry folder = list.Items[index] as Entry;
                    if (folder != null && Directory.Exists(folder.FullPath)) pane.DropTargetPath = folder.FullPath;
                }
            }
        };
        target.DragDrop += (s, e) =>
        {
            if (e.Effect == DragDropEffects.None) return;
            string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            string destination = pane.DropTargetPath;
            if (String.IsNullOrWhiteSpace(destination) || !Directory.Exists(destination)) destination = pane.CurrentPath;
            bool move = e.Effect != DragDropEffects.Copy;
            pane.DropTargetPath = null;
            TransferFilesToFolder(pane, paths, destination, move, false);
        };
    }

    private void SetTransferDragEffect(TransferPaneState pane, DragEventArgs e)
    {
        if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effect = DragDropEffects.None; return; }
        if (e.Data.GetDataPresent("DosyaGezgini.TransferSourceSide"))
        {
            object source = e.Data.GetData("DosyaGezgini.TransferSourceSide");
            if (source is int && (int)source == pane.Side) { e.Effect = DragDropEffects.None; return; }
        }
        string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[] ?? new string[0];
        bool containsFolder = paths.Any(Directory.Exists);
        bool copyRequested = containsFolder || (e.KeyState & 8) != 0 || transferCopyMode.Checked;
        e.Effect = copyRequested ? DragDropEffects.Copy : DragDropEffects.Move;
    }

    private void RefreshTransferCutAppearance()
    {
        foreach (TransferPaneState pane in new TransferPaneState[] { transferLeftPane, transferRightPane })
        {
            if (pane == null || pane.FileList == null) continue;
            foreach (ListViewItem item in pane.FileList.Items)
            {
                Entry entry = item.Tag as Entry;
                bool cut = transferClipboardIsCut && entry != null && transferClipboardPaths.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase);
                item.ForeColor = cut ? Color.Gray : pane.FileList.ForeColor;
            }
        }
    }

    private void SaveTransferPanePaths()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(transferPanePathsFile));
            File.WriteAllLines(transferPanePathsFile, new string[]
            {
                transferLeftPane == null ? String.Empty : StorePortablePath(transferLeftPane.CurrentPath),
                transferRightPane == null ? String.Empty : StorePortablePath(transferRightPane.CurrentPath),
                transferLeftPane != null && transferLeftPane.FolderPaneHidden ? "1" : "0",
                transferRightPane != null && transferRightPane.FolderPaneHidden ? "1" : "0"
            });
        }
        catch { }
    }

    private void InitializeTrayIcon()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("Dosya Gezgini'ni aç", null, (s, e) => RestoreFromFloatingIcon());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Çıkış", null, (s, e) =>
        {
            exitRequested = true;
            if (miniRestoreIcon != null) { miniRestoreIcon.Close(); miniRestoreIcon = null; }
            Close();
        });
        trayIcon = new NotifyIcon
        {
            Icon = Icon,
            Text = "Dosya Gezgini",
            ContextMenuStrip = menu,
            Visible = true
        };
        trayIcon.DoubleClick += (s, e) => RestoreFromFloatingIcon();
    }

    private static Icon CreateExplorerIcon()
    {
        var bitmap = new Bitmap(32, 32);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var folder = new SolidBrush(Color.FromArgb(255, 190, 45)))
        using (var edge = new Pen(Color.FromArgb(38, 66, 105), 1.5f))
        using (var paper = new SolidBrush(Color.White))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            g.FillRectangle(paper, 16, 3, 12, 16);
            g.DrawRectangle(edge, 16, 3, 12, 16);
            g.DrawLine(edge, 19, 8, 25, 8);
            g.DrawLine(edge, 19, 11, 25, 11);
            g.FillRectangle(folder, 3, 11, 26, 17);
            g.FillRectangle(folder, 5, 8, 10, 5);
            g.DrawRectangle(edge, 3, 11, 26, 17);
        }
        IntPtr handle = bitmap.GetHicon();
        Icon icon = (Icon)Icon.FromHandle(handle).Clone();
        DestroyIcon(handle);
        bitmap.Dispose();
        return icon;
    }

    private static Bitmap CreateToolbarFolderIcon()
    {
        var bitmap = new Bitmap(16, 16);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var folder = new SolidBrush(Color.FromArgb(255, 194, 55)))
        using (var tab = new SolidBrush(Color.FromArgb(255, 218, 105)))
        using (var outline = new Pen(Color.FromArgb(84, 92, 103), 1F))
        using (var path = new GraphicsPath())
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            g.FillRectangle(tab, 1, 3, 7, 4);
            path.AddPolygon(new Point[]
            {
                new Point(1, 6), new Point(6, 6), new Point(8, 8),
                new Point(15, 8), new Point(14, 14), new Point(1, 14)
            });
            g.FillPath(folder, path);
            g.DrawPath(outline, path);
            g.DrawLine(outline, 2, 9, 14, 9);
        }
        return bitmap;
    }

    private static Bitmap CreatePreviewButtonIcon()
    {
        var bitmap = new Bitmap(14, 14);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var lens = new Pen(Color.FromArgb(43, 85, 133), 1.6F))
        using (var detail = new SolidBrush(Color.FromArgb(84, 158, 196)))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            g.DrawEllipse(lens, 1.4F, 1.2F, 8.2F, 8.2F);
            g.FillEllipse(detail, 4.2F, 4.0F, 2.5F, 2.5F);
            g.DrawLine(lens, 8.5F, 8.5F, 12.6F, 12.6F);
        }
        return bitmap;
    }

    private static Bitmap CreateShortcutButtonIcon()
    {
        var bitmap = new Bitmap(14, 14);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var fill = new SolidBrush(Color.FromArgb(255, 193, 53)))
        using (var outline = new Pen(Color.FromArgb(118, 86, 20), 0.9F))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            PointF[] points = new PointF[10];
            for (int i = 0; i < points.Length; i++)
            {
                double angle = -Math.PI / 2.0 + i * Math.PI / 5.0;
                float radius = i % 2 == 0 ? 6.2F : 2.8F;
                points[i] = new PointF(7F + (float)Math.Cos(angle) * radius,
                    7F + (float)Math.Sin(angle) * radius);
            }
            g.FillPolygon(fill, points);
            g.DrawPolygon(outline, points);
        }
        return bitmap;
    }

    private static Bitmap CreateHistoryArrowIcon(bool forward)
    {
        var bitmap = new Bitmap(14, 14);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var pen = new Pen(Color.FromArgb(51, 74, 98), 1.7F))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            if (forward)
            {
                g.DrawLine(pen, 2, 7, 11, 7);
                g.DrawLine(pen, 7, 3, 11, 7);
                g.DrawLine(pen, 7, 11, 11, 7);
            }
            else
            {
                g.DrawLine(pen, 3, 7, 12, 7);
                g.DrawLine(pen, 3, 7, 7, 3);
                g.DrawLine(pen, 3, 7, 7, 11);
            }
        }
        return bitmap;
    }

    private static Bitmap CreateUpButtonIcon()
    {
        var bitmap = new Bitmap(14, 14);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var pen = new Pen(Color.FromArgb(51, 74, 98), 1.7F))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            g.DrawLine(pen, 7, 12, 7, 2);
            g.DrawLine(pen, 3, 6, 7, 2);
            g.DrawLine(pen, 11, 6, 7, 2);
        }
        return bitmap;
    }

    private static Bitmap CreateRefreshButtonIcon()
    {
        var bitmap = new Bitmap(14, 14);
        using (Graphics g = Graphics.FromImage(bitmap))
        using (var pen = new Pen(Color.FromArgb(51, 74, 98), 1.5F))
        using (var fill = new SolidBrush(Color.FromArgb(51, 74, 98)))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            g.DrawArc(pen, 2.2F, 2.2F, 9.2F, 9.2F, -55F, 285F);
            g.FillPolygon(fill, new Point[] { new Point(10, 1), new Point(13, 2), new Point(11, 5) });
        }
        return bitmap;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr handle);

    private static void StyleHeading(Label label)
    {
        label.Height = 34;
        label.TextAlign = ContentAlignment.MiddleLeft;
        label.Font = new Font("Segoe UI Semibold", 10F);
        label.ForeColor = Color.FromArgb(38, 51, 70);
    }

    private void StyleIconButton(Button button, string glyph, string tip, Font glyphFont)
    {
        button.Text = glyph;
        button.Font = glyphFont;
        button.TextAlign = ContentAlignment.MiddleCenter;
        button.Padding = new Padding(3);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        button.BackColor = Color.FromArgb(239, 245, 255);
        button.ForeColor = Color.FromArgb(38, 51, 70);
        button.AccessibleName = tip;
        actionTooltips.SetToolTip(button, tip);
    }

    private void OpenInWindowsExplorer(string path)
    {
        if (String.IsNullOrWhiteSpace(path) || (!Directory.Exists(path) && !File.Exists(path)))
        {
            status.Text = "Windows Explorer'da açılacak konum bulunamadı.";
            return;
        }
        try
        {
            string arguments = File.Exists(path) ? "/select,\"" + path + "\"" : "\"" + path + "\"";
            Process.Start(new ProcessStartInfo("explorer.exe", arguments) { UseShellExecute = true });
            MinimizeToFloatingIcon();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Windows Explorer", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void ShareWithWindows(string path)
    {
        if (String.IsNullOrWhiteSpace(path) || (!File.Exists(path) && !Directory.Exists(path))) return;
        try
        {
            // Let the Windows Shell prepare the same Unicode file item that File Explorer
            // sends to share targets. Some targets misread the WinRT storage-item route.
            if (TryShareUsingWindowsShell(path)) return;
            EnsureWindowsShareManager();
            windowsSharePendingPath = path;
            windowsSharePendingStorageItem = GetWindowsStorageItem(path);
            if (!windowsShareHandlerRegistered)
            {
                AddWindowsRuntimeShareHandler();
                windowsShareHandlerRegistered = true;
            }
            IDataTransferManagerInterop interop = GetDataTransferManagerInterop();
            Marshal.ThrowExceptionForHR(interop.ShowShareUIForWindow(Handle));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Windows paylaşım ekranı açılamadı.\r\n\r\n" + ex.Message,
                "Windows ile Paylaş", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void ShareWithWhatsAppSafeFileName(string path)
    {
        if (String.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
        string shareRoot = Path.Combine(Path.GetTempPath(), "DosyaGezgini-WhatsAppPaylasim");
        string operationFolder = null;
        try
        {
            Directory.CreateDirectory(shareRoot);
            foreach (string oldFolder in Directory.GetDirectories(shareRoot))
            {
                try
                {
                    if (DateTime.Now - Directory.GetLastWriteTime(oldFolder) > TimeSpan.FromDays(2)) Directory.Delete(oldFolder, true);
                }
                catch { }
            }
            operationFolder = Path.Combine(shareRoot, DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(operationFolder);
            string extension = Path.GetExtension(path);
            string stem = Path.GetFileNameWithoutExtension(path);
            string safeName = ToWhatsAppSafeAsciiName(stem) + extension;
            string copyPath = Path.Combine(operationFolder, safeName);
            File.Copy(path, copyPath, false);
            ShareWithWindows(copyPath);
            status.Text = "WhatsApp uyumlu geçici kopya paylaşıma hazır: " + safeName + " (2 günden eski kopyalar sonraki kullanımda temizlenir).";
        }
        catch (Exception ex)
        {
            if (!String.IsNullOrWhiteSpace(operationFolder))
            {
                try { if (Directory.Exists(operationFolder)) Directory.Delete(operationFolder, true); } catch { }
            }
            MessageBox.Show(this, "Paylaşım için geçici kopya hazırlanamadı. Özgün dosya değiştirilmedi.\r\n\r\n" + ex.Message,
                "WhatsApp uyumlu paylaşım", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static string ToWhatsAppSafeAsciiName(string value)
    {
        if (String.IsNullOrEmpty(value)) return "Dosya";
        string mapped = value
            .Replace('Ç', 'C').Replace('ç', 'c').Replace('Ğ', 'G').Replace('ğ', 'g')
            .Replace('İ', 'I').Replace('ı', 'i').Replace('Ö', 'O').Replace('ö', 'o')
            .Replace('Ş', 'S').Replace('ş', 's').Replace('Ü', 'U').Replace('ü', 'u')
            .Normalize(NormalizationForm.FormD);
        var output = new StringBuilder(mapped.Length);
        foreach (char ch in mapped)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) == UnicodeCategory.NonSpacingMark) continue;
            output.Append(ch <= 0x7f && !Path.GetInvalidFileNameChars().Contains(ch) ? ch : '_');
        }
        string result = Regex.Replace(output.ToString(), "_+", "_").Trim(' ', '.');
        return String.IsNullOrWhiteSpace(result) ? "Dosya" : result;
    }

    private static bool TryShareUsingWindowsShell(string path)
    {
        object shell = null;
        object folder = null;
        object item = null;
        object verbs = null;
        try
        {
            Type shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null) return false;
            shell = Activator.CreateInstance(shellType);
            folder = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, null, shell,
                new object[] { Path.GetDirectoryName(path) });
            if (folder == null) return false;
            item = folder.GetType().InvokeMember("ParseName", BindingFlags.InvokeMethod, null, folder,
                new object[] { Path.GetFileName(path) });
            if (item == null) return false;
            verbs = item.GetType().InvokeMember("Verbs", BindingFlags.GetProperty, null, item, null);
            if (verbs == null) return false;
            int count = Convert.ToInt32(verbs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, verbs, null), CultureInfo.InvariantCulture);
            for (int i = 0; i < count; i++)
            {
                object verb = verbs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, verbs, new object[] { i });
                if (verb == null) continue;
                string name = Convert.ToString(verb.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, verb, null), CultureInfo.CurrentCulture);
                string normalized = (name ?? String.Empty).Replace("&", String.Empty).Replace("…", String.Empty).Trim();
                if (!String.Equals(normalized, "Paylaş", StringComparison.CurrentCultureIgnoreCase) &&
                    !String.Equals(normalized, "Share", StringComparison.CurrentCultureIgnoreCase)) continue;
                verb.GetType().InvokeMember("DoIt", BindingFlags.InvokeMethod, null, verb, null);
                if (Marshal.IsComObject(verb)) Marshal.ReleaseComObject(verb);
                return true;
            }
            return false;
        }
        catch { return false; }
        finally
        {
            if (verbs != null && Marshal.IsComObject(verbs)) Marshal.ReleaseComObject(verbs);
            if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
            if (folder != null && Marshal.IsComObject(folder)) Marshal.ReleaseComObject(folder);
            if (shell != null && Marshal.IsComObject(shell)) Marshal.ReleaseComObject(shell);
        }
    }

    private void EnsureWindowsShareManager()
    {
        if (windowsShareManager != null) return;
        IDataTransferManagerInterop interop = GetDataTransferManagerInterop();
        Guid managerIid = new Guid("a5caee9b-8708-49d1-8d36-67d25a8da00c");
        IntPtr managerPointer = IntPtr.Zero;
        Marshal.ThrowExceptionForHR(interop.GetForWindow(Handle, ref managerIid, out managerPointer));
        if (managerPointer == IntPtr.Zero) throw new InvalidOperationException("Windows paylaşım yöneticisi alınamadı.");
        try { windowsShareManager = Marshal.GetObjectForIUnknown(managerPointer); }
        finally { Marshal.Release(managerPointer); }

        Type managerType = Type.GetType("Windows.ApplicationModel.DataTransfer.DataTransferManager, Windows.ApplicationModel, ContentType=WindowsRuntime", true);
        windowsShareDataRequestedEvent = managerType.GetEvent("DataRequested", BindingFlags.Public | BindingFlags.Instance);
        if (windowsShareDataRequestedEvent == null) throw new InvalidOperationException("Windows paylaşım verisi olayı bulunamadı.");
        Type handlerType = windowsShareDataRequestedEvent.EventHandlerType;
        Type tokenType = Type.GetType("System.Runtime.InteropServices.WindowsRuntime.EventRegistrationToken, System.Runtime.InteropServices.WindowsRuntime", true);
        Type addDelegateType = typeof(Func<,>).MakeGenericType(handlerType, tokenType);
        Type removeDelegateType = typeof(Action<>).MakeGenericType(tokenType);
        windowsShareAddEventHandler = Delegate.CreateDelegate(addDelegateType, windowsShareManager, windowsShareDataRequestedEvent.GetAddMethod());
        windowsShareRemoveEventHandler = Delegate.CreateDelegate(removeDelegateType, windowsShareManager, windowsShareDataRequestedEvent.GetRemoveMethod());
        ParameterInfo[] parameters = handlerType.GetMethod("Invoke").GetParameters();
        var sender = System.Linq.Expressions.Expression.Parameter(parameters[0].ParameterType, "sender");
        var args = System.Linq.Expressions.Expression.Parameter(parameters[1].ParameterType, "args");
        MethodInfo callback = GetType().GetMethod("HandleWindowsShareDataRequested", BindingFlags.Instance | BindingFlags.NonPublic);
        var call = System.Linq.Expressions.Expression.Call(System.Linq.Expressions.Expression.Constant(this), callback,
            System.Linq.Expressions.Expression.Convert(sender, typeof(object)),
            System.Linq.Expressions.Expression.Convert(args, typeof(object)));
        windowsShareEventHandler = System.Linq.Expressions.Expression.Lambda(handlerType, call, sender, args).Compile();
    }

    private void AddWindowsRuntimeShareHandler()
    {
        Type runtimeMarshal = Type.GetType("System.Runtime.InteropServices.WindowsRuntime.WindowsRuntimeMarshal, System.Runtime.InteropServices.WindowsRuntime", true);
        MethodInfo add = runtimeMarshal.GetMethods(BindingFlags.Public | BindingFlags.Static)
            .First(m => m.Name == "AddEventHandler" && m.IsGenericMethodDefinition && m.GetParameters().Length == 3)
            .MakeGenericMethod(windowsShareDataRequestedEvent.EventHandlerType);
        add.Invoke(null, new object[] { windowsShareAddEventHandler, windowsShareRemoveEventHandler, windowsShareEventHandler });
    }

    private void HandleWindowsShareDataRequested(object sender, object args)
    {
        string path = windowsSharePendingPath;
        if (String.IsNullOrWhiteSpace(path) || windowsSharePendingStorageItem == null) return;
        object storageItem = windowsSharePendingStorageItem;
        Type argsType = Type.GetType("Windows.ApplicationModel.DataTransfer.DataRequestedEventArgs, Windows.ApplicationModel, ContentType=WindowsRuntime", true);
        object request = argsType.GetProperty("Request").GetValue(args, null);
        object data = request.GetType().GetProperty("Data").GetValue(request, null);
        object properties = data.GetType().GetProperty("Properties").GetValue(data, null);
        properties.GetType().GetProperty("Title").SetValue(properties, Path.GetFileName(path), null);
        Type storageItemInterface = Type.GetType("Windows.Storage.IStorageItem, Windows.Storage, ContentType=WindowsRuntime", true);
        Array storageItems = Array.CreateInstance(storageItemInterface, 1);
        storageItems.SetValue(storageItem, 0);
        MethodInfo setStorageItems = data.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .First(m => m.Name == "SetStorageItems" && m.GetParameters().Length == 2);
        setStorageItems.Invoke(data, new object[] { storageItems, true });
    }

    private static object GetWindowsStorageItem(string path)
    {
        bool isDirectory = Directory.Exists(path);
        string storageTypeName = isDirectory
            ? "Windows.Storage.StorageFolder, Windows.Storage, ContentType=WindowsRuntime"
            : "Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime";
        Type storageType = Type.GetType(storageTypeName, true);
        string asyncMethodName = isDirectory ? "GetFolderFromPathAsync" : "GetFileFromPathAsync";
        object asyncOperation = storageType.GetMethod(asyncMethodName, BindingFlags.Public | BindingFlags.Static)
            .Invoke(null, new object[] { path });
        Type runtimeExtensions = Type.GetType("System.WindowsRuntimeSystemExtensions, System.Runtime.WindowsRuntime", true);
        MethodInfo asTask = runtimeExtensions.GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(m => m.Name == "AsTask" && m.IsGenericMethodDefinition && m.GetParameters().Length == 1)
            .First(m => m.GetParameters()[0].ParameterType.IsGenericType &&
                m.GetParameters()[0].ParameterType.GetGenericTypeDefinition().FullName == "Windows.Foundation.IAsyncOperation`1")
            .MakeGenericMethod(storageType);
        object task = asTask.Invoke(null, new object[] { asyncOperation });
        task.GetType().GetMethod("Wait", Type.EmptyTypes).Invoke(task, null);
        return task.GetType().GetProperty("Result").GetValue(task, null);
    }

    private static IDataTransferManagerInterop GetDataTransferManagerInterop()
    {
        const string runtimeClass = "Windows.ApplicationModel.DataTransfer.DataTransferManager";
        IntPtr className = IntPtr.Zero;
        IntPtr factory = IntPtr.Zero;
        try
        {
            Marshal.ThrowExceptionForHR(WindowsCreateString(runtimeClass, (uint)runtimeClass.Length, out className));
            Guid interopIid = new Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8");
            Marshal.ThrowExceptionForHR(RoGetActivationFactory(className, ref interopIid, out factory));
            return (IDataTransferManagerInterop)Marshal.GetTypedObjectForIUnknown(factory, typeof(IDataTransferManagerInterop));
        }
        finally
        {
            if (factory != IntPtr.Zero) Marshal.Release(factory);
            if (className != IntPtr.Zero) WindowsDeleteString(className);
        }
    }

    [ComImport, Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDataTransferManagerInterop
    {
        [PreserveSig] int GetForWindow(IntPtr appWindow, ref Guid riid, out IntPtr dataTransferManager);
        [PreserveSig] int ShowShareUIForWindow(IntPtr appWindow);
    }

    [DllImport("combase.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int WindowsCreateString(string sourceString, uint length, out IntPtr hString);
    [DllImport("combase.dll", PreserveSig = true)]
    private static extern int WindowsDeleteString(IntPtr hString);
    [DllImport("combase.dll", PreserveSig = true)]
    private static extern int RoGetActivationFactory(IntPtr activatableClassId, ref Guid iid, out IntPtr factory);
    private void ShowQuickLocationsMenu()
    {
        PopulateQuickLocationsMenu(quickLocationsMenu, path => Navigate(path), true);
        quickLocationsMenu.Show(quickLocations, new Point(0, quickLocations.Height));
    }

    private void PopulateQuickLocationsMenu(ContextMenuStrip menu, Action<string> navigate, bool includeGallery)
    {
        menu.Items.Clear();
        AddQuickLocation(menu, "Giriş", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), navigate);
        AddQuickLocation(menu, "Belgeler", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), navigate);
        AddQuickLocation(menu, "İndirilenler", GetDownloadsFolderPath(), navigate);
        AddQuickLocation(menu, "Masaüstü", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), navigate);
        AddQuickLocation(menu, "Resimler", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), navigate);
        AddQuickLocation(menu, "Müzikler", Environment.GetFolderPath(Environment.SpecialFolder.MyMusic), navigate);
        AddQuickLocation(menu, "Videolar", Environment.GetFolderPath(Environment.SpecialFolder.MyVideos), navigate);

        string oneDrive = Environment.GetEnvironmentVariable("OneDrive");
        if (String.IsNullOrWhiteSpace(oneDrive) || !Directory.Exists(oneDrive))
            oneDrive = Environment.GetEnvironmentVariable("OneDriveConsumer");
        if (!String.IsNullOrWhiteSpace(oneDrive)) AddQuickLocation(menu, "OneDrive", oneDrive, navigate);

        if (includeGallery)
        {
            if (menu.Items.Count > 0) menu.Items.Add(new ToolStripSeparator());
            ToolStripMenuItem gallery = new ToolStripMenuItem("Galeri");
            gallery.Click += (s, e) => OpenWindowsShellLocation("shell:Gallery");
            menu.Items.Add(gallery);
        }
    }

    private static string GetDownloadsFolderPath()
    {
        const string downloadsFolderId = "{374DE290-123F-4565-9164-39C4925E467B}";
        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"))
            {
                if (key != null)
                {
                    object value = key.GetValue(downloadsFolderId);
                    string configuredPath = value as string;
                    if (!String.IsNullOrWhiteSpace(configuredPath))
                    {
                        configuredPath = Environment.ExpandEnvironmentVariables(configuredPath);
                        if (Directory.Exists(configuredPath)) return configuredPath;
                    }
                }
            }
        }
        catch { }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
    }

    private void AddQuickLocation(ContextMenuStrip menu, string label, string path, Action<string> navigate)
    {
        if (String.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;
        foreach (ToolStripItem existing in menu.Items)
        {
            ToolStripMenuItem menuItem = existing as ToolStripMenuItem;
            if (menuItem != null && String.Equals(menuItem.Tag as string, path, StringComparison.OrdinalIgnoreCase)) return;
        }
        var item = new ToolStripMenuItem(label) { Tag = path };
        item.Click += (s, e) => navigate(path);
        menu.Items.Add(item);
    }

    private void OpenWindowsShellLocation(string shellLocation)
    {
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", shellLocation) { UseShellExecute = true });
            MinimizeToFloatingIcon();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Windows konumu", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private string[] GetSelectedOrCheckedFilePaths()
    {
        IEnumerable<ListViewItem> items = files.CheckedItems.Count > 0
            ? files.CheckedItems.Cast<ListViewItem>()
            : files.SelectedItems.Cast<ListViewItem>();
        return items.Select(row => row.Tag as Entry)
            .Where(entry => entry != null && File.Exists(entry.FullPath))
            .Select(entry => entry.FullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private string[] GetSelectedOrCheckedFileAndFolderPaths()
    {
        IEnumerable<ListViewItem> items = files.CheckedItems.Cast<ListViewItem>()
            .Concat(files.SelectedItems.Cast<ListViewItem>())
            .Distinct();
        return items.Select(row => row.Tag as Entry)
            .Where(entry => entry != null && (File.Exists(entry.FullPath) || Directory.Exists(entry.FullPath)))
            .Select(entry => entry.FullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static Entry GetSelectedEntry(ListView list)
    {
        return list != null && list.SelectedItems.Count > 0 ? list.SelectedItems[0].Tag as Entry : null;
    }

    private static string[] GetSelectedFilePaths(ListView list)
    {
        if (list == null) return new string[0];
        return list.SelectedItems.Cast<ListViewItem>()
            .Select(row => row.Tag as Entry)
            .Where(entry => entry != null && File.Exists(entry.FullPath))
            .Select(entry => entry.FullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string[] GetSelectedListPaths(ListView list)
    {
        if (list == null) return new string[0];
        return list.SelectedItems.Cast<ListViewItem>()
            .Select(row => row.Tag as Entry)
            .Where(entry => entry != null && (File.Exists(entry.FullPath) || Directory.Exists(entry.FullPath)))
            .Select(entry => entry.FullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string[] GetSelectedFolderPaths(ListBox list)
    {
        if (list == null) return new string[0];
        return list.SelectedItems.Cast<object>().OfType<Entry>()
            .Where(entry => !entry.IsParent && Directory.Exists(entry.FullPath) &&
                !String.Equals(entry.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase))
            .Select(entry => entry.FullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private void SetFolderListClipboard(ListBox list, bool cut)
    {
        string[] selected = GetSelectedFolderPaths(list);
        if (selected.Length == 0) return;
        try
        {
            SetFileClipboard(selected, cut);
            pendingFolderClipboardPath = null;
            pendingCutPaths.Clear();
            if (cut) pendingCutPaths.AddRange(selected);
            RefreshCutFileAppearance();
            list.Invalidate();
            status.Text = selected.Length + " klasör " + (cut ? "kesildi." : "kopyalandı.");
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, cut ? "Kes" : "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void PasteClipboardIntoFolder(string destinationFolder)
    {
        if (!String.IsNullOrWhiteSpace(pendingFolderClipboardPath) && Directory.Exists(pendingFolderClipboardPath))
        {
            PasteFolderClipboard(new Entry(String.Empty, destinationFolder));
            return;
        }
        if (String.IsNullOrWhiteSpace(destinationFolder) || !Directory.Exists(destinationFolder)) return;
        try
        {
            string[] paths = Clipboard.GetFileDropList().Cast<string>()
                .Where(path => !String.IsNullOrWhiteSpace(path) && (File.Exists(path) || Directory.Exists(path)))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (paths.Length == 0) return;
            bool moving = ClipboardRequestsMove();
            TransferFilesToFolder(null, paths, destinationFolder, moving, false);
            if (moving && paths.All(path => !File.Exists(path) && !Directory.Exists(path)))
            {
                try { Clipboard.Clear(); } catch { }
                pendingCutPaths.Clear();
            }
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void ConfigureFileListKeyboard(ListView list, Action open, Action delete, Action rename, Action cut, Action copy, Action paste)
    {
        list.KeyDown += (s, e) =>
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                foreach (ListViewItem item in list.Items) { item.Selected = true; if (list.CheckBoxes) item.Checked = true; }
                if (list.Items.Count > 0) list.FocusedItem = list.Items[list.Items.Count - 1];
            }
            else if (e.Control && e.KeyCode == Keys.C) copy();
            else if (e.Control && e.KeyCode == Keys.X) cut();
            else if (e.Control && e.KeyCode == Keys.V) paste();
            else if (e.KeyCode == Keys.Enter) open();
            else if (e.KeyCode == Keys.Delete) delete();
            else if (e.KeyCode == Keys.F2) rename();
            else return;
            e.Handled = e.SuppressKeyPress = true;
        };
    }

    private void CopyListSelection(ListView list)
    {
        string[] selected = GetSelectedListPaths(list);
        if (selected.Length == 0) return;
        try
        {
            SetFileClipboard(selected, false);
            pendingCutPaths.Clear();
            pendingFolderClipboardPath = null;
            RefreshCutFileAppearance();
            status.Text = selected.Length + " dosya/klasör kopyalamak için panoya alındı.";
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void CutListSelection(ListView list)
    {
        string[] selected = GetSelectedListPaths(list);
        if (selected.Length == 0) return;
        try
        {
            SetFileClipboard(selected, true);
            pendingCutPaths.Clear();
            pendingCutPaths.AddRange(selected);
            pendingFolderClipboardPath = null;
            RefreshCutFileAppearance();
            foreach (Control folderControl in new Control[] { folders, transferLeftPane == null ? null : transferLeftPane.FolderList, transferRightPane == null ? null : transferRightPane.FolderList })
                if (folderControl != null) folderControl.Invalidate();
            status.Text = selected.Length + " dosya/klasör kesildi. Hedef Gezgin konumunda Yapıştır seçin.";
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Kes", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void OpenSelectedListFile(ListView list)
    {
        Entry selected = GetSelectedEntry(list);
        if (selected == null) return;
        if (File.Exists(selected.FullPath)) OpenFile(selected.FullPath);
        else if (Directory.Exists(selected.FullPath)) RevealInExplorer(selected.FullPath);
        else MessageBox.Show(this, "Bu öğenin hedefi bulunamadı:\n" + selected.FullPath,
            "Hedef bulunamadı", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void CopySelectedFiles()
    {
        string[] selected = GetSelectedOrCheckedFilePaths();
        if (selected.Length == 0) return;
        try
        {
            SetFileClipboard(selected, false);
            pendingCutPaths.Clear();
            RefreshCutFileAppearance();
            status.Text = selected.Length + " dosya kopyalamak için panoya alındı.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Dosyalar panoya kopyalanamadı.\r\n\r\n" + ex.Message,
                "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void CutSelectedFiles()
    {
        string[] selected = GetSelectedOrCheckedFilePaths();
        if (selected.Length == 0) return;
        try
        {
            SetFileClipboard(selected, true);
            pendingCutPaths.Clear();
            pendingCutPaths.AddRange(selected);
            RefreshCutFileAppearance();
            status.Text = selected.Length + " dosya kesildi. Hedef klasörde Yapıştır seçin.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Dosyalar kesilmek üzere panoya alınamadı.\r\n\r\n" + ex.Message,
                "Kes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void CopySelectedExplorerItems(bool cut)
    {
        string[] selected = GetSelectedOrCheckedFileAndFolderPaths();
        if (selected.Length == 0) return;
        try
        {
            pendingFolderClipboardPath = null;
            SetFileClipboard(selected, cut);
            pendingCutPaths.Clear();
            if (cut) pendingCutPaths.AddRange(selected.Where(File.Exists));
            RefreshCutFileAppearance();
            foreach (ListViewItem row in files.Items)
            {
                Entry entry = row.Tag as Entry;
                if (entry != null && Directory.Exists(entry.FullPath))
                    row.ForeColor = cut && selected.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase) ? Color.Gray : files.ForeColor;
            }
            status.Text = selected.Length + (cut ? " öğe kesildi. Dosyalar bölümünde Yapıştır seçin." : " öğe kopyalamak için panoya alındı.");
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, cut ? "Kes" : "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void SetFileClipboard(IEnumerable<string> paths, bool move)
    {
        var data = new DataObject();
        var fileDrop = new StringCollection();
        fileDrop.AddRange(paths.Where(path => !String.IsNullOrWhiteSpace(path)).ToArray());
        data.SetFileDropList(fileDrop);
        if (clipboardDropEffectStream != null) clipboardDropEffectStream.Dispose();
        clipboardDropEffectStream = new MemoryStream(BitConverter.GetBytes(move ? 2 : 1));
        data.SetData(DataFormats.GetFormat("Preferred DropEffect").Name, false, clipboardDropEffectStream);
        Clipboard.SetDataObject(data, true);
    }

    private static bool ClipboardRequestsMove()
    {
        try
        {
            object effect = Clipboard.GetData(DataFormats.GetFormat("Preferred DropEffect").Name);
            if (effect is MemoryStream)
            {
                using (var stream = new MemoryStream(((MemoryStream)effect).ToArray()))
                using (var reader = new BinaryReader(stream)) return (reader.ReadInt32() & 2) != 0;
            }
            if (effect is byte[] && ((byte[])effect).Length >= 4)
                return (BitConverter.ToInt32((byte[])effect, 0) & 2) != 0;
            if (effect is int) return (((int)effect) & 2) != 0;
        }
        catch { }
        return false;
    }

    private void RefreshCutFileAppearance()
    {
        foreach (ListViewItem item in files.Items)
        {
            Entry entry = item.Tag as Entry;
            if (entry == null) continue;
            bool cut = pendingCutPaths.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase);
            item.ForeColor = cut ? Color.Gray : files.ForeColor;
        }
    }

    private void PasteFiles()
    {
        if (!String.IsNullOrWhiteSpace(pendingFolderClipboardPath) && Directory.Exists(pendingFolderClipboardPath))
        {
            PasteFolderClipboard(null);
            return;
        }
        StringCollection fileDrop;
        try { fileDrop = Clipboard.GetFileDropList(); }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Pano okunamadı.\r\n\r\n" + ex.Message,
                "Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (fileDrop == null || fileDrop.Count == 0) return;

        bool isCut = ClipboardRequestsMove();
        string[] clipboardPaths = fileDrop.Cast<string>()
            .Where(path => !String.IsNullOrWhiteSpace(path) && (File.Exists(path) || Directory.Exists(path)))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (clipboardPaths.Any(Directory.Exists))
        {
            TransferFilesToFolder(null, clipboardPaths, currentPath, isCut, false);
            if (isCut && clipboardPaths.All(path => !File.Exists(path) && !Directory.Exists(path)))
            {
                try { Clipboard.Clear(); } catch { }
                pendingCutPaths.Clear();
            }
            RefreshCutFileAppearance();
            return;
        }
        bool applyToAll = false;
        FileConflictChoice applyChoice = FileConflictChoice.Skip;
        int copied = 0;
        string lastCopiedPath = null;
        var remainingCutPaths = new List<string>();
        var movedRecords = new List<MoveRecord>();
        foreach (string source in fileDrop.Cast<string>())
        {
            if (String.IsNullOrWhiteSpace(source) || !File.Exists(source)) continue;
            string destination = Path.Combine(currentPath, Path.GetFileName(source));
            if (String.Equals(source, destination, StringComparison.OrdinalIgnoreCase))
            {
                if (isCut) remainingCutPaths.Add(source);
                else destination = GetUnusedFilePath(destination);
                if (isCut) continue;
            }
            string backup = null;
            try
            {
                if (File.Exists(destination))
                {
                    FileConflictChoice choice;
                    if (applyToAll) choice = applyChoice;
                    else
                    {
                        choice = FileConflictDialog.Ask(this, source, destination, out applyToAll);
                        if (applyToAll) applyChoice = choice;
                    }
                    if (choice == FileConflictChoice.Skip) { if (isCut) remainingCutPaths.Add(source); continue; }
                    if (choice == FileConflictChoice.KeepBoth) destination = GetUnusedFilePath(destination);
                    if (isCut)
                    {
                        if (choice == FileConflictChoice.Replace)
                        {
                            string backupDirectory = Path.Combine(Path.GetDirectoryName(historyFile), "MoveUndo");
                            Directory.CreateDirectory(backupDirectory);
                            backup = Path.Combine(backupDirectory, Guid.NewGuid().ToString("N") + ".bak");
                            MoveFileAcrossVolumes(destination, backup);
                        }
                        MoveFileAcrossVolumes(source, destination);
                    }
                    else File.Copy(source, destination, choice == FileConflictChoice.Replace);
                }
                else if (isCut) MoveFileAcrossVolumes(source, destination);
                else File.Copy(source, destination, false);
                copied++;
                lastCopiedPath = destination;
                if (isCut) movedRecords.Add(new MoveRecord(source, destination, backup));
            }
            catch (Exception ex)
            {
                if (isCut && File.Exists(source)) remainingCutPaths.Add(source);
                if (!String.IsNullOrEmpty(backup) && File.Exists(backup) && !File.Exists(destination))
                {
                    try { MoveFileAcrossVolumes(backup, destination); } catch { }
                }
                MessageBox.Show(this, "Dosya kopyalanamadı:\r\n" + source + "\r\n\r\n" + ex.Message,
                    "Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        if (movedRecords.Count > 0)
        {
            ClearLastMoveUndo();
            lastMoveRecords.AddRange(movedRecords);
            lastMoveOriginView = Path.GetDirectoryName(movedRecords[0].SourcePath);
        }
        if (isCut)
        {
            pendingCutPaths.Clear();
            pendingCutPaths.AddRange(remainingCutPaths);
            try
            {
                if (remainingCutPaths.Count == 0) Clipboard.Clear();
                else SetFileClipboard(remainingCutPaths, true);
            }
            catch { }
            RefreshCutFileAppearance();
        }
        if (copied > 0) Navigate(currentPath, null, lastCopiedPath);
        else RefreshFileRows();
        status.Text = copied + (isCut ? " dosya taşındı." : " dosya kopyalandı.") +
            (movedRecords.Count > 0 ? " Sağ tık menüsünden Geri Al Taşı ile geri alabilirsiniz." : String.Empty);
    }

    private void MoveSelectedFiles()
    {
        string[] selected = GetSelectedOrCheckedFilePaths();
        if (selected.Length == 0) return;

        using (var dialog = new FolderBrowserDialog())
        {
            dialog.Description = "Seçili dosyaların taşınacağı klasörü seçin.";
            dialog.SelectedPath = currentPath;
            dialog.ShowNewFolderButton = true;
            if (dialog.ShowDialog(this) != DialogResult.OK || !Directory.Exists(dialog.SelectedPath)) return;
            MoveFilesToFolder(selected, dialog.SelectedPath);
        }
    }

    private void MoveSelectedExplorerItems()
    {
        string[] selected = GetSelectedOrCheckedFileAndFolderPaths();
        if (selected.Length == 0) return;
        using (var dialog = new FolderBrowserDialog())
        {
            dialog.Description = "Seçili dosya ve klasörlerin taşınacağı klasörü seçin.";
            dialog.SelectedPath = currentPath;
            dialog.ShowNewFolderButton = true;
            if (dialog.ShowDialog(this) != DialogResult.OK || !Directory.Exists(dialog.SelectedPath)) return;
            TransferFilesToFolder(null, selected, dialog.SelectedPath, true, false);
        }
    }

    private static bool IsInternalFileMove(IDataObject data)
    {
        if (data == null || !data.GetDataPresent(DataFormats.FileDrop) ||
            !data.GetDataPresent("DosyaGezgini.InternalFileMove")) return false;
        string[] paths = data.GetData(DataFormats.FileDrop) as string[];
        return paths != null && paths.Length > 0;
    }

    private static DragDropEffects GetFileDropEffect(DragEventArgs e)
    {
        if (e == null || e.Data == null || !HasDroppedFiles(e.Data)) return DragDropEffects.None;
        bool physicalFiles = false;
        try { physicalFiles = e.Data.GetDataPresent(DataFormats.FileDrop); } catch { }
        if (!physicalFiles)
            return (e.AllowedEffect & DragDropEffects.Copy) != 0 ? DragDropEffects.Copy : DragDropEffects.None;
        if (IsInternalFileMove(e.Data))
        {
            string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[] ?? new string[0];
            if (paths.Any(Directory.Exists)) return (e.AllowedEffect & DragDropEffects.Copy) != 0 ? DragDropEffects.Copy : DragDropEffects.None;
            if ((e.KeyState & 8) != 0) return (e.AllowedEffect & DragDropEffects.Copy) != 0 ? DragDropEffects.Copy : DragDropEffects.None;
            return (e.AllowedEffect & DragDropEffects.Move) != 0 ? DragDropEffects.Move : DragDropEffects.None;
        }
        if ((e.KeyState & 8) != 0)
            return (e.AllowedEffect & DragDropEffects.Copy) != 0 ? DragDropEffects.Copy : DragDropEffects.None;
        if ((e.AllowedEffect & DragDropEffects.Move) != 0) return DragDropEffects.Move;
        if ((e.AllowedEffect & DragDropEffects.Copy) != 0) return DragDropEffects.Copy;
        return DragDropEffects.None;
    }

    private static bool HasDroppedFiles(IDataObject data)
    {
        if (data == null) return false;
        try { if (data.GetDataPresent(DataFormats.FileDrop)) return true; }
        catch { }
        try
        {
            string[] formats = data.GetFormats(false);
            if (formats.Any(format => String.Equals(format, "FileGroupDescriptorW", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(format, "FileGroupDescriptor", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(format, "FileNameW", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(format, "FileName", StringComparison.OrdinalIgnoreCase))) return true;
        }
        catch { }
        return HasDataFormat(data, "FileGroupDescriptorW") || HasDataFormat(data, "FileGroupDescriptor") ||
            HasDataFormat(data, "FileNameW") || HasDataFormat(data, "FileName");
    }

    private static bool HasDataFormat(IDataObject data, string format)
    {
        try { return data.GetDataPresent(format, false); }
        catch { return false; }
    }

    private static string GetDropFormatsDescription(IDataObject data)
    {
        if (data == null) return "veri yok";
        try
        {
            string[] formats = data.GetFormats(false);
            string description = String.Join(", ", formats.Take(12).ToArray());
            if (formats.Length > 12) description += ", …";
            return String.IsNullOrWhiteSpace(description) ? "format listesi boş" : description;
        }
        catch (Exception ex) { return "format listesi okunamadı (" + ex.GetType().Name + ")"; }
    }

    private static string[] GetFileDropPaths(IDataObject data)
    {
        object value = data.GetData(DataFormats.FileDrop);
        string[] paths = value as string[];
        if (paths != null) return paths;
        var collection = value as StringCollection;
        if (collection != null) return collection.Cast<string>().ToArray();
        string singlePath = value as string;
        if (!String.IsNullOrWhiteSpace(singlePath)) return new[] { singlePath };
        return GetOleShellFileDropPaths(data);
    }

    private static string[] GetOleShellFileDropPaths(IDataObject data)
    {
        var oleData = data as System.Runtime.InteropServices.ComTypes.IDataObject;
        if (oleData == null) return new string[0];
        var formatEtc = new System.Runtime.InteropServices.ComTypes.FORMATETC
        {
            cfFormat = 15,
            dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL
        };
        System.Runtime.InteropServices.ComTypes.STGMEDIUM medium;
        try { oleData.GetData(ref formatEtc, out medium); }
        catch { return new string[0]; }
        try
        {
            if (medium.tymed != System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL) return new string[0];
            IntPtr handle = medium.unionmember;
            uint count = DragQueryFile(handle, UInt32.MaxValue, null, 0);
            var paths = new List<string>();
            for (uint index = 0; index < count; index++)
            {
                uint length = DragQueryFile(handle, index, null, 0);
                if (length == 0 || length > 32767) continue;
                var buffer = new StringBuilder((int)length + 1);
                if (DragQueryFile(handle, index, buffer, (uint)buffer.Capacity) > 0) paths.Add(buffer.ToString());
            }
            return paths.ToArray();
        }
        finally { ReleaseStgMedium(ref medium); }
    }

    private void ProcessExplorerDrop(IDataObject data, string destination, bool move)
    {
        if (data == null || !Directory.Exists(destination)) return;
        try
        {
            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] paths = GetFileDropPaths(data);
                if (paths != null) TransferExplorerDrop(paths, destination, move);
                if (paths != null && paths.Length > 0) return;
            }

            string[] saved = SaveVirtualFileDrop(data, destination);
            if (saved.Length == 0) return;
            if (String.Equals(destination, currentPath, StringComparison.OrdinalIgnoreCase))
            {
                RefreshFileRows();
                foreach (ListViewItem item in files.Items)
                {
                    Entry entry = item.Tag as Entry;
                    if (entry != null && saved.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase))
                    {
                        item.Selected = true;
                        item.Checked = true;
                        files.FocusedItem = item;
                    }
                }
            }
            else RefreshFolderTreeChildrenAtPath(Path.GetDirectoryName(destination));
            status.Text = saved.Length + " mail eki dosya olarak kaydedildi.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Mail uygulamasından sürüklenen ek alınamadı.\r\n\r\n" + ex.GetBaseException().Message +
                "\r\n\r\nAktarım biçimleri: " + GetDropFormatsDescription(data),
                "Dosya bırakma", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static string[] SaveVirtualFileDrop(IDataObject data, string destination)
    {
        string[] formats;
        try { formats = data.GetFormats(false); } catch { formats = new string[0]; }
        bool unicodeNames = HasDataFormat(data, "FileGroupDescriptorW") ||
            formats.Any(format => String.Equals(format, "FileGroupDescriptorW", StringComparison.OrdinalIgnoreCase));
        bool hasAnsiNames = HasDataFormat(data, "FileGroupDescriptor") ||
            formats.Any(format => String.Equals(format, "FileGroupDescriptor", StringComparison.OrdinalIgnoreCase));
        bool hasDescriptors = unicodeNames || hasAnsiNames;
        var namesAndSizes = new List<Tuple<string, long>>();
        if (hasDescriptors)
        {
            string descriptorFormat = unicodeNames ? "FileGroupDescriptorW" : "FileGroupDescriptor";
            byte[] descriptor = ReadClipboardPayload(data.GetData(descriptorFormat, false));
            if (descriptor.Length < 4) throw new InvalidDataException("Mail uygulaması ek dosyalarının listesini göndermedi.");
            int count = BitConverter.ToInt32(descriptor, 0);
            int descriptorSize = unicodeNames ? 592 : 332;
            if (count < 1 || count > 1000 || descriptor.Length < 4L + (long)count * descriptorSize)
                throw new InvalidDataException("Mail uygulamasından gelen dosya listesi geçersiz.");
            for (int index = 0; index < count; index++)
            {
                int offset = 4 + index * descriptorSize;
                string name = ReadVirtualFileName(descriptor, offset + 72, unicodeNames);
                long size = ((long)BitConverter.ToUInt32(descriptor, offset + 64) << 32) |
                    BitConverter.ToUInt32(descriptor, offset + 68);
                namesAndSizes.Add(Tuple.Create(name, size));
            }
        }
        else
        {
            bool unicodeName = HasDataFormat(data, "FileNameW") ||
                formats.Any(format => String.Equals(format, "FileNameW", StringComparison.OrdinalIgnoreCase));
            string nameFormat = unicodeName ? "FileNameW" : "FileName";
            if (!HasDataFormat(data, nameFormat) && !formats.Any(format => String.Equals(format, nameFormat, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException("Mail uygulaması dosya adını ek biçiminde paylaşmadı.");
            object rawName = data.GetData(nameFormat, false);
            string name = rawName as string;
            if (name == null)
            {
                byte[] nameBytes = ReadClipboardPayload(rawName);
                name = unicodeName ? Encoding.Unicode.GetString(nameBytes).TrimEnd('\0') : Encoding.Default.GetString(nameBytes).TrimEnd('\0');
            }
            namesAndSizes.Add(Tuple.Create(name, 0L));
        }

        var created = new List<string>();
        var oleData = data as System.Runtime.InteropServices.ComTypes.IDataObject;
        if (oleData == null)
            throw new NotSupportedException("Bu mail uygulaması ek dosyalarını Windows OLE akışıyla paylaşmadı.");

        ushort contentsFormat = RegisterClipboardFormat("FileContents");
        if (contentsFormat == 0) throw new InvalidOperationException("Windows dosya eki aktarım biçimi kullanılamıyor.");

        try
        {
            for (int index = 0; index < namesAndSizes.Count; index++)
            {
                string name = namesAndSizes[index].Item1;
                name = Path.GetFileName(name);
                if (String.IsNullOrWhiteSpace(name) || name == "." || name == "..") name = "MailEki-" + (index + 1) + ".dat";
                foreach (char invalid in Path.GetInvalidFileNameChars()) name = name.Replace(invalid, '_');

                long declaredSize = namesAndSizes[index].Item2;
                string output = Path.Combine(destination, name);
                if (File.Exists(output) || Directory.Exists(output)) output = GetUnusedFilePath(output);
                bool completed = false;
                try
                {
                    SaveVirtualFileContent(oleData, contentsFormat, index, output, declaredSize);
                    completed = true;
                    created.Add(output);
                }
                finally
                {
                    if (!completed) { try { if (File.Exists(output)) File.Delete(output); } catch { } }
                }
            }
        }
        catch
        {
            foreach (string path in created) { try { File.Delete(path); } catch { } }
            throw;
        }
        return created.ToArray();
    }

    private static byte[] ReadClipboardPayload(object payload)
    {
        byte[] bytes = payload as byte[];
        if (bytes != null) return bytes;
        var stream = payload as Stream;
        if (stream == null) throw new InvalidDataException("Mail uygulamasının ek listesi okunamadı.");
        using (var copy = new MemoryStream())
        {
            stream.CopyTo(copy);
            return copy.ToArray();
        }
    }

    private static string ReadVirtualFileName(byte[] descriptor, int offset, bool unicode)
    {
        if (unicode)
        {
            int end = offset;
            int limit = Math.Min(descriptor.Length, offset + 520);
            while (end + 1 < limit && (descriptor[end] != 0 || descriptor[end + 1] != 0)) end += 2;
            return Encoding.Unicode.GetString(descriptor, offset, end - offset);
        }
        int ansiEnd = offset;
        int ansiLimit = Math.Min(descriptor.Length, offset + 260);
        while (ansiEnd < ansiLimit && descriptor[ansiEnd] != 0) ansiEnd++;
        return Encoding.Default.GetString(descriptor, offset, ansiEnd - offset);
    }

    private static void SaveVirtualFileContent(System.Runtime.InteropServices.ComTypes.IDataObject oleData,
        ushort format, int index, string output, long declaredSize)
    {
        var formatEtc = new System.Runtime.InteropServices.ComTypes.FORMATETC
        {
            cfFormat = unchecked((short)format),
            dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT,
            lindex = index,
            ptd = IntPtr.Zero,
            tymed = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTREAM |
                System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL
        };
        System.Runtime.InteropServices.ComTypes.STGMEDIUM medium;
        oleData.GetData(ref formatEtc, out medium);
        try
        {
            using (var outputStream = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                if (medium.tymed == System.Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTREAM)
                {
                    var input = (System.Runtime.InteropServices.ComTypes.IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                    IntPtr readCount = Marshal.AllocHGlobal(sizeof(int));
                    try
                    {
                        byte[] buffer = new byte[64 * 1024];
                        while (true)
                        {
                            Marshal.WriteInt32(readCount, 0);
                            input.Read(buffer, buffer.Length, readCount);
                            int read = Marshal.ReadInt32(readCount);
                            if (read <= 0) break;
                            outputStream.Write(buffer, 0, read);
                        }
                    }
                    finally { Marshal.FreeHGlobal(readCount); }
                }
                else if (medium.tymed == System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL)
                {
                    ulong available = GlobalSize(medium.unionmember).ToUInt64();
                    long remaining = declaredSize > 0 ? Math.Min((long)available, declaredSize) : (long)available;
                    IntPtr locked = GlobalLock(medium.unionmember);
                    if (locked == IntPtr.Zero && remaining > 0) throw new IOException("Mail ekinin içeriği bellekte açılamadı.");
                    try
                    {
                        byte[] buffer = new byte[64 * 1024];
                        long position = 0;
                        while (position < remaining)
                        {
                            int length = (int)Math.Min(buffer.Length, remaining - position);
                            Marshal.Copy(IntPtr.Add(locked, checked((int)position)), buffer, 0, length);
                            outputStream.Write(buffer, 0, length);
                            position += length;
                        }
                    }
                    finally { if (locked != IntPtr.Zero) GlobalUnlock(medium.unionmember); }
                }
                else throw new NotSupportedException("Mail uygulaması ek içeriğini desteklenmeyen biçimde sundu.");
            }
        }
        finally { ReleaseStgMedium(ref medium); }
    }

    private void TransferExplorerDrop(string[] paths, string destination, bool move)
    {
        if (paths != null && paths.Any(path => File.Exists(path) || Directory.Exists(path)))
        {
            TransferFilesToFolder(null, paths, destination, move, false);
            return;
        }
        if (move)
        {
            MoveFilesToFolder(paths, destination);
            return;
        }
        if (paths == null || paths.Length == 0 || !Directory.Exists(destination)) return;
        bool applyToAll = false;
        FileConflictChoice applyChoice = FileConflictChoice.Skip;
        var completed = new List<string>();
        int skipped = 0;
        foreach (string source in paths.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            string target = Path.Combine(destination, Path.GetFileName(source));
            if (String.Equals(source, target, StringComparison.OrdinalIgnoreCase)) target = GetUnusedFilePath(target);
            try
            {
                if (File.Exists(target))
                {
                    FileConflictChoice choice;
                    if (applyToAll) choice = applyChoice;
                    else
                    {
                        choice = FileConflictDialog.Ask(this, source, target, out applyToAll);
                        if (applyToAll) applyChoice = choice;
                    }
                    if (choice == FileConflictChoice.Skip) { skipped++; continue; }
                    if (choice == FileConflictChoice.KeepBoth) target = GetUnusedFilePath(target);
                }
                File.Copy(source, target, File.Exists(target));
                completed.Add(target);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Dosya kopyalanamadı:\r\n" + source + "\r\n\r\n" + ex.Message,
                    "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        if (completed.Count > 0) Navigate(destination, null, completed[completed.Count - 1]);
        status.Text = completed.Count + " dosya kopyalandı." + (skipped > 0 ? " " + skipped + " dosya atlandı." : String.Empty);
    }

    private void MoveFilesToFolder(string[] selected, string targetFolder)
    {
        if (selected == null || selected.Length == 0 || !Directory.Exists(targetFolder)) return;
        if (selected.All(path => String.Equals(Path.GetDirectoryName(path), targetFolder, StringComparison.OrdinalIgnoreCase)))
        {
            status.Text = "Seçili dosyalar zaten bu klasörde.";
            return;
        }

        bool applyToAll = false;
        FileConflictChoice applyChoice = FileConflictChoice.Skip;
        var moved = new List<MoveRecord>();
        int skipped = 0;
        string lastDestination = null;
        foreach (string source in selected.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            string destination = Path.Combine(targetFolder, Path.GetFileName(source));
            if (String.Equals(source, destination, StringComparison.OrdinalIgnoreCase)) continue;
            string backup = null;
            bool replacedDestinationRemoved = false;
            try
            {
                if (File.Exists(destination))
                {
                    FileConflictChoice choice;
                    if (applyToAll) choice = applyChoice;
                    else
                    {
                        choice = FileConflictDialog.Ask(this, source, destination, out applyToAll);
                        if (applyToAll) applyChoice = choice;
                    }
                    if (choice == FileConflictChoice.Skip) { skipped++; continue; }
                    if (choice == FileConflictChoice.KeepBoth) destination = GetUnusedFilePath(destination);
                    else if (choice == FileConflictChoice.Replace)
                    {
                        string backupDirectory = Path.Combine(Path.GetDirectoryName(historyFile), "MoveUndo");
                        Directory.CreateDirectory(backupDirectory);
                        backup = Path.Combine(backupDirectory, Guid.NewGuid().ToString("N") + ".bak");
                        File.Copy(destination, backup, false);
                        File.Delete(destination);
                        replacedDestinationRemoved = true;
                    }
                }
                MoveFileAcrossVolumes(source, destination);
                moved.Add(new MoveRecord(source, destination, backup));
                lastDestination = destination;
            }
            catch (Exception ex)
            {
                if (!String.IsNullOrEmpty(backup))
                {
                    try
                    {
                        if (replacedDestinationRemoved && !File.Exists(destination) && File.Exists(backup))
                            File.Move(backup, destination);
                        else File.Delete(backup);
                    }
                    catch { }
                }
                MessageBox.Show(this, "Dosya taşınamadı:\r\n" + source + "\r\n\r\n" + ex.Message,
                    "Taşı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        if (moved.Count == 0)
        {
            status.Text = "Taşınan dosya yok.";
            return;
        }
        ClearLastMoveUndo();
        lastMoveRecords.AddRange(moved);
        lastMoveOriginView = currentPath;
        Navigate(targetFolder, null, lastDestination);
        status.Text = moved.Count + " dosya taşındı. Sağ tık menüsünden Geri Al ile geri alabilirsiniz." +
            (skipped > 0 ? " " + skipped + " dosya atlandı." : String.Empty);
    }

    private void UndoLastMove()
    {
        if (lastMoveRecords.Count == 0) return;
        var remaining = new List<MoveRecord>();
        var restoredPaths = new List<string>();
        int restored = 0;
        for (int index = lastMoveRecords.Count - 1; index >= 0; index--)
        {
            MoveRecord record = lastMoveRecords[index];
            try
            {
                if (File.Exists(record.SourcePath) || Directory.Exists(record.SourcePath))
                    throw new IOException("Kaynak konumda aynı adlı öğe zaten var: " + record.SourcePath);
                bool destinationIsDirectory = Directory.Exists(record.DestinationPath);
                bool destinationExists = File.Exists(record.DestinationPath) || destinationIsDirectory;
                bool backupExists = !String.IsNullOrEmpty(record.BackupPath) && File.Exists(record.BackupPath);
                if (!destinationExists && !backupExists)
                    throw new FileNotFoundException("Taşınan dosya hedef konumda bulunamadı.", record.DestinationPath);
                if (destinationExists)
                {
                    if (destinationIsDirectory) MoveDirectoryAcrossVolumes(record.DestinationPath, record.SourcePath);
                    else MoveFileAcrossVolumes(record.DestinationPath, record.SourcePath);
                }
                if (backupExists)
                {
                    if (File.Exists(record.DestinationPath))
                        throw new IOException("Yedek dosya geri yüklenemedi; hedefte aynı adlı dosya var.");
                    try { MoveFileAcrossVolumes(record.BackupPath, record.DestinationPath); }
                    catch
                    {
                        if (destinationExists && File.Exists(record.SourcePath) && !File.Exists(record.DestinationPath))
                            MoveFileAcrossVolumes(record.SourcePath, record.DestinationPath);
                        throw;
                    }
                }
                restoredPaths.Add(record.SourcePath);
                restored++;
            }
            catch (Exception ex)
            {
                remaining.Add(record);
                MessageBox.Show(this, "Taşıma geri alınamadı:\r\n" + record.DestinationPath + "\r\n\r\n" + ex.Message,
                    "Geri Al", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        lastMoveRecords.Clear();
        lastMoveRecords.AddRange(remaining);
        string returnPath = lastMoveOriginView;
        if (String.IsNullOrWhiteSpace(returnPath) || !Directory.Exists(returnPath)) returnPath = currentPath;
        string selectPath = restoredPaths.Count > 0 ? restoredPaths[0] : null;
        Navigate(returnPath, null, selectPath);
        foreach (TransferPaneState pane in new TransferPaneState[] { transferLeftPane, transferRightPane })
        {
            if (pane == null || String.IsNullOrWhiteSpace(pane.CurrentPath) || !Directory.Exists(pane.CurrentPath)) continue;
            bool containsMoveSource = restoredPaths.Any(path => String.Equals(Path.GetDirectoryName(path), pane.CurrentPath, StringComparison.OrdinalIgnoreCase));
            bool containsMoveDestination = lastMoveRecords.Any(record => String.Equals(Path.GetDirectoryName(record.DestinationPath), pane.CurrentPath, StringComparison.OrdinalIgnoreCase));
            if (containsMoveSource || containsMoveDestination)
            {
                TransferNavigate(pane, pane.CurrentPath);
                SelectTransferFiles(pane, restoredPaths.Where(path => String.Equals(Path.GetDirectoryName(path), pane.CurrentPath, StringComparison.OrdinalIgnoreCase)));
            }
        }
        if (lastMoveRecords.Count == 0)
        {
            ClearLastMoveUndo();
            status.Text = restored + " dosya için son taşıma geri alındı.";
        }
        else status.Text = restored + " dosya geri alındı; " + lastMoveRecords.Count + " işlem geri alınamadı.";
    }

    private void ClearLastMoveUndo()
    {
        foreach (MoveRecord record in lastMoveRecords)
            if (!String.IsNullOrEmpty(record.BackupPath)) { try { File.Delete(record.BackupPath); } catch { } }
        lastMoveRecords.Clear();
        lastMoveOriginView = null;
    }

    private static void MoveFileAcrossVolumes(string source, string destination)
    {
        string destinationDirectory = Path.GetDirectoryName(destination);
        if (!Directory.Exists(destinationDirectory)) Directory.CreateDirectory(destinationDirectory);
        try
        {
            File.Move(source, destination);
        }
        catch (IOException)
        {
            if (!File.Exists(source) || File.Exists(destination)) throw;
            File.Copy(source, destination, false);
            try { File.Delete(source); }
            catch
            {
                try { File.Delete(destination); } catch { }
                throw;
            }
        }
    }

    private static void MoveDirectoryAcrossVolumes(string source, string destination)
    {
        try { Directory.Move(source, destination); }
        catch (IOException)
        {
            if (!Directory.Exists(source) || Directory.Exists(destination) || File.Exists(destination)) throw;
            CopyDirectoryTree(source, destination);
            try { Directory.Delete(source, true); }
            catch
            {
                try { if (Directory.Exists(destination)) Directory.Delete(destination, true); } catch { }
                throw;
            }
        }
    }

    private static string GetUnusedFilePath(string path)
    {
        string directory = Path.GetDirectoryName(path);
        string stem = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);
        for (int index = 1; ; index++)
        {
            string candidate = Path.Combine(directory, stem + " (" + index + ")" + extension);
            if (!File.Exists(candidate) && !Directory.Exists(candidate)) return candidate;
        }
    }

    private void RenameSelectedFile(Entry selected)
    {
        if (selected == null || !File.Exists(selected.FullPath)) return;
        string oldName = Path.GetFileName(selected.FullPath);
        using (var dialog = new RenameFileDialog(oldName))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string newName = dialog.NewName.Trim();
            if (newName.Length == 0 || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                newName == "." || newName == "..")
            {
                MessageBox.Show(this, "Geçerli bir dosya adı girin.", "Yeniden adlandır",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string destination = Path.Combine(Path.GetDirectoryName(selected.FullPath), newName);
            if (String.Equals(destination, selected.FullPath, StringComparison.OrdinalIgnoreCase)) return;
            if (File.Exists(destination))
            {
                MessageBox.Show(this, "Bu adla başka bir dosya zaten var.", "Yeniden adlandır",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                File.Move(selected.FullPath, destination);
                Navigate(currentPath, null, destination);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Dosya yeniden adlandırılamadı.\r\n\r\n" + ex.Message,
                    "Yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    private void SetFolderClipboard(Entry entry, bool cut)
    {
        if (entry == null || !Directory.Exists(entry.FullPath)) return;
        pendingFolderClipboardPath = entry.FullPath;
        pendingFolderClipboardIsCut = cut;
        status.Text = "Klasör " + (cut ? "kesildi; hedef klasörde sağ tık > Yapıştır seçin." : "kopyalandı; hedef klasörde sağ tık > Yapıştır seçin.");
    }

    private bool ConfirmFolderMove(IEnumerable<string> folderPaths, string destinationFolder)
    {
        string[] paths = (folderPaths ?? Enumerable.Empty<string>()).Where(Directory.Exists).ToArray();
        if (paths.Length == 0) return true;
        string names = String.Join("\r\n", paths.Take(6).Select(path => "• " + Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))).ToArray());
        if (paths.Length > 6) names += "\r\n… ve " + (paths.Length - 6) + " klasör daha";
        string message = paths.Length + " klasör şu konuma taşınacak:\r\n\r\n" + names +
            "\r\n\r\nHedef: " + destinationFolder + "\r\n\r\nTaşıma işlemi yapılsın mı?";
        return MessageBox.Show(this, message, "Klasör taşıma onayı", MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
    }

    private void PasteFolderClipboard(Entry target)
    {
        string source = pendingFolderClipboardPath;
        string destinationFolder = target != null && Directory.Exists(target.FullPath) ? target.FullPath : currentPath;
        if (String.IsNullOrWhiteSpace(source) || !Directory.Exists(source) || !Directory.Exists(destinationFolder)) return;
        if (pendingFolderClipboardIsCut && String.Equals(Path.GetDirectoryName(source), destinationFolder, StringComparison.OrdinalIgnoreCase)) return;
        if (String.Equals(source, destinationFolder, StringComparison.OrdinalIgnoreCase) || destinationFolder.StartsWith(source.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return;
        bool moving = pendingFolderClipboardIsCut;
        if (moving && !ConfirmFolderMove(new[] { source }, destinationFolder)) return;
        string destination = Path.Combine(destinationFolder, Path.GetFileName(source));
        if (Directory.Exists(destination) || File.Exists(destination)) destination = GetUnusedFilePath(destination);
        try
        {
            if (moving) Directory.Move(source, destination);
            else CopyDirectoryTree(source, destination);
            if (moving) pendingFolderClipboardPath = null;
            Navigate(currentPath, null, null, false);
            if (moving)
            {
                string sourceParent = Path.GetDirectoryName(source);
                if (!String.IsNullOrWhiteSpace(sourceParent) && Directory.Exists(sourceParent))
                    RefreshFolderTreeChildrenAtPath(sourceParent);
            }
            RefreshFolderTreeChildrenAtPath(destinationFolder);
            status.Text = moving ? "Klasör taşındı." : "Klasör kopyalandı.";
        }
        catch (Exception ex) { MessageBox.Show(this, "Klasör aktarılamadı.\r\n\r\n" + ex.Message, "Klasör işlemi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void RenameSelectedFolder(Entry selected)
    {
        if (selected == null || selected.IsParent || !Directory.Exists(selected.FullPath)) return;
        string oldName = Path.GetFileName(selected.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        using (var dialog = new RenameFileDialog(oldName, true))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string newName = dialog.NewName.Trim();
            if (newName.Length == 0 || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || newName == "." || newName == "..")
            {
                MessageBox.Show(this, "Geçerli bir klasör adı girin.", "Klasörü yeniden adlandır",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string parent = Path.GetDirectoryName(selected.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            if (String.IsNullOrWhiteSpace(parent)) return;
            string destination = Path.Combine(parent, newName);
            if (String.Equals(destination, selected.FullPath, StringComparison.OrdinalIgnoreCase)) return;
            if (Directory.Exists(destination) || File.Exists(destination))
            {
                MessageBox.Show(this, "Bu adla başka bir dosya veya klasör zaten var.", "Klasörü yeniden adlandır",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                Directory.Move(selected.FullPath, destination);
                RemapRenamedFolderReferences(selected.FullPath, destination);
                // Keep the renamed folder visible and selected in its parent directory.
                Navigate(parent, new Entry(newName, destination));
                for (int i = 0; i < folders.Items.Count; i++)
                {
                    Entry folder = folders.Items[i] as Entry;
                    if (folder != null && String.Equals(folder.FullPath, destination, StringComparison.OrdinalIgnoreCase))
                    {
                        folders.SelectedIndex = i;
                        break;
                    }
                }
                RefreshFolderTreeChildrenAtPath(parent);
                SelectFolderTreePath(destination);
                folderTree.Invalidate();
                status.Text = "Klasör yeniden adlandırıldı: " + newName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Klasör yeniden adlandırılamadı. Klasör başka bir uygulamada açık veya kullanılıyor olabilir.\r\n\r\n" + ex.Message,
                    "Klasörü yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    private static string RemapFolderPath(string path, string oldRoot, string newRoot)
    {
        if (String.IsNullOrWhiteSpace(path) || !IsPathWithin(path, oldRoot)) return path;
        try
        {
            string fullPath = Path.GetFullPath(path);
            string fullOld = Path.GetFullPath(oldRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string fullNew = Path.GetFullPath(newRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (String.Equals(fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), fullOld, StringComparison.OrdinalIgnoreCase))
                return fullNew;
            string suffix = fullPath.Substring(fullOld.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return Path.Combine(fullNew, suffix);
        }
        catch { return path; }
    }

    private void RemapRenamedFolderReferences(string oldPath, string newPath)
    {
        currentPath = RemapFolderPath(currentPath, oldPath, newPath);
        for (int index = 0; index < navigationHistory.Count; index++)
            navigationHistory[index] = RemapFolderPath(navigationHistory[index], oldPath, newPath);

        for (int index = 0; index < shortcuts.Count; index++)
        {
            Entry entry = shortcuts[index];
            string updated = RemapFolderPath(entry.FullPath, oldPath, newPath);
            if (!String.Equals(updated, entry.FullPath, StringComparison.OrdinalIgnoreCase))
                shortcuts[index] = new Entry(DisplayName(updated), updated, entry.AddedAt);
        }
        for (int index = 0; index < history.Count; index++)
        {
            HistoryEntry entry = history[index];
            string updated = RemapFolderPath(entry.FullPath, oldPath, newPath);
            if (!String.Equals(updated, entry.FullPath, StringComparison.OrdinalIgnoreCase))
                history[index] = new HistoryEntry(updated, entry.OpenedAt);
        }
        foreach (TabPage page in folderTabs.TabPages)
        {
            FolderTabState state = page.Tag as FolderTabState;
            if (state == null) continue;
            state.Path = RemapFolderPath(state.Path, oldPath, newPath);
            state.History = state.History.Select(path => RemapFolderPath(path, oldPath, newPath)).ToList();
            page.Text = String.IsNullOrWhiteSpace(state.CustomName) ? GetFolderTabTitle(state.Path) : state.CustomName;
            page.ToolTipText = state.Path;
        }
        foreach (WorkProfile profile in workProfiles)
        {
            string oldRoot = profile.RootPath;
            List<string> portablePaths = profile.PortableRelatedPaths
                .Select(relative => ResolveRelativeWorkPath(oldRoot, relative)).ToList();
            profile.RootPath = RemapFolderPath(profile.RootPath, oldPath, newPath);
            profile.LastPath = RemapFolderPath(profile.LastPath, oldPath, newPath);
            for (int index = 0; index < profile.RelatedFolders.Count; index++)
                profile.RelatedFolders[index] = RemapFolderPath(profile.RelatedFolders[index], oldPath, newPath);
            profile.PortableRelatedPaths.Clear();
            foreach (string portablePath in portablePaths)
            {
                string remapped = RemapFolderPath(portablePath, oldPath, newPath);
                string relative;
                if (!String.IsNullOrWhiteSpace(remapped) && TryGetRelativeWorkPath(profile.RootPath, remapped, out relative) &&
                    !profile.PortableRelatedPaths.Contains(relative, StringComparer.OrdinalIgnoreCase))
                    profile.PortableRelatedPaths.Add(relative);
            }
        }
        SaveShortcuts();
        SaveHistory();
        SaveNavigationHistory();
        SaveFolderTabSessions();
        SaveWorkProfiles();
        SaveWorkProfileLocations();
        RefreshShortcutList();
        RefreshHistoryList();
        RefreshWorkProfileList();
        RefreshWorkProfileDetails();
    }

    private void ZipSelectedFiles()
    {
        string[] selected = GetSelectedOrCheckedFilePaths();
        if (selected.Length == 0) return;
        string zipName = selected.Length == 1
            ? Path.GetFileNameWithoutExtension(selected[0]) + ".zip"
            : "SecilenDosyalar-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".zip";
        string zipPath = GetUnusedFilePath(Path.Combine(currentPath, zipName));
        try
        {
            using (var output = new FileStream(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create))
            {
                foreach (string path in selected)
                {
                    ZipArchiveEntry entry = archive.CreateEntry(Path.GetFileName(path), CompressionLevel.Optimal);
                    using (Stream input = File.OpenRead(path))
                    using (Stream target = entry.Open()) input.CopyTo(target);
                }
            }
        }
        catch (Exception ex)
        {
            try { if (File.Exists(zipPath)) File.Delete(zipPath); } catch { }
            MessageBox.Show(this, "ZIP arşivi oluşturulamadı.\r\n\r\n" + ex.Message,
                "ZIP oluştur", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        Navigate(currentPath, null, zipPath);
        AddToHistory(zipPath);
        try { Process.Start(new ProcessStartInfo(zipPath) { UseShellExecute = true }); }
        catch (Exception ex)
        {
            MessageBox.Show(this, "ZIP dosyası oluşturuldu ancak varsayılan ZIP uygulaması açılamadı.\r\n\r\n" + ex.Message,
                "ZIP oluştur", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void ExtractSelectedArchive(Entry selected, ArchiveExtractionTarget extractionTarget)
    {
        if (selected == null || !File.Exists(selected.FullPath)) return;
        string extension = Path.GetExtension(selected.FullPath);
        bool isZip = String.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase);
        bool isRar = String.Equals(extension, ".rar", StringComparison.OrdinalIgnoreCase);
        if (!isZip && !isRar) return;

        string destination = currentPath;
        if (extractionTarget == ArchiveExtractionTarget.ArchiveFolder)
        {
            destination = Path.Combine(currentPath, Path.GetFileNameWithoutExtension(selected.FullPath));
            if (Directory.Exists(destination))
            {
                DialogResult answer = MessageBox.Show(this,
                    "'" + Path.GetFileName(destination) + "' klasörü zaten var. Arşiv bu klasöre ayıklansın mı?",
                    "Ayıklama konumu", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes) return;
            }
            else Directory.CreateDirectory(destination);
        }
        else if (extractionTarget == ArchiveExtractionTarget.ChooseFolder)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Arşivin ayıklanacağı klasörü seçin.";
                dialog.SelectedPath = currentPath;
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                destination = dialog.SelectedPath;
            }
        }

        string winRar = FindWinRarExecutable();
        if (!String.IsNullOrWhiteSpace(winRar))
        {
            try
            {
                string targetWithSeparator = destination.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                var start = new ProcessStartInfo
                {
                    FileName = winRar,
                    Arguments = "x " + QuoteProcessArgument(selected.FullPath) + " " + QuoteProcessArgument(targetWithSeparator),
                    UseShellExecute = false,
                    CreateNoWindow = false,
                    WorkingDirectory = Path.GetDirectoryName(winRar)
                };
                Process.Start(start);
                if (!String.Equals(destination, currentPath, StringComparison.OrdinalIgnoreCase)) Navigate(destination);
                status.Text = "WinRAR arşiv ayıklamayı başlattı: " + Path.GetFileName(selected.FullPath);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "WinRAR ile ayıklama başlatılamadı.\r\n\r\n" + ex.Message,
                    "Arşiv ayıklama", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        if (isRar)
        {
            MessageBox.Show(this, "RAR arşivlerini açmak için WinRAR kurulmalıdır. Bu bilgisayarda WinRAR bulunamadı.",
                "WinRAR bulunamadı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            int extracted = ExtractZipArchive(selected.FullPath, destination);
            if (!String.Equals(destination, currentPath, StringComparison.OrdinalIgnoreCase)) Navigate(destination);
            else Navigate(currentPath, null, null, false);
            status.Text = extracted + " öğe ZIP arşivinden ayıklandı.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "ZIP arşivi ayıklanamadı.\r\n\r\n" + ex.Message,
                "Arşiv ayıklama", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static string FindWinRarExecutable()
    {
        string[] registryKeys = new string[]
        {
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe"
        };
        foreach (string keyPath in registryKeys)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(keyPath))
                {
                    string registered = key == null ? null : key.GetValue(null) as string;
                    if (!String.IsNullOrWhiteSpace(registered))
                    {
                        registered = Environment.ExpandEnvironmentVariables(registered.Trim().Trim('"'));
                        if (File.Exists(registered)) return registered;
                    }
                }
            }
            catch { }
        }
        string[] candidates = new string[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WinRAR", "WinRAR.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "WinRAR", "WinRAR.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "WinRAR", "WinRAR.exe")
        };
        return candidates.FirstOrDefault(File.Exists);
    }

    private static string QuoteProcessArgument(string value)
    {
        return "\"" + (value ?? String.Empty).Replace("\\\"", "\\\\\"") + "\"";
    }

    private int ExtractZipArchive(string archivePath, string destination)
    {
        string root = Path.GetFullPath(destination).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        Directory.CreateDirectory(destination);
        bool applyToAll = false;
        FileConflictChoice applyChoice = FileConflictChoice.Skip;
        int extracted = 0;
        using (var input = new FileStream(archivePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var archive = new ZipArchive(input, ZipArchiveMode.Read))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                string relative = (entry.FullName ?? String.Empty).Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
                if (relative.Length == 0) continue;
                if (Path.IsPathRooted(relative)) throw new InvalidDataException("Arşivde kök dışına çıkan bir yol var: " + entry.FullName);
                string target = Path.GetFullPath(Path.Combine(destination, relative));
                if (!target.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Arşivde hedef klasör dışına çıkan bir yol var: " + entry.FullName);
                bool directoryEntry = entry.FullName.EndsWith("/", StringComparison.Ordinal) || entry.FullName.EndsWith("\\", StringComparison.Ordinal);
                if (directoryEntry)
                {
                    if (File.Exists(target)) throw new IOException("Ayıklanacak klasörün yerinde aynı adlı dosya var: " + target);
                    Directory.CreateDirectory(target);
                    continue;
                }
                string parent = Path.GetDirectoryName(target);
                if (!String.IsNullOrWhiteSpace(parent)) Directory.CreateDirectory(parent);
                if (Directory.Exists(target)) throw new IOException("Ayıklanacak dosyanın yerinde aynı adlı klasör var: " + target);
                bool overwrite = false;
                if (File.Exists(target))
                {
                    FileConflictChoice choice;
                    if (applyToAll) choice = applyChoice;
                    else
                    {
                        choice = FileConflictDialog.Ask(this, archivePath + "!" + entry.FullName, target, out applyToAll);
                        if (applyToAll) applyChoice = choice;
                    }
                    if (choice == FileConflictChoice.Skip) continue;
                    if (choice == FileConflictChoice.KeepBoth) target = GetUnusedFilePath(target);
                    else overwrite = true;
                }
                using (Stream source = entry.Open())
                using (var output = new FileStream(target, overwrite ? FileMode.Create : FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    source.CopyTo(output);
                extracted++;
            }
        }
        return extracted;
    }

    private void ConfigureSearchContextMenu()
    {
        ConfigureItemContextMenu(deepSearchResults, false, true);
    }

    private void DeleteSearchResult(Entry entry)
    {
        if (entry == null || !File.Exists(entry.FullPath)) return;
        DialogResult answer = MessageBox.Show(this, "'" + entry.ToString() + "' dosyası silinsin mi?", "Silme onayı",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes) return;
        try
        {
            MovePathsToRecycleBin(new string[] { entry.FullPath });
            foreach (ListViewItem row in deepSearchResults.Items.Cast<ListViewItem>().ToArray())
            {
                Entry item = row.Tag as Entry;
                if (item != null && String.Equals(item.FullPath, entry.FullPath, StringComparison.OrdinalIgnoreCase)) deepSearchResults.Items.Remove(row);
            }
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private static bool HasFileDropClipboard()
    {
        try { return Clipboard.ContainsFileDropList() && Clipboard.GetFileDropList().Count > 0; }
        catch { return false; }
    }

    private static string GetDefaultEdefterTemplateFolder()
    {
        return PortableDataPath("Şablonlar");
    }

    private int LoadThemeIndex()
    {
        try
        {
            int value;
            if (File.Exists(themeFile) && Int32.TryParse(File.ReadAllText(themeFile).Trim(), out value) && value >= 0 && value <= 5)
                return value;
        }
        catch { }
        return 0;
    }

    private void LoadPreferences()
    {
        try
        {
            if (!File.Exists(preferencesFile)) return;
            foreach (string line in File.ReadAllLines(preferencesFile))
            {
                int split = line.IndexOf('=');
                if (split < 0) continue;
                string key = line.Substring(0, split).Trim();
                string value = line.Substring(split + 1).Trim();
                bool enabled = value == "1" || String.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
                if (key == "closeToTray") closeToTrayOnExit = enabled;
                else if (key == "keepOnTop") keepWindowOnTop = enabled;
                else if (key == "rememberFolderSort") rememberFolderSort = enabled;
                else if (key == "startMaximized") startMaximizedOnLaunch = enabled;
                else if (key == "minimizeAfterOpeningFile") minimizeAfterOpeningFile = enabled;
                else if (key == "classicFolderView") classicFolderView = enabled;
                else if (key == "historyRetentionCount")
                {
                    int count;
                    if (Int32.TryParse(value, out count)) historyRetentionCount = Math.Max(1, Math.Min(30, count));
                }
                else if (key == "historyRetentionUnit" && new[] { "Gün", "Hafta", "Ay", "Yıl" }.Contains(value)) historyRetentionUnit = value;
                else if (key == "preserveHistoryRetention") preserveHistoryRetention = enabled;
                else if (key == "edefterTemplateFolder") edefterTemplateFolder = value;
                else if (key == "oneDriveRoot") oneDriveRoot = value;
                else if (key == "oneDriveRootEnabled") oneDriveRootEnabled = enabled;
                else if (key == "activeTab") savedActiveTabName = value;
                else if (key == "searchSubfolders") searchSubfolders = enabled;
                else if (key == "automaticUpdateCheck") automaticUpdateCheck = enabled;
                else if (key == "updateManifestUrl")
                    updateManifestUrl = String.IsNullOrWhiteSpace(value) ? DefaultUpdateManifestUrl : value;
            }
        }
        catch { }
    }

    private void SavePreferences()
    {
        try
        {
            string directory = Path.GetDirectoryName(preferencesFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllLines(preferencesFile, new string[]
            {
                "closeToTray=" + (closeToTrayOnExit ? "1" : "0"),
                "keepOnTop=" + (keepWindowOnTop ? "1" : "0"),
                "rememberFolderSort=" + (rememberFolderSort ? "1" : "0"),
                "startMaximized=" + (startMaximizedOnLaunch ? "1" : "0"),
                "minimizeAfterOpeningFile=" + (minimizeAfterOpeningFile ? "1" : "0"),
                "classicFolderView=" + (classicFolderView ? "1" : "0"),
                "historyRetentionCount=" + historyRetentionCount.ToString(CultureInfo.InvariantCulture),
                "historyRetentionUnit=" + historyRetentionUnit,
                "preserveHistoryRetention=" + (preserveHistoryRetention ? "1" : "0"),
                "edefterTemplateFolder=" + edefterTemplateFolder,
                "oneDriveRoot=" + oneDriveRoot,
                "oneDriveRootEnabled=" + (oneDriveRootEnabled ? "1" : "0"),
                "activeTab=" + (tabs.SelectedTab == null ? savedActiveTabName : tabs.SelectedTab.Text),
                "searchSubfolders=" + (searchSubfolders ? "1" : "0"),
                "automaticUpdateCheck=" + (automaticUpdateCheck ? "1" : "0"),
                "updateManifestUrl=" + updateManifestUrl
            });
        }
        catch { }
    }

    private void InitializeUpdateWorkers()
    {
        updateCheckWorker.DoWork += (s, e) => e.Result = ReadUpdateRelease(Convert.ToString(e.Argument, CultureInfo.InvariantCulture));
        updateCheckWorker.RunWorkerCompleted += (s, e) =>
        {
            bool manual = manualUpdateCheck;
            manualUpdateCheck = false;
            if (e.Error != null)
            {
                if (manual) MessageBox.Show(this, "Güncelleme kontrol edilemedi. Adresi ve internet bağlantısını kontrol edin.\r\n\r\n" + e.Error.Message,
                    "Güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            UpdateRelease release = e.Result as UpdateRelease;
            if (release == null)
            {
                if (manual) MessageBox.Show(this, "Program güncel (" + CurrentAppVersion + ").", "Güncelleme kontrolü", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            pendingUpdateRelease = release;
            string message = "Yeni sürüm " + release.Version + " bulundu.\r\nKurulu sürüm: " + CurrentAppVersion +
                (String.IsNullOrWhiteSpace(release.Notes) ? String.Empty : "\r\n\r\nDeğişiklikler:\r\n" + release.Notes) +
                "\r\n\r\nŞimdi indirip kurmak ister misiniz?";
            if (MessageBox.Show(this, message, "Program güncellemesi", MessageBoxButtons.YesNo,
                MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            string updater = Path.Combine(Application.StartupPath, UpdaterExecutableName);
            if (!File.Exists(updater))
            {
                MessageBox.Show(this, "Güncelleyici bulunamadı:\r\n" + updater + "\r\n\r\nİlk kurulumda " + UpdaterExecutableName + " dosyasını programla aynı klasöre koyun.",
                    "Güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!updateInstallWorker.IsBusy) updateInstallWorker.RunWorkerAsync(release);
        };
        updateInstallWorker.DoWork += (s, e) => e.Result = DownloadAndVerifyUpdate((UpdateRelease)e.Argument);
        updateInstallWorker.RunWorkerCompleted += (s, e) =>
        {
            if (e.Error != null)
            {
                MessageBox.Show(this, "Güncelleme indirilemedi veya doğrulanamadı. Program değiştirilmedi.\r\n\r\n" + e.Error.Message,
                    "Güncelleme başarısız", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string stagedPath = Convert.ToString(e.Result, CultureInfo.InvariantCulture);
            string updater = Path.Combine(Application.StartupPath, UpdaterExecutableName);
            string payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture) + "|" +
                Application.ExecutablePath + "|" + stagedPath));
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = updater,
                    Arguments = "\"" + payload + "\"",
                    WorkingDirectory = Application.StartupPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
                exitRequested = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Güncelleme dosyası hazırlandı ancak güncelleyici başlatılamadı. Program açık kalıyor.\r\n\r\n" + ex.Message,
                    "Güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        };
    }

    private void CheckForUpdates(bool manual)
    {
        if (updateCheckWorker.IsBusy)
        {
            manualUpdateCheck = manualUpdateCheck || manual;
            return;
        }
        string url = (updateManifestUrl ?? String.Empty).Trim();
        if (url.Length == 0)
        {
            if (manual) MessageBox.Show(this, "Ayarlar > İnternet üzerinden program güncellemesi bölümüne latest.xml adresini girin.",
                "Güncelleme adresi eksik", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        manualUpdateCheck = manual;
        status.Text = "Program güncellemesi kontrol ediliyor...";
        updateCheckWorker.RunWorkerAsync(url);
    }

    private static UpdateRelease ReadUpdateRelease(string manifestUrl)
    {
        Uri uri;
        if (!Uri.TryCreate(manifestUrl, UriKind.Absolute, out uri) || uri.Scheme != Uri.UriSchemeHttps)
            throw new InvalidOperationException("Güncelleme adresi HTTPS olmalıdır.");
        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
        var request = (HttpWebRequest)WebRequest.Create(uri);
        request.Method = "GET";
        request.UserAgent = "DosyaGezgini-Updater/1.0";
        request.Timeout = 20000;
        request.ReadWriteTimeout = 20000;
        using (WebResponse response = request.GetResponse())
        using (Stream stream = response.GetResponseStream())
        using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
        {
            var document = new XmlDocument { XmlResolver = null };
            document.Load(reader);
            XmlElement root = document.DocumentElement;
            if (root == null || root.Name != "update") throw new InvalidDataException("latest.xml biçimi tanınmadı.");
            string versionText = root.GetAttribute("version");
            Version latest, current;
            if (!Version.TryParse(versionText, out latest) || !Version.TryParse(CurrentAppVersion, out current))
                throw new InvalidDataException("Sürüm numarası geçersiz.");
            if (latest <= current) return null;
            string downloadUrl = GetXmlChildText(root, "downloadUrl");
            Uri downloadUri;
            if (!Uri.TryCreate(downloadUrl, UriKind.Absolute, out downloadUri) || downloadUri.Scheme != Uri.UriSchemeHttps)
                throw new InvalidDataException("İndirme adresi HTTPS olmalıdır.");
            string hash = GetXmlChildText(root, "sha256").Trim();
            if (hash.Length != 64 || hash.Any(ch => !Uri.IsHexDigit(ch))) throw new InvalidDataException("SHA-256 özeti geçersiz.");
            string updaterUrl = GetXmlChildText(root, "updaterUrl");
            string updaterHash = GetXmlChildText(root, "updaterSha256").Trim();
            Uri updaterUri = null;
            if (updaterUrl.Length > 0 || updaterHash.Length > 0)
            {
                if (!Uri.TryCreate(updaterUrl, UriKind.Absolute, out updaterUri) || updaterUri.Scheme != Uri.UriSchemeHttps)
                    throw new InvalidDataException("Güncelleme yardımcısı indirme adresi HTTPS olmalıdır.");
                if (updaterHash.Length != 64 || updaterHash.Any(ch => !Uri.IsHexDigit(ch)))
                    throw new InvalidDataException("Güncelleme yardımcısının SHA-256 özeti geçersiz.");
            }
            return new UpdateRelease(versionText, downloadUri, hash, GetXmlChildText(root, "notes"), updaterUri, updaterHash);
        }
    }

    private static string GetXmlChildText(XmlElement root, string name)
    {
        XmlNode node = root.SelectSingleNode(name);
        return node == null ? String.Empty : node.InnerText.Trim();
    }

    private static string DownloadAndVerifyUpdate(UpdateRelease release)
    {
        if (release == null) throw new InvalidOperationException("Güncelleme bilgisi bulunamadı.");
        string directory = Path.Combine(Path.GetTempPath(), "DosyaGezginiUpdate-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string destination = Path.Combine(directory, "DosyaGezgini.new.exe");
        DownloadAndVerifyFile(release.DownloadUri, release.Sha256, destination);

        string updaterPath = Path.Combine(Application.StartupPath, UpdaterExecutableName);
        if (!File.Exists(updaterPath))
        {
            if (release.UpdaterUri == null || String.IsNullOrWhiteSpace(release.UpdaterSha256))
                throw new InvalidDataException("Güncelleme yardımcısı bulunamadı ve latest.xml içinde indirme bilgisi yok.");
            string stagedUpdater = Path.Combine(directory, UpdaterExecutableName);
            DownloadAndVerifyFile(release.UpdaterUri, release.UpdaterSha256, stagedUpdater);
            File.Copy(stagedUpdater, updaterPath, false);
        }
        return destination;
    }

    private static void DownloadAndVerifyFile(Uri downloadUri, string expectedHash, string destination)
    {
        ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
        var request = (HttpWebRequest)WebRequest.Create(downloadUri);
        request.Method = "GET";
        request.UserAgent = "DosyaGezgini-Updater/1.0";
        request.Timeout = 60000;
        request.ReadWriteTimeout = 60000;
        using (WebResponse response = request.GetResponse())
        using (Stream input = response.GetResponseStream())
        using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            byte[] buffer = new byte[81920];
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0) output.Write(buffer, 0, read);
        }
        string actualHash;
        using (var sha = SHA256.Create())
        using (var input = File.OpenRead(destination))
            actualHash = BitConverter.ToString(sha.ComputeHash(input)).Replace("-", String.Empty);
        if (!String.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase))
        {
            try { File.Delete(destination); } catch { }
            throw new InvalidDataException("İndirilen dosyanın SHA-256 özeti eşleşmedi.");
        }
    }

    private string StorePortablePath(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return path ?? String.Empty;
        string full = path;
        try { full = Path.GetFullPath(path); } catch { }
        if (!oneDriveRootEnabled || String.IsNullOrWhiteSpace(oneDriveRoot)) return full;
        string root = oneDriveRoot.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (String.Equals(full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), root,
            StringComparison.OrdinalIgnoreCase)) return "OneDrive:\\";
        string prefix = root + Path.DirectorySeparatorChar;
        if (full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            return "OneDrive:\\" + full.Substring(prefix.Length).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        return full;
    }

    private string ResolvePortablePath(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return path ?? String.Empty;
        if (!path.StartsWith("OneDrive:\\", StringComparison.OrdinalIgnoreCase) &&
            !String.Equals(path, "OneDrive:", StringComparison.OrdinalIgnoreCase)) return path;
        if (String.IsNullOrWhiteSpace(oneDriveRoot) || !Directory.Exists(oneDriveRoot)) return String.Empty;
        string relative = path.Length <= 9 ? String.Empty : path.Substring(9);
        relative = relative.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, '\\', '/');
        string resolved = Path.GetFullPath(Path.Combine(oneDriveRoot, relative));
        if (!IsPathWithin(resolved, oneDriveRoot) && !String.Equals(resolved.TrimEnd(Path.DirectorySeparatorChar),
            oneDriveRoot.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase)) return String.Empty;
        return resolved;
    }

    private void PromptForPortableOneDriveRootIfNeeded()
    {
        if (Directory.Exists(oneDriveRoot)) return;
        string[] portableFiles = new string[]
        {
            settingsFile, navigationHistoryFile, folderTabsFile, shortcutsFile, historyFile,
            workProfileLocationsFile, workProfilesFile, sharedDevelopmentFile, transferPanePathsFile
        };
        bool hasPortablePaths = false;
        foreach (string file in portableFiles)
        {
            try
            {
                if (!File.Exists(file)) continue;
                using (var reader = new StreamReader(file))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                        if (line.IndexOf("OneDrive:\\", StringComparison.OrdinalIgnoreCase) >= 0)
                        { hasPortablePaths = true; break; }
                }
                if (hasPortablePaths) break;
            }
            catch { }
        }
        if (!hasPortablePaths) return;
        using (var dialog = new FolderBrowserDialog())
        {
            dialog.Description = "Taşınabilir kayıtları bu bilgisayarda açmak için OneDrive kök klasörünüzü seçin.";
            if (MessageBox.Show(this,
                "Taşınabilir OneDrive yolları bulundu. Bu bilgisayardaki OneDrive ana klasörünü seçerek Gezgin, Kısayollarım, Geçmiş, Çiftli İşlem ve İş Klasörleri kayıtlarını açabilirsiniz.",
                "OneDrive klasörünü seçin", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            oneDriveRoot = dialog.SelectedPath;
            oneDriveRootEnabled = true;
            SavePreferences();
        }
    }

    private void LoadFolderSortPreferences()
    {
        try
        {
            if (!File.Exists(folderSortFile)) return;
            foreach (string line in File.ReadAllLines(folderSortFile))
            {
                string[] parts = line.Split('|');
                if (parts.Length != 3) continue;
                string path = Encoding.UTF8.GetString(Convert.FromBase64String(parts[0]));
                int column;
                if (!Int32.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out column) || column < -1 || column > 3) continue;
                folderSortPreferences[path] = new FileSortPreference(column, parts[2] == "1");
            }
        }
        catch { }
    }

    private void LoadListSortPreferences()
    {
        try
        {
            if (File.Exists(listSortFile))
            {
                foreach (string line in File.ReadAllLines(listSortFile))
                {
                    string[] parts = line.Split('|');
                    int column;
                    if (parts.Length != 3 || !Int32.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out column)) continue;
                    if (column < -1 || column > 3) continue;
                    listSortPreferences[parts[0]] = new FileSortPreference(column, parts[2] == "1");
                }
            }
            FileSortPreference preference;
            if (listSortPreferences.TryGetValue("Gezgin", out preference)) { fileSortColumn = preference.Column; fileSortAscending = preference.Ascending; }
            if (listSortPreferences.TryGetValue("Kısayollarım", out preference)) { shortcutSortColumn = preference.Column; shortcutSortAscending = preference.Ascending; }
            if (listSortPreferences.TryGetValue("Geçmiş", out preference)) { historySortColumn = preference.Column; historySortAscending = preference.Ascending; }
            if (listSortPreferences.TryGetValue("Arama", out preference)) { searchSortColumn = preference.Column; searchSortAscending = preference.Ascending; }
        }
        catch { }
    }

    private void SaveListSortPreference(string key, int column, bool ascending)
    {
        listSortPreferences[key] = new FileSortPreference(column, ascending);
        try
        {
            string directory = Path.GetDirectoryName(listSortFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllLines(listSortFile, listSortPreferences.Select(pair =>
                pair.Key + "|" + pair.Value.Column.ToString(CultureInfo.InvariantCulture) + "|" + (pair.Value.Ascending ? "1" : "0")).ToArray());
        }
        catch { }
    }

    private void LoadColumnLayouts()
    {
        try
        {
            if (!File.Exists(columnLayoutsFile)) return;
            foreach (string line in File.ReadAllLines(columnLayoutsFile))
            {
                int separator = line.IndexOf('|');
                if (separator <= 0) continue;
                string key = line.Substring(0, separator);
                string[] values = line.Substring(separator + 1).Split(';');
                var layout = new ColumnLayout(values.Length);
                bool valid = true;
                for (int i = 0; i < values.Length; i++)
                {
                    string[] pair = values[i].Split(',');
                    if (pair.Length != 2 ||
                        !Int32.TryParse(pair[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out layout.DisplayIndexes[i]) ||
                        !Int32.TryParse(pair[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out layout.Widths[i]))
                    {
                        valid = false;
                        break;
                    }
                }
                if (valid) columnLayouts[key] = layout;
            }
        }
        catch { }
    }

    private void LoadLinkedPaneWidths()
    {
        try
        {
            if (!File.Exists(linkedPaneWidthsFile)) return;
            foreach (string line in File.ReadAllLines(linkedPaneWidthsFile))
            {
                int separator = line.IndexOf('|');
                int width;
                if (separator <= 0 || !Int32.TryParse(line.Substring(separator + 1), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out width)) continue;
                string key = line.Substring(0, separator);
                if (!paneWidthPreferences.ContainsKey(key))
                    paneWidthPreferences[key] = new PaneWidthPreference(width, LoadSplitterDistance());
            }
        }
        catch { }
    }

    private void LoadTabPaneWidths()
    {
        try
        {
            if (!File.Exists(tabPaneWidthsFile)) return;
            foreach (string line in File.ReadAllLines(tabPaneWidthsFile))
            {
                string[] parts = line.Split('|');
                int primary, explorerFolder;
                if (parts.Length != 3 ||
                    !Int32.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out primary) ||
                    !Int32.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out explorerFolder)) continue;
                paneWidthPreferences[parts[0]] = new PaneWidthPreference(primary, explorerFolder);
            }
        }
        catch { }
    }

    private void SaveTabPaneWidths()
    {
        try
        {
            string directory = Path.GetDirectoryName(tabPaneWidthsFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllLines(tabPaneWidthsFile, paneWidthPreferences.Select(pair =>
                pair.Key + "|" + pair.Value.PrimaryWidth.ToString(CultureInfo.InvariantCulture) + "|" +
                pair.Value.ExplorerFolderWidth.ToString(CultureInfo.InvariantCulture)).ToArray());
        }
        catch { }
    }

    private void SaveActivePaneWidths()
    {
        if (!panePersistenceReady || applyingPaneWidths || folderFilesSplit == null || String.IsNullOrEmpty(activePaneKey) ||
            activePaneKey == "Ayarlar" || activePaneKey == "İş Klasörleri") return;
        int primaryWidth;
        int explorerFolderWidth = -1;
        if (String.Equals(activePaneKey, "Gezgin", StringComparison.Ordinal))
        {
            if (folderFilesSplit.ClientSize.Width <= folderFilesSplit.Panel1MinSize + folderFilesSplit.Panel2MinSize + folderFilesSplit.SplitterWidth) return;
            primaryWidth = folderFilesSplit.SplitterDistance;
        }
        else
        {
            SplitContainer linked;
            if (!linkedPaneSplits.TryGetValue(activePaneKey, out linked) ||
                linked.ClientSize.Width <= linked.Panel1MinSize + linked.Panel2MinSize + linked.SplitterWidth) return;
            primaryWidth = linked.SplitterDistance;
            if (folderFilesSplit.ClientSize.Width > folderFilesSplit.Panel1MinSize + folderFilesSplit.Panel2MinSize + folderFilesSplit.SplitterWidth)
                explorerFolderWidth = folderFilesSplit.SplitterDistance;
        }
        paneWidthPreferences[activePaneKey] = new PaneWidthPreference(primaryWidth, explorerFolderWidth);
        SaveTabPaneWidths();
    }

    private void ScheduleActivePaneWidthRestore(string key)
    {
        try
        {
            BeginInvoke((MethodInvoker)(() => BeginInvoke((MethodInvoker)(() =>
            {
                if (!String.Equals(activePaneKey, key, StringComparison.Ordinal)) return;
                PaneWidthPreference preference;
                bool hasPreference = paneWidthPreferences.TryGetValue(key, out preference);
                applyingPaneWidths = true;
                try
                {
                    if (String.Equals(key, "Gezgin", StringComparison.Ordinal))
                    {
                        int width = hasPreference ? preference.PrimaryWidth : LoadSplitterDistance();
                        RestoreSplitterDistanceWhenReady(folderFilesSplit, width);
                    }
                    else
                    {
                        SplitContainer linked;
                        if (linkedPaneSplits.TryGetValue(key, out linked) && hasPreference)
                            RestoreSplitterDistanceWhenReady(linked, preference.PrimaryWidth);
                        int explorerWidth = hasPreference && preference.ExplorerFolderWidth > 0
                            ? preference.ExplorerFolderWidth : LoadSplitterDistance();
                        RestoreSplitterDistanceWhenReady(folderFilesSplit, explorerWidth);
                    }
                }
                finally { applyingPaneWidths = false; }
            }))));
        }
        catch { }
    }

    private static void RestoreSplitterDistanceWhenReady(SplitContainer split, int distance)
    {
        if (split == null || split.IsDisposed) return;
        int minimum = split.Panel1MinSize + split.Panel2MinSize + split.SplitterWidth;
        if (split.ClientSize.Width <= minimum) return;
        RestoreSplitterDistance(split, distance);
    }

    private void AttachColumnLayout(ListView list, string key)
    {
        list.AllowColumnReorder = true;
        columnLayoutLists[key] = list;
        ColumnLayout saved;
        if (columnLayouts.TryGetValue(key, out saved) && saved.DisplayIndexes.Length == list.Columns.Count)
        {
            ApplyColumnLayout(list, saved);
            if (list == deepSearchResults) deepSearchColumnsCustomized = true;
        }
        list.ColumnWidthChanged += (s, e) =>
        {
            if (restoringColumnLayout) return;
            if (list == deepSearchResults) deepSearchColumnsCustomized = true;
            columnLayouts[key] = CaptureColumnLayout(list);
            SaveColumnLayouts();
        };
        list.ColumnReordered += (s, e) =>
        {
            if (restoringColumnLayout) return;
            if (list == deepSearchResults) deepSearchColumnsCustomized = true;
            try
            {
                BeginInvoke((MethodInvoker)(() =>
                {
                    if (IsDisposed || restoringColumnLayout) return;
                    columnLayouts[key] = CaptureColumnLayout(list);
                    SaveColumnLayouts();
                }));
            }
            catch { }
        };
    }

    private void RestoreAllColumnLayouts()
    {
        foreach (KeyValuePair<string, ListView> entry in columnLayoutLists)
        {
            ColumnLayout layout;
            if (columnLayouts.TryGetValue(entry.Key, out layout) &&
                layout.DisplayIndexes.Length == entry.Value.Columns.Count)
                ApplyColumnLayout(entry.Value, layout);
        }
        if (columnLayouts.ContainsKey("Arama")) deepSearchColumnsCustomized = true;
    }

    private ColumnLayout CaptureColumnLayout(ListView list)
    {
        var layout = new ColumnLayout(list.Columns.Count);
        for (int i = 0; i < list.Columns.Count; i++)
        {
            layout.DisplayIndexes[i] = list.Columns[i].DisplayIndex;
            layout.Widths[i] = list.Columns[i].Width;
        }
        return layout;
    }

    private void ApplyColumnLayout(ListView list, ColumnLayout layout)
    {
        restoringColumnLayout = true;
        try
        {
            for (int i = 0; i < list.Columns.Count; i++)
                list.Columns[i].Width = Math.Max(24, layout.Widths[i]);
            int[] order = Enumerable.Range(0, list.Columns.Count)
                .OrderBy(i => layout.DisplayIndexes[i]).ToArray();
            for (int display = 0; display < order.Length; display++)
                list.Columns[order[display]].DisplayIndex = display;
        }
        catch { }
        finally { restoringColumnLayout = false; }
    }

    private void SaveColumnLayouts()
    {
        try
        {
            foreach (KeyValuePair<string, ListView> entry in columnLayoutLists)
                if (!entry.Value.IsDisposed && entry.Value.Columns.Count > 0)
                    columnLayouts[entry.Key] = CaptureColumnLayout(entry.Value);
            string directory = Path.GetDirectoryName(columnLayoutsFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllLines(columnLayoutsFile, columnLayouts.Select(entry =>
            {
                string[] values = new string[entry.Value.DisplayIndexes.Length];
                for (int i = 0; i < values.Length; i++)
                    values[i] = entry.Value.DisplayIndexes[i].ToString(CultureInfo.InvariantCulture) + "," +
                        entry.Value.Widths[i].ToString(CultureInfo.InvariantCulture);
                return entry.Key + "|" + String.Join(";", values);
            }).ToArray());
        }
        catch { }
    }

    private void SaveFolderSortPreferences()
    {
        try
        {
            string directory = Path.GetDirectoryName(folderSortFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllLines(folderSortFile, folderSortPreferences.Select(pair =>
                Convert.ToBase64String(Encoding.UTF8.GetBytes(pair.Key)) + "|" +
                pair.Value.Column.ToString(CultureInfo.InvariantCulture) + "|" + (pair.Value.Ascending ? "1" : "0")).ToArray());
        }
        catch { }
    }

    private void RememberCurrentFolderSort()
    {
        if (String.IsNullOrEmpty(currentPath)) return;
        folderSortPreferences[currentPath] = new FileSortPreference(fileSortColumn, fileSortAscending);
        SaveFolderSortPreferences();
    }

    private void ApplyFolderSortPreference(string path)
    {
        FileSortPreference preference;
        if (rememberFolderSort && folderSortPreferences.TryGetValue(path, out preference))
        {
            fileSortColumn = preference.Column;
            fileSortAscending = preference.Ascending;
        }
        else if (listSortPreferences.TryGetValue("Gezgin", out preference))
        {
            fileSortColumn = preference.Column;
            fileSortAscending = preference.Ascending;
        }
        else
        {
            fileSortColumn = -1;
            fileSortAscending = true;
        }
        SetSortIndicators(files, "Ad", "Değiştirme tarihi", fileSortColumn, fileSortAscending);
    }

    private void SaveThemeIndex(int value)
    {
        try
        {
            string directory = Path.GetDirectoryName(themeFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(themeFile, value.ToString(CultureInfo.InvariantCulture));
        }
        catch { }
    }

    private void ApplyTheme()
    {
        ThemeColors colors = ThemeColors.ForIndex(themeSelector.SelectedIndex);
        BackColor = colors.Base;
        ForeColor = colors.Text;
        ApplyThemeToControls(this, colors);
        status.BackColor = colors.Surface;
        status.ForeColor = colors.Muted;
        foreach (TabPage page in tabs.TabPages) page.BackColor = colors.Base;
    }

    private void ApplyThemeToControls(Control parent, ThemeColors colors)
    {
        foreach (Control control in parent.Controls)
        {
            if (control is SplitContainer)
            {
                var split = (SplitContainer)control;
                split.BackColor = colors.Split;
                split.Panel1.BackColor = colors.Base;
                split.Panel2.BackColor = colors.Base;
            }
            else if (control is TabControl)
            {
                control.BackColor = colors.Base;
                control.ForeColor = colors.Text;
            }
            else if (control is TabPage)
            {
                control.BackColor = colors.Base;
                control.ForeColor = colors.Text;
            }
            else if (control is Button)
            {
                var button = (Button)control;
                button.UseVisualStyleBackColor = false;
                button.BackColor = colors.Button;
                button.ForeColor = colors.Text;
                button.FlatAppearance.BorderColor = colors.Border;
            }
            else if (control is CheckBox || control is RadioButton || control is GroupBox)
            {
                control.BackColor = colors.Base;
                control.ForeColor = colors.Text;
            }
            else if (control is ListView || control is ListBox || control is TreeView)
            {
                control.BackColor = colors.Surface;
                control.ForeColor = colors.Text;
            }
            else if (control is TextBoxBase || control is ComboBox)
            {
                control.BackColor = colors.Surface;
                control.ForeColor = colors.Text;
            }
            else if (control is Label)
            {
                control.ForeColor = control == status ? colors.Muted : colors.Text;
                if (control == status) control.BackColor = colors.Surface;
                if (control is LinkLabel)
                {
                    var link = (LinkLabel)control;
                    link.LinkColor = colors.Text;
                    link.ActiveLinkColor = colors.Muted;
                    link.VisitedLinkColor = colors.Text;
                }
            }
            else if (control is Panel)
            {
                control.BackColor = control is SplitterPanel ? colors.Base : colors.Surface;
            }

            ApplyThemeToControls(control, colors);
        }
    }

    private sealed class ThemeColors
    {
        public readonly Color Base, Surface, Split, Button, Border, Text, Muted;

        private ThemeColors(Color baseColor, Color surface, Color split, Color button,
            Color border, Color text, Color muted)
        {
            Base = baseColor; Surface = surface; Split = split; Button = button;
            Border = border; Text = text; Muted = muted;
        }

        public static ThemeColors ForIndex(int index)
        {
            if (index == 1)
                return new ThemeColors(Color.FromArgb(239, 247, 241), Color.FromArgb(250, 253, 250),
                    Color.FromArgb(204, 224, 210), Color.FromArgb(224, 242, 229), Color.FromArgb(157, 190, 168),
                    Color.FromArgb(34, 66, 45), Color.FromArgb(77, 110, 87));
            if (index == 2)
                return new ThemeColors(Color.FromArgb(25, 28, 34), Color.FromArgb(35, 39, 47),
                    Color.FromArgb(14, 17, 22), Color.FromArgb(49, 56, 67), Color.FromArgb(73, 82, 96),
                    Color.FromArgb(232, 236, 242), Color.FromArgb(171, 181, 194));
            if (index == 3)
                return new ThemeColors(Color.FromArgb(35, 22, 52), Color.FromArgb(49, 32, 72),
                    Color.FromArgb(24, 14, 38), Color.FromArgb(83, 52, 119), Color.FromArgb(132, 91, 174),
                    Color.FromArgb(249, 241, 255), Color.FromArgb(202, 177, 225));
            if (index == 4)
                return new ThemeColors(Color.FromArgb(255, 239, 224), Color.FromArgb(255, 249, 242),
                    Color.FromArgb(239, 199, 170), Color.FromArgb(255, 218, 193), Color.FromArgb(211, 133, 104),
                    Color.FromArgb(83, 38, 28), Color.FromArgb(138, 83, 65));
            if (index == 5)
                return new ThemeColors(Color.FromArgb(224, 248, 246), Color.FromArgb(242, 255, 253),
                    Color.FromArgb(174, 221, 215), Color.FromArgb(198, 244, 236), Color.FromArgb(84, 171, 157),
                    Color.FromArgb(18, 66, 62), Color.FromArgb(57, 119, 111));
            return new ThemeColors(Color.FromArgb(246, 248, 251), Color.White,
                Color.FromArgb(226, 231, 238), Color.FromArgb(239, 245, 255), Color.FromArgb(202, 212, 225),
                Color.FromArgb(38, 51, 70), Color.FromArgb(82, 94, 112));
        }
    }

    private sealed class NaturalNameComparer : IComparer<string>
    {
        public static readonly NaturalNameComparer Instance = new NaturalNameComparer();

        public int Compare(string first, string second)
        {
            if (ReferenceEquals(first, second)) return 0;
            if (first == null) return -1;
            if (second == null) return 1;
            return StrCmpLogicalW(first, second);
        }
    }

    private static void StyleList(ListBox list)
    {
        list.BorderStyle = BorderStyle.None;
        list.BackColor = Color.White;
        list.ForeColor = Color.FromArgb(38, 51, 70);
        list.Font = new Font("Segoe UI", 10F);
        list.IntegralHeight = false;
        list.ItemHeight = 28;
    }

    private static void StyleDetailsList(ListView list, string nameColumn, string dateColumn)
    {
        list.View = View.Details;
        list.FullRowSelect = true;
        list.GridLines = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        list.BorderStyle = BorderStyle.None;
        list.BackColor = Color.White;
        list.ForeColor = Color.FromArgb(38, 51, 70);
        list.Font = new Font("Segoe UI", 10F);
        list.Columns.Add(nameColumn, 450, HorizontalAlignment.Left);
        list.Columns.Add(dateColumn, 190, HorizontalAlignment.Left);
        list.Columns.Add("Boyut", 120, HorizontalAlignment.Right);
    }

    private void InitializeFilter(ComboBox combo, Action changed)
    {
        combo.Dock = DockStyle.Bottom;
        combo.Height = 30;
        combo.DropDownStyle = ComboBoxStyle.DropDownList;
        combo.Font = new Font("Segoe UI", 9F);
        combo.Items.AddRange(new object[]
        {
            "Word dosyaları", "Excel dosyaları", "Acrobat dosyaları",
            "UDF dosyaları", "ZIP dosyaları", "Tüm dosyalar"
        });
        combo.SelectedIndexChanged += (s, e) => changed();
        combo.SelectedItem = "Tüm dosyalar";
    }

    private void InitializeDeepSearchPage(TabPage page)
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 74, BackColor = Color.White, Padding = new Padding(10, 7, 10, 5) };
        deepSearchBox.SetBounds(10, 7, 640, 30);
        deepSearchBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        deepSearchBox.Font = new Font("Segoe UI", 10F);
        deepSearchBox.BorderStyle = BorderStyle.FixedSingle;
        deepSearchBox.TabStop = true;
        deepSearchBox.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Tab) e.IsInputKey = true; };
        deepSearchBox.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Tab)
            {
                CycleSearchFocus(deepSearchBox, !e.Shift);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                StartDeepSearch();
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down && deepSearchResults.Items.Count > 0)
            {
                MoveListViewSelection(deepSearchResults, 1);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        };
        deepSearchButton.Text = "Ara";
        deepSearchButton.SetBounds(662, 7, 76, 30);
        deepSearchButton.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        deepSearchButton.FlatStyle = FlatStyle.Flat;
        deepSearchButton.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        deepSearchButton.BackColor = Color.FromArgb(239, 245, 255);
        deepSearchButton.Click += (s, e) => StartDeepSearch();
        clearDeepSearchButton.Text = "Temizle";
        clearDeepSearchButton.SetBounds(494, 7, 76, 30);
        clearDeepSearchButton.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        clearDeepSearchButton.FlatStyle = FlatStyle.Flat;
        clearDeepSearchButton.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        clearDeepSearchButton.BackColor = Color.FromArgb(245, 246, 248);
        clearDeepSearchButton.Click += (s, e) => ClearDeepSearchResults();
        cancelDeepSearchButton.Text = "Durdur";
        cancelDeepSearchButton.SetBounds(578, 7, 76, 30);
        cancelDeepSearchButton.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        cancelDeepSearchButton.Enabled = false;
        cancelDeepSearchButton.FlatStyle = FlatStyle.Flat;
        cancelDeepSearchButton.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        cancelDeepSearchButton.BackColor = Color.FromArgb(250, 239, 239);
        cancelDeepSearchButton.Click += (s, e) => CancelDeepSearch();
        searchProgress.SetBounds(10, 44, 728, 16);
        searchProgress.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        searchProgress.Style = ProgressBarStyle.Marquee;
        searchProgress.MarqueeAnimationSpeed = 28;
        searchProgress.Visible = false;
        toolbar.Controls.Add(deepSearchBox);
        toolbar.Controls.Add(deepSearchButton);
        toolbar.Controls.Add(clearDeepSearchButton);
        toolbar.Controls.Add(cancelDeepSearchButton);
        toolbar.Controls.Add(searchProgress);
        toolbar.Resize += (s, e) =>
        {
            deepSearchButton.Left = Math.Max(10, toolbar.ClientSize.Width - deepSearchButton.Width - 10);
            clearDeepSearchButton.Left = deepSearchButton.Left - clearDeepSearchButton.Width - 8;
            cancelDeepSearchButton.Left = clearDeepSearchButton.Left - cancelDeepSearchButton.Width - 8;
            deepSearchBox.Width = Math.Max(80, cancelDeepSearchButton.Left - deepSearchBox.Left - 12);
            searchProgress.Width = Math.Max(0, toolbar.ClientSize.Width - 20);
        };

        deepSearchResults.Dock = DockStyle.Fill;
        deepSearchResults.View = View.Details;
        deepSearchResults.FullRowSelect = true;
        deepSearchResults.GridLines = true;
        deepSearchResults.HideSelection = false;
        deepSearchResults.MultiSelect = true;
        deepSearchResults.BorderStyle = BorderStyle.None;
        deepSearchResults.BackColor = Color.White;
        deepSearchResults.ForeColor = Color.FromArgb(38, 51, 70);
        deepSearchResults.Font = new Font("Segoe UI", 10F);
        deepSearchResults.Columns.Add("Dosya adı", 560, HorizontalAlignment.Left);
        deepSearchResults.Columns.Add("Klasör", 700, HorizontalAlignment.Left);
        deepSearchResults.Columns.Add("Boyut", 120, HorizontalAlignment.Right);
        deepSearchBoldFont = new Font(deepSearchResults.Font, FontStyle.Bold);
        deepSearchResults.OwnerDraw = true;
        deepSearchResults.DrawColumnHeader += (s, e) => e.DrawDefault = true;
        deepSearchResults.DrawItem += (s, e) => e.DrawDefault = false;
        deepSearchResults.DrawSubItem += DrawSearchResultBold;
        deepSearchResults.Resize += (s, e) => ResizeDeepSearchColumns();
        deepSearchResults.ColumnClick += (s, e) => SortDeepSearchResults(e.Column);
        deepSearchResults.MouseClick += (s, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            ListViewItem row = deepSearchResults.HitTest(e.Location).Item;
            Entry entry = row == null ? null : row.Tag as Entry;
            if (entry != null) RevealInExplorer(entry.FullPath, true);
        };
        deepSearchResults.MouseDoubleClick += (s, e) => OpenListItemInExplorer(deepSearchResults, e.Location);
        ConfigureSearchContextMenu();
        ResizeDeepSearchColumns();
        ApplyDeepSearchSort();
        page.Controls.Add(deepSearchResults);
        page.Controls.Add(toolbar);

        var tip = new ToolTip();
        tip.SetToolTip(deepSearchBox, "* sıfır veya daha çok karakter; ? tek karakter. Joker yoksa adın içinde arar. Birden çok kalıbı ; ile ayırın.");
        tip.SetToolTip(deepSearchButton, "Seçili klasörde ara");
        tip.SetToolTip(clearDeepSearchButton, "Arama sonuçlarını ve bekleyen aramayı temizle");
        tip.SetToolTip(cancelDeepSearchButton, "Çalışan aramayı ve sıradaki aramayı durdur");
        searchWorker.WorkerSupportsCancellation = true;
        searchWorker.WorkerReportsProgress = true;
        searchWorker.DoWork += SearchWorkerDoWork;
        searchWorker.ProgressChanged += (s, e) =>
        {
            if (e.UserState is string) status.Text = (string)e.UserState;
        };
        searchWorker.RunWorkerCompleted += SearchWorkerCompleted;
    }

    private void InitializeMediaPage(TabPage page)
    {
        page.BackColor = Color.FromArgb(244, 247, 251);
        var mediaTabs = new TabControl { Dock = DockStyle.Fill };
        var newsPage = new TabPage("Haber başlıkları") { BackColor = Color.White };
        var worldNewsPage = new TabPage("Yabancı haberler") { BackColor = Color.White };
        var sitesPage = new TabPage("Siteler") { BackColor = Color.FromArgb(244, 247, 251) };
        var researchPage = new TabPage("Araştırma ve güncelleme") { BackColor = Color.FromArgb(244, 247, 251) };
        mediaTabs.TabPages.Add(newsPage);
        mediaTabs.TabPages.Add(worldNewsPage);
        mediaTabs.TabPages.Add(sitesPage);
        mediaTabs.TabPages.Add(researchPage);
        InitializeWorldNewsPage(worldNewsPage);

        var newsToolbar = new Panel { Dock = DockStyle.Top, Height = 46, BackColor = Color.White, Padding = new Padding(10, 7, 10, 5) };
        mediaNewsRefreshButton.Text = "Haberleri yenile";
        mediaNewsRefreshButton.Size = new Size(120, 28);
        mediaNewsRefreshButton.Dock = DockStyle.Right;
        mediaNewsRefreshButton.FlatStyle = FlatStyle.Flat;
        mediaNewsRefreshButton.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        mediaNewsRefreshButton.BackColor = Color.FromArgb(239, 245, 255);
        mediaNewsRefreshButton.Click += (s, e) => StartMediaNewsRefresh(true);
        mediaNewsCategoryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        mediaNewsCategoryCombo.Items.AddRange(new object[] { "Tüm haberler", "Manşet", "Son dakika", "Güncel", "Gündem", "Türkiye", "Dünyadan", "Ekonomi", "Spor", "Sağlık", "Eğitim", "Teknoloji", "Kültür-Sanat", "Yaşam", "İş / Yaşam" });
        mediaNewsCategoryCombo.SelectedIndex = 0;
        mediaNewsCategoryCombo.Width = 145;
        mediaNewsCategoryCombo.Dock = DockStyle.Right;
        mediaNewsCategoryCombo.SelectedIndexChanged += (s, e) => ApplyMediaNewsFilter();
        var domesticCategoryLabel = new Label { Text = "Kategori:", Width = 57, Dock = DockStyle.Right, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 8.5F) };
        mediaNewsStatus.Text = "Resmî RSS akışlarından haber başlıkları alınacak.";
        mediaNewsStatus.Dock = DockStyle.Fill;
        mediaNewsStatus.TextAlign = ContentAlignment.MiddleLeft;
        mediaNewsStatus.Font = new Font("Segoe UI", 9F);
        mediaNewsStatus.ForeColor = Color.FromArgb(93, 105, 120);
        newsToolbar.Controls.Add(mediaNewsStatus);
        newsToolbar.Controls.Add(mediaNewsCategoryCombo);
        newsToolbar.Controls.Add(domesticCategoryLabel);
        newsToolbar.Controls.Add(mediaNewsRefreshButton);
        mediaNewsList.Dock = DockStyle.Fill;
        mediaNewsList.View = View.Details;
        mediaNewsList.FullRowSelect = true;
        mediaNewsList.HideSelection = false;
        mediaNewsList.MultiSelect = false;
        mediaNewsList.GridLines = true;
        mediaNewsList.BorderStyle = BorderStyle.None;
        mediaNewsList.BackColor = Color.White;
        mediaNewsList.Font = new Font("Segoe UI", 9.5F);
        mediaNewsList.Columns.Add("Kaynak", 150, HorizontalAlignment.Left);
        mediaNewsList.Columns.Add("Kategori", 105, HorizontalAlignment.Left);
        mediaNewsList.Columns.Add("Yayın zamanı", 140, HorizontalAlignment.Left);
        mediaNewsList.Columns.Add("Haber başlığı", 600, HorizontalAlignment.Left);
        mediaNewsList.Resize += (s, e) =>
        {
            int available = mediaNewsList.ClientSize.Width - 150 - 105 - 140 - 24;
            if (available > 80) mediaNewsList.Columns[3].Width = available;
        };
        mediaNewsList.SelectedIndexChanged += (s, e) => UpdateDomesticNewsDetails();
        mediaNewsList.MouseDoubleClick += (s, e) =>
        {
            ListViewItem row = mediaNewsList.GetItemAt(e.X, e.Y);
            MediaHeadline headline = row == null ? null : row.Tag as MediaHeadline;
            if (headline != null) OpenMediaSite(headline.Url);
        };
        mediaNewsDetails.Dock = DockStyle.Fill;
        mediaNewsDetails.ReadOnly = true;
        mediaNewsDetails.BorderStyle = BorderStyle.None;
        mediaNewsDetails.BackColor = Color.FromArgb(250, 251, 253);
        mediaNewsDetails.Font = new Font("Segoe UI", 9F);
        var domesticDetailPanel = new Panel { Dock = DockStyle.Bottom, Height = 118, BackColor = Color.FromArgb(250, 251, 253), Padding = new Padding(10, 5, 10, 7) };
        mediaNewsDetailTitle.Text = "Haber özeti";
        mediaNewsDetailTitle.Dock = DockStyle.Fill;
        mediaNewsDetailTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        mediaNewsDetailTitle.ForeColor = Color.FromArgb(38, 51, 70);
        mediaNewsOpenButton.Text = "Haberi aç";
        mediaNewsOpenButton.Dock = DockStyle.Right;
        mediaNewsOpenButton.Width = 90;
        mediaNewsOpenButton.Enabled = false;
        mediaNewsOpenButton.Click += (s, e) =>
        {
            if (mediaNewsList.SelectedItems.Count > 0 && mediaNewsList.SelectedItems[0].Tag is MediaHeadline)
                OpenMediaSite(((MediaHeadline)mediaNewsList.SelectedItems[0].Tag).Url);
        };
        var domesticDetailHeader = new Panel { Dock = DockStyle.Top, Height = 24 };
        domesticDetailHeader.Controls.Add(mediaNewsDetailTitle);
        domesticDetailHeader.Controls.Add(mediaNewsOpenButton);
        domesticDetailPanel.Controls.Add(mediaNewsDetails);
        domesticDetailPanel.Controls.Add(domesticDetailHeader);
        newsPage.Controls.Add(mediaNewsList);
        newsPage.Controls.Add(domesticDetailPanel);
        newsPage.Controls.Add(newsToolbar);

        var apiPanel = new Panel { Dock = DockStyle.Top, Height = 105, BackColor = Color.White, Padding = new Padding(14, 10, 14, 8) };
        var apiTitle = new Label { Text = "ChatGPT ile kaynak araştırması", AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.FromArgb(38, 51, 70), Location = new Point(14, 8) };
        mediaApiKeyBox.SetBounds(14, 34, 390, 28);
        mediaApiKeyBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        mediaApiKeyBox.UseSystemPasswordChar = true;
        mediaApiKeyBox.AccessibleName = "OpenAI API anahtarı";
        var saveKey = new Button { Text = "Anahtarı kaydet", Size = new Size(112, 28), Location = new Point(414, 34), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var clearKey = new Button { Text = "Anahtarı sil", Size = new Size(88, 28), Location = new Point(532, 34), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        mediaResearchButton.Text = "Kaynakları araştır ve güncelle";
        mediaResearchButton.Size = new Size(205, 28);
        mediaResearchButton.Location = new Point(628, 34);
        mediaResearchButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        mediaResearchButton.Click += (s, e) => StartMediaResearch();
        saveKey.Click += (s, e) => SaveMediaApiKey();
        clearKey.Click += (s, e) => ClearMediaApiKey();
        var apiNote = new Label
        {
            Text = "Anahtar bu bilgisayarda şifreli saklanır. Araştırma, varsayılan tarayıcı yerine OpenAI API web araması kullanır; API kullanımı ChatGPT aboneliğinden ayrı ücretlendirilebilir.",
            AutoSize = false, Height = 26, Location = new Point(14, 66), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            ForeColor = Color.FromArgb(93, 105, 120), Font = new Font("Segoe UI", 8.5F)
        };
        apiPanel.Controls.Add(apiTitle);
        apiPanel.Controls.Add(mediaApiKeyBox);
        apiPanel.Controls.Add(saveKey);
        apiPanel.Controls.Add(clearKey);
        apiPanel.Controls.Add(mediaResearchButton);
        apiPanel.Controls.Add(apiNote);
        apiPanel.Resize += (s, e) =>
        {
            int right = apiPanel.ClientSize.Width - apiPanel.Padding.Right;
            mediaResearchButton.Left = right - mediaResearchButton.Width;
            clearKey.Left = mediaResearchButton.Left - clearKey.Width - 8;
            saveKey.Left = clearKey.Left - saveKey.Width - 8;
            mediaApiKeyBox.Width = Math.Max(150, saveKey.Left - mediaApiKeyBox.Left - 10);
            apiNote.Width = Math.Max(100, right - apiNote.Left);
        };

        var reportPanel = new Panel { Dock = DockStyle.Bottom, Height = 174, BackColor = Color.White, Padding = new Padding(12, 8, 12, 10) };
        var reportHeader = new Panel { Dock = DockStyle.Top, Height = 24 };
        var reportTitle = new Label { Text = "Son araştırma raporu", Dock = DockStyle.Left, Width = 190, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(38, 51, 70), TextAlign = ContentAlignment.MiddleLeft };
        mediaResearchStatus.Text = "Kaynak listesi hazır. Siteler varsayılan web tarayıcınızda açılır.";
        mediaResearchStatus.Dock = DockStyle.Fill;
        mediaResearchStatus.Font = new Font("Segoe UI", 8.5F);
        mediaResearchStatus.ForeColor = Color.FromArgb(93, 105, 120);
        reportHeader.Controls.Add(mediaResearchStatus);
        reportHeader.Controls.Add(reportTitle);
        mediaResearchOutput.Dock = DockStyle.Fill;
        mediaResearchOutput.ReadOnly = true;
        mediaResearchOutput.BorderStyle = BorderStyle.FixedSingle;
        mediaResearchOutput.BackColor = Color.FromArgb(250, 251, 253);
        mediaResearchOutput.Font = new Font("Segoe UI", 9F);
        reportPanel.Controls.Add(mediaResearchOutput);
        reportPanel.Controls.Add(reportHeader);

        var catalog = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(12), BackColor = Color.FromArgb(244, 247, 251) };
        catalog.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        catalog.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        catalog.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        catalog.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        AddMediaCategory(catalog, 0, 0, "Haber ve doğrulama", new Tuple<string, string, string>[]
        {
            Tuple.Create("TRT Haber", "Güncel haberler, videolar ve canlı yayın.", "https://www.trthaber.com/"),
            Tuple.Create("Anadolu Ajansı", "Haber, fotoğraf ve video içerikleri; profesyonel servisleri ayrıca ücretli olabilir.", "https://www.aa.com.tr/tr"),
            Tuple.Create("BBC News Türkçe", "Türkiye ve dünyadan Türkçe haberler.", "https://www.bbc.com/turkce"),
            Tuple.Create("Teyit", "İddia ve görseller için doğrulama çalışmaları.", "https://teyit.org/")
        });
        AddMediaCategory(catalog, 0, 1, "Canlı TV ve video", new Tuple<string, string, string>[]
        {
            Tuple.Create("tabii", "TRT'nin yayın platformu; Türkiye'de ücretsiz erişim, bazı özellikler hesap isteyebilir.", "https://www.tabii.com/tr"),
            Tuple.Create("TRT Haber canlı", "TRT Haber'in resmî canlı yayın sayfası.", "https://www.trthaber.com/canli-yayin-izle.html"),
            Tuple.Create("Diyanet TV", "Canlı yayın ve program arşivi.", "https://www.diyanet.tv/"),
            Tuple.Create("YouTube", "Ücretsiz video platformu; reklamlar ve hesap koşulları uygulanabilir.", "https://www.youtube.com/")
        });
        AddMediaCategory(catalog, 1, 0, "Radyo, podcast ve sesli içerik", new Tuple<string, string, string>[]
        {
            Tuple.Create("TRT Dinle", "Ücretsiz radyo, podcast, sesli kitap ve radyo tiyatrosu içerikleri.", "https://www.trtdinle.com/"),
            Tuple.Create("TRT Radyo kanalları", "TRT'nin ulusal, bölgesel ve uluslararası canlı radyo yayınları.", "https://radyo.trt.net.tr/kanallar"),
            Tuple.Create("Radyo Evi", "Çok sayıda Türkiye radyosunu bir araya getiren ücretsiz dinleme dizini.", "https://radyoevi.net/"),
            Tuple.Create("TRT Haber RSS", "TRT Haber'in ücretsiz RSS başlık akışları.", "https://www.trthaber.com/sitene_ekle.html")
        });
        AddMediaCategory(catalog, 1, 1, "Eğitim ve kültür", new Tuple<string, string, string>[]
        {
            Tuple.Create("MEBİ", "Ücretsiz bireysel öğrenme, sınav hazırlığı ve video çözümleri; giriş gerekebilir.", "https://mebi.eba.gov.tr/"),
            Tuple.Create("EBA", "MEB'in eğitim içerikleri ve etkileşimli öğrenme platformu; kullanıcı girişi gerekebilir.", "https://www.eba.gov.tr/"),
            Tuple.Create("Türkiye Kültür Portalı", "Müzeler, kültürel miras, etkinlik ve şehir rehberleri.", "https://www.kulturportali.gov.tr/"),
            Tuple.Create("TÜBİTAK Bilim Genç", "Bilim haberleri, içerikler ve etkinlikler.", "https://bilimgenc.tubitak.gov.tr/")
        });

        sitesPage.Controls.Add(catalog);
        researchPage.Controls.Add(reportPanel);
        researchPage.Controls.Add(apiPanel);
        page.Controls.Add(mediaTabs);
        LoadMediaApiKey();
        LoadMediaResearchReport();
        mediaNewsWorker = new BackgroundWorker();
        mediaNewsWorker.DoWork += (s, e) => e.Result = FetchMediaNewsHeadlines();
        mediaNewsWorker.RunWorkerCompleted += (s, e) =>
        {
            mediaNewsRefreshButton.Enabled = true;
            mediaNewsLastRefresh = DateTime.Now;
            if (e.Error != null)
            {
                mediaNewsStatus.Text = "Haber akışları alınamadı: " + e.Error.GetBaseException().Message;
                return;
            }
            var result = e.Result as Tuple<List<MediaHeadline>, string>;
            if (result == null) return;
            if (result.Item1.Count > 0)
            {
                PreserveMediaNewsTranslations(result.Item1, mediaNewsCache);
                mediaNewsCache = result.Item1;
                SaveMediaNewsCache();
            }
            ApplyMediaNewsFilter();
            string partial = String.IsNullOrWhiteSpace(result.Item2) ? "" : " • Erişilemeyen akışlar: " + result.Item2;
            mediaNewsStatus.Text = result.Item1.Count.ToString(CultureInfo.CurrentCulture) + " başlık • " + DateTime.Now.ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture) + partial + " • Haberi açmak için çift tıklayın.";
        };
        LoadMediaNewsCache();
        ApplyMediaNewsFilter();
        ApplyWorldNewsFilter();
        page.Enter += (s, e) =>
        {
            if (mediaNewsInitialRefreshStarted) return;
            mediaNewsInitialRefreshStarted = true;
            StartMediaNewsRefresh(false);
        };
        mediaResearchWorker = new BackgroundWorker();
        mediaResearchWorker.DoWork += (s, e) => e.Result = RequestMediaResearch((string)e.Argument);
        mediaResearchWorker.RunWorkerCompleted += (s, e) =>
        {
            mediaResearchButton.Enabled = true;
            if (e.Error != null)
            {
                mediaResearchStatus.Text = "Araştırma başarısız.";
                mediaResearchOutput.Text = e.Error.GetBaseException().Message;
                return;
            }
            string report = Convert.ToString(e.Result, CultureInfo.CurrentCulture);
            mediaResearchOutput.Text = report;
            mediaResearchStatus.Text = "Güncellendi: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture);
            try { File.WriteAllText(mediaResearchFile, mediaResearchStatus.Text + Environment.NewLine + Environment.NewLine + report, new UTF8Encoding(true)); } catch { }
        };
    }

    private sealed class MediaHeadline
    {
        public string Source;
        public string Category;
        public string Title;
        public string Summary;
        public string TranslatedTitle;
        public string TranslatedSummary;
        public string Url;
        public DateTime Published;
    }

    private sealed class WorldNewsProgressSnapshot
    {
        public string Message;
        public List<MediaHeadline> Headlines;
        public int CompletedFeeds;
    }

    private void LoadMediaNewsCache()
    {
        try
        {
            if (!File.Exists(mediaNewsCacheFile)) return;
            var serializer = new JavaScriptSerializer { MaxJsonLength = 20000000 };
            var root = serializer.DeserializeObject(File.ReadAllText(mediaNewsCacheFile, Encoding.UTF8)) as IDictionary<string, object>;
            if (root == null) return;
            mediaNewsCache = DeserializeMediaHeadlines(root.ContainsKey("domestic") ? root["domestic"] : null);
            mediaWorldNewsCache = DeserializeMediaHeadlines(root.ContainsKey("world") ? root["world"] : null);
            object savedAt;
            string timestamp = root.TryGetValue("savedAt", out savedAt) ? Convert.ToString(savedAt, CultureInfo.InvariantCulture) : String.Empty;
            DateTime saved;
            string when = DateTime.TryParse(timestamp, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out saved)
                ? saved.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture) : "önceki oturumdan";
            if (mediaNewsCache.Count > 0) mediaNewsStatus.Text = mediaNewsCache.Count.ToString(CultureInfo.CurrentCulture) + " kayıtlı haber yüklendi (" + when + "); ilk güncelleme bekleniyor.";
            if (mediaWorldNewsCache.Count > 0) mediaWorldStatus.Text = mediaWorldNewsCache.Count.ToString(CultureInfo.CurrentCulture) + " kayıtlı yabancı haber yüklendi (" + when + ").";
        }
        catch { }
    }

    private static List<MediaHeadline> DeserializeMediaHeadlines(object value)
    {
        var output = new List<MediaHeadline>();
        object[] rows = value as object[];
        if (rows == null) return output;
        foreach (object row in rows)
        {
            var fields = row as IDictionary<string, object>;
            if (fields == null) continue;
            Func<string, string> get = key => { object field; return fields.TryGetValue(key, out field) && field != null ? Convert.ToString(field, CultureInfo.InvariantCulture) : String.Empty; };
            DateTime published;
            DateTime.TryParse(get("Published"), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out published);
            output.Add(new MediaHeadline
            {
                Source = get("Source"), Category = get("Category"), Title = get("Title"), Summary = get("Summary"),
                TranslatedTitle = get("TranslatedTitle"), TranslatedSummary = get("TranslatedSummary"), Url = get("Url"), Published = published
            });
        }
        return output;
    }

    private void SaveMediaNewsCache()
    {
        try
        {
            string directory = Path.GetDirectoryName(mediaNewsCacheFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            Func<MediaHeadline, Dictionary<string, object>> serialize = item => new Dictionary<string, object>
            {
                { "Source", item.Source ?? String.Empty }, { "Category", item.Category ?? String.Empty }, { "Title", item.Title ?? String.Empty },
                { "Summary", item.Summary ?? String.Empty }, { "TranslatedTitle", item.TranslatedTitle ?? String.Empty },
                { "TranslatedSummary", item.TranslatedSummary ?? String.Empty }, { "Url", item.Url ?? String.Empty },
                { "Published", item.Published == DateTime.MinValue ? String.Empty : item.Published.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture) }
            };
            var root = new Dictionary<string, object>
            {
                { "savedAt", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) },
                { "domestic", mediaNewsCache.Select(serialize).ToArray() },
                { "world", mediaWorldNewsCache.Select(serialize).ToArray() }
            };
            string json = new JavaScriptSerializer { MaxJsonLength = 20000000 }.Serialize(root);
            string temporary = mediaNewsCacheFile + ".tmp";
            File.WriteAllText(temporary, json, new UTF8Encoding(false));
            if (File.Exists(mediaNewsCacheFile))
            {
                try { File.Replace(temporary, mediaNewsCacheFile, null); }
                catch { File.Delete(mediaNewsCacheFile); File.Move(temporary, mediaNewsCacheFile); }
            }
            else File.Move(temporary, mediaNewsCacheFile);
        }
        catch { }
    }

    private static void PreserveMediaNewsTranslations(List<MediaHeadline> fresh, IEnumerable<MediaHeadline> previous)
    {
        if (fresh == null || previous == null) return;
        var oldByUrl = previous.Where(x => x != null && !String.IsNullOrWhiteSpace(x.Url))
            .GroupBy(x => x.Url, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        foreach (MediaHeadline item in fresh)
        {
            MediaHeadline old;
            if (item == null || String.IsNullOrWhiteSpace(item.Url) || !oldByUrl.TryGetValue(item.Url, out old)) continue;
            if (String.Equals(item.Title, old.Title, StringComparison.Ordinal) && !String.IsNullOrWhiteSpace(old.TranslatedTitle))
                item.TranslatedTitle = old.TranslatedTitle;
            if (String.Equals(item.Summary, old.Summary, StringComparison.Ordinal) && !String.IsNullOrWhiteSpace(old.TranslatedSummary))
                item.TranslatedSummary = old.TranslatedSummary;
        }
    }

    private void InitializeWorldNewsPage(TabPage page)
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = Color.White, Padding = new Padding(10, 7, 10, 5) };
        var worldStatusTip = new ToolTip();
        worldStatusTip.SetToolTip(mediaWorldStatus, "Akış yenileme sonucu ve başarısız kaynak ayrıntıları burada gösterilir.");
        mediaWorldStatus.Text = "Yabancı kaynakların RSS akışları hazır.";
        mediaWorldStatus.SetBounds(10, 9, 270, 28);
        mediaWorldStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        mediaWorldStatus.TextAlign = ContentAlignment.MiddleLeft;
        mediaWorldStatus.Font = new Font("Segoe UI", 8.5F);
        mediaWorldStatus.ForeColor = Color.FromArgb(93, 105, 120);
        mediaWorldCategoryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        mediaWorldCategoryCombo.Items.AddRange(new object[] { "Tüm kategoriler", "Dünyadan", "Ekonomi", "Spor", "Teknoloji", "Bilim" });
        mediaWorldCategoryCombo.SelectedIndex = 0;
        mediaWorldCategoryCombo.SetBounds(0, 8, 140, 28);
        mediaWorldCategoryCombo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        mediaWorldRefreshButton.Text = "Yenile";
        mediaWorldRefreshButton.SetBounds(0, 8, 82, 28);
        mediaWorldRefreshButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        mediaWorldRefreshButton.FlatStyle = FlatStyle.Flat;
        mediaWorldRefreshButton.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        mediaWorldRefreshButton.BackColor = Color.FromArgb(239, 245, 255);
        mediaWorldTranslateButton.Text = "Çevrimdışı çevir";
        mediaWorldTranslateButton.SetBounds(0, 8, 125, 28);
        mediaWorldTranslateButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        mediaWorldTranslateButton.FlatStyle = FlatStyle.Flat;
        mediaWorldTranslateButton.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
        mediaWorldTranslateButton.BackColor = Color.FromArgb(239, 245, 255);
        var translationTip = new ToolTip();
        translationTip.SetToolTip(mediaWorldTranslateButton, "Argos English–Türkçe motoru. İlk kullanımda bu bilgisayara ücretsiz bileşenler indirilir; sonraki çeviriler internet ve API anahtarı olmadan yapılır.");
        var categoryLabel = new Label { Text = "Kategori:", AutoSize = true, Anchor = AnchorStyles.Top | AnchorStyles.Right, TextAlign = ContentAlignment.MiddleRight };
        categoryLabel.SetBounds(0, 13, 55, 18);
        toolbar.Controls.Add(mediaWorldStatus);
        toolbar.Controls.Add(categoryLabel);
        toolbar.Controls.Add(mediaWorldCategoryCombo);
        toolbar.Controls.Add(mediaWorldRefreshButton);
        toolbar.Controls.Add(mediaWorldTranslateButton);
        toolbar.Resize += (s, e) =>
        {
            int right = toolbar.ClientSize.Width - toolbar.Padding.Right;
            mediaWorldTranslateButton.Left = right - mediaWorldTranslateButton.Width;
            mediaWorldRefreshButton.Left = mediaWorldTranslateButton.Left - mediaWorldRefreshButton.Width - 6;
            mediaWorldCategoryCombo.Left = mediaWorldRefreshButton.Left - mediaWorldCategoryCombo.Width - 6;
            categoryLabel.Left = mediaWorldCategoryCombo.Left - categoryLabel.Width - 4;
            mediaWorldStatus.Width = Math.Max(80, categoryLabel.Left - mediaWorldStatus.Left - 8);
        };

        mediaWorldNewsList.Dock = DockStyle.Fill;
        mediaWorldNewsList.View = View.Details;
        mediaWorldNewsList.FullRowSelect = true;
        mediaWorldNewsList.HideSelection = false;
        mediaWorldNewsList.MultiSelect = false;
        mediaWorldNewsList.GridLines = true;
        mediaWorldNewsList.BorderStyle = BorderStyle.None;
        mediaWorldNewsList.BackColor = Color.White;
        mediaWorldNewsList.Font = new Font("Segoe UI", 9.5F);
        mediaWorldNewsList.Columns.Add("Kaynak", 150, HorizontalAlignment.Left);
        mediaWorldNewsList.Columns.Add("Kategori", 105, HorizontalAlignment.Left);
        mediaWorldNewsList.Columns.Add("Yayın zamanı", 140, HorizontalAlignment.Left);
        mediaWorldNewsList.Columns.Add("Başlık (Türkçe çeviri varsa)", 550, HorizontalAlignment.Left);
        mediaWorldNewsList.SelectedIndexChanged += (s, e) => UpdateWorldNewsDetails();
        mediaWorldNewsList.MouseDoubleClick += (s, e) =>
        {
            ListViewItem row = mediaWorldNewsList.GetItemAt(e.X, e.Y);
            MediaHeadline item = row == null ? null : row.Tag as MediaHeadline;
            if (item != null) OpenMediaSite(item.Url);
        };
        mediaWorldNewsList.Resize += (s, e) => ResizeWorldNewsColumns();
        mediaWorldDetails.Dock = DockStyle.Fill;
        mediaWorldDetails.ReadOnly = true;
        mediaWorldDetails.BorderStyle = BorderStyle.None;
        mediaWorldDetails.BackColor = Color.FromArgb(250, 251, 253);
        mediaWorldDetails.Font = new Font("Segoe UI", 9F);
        var detailPanel = new Panel { Dock = DockStyle.Bottom, Height = 150, BackColor = Color.FromArgb(250, 251, 253), Padding = new Padding(10, 6, 10, 8) };
        var detailHeader = new Panel { Dock = DockStyle.Top, Height = 27 };
        mediaWorldDetailTitle.Text = "Haber özeti";
        mediaWorldDetailTitle.Dock = DockStyle.Fill;
        mediaWorldDetailTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        mediaWorldDetailTitle.ForeColor = Color.FromArgb(38, 51, 70);
        mediaWorldOpenButton.Text = "Haberi aç";
        mediaWorldOpenButton.Dock = DockStyle.Right;
        mediaWorldOpenButton.Width = 92;
        mediaWorldOpenButton.Enabled = false;
        mediaWorldOpenButton.Click += (s, e) =>
        {
            if (mediaWorldNewsList.SelectedItems.Count > 0 && mediaWorldNewsList.SelectedItems[0].Tag is MediaHeadline)
                OpenMediaSite(((MediaHeadline)mediaWorldNewsList.SelectedItems[0].Tag).Url);
        };
        detailHeader.Controls.Add(mediaWorldDetailTitle);
        detailHeader.Controls.Add(mediaWorldOpenButton);
        detailPanel.Controls.Add(mediaWorldDetails);
        detailPanel.Controls.Add(detailHeader);
        page.Controls.Add(mediaWorldNewsList);
        page.Controls.Add(detailPanel);
        page.Controls.Add(toolbar);

        mediaWorldCategoryCombo.SelectedIndexChanged += (s, e) => ApplyWorldNewsFilter();
        mediaWorldRefreshButton.Click += (s, e) => StartWorldNewsRefresh(true);
        mediaWorldTranslateButton.Click += (s, e) => StartWorldNewsTranslation();
        mediaWorldNewsWorker = new BackgroundWorker();
        mediaWorldNewsWorker.WorkerReportsProgress = true;
        mediaWorldNewsWorker.DoWork += (s, e) => e.Result = FetchWorldNewsHeadlines(s as BackgroundWorker);
        mediaWorldNewsWorker.ProgressChanged += (s, e) =>
        {
            WorldNewsProgressSnapshot progress = e.UserState as WorldNewsProgressSnapshot;
            if (progress == null) return;
            if (progress.CompletedFeeds >= mediaWorldDisplayedProgressCount)
            {
                mediaWorldDisplayedProgressCount = progress.CompletedFeeds;
                if (progress.Headlines != null && progress.Headlines.Count > 0)
                {
                    PreserveMediaNewsTranslations(progress.Headlines, mediaWorldNewsTranslationBaseline);
                    mediaWorldNewsCache = progress.Headlines;
                    ApplyWorldNewsFilter();
                }
                mediaWorldStatus.Text = progress.Message;
            }
        };
        mediaWorldNewsWorker.RunWorkerCompleted += (s, e) =>
        {
            mediaWorldRefreshButton.Enabled = true;
            if (e.Error != null) { mediaWorldNewsTranslationBaseline.Clear(); mediaWorldStatus.Text = "Yabancı haberler alınamadı: " + e.Error.GetBaseException().Message; return; }
            var result = e.Result as Tuple<List<MediaHeadline>, string>;
            if (result == null) return;
            mediaWorldNewsLastRefresh = DateTime.Now;
            if (result.Item1.Count > 0)
            {
                PreserveMediaNewsTranslations(result.Item1, mediaWorldNewsTranslationBaseline);
                mediaWorldNewsCache = result.Item1;
                SaveMediaNewsCache();
            }
            mediaWorldNewsTranslationBaseline.Clear();
            ApplyWorldNewsFilter();
            string issue = String.IsNullOrWhiteSpace(result.Item2) ? "" : " • Sorunlu akışlar: " + result.Item2;
            mediaWorldStatus.Text = result.Item1.Count == 0
                ? (mediaWorldNewsCache.Count > 0 ? "Yeni haber alınamadı; önceki liste korunuyor." : "Akışlardan haber alınamadı.") + issue
                : mediaWorldNewsCache.Count.ToString(CultureInfo.CurrentCulture) + " haber • " + DateTime.Now.ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture) + issue;
            mediaWorldStatus.ForeColor = result.Item1.Count == 0 ? Color.Firebrick : Color.FromArgb(93, 105, 120);
            worldStatusTip.SetToolTip(mediaWorldStatus, String.IsNullOrWhiteSpace(result.Item2) ? mediaWorldStatus.Text : "Başarısız akışlar: " + result.Item2);
        };
        mediaWorldTranslateWorker = new BackgroundWorker();
        mediaWorldTranslateWorker.WorkerReportsProgress = true;
        mediaWorldTranslateWorker.DoWork += (s, e) => TranslateWorldNews((List<MediaHeadline>)e.Argument, s as BackgroundWorker);
        mediaWorldTranslateWorker.ProgressChanged += (s, e) =>
        {
            if (e.UserState is string) mediaWorldStatus.Text = (string)e.UserState;
        };
        mediaWorldTranslateWorker.RunWorkerCompleted += (s, e) =>
        {
            mediaWorldTranslateButton.Enabled = true;
            if (e.Error != null)
            {
                Exception error = e.Error.GetBaseException();
                mediaWorldStatus.Text = "Çeviri kurulumu/işlemi başarısız: " + error.Message;
                ShowArgosDiagnostic(e.Error);
                return;
            }
            SaveMediaNewsCache();
            ApplyWorldNewsFilter();
            mediaWorldStatus.Text = "Haber başlıkları ve RSS özetleri bu bilgisayarda çevrildi. • " + DateTime.Now.ToString("HH:mm", CultureInfo.CurrentCulture);
        };
        page.Enter += (s, e) =>
        {
            if (mediaWorldNewsInitialRefreshStarted) return;
            mediaWorldNewsInitialRefreshStarted = true;
            StartWorldNewsRefresh(false);
        };
    }

    private void ResizeWorldNewsColumns()
    {
        if (mediaWorldNewsList.Columns.Count < 4) return;
        int remainder = mediaWorldNewsList.ClientSize.Width - 150 - 105 - 140 - 24;
        if (remainder > 80) mediaWorldNewsList.Columns[3].Width = remainder;
    }

    private void ApplyWorldNewsFilter()
    {
        if (mediaWorldNewsList.IsDisposed) return;
        string category = Convert.ToString(mediaWorldCategoryCombo.SelectedItem, CultureInfo.CurrentCulture);
        mediaWorldNewsList.BeginUpdate();
        mediaWorldNewsList.Items.Clear();
        foreach (MediaHeadline headline in mediaWorldNewsCache.Where(x => category == "Tüm kategoriler" || String.Equals(x.Category, category, StringComparison.CurrentCultureIgnoreCase)))
        {
            var row = new ListViewItem(headline.Source);
            row.SubItems.Add(headline.Category ?? "");
            row.SubItems.Add(headline.Published == DateTime.MinValue ? "" : headline.Published.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture));
            row.SubItems.Add(String.IsNullOrWhiteSpace(headline.TranslatedTitle) ? headline.Title : headline.TranslatedTitle);
            row.Tag = headline;
            mediaWorldNewsList.Items.Add(row);
        }
        mediaWorldNewsList.EndUpdate();
        UpdateWorldNewsDetails();
        ResizeWorldNewsColumns();
    }

    private void UpdateWorldNewsDetails()
    {
        MediaHeadline item = mediaWorldNewsList.SelectedItems.Count == 0 ? null : mediaWorldNewsList.SelectedItems[0].Tag as MediaHeadline;
        if (item == null)
        {
            mediaWorldDetailTitle.Text = "Haber özeti";
            mediaWorldDetails.Clear();
            mediaWorldOpenButton.Enabled = false;
            return;
        }
        mediaWorldDetailTitle.Text = item.Source + " • " + item.Category + (item.Published == DateTime.MinValue ? "" : " • " + item.Published.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture));
        string summary = String.IsNullOrWhiteSpace(item.TranslatedSummary) ? item.Summary : item.TranslatedSummary;
        mediaWorldDetails.Text = (String.IsNullOrWhiteSpace(item.TranslatedTitle) ? item.Title : item.TranslatedTitle) + Environment.NewLine + Environment.NewLine +
            (String.IsNullOrWhiteSpace(summary) ? "Kaynak, RSS akışında ayrıca bir özet sunmamış." : summary);
        mediaWorldOpenButton.Enabled = true;
    }

    private void StartWorldNewsRefresh(bool force)
    {
        if (mediaWorldNewsWorker == null || mediaWorldNewsWorker.IsBusy) return;
        if (!force && mediaWorldNewsLastRefresh != DateTime.MinValue && DateTime.Now - mediaWorldNewsLastRefresh < TimeSpan.FromMinutes(5)) return;
        mediaWorldNewsTranslationBaseline = mediaWorldNewsCache.ToList();
        mediaWorldDisplayedProgressCount = -1;
        mediaWorldRefreshButton.Enabled = false;
        mediaWorldStatus.Text = "Yabancı haber akışları kontrol ediliyor...";
        mediaWorldNewsWorker.RunWorkerAsync();
    }

    private List<MediaHeadline> FetchWorldNewsHeadlines(BackgroundWorker worker)
    {
        var feeds = new Tuple<string, string, string>[]
        {
            Tuple.Create("Google News", "Dünyadan", "https://news.google.com/rss/headlines/section/topic/WORLD?hl=en-US&gl=US&ceid=US:en"),
            Tuple.Create("Google News", "Ekonomi", "https://news.google.com/rss/headlines/section/topic/BUSINESS?hl=en-US&gl=US&ceid=US:en"),
            Tuple.Create("Google News", "Spor", "https://news.google.com/rss/headlines/section/topic/SPORTS?hl=en-US&gl=US&ceid=US:en"),
            Tuple.Create("Google News", "Teknoloji", "https://news.google.com/rss/headlines/section/topic/TECHNOLOGY?hl=en-US&gl=US&ceid=US:en"),
            Tuple.Create("Google News", "Bilim", "https://news.google.com/rss/headlines/section/topic/SCIENCE?hl=en-US&gl=US&ceid=US:en"),
            Tuple.Create("Google News", "Dünyadan", "https://news.google.com/rss?hl=en-US&gl=US&ceid=US:en"),
            Tuple.Create("BBC News", "Dünyadan", "https://feeds.bbci.co.uk/news/world/rss.xml"),
            Tuple.Create("BBC News", "Ekonomi", "https://feeds.bbci.co.uk/news/business/rss.xml"),
            Tuple.Create("BBC Sport", "Spor", "https://feeds.bbci.co.uk/sport/rss.xml"),
            Tuple.Create("BBC News", "Teknoloji", "https://feeds.bbci.co.uk/news/technology/rss.xml"),
            Tuple.Create("The Guardian", "Dünyadan", "https://www.theguardian.com/world/rss"),
            Tuple.Create("The Guardian", "Ekonomi", "https://www.theguardian.com/uk/business/rss"),
            Tuple.Create("The Guardian", "Spor", "https://www.theguardian.com/uk/sport/rss"),
            Tuple.Create("The Guardian", "Bilim", "https://www.theguardian.com/science/rss"),
            Tuple.Create("Le Monde (English)", "Dünyadan", "https://www.lemonde.fr/en/international/rss_full.xml"),
            Tuple.Create("Le Monde (English)", "Spor", "https://www.lemonde.fr/en/sports/rss_full.xml")
        };
        var results = new ConcurrentBag<MediaHeadline>();
        var failed = new ConcurrentBag<string>();
        var knownUrls = new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        int completedFeeds = 0;
        Parallel.ForEach(feeds, new ParallelOptions { MaxDegreeOfParallelism = 8 }, feed =>
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(feed.Item3);
                request.Method = "GET";
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36";
                request.Accept = "application/rss+xml, application/atom+xml, application/xml, text/xml, */*";
                request.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9,tr;q=0.8";
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Timeout = 7000;
                request.ReadWriteTimeout = 7000;
                using (var response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreComments = true }))
                {
                    var document = new XmlDocument { XmlResolver = null };
                    document.Load(reader);
                    XmlNodeList items = document.SelectNodes("//*[local-name()='item' or local-name()='entry']");
                    if (items == null || items.Count == 0) { failed.Add(feed.Item1 + " / " + feed.Item2 + ": RSS içinde haber girdisi yok"); return; }
                    int accepted = 0;
                    foreach (XmlNode node in items)
                    {
                        string title = CleanMediaFeedText(GetMediaXmlValue(node, "title"));
                        string url = GetMediaNewsLink(node);
                        if (String.IsNullOrWhiteSpace(title) || String.IsNullOrWhiteSpace(url) || !knownUrls.TryAdd(url, true)) continue;
                        string rawSummary = GetMediaXmlValue(node, "encoded");
                        if (String.IsNullOrWhiteSpace(rawSummary)) rawSummary = GetMediaXmlValue(node, "description");
                        if (String.IsNullOrWhiteSpace(rawSummary)) rawSummary = GetMediaXmlValue(node, "summary");
                        if (String.IsNullOrWhiteSpace(rawSummary)) rawSummary = GetMediaXmlValue(node, "content");
                        string summary = CleanMediaFeedText(rawSummary);
                        if (summary.Length > 1200) summary = summary.Substring(0, 1200).TrimEnd() + "…";
                        results.Add(new MediaHeadline
                        {
                            Source = feed.Item1, Category = feed.Item2, Title = title, Summary = summary, Url = url,
                            Published = ParseMediaNewsDate(GetMediaXmlValue(node, "pubDate"), GetMediaXmlValue(node, "published"), GetMediaXmlValue(node, "updated"), GetMediaXmlValue(node, "date"))
                        });
                        if (++accepted >= 12) break;
                    }
                }
            }
            catch (Exception ex)
            {
                WebException webError = ex as WebException;
                string detail = webError != null && webError.Response is HttpWebResponse
                    ? "HTTP " + (int)((HttpWebResponse)webError.Response).StatusCode
                    : ex.GetBaseException().Message;
                failed.Add(feed.Item1 + " / " + feed.Item2 + ": " + detail);
            }
            finally
            {
                int completed = Interlocked.Increment(ref completedFeeds);
                if (worker != null && worker.WorkerReportsProgress)
                {
                    var snapshot = results.OrderByDescending(x => x.Published).Take(100).ToList();
                    worker.ReportProgress(0, new WorldNewsProgressSnapshot
                    {
                        Message = "Yabancı haber akışları kontrol ediliyor: " + completed.ToString(CultureInfo.CurrentCulture) + "/" + feeds.Length.ToString(CultureInfo.CurrentCulture) + " kaynak • " + results.Count.ToString(CultureInfo.CurrentCulture) + " haber bulundu.",
                        Headlines = snapshot,
                        CompletedFeeds = completed
                    });
                }
            }
        });
        return results.OrderByDescending(x => x.Published).Take(100).ToList();
    }

    private static string CleanMediaFeedText(string text)
    {
        if (String.IsNullOrWhiteSpace(text)) return String.Empty;
        string clean = Regex.Replace(text, "<[^>]+>", " ");
        clean = WebUtility.HtmlDecode(clean);
        return Regex.Replace(clean, "\\s+", " ").Trim();
    }

    private void StartWorldNewsTranslation()
    {
        if (mediaWorldNewsCache.Count == 0) { MessageBox.Show(this, "Önce yabancı haber akışlarını yenileyin.", "Çeviri", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        if (mediaWorldTranslateWorker.IsBusy) return;
        string category = Convert.ToString(mediaWorldCategoryCombo.SelectedItem, CultureInfo.CurrentCulture);
        List<MediaHeadline> items = mediaWorldNewsCache.Where(x => category == "Tüm kategoriler" || String.Equals(x.Category, category, StringComparison.CurrentCultureIgnoreCase)).ToList();
        if (items.Count == 0) return;
        if (!IsArgosEngineReady())
        {
            DialogResult answer = MessageBox.Show(this,
                "Ücretsiz yerel çeviri motoru hazırlanacak. Önce bilgisayarda uyumlu bir Python kurulumu aranır; bulunursa mevcut Python kullanılır. Bulunamazsa yalnızca uygulamanın kendi çalışma klasörüne Python kurulur. Argos bileşenleri ve yaklaşık 119 MB İngilizce–Türkçe modeli de indirilir. Devam edilsin mi?",
                "Yerel çeviri motoru", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) return;
        }
        mediaWorldTranslateButton.Enabled = false;
        mediaWorldStatus.Text = items.Count.ToString(CultureInfo.CurrentCulture) + " haber yerel motorla Türkçeye çevriliyor...";
        mediaWorldTranslateWorker.RunWorkerAsync(items);
    }

    private string ArgosPythonExePath()
    {
        string privatePython = Path.Combine(argosRuntimeRoot, "Python313", "python.exe");
        if (File.Exists(privatePython)) return privatePython;
        string storedPython = Path.Combine(argosRuntimeRoot, "python-executable.txt");
        try
        {
            if (File.Exists(storedPython))
            {
                string saved = File.ReadAllText(storedPython, Encoding.UTF8).Trim();
                if (File.Exists(saved)) return saved;
            }
        }
        catch { }
        if (!systemPythonDiscoveryComplete)
        {
            discoveredSystemPythonPath = FindInstalledPythonExecutable();
            systemPythonDiscoveryComplete = true;
        }
        return discoveredSystemPythonPath ?? privatePython;
    }

    private string FindInstalledPythonExecutable()
    {
        ArgosLog("PYTHON ARAMA BAŞLADI | PATH, kullanıcı Python klasörleri ve PythonCore kayıtları kontrol ediliyor.");
        var candidates = new List<string>();
        try
        {
            foreach (string folder in (Environment.GetEnvironmentVariable("PATH") ?? String.Empty).Split(Path.PathSeparator))
            {
                if (String.IsNullOrWhiteSpace(folder)) continue;
                candidates.Add(Path.Combine(folder.Trim('"'), "python.exe"));
                candidates.Add(Path.Combine(folder.Trim('"'), "python3.exe"));
            }
        }
        catch { }
        try
        {
            string localPrograms = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python");
            if (Directory.Exists(localPrograms))
                foreach (string directory in Directory.GetDirectories(localPrograms).OrderByDescending(path => path, StringComparer.OrdinalIgnoreCase))
                    candidates.Add(Path.Combine(directory, "python.exe"));
        }
        catch { }
        foreach (RegistryKey hive in new RegistryKey[] { Registry.CurrentUser, Registry.LocalMachine })
        {
            try
            {
                using (RegistryKey pythonCore = hive.OpenSubKey("Software\\Python\\PythonCore"))
                {
                    if (pythonCore == null) continue;
                    foreach (string version in pythonCore.GetSubKeyNames())
                    using (RegistryKey install = pythonCore.OpenSubKey(version + "\\InstallPath"))
                    {
                        if (install == null) continue;
                        string executable = install.GetValue("ExecutablePath") as string;
                        if (!String.IsNullOrWhiteSpace(executable)) candidates.Add(executable);
                        string directory = install.GetValue(null) as string;
                        if (!String.IsNullOrWhiteSpace(directory)) candidates.Add(Path.Combine(directory, "python.exe"));
                    }
                }
            }
            catch { }
        }
        foreach (string candidate in candidates.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (candidate.IndexOf("\\WindowsApps\\", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ArgosLog("PYTHON ADAYI ATLANDI | Microsoft Store yönlendirme takma adı: " + candidate);
                continue;
            }
            try
            {
                var info = new ProcessStartInfo
                {
                    FileName = candidate,
                    Arguments = "-c \"import sys; print('%d.%d' % sys.version_info[:2])\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (Process process = Process.Start(info))
                {
                    if (process == null) continue;
                    if (!process.WaitForExit(3000))
                    {
                        try { process.Kill(); } catch { }
                        ArgosLog("PYTHON ADAYI ZAMAN AŞIMI | " + candidate);
                        continue;
                    }
                    string version = process.StandardOutput.ReadToEnd().Trim();
                    Version parsed;
                    if (process.ExitCode == 0 && Version.TryParse(version, out parsed) && parsed.Major == 3 && parsed.Minor >= 8 && parsed.Minor <= 14)
                    {
                        ArgosLog("UYUMLU PYTHON BULUNDU | Sürüm=" + version + " | Yol=" + candidate);
                        return candidate;
                    }
                    ArgosLog("PYTHON ADAYI UYGUN DEĞİL | Yol=" + candidate + " | Sürüm çıktısı=" + version + " | Çıkış kodu=" + process.ExitCode);
                }
            }
            catch (Exception ex) { ArgosLog("PYTHON ADAYI ÇALIŞTIRILAMADI | Yol=" + candidate + " | " + ex.GetType().Name + ": " + ex.Message); }
        }
        ArgosLog("UYUMLU PYTHON BULUNAMADI | Uygulamaya özel Python kurulumu kullanılacak.");
        return null;
    }

    private string ArgosPackagesPath()
    {
        return Path.Combine(argosRuntimeRoot, "packages");
    }

    private string ArgosPythonPackagesPath()
    {
        return Path.Combine(argosRuntimeRoot, "python-packages");
    }

    private string ArgosReadyPath()
    {
        return Path.Combine(argosRuntimeRoot, "en-tr.ready");
    }

    private bool IsArgosEngineReady()
    {
        string privatePython = Path.Combine(argosRuntimeRoot, "Python313", "python.exe");
        string savedPython = Path.Combine(argosRuntimeRoot, "python-executable.txt");
        string librariesReady = Path.Combine(argosRuntimeRoot, "argos-libraries.ready");
        bool hasPython = File.Exists(privatePython);
        try { if (!hasPython && File.Exists(savedPython)) hasPython = File.Exists(File.ReadAllText(savedPython, Encoding.UTF8).Trim()); } catch { }
        return hasPython && File.Exists(librariesReady) && File.Exists(ArgosReadyPath());
    }

    private void TranslateWorldNews(List<MediaHeadline> items, BackgroundWorker worker)
    {
        EnsureArgosTranslator(worker);
        Directory.CreateDirectory(argosRuntimeRoot);
        string inputFile = Path.Combine(argosRuntimeRoot, "translate-input-" + Guid.NewGuid().ToString("N") + ".json");
        string outputFile = Path.Combine(argosRuntimeRoot, "translate-output-" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            var inputItems = items.Select(x => new Dictionary<string, object> { { "title", x.Title }, { "summary", x.Summary ?? "" } }).ToArray();
            var serializer = new JavaScriptSerializer { MaxJsonLength = 10000000 };
            File.WriteAllText(inputFile, serializer.Serialize(inputItems), new UTF8Encoding(false));
            worker.ReportProgress(0, "Argos çevrimdışı motoruyla " + items.Count.ToString(CultureInfo.CurrentCulture) + " haber çevriliyor...");
            RunArgosPython("translate " + QuoteProcessArgument(argosRuntimeRoot) + " " + QuoteProcessArgument(inputFile) + " " + QuoteProcessArgument(outputFile), worker, "Haberler çevriliyor...");
            object parsed = serializer.DeserializeObject(File.ReadAllText(outputFile, Encoding.UTF8));
            object[] translated = parsed as object[];
            if (translated == null || translated.Length != items.Count) throw new InvalidOperationException("Yerel çeviri yanıtı beklenen biçimde değil.");
            for (int i = 0; i < translated.Length; i++)
            {
                var row = translated[i] as IDictionary<string, object>;
                if (row == null) continue;
                object title, summary;
                if (row.TryGetValue("title", out title)) items[i].TranslatedTitle = Convert.ToString(title, CultureInfo.CurrentCulture);
                if (row.TryGetValue("summary", out summary)) items[i].TranslatedSummary = Convert.ToString(summary, CultureInfo.CurrentCulture);
            }
        }
        finally
        {
            try { if (File.Exists(inputFile)) File.Delete(inputFile); } catch { }
            try { if (File.Exists(outputFile)) File.Delete(outputFile); } catch { }
        }
    }

    private void EnsureArgosTranslator(BackgroundWorker worker)
    {
        Directory.CreateDirectory(argosRuntimeRoot);
        Directory.CreateDirectory(ArgosPackagesPath());
        string python = ArgosPythonExePath();
        ArgosLog("BAŞLANGIÇ | Python adayı=" + python + " | Var=" + File.Exists(python) + " | Windows=" + Environment.OSVersion);
        if (!File.Exists(python))
        {
            string installer = Path.Combine(argosRuntimeRoot, "python-3.13.17-amd64.exe");
            ReportArgosProgress(worker, "Resmî Python çalışma zamanı indiriliyor...");
            DownloadArgosFile(ArgosPythonInstallerUrl, installer, worker, "Python çalışma zamanı indiriliyor");
            string actualHash;
            using (SHA256 sha = SHA256.Create())
            using (FileStream input = File.OpenRead(installer))
                actualHash = BitConverter.ToString(sha.ComputeHash(input)).Replace("-", "").ToLowerInvariant();
            if (!String.Equals(actualHash, ArgosPythonInstallerSha256, StringComparison.OrdinalIgnoreCase))
            {
                try { File.Delete(installer); } catch { }
                throw new InvalidOperationException("Python kurulum dosyasının SHA-256 doğrulaması başarısız oldu.");
            }
            string installDirectory = Path.GetDirectoryName(python);
            string installArguments = "/quiet InstallAllUsers=0 TargetDir=" + QuoteProcessArgument(installDirectory) +
                " Include_pip=1 Include_launcher=0 Include_doc=0 Include_tcltk=0 Include_test=0 Include_tools=0 PrependPath=0 Shortcuts=0 AssociateFiles=0";
            ReportArgosProgress(worker, "Python bu Windows hesabına kuruluyor...");
            RunArgosProcess(installer, installArguments, null, worker, "Python yerel olarak kuruluyor...");
            try { File.Delete(installer); } catch { }
            if (!File.Exists(python)) throw new InvalidOperationException("Python yükleyicisi çalıştı ancak beklenen python.exe oluşmadı. Beklenen konum: " + python + ". Yükleme klasörüne yazma izni, Smart App Control/antivirüs engeli veya yükleyici hedef yolu bu duruma neden olabilir.");
        }
        else if (!String.Equals(Path.GetFullPath(python), Path.GetFullPath(Path.Combine(argosRuntimeRoot, "Python313", "python.exe")), StringComparison.OrdinalIgnoreCase))
        {
            try { File.WriteAllText(Path.Combine(argosRuntimeRoot, "python-executable.txt"), python, new UTF8Encoding(false)); } catch { }
            ReportArgosProgress(worker, "Bilgisayarda uyumlu Python bulundu; mevcut kurulum kullanılacak: " + python);
        }

        string librariesReady = Path.Combine(argosRuntimeRoot, "argos-libraries.ready");
        if (!File.Exists(librariesReady))
        {
            ReportArgosProgress(worker, "Argos Translate bileşenleri indiriliyor ve kuruluyor...");
            Directory.CreateDirectory(ArgosPythonPackagesPath());
            string target = QuoteProcessArgument(ArgosPythonPackagesPath());
            string pipArguments = "-m pip install --disable-pip-version-check --no-warn-script-location --upgrade --target " + target + " ctranslate2 sentencepiece sacremoses minisbd packaging";
            string supportOutput = RunArgosProcess(python, pipArguments, ArgosRuntimeEnvironment(), worker, "Hafif çeviri bağımlılıkları kuruluyor...");
            string argosArguments = "-m pip install --disable-pip-version-check --no-warn-script-location --upgrade --target " + target + " --no-deps argostranslate==" + ArgosPackageVersion;
            string argosOutput = RunArgosProcess(python, argosArguments, ArgosRuntimeEnvironment(), worker, "Argos Translate kuruluyor (Stanza/Spacy ek paketleri atlanıyor)...");
            File.WriteAllText(librariesReady, "argostranslate " + ArgosPackageVersion + Environment.NewLine + supportOutput + Environment.NewLine + argosOutput, new UTF8Encoding(false));
        }

        if (!File.Exists(ArgosReadyPath()))
        {
            string modelFile = Path.Combine(argosRuntimeRoot, ArgosModelFileName);
            ReportArgosProgress(worker, "İngilizce–Türkçe modeli (yaklaşık 119 MB) indiriliyor...");
            DownloadArgosFile(ArgosModelUrl, modelFile, worker, "Argos İngilizce–Türkçe modeli indiriliyor");
            using (var archive = System.IO.Compression.ZipFile.OpenRead(modelFile))
            {
                if (!archive.Entries.Any(entry => String.Equals(Path.GetFileName(entry.FullName), "metadata.json", StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("İndirilen Argos modeli geçerli değil (metadata.json bulunamadı).");
            }
            ReportArgosProgress(worker, "İngilizce–Türkçe modeli yerel motor dizinine kuruluyor...");
            RunArgosPython("install " + QuoteProcessArgument(argosRuntimeRoot) + " " + QuoteProcessArgument(modelFile), worker, "Argos modeli kuruluyor...");
            File.WriteAllText(ArgosReadyPath(), "English -> Turkish, Argos 1.5" + Environment.NewLine, new UTF8Encoding(false));
            try { File.Delete(modelFile); } catch { }
        }
    }

    private void DownloadArgosFile(string url, string targetPath, BackgroundWorker worker, string description)
    {
        string partialPath = targetPath + ".part";
        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.UserAgent = "DosyaGezgini/ArgosBootstrap";
            request.Timeout = 180000;
            request.ReadWriteTimeout = 180000;
            using (var response = (HttpWebResponse)request.GetResponse())
            using (Stream input = response.GetResponseStream())
            using (var output = new FileStream(partialPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                long total = response.ContentLength;
                long received = 0;
                int previousPercent = -1;
                byte[] buffer = new byte[128 * 1024];
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    output.Write(buffer, 0, read);
                    received += read;
                    if (total > 0)
                    {
                        int percent = (int)Math.Min(100, received * 100 / total);
                        if (percent != previousPercent)
                        {
                            previousPercent = percent;
                            ReportArgosProgress(worker, description + " • %" + percent.ToString(CultureInfo.CurrentCulture));
                        }
                    }
                }
            }
            if (File.Exists(targetPath)) File.Delete(targetPath);
            File.Move(partialPath, targetPath);
        }
        catch (Exception ex)
        {
            try { if (File.Exists(partialPath)) File.Delete(partialPath); } catch { }
            ArgosLog("İNDİRME HATASI | Aşama=" + description + " | URL=" + url + " | Hedef=" + targetPath + " | " + ex);
            throw new InvalidOperationException(description + " başarısız oldu. URL: " + url + ". Hedef: " + targetPath + ". Ayrıntı: " + ex.GetType().Name + ": " + ex.Message, ex);
        }
    }

    private string RunArgosPython(string arguments, BackgroundWorker worker, string progressMessage)
    {
        string script = Path.Combine(argosRuntimeRoot, "argos_translate.py");
        if (!File.Exists(script))
        {
            using (Stream resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("DosyaGezgini.argos_translate.py"))
            {
                if (resource == null) throw new InvalidOperationException("Argos Python köprü dosyası program kaynağında bulunamadı.");
                using (var destination = new FileStream(script, FileMode.Create, FileAccess.Write, FileShare.None)) resource.CopyTo(destination);
            }
        }
        return RunArgosProcess(ArgosPythonExePath(), QuoteProcessArgument(script) + " " + arguments, ArgosRuntimeEnvironment(), worker, progressMessage);
    }

    private Dictionary<string, string> ArgosRuntimeEnvironment()
    {
        string previousPythonPath = Environment.GetEnvironmentVariable("PYTHONPATH");
        string pythonPath = ArgosPythonPackagesPath();
        if (!String.IsNullOrWhiteSpace(previousPythonPath)) pythonPath += Path.PathSeparator + previousPythonPath;
        return new Dictionary<string, string>
        {
            { "ARGOS_PACKAGES_DIR", ArgosPackagesPath() },
            { "ARGOS_DEVICE_TYPE", "cpu" },
            { "ARGOS_CHUNK_TYPE", "MINISBD" },
            { "PYTHONUTF8", "1" },
            { "PYTHONPATH", pythonPath },
            { "PIP_DISABLE_PIP_VERSION_CHECK", "1" }
        };
    }

    private string RunArgosProcess(string executable, string arguments, Dictionary<string, string> environment, BackgroundWorker worker, string progressMessage)
    {
        ReportArgosProgress(worker, progressMessage);
        ArgosLog("İŞLEM BAŞLADI | Aşama=" + progressMessage + " | Program=" + executable + " | Argümanlar=" + arguments);
        var info = new ProcessStartInfo
        {
            FileName = executable,
            Arguments = arguments,
            WorkingDirectory = argosRuntimeRoot,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        if (environment != null) foreach (KeyValuePair<string, string> item in environment) info.EnvironmentVariables[item.Key] = item.Value;
        var log = new StringBuilder();
        using (var process = new Process { StartInfo = info })
        {
            process.OutputDataReceived += (s, e) =>
            {
                if (e.Data == null) return;
                lock (log) log.AppendLine(e.Data);
                if (e.Data.StartsWith("Collecting ", StringComparison.OrdinalIgnoreCase) ||
                    e.Data.StartsWith("Downloading ", StringComparison.OrdinalIgnoreCase) ||
                    e.Data.StartsWith("Installing collected packages", StringComparison.OrdinalIgnoreCase) ||
                    e.Data.StartsWith("Successfully installed", StringComparison.OrdinalIgnoreCase))
                    ReportArgosProgress(worker, progressMessage + " • " + e.Data);
            };
            process.ErrorDataReceived += (s, e) =>
            {
                if (e.Data == null) return;
                lock (log) log.AppendLine(e.Data);
                if (e.Data.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) || e.Data.StartsWith("WARNING", StringComparison.OrdinalIgnoreCase))
                    ReportArgosProgress(worker, e.Data);
            };
            if (!process.Start()) throw new InvalidOperationException("Argos yardımcı işlemi başlatılamadı.");
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            while (!process.WaitForExit(250))
            {
                if (worker != null && worker.CancellationPending)
                {
                    try { process.Kill(); } catch { }
                    throw new OperationCanceledException("Argos işlemi kullanıcı tarafından durduruldu.");
                }
            }
            process.WaitForExit();
            string output;
            lock (log) output = log.ToString().Trim();
            ArgosLog("İŞLEM BİTTİ | Aşama=" + progressMessage + " | Çıkış kodu=" + process.ExitCode + Environment.NewLine + output);
            if (process.ExitCode != 0)
                throw new InvalidOperationException("Yardımcı işlem başarısız oldu. Aşama: " + progressMessage + ". Program: " + executable + ". Çıkış kodu: " + process.ExitCode.ToString(CultureInfo.CurrentCulture) + ". Çıktı: " + output);
            return output;
        }
    }

    private void ReportArgosProgress(BackgroundWorker worker, string message)
    {
        ArgosLog("DURUM | " + message);
        if (worker != null && worker.WorkerReportsProgress) worker.ReportProgress(0, message);
    }

    private string ArgosLogPath { get { return Path.Combine(argosRuntimeRoot, "argos-diagnostic.log"); } }

    private void ArgosLog(string message)
    {
        try
        {
            Directory.CreateDirectory(argosRuntimeRoot);
            File.AppendAllText(ArgosLogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + " | " + message + Environment.NewLine, new UTF8Encoding(false));
        }
        catch { }
    }

    private void ShowArgosDiagnostic(Exception error)
    {
        string details = "Argos çevrimdışı çeviri kurulumu veya çalıştırılması başarısız oldu." + Environment.NewLine +
            "Windows: " + Environment.OSVersion + " (64-bit işletim sistemi: " + Environment.Is64BitOperatingSystem + ")" + Environment.NewLine +
            "Seçilen Python: " + ArgosPythonExePath() + Environment.NewLine +
            "Argos çalışma klasörü: " + argosRuntimeRoot + Environment.NewLine +
            "Tanılama günlüğü: " + ArgosLogPath + Environment.NewLine + Environment.NewLine +
            "Hata ayrıntısı:" + Environment.NewLine + error;
        try
        {
            if (File.Exists(ArgosLogPath))
            {
                string log = File.ReadAllText(ArgosLogPath, Encoding.UTF8);
                if (log.Length > 24000) log = log.Substring(log.Length - 24000);
                details += Environment.NewLine + Environment.NewLine + "Son kurulum kayıtları:" + Environment.NewLine + log;
            }
        }
        catch { }
        using (var dialog = new Form { Text = "Çevrimdışı çeviri tanılama", StartPosition = FormStartPosition.CenterParent, Size = new Size(820, 620), MinimumSize = new Size(640, 420), ShowInTaskbar = false })
        {
            var text = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill, Font = new Font("Consolas", 9F), Text = details };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
            var close = new Button { Text = "Kapat", Width = 90, DialogResult = DialogResult.OK };
            var copy = new Button { Text = "Ayrıntıları kopyala", Width = 145 };
            copy.Click += (s, e) => { try { Clipboard.SetText(details); } catch { } };
            var openLog = new Button { Text = "Günlük dosyasını aç", Width = 145 };
            openLog.Click += (s, e) => { try { if (File.Exists(ArgosLogPath)) Process.Start(new ProcessStartInfo(ArgosLogPath) { UseShellExecute = true }); } catch { } };
            buttons.Controls.Add(close); buttons.Controls.Add(copy); buttons.Controls.Add(openLog);
            dialog.Controls.Add(text); dialog.Controls.Add(buttons); dialog.AcceptButton = close;
            dialog.ShowDialog(this);
        }
    }

    private void StartMediaNewsRefresh(bool force)
    {
        if (mediaNewsWorker == null || mediaNewsWorker.IsBusy) return;
        if (!force && mediaNewsLastRefresh != DateTime.MinValue && DateTime.Now - mediaNewsLastRefresh < TimeSpan.FromMinutes(5)) return;
        mediaNewsRefreshButton.Enabled = false;
        mediaNewsStatus.Text = "Haber akışları kontrol ediliyor...";
        mediaNewsWorker.RunWorkerAsync();
    }

    private void ApplyMediaNewsFilter()
    {
        if (mediaNewsList.IsDisposed || mediaNewsList.Columns.Count < 4) return;
        string category = Convert.ToString(mediaNewsCategoryCombo.SelectedItem, CultureInfo.CurrentCulture);
        mediaNewsList.BeginUpdate();
        mediaNewsList.Items.Clear();
        foreach (MediaHeadline headline in mediaNewsCache.Where(x => category == "Tüm haberler" || String.Equals(x.Category, category, StringComparison.CurrentCultureIgnoreCase)))
        {
            var row = new ListViewItem(headline.Source);
            row.SubItems.Add(headline.Category ?? "");
            row.SubItems.Add(headline.Published == DateTime.MinValue ? "" : headline.Published.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture));
            row.SubItems.Add(headline.Title);
            row.Tag = headline;
            mediaNewsList.Items.Add(row);
        }
        mediaNewsList.EndUpdate();
        UpdateDomesticNewsDetails();
    }

    private void UpdateDomesticNewsDetails()
    {
        MediaHeadline item = mediaNewsList.SelectedItems.Count == 0 ? null : mediaNewsList.SelectedItems[0].Tag as MediaHeadline;
        if (item == null)
        {
            mediaNewsDetailTitle.Text = "Haber özeti";
            mediaNewsDetails.Clear();
            mediaNewsOpenButton.Enabled = false;
            return;
        }
        mediaNewsDetailTitle.Text = item.Source + " • " + item.Category + (item.Published == DateTime.MinValue ? "" : " • " + item.Published.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture));
        mediaNewsDetails.Text = String.IsNullOrWhiteSpace(item.Summary) ? "Kaynak, RSS akışında ayrıca bir özet sunmamış." : item.Summary;
        mediaNewsOpenButton.Enabled = true;
    }

    private Tuple<List<MediaHeadline>, string> FetchMediaNewsHeadlines()
    {
        var feeds = new Tuple<string, string, string>[]
        {
            Tuple.Create("TRT Haber", "Manşet", "https://www.trthaber.com/manset_articles.rss"),
            Tuple.Create("TRT Haber", "Son dakika", "https://www.trthaber.com/sondakika_articles.rss"),
            Tuple.Create("TRT Haber", "Gündem", "https://www.trthaber.com/gundem_articles.rss"),
            Tuple.Create("TRT Haber", "Türkiye", "https://www.trthaber.com/turkiye_articles.rss"),
            Tuple.Create("TRT Haber", "Dünyadan", "https://www.trthaber.com/dunya_articles.rss"),
            Tuple.Create("TRT Haber", "Ekonomi", "https://www.trthaber.com/ekonomi_articles.rss"),
            Tuple.Create("TRT Haber", "Spor", "https://www.trthaber.com/spor_articles.rss"),
            Tuple.Create("TRT Haber", "Yaşam", "https://www.trthaber.com/yasam_articles.rss"),
            Tuple.Create("TRT Haber", "Sağlık", "https://www.trthaber.com/saglik_articles.rss"),
            Tuple.Create("TRT Haber", "Kültür-Sanat", "https://www.trthaber.com/kultur_sanat_articles.rss"),
            Tuple.Create("TRT Haber", "Teknoloji", "https://www.trthaber.com/bilim_teknoloji_articles.rss"),
            Tuple.Create("TRT Haber", "Güncel", "https://www.trthaber.com/guncel_articles.rss"),
            Tuple.Create("TRT Haber", "Eğitim", "https://www.trthaber.com/egitim_articles.rss"),
            Tuple.Create("Habertürk", "Manşet", "https://www.haberturk.com/rss/manset.xml"),
            Tuple.Create("Habertürk", "Ekonomi", "https://www.haberturk.com/rss/ekonomi.xml"),
            Tuple.Create("Habertürk", "Spor", "https://www.haberturk.com/rss/spor.xml"),
            Tuple.Create("Habertürk", "Dünyadan", "https://www.haberturk.com/rss/kategori/dunya.xml"),
            Tuple.Create("Habertürk", "Gündem", "https://www.haberturk.com/rss/kategori/gundem.xml"),
            Tuple.Create("Habertürk", "Sağlık", "https://www.haberturk.com/rss/kategori/saglik.xml"),
            Tuple.Create("Habertürk", "Yaşam", "https://www.haberturk.com/rss/kategori/yasam.xml"),
            Tuple.Create("Habertürk", "Kültür-Sanat", "https://www.haberturk.com/rss/kategori/kultur-sanat.xml"),
            Tuple.Create("Habertürk", "Teknoloji", "https://www.haberturk.com/rss/kategori/teknoloji.xml"),
            Tuple.Create("Habertürk", "İş / Yaşam", "https://www.haberturk.com/rss/kategori/is-yasam.xml"),
            Tuple.Create("Anadolu Ajansı", "Güncel", "https://www.aa.com.tr/tr/rss/default?cat=guncel")
        };
        var results = new ConcurrentBag<MediaHeadline>();
        var failed = new ConcurrentBag<string>();
        var knownUrls = new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        Parallel.ForEach(feeds, new ParallelOptions { MaxDegreeOfParallelism = 6 }, feed =>
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(feed.Item3);
                request.Method = "GET";
                request.UserAgent = "DosyaGezgini/1.0 RSS Reader";
                request.Accept = "application/rss+xml, application/atom+xml, application/xml, text/xml, */*";
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Timeout = 15000;
                request.ReadWriteTimeout = 15000;
                using (var response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreComments = true }))
                {
                    var document = new XmlDocument { XmlResolver = null };
                    document.Load(reader);
                    XmlNodeList items = document.SelectNodes("//*[local-name()='item' or local-name()='entry']");
                    if (items == null) return;
                    int accepted = 0;
                    foreach (XmlNode item in items)
                    {
                        string title = CleanMediaFeedText(GetMediaXmlValue(item, "title"));
                        string url = GetMediaNewsLink(item);
                        if (String.IsNullOrWhiteSpace(title) || String.IsNullOrWhiteSpace(url) || !knownUrls.TryAdd(url, true)) continue;
                        DateTime published = ParseMediaNewsDate(GetMediaXmlValue(item, "pubDate"), GetMediaXmlValue(item, "published"), GetMediaXmlValue(item, "updated"), GetMediaXmlValue(item, "date"));
                        string summary = GetMediaXmlValue(item, "encoded");
                        if (String.IsNullOrWhiteSpace(summary)) summary = GetMediaXmlValue(item, "description");
                        if (String.IsNullOrWhiteSpace(summary)) summary = GetMediaXmlValue(item, "summary");
                        summary = CleanMediaFeedText(summary);
                        if (summary.Length > 1200) summary = summary.Substring(0, 1200).TrimEnd() + "…";
                        results.Add(new MediaHeadline { Source = feed.Item1, Category = feed.Item2, Title = title, Summary = summary, Url = url, Published = published });
                        accepted++;
                        if (accepted >= 10) break;
                    }
                }
            }
            catch { failed.Add(feed.Item1 + " / " + feed.Item2); }
        });
        return Tuple.Create(results.OrderByDescending(headline => headline.Published).ToList(), String.Join(", ", failed));
    }

    private static string GetMediaXmlValue(XmlNode parent, string localName)
    {
        XmlNode node = parent == null ? null : parent.SelectSingleNode("./*[local-name()='" + localName + "']");
        return node == null ? String.Empty : node.InnerText;
    }

    private static string GetMediaNewsLink(XmlNode item)
    {
        XmlNode link = item == null ? null : item.SelectSingleNode("./*[local-name()='link']");
        string value = link == null ? String.Empty : link.InnerText.Trim();
        if (String.IsNullOrWhiteSpace(value) && link != null && link.Attributes != null && link.Attributes["href"] != null)
            value = link.Attributes["href"].Value;
        Uri uri;
        return Uri.TryCreate(value, UriKind.Absolute, out uri) &&
            (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) ? uri.AbsoluteUri : String.Empty;
    }

    private static DateTime ParseMediaNewsDate(params string[] candidates)
    {
        foreach (string candidate in candidates)
        {
            DateTime parsed;
            if (!String.IsNullOrWhiteSpace(candidate) && DateTime.TryParse(candidate, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed)) return parsed;
        }
        return DateTime.MinValue;
    }

    private void AddMediaCategory(TableLayoutPanel catalog, int row, int column, string title, IEnumerable<Tuple<string, string, string>> sites)
    {
        var group = new GroupBox { Text = title, Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(38, 51, 70), BackColor = Color.White, Padding = new Padding(10) };
        var list = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new Padding(5, 8, 5, 4), BackColor = Color.White };
        foreach (Tuple<string, string, string> site in sites)
        {
            var item = new Panel { Height = 51, Width = 430, Margin = new Padding(2, 1, 2, 2), BackColor = Color.White };
            var name = new LinkLabel { Text = site.Item1, AutoSize = true, Location = new Point(2, 1), Font = new Font("Segoe UI", 9F, FontStyle.Bold), LinkColor = Color.FromArgb(21, 91, 158), ActiveLinkColor = Color.FromArgb(15, 67, 117), Tag = site.Item3 };
            name.Links.Add(0, site.Item1.Length, site.Item3);
            name.LinkClicked += (s, e) => OpenMediaSite(Convert.ToString(((Control)s).Tag, CultureInfo.InvariantCulture));
            var description = new Label { Text = site.Item2, Location = new Point(3, 23), Height = 25, AutoEllipsis = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.FromArgb(93, 105, 120), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            item.Controls.Add(name);
            item.Controls.Add(description);
            item.Resize += (s, e) => description.Width = Math.Max(80, item.ClientSize.Width - 8);
            list.Controls.Add(item);
            var separator = new Panel { Height = 1, Width = 420, BackColor = Color.FromArgb(235, 239, 244), Margin = new Padding(2, 0, 2, 1) };
            list.Controls.Add(separator);
        }
        group.Controls.Add(list);
        catalog.Controls.Add(group, column, row);
    }

    private void OpenMediaSite(string url)
    {
        if (String.IsNullOrWhiteSpace(url)) return;
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Site açılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void SaveMediaApiKey()
    {
        string key = (mediaApiKeyBox.Text ?? String.Empty).Trim();
        if (key.Length < 20)
        {
            MessageBox.Show(this, "Geçerli bir OpenAI API anahtarı girin.", "API anahtarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try
        {
            string directory = Path.GetDirectoryName(mediaApiKeyFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            byte[] encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(key), null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(mediaApiKeyFile, encrypted);
            mediaApiKey = key;
            mediaApiKeyBox.Clear();
            mediaResearchStatus.Text = "API anahtarı bu Windows hesabı için şifreli kaydedildi.";
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "API anahtarı kaydedilemedi", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }

    private void ClearMediaApiKey()
    {
        try { if (File.Exists(mediaApiKeyFile)) File.Delete(mediaApiKeyFile); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "API anahtarı silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        mediaApiKey = String.Empty;
        mediaApiKeyBox.Clear();
        mediaResearchStatus.Text = "API anahtarı silindi.";
    }

    private void LoadMediaApiKey()
    {
        try
        {
            if (!File.Exists(mediaApiKeyFile)) return;
            byte[] encrypted = File.ReadAllBytes(mediaApiKeyFile);
            mediaApiKey = Encoding.UTF8.GetString(ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser));
        }
        catch { mediaApiKey = String.Empty; }
    }

    private void LoadMediaResearchReport()
    {
        try
        {
            if (!File.Exists(mediaResearchFile)) return;
            string saved = File.ReadAllText(mediaResearchFile, Encoding.UTF8);
            int split = saved.IndexOf(Environment.NewLine + Environment.NewLine, StringComparison.Ordinal);
            if (split >= 0)
            {
                mediaResearchStatus.Text = saved.Substring(0, split).Trim();
                mediaResearchOutput.Text = saved.Substring(split + 2 * Environment.NewLine.Length);
            }
            else mediaResearchOutput.Text = saved;
        }
        catch { }
    }

    private void StartMediaResearch()
    {
        if (String.IsNullOrWhiteSpace(mediaApiKey))
        {
            MessageBox.Show(this, "Önce OpenAI API anahtarınızı girip kaydedin.", "Kaynak araştırması", MessageBoxButtons.OK, MessageBoxIcon.Information);
            mediaApiKeyBox.Focus();
            return;
        }
        if (mediaResearchWorker == null || mediaResearchWorker.IsBusy) return;
        mediaResearchButton.Enabled = false;
        mediaResearchStatus.Text = "Web kaynakları araştırılıyor...";
        mediaResearchWorker.RunWorkerAsync(mediaApiKey);
    }

    private string RequestMediaResearch(string apiKey)
    {
        var sites = new string[]
        {
            "TRT Haber — https://www.trthaber.com/ — news, live video, RSS",
            "Anadolu Ajansı — https://www.aa.com.tr/tr — news, photo, video; some professional services may be paid",
            "BBC News Türkçe — https://www.bbc.com/turkce — Turkish news",
            "Teyit — https://teyit.org/ — fact checking",
            "tabii — https://www.tabii.com/tr — TRT streaming; free access in Türkiye, account/features may vary",
            "TRT Haber live — https://www.trthaber.com/canli-yayin-izle.html",
            "Diyanet TV — https://www.diyanet.tv/",
            "YouTube — https://www.youtube.com/ — free access with ads/account limits",
            "TRT Dinle — https://www.trtdinle.com/ — radio, podcasts and audio library",
            "TRT radio — https://radyo.trt.net.tr/kanallar — live radio channels",
            "Radyo Evi — https://radyoevi.net/ — free radio directory",
            "TRT Haber RSS — https://www.trthaber.com/sitene_ekle.html",
            "MEBİ — https://mebi.eba.gov.tr/ — free learning platform; sign-in may be required",
            "EBA — https://www.eba.gov.tr/ — education platform; sign-in may be required",
            "Türkiye Kültür Portalı — https://www.kulturportali.gov.tr/ — culture, heritage, museums and events",
            "TÜBİTAK Bilim Genç — https://bilimgenc.tubitak.gov.tr/ — science content"
        };
        string prompt = "Türkçe bir kaynak güncelleme raporu hazırla. Bugün " + DateTime.Now.ToString("dd MMMM yyyy", CultureInfo.GetCultureInfo("tr-TR")) + ". " +
            "Aşağıdaki Türkiye'den ücretsiz veya ücretsiz erişim katmanı olan haber, video, canlı TV/radyo, eğitim ve kültür sitelerini web'de araştır. " +
            "Her site için erişilebilirlik ve doğru resmî adresi, ücretsiz/hesap/reklam/bölge koşullarını, yakın dönemde hizmet veya içerik değişikliği varsa kısa özetini ver. " +
            "Yanlış, kapanmış veya değişmiş bağlantıları özellikle belirt ve önerdiğin yeni URL'yi göster. Kaynak olarak tıklanabilir tam bağlantılar ekle. " +
            "Site içeriğini doğrudan kesin bilgi gibi varsayma; doğrulayamadıklarını 'doğrulanamadı' diye yaz. Rapor sonunda eklenebilecek 3 yeni ücretsiz Türkiye kaynağı öner.\n\n" + String.Join("\n", sites);
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        var requestBody = new Dictionary<string, object>
        {
            { "model", "gpt-6-luna" },
            { "tools", new object[] { new Dictionary<string, object> { { "type", "web_search" } } } },
            { "tool_choice", "auto" },
            { "max_output_tokens", 2800 },
            { "input", prompt }
        };
        var serializer = new JavaScriptSerializer { MaxJsonLength = 10000000 };
        byte[] body = Encoding.UTF8.GetBytes(serializer.Serialize(requestBody));
        var request = (HttpWebRequest)WebRequest.Create("https://api.openai.com/v1/responses");
        request.Method = "POST";
        request.ContentType = "application/json";
        request.Accept = "application/json";
        request.Headers[HttpRequestHeader.Authorization] = "Bearer " + apiKey;
        request.Timeout = 120000;
        request.ReadWriteTimeout = 120000;
        using (Stream output = request.GetRequestStream()) output.Write(body, 0, body.Length);
        using (var response = (HttpWebResponse)request.GetResponse())
        using (var input = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
        {
            object parsed = serializer.DeserializeObject(input.ReadToEnd());
            var root = parsed as IDictionary<string, object>;
            if (root == null) throw new InvalidOperationException("API yanıtı beklenen biçimde değil.");
            object outputItems;
            if (!root.TryGetValue("output", out outputItems)) throw new InvalidOperationException("API yanıtında araştırma sonucu bulunamadı.");
            var report = new StringBuilder();
            foreach (object outputItem in outputItems as object[] ?? new object[0])
            {
                var item = outputItem as IDictionary<string, object>;
                if (item == null || !String.Equals(Convert.ToString(item.ContainsKey("type") ? item["type"] : null), "message", StringComparison.Ordinal)) continue;
                object content;
                if (!item.TryGetValue("content", out content)) continue;
                foreach (object block in content as object[] ?? new object[0])
                {
                    var part = block as IDictionary<string, object>;
                    if (part == null || !String.Equals(Convert.ToString(part.ContainsKey("type") ? part["type"] : null), "output_text", StringComparison.Ordinal)) continue;
                    object text;
                    if (part.TryGetValue("text", out text)) report.AppendLine(Convert.ToString(text, CultureInfo.CurrentCulture));
                }
            }
            if (report.Length == 0) throw new InvalidOperationException("Araştırma metni yanıtta bulunamadı.");
            return report.ToString().Trim();
        }
    }

    private string SendMediaResponsesRequest(Dictionary<string, object> requestBody, string apiKey)
    {
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        var serializer = new JavaScriptSerializer { MaxJsonLength = 10000000 };
        byte[] body = Encoding.UTF8.GetBytes(serializer.Serialize(requestBody));
        var request = (HttpWebRequest)WebRequest.Create("https://api.openai.com/v1/responses");
        request.Method = "POST";
        request.ContentType = "application/json";
        request.Accept = "application/json";
        request.Headers[HttpRequestHeader.Authorization] = "Bearer " + apiKey;
        request.Timeout = 120000;
        request.ReadWriteTimeout = 120000;
        using (Stream output = request.GetRequestStream()) output.Write(body, 0, body.Length);
        using (var response = (HttpWebResponse)request.GetResponse())
        using (var input = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
        {
            object parsed = serializer.DeserializeObject(input.ReadToEnd());
            var root = parsed as IDictionary<string, object>;
            if (root == null) throw new InvalidOperationException("API yanıtı beklenen biçimde değil.");
            object outputItems;
            if (!root.TryGetValue("output", out outputItems)) throw new InvalidOperationException("API yanıtında metin bulunamadı.");
            var outputText = new StringBuilder();
            foreach (object outputItem in outputItems as object[] ?? new object[0])
            {
                var item = outputItem as IDictionary<string, object>;
                if (item == null || !String.Equals(Convert.ToString(item.ContainsKey("type") ? item["type"] : null), "message", StringComparison.Ordinal)) continue;
                object content;
                if (!item.TryGetValue("content", out content)) continue;
                foreach (object block in content as object[] ?? new object[0])
                {
                    var part = block as IDictionary<string, object>;
                    if (part == null || !String.Equals(Convert.ToString(part.ContainsKey("type") ? part["type"] : null), "output_text", StringComparison.Ordinal)) continue;
                    object text;
                    if (part.TryGetValue("text", out text)) outputText.AppendLine(Convert.ToString(text, CultureInfo.CurrentCulture));
                }
            }
            if (outputText.Length == 0) throw new InvalidOperationException("API yanıtında çeviri metni bulunamadı.");
            return outputText.ToString().Trim();
        }
    }

    private void InitializeFolderTabBar(Control explorerView, TabPage explorerPage)
    {
        explorerWorkspace.Dock = DockStyle.Fill;
        var tabStrip = new Panel { Dock = DockStyle.Top, Height = 32, BackColor = Color.FromArgb(246, 248, 251) };
        folderTabs.Dock = DockStyle.Fill;
        folderTabs.HotTrack = true;
        folderTabs.SizeMode = TabSizeMode.Normal;
        folderTabs.AllowDrop = true;
        int dragTabIndex = -1;
        folderTabs.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) dragTabIndex = GetFolderTabAt(e.Location); };
        folderTabs.MouseMove += (s, e) =>
        {
            if (dragTabIndex < 0 || (Control.MouseButtons & MouseButtons.Left) == 0) return;
            TabPage moving = folderTabs.TabPages[dragTabIndex];
            dragTabIndex = -1;
            folderTabs.DoDragDrop(moving, DragDropEffects.Move);
        };
        folderTabs.DragOver += (s, e) => e.Effect = e.Data.GetDataPresent(typeof(TabPage)) ? DragDropEffects.Move : DragDropEffects.None;
        folderTabs.DragDrop += (s, e) =>
        {
            TabPage moving = e.Data.GetData(typeof(TabPage)) as TabPage;
            int target = GetFolderTabAt(folderTabs.PointToClient(new Point(e.X, e.Y)));
            if (moving == null || target < 0 || moving == folderTabs.TabPages[target]) return;
            int oldIndex = folderTabs.TabPages.IndexOf(moving);
            folderTabs.TabPages.Remove(moving);
            folderTabs.TabPages.Insert(Math.Min(target, folderTabs.TabPages.Count), moving);
            folderTabs.SelectedTab = moving;
            SaveFolderTabSessions();
        };
        newFolderTabButton.Text = "+";
        newFolderTabButton.Dock = DockStyle.Right;
        newFolderTabButton.Width = 36;
        newFolderTabButton.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
        newFolderTabButton.Margin = Padding.Empty;
        newFolderTabButton.Click += (s, e) => CreateFolderTab(currentPath);
        actionTooltips.SetToolTip(newFolderTabButton, "Yeni klasör sekmesi (Ctrl+T)");
        tabStrip.Controls.Add(folderTabs);
        tabStrip.Controls.Add(newFolderTabButton);
        folderTabs.SelectedIndexChanged += (s, e) => SwitchFolderTab();
        folderTabs.MouseClick += (s, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            int index = GetFolderTabAt(e.Location);
            if (index < 0) return;
            TabPage clickedTab = folderTabs.TabPages[index];
            if (folderTabs.SelectedTab != clickedTab) folderTabs.SelectedTab = clickedTab;
            else ActivateFolderTab(clickedTab);
        };
        folderTabs.MouseUp += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            int index = GetFolderTabAt(e.Location);
            if (index < 0) return;
            contextFolderTab = folderTabs.TabPages[index];
            if (folderTabs.SelectedTab != contextFolderTab) folderTabs.SelectedTab = contextFolderTab;
        };
        folderTabs.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Middle) return;
            int index = GetFolderTabAt(e.Location);
            if (index >= 0) CloseFolderTab(folderTabs.TabPages[index]);
        };
        var menu = new ContextMenuStrip();
        var newTabItem = new ToolStripMenuItem("Yeni sekme");
        var renameTabItem = new ToolStripMenuItem("Sekmeye isim ver...");
        var closeItem = new ToolStripMenuItem("Sekmeyi kapat");
        var closeOthersItem = new ToolStripMenuItem("Diğer sekmeleri kapat");
        var reopenItem = new ToolStripMenuItem("Kapatılan sekmeyi geri aç");
        newTabItem.Click += (s, e) => CreateFolderTab(currentPath);
        renameTabItem.Click += (s, e) => RenameFolderTab(contextFolderTab ?? folderTabs.SelectedTab);
        closeItem.Click += (s, e) => CloseFolderTab(contextFolderTab ?? folderTabs.SelectedTab);
        closeOthersItem.Click += (s, e) => CloseOtherFolderTabs(contextFolderTab ?? folderTabs.SelectedTab);
        reopenItem.Click += (s, e) => ReopenClosedFolderTab();
        menu.Items.Add(newTabItem);
        menu.Items.Add(renameTabItem);
        menu.Items.Add(closeItem);
        menu.Items.Add(closeOthersItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(reopenItem);
        menu.Opening += (s, e) =>
        {
            if (contextFolderTab == null) contextFolderTab = folderTabs.SelectedTab;
            closeItem.Enabled = folderTabs.TabPages.Count > 1;
            renameTabItem.Enabled = contextFolderTab != null;
            closeOthersItem.Enabled = folderTabs.TabPages.Count > 1;
            reopenItem.Enabled = recentlyClosedFolderTabs.Count > 0;
        };
        folderTabs.ContextMenuStrip = menu;
        explorerWorkspace.Controls.Add(explorerView);
        explorerWorkspace.Controls.Add(tabStrip);
        explorerPage.Controls.Add(explorerWorkspace);
    }

    private int GetFolderTabAt(Point point)
    {
        for (int i = 0; i < folderTabs.TabPages.Count; i++)
            if (folderTabs.GetTabRect(i).Contains(point)) return i;
        return -1;
    }

    private TabPage CreateFolderTabPage(string path, IEnumerable<string> historyPaths = null, int historyIndex = -1, string customName = null)
    {
        var recent = TakeLastItems((historyPaths ?? new string[] { path }).Where(item => !String.IsNullOrWhiteSpace(item)), 20);
        if (recent.Count == 0 && !String.IsNullOrWhiteSpace(path)) recent.Add(path);
        int index = recent.Count == 0 ? -1 : Math.Max(0, Math.Min(recent.Count - 1, historyIndex < 0 ? recent.Count - 1 : historyIndex));
        var state = new FolderTabState(path, recent, index) { CustomName = customName ?? String.Empty };
        string title = String.IsNullOrWhiteSpace(state.CustomName) ? GetFolderTabTitle(path) : state.CustomName;
        var page = new TabPage(title) { Tag = state, ToolTipText = path };
        return page;
    }

    private static string GetFolderTabTitle(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return "Gezgin";
        string trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string name = Path.GetFileName(trimmed);
        return String.IsNullOrEmpty(name) ? path : name;
    }

    private static List<T> TakeLastItems<T>(IEnumerable<T> source, int count)
    {
        List<T> list = source.ToList();
        return list.Skip(Math.Max(0, list.Count - count)).ToList();
    }

    private void LoadFolderTabSessions()
    {
        suppressFolderTabEvents = true;
        int selectedIndex = 0;
        try
        {
            folderTabs.TabPages.Clear();
            if (File.Exists(folderTabsFile))
            {
                var document = new XmlDocument();
                document.Load(folderTabsFile);
                Int32.TryParse(document.DocumentElement == null || document.DocumentElement.Attributes["selected"] == null
                    ? null : document.DocumentElement.Attributes["selected"].Value, out selectedIndex);
                foreach (XmlNode node in document.SelectNodes("/FolderTabs/Tab"))
                {
                    string path = ResolvePortablePath(node.Attributes["path"] == null ? String.Empty : node.Attributes["path"].Value);
                    if (String.IsNullOrWhiteSpace(path)) continue;
                    var history = new List<string>();
                    foreach (XmlNode item in node.SelectNodes("History/Path"))
                    {
                        string value = ResolvePortablePath(item.Attributes["value"] == null ? String.Empty : item.Attributes["value"].Value);
                        if (!String.IsNullOrWhiteSpace(value)) history.Add(value);
                    }
                    int index = -1;
                    Int32.TryParse(node.Attributes["historyIndex"] == null ? null : node.Attributes["historyIndex"].Value, out index);
                    string customName = node.Attributes["name"] == null ? String.Empty : node.Attributes["name"].Value;
                    TabPage page = CreateFolderTabPage(path, history, index, customName);
                    FolderTabState state = page.Tag as FolderTabState;
                    string shortcutPath = ResolvePortablePath(node.Attributes["shortcutPath"] == null ? String.Empty : node.Attributes["shortcutPath"].Value);
                    if (state != null) state.ShortcutPath = !String.IsNullOrWhiteSpace(shortcutPath)
                        ? shortcutPath : (String.IsNullOrWhiteSpace(customName) ? String.Empty : path);
                    folderTabs.TabPages.Add(page);
                }
            }
        }
        catch { folderTabs.TabPages.Clear(); }
        if (folderTabs.TabPages.Count == 0)
        {
            List<string> initialHistory = TakeLastItems(navigationHistory.Where(Directory.Exists), 20);
            if (!initialHistory.Contains(currentPath, StringComparer.OrdinalIgnoreCase)) initialHistory.Add(currentPath);
            folderTabs.TabPages.Add(CreateFolderTabPage(currentPath, initialHistory, initialHistory.Count - 1));
        }
        folderTabs.SelectedIndex = Math.Max(0, Math.Min(folderTabs.TabPages.Count - 1, selectedIndex));
        activeFolderTab = folderTabs.SelectedTab;
        RestoreFolderTabNavigation(activeFolderTab);
        FolderTabState selectedState = activeFolderTab == null ? null : activeFolderTab.Tag as FolderTabState;
        if (selectedState != null && selectedState.IsUnavailable)
            status.Text = "Sekmenin klasörü şu anda bağlı değil; yol saklandı: " + selectedState.Path;
        suppressFolderTabEvents = false;
    }

    private void SwitchFolderTab()
    {
        if (suppressFolderTabEvents || folderTabs.SelectedTab == null) return;
        CaptureActiveFolderTab();
        activeFolderTab = folderTabs.SelectedTab;
        ActivateFolderTab(activeFolderTab);
        UpdateNavigationButtons();
        SaveFolderTabSessions();
    }

    private void ActivateFolderTab(TabPage page)
    {
        if (suppressFolderTabEvents || page == null) return;
        FolderTabState state = page.Tag as FolderTabState;
        if (state == null) return;
        if (page != activeFolderTab)
        {
            CaptureActiveFolderTab();
            activeFolderTab = page;
        }
        bool wasUnavailable = state.IsUnavailable;
        string configuredRoot = state.ShortcutPath ?? String.Empty;
        string targetPath = String.IsNullOrWhiteSpace(state.Path) ? currentPath : state.Path;
        // A tab remembers its own last folder, but never restores outside its configured root.
        // Moving above that root therefore returns to the root when the tab is selected again.
        if (!String.IsNullOrWhiteSpace(configuredRoot) && !IsPathWithin(targetPath, configuredRoot)) targetPath = configuredRoot;
        if (!Directory.Exists(targetPath))
        {
            if (!String.IsNullOrWhiteSpace(configuredRoot) && Directory.Exists(configuredRoot)) targetPath = configuredRoot;
            if (!Directory.Exists(targetPath))
            {
                state.Path = targetPath;
                state.IsUnavailable = true;
                RestoreFolderTabNavigation(page);
                status.Text = "Sekmenin klasörü şu anda bağlı değil; yol saklandı: " + targetPath;
                page.ToolTipText = targetPath + " (bağlı değil)";
                return;
            }
        }
        string displayedPath = currentPath;
        state.Path = targetPath;
        state.IsUnavailable = false;
        page.ToolTipText = targetPath;
        RestoreFolderTabNavigation(page);
        if (wasUnavailable || !String.Equals(displayedPath, targetPath, StringComparison.OrdinalIgnoreCase))
            Navigate(targetPath, null, null, false);
    }

    private void CaptureActiveFolderTab()
    {
        if (activeFolderTab == null) return;
        FolderTabState state = activeFolderTab.Tag as FolderTabState;
        if (state == null) return;
        if (state.IsUnavailable) return;
        state.Path = currentPath;
        state.History = navigationHistory.ToList();
        state.HistoryIndex = navigationHistoryIndex;
        state.IsUnavailable = false;
        activeFolderTab.Text = String.IsNullOrWhiteSpace(state.CustomName) ? GetFolderTabTitle(currentPath) : state.CustomName;
        activeFolderTab.ToolTipText = currentPath;
    }

    private void RestoreFolderTabNavigation(TabPage page)
    {
        FolderTabState state = page == null ? null : page.Tag as FolderTabState;
        if (state == null) return;
        navigationHistory.Clear();
        navigationHistory.AddRange(TakeLastItems(state.History.Where(path => !String.IsNullOrWhiteSpace(path)), 20));
        navigationHistoryIndex = navigationHistory.Count == 0 ? -1 : Math.Max(0, Math.Min(navigationHistory.Count - 1, state.HistoryIndex));
        if (String.IsNullOrWhiteSpace(state.Path))
            state.Path = navigationHistoryIndex >= 0 ? navigationHistory[navigationHistoryIndex] : currentPath;
        state.IsUnavailable = !Directory.Exists(state.Path);
        if (!state.IsUnavailable) currentPath = state.Path;
    }

    private void CreateFolderTab(string path)
    {
        if (!Directory.Exists(path)) path = currentPath;
        CaptureActiveFolderTab();
        var page = CreateFolderTabPage(path, new string[] { path }, 0);
        folderTabs.TabPages.Add(page);
        folderTabs.SelectedTab = page;
    }

    private void RenameFolderTab(TabPage page)
    {
        if (page == null) return;
        FolderTabState state = page.Tag as FolderTabState;
        if (state == null) return;
        using (var dialog = new FolderTabNameDialog(state.CustomName, GetFolderTabTitle(state.Path)))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            if (page == activeFolderTab) CaptureActiveFolderTab();
            state.CustomName = dialog.TabName.Trim();
            state.ShortcutPath = String.IsNullOrWhiteSpace(state.CustomName) ? String.Empty : state.Path;
            page.Text = String.IsNullOrWhiteSpace(state.CustomName) ? GetFolderTabTitle(state.Path) : state.CustomName;
            page.ToolTipText = state.Path;
            SaveFolderTabSessions();
        }
    }

    private void CloseFolderTab(TabPage page)
    {
        if (page == null || folderTabs.TabPages.Count == 0) return;
        if (page == activeFolderTab) CaptureActiveFolderTab();
        if (folderTabs.TabPages.Count == 1)
        {
            return;
        }
        recentlyClosedFolderTabs.Push(page);
        folderTabs.TabPages.Remove(page);
        if (page == activeFolderTab) activeFolderTab = folderTabs.SelectedTab;
        SaveFolderTabSessions();
    }

    private void CloseOtherFolderTabs(TabPage keep)
    {
        if (keep == null) return;
        if (keep != activeFolderTab) folderTabs.SelectedTab = keep;
        CaptureActiveFolderTab();
        foreach (TabPage page in folderTabs.TabPages.Cast<TabPage>().Where(item => item != keep).ToArray())
        {
            recentlyClosedFolderTabs.Push(page);
            folderTabs.TabPages.Remove(page);
        }
        folderTabs.SelectedTab = keep;
        activeFolderTab = keep;
        SaveFolderTabSessions();
    }

    private void ReopenClosedFolderTab()
    {
        while (recentlyClosedFolderTabs.Count > 0)
        {
            TabPage page = recentlyClosedFolderTabs.Pop();
            FolderTabState state = page.Tag as FolderTabState;
            if (state == null || String.IsNullOrWhiteSpace(state.Path)) continue;
            folderTabs.TabPages.Add(page);
            folderTabs.SelectedTab = page;
            return;
        }
    }

    private void SaveFolderTabSessions()
    {
        try
        {
            CaptureActiveFolderTab();
            string directory = Path.GetDirectoryName(folderTabsFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            var document = new XmlDocument();
            XmlElement root = document.CreateElement("FolderTabs");
            root.SetAttribute("selected", Math.Max(0, folderTabs.SelectedIndex).ToString(CultureInfo.InvariantCulture));
            document.AppendChild(root);
            foreach (TabPage page in folderTabs.TabPages)
            {
                FolderTabState state = page.Tag as FolderTabState;
                if (state == null || String.IsNullOrWhiteSpace(state.Path)) continue;
                XmlElement tab = document.CreateElement("Tab");
                tab.SetAttribute("path", StorePortablePath(state.Path));
                tab.SetAttribute("name", state.CustomName ?? String.Empty);
                tab.SetAttribute("shortcutPath", StorePortablePath(state.ShortcutPath ?? String.Empty));
                tab.SetAttribute("historyIndex", state.HistoryIndex.ToString(CultureInfo.InvariantCulture));
                XmlElement historyNode = document.CreateElement("History");
                foreach (string path in TakeLastItems(state.History.Where(item => !String.IsNullOrWhiteSpace(item)), 20))
                {
                    XmlElement item = document.CreateElement("Path");
                    item.SetAttribute("value", StorePortablePath(path));
                    historyNode.AppendChild(item);
                }
                tab.AppendChild(historyNode);
                root.AppendChild(tab);
            }
            using (var writer = XmlWriter.Create(folderTabsFile, new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) }))
                document.Save(writer);
        }
        catch { }
    }

    private void InitializeWorkspacesPage(TabPage page)
    {
        page.BackColor = Color.FromArgb(246, 248, 251);
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = Color.FromArgb(246, 248, 251),
            Padding = new Padding(18)
        };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var header = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(246, 248, 251) };
        var heading = new Label
        {
            Text = "İş Klasörleri",
            AutoSize = true,
            Font = new Font("Segoe UI Semibold", 18F),
            ForeColor = Color.FromArgb(38, 51, 70),
            Location = new Point(2, 2)
        };
        var subtitle = new Label
        {
            Text = "Her dosya için kaldığınız yeri ve ilgili klasörleri hatırlayan çalışma alanları.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5F),
            ForeColor = Color.FromArgb(82, 94, 112),
            Location = new Point(4, 39)
        };
        header.Controls.Add(heading);
        header.Controls.Add(subtitle);

        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 5, 0, 4),
            BackColor = Color.FromArgb(246, 248, 251)
        };
        ConfigureWorkspaceButton(newWorkProfileButton, "+  Yeni iş", 110, true);
        ConfigureWorkspaceButton(openWorkProfileButton, "İşi aç", 90, false);
        ConfigureWorkspaceButton(editWorkProfileButton, "Düzenle", 90, false);
        ConfigureWorkspaceButton(removeWorkProfileButton, "İş bitti", 100, false);
        newWorkProfileButton.Click += (s, e) => CreateWorkProfile();
        openWorkProfileButton.Click += (s, e) => OpenSelectedWorkProfile();
        editWorkProfileButton.Click += (s, e) => EditSelectedWorkProfile();
        removeWorkProfileButton.Click += (s, e) => ToggleSelectedWorkProfileCompleted();
        toolbar.Controls.Add(newWorkProfileButton);
        toolbar.Controls.Add(openWorkProfileButton);
        toolbar.Controls.Add(editWorkProfileButton);
        toolbar.Controls.Add(removeWorkProfileButton);

        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            Size = new Size(1000, 600),
            SplitterDistance = 405,
            BackColor = Color.FromArgb(226, 231, 238),
            BorderStyle = BorderStyle.FixedSingle
        };
        split.Panel1.Padding = new Padding(0, 0, 8, 0);
        split.Panel2.Padding = new Padding(14, 10, 2, 2);

        var listTitle = new Label
        {
            Text = "KAYITLI İŞLER",
            Dock = DockStyle.Top,
            Height = 32,
            Font = new Font("Segoe UI Semibold", 9F),
            ForeColor = Color.FromArgb(82, 94, 112),
            TextAlign = ContentAlignment.MiddleLeft
        };
        workProfileList.Dock = DockStyle.Fill;
        workProfileList.View = View.Details;
        workProfileList.FullRowSelect = true;
        workProfileList.HideSelection = false;
        workProfileList.MultiSelect = false;
        workProfileList.BorderStyle = BorderStyle.None;
        workProfileList.Font = new Font("Segoe UI", 9.5F);
        workProfileActiveFont = new Font(workProfileList.Font, FontStyle.Bold);
        workProfileList.Columns.Clear();
        workProfileList.Columns.Add("İş adı", 145, HorizontalAlignment.Left);
        workProfileList.Columns.Add("Son açılış", 105, HorizontalAlignment.Left);
        workProfileList.Columns.Add("Tamamlanma Durumu", 150, HorizontalAlignment.Left);
        workProfileList.SelectedIndexChanged += (s, e) =>
        {
            if (!refreshingWorkProfileList) RefreshWorkProfileDetails();
        };
        workProfileList.MouseDoubleClick += (s, e) =>
        {
            if (workProfileList.HitTest(e.Location).Item != null) OpenSelectedWorkProfile();
        };
        split.Panel1.Controls.Add(workProfileList);
        split.Panel1.Controls.Add(listTitle);

        var detailLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            BackColor = Color.White,
            Padding = new Padding(22, 18, 18, 16)
        };
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        workProfileTitle.Dock = DockStyle.Fill;
        workProfileTitle.Font = new Font("Segoe UI Semibold", 17F);
        workProfileTitle.ForeColor = Color.FromArgb(38, 51, 70);
        workProfileTitle.TextAlign = ContentAlignment.MiddleLeft;
        workProfileRoot.Dock = DockStyle.Fill;
        workProfileRoot.Font = new Font("Segoe UI", 9.5F);
        workProfileRoot.ForeColor = Color.FromArgb(55, 65, 81);
        workProfileRoot.AutoEllipsis = true;
        workProfileRoot.TextAlign = ContentAlignment.MiddleLeft;
        workProfileNotes.Dock = DockStyle.Fill;
        workProfileNotes.Font = new Font("Segoe UI", 9.5F);
        workProfileNotes.ForeColor = Color.FromArgb(82, 94, 112);
        workProfileNotes.BackColor = Color.White;
        workProfileNotes.BorderStyle = BorderStyle.None;
        workProfileNotes.Multiline = true;
        workProfileNotes.ReadOnly = true;
        workProfileNotes.WordWrap = true;
        workProfileNotes.ScrollBars = ScrollBars.Vertical;
        workProfileLastOpened.Dock = DockStyle.Fill;
        workProfileLastOpened.Font = new Font("Segoe UI", 8.5F);
        workProfileLastOpened.ForeColor = Color.FromArgb(110, 120, 135);
        workProfileLastOpened.TextAlign = ContentAlignment.MiddleLeft;

        var relatedHeader = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
        var relatedTitle = new Label
        {
            Text = "İLGİLİ KLASÖRLER",
            Dock = DockStyle.Left,
            Width = 180,
            Font = new Font("Segoe UI Semibold", 9F),
            ForeColor = Color.FromArgb(82, 94, 112),
            TextAlign = ContentAlignment.MiddleLeft
        };
        var relatedButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.RightToLeft,
            BackColor = Color.White
        };
        ConfigureWorkspaceButton(addWorkProfileFolderButton, "+ Klasör ekle", 108, false);
        ConfigureWorkspaceButton(addWorkProfileFileButton, "+ Dosya ekle", 102, false);
        ConfigureWorkspaceButton(removeWorkProfileFolderButton, "Listeden çıkar", 108, false);
        addWorkProfileFolderButton.Click += (s, e) => AddRelatedWorkFolder();
        addWorkProfileFileButton.Click += (s, e) => AddRelatedWorkFile();
        removeWorkProfileFolderButton.Click += (s, e) => RemoveRelatedWorkFolder();
        relatedButtons.Controls.Add(removeWorkProfileFolderButton);
        relatedButtons.Controls.Add(addWorkProfileFileButton);
        relatedButtons.Controls.Add(addWorkProfileFolderButton);
        relatedHeader.Controls.Add(relatedTitle);
        relatedHeader.Controls.Add(relatedButtons);

        workProfileFolders.Dock = DockStyle.Fill;
        workProfileFolders.View = View.Details;
        workProfileFolders.FullRowSelect = true;
        workProfileFolders.HideSelection = false;
        workProfileFolders.MultiSelect = false;
        workProfileFolders.BorderStyle = BorderStyle.FixedSingle;
        workProfileFolders.Font = new Font("Segoe UI", 9F);
        workProfileFolders.Columns.Clear();
        workProfileFolders.Columns.Add("Öğe", 185, HorizontalAlignment.Left);
        workProfileFolders.Columns.Add("Tür", 70, HorizontalAlignment.Left);
        workProfileFolders.Columns.Add("Konum", 450, HorizontalAlignment.Left);
        workProfileFolders.MouseDoubleClick += (s, e) => OpenSelectedRelatedWorkFolder();
        workProfileFolders.SelectedIndexChanged += (s, e) => SetWorkProfileActionsEnabled(GetSelectedWorkProfile() != null);

        detailLayout.Controls.Add(workProfileTitle, 0, 0);
        detailLayout.Controls.Add(workProfileRoot, 0, 1);
        detailLayout.Controls.Add(workProfileNotes, 0, 2);
        detailLayout.Controls.Add(workProfileLastOpened, 0, 3);
        detailLayout.Controls.Add(relatedHeader, 0, 4);
        detailLayout.Controls.Add(workProfileFolders, 0, 5);
        split.Panel2.Controls.Add(detailLayout);

        layout.Controls.Add(header, 0, 0);
        layout.Controls.Add(toolbar, 0, 1);
        layout.Controls.Add(split, 0, 2);
        page.Controls.Add(layout);
        SetWorkProfileActionsEnabled(false);
        ShowEmptyWorkProfileDetails();
    }

    private void InitializeDevelopmentPage(TabPage page)
    {
        page.BackColor = Color.FromArgb(246, 248, 251);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(16), BackColor = page.BackColor };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var header = new Panel { Dock = DockStyle.Fill, BackColor = page.BackColor };
        header.Controls.Add(new Label
        {
            Text = "Geliştirme Takibi",
            AutoSize = true,
            Font = new Font("Segoe UI Semibold", 17F),
            ForeColor = Color.FromArgb(38, 51, 70),
            Location = new Point(2, 0)
        });
        header.Controls.Add(new Label
        {
            Text = "Talepler ve sürüm notları program klasöründeki ortak dosyada saklanır.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(82, 94, 112),
            Location = new Point(4, 38)
        });

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, Size = new Size(1100, 600), SplitterDistance = 560, BackColor = Color.FromArgb(226, 231, 238), BorderStyle = BorderStyle.FixedSingle };
        split.Panel1.Padding = new Padding(0, 0, 8, 0);
        split.Panel2.Padding = new Padding(8, 0, 0, 0);
        var taskPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(12) };
        taskPanel.Controls.Add(new Label { Text = "EKLENECEKLER", Dock = DockStyle.Top, Height = 30, Font = new Font("Segoe UI Semibold", 10F), ForeColor = Color.FromArgb(55, 65, 81), TextAlign = ContentAlignment.MiddleLeft });
        developmentTaskList.Dock = DockStyle.Fill;
        developmentTaskList.View = View.Details;
        developmentTaskList.FullRowSelect = true;
        developmentTaskList.HideSelection = false;
        developmentTaskList.MultiSelect = false;
        developmentTaskList.BorderStyle = BorderStyle.FixedSingle;
        developmentTaskList.Font = new Font("Segoe UI", 9.5F);
        developmentTaskList.Columns.Clear();
        developmentTaskList.Columns.Add("İstek", 330, HorizontalAlignment.Left);
        developmentTaskList.Columns.Add("Durum", 105, HorizontalAlignment.Left);
        developmentTaskList.Columns.Add("Eklenme", 125, HorizontalAlignment.Left);
        activeDevelopmentTaskFont = new Font(developmentTaskList.Font, FontStyle.Bold);
        developmentTaskList.SelectedIndexChanged += (s, e) =>
        {
            UpdateDevelopmentTaskActions();
            UpdateDevelopmentTaskDetails();
        };
        developmentTaskList.MouseDoubleClick += (s, e) => ToggleSelectedDevelopmentTask();
        var taskInput = new Panel { Dock = DockStyle.Top, Height = 74, Padding = new Padding(0, 4, 0, 6), BackColor = Color.White };
        developmentTaskBox.SetBounds(0, 5, 360, 28);
        developmentTaskBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        developmentTaskBox.Font = new Font("Segoe UI", 9.5F);
        actionTooltips.SetToolTip(developmentTaskBox, "İstenen düzeltme veya eklentiyi yazın.");
        addDevelopmentTaskButton.Text = "İstek ekle";
        addDevelopmentTaskButton.SetBounds(368, 4, 95, 30);
        addDevelopmentTaskButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        addDevelopmentTaskButton.Click += (s, e) => AddDevelopmentTask();
        completeDevelopmentTaskButton.Text = "Eklendi olarak işaretle";
        completeDevelopmentTaskButton.SetBounds(0, 39, 170, 28);
        completeDevelopmentTaskButton.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        completeDevelopmentTaskButton.Enabled = false;
        completeDevelopmentTaskButton.Click += (s, e) => ToggleSelectedDevelopmentTask();
        editDevelopmentTaskButton.Text = "Düzelt";
        editDevelopmentTaskButton.SetBounds(178, 39, 82, 28);
        editDevelopmentTaskButton.Enabled = false;
        editDevelopmentTaskButton.Click += (s, e) => EditSelectedDevelopmentTask();
        deleteDevelopmentTaskButton.Text = "Sil";
        deleteDevelopmentTaskButton.SetBounds(268, 39, 70, 28);
        deleteDevelopmentTaskButton.Enabled = false;
        deleteDevelopmentTaskButton.Click += (s, e) => DeleteSelectedDevelopmentTask();
        developmentChangeScaleSelector.DropDownStyle = ComboBoxStyle.DropDownList;
        developmentChangeScaleSelector.Items.AddRange(new object[] { "Küçük değişiklik", "Büyük değişiklik" });
        developmentChangeScaleSelector.SelectedIndex = 0;
        developmentChangeScaleSelector.SetBounds(346, 39, 155, 28);
        developmentChangeScaleSelector.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        actionTooltips.SetToolTip(developmentChangeScaleSelector, "Tamamlarken sürüm türünü belirler.");
        taskInput.Controls.Add(developmentTaskBox);
        taskInput.Controls.Add(addDevelopmentTaskButton);
        taskInput.Controls.Add(completeDevelopmentTaskButton);
        taskInput.Controls.Add(editDevelopmentTaskButton);
        taskInput.Controls.Add(deleteDevelopmentTaskButton);
        taskInput.Controls.Add(developmentChangeScaleSelector);
        taskPanel.Controls.Add(developmentTaskList);
        taskPanel.Controls.Add(taskInput);
        var taskDetailsPanel = new Panel { Dock = DockStyle.Bottom, Height = 150, Padding = new Padding(0, 12, 0, 0), BackColor = Color.White };
        developmentTaskDetails.Dock = DockStyle.Fill;
        developmentTaskDetails.Multiline = true;
        developmentTaskDetails.ReadOnly = true;
        developmentTaskDetails.WordWrap = true;
        developmentTaskDetails.ScrollBars = ScrollBars.Vertical;
        developmentTaskDetails.BorderStyle = BorderStyle.FixedSingle;
        developmentTaskDetails.BackColor = Color.White;
        developmentTaskDetails.Font = new Font("Segoe UI", 9.5F);
        taskDetailsPanel.Controls.Add(developmentTaskDetails);
        taskDetailsPanel.Controls.Add(new Label { Text = "İSTEK AÇIKLAMASI", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI Semibold", 9F), ForeColor = Color.FromArgb(55, 65, 81), TextAlign = ContentAlignment.MiddleLeft });
        taskPanel.Controls.Add(taskDetailsPanel);
        split.Panel1.Controls.Add(taskPanel);

        var versionPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(12) };
        versionPanel.Controls.Add(new Label { Text = "VERSİYON DEĞİŞİKLİKLERİ", Dock = DockStyle.Top, Height = 30, Font = new Font("Segoe UI Semibold", 10F), ForeColor = Color.FromArgb(55, 65, 81), TextAlign = ContentAlignment.MiddleLeft });
        developmentVersionList.Dock = DockStyle.Fill;
        developmentVersionList.View = View.Details;
        developmentVersionList.FullRowSelect = true;
        developmentVersionList.HideSelection = false;
        developmentVersionList.MultiSelect = false;
        developmentVersionList.BorderStyle = BorderStyle.FixedSingle;
        developmentVersionList.Font = new Font("Segoe UI", 9.5F);
        developmentVersionList.Columns.Clear();
        developmentVersionList.Columns.Add("Versiyon", 82, HorizontalAlignment.Left);
        developmentVersionList.Columns.Add("Değişiklik tarihi", 125, HorizontalAlignment.Left);
        developmentVersionList.Columns.Add("Eklenenler / değişiklikler", 310, HorizontalAlignment.Left);
        developmentVersionList.SelectedIndexChanged += (s, e) => UpdateDevelopmentVersionDetails();
        var versionDetailsPanel = new Panel { Dock = DockStyle.Bottom, Height = 165, Padding = new Padding(0, 12, 0, 0), BackColor = Color.White };
        developmentVersionDetails.Dock = DockStyle.Fill;
        developmentVersionDetails.Multiline = true;
        developmentVersionDetails.ReadOnly = true;
        developmentVersionDetails.WordWrap = true;
        developmentVersionDetails.ScrollBars = ScrollBars.Vertical;
        developmentVersionDetails.BorderStyle = BorderStyle.FixedSingle;
        developmentVersionDetails.BackColor = Color.White;
        developmentVersionDetails.Font = new Font("Segoe UI", 9.5F);
        versionDetailsPanel.Controls.Add(developmentVersionDetails);
        versionDetailsPanel.Controls.Add(new Label { Text = "DEĞİŞİKLİK AÇIKLAMASI", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI Semibold", 9F), ForeColor = Color.FromArgb(55, 65, 81), TextAlign = ContentAlignment.MiddleLeft });
        versionPanel.Controls.Add(versionDetailsPanel);
        versionPanel.Controls.Add(developmentVersionList);
        split.Panel2.Controls.Add(versionPanel);

        layout.Controls.Add(header, 0, 0);
        layout.Controls.Add(split, 0, 1);
        page.Controls.Add(layout);
        developmentTaskBox.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { AddDevelopmentTask(); e.SuppressKeyPress = true; } };
        taskInput.Resize += (s, e) =>
        {
            developmentTaskBox.Width = Math.Max(120, taskInput.ClientSize.Width - addDevelopmentTaskButton.Width - 8);
            addDevelopmentTaskButton.Left = developmentTaskBox.Right + 8;
            developmentChangeScaleSelector.Width = Math.Max(120, taskInput.ClientSize.Width - developmentChangeScaleSelector.Left - 2);
        };
    }

    private void LoadDevelopmentBoard()
    {
        developmentTasks.Clear();
        deletedDevelopmentTaskIds.Clear();
        developmentVersions.Clear();
        MergeDevelopmentBoardFromDisk();
        RefreshDevelopmentBoard();
    }

    private void MergeDevelopmentBoardFromDisk()
    {
        try
        {
            if (!File.Exists(sharedDevelopmentFile)) return;
            var document = new XmlDocument();
            document.Load(sharedDevelopmentFile);
            foreach (XmlNode deletedNode in document.SelectNodes("/DevelopmentBoard/DeletedTasks/Task"))
            {
                Guid deletedId;
                if (Guid.TryParse(deletedNode.InnerText, out deletedId))
                {
                    deletedDevelopmentTaskIds.Add(deletedId);
                    developmentTasks.RemoveAll(item => item.Id == deletedId);
                }
            }
            foreach (XmlNode node in document.SelectNodes("/DevelopmentBoard/Tasks/Task"))
            {
                Guid id;
                if (!Guid.TryParse(node.Attributes["Id"] == null ? null : node.Attributes["Id"].Value, out id)) continue;
                if (deletedDevelopmentTaskIds.Contains(id)) continue;
                var incoming = new DevelopmentTask
                {
                    Id = id,
                    Text = node["Text"] == null ? String.Empty : node["Text"].InnerText,
                    IsCompleted = ParseDevelopmentBoolean(node["IsCompleted"] == null ? null : node["IsCompleted"].InnerText),
                    CreatedAt = ParseDevelopmentDate(node["CreatedAt"] == null ? null : node["CreatedAt"].InnerText),
                    UpdatedAt = ParseDevelopmentDate(node["UpdatedAt"] == null ? null : node["UpdatedAt"].InnerText)
                };
                DevelopmentTask existing = developmentTasks.FirstOrDefault(item => item.Id == incoming.Id);
                if (existing == null) developmentTasks.Add(incoming);
                else if (incoming.UpdatedAt >= existing.UpdatedAt)
                {
                    existing.Text = incoming.Text;
                    existing.IsCompleted = incoming.IsCompleted;
                    existing.CreatedAt = incoming.CreatedAt;
                    existing.UpdatedAt = incoming.UpdatedAt;
                }
            }
            foreach (XmlNode node in document.SelectNodes("/DevelopmentBoard/Versions/Version"))
            {
                Guid id;
                if (!Guid.TryParse(node.Attributes["Id"] == null ? null : node.Attributes["Id"].Value, out id)) continue;
                int number;
                if (!Int32.TryParse(node.Attributes["Number"] == null ? null : node.Attributes["Number"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 0;
                var incoming = new DevelopmentVersion
                {
                    Id = id,
                    Number = number,
                    Major = ParseVersionPart(node.Attributes["Major"] == null ? null : node.Attributes["Major"].Value, 1),
                    Minor = ParseVersionPart(node.Attributes["Minor"] == null ? null : node.Attributes["Minor"].Value, 0),
                    Patch = ParseVersionPart(node.Attributes["Patch"] == null ? null : node.Attributes["Patch"].Value, number),
                    UsedAt = ParseDevelopmentDate(node["UsedAt"] == null ? null : node["UsedAt"].InnerText),
                    UpdatedAt = ParseDevelopmentDate(node["UpdatedAt"] == null ? null : node["UpdatedAt"].InnerText)
                };
                if (incoming.UpdatedAt == DateTime.MinValue) incoming.UpdatedAt = incoming.UsedAt;
                XmlNode changes = node["Changes"];
                if (changes != null)
                    foreach (XmlNode change in changes.SelectNodes("Change"))
                        if (!String.IsNullOrWhiteSpace(change.InnerText)) incoming.Changes.Add(change.InnerText);
                if (incoming.Changes.Count == 1 && String.Equals(incoming.Changes[0], "Yeni kullanım oturumu", StringComparison.OrdinalIgnoreCase)) continue;
                DevelopmentVersion existing = developmentVersions.FirstOrDefault(item => item.Id == incoming.Id);
                if (existing == null) developmentVersions.Add(incoming);
                else if (incoming.UsedAt >= existing.UsedAt)
                {
                    existing.Number = incoming.Number;
                    existing.Major = incoming.Major;
                    existing.Minor = incoming.Minor;
                    existing.Patch = incoming.Patch;
                    existing.UsedAt = incoming.UsedAt;
                    existing.UpdatedAt = incoming.UpdatedAt;
                    existing.Changes.Clear();
                    existing.Changes.AddRange(incoming.Changes);
                }
            }
        }
        catch (Exception ex)
        {
            status.Text = "Geliştirme kayıtları okunamadı: " + ex.Message;
        }
    }

    private void StartDevelopmentSession()
    {
        MergeDevelopmentBoardFromDisk();
        int oldSessionCount = developmentVersions.RemoveAll(version =>
            version.Changes.Count == 1 && String.Equals(version.Changes[0], "Yeni kullanım oturumu", StringComparison.OrdinalIgnoreCase));
        bool addedBaseline = false;
        if (developmentVersions.Count == 0)
        {
            developmentVersions.Add(new DevelopmentVersion
            {
                Id = Guid.NewGuid(), Number = 0, Major = 1, Minor = 0, Patch = 0,
                UsedAt = DateTime.Now, UpdatedAt = DateTime.Now,
                Changes = { "Başlangıç sürümü" }
            });
            addedBaseline = true;
        }
        bool addedHistory = AddBuiltInHistoricalVersionsIfNeeded();
        if (oldSessionCount > 0 || addedBaseline || addedHistory)
            SaveDevelopmentBoard();
        RefreshDevelopmentBoard();
    }

    private bool AddBuiltInHistoricalVersionsIfNeeded()
    {
        if (developmentVersions.Any(version => version.Number > 0)) return false;
        string[][] historicalChanges = new string[][]
        {
            new string[] { "Windows Explorer tarzı ana Gezgin: adres çubuğu, klasör/dosya panelleri, varsayılan uygulamada açma ve geri/ileri/yukarı gezinme." },
            new string[] { "Kısayollarım ve Geçmiş sekmeleri; dosya açılışlarının kaydı, sekmelerden Gezgin'e gitme ve son klasör/navigasyon durumunun korunması." },
            new string[] { "Dosya ayrıntıları: değiştirme tarihi, boyut ve tür; doğal sıralama, filtreler, sütun ve panel genişlikleri, pencere boyutunun ve temaların hatırlanması." },
            new string[] { "Alt klasörlerde * ve ? destekli arama, ilerleme göstergesi, arama sonucu seçimi ve dosya/klasör bilgilerinin listelenmesi." },
            new string[] { "Aynı formda ve ayrı pencerede ön izleme; resim, UDF ve çok satırlı metin dosyaları için görüntüleme." },
            new string[] { "XML e-defter ön izleme eklentisi; Yevmiye, Kebir ve Berat şablonları, şablon klasörü ayarı ve Excel/PDF dışa aktarma." },
            new string[] { "Sağ tık menüleri, varsayılan mail uygulaması/Outlook New ile dosya gönderme, Windows Explorer ile açma ve dosya sürükle-bırak." },
            new string[] { "İş Klasörleri profilleri; dosya ve klasör ekleme, iş tamamlama durumu, farklı bilgisayarlara uyarlanabilir göreli yollar ve ortak iş akışı dosyası." },
            new string[] { "Dosya işlemleri: çoklu seçim, checkbox ve Ctrl+A; kes/kopyala/yapıştır, çakışma seçenekleri, yeniden adlandırma, ZIP, Yeni dosyalar ve yazdırma." },
            new string[] { "Klasör ağacı ve klasik görünüm; çift panelli aktarım, sürükle-bırak ile taşıma, Ctrl ile kopyalama, hedef dosyanın seçilmesi ve taşıma geri alma." },
            new string[] { "OneDrive bulut eşitleme durumları için simgeler ve çok dosyalı klasörlerde arayüzü bekletmeyen arka plan durum sorguları." },
            new string[] { "Tek örnek çalışma, sistem tepsisi, dosya açılışında küçülme ve geri yükleme, açık dosya seçiminin korunması ve minimize simgesi animasyonları." },
            new string[] { "Ortak geliştirme istek listesi, istekleri düzeltme/silme/tamamlama ve küçük/büyük değişikliklere göre sürüm geçmişi." }
        };
        for (int index = 0; index < historicalChanges.Length; index++)
        {
            int sequence = index + 1;
            var version = new DevelopmentVersion
            {
                Id = Guid.NewGuid(),
                Number = sequence,
                Major = 1,
                Minor = sequence,
                Patch = 0,
                UsedAt = DateTime.MinValue,
                UpdatedAt = DateTime.MinValue
            };
            version.Changes.AddRange(historicalChanges[index]);
            developmentVersions.Add(version);
        }
        return true;
    }

    private static DateTime ParseDevelopmentDate(string value)
    {
        DateTime date;
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date) ? date : DateTime.MinValue;
    }

    private static bool ParseDevelopmentBoolean(string value)
    {
        bool result;
        return Boolean.TryParse(value, out result) && result;
    }

    private static int ParseVersionPart(string value, int fallback)
    {
        int result;
        return Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result) && result >= 0 ? result : Math.Max(0, fallback);
    }

    private void SaveDevelopmentBoard()
    {
        string temporary = null;
        try
        {
            MergeDevelopmentBoardFromDisk();
            var document = new XmlDocument();
            XmlElement root = document.CreateElement("DevelopmentBoard");
            root.SetAttribute("FormatVersion", "1");
            document.AppendChild(root);
            XmlElement tasks = document.CreateElement("Tasks");
            root.AppendChild(tasks);
            XmlElement deletedTasks = document.CreateElement("DeletedTasks");
            root.AppendChild(deletedTasks);
            foreach (Guid deletedId in deletedDevelopmentTaskIds)
                AppendDevelopmentText(document, deletedTasks, "Task", deletedId.ToString("D"));
            foreach (DevelopmentTask task in developmentTasks.OrderBy(item => item.CreatedAt))
            {
                XmlElement node = document.CreateElement("Task");
                node.SetAttribute("Id", task.Id.ToString("D"));
                AppendDevelopmentText(document, node, "Text", task.Text);
                AppendDevelopmentText(document, node, "IsCompleted", task.IsCompleted.ToString(CultureInfo.InvariantCulture));
                AppendDevelopmentText(document, node, "CreatedAt", task.CreatedAt.ToString("o", CultureInfo.InvariantCulture));
                AppendDevelopmentText(document, node, "UpdatedAt", task.UpdatedAt.ToString("o", CultureInfo.InvariantCulture));
                tasks.AppendChild(node);
            }
            XmlElement versions = document.CreateElement("Versions");
            root.AppendChild(versions);
            foreach (DevelopmentVersion version in developmentVersions.OrderBy(item => item.Number))
            {
                XmlElement node = document.CreateElement("Version");
                node.SetAttribute("Id", version.Id.ToString("D"));
                node.SetAttribute("Number", version.Number.ToString(CultureInfo.InvariantCulture));
                node.SetAttribute("Major", version.Major.ToString(CultureInfo.InvariantCulture));
                node.SetAttribute("Minor", version.Minor.ToString(CultureInfo.InvariantCulture));
                node.SetAttribute("Patch", version.Patch.ToString(CultureInfo.InvariantCulture));
                AppendDevelopmentText(document, node, "UsedAt", version.UsedAt.ToString("o", CultureInfo.InvariantCulture));
                AppendDevelopmentText(document, node, "UpdatedAt", version.UpdatedAt.ToString("o", CultureInfo.InvariantCulture));
                XmlElement changes = document.CreateElement("Changes");
                foreach (string change in version.Changes) AppendDevelopmentText(document, changes, "Change", change);
                node.AppendChild(changes);
                versions.AppendChild(node);
            }
            string folder = Path.GetDirectoryName(sharedDevelopmentFile);
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            temporary = sharedDevelopmentFile + "." + Guid.NewGuid().ToString("N") + ".tmp";
            document.Save(temporary);
            if (File.Exists(sharedDevelopmentFile))
            {
                try { File.Replace(temporary, sharedDevelopmentFile, null); }
                catch { File.Copy(temporary, sharedDevelopmentFile, true); File.Delete(temporary); }
            }
            else File.Move(temporary, sharedDevelopmentFile);
            status.Text = "Geliştirme kayıtları eşitlenmek üzere kaydedildi.";
        }
        catch (Exception ex)
        {
            status.Text = "Geliştirme kayıtları kaydedilemedi: " + ex.Message;
            MessageBox.Show(this, "Ortak geliştirme dosyasına kayıt yapılamadı:\r\n" + sharedDevelopmentFile + "\r\n\r\n" + ex.Message,
                "Geliştirme kayıtları", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally { if (!String.IsNullOrEmpty(temporary) && File.Exists(temporary)) { try { File.Delete(temporary); } catch { } } }
    }

    private static void AppendDevelopmentText(XmlDocument document, XmlNode parent, string name, string value)
    {
        XmlElement element = document.CreateElement(name);
        element.InnerText = value ?? String.Empty;
        parent.AppendChild(element);
    }

    private void AddDevelopmentTask()
    {
        string text = developmentTaskBox.Text.Trim();
        if (text.Length == 0) { developmentTaskBox.Focus(); return; }
        MergeDevelopmentBoardFromDisk();
        var task = new DevelopmentTask { Id = Guid.NewGuid(), Text = text, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now };
        developmentTasks.Add(task);
        developmentTaskBox.Clear();
        SaveDevelopmentBoard();
        RefreshDevelopmentBoard(task.Id);
    }

    private void ToggleSelectedDevelopmentTask()
    {
        DevelopmentTask task = developmentTaskList.SelectedItems.Count == 0 ? null : developmentTaskList.SelectedItems[0].Tag as DevelopmentTask;
        if (task == null) return;
        MergeDevelopmentBoardFromDisk();
        task = developmentTasks.FirstOrDefault(item => item.Id == task.Id) ?? task;
        task.IsCompleted = !task.IsCompleted;
        task.UpdatedAt = DateTime.Now;
        if (task.IsCompleted) AddDevelopmentVersionForTask(task, developmentChangeScaleSelector.SelectedIndex == 1);
        SaveDevelopmentBoard();
        RefreshDevelopmentBoard(task.Id);
    }

    private void UpdateDevelopmentTaskActions()
    {
        DevelopmentTask task = developmentTaskList.SelectedItems.Count == 0 ? null : developmentTaskList.SelectedItems[0].Tag as DevelopmentTask;
        completeDevelopmentTaskButton.Enabled = task != null;
        editDevelopmentTaskButton.Enabled = task != null;
        deleteDevelopmentTaskButton.Enabled = task != null;
        completeDevelopmentTaskButton.Text = task != null && task.IsCompleted ? "Yeniden etkinleştir" : "Eklendi olarak işaretle";
    }

    private void EditSelectedDevelopmentTask()
    {
        DevelopmentTask selected = developmentTaskList.SelectedItems.Count == 0 ? null : developmentTaskList.SelectedItems[0].Tag as DevelopmentTask;
        if (selected == null) return;
        MergeDevelopmentBoardFromDisk();
        DevelopmentTask task = developmentTasks.FirstOrDefault(item => item.Id == selected.Id);
        if (task == null) return;
        using (var dialog = new DevelopmentTaskDialog(task.Text))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string revisedText = dialog.TaskText.Trim();
            if (revisedText.Length == 0) return;
            task.Text = revisedText;
            task.UpdatedAt = DateTime.Now;
        }
        SaveDevelopmentBoard();
        RefreshDevelopmentBoard(selected.Id);
    }

    private void DeleteSelectedDevelopmentTask()
    {
        DevelopmentTask selected = developmentTaskList.SelectedItems.Count == 0 ? null : developmentTaskList.SelectedItems[0].Tag as DevelopmentTask;
        if (selected == null) return;
        DialogResult answer = MessageBox.Show(this, "Seçili isteği ortak listeden silmek istiyor musunuz?\r\n\r\n" + selected.Text,
            "İsteği sil", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (answer != DialogResult.Yes) return;
        MergeDevelopmentBoardFromDisk();
        deletedDevelopmentTaskIds.Add(selected.Id);
        developmentTasks.RemoveAll(item => item.Id == selected.Id);
        SaveDevelopmentBoard();
        RefreshDevelopmentBoard();
    }

    private void AddDevelopmentVersionForTask(DevelopmentTask task, bool majorChange)
    {
        DevelopmentVersion previous = developmentVersions.OrderByDescending(item => item.Number).FirstOrDefault();
        if (previous == null) previous = new DevelopmentVersion { Number = 0, Major = 1, Minor = 0, Patch = 0 };
        var version = new DevelopmentVersion
        {
            Id = Guid.NewGuid(),
            Number = developmentVersions.Count == 0 ? 1 : developmentVersions.Max(item => item.Number) + 1,
            Major = previous.Major,
            Minor = previous.Minor,
            Patch = previous.Patch,
            UsedAt = DateTime.Now,
            UpdatedAt = DateTime.Now
        };
        if (majorChange) { version.Major++; version.Minor = 0; version.Patch = 0; }
        else { version.Minor++; version.Patch = 0; }
        version.Changes.Add((majorChange ? "Büyük değişiklik: " : "Küçük değişiklik: ") + task.Text);
        developmentVersions.Add(version);
    }

    private void RefreshDevelopmentBoard(Guid? selectTaskId = null, Guid? selectVersionId = null)
    {
        if (developmentTaskList.IsDisposed || developmentVersionList.IsDisposed) return;
        developmentTaskList.BeginUpdate();
        developmentTaskList.Items.Clear();
        foreach (DevelopmentTask task in developmentTasks.OrderBy(item => item.IsCompleted).ThenByDescending(item => item.CreatedAt))
        {
            var row = new ListViewItem(task.Text);
            row.SubItems.Add(task.IsCompleted ? "Tamamlandı" : "Eklenecek");
            row.SubItems.Add(task.CreatedAt == DateTime.MinValue ? String.Empty : task.CreatedAt.ToString("dd.MM.yyyy HH:mm"));
            row.Tag = task;
            if (task.IsCompleted)
            {
                row.ForeColor = Color.Gray;
                foreach (ListViewItem.ListViewSubItem subItem in row.SubItems) subItem.ForeColor = Color.Gray;
            }
            else
            {
                row.ForeColor = Color.DarkRed;
                row.Font = activeDevelopmentTaskFont ?? developmentTaskList.Font;
                foreach (ListViewItem.ListViewSubItem subItem in row.SubItems) subItem.ForeColor = Color.DarkRed;
            }
            developmentTaskList.Items.Add(row);
            if (selectTaskId.HasValue && task.Id == selectTaskId.Value) row.Selected = true;
        }
        developmentTaskList.EndUpdate();
        UpdateDevelopmentTaskDetails();
        developmentVersionList.BeginUpdate();
        developmentVersionList.Items.Clear();
        bool selectedVersion = false;
        foreach (DevelopmentVersion version in developmentVersions.OrderByDescending(item => item.Number))
        {
            var row = new ListViewItem(version.Major.ToString(CultureInfo.InvariantCulture) + "." +
                version.Minor.ToString(CultureInfo.InvariantCulture) + "." + version.Patch.ToString(CultureInfo.InvariantCulture));
            row.SubItems.Add(version.UsedAt == DateTime.MinValue ? "—" : version.UsedAt.ToString("dd.MM.yyyy HH:mm:ss"));
            row.SubItems.Add(version.Changes.Count == 0 ? "—" : String.Join("; ", version.Changes.ToArray()));
            row.Tag = version;
            developmentVersionList.Items.Add(row);
            if ((selectVersionId.HasValue && version.Id == selectVersionId.Value) || (!selectVersionId.HasValue && !selectedVersion))
            {
                row.Selected = true;
                developmentVersionList.FocusedItem = row;
                selectedVersion = true;
            }
        }
        developmentVersionList.EndUpdate();
        UpdateDevelopmentVersionDetails();
        UpdateDevelopmentTaskActions();
    }

    private void UpdateDevelopmentVersionDetails()
    {
        DevelopmentVersion version = developmentVersionList.SelectedItems.Count == 0
            ? null : developmentVersionList.SelectedItems[0].Tag as DevelopmentVersion;
        if (version == null)
        {
            developmentVersionDetails.Clear();
            return;
        }
        string versionName = version.Major.ToString(CultureInfo.InvariantCulture) + "." +
            version.Minor.ToString(CultureInfo.InvariantCulture) + "." + version.Patch.ToString(CultureInfo.InvariantCulture);
        string dateTime = version.UsedAt == DateTime.MinValue
            ? "Tarih/saat: Geçmiş kayıt için kaynaklarda bulunamadı."
            : "Tarih/saat: " + version.UsedAt.ToString("dd.MM.yyyy HH:mm:ss");
        string changes = version.Changes.Count == 0 ? "Bu sürüm için açıklama eklenmemiş." : String.Join(Environment.NewLine, version.Changes.ToArray());
        developmentVersionDetails.Text = "Sürüm: " + versionName + Environment.NewLine + dateTime + Environment.NewLine + Environment.NewLine + changes;
        developmentVersionDetails.SelectionStart = 0;
        developmentVersionDetails.SelectionLength = 0;
        developmentVersionDetails.ScrollToCaret();
    }

    private void UpdateDevelopmentTaskDetails()
    {
        DevelopmentTask task = developmentTaskList.SelectedItems.Count == 0
            ? null : developmentTaskList.SelectedItems[0].Tag as DevelopmentTask;
        if (task == null)
        {
            developmentTaskDetails.Clear();
            return;
        }
        string statusText = task.IsCompleted ? "Tamamlandı" : "Eklenecek";
        string addedAt = task.CreatedAt == DateTime.MinValue
            ? "Eklenme tarihi: bilinmiyor"
            : "Eklenme tarihi: " + task.CreatedAt.ToString("dd.MM.yyyy HH:mm:ss");
        developmentTaskDetails.Text = "Durum: " + statusText + Environment.NewLine + addedAt + Environment.NewLine + Environment.NewLine + task.Text;
        developmentTaskDetails.SelectionStart = 0;
        developmentTaskDetails.SelectionLength = 0;
        developmentTaskDetails.ScrollToCaret();
    }

    private static void ConfigureWorkspaceButton(Button button, string text, int width, bool emphasized)
    {
        button.Text = text;
        button.Width = width;
        button.Height = 34;
        button.Margin = new Padding(0, 0, 8, 0);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = emphasized ? Color.FromArgb(49, 100, 163) : Color.FromArgb(202, 212, 225);
        button.BackColor = emphasized ? Color.FromArgb(49, 100, 163) : Color.White;
        button.ForeColor = emphasized ? Color.White : Color.FromArgb(38, 51, 70);
        button.Font = new Font(emphasized ? "Segoe UI Semibold" : "Segoe UI", 9F);
        button.Cursor = Cursors.Hand;
    }

    private void CreateWorkProfile()
    {
        using (var dialog = new WorkProfileDialog())
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            if (workProfiles.Any(existingProfile => String.Equals(existingProfile.Name, dialog.ProfileName, StringComparison.CurrentCultureIgnoreCase)))
            {
                MessageBox.Show(this, "Bu adla kayıtlı bir iş zaten var.", "İş klasörü oluştur",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var profile = new WorkProfile
            {
                Id = Guid.NewGuid(),
                Name = dialog.ProfileName.Trim(),
                RootPath = Path.GetFullPath(dialog.RootPath),
                LastPath = Path.GetFullPath(dialog.RootPath),
                Notes = dialog.Notes.Trim(),
                CreatedAt = DateTime.Now,
                LastOpenedAt = DateTime.MinValue
            };
            workProfiles.Add(profile);
            SaveWorkProfiles();
            SaveWorkProfileLocations();
            LoadPortableShortcuts();
            SaveShortcuts();
            RefreshWorkProfileList(profile.Id);
            RefreshShortcutList();
        }
    }

    private void EditSelectedWorkProfile()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null) return;
        using (var dialog = new WorkProfileDialog(profile.Name, profile.RootPath, profile.Notes))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            if (workProfiles.Any(item => item.Id != profile.Id &&
                String.Equals(item.Name, dialog.ProfileName, StringComparison.CurrentCultureIgnoreCase)))
            {
                MessageBox.Show(this, "Bu adla kayıtlı başka bir iş zaten var.", "İş klasörünü düzenle",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            profile.Name = dialog.ProfileName.Trim();
            profile.Notes = dialog.Notes.Trim();
            string newRoot = Path.GetFullPath(dialog.RootPath);
            if (!String.Equals(profile.RootPath, newRoot, StringComparison.OrdinalIgnoreCase))
            {
                if (!String.IsNullOrWhiteSpace(profile.RootPath))
                    shortcuts.RemoveAll(item => IsPathWithin(item.FullPath, profile.RootPath) &&
                        !String.Equals(item.FullPath, profile.RootPath, StringComparison.OrdinalIgnoreCase));
                RemoveResolvedPortableItems(profile);
                profile.RootPath = newRoot;
                if (!GetValidWorkProfilePath(profile, profile.LastPath))
                    profile.LastPath = profile.RootPath;
                RestoreResolvedPortableItems(profile);
            }
            SaveWorkProfileLocations();
            LoadPortableShortcuts();
            SaveShortcuts();
            RefreshWorkProfileList(profile.Id);
            RefreshShortcutList();
        }
    }

    private void ToggleSelectedWorkProfileCompleted()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null) return;
        bool completing = !profile.IsCompleted;
        DialogResult result = MessageBox.Show(this,
            completing
                ? "‘" + profile.Name + "’ işini tamamlandı olarak işaretleyeyim mi? Profil ve ekli öğeler korunur."
                : "‘" + profile.Name + "’ işini yeniden etkinleştireyim mi?",
            completing ? "İş bitti" : "İşi yeniden aç", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result != DialogResult.Yes) return;
        profile.IsCompleted = completing;
        if (completing && activeWorkProfile != null && activeWorkProfile.Id == profile.Id) activeWorkProfile = null;
        SaveWorkProfiles();
        RefreshWorkProfileList(profile.Id);
    }

    private void OpenSelectedWorkProfile()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null) return;
        string target = GetValidWorkProfilePath(profile, profile.LastPath) ? profile.LastPath : profile.RootPath;
        if (!Directory.Exists(target))
        {
            MessageBox.Show(this, "İş klasörü bulunamadı:\r\n" + target, "İş klasörünü aç",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        activeWorkProfile = profile;
        profile.LastPath = target;
        profile.LastOpenedAt = DateTime.Now;
        SaveWorkProfileLocations();
        RefreshWorkProfileList(profile.Id);
        tabs.SelectedIndex = 0;
        Navigate(target);
    }

    private void AddRelatedWorkFolder()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null) return;
        using (var dialog = new FolderBrowserDialog())
        {
            dialog.Description = "Bu işe bağlı klasörü seçin.";
            if (Directory.Exists(profile.RootPath)) dialog.SelectedPath = profile.RootPath;
            dialog.ShowNewFolderButton = false;
            if (dialog.ShowDialog(this) != DialogResult.OK || !Directory.Exists(dialog.SelectedPath)) return;
            string path = Path.GetFullPath(dialog.SelectedPath);
            if (String.Equals(path, profile.RootPath, StringComparison.OrdinalIgnoreCase) ||
                profile.RelatedFolders.Any(existing => String.Equals(existing, path, StringComparison.OrdinalIgnoreCase))) return;
            profile.RelatedFolders.Add(path);
            RegisterPortableWorkPath(profile, path);
            SaveWorkProfiles();
            SaveWorkProfileLocations();
            RefreshWorkProfileDetails();
        }
    }

    private void AddRelatedWorkFile()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null) return;
        using (var dialog = new OpenFileDialog())
        {
            dialog.Title = "Bu işe bağlı dosyayı seçin";
            dialog.Multiselect = false;
            dialog.CheckFileExists = true;
            dialog.InitialDirectory = Directory.Exists(profile.RootPath) ? profile.RootPath : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string path = Path.GetFullPath(dialog.FileName);
            if (profile.RelatedFolders.Any(existing => String.Equals(existing, path, StringComparison.OrdinalIgnoreCase))) return;
            profile.RelatedFolders.Add(path);
            RegisterPortableWorkPath(profile, path);
            SaveWorkProfiles();
            SaveWorkProfileLocations();
            RefreshWorkProfileDetails();
        }
    }

    private void RemoveRelatedWorkFolder()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null || workProfileFolders.SelectedItems.Count == 0) return;
        string path = workProfileFolders.SelectedItems[0].Tag as string;
        if (String.IsNullOrWhiteSpace(path)) return;
        string relativePath;
        if (TryGetRelativeWorkPath(profile.RootPath, path, out relativePath))
            profile.PortableRelatedPaths.RemoveAll(item => String.Equals(item, relativePath, StringComparison.OrdinalIgnoreCase));
        profile.RelatedFolders.RemoveAll(existing => String.Equals(existing, path, StringComparison.OrdinalIgnoreCase));
        SaveWorkProfiles();
        SaveWorkProfileLocations();
        RefreshWorkProfileDetails();
    }

    private void OpenSelectedRelatedWorkFolder()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        if (profile == null || workProfileFolders.SelectedItems.Count == 0) return;
        string path = workProfileFolders.SelectedItems[0].Tag as string;
        if (String.IsNullOrWhiteSpace(path)) return;
        if (File.Exists(path))
        {
            string parent = Path.GetDirectoryName(path);
            if (String.IsNullOrEmpty(parent) || !Directory.Exists(parent)) return;
            activeWorkProfile = profile;
            profile.LastPath = parent;
            profile.LastOpenedAt = DateTime.Now;
            SaveWorkProfileLocations();
            tabs.SelectedIndex = 0;
            Navigate(parent, null, path);
            return;
        }
        if (!Directory.Exists(path)) return;
        activeWorkProfile = profile;
        profile.LastPath = path;
        profile.LastOpenedAt = DateTime.Now;
        SaveWorkProfileLocations();
        tabs.SelectedIndex = 0;
        Navigate(path);
    }

    private WorkProfile GetSelectedWorkProfile()
    {
        return workProfileList.SelectedItems.Count == 0 ? null : workProfileList.SelectedItems[0].Tag as WorkProfile;
    }

    private void SetWorkProfileActionsEnabled(bool enabled)
    {
        openWorkProfileButton.Enabled = enabled;
        editWorkProfileButton.Enabled = enabled;
        removeWorkProfileButton.Enabled = enabled;
        WorkProfile profile = GetSelectedWorkProfile();
        bool canEditRelated = enabled && profile != null && !profile.IsCompleted;
        workProfileFolders.Enabled = canEditRelated;
        addWorkProfileFolderButton.Enabled = canEditRelated;
        addWorkProfileFileButton.Enabled = canEditRelated;
        removeWorkProfileFolderButton.Enabled = canEditRelated && workProfileFolders.SelectedItems.Count > 0;
    }

    private void ShowEmptyWorkProfileDetails()
    {
        workProfileFolders.Enabled = false;
        workProfileTitle.Text = workProfiles.Count == 0 ? "Henüz iş profili yok" : "Bir iş profili seçin";
        workProfileRoot.Text = workProfiles.Count == 0
            ? "Yeni İş ile ana klasörünüzü kaydedin; dosyalarınıza kaldığınız yerden dönün."
            : String.Empty;
        workProfileNotes.Text = String.Empty;
        workProfileLastOpened.Text = String.Empty;
        workProfileFolders.Items.Clear();
    }

    private void RefreshWorkProfileDetails()
    {
        WorkProfile profile = GetSelectedWorkProfile();
        SetWorkProfileActionsEnabled(profile != null);
        if (profile == null)
        {
            ShowEmptyWorkProfileDetails();
            return;
        }
        workProfileTitle.Text = (profile.IsCompleted ? "✓  " : String.Empty) + profile.Name;
        removeWorkProfileButton.Text = profile.IsCompleted ? "Yeniden aç" : "İş bitti";
        workProfileRoot.Text = "Ana klasör   " + (String.IsNullOrWhiteSpace(profile.RootPath) ? "Bu bilgisayarda henüz seçilmedi (Düzenle ile belirleyin)." : profile.RootPath) +
            (String.IsNullOrWhiteSpace(profile.LastPath) ? String.Empty : "\r\nKaldığınız yer   " + profile.LastPath);
        workProfileNotes.Text = String.IsNullOrWhiteSpace(profile.Notes) ? "Açıklama eklenmemiş." : profile.Notes;
        workProfileLastOpened.Text = profile.LastOpenedAt == DateTime.MinValue
            ? "Henüz açılmadı"
            : "Son kullanım   " + profile.LastOpenedAt.ToString("dd.MM.yyyy HH:mm");
        workProfileFolders.BeginUpdate();
        workProfileFolders.Items.Clear();
        foreach (string path in profile.RelatedFolders.OrderBy(item => Path.GetFileName(item), NaturalNameComparer.Instance))
        {
            bool isFile = File.Exists(path) || (!Directory.Exists(path) && Path.HasExtension(path));
            string name = isFile ? Path.GetFileName(path) : Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            var row = new ListViewItem(new string[] { name, isFile ? "Dosya" : "Klasör", path }) { Tag = path };
            if (!File.Exists(path) && !Directory.Exists(path)) row.ForeColor = Color.Firebrick;
            workProfileFolders.Items.Add(row);
        }
        workProfileFolders.EndUpdate();
    }

    private void RefreshWorkProfileList(Guid? selectId = null)
    {
        Guid? previousId = selectId;
        if (!previousId.HasValue)
        {
            WorkProfile selected = GetSelectedWorkProfile();
            if (selected != null) previousId = selected.Id;
        }
        refreshingWorkProfileList = true;
        workProfileList.BeginUpdate();
        workProfileList.Items.Clear();
        foreach (WorkProfile profile in workProfiles.OrderBy(item => item.IsCompleted)
            .ThenByDescending(item => item.LastOpenedAt)
            .ThenByDescending(item => item.CreatedAt))
        {
            string lastOpened = profile.LastOpenedAt == DateTime.MinValue
                ? "Henüz"
                : profile.LastOpenedAt.ToString("dd.MM.yy HH:mm");
            var row = new ListViewItem(new string[] { profile.Name, lastOpened, profile.IsCompleted ? "Bitti" : "Devam ediyor" }) { Tag = profile };
            row.ToolTipText = profile.RootPath;
            if (profile.IsCompleted)
            {
                row.ForeColor = Color.Gray;
                row.Font = workProfileList.Font;
            }
            else
            {
                row.ForeColor = Color.DarkRed;
                row.Font = workProfileActiveFont;
            }
            workProfileList.Items.Add(row);
            if (previousId.HasValue && profile.Id == previousId.Value) row.Selected = true;
        }
        workProfileList.EndUpdate();
        refreshingWorkProfileList = false;
        if (previousId.HasValue)
        {
            ListViewItem selectedRow = workProfileList.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => ((WorkProfile)item.Tag).Id == previousId.Value);
            if (selectedRow != null) workProfileList.FocusedItem = selectedRow;
        }
        RefreshWorkProfileDetails();
        removeWorkProfileButton.Text = GetSelectedWorkProfile() != null && GetSelectedWorkProfile().IsCompleted
            ? "Yeniden aç" : "İş bitti";
    }

    private void LoadWorkProfiles()
    {
        try
        {
            string source = workProfilesFile;
            if (!File.Exists(source)) return;
            var document = new XmlDocument();
            document.Load(source);
            foreach (XmlNode node in document.SelectNodes("/WorkProfiles/Profile"))
            {
                Guid id;
                if (!Guid.TryParse(node.Attributes["id"] == null ? null : node.Attributes["id"].Value, out id)) id = Guid.NewGuid();
                DateTime created;
                if (!DateTime.TryParse(node.Attributes["created"] == null ? null : node.Attributes["created"].Value,
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out created)) created = DateTime.Now;
                DateTime lastOpened;
                if (!DateTime.TryParse(node.Attributes["lastOpened"] == null ? null : node.Attributes["lastOpened"].Value,
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out lastOpened)) lastOpened = DateTime.MinValue;
                string name = node.Attributes["name"] == null ? String.Empty : node.Attributes["name"].Value;
                if (String.IsNullOrWhiteSpace(name)) continue;
                var profile = new WorkProfile
                {
                    Id = id,
                    Name = name,
                    Notes = node.Attributes["notes"] == null ? String.Empty : node.Attributes["notes"].Value,
                    CreatedAt = created,
                    LastOpenedAt = lastOpened,
                    IsCompleted = node.Attributes["completed"] != null && node.Attributes["completed"].Value == "true"
                };
                foreach (XmlNode entry in node.SelectNodes("Items/Item"))
                {
                    string relativePath = entry.Attributes["path"] == null ? String.Empty : entry.Attributes["path"].Value;
                    if (!String.IsNullOrWhiteSpace(relativePath)) profile.PortableRelatedPaths.Add(relativePath);
                }
                workProfiles.Add(profile);
            }
            foreach (XmlNode shortcut in document.SelectNodes("/WorkProfiles/PortableShortcuts/Shortcut"))
            {
                Guid profileId;
                if (!Guid.TryParse(shortcut.Attributes["profileId"] == null ? null : shortcut.Attributes["profileId"].Value, out profileId)) continue;
                string relative = shortcut.Attributes["path"] == null ? String.Empty : shortcut.Attributes["path"].Value;
                if (String.IsNullOrWhiteSpace(relative)) continue;
                DateTime added;
                if (!DateTime.TryParse(shortcut.Attributes["added"] == null ? null : shortcut.Attributes["added"].Value,
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out added)) added = DateTime.Now;
                portableShortcutRefs.Add(Tuple.Create(profileId, relative, added));
            }
            LoadWorkProfileLocations();
            foreach (WorkProfile profile in workProfiles) RestoreResolvedPortableItems(profile);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "İş klasörü profilleri okunamadı.\r\n\r\n" + ex.Message,
                "İş Klasörleri", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void SaveWorkProfiles()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(workProfilesFile));
            var document = new XmlDocument();
            XmlElement root = document.CreateElement("WorkProfiles");
            document.AppendChild(root);
            foreach (WorkProfile profile in workProfiles)
            {
                XmlElement item = document.CreateElement("Profile");
                item.SetAttribute("id", profile.Id.ToString("D"));
                item.SetAttribute("name", profile.Name ?? String.Empty);
                item.SetAttribute("notes", profile.Notes ?? String.Empty);
                item.SetAttribute("created", profile.CreatedAt.ToString("o", CultureInfo.InvariantCulture));
                item.SetAttribute("completed", profile.IsCompleted ? "true" : "false");
                XmlElement entries = document.CreateElement("Items");
                foreach (string path in profile.RelatedFolders)
                {
                    string relativePath;
                    if (TryGetRelativeWorkPath(profile.RootPath, path, out relativePath) &&
                        !profile.PortableRelatedPaths.Contains(relativePath, StringComparer.OrdinalIgnoreCase))
                        profile.PortableRelatedPaths.Add(relativePath);
                }
                foreach (string relativePath in profile.PortableRelatedPaths.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    XmlElement entry = document.CreateElement("Item");
                    entry.SetAttribute("path", relativePath);
                    entries.AppendChild(entry);
                }
                item.AppendChild(entries);
                root.AppendChild(item);
            }
            XmlElement portableShortcuts = document.CreateElement("PortableShortcuts");
            var savedShortcutKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var reference in portableShortcutRefs)
            {
                if (!workProfiles.Any(profile => profile.Id == reference.Item1)) continue;
                string key = reference.Item1.ToString("D") + "|" + reference.Item2;
                if (!savedShortcutKeys.Add(key)) continue;
                XmlElement shortcut = document.CreateElement("Shortcut");
                shortcut.SetAttribute("profileId", reference.Item1.ToString("D"));
                shortcut.SetAttribute("path", reference.Item2);
                shortcut.SetAttribute("added", reference.Item3.ToString("o", CultureInfo.InvariantCulture));
                portableShortcuts.AppendChild(shortcut);
            }
            foreach (Entry shortcutEntry in shortcuts)
            {
                foreach (WorkProfile profile in workProfiles.OrderByDescending(item => item.RootPath == null ? 0 : item.RootPath.Length))
                {
                    string relative;
                    if (String.IsNullOrWhiteSpace(profile.RootPath) ||
                        String.Equals(shortcutEntry.FullPath, profile.RootPath, StringComparison.OrdinalIgnoreCase) ||
                        !TryGetRelativeWorkPath(profile.RootPath, shortcutEntry.FullPath, out relative)) continue;
                    string key = profile.Id.ToString("D") + "|" + relative;
                    if (!savedShortcutKeys.Add(key)) break;
                    XmlElement shortcut = document.CreateElement("Shortcut");
                    shortcut.SetAttribute("profileId", profile.Id.ToString("D"));
                    shortcut.SetAttribute("path", relative.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/'));
                    shortcut.SetAttribute("added", shortcutEntry.AddedAt.ToString("o", CultureInfo.InvariantCulture));
                    portableShortcuts.AppendChild(shortcut);
                    if (!portableShortcutRefs.Any(reference => reference.Item1 == profile.Id &&
                        String.Equals(reference.Item2, relative, StringComparison.OrdinalIgnoreCase)))
                        portableShortcutRefs.Add(Tuple.Create(profile.Id, relative, shortcutEntry.AddedAt));
                    break;
                }
            }
            root.AppendChild(portableShortcuts);
            using (var writer = XmlWriter.Create(workProfilesFile, new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) }))
                document.Save(writer);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "İş klasörü profilleri kaydedilemedi.\r\n\r\n" + ex.Message,
                "İş Klasörleri", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void LoadWorkProfileLocations()
    {
        try
        {
            string source = workProfileLocationsFile;
            if (!File.Exists(source)) return;
            var document = new XmlDocument();
            document.Load(source);
            foreach (XmlNode node in document.SelectNodes("/WorkProfileLocations/Profile"))
            {
                Guid id;
                if (!Guid.TryParse(node.Attributes["id"] == null ? null : node.Attributes["id"].Value, out id)) continue;
                WorkProfile profile = workProfiles.FirstOrDefault(item => item.Id == id);
                if (profile == null) continue;
                profile.RootPath = ResolvePortablePath(node.Attributes["root"] == null ? String.Empty : node.Attributes["root"].Value);
                profile.LastPath = ResolvePortablePath(node.Attributes["lastPath"] == null ? profile.RootPath : node.Attributes["lastPath"].Value);
                DateTime lastOpened;
                if (DateTime.TryParse(node.Attributes["lastOpened"] == null ? null : node.Attributes["lastOpened"].Value,
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out lastOpened)) profile.LastOpenedAt = lastOpened;
                foreach (XmlNode entry in node.SelectNodes("Items/Item"))
                {
                    string path = ResolvePortablePath(entry.Attributes["path"] == null ? String.Empty : entry.Attributes["path"].Value);
                    if (!String.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path) &&
                        !IsPathWithin(path, profile.RootPath) && !profile.RelatedFolders.Contains(path, StringComparer.OrdinalIgnoreCase))
                        profile.RelatedFolders.Add(path);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Bu bilgisayardaki iş klasörü konumları okunamadı.\r\n\r\n" + ex.Message,
                "İş Klasörleri", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void SaveWorkProfileLocations()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(workProfileLocationsFile));
            var document = new XmlDocument();
            XmlElement root = document.CreateElement("WorkProfileLocations");
            document.AppendChild(root);
            foreach (WorkProfile profile in workProfiles)
            {
                XmlElement item = document.CreateElement("Profile");
                item.SetAttribute("id", profile.Id.ToString("D"));
                item.SetAttribute("root", StorePortablePath(profile.RootPath ?? String.Empty));
                item.SetAttribute("lastPath", StorePortablePath(profile.LastPath ?? String.Empty));
                item.SetAttribute("lastOpened", profile.LastOpenedAt == DateTime.MinValue
                    ? String.Empty : profile.LastOpenedAt.ToString("o", CultureInfo.InvariantCulture));
                XmlElement entries = document.CreateElement("Items");
                foreach (string path in profile.RelatedFolders)
                {
                    string unusedRelativePath;
                    if (TryGetRelativeWorkPath(profile.RootPath, path, out unusedRelativePath)) continue;
                    XmlElement entry = document.CreateElement("Item");
                    entry.SetAttribute("path", StorePortablePath(path));
                    entries.AppendChild(entry);
                }
                item.AppendChild(entries);
                root.AppendChild(item);
            }
            using (var writer = XmlWriter.Create(workProfileLocationsFile, new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) }))
                document.Save(writer);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Bu bilgisayarın iş klasörü konumları kaydedilemedi.\r\n\r\n" + ex.Message,
                "İş Klasörleri", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static bool TryGetRelativeWorkPath(string rootPath, string itemPath, out string relativePath)
    {
        relativePath = String.Empty;
        if (String.IsNullOrWhiteSpace(rootPath) || String.IsNullOrWhiteSpace(itemPath) || !IsPathWithin(itemPath, rootPath)) return false;
        try
        {
            string root = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string item = Path.GetFullPath(itemPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (String.Equals(root, item, StringComparison.OrdinalIgnoreCase)) relativePath = ".";
            else relativePath = item.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            relativePath = relativePath.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
            return true;
        }
        catch { return false; }
    }

    private static void RegisterPortableWorkPath(WorkProfile profile, string path)
    {
        string relativePath;
        if (TryGetRelativeWorkPath(profile.RootPath, path, out relativePath) &&
            !profile.PortableRelatedPaths.Contains(relativePath, StringComparer.OrdinalIgnoreCase))
            profile.PortableRelatedPaths.Add(relativePath);
    }

    private static string ResolveRelativeWorkPath(string rootPath, string relativePath)
    {
        if (String.IsNullOrWhiteSpace(rootPath) || String.IsNullOrWhiteSpace(relativePath)) return String.Empty;
        if (relativePath == ".") return Path.GetFullPath(rootPath);
        string nativeRelative = relativePath.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
        if (Path.IsPathRooted(nativeRelative)) return String.Empty;
        string resolved = Path.GetFullPath(Path.Combine(rootPath, nativeRelative));
        return IsPathWithin(resolved, rootPath) ? resolved : String.Empty;
    }

    private static void RemoveResolvedPortableItems(WorkProfile profile)
    {
        foreach (string relative in profile.PortableRelatedPaths)
        {
            string resolved = ResolveRelativeWorkPath(profile.RootPath, relative);
            if (!String.IsNullOrEmpty(resolved))
                profile.RelatedFolders.RemoveAll(path => String.Equals(path, resolved, StringComparison.OrdinalIgnoreCase));
        }
    }

    private static void RestoreResolvedPortableItems(WorkProfile profile)
    {
        foreach (string relative in profile.PortableRelatedPaths)
        {
            string resolved = ResolveRelativeWorkPath(profile.RootPath, relative);
            if (!String.IsNullOrEmpty(resolved) && !profile.RelatedFolders.Contains(resolved, StringComparer.OrdinalIgnoreCase))
                profile.RelatedFolders.Add(resolved);
        }
    }

    private bool GetValidWorkProfilePath(WorkProfile profile, string path)
    {
        if (String.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return false;
        if (IsPathWithin(path, profile.RootPath)) return true;
        return profile.RelatedFolders.Any(folder => IsPathWithin(path, folder));
    }

    private static bool IsPathWithin(string path, string root)
    {
        try
        {
            string candidate = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string basePath = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return String.Equals(candidate, basePath, StringComparison.OrdinalIgnoreCase) ||
                candidate.StartsWith(basePath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private void RememberActiveWorkProfilePath(string path)
    {
        if (activeWorkProfile == null || !GetValidWorkProfilePath(activeWorkProfile, path)) return;
        activeWorkProfile.LastPath = path;
        activeWorkProfile.LastOpenedAt = DateTime.Now;
        SaveWorkProfileLocations();
    }

    private void InitializeSettingsPage(TabPage page)
    {
        var content = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(28) };
        var heading = new Label
        {
            Text = "Program ayarları",
            AutoSize = true,
            Location = new Point(28, 20),
            Font = new Font("Segoe UI Semibold", 15F)
        };
        var themeLabel = new Label { Text = "Tema", AutoSize = true, Location = new Point(30, 67) };
        themeSelector.DropDownStyle = ComboBoxStyle.DropDownList;
        themeSelector.SetBounds(30, 91, 190, 30);
        themeSelector.Items.AddRange(new object[]
        {
            "Mavi tema", "Yeşil tema", "Siyah tema", "Mor gece", "Mercan", "Turkuaz"
        });
        themeSelector.SelectedIndex = LoadThemeIndex();
        themeSelector.SelectedIndexChanged += (s, e) =>
        {
            SaveThemeIndex(themeSelector.SelectedIndex);
            ApplyTheme();
        };
        actionTooltips.SetToolTip(themeSelector, "Programın renk temasını seçin");

        var closeGroup = new GroupBox
        {
            Text = "Pencere kapatıldığında",
            Location = new Point(28, 140),
            Size = new Size(490, 100),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        closeToTrayOption.Text = "Tepsi simgesine küçült; çıkış için tepsi menüsünü kullan";
        closeToTrayOption.AutoSize = true;
        closeToTrayOption.Location = new Point(14, 25);
        closeToTrayOption.Checked = closeToTrayOnExit;
        normalCloseOption.Text = "Normal kapat; pencere kapanınca programdan çık";
        normalCloseOption.AutoSize = true;
        normalCloseOption.Location = new Point(14, 55);
        normalCloseOption.Checked = !closeToTrayOnExit;
        closeGroup.Controls.Add(closeToTrayOption);
        closeGroup.Controls.Add(normalCloseOption);

        minimizeAfterOpeningFileOption.Text = "Dosya açıldığında programı simge durumuna küçült";
        minimizeAfterOpeningFileOption.AutoSize = true;
        minimizeAfterOpeningFileOption.Location = new Point(30, 248);
        minimizeAfterOpeningFileOption.Checked = minimizeAfterOpeningFile;
        minimizeAfterOpeningFileOption.CheckedChanged += (s, e) =>
        {
            minimizeAfterOpeningFile = minimizeAfterOpeningFileOption.Checked;
            SavePreferences();
        };
        keepOnTopOption.Text = "Dosya Gezgini penceresini sürekli en üstte tut";
        keepOnTopOption.AutoSize = true;
        keepOnTopOption.Location = new Point(30, 280);
        keepOnTopOption.Checked = keepWindowOnTop;
        rememberFolderSortOption.Text = "Her klasörün dosya sıralama şeklini ayrı ayrı hatırla";
        rememberFolderSortOption.AutoSize = true;
        rememberFolderSortOption.Location = new Point(30, 315);
        rememberFolderSortOption.Checked = rememberFolderSort;
        startMaximizedOption.Text = "Programı açılışta maksimize başlat";
        startMaximizedOption.AutoSize = true;
        startMaximizedOption.Location = new Point(30, 350);
        startMaximizedOption.Checked = startMaximizedOnLaunch;

        var folderViewGroup = new GroupBox
        {
            Text = "Gezgin klasör görünümü",
            Location = new Point(28, 395),
            Size = new Size(490, 76),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        classicFolderViewOption.Location = new Point(16, 30);
        treeFolderViewOption.Location = new Point(100, 30);
        folderViewGroup.Controls.Add(classicFolderViewOption);
        folderViewGroup.Controls.Add(treeFolderViewOption);
        classicFolderViewOption.CheckedChanged += (s, e) =>
        {
            if (!classicFolderViewOption.Checked) return;
            classicFolderView = true;
            ApplyFolderViewMode();
            SavePreferences();
        };
        treeFolderViewOption.CheckedChanged += (s, e) =>
        {
            if (!treeFolderViewOption.Checked) return;
            classicFolderView = false;
            ApplyFolderViewMode();
            SavePreferences();
        };

        var searchOptionsGroup = new GroupBox
        {
            Text = "Arama",
            Location = new Point(28, 475),
            Size = new Size(620, 60),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        searchSubfoldersOption.Text = "Alt klasörlerde arama yap";
        searchSubfoldersOption.AutoSize = true;
        searchSubfoldersOption.Location = new Point(14, 24);
        searchSubfoldersOption.Checked = searchSubfolders;
        actionTooltips.SetToolTip(searchSubfoldersOption, "Açıksa seçili klasörün altındaki klasörler de taranır.");
        searchSubfoldersOption.CheckedChanged += (s, e) =>
        {
            searchSubfolders = searchSubfoldersOption.Checked;
            SavePreferences();
        };
        searchOptionsGroup.Controls.Add(searchSubfoldersOption);

        var historyRetentionGroup = new GroupBox
        {
            Text = "Geçmiş temizleme aralığı",
            Location = new Point(28, 545),
            Size = new Size(620, 76),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        preserveHistoryRetentionOption.Text = String.Empty;
        preserveHistoryRetentionOption.AutoSize = true;
        preserveHistoryRetentionOption.Location = new Point(12, 31);
        preserveHistoryRetentionOption.Checked = preserveHistoryRetention;
        actionTooltips.SetToolTip(preserveHistoryRetentionOption,
            "Seçiliyse seçilen dönem içindeki geçmiş korunur, daha eski kayıtlar temizlenir. Seçili değilse tüm geçmiş temizlenir.");
        var retentionLabel = new Label { Text = "Son", AutoSize = true, Location = new Point(35, 32) };
        historyRetentionCountSelector.DropDownStyle = ComboBoxStyle.DropDownList;
        historyRetentionCountSelector.SetBounds(63, 27, 60, 28);
        for (int i = 1; i <= 30; i++) historyRetentionCountSelector.Items.Add(i.ToString(CultureInfo.InvariantCulture));
        historyRetentionCountSelector.SelectedIndex = Math.Max(0, Math.Min(29, historyRetentionCount - 1));
        historyRetentionUnitSelector.DropDownStyle = ComboBoxStyle.DropDownList;
        historyRetentionUnitSelector.SetBounds(130, 27, 88, 28);
        historyRetentionUnitSelector.Items.AddRange(new object[] { "Gün", "Hafta", "Ay", "Yıl" });
        historyRetentionUnitSelector.SelectedItem = historyRetentionUnit;
        if (historyRetentionUnitSelector.SelectedIndex < 0) historyRetentionUnitSelector.SelectedIndex = 0;
        var retentionSuffix = new Label { Text = "içindeki kayıtları koru", AutoSize = true, Location = new Point(226, 32) };
        var clearHistoryButton = new Button { Text = "Geçmişi Temizle", Location = new Point(430, 25), Size = new Size(168, 32) };
        clearHistoryButton.Click += (s, e) => ClearHistory();
        historyRetentionGroup.Controls.Add(preserveHistoryRetentionOption);
        historyRetentionGroup.Controls.Add(retentionLabel);
        historyRetentionGroup.Controls.Add(historyRetentionCountSelector);
        historyRetentionGroup.Controls.Add(historyRetentionUnitSelector);
        historyRetentionGroup.Controls.Add(retentionSuffix);
        historyRetentionGroup.Controls.Add(clearHistoryButton);
        Action updateRetentionEnabledState = () =>
        {
            historyRetentionCountSelector.Enabled = preserveHistoryRetentionOption.Checked;
            historyRetentionUnitSelector.Enabled = preserveHistoryRetentionOption.Checked;
            retentionLabel.Enabled = preserveHistoryRetentionOption.Checked;
            retentionSuffix.Enabled = preserveHistoryRetentionOption.Checked;
        };
        updateRetentionEnabledState();
        preserveHistoryRetentionOption.CheckedChanged += (s, e) =>
        {
            preserveHistoryRetention = preserveHistoryRetentionOption.Checked;
            updateRetentionEnabledState();
            SavePreferences();
            if (preserveHistoryRetention) PruneHistoryByRetention();
            RefreshHistoryList();
        };
        var templateGroup = new GroupBox
        {
            Text = "E-defter şablon dosyaları",
            Location = new Point(28, 640),
            Size = new Size(620, 106),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        edefterTemplateFolderBox.SetBounds(14, 27, 420, 28);
        edefterTemplateFolderBox.Text = String.IsNullOrWhiteSpace(edefterTemplateFolder)
            ? GetDefaultEdefterTemplateFolder() : edefterTemplateFolder;
        var browseTemplates = new Button { Text = "Gözat..." };
        browseTemplates.SetBounds(442, 26, 78, 30);
        var defaultTemplates = new Button { Text = "Varsayılan", Location = new Point(528, 26), Size = new Size(78, 30) };
        var templateHint = new Label
        {
            Text = "Önce Dosya\\Şablonlar klasörü kullanılır. Başka bilgisayarda farklı konum seçebilirsiniz.",
            Location = new Point(14, 62),
            Size = new Size(590, 30),
            ForeColor = Color.FromArgb(82, 94, 112)
        };
        browseTemplates.Click += (s, e) =>
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "yevmiye.xslt, kebir.xslt ve berat.xslt dosyalarının bulunduğu klasörü seçin.";
                if (Directory.Exists(edefterTemplateFolderBox.Text)) dialog.SelectedPath = edefterTemplateFolderBox.Text;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    edefterTemplateFolderBox.Text = dialog.SelectedPath;
                    edefterTemplateFolder = dialog.SelectedPath;
                    SavePreferences();
                }
            }
        };
        defaultTemplates.Click += (s, e) =>
        {
            edefterTemplateFolder = String.Empty;
            edefterTemplateFolderBox.Text = GetDefaultEdefterTemplateFolder();
            SavePreferences();
        };
        actionTooltips.SetToolTip(edefterTemplateFolderBox, "Şablon klasörünü programın yanındaki Şablonlar klasöründen farklı bir yere yönlendirebilirsiniz.");
        edefterTemplateFolderBox.Leave += (s, e) =>
        {
            string entered = edefterTemplateFolderBox.Text.Trim().Trim('"');
            edefterTemplateFolder = String.Equals(entered, GetDefaultEdefterTemplateFolder(), StringComparison.OrdinalIgnoreCase)
                ? String.Empty : entered;
            SavePreferences();
        };
        templateGroup.Controls.Add(edefterTemplateFolderBox);
        templateGroup.Controls.Add(browseTemplates);
        templateGroup.Controls.Add(defaultTemplates);
        templateGroup.Controls.Add(templateHint);
        historyRetentionCountSelector.SelectedIndexChanged += (s, e) => SaveHistoryRetentionFromControls();
        historyRetentionUnitSelector.SelectedIndexChanged += (s, e) => SaveHistoryRetentionFromControls();

        var oneDriveGroup = new GroupBox
        {
            Text = "Taşınabilir OneDrive verileri",
            Location = new Point(28, 760),
            Size = new Size(620, 158),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        oneDriveRootOption.Text = "OneDrive yollarını taşınabilir biçimde kaydet";
        oneDriveRootOption.AutoSize = true;
        oneDriveRootOption.Location = new Point(14, 24);
        oneDriveRootOption.Checked = oneDriveRootEnabled;
        oneDriveRootBox.SetBounds(14, 52, 450, 27);
        oneDriveRootBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        oneDriveRootBox.Text = oneDriveRoot;
        var browseOneDrive = new Button { Text = "Klasör seç...", Location = new Point(474, 50), Size = new Size(126, 30), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        var saveOneDrive = new Button { Text = "Kaydet", Location = new Point(14, 84), Size = new Size(90, 28) };
        var clearOneDrive = new Button { Text = "Ayarı temizle", Location = new Point(110, 84), Size = new Size(110, 28) };
        var oneDriveStatus = new Label { AutoEllipsis = true, Location = new Point(230, 88), Size = new Size(370, 22), ForeColor = Color.FromArgb(55, 65, 81) };
        var oneDriveHint = new Label
        {
            Text = "Kök klasör yolu bu bilgisayarda saklanır; ortak kayıtlar OneDrive:\\göreli-yol biçimini kullanır.",
            Location = new Point(14, 118),
            Size = new Size(590, 26),
            ForeColor = Color.FromArgb(82, 94, 112)
        };
        Action updateOneDriveStatus = () =>
        {
            string displayRoot = oneDriveRoot.Trim();
            oneDriveStatus.Text = oneDriveRootEnabled && Directory.Exists(displayRoot)
                ? "Etkin · " + displayRoot
                : (Directory.Exists(displayRoot) ? "Klasör seçili · taşınabilir kayıt kapalı" : "OneDrive klasörü seçilmedi");
        };
        updateOneDriveStatus();
        oneDriveRootBox.TextChanged += (s, e) =>
        {
            string entered = oneDriveRootBox.Text.Trim().Trim('"');
            oneDriveStatus.Text = Directory.Exists(entered) ? "Kaydedilmemiş seçim · " + entered : "Geçerli bir klasör seçin.";
        };
        oneDriveRootOption.CheckedChanged += (s, e) =>
        {
            oneDriveStatus.Text = oneDriveRootOption.Checked ? "Etkinleştirme Kaydet'e basınca uygulanır." : "Devre dışı bırakma Kaydet'e basınca uygulanır.";
        };
        browseOneDrive.Click += (s, e) =>
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Bu bilgisayardaki OneDrive kök klasörünü seçin.";
                if (Directory.Exists(oneDriveRootBox.Text)) dialog.SelectedPath = oneDriveRootBox.Text;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    oneDriveRootBox.Text = dialog.SelectedPath;
                    oneDriveRootOption.Checked = true;
                }
            }
        };
        saveOneDrive.Click += (s, e) =>
        {
            string selectedRoot = oneDriveRootBox.Text.Trim().Trim('"');
            if (oneDriveRootOption.Checked && !Directory.Exists(selectedRoot))
            {
                MessageBox.Show(this, "Önce geçerli bir OneDrive klasörü seçin.", "OneDrive ayarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            oneDriveRoot = Directory.Exists(selectedRoot) ? selectedRoot : String.Empty;
            oneDriveRootEnabled = oneDriveRootOption.Checked && !String.IsNullOrWhiteSpace(oneDriveRoot);
            SavePreferences();
            if (oneDriveRootEnabled) SavePortablePathData();
            updateOneDriveStatus();
            MessageBox.Show(this, oneDriveRootEnabled
                ? "OneDrive klasör ayarı kaydedildi. Taşınabilir dosya kayıtları bu köke göre saklanacak."
                : "OneDrive taşınabilir kayıt ayarı kapatıldı.",
                "OneDrive ayarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        clearOneDrive.Click += (s, e) =>
        {
            oneDriveRoot = String.Empty;
            oneDriveRootEnabled = false;
            oneDriveRootBox.Clear();
            oneDriveRootOption.Checked = false;
            SavePreferences();
            updateOneDriveStatus();
        };
        oneDriveGroup.Controls.Add(oneDriveRootOption);
        oneDriveGroup.Controls.Add(oneDriveRootBox);
        oneDriveGroup.Controls.Add(browseOneDrive);
        oneDriveGroup.Controls.Add(saveOneDrive);
        oneDriveGroup.Controls.Add(clearOneDrive);
        oneDriveGroup.Controls.Add(oneDriveStatus);
        oneDriveGroup.Controls.Add(oneDriveHint);

        var folderTabManagementGroup = new GroupBox
        {
            Text = "Gezgin klasör sekmeleri",
            Location = new Point(28, 930),
            Size = new Size(620, 250),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        var folderTabManagementList = new ListView
        {
            Location = new Point(12, 24),
            Size = new Size(590, 164),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            View = View.Details,
            FullRowSelect = true,
            HideSelection = false,
            MultiSelect = false,
            GridLines = true
        };
        folderTabManagementList.Columns.Add("Sekme adı", 150, HorizontalAlignment.Left);
        folderTabManagementList.Columns.Add("Klasör kısayolu", 420, HorizontalAlignment.Left);
        var renameManagedTab = new Button { Text = "Adı değiştir...", Location = new Point(12, 197), Size = new Size(130, 32) };
        var changeManagedTabPath = new Button { Text = "Klasör kısayolunu değiştir...", Location = new Point(150, 197), Size = new Size(230, 32) };
        var deleteManagedTab = new Button { Text = "Sekmeyi sil", Location = new Point(388, 197), Size = new Size(110, 32) };
        renameManagedTab.Enabled = changeManagedTabPath.Enabled = deleteManagedTab.Enabled = false;
        Action refreshFolderTabManagementList = () =>
        {
            CaptureActiveFolderTab();
            folderTabManagementList.BeginUpdate();
            folderTabManagementList.Items.Clear();
            foreach (TabPage tab in folderTabs.TabPages)
            {
                FolderTabState state = tab.Tag as FolderTabState;
                if (state == null) continue;
                string displayedName = String.IsNullOrWhiteSpace(state.CustomName) ? GetFolderTabTitle(state.Path) : state.CustomName;
                string shortcutPath = String.IsNullOrWhiteSpace(state.ShortcutPath) ? state.Path : state.ShortcutPath;
                var row = new ListViewItem(new[] { displayedName, shortcutPath }) { Tag = tab, ToolTipText = shortcutPath };
                folderTabManagementList.Items.Add(row);
            }
            folderTabManagementList.EndUpdate();
            bool hasSelected = folderTabManagementList.SelectedItems.Count > 0;
            renameManagedTab.Enabled = changeManagedTabPath.Enabled = deleteManagedTab.Enabled = hasSelected;
        };
        folderTabManagementList.SelectedIndexChanged += (s, e) =>
        {
            bool hasSelected = folderTabManagementList.SelectedItems.Count > 0;
            renameManagedTab.Enabled = changeManagedTabPath.Enabled = deleteManagedTab.Enabled = hasSelected;
        };
        folderTabManagementList.ItemDrag += (s, e) =>
        {
            ListViewItem dragged = e.Item as ListViewItem;
            if (dragged != null) folderTabManagementList.DoDragDrop(dragged, DragDropEffects.Move);
        };
        folderTabManagementList.AllowDrop = true;
        folderTabManagementList.DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(typeof(ListViewItem)) ? DragDropEffects.Move : DragDropEffects.None;
        folderTabManagementList.DragOver += (s, e) => e.Effect = e.Data.GetDataPresent(typeof(ListViewItem)) ? DragDropEffects.Move : DragDropEffects.None;
        folderTabManagementList.DragDrop += (s, e) =>
        {
            ListViewItem dragged = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
            if (dragged == null || dragged.Tag == null) return;
            Point point = folderTabManagementList.PointToClient(new Point(e.X, e.Y));
            ListViewItem targetRow = folderTabManagementList.GetItemAt(point.X, point.Y);
            TabPage moving = dragged.Tag as TabPage;
            TabPage targetTab = targetRow == null ? null : targetRow.Tag as TabPage;
            int oldIndex = moving == null ? -1 : folderTabs.TabPages.IndexOf(moving);
            int targetIndex = targetTab == null ? folderTabs.TabPages.Count - 1 : folderTabs.TabPages.IndexOf(targetTab);
            if (oldIndex < 0 || targetIndex < 0 || oldIndex == targetIndex) return;
            folderTabs.TabPages.Remove(moving);
            if (oldIndex < targetIndex) targetIndex--;
            folderTabs.TabPages.Insert(Math.Max(0, Math.Min(targetIndex, folderTabs.TabPages.Count)), moving);
            folderTabs.SelectedTab = moving;
            SaveFolderTabSessions();
            refreshFolderTabManagementList();
        };
        renameManagedTab.Click += (s, e) =>
        {
            if (folderTabManagementList.SelectedItems.Count == 0) return;
            TabPage selectedTab = folderTabManagementList.SelectedItems[0].Tag as TabPage;
            RenameFolderTab(selectedTab);
            refreshFolderTabManagementList();
        };
        changeManagedTabPath.Click += (s, e) =>
        {
            if (folderTabManagementList.SelectedItems.Count == 0) return;
            TabPage selectedTab = folderTabManagementList.SelectedItems[0].Tag as TabPage;
            FolderTabState state = selectedTab == null ? null : selectedTab.Tag as FolderTabState;
            if (state == null) return;
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Bu Gezgin sekmesinin açacağı klasörü seçin.";
                if (Directory.Exists(state.Path)) dialog.SelectedPath = state.Path;
                if (dialog.ShowDialog(this) != DialogResult.OK || !Directory.Exists(dialog.SelectedPath)) return;
                state.Path = dialog.SelectedPath;
                state.ShortcutPath = dialog.SelectedPath;
                state.History = new List<string> { state.Path };
                state.HistoryIndex = 0;
                selectedTab.ToolTipText = state.Path;
                selectedTab.Text = String.IsNullOrWhiteSpace(state.CustomName) ? GetFolderTabTitle(state.Path) : state.CustomName;
                if (selectedTab == activeFolderTab) Navigate(state.Path, null, null, false);
                else SaveFolderTabSessions();
                refreshFolderTabManagementList();
            }
        };
        deleteManagedTab.Click += (s, e) =>
        {
            if (folderTabManagementList.SelectedItems.Count == 0) return;
            TabPage selectedTab = folderTabManagementList.SelectedItems[0].Tag as TabPage;
            if (selectedTab == null) return;
            if (MessageBox.Show(this, "Seçili Gezgin sekmesi silinsin mi? Klasör ve içindeki dosyalar etkilenmez.",
                "Sekmeyi sil", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            if (folderTabs.TabPages.Count > 1) CloseFolderTab(selectedTab);
            else
            {
                CaptureActiveFolderTab();
                var replacement = CreateFolderTabPage(currentPath, navigationHistory, navigationHistoryIndex);
                suppressFolderTabEvents = true;
                folderTabs.TabPages.Add(replacement);
                folderTabs.TabPages.Remove(selectedTab);
                folderTabs.SelectedTab = replacement;
                activeFolderTab = replacement;
                suppressFolderTabEvents = false;
                SaveFolderTabSessions();
            }
            refreshFolderTabManagementList();
        };
        folderTabManagementGroup.Controls.Add(folderTabManagementList);
        folderTabManagementGroup.Controls.Add(renameManagedTab);
        folderTabManagementGroup.Controls.Add(changeManagedTabPath);
        folderTabManagementGroup.Controls.Add(deleteManagedTab);
        actionTooltips.SetToolTip(folderTabManagementList, "Sekme adı ve açacağı klasör kısayolu.");
        page.Enter += (s, e) => refreshFolderTabManagementList();

        var updateGroup = new GroupBox
        {
            Text = "İnternet üzerinden program güncellemesi",
            Location = new Point(28, 1190),
            Size = new Size(620, 132),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        automaticUpdateCheckOption.Text = "Program açılışında yeni sürüm kontrol et";
        automaticUpdateCheckOption.AutoSize = true;
        automaticUpdateCheckOption.Location = new Point(14, 24);
        automaticUpdateCheckOption.Checked = automaticUpdateCheck;
        automaticUpdateCheckOption.CheckedChanged += (s, e) =>
        {
            automaticUpdateCheck = automaticUpdateCheckOption.Checked;
            SavePreferences();
        };
        var manifestUrlLabel = new Label { Text = "Güncelleme adresi (latest.xml):", AutoSize = true, Location = new Point(14, 57) };
        updateManifestUrlBox.SetBounds(14, 78, 456, 28);
        updateManifestUrlBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        updateManifestUrlBox.Text = updateManifestUrl;
        updateManifestUrlBox.Leave += (s, e) =>
        {
            updateManifestUrl = updateManifestUrlBox.Text.Trim();
            SavePreferences();
        };
        var checkUpdateButton = new Button { Text = "Şimdi kontrol et", Location = new Point(478, 77), Size = new Size(124, 30), Anchor = AnchorStyles.Top | AnchorStyles.Right };
        checkUpdateButton.Click += (s, e) =>
        {
            updateManifestUrl = updateManifestUrlBox.Text.Trim();
            SavePreferences();
            CheckForUpdates(true);
        };
        var updateHint = new Label
        {
            Text = "HTTPS XML adresi gerekir. Yeni sürüm, indirme bağlantısı ve SHA-256 özeti bu XML'den okunur.",
            AutoEllipsis = true,
            Location = new Point(14, 108),
            Size = new Size(588, 18),
            ForeColor = Color.FromArgb(82, 94, 112)
        };
        updateGroup.Controls.Add(automaticUpdateCheckOption);
        updateGroup.Controls.Add(manifestUrlLabel);
        updateGroup.Controls.Add(updateManifestUrlBox);
        updateGroup.Controls.Add(checkUpdateButton);
        updateGroup.Controls.Add(updateHint);

        closeToTrayOption.CheckedChanged += (s, e) =>
        {
            if (!closeToTrayOption.Checked) return;
            closeToTrayOnExit = true;
            SavePreferences();
        };
        normalCloseOption.CheckedChanged += (s, e) =>
        {
            if (!normalCloseOption.Checked) return;
            closeToTrayOnExit = false;
            SavePreferences();
        };
        keepOnTopOption.CheckedChanged += (s, e) =>
        {
            keepWindowOnTop = keepOnTopOption.Checked;
            TopMost = keepWindowOnTop;
            SavePreferences();
        };
        rememberFolderSortOption.CheckedChanged += (s, e) =>
        {
            rememberFolderSort = rememberFolderSortOption.Checked;
            if (rememberFolderSort)
            {
                if (folderSortPreferences.ContainsKey(currentPath)) ApplyFolderSortPreference(currentPath);
                else RememberCurrentFolderSort();
                RefreshFileRows();
            }
            SavePreferences();
        };
        startMaximizedOption.CheckedChanged += (s, e) =>
        {
            startMaximizedOnLaunch = startMaximizedOption.Checked;
            SavePreferences();
            if (startMaximizedOnLaunch)
            {
                SaveMaximizedState(true);
                WindowState = FormWindowState.Maximized;
            }
            else if (WindowState == FormWindowState.Maximized)
            {
                SaveMaximizedState(false);
                WindowState = FormWindowState.Normal;
            }
        };
        content.Controls.Add(searchOptionsGroup);
        content.Controls.Add(historyRetentionGroup);
        content.Controls.Add(templateGroup);
        content.Controls.Add(oneDriveGroup);
        content.Controls.Add(folderTabManagementGroup);
        content.Controls.Add(updateGroup);

        content.Controls.Add(heading);
        content.Controls.Add(themeLabel);
        content.Controls.Add(themeSelector);
        content.Controls.Add(closeGroup);
        content.Controls.Add(minimizeAfterOpeningFileOption);
        content.Controls.Add(keepOnTopOption);
        content.Controls.Add(rememberFolderSortOption);
        content.Controls.Add(startMaximizedOption);
        content.Controls.Add(folderViewGroup);
        page.Controls.Add(content);
    }

    private SplitContainer CreateLinkedExplorerSplit(TabPage page, string key)
    {
        Control[] existing = page.Controls.Cast<Control>().ToArray();
        page.Controls.Clear();
        var linked = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            Size = new Size(1000, 600),
            SplitterDistance = 400,
            BackColor = Color.FromArgb(226, 231, 238),
            BorderStyle = BorderStyle.FixedSingle
        };
        foreach (Control control in existing) linked.Panel1.Controls.Add(control);
        page.Controls.Add(linked);
        linkedPaneSplits[key] = linked;
        linked.SplitterMoved += (s, e) =>
        {
            if (applyingPaneWidths || !String.Equals(activePaneKey, key, StringComparison.Ordinal)) return;
            SaveActivePaneWidths();
        };
        return linked;
    }

    private void StartDeepSearch()
    {
        string query = deepSearchBox.Text.Trim();
        if (query.Length == 0)
        {
            status.Text = "Arama kutusuna dosya adı veya kalıp yazın.";
            deepSearchBox.Focus();
            return;
        }
        string root = currentPath;
        Entry selectedFolder = folders.SelectedItem as Entry;
        if (selectedFolder != null && !selectedFolder.IsParent && Directory.Exists(selectedFolder.FullPath))
            root = selectedFolder.FullPath;
        Regex[] patterns = BuildSearchPatterns(query);
        if (patterns.Length == 0)
        {
            status.Text = "Arama kalıbı yazın. Örnek: rapor*.pdf veya *.docx;*.xlsx";
            deepSearchBox.Focus();
            return;
        }
        SearchRequest request = new SearchRequest(root, patterns, query, searchSubfolders);
        if (searchWorker.IsBusy)
        {
            pendingSearch = request;
            searchWorker.CancelAsync();
            status.Text = "Önceki arama durduruluyor...";
            searchProgress.Visible = true;
            return;
        }
        BeginDeepSearch(request);
    }

    private void BeginDeepSearch(SearchRequest request)
    {
        deepSearchResults.Items.Clear();
        ResizeDeepSearchColumns();
        searchProgress.Visible = true;
        cancelDeepSearchButton.Enabled = true;
        status.Text = "Aranıyor: " + request.Root;
        searchWorker.RunWorkerAsync(request);
    }

    private void CancelDeepSearch()
    {
        pendingSearch = null;
        if (!searchWorker.IsBusy) return;
        cancelDeepSearchButton.Enabled = false;
        status.Text = "Arama durduruluyor...";
        searchWorker.CancelAsync();
    }

    private void ClearDeepSearchResults()
    {
        pendingSearch = null;
        if (searchWorker.IsBusy) searchWorker.CancelAsync();
        deepSearchResults.Items.Clear();
        ResizeDeepSearchColumns();
        searchProgress.Visible = false;
        cancelDeepSearchButton.Enabled = false;
        status.Text = "Arama sonuçları temizlendi.";
        try { if (File.Exists(deepSearchResultsFile)) File.Delete(deepSearchResultsFile); } catch { }
    }

    private static Regex[] BuildSearchPatterns(string query)
    {
        var patterns = new List<Regex>();
        foreach (string raw in query.Split(';'))
        {
            string pattern = raw.Trim();
            if (pattern.Length == 0) continue;
            if (pattern.IndexOf('*') < 0 && pattern.IndexOf('?') < 0) pattern = "*" + pattern + "*";
            string expression = "^" + Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".") + "$";
            patterns.Add(new Regex(expression, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
        }
        return patterns.ToArray();
    }

    private void SearchWorkerDoWork(object sender, DoWorkEventArgs e)
    {
        var worker = (BackgroundWorker)sender;
        var request = (SearchRequest)e.Argument;
        var found = new List<Entry>();
        var allFiles = new List<string>();
        var pending = new Stack<Tuple<string, int>>();
        var visitedDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        pending.Push(Tuple.Create(request.Root, 0));
        int reportFolderCount = 0;
        while (pending.Count > 0)
        {
            if (worker.CancellationPending) { e.Cancel = true; return; }
            Tuple<string, int> current = pending.Pop();
            string directory = current.Item1;
            try { directory = Path.GetFullPath(directory); }
            catch { continue; }
            if (!visitedDirectories.Add(directory)) continue;
            string[] currentFiles;
            try { currentFiles = Directory.GetFiles(directory); }
            catch { currentFiles = new string[0]; }
            allFiles.AddRange(currentFiles);
            reportFolderCount++;
            if (reportFolderCount == 1 || reportFolderCount % 8 == 0)
                worker.ReportProgress(0, "Klasörler: " + reportFolderCount + " · Dosyalar sayılıyor: " + allFiles.Count + " · Şu an: " + GetSafeFolderName(directory));
            if (request.SearchSubfolders && current.Item2 < 128)
            {
                string[] subdirectories;
                try { subdirectories = Directory.GetDirectories(directory); }
                catch { subdirectories = new string[0]; }
                foreach (string subdirectory in subdirectories)
                {
                    // Follow valid reparse folders too (some cloud providers use them).
                    // The visited set and depth limit prevent cyclic links from looping forever.
                    pending.Push(Tuple.Create(subdirectory, current.Item2 + 1));
                }
            }
        }
        int totalFiles = allFiles.Count;
        worker.ReportProgress(0, "Taranan klasör: " + reportFolderCount + " · Toplam dosya: " + totalFiles + " · Eşleşmeler aranıyor...");
        for (int index = 0; index < allFiles.Count; index++)
        {
            if (worker.CancellationPending) { e.Cancel = true; return; }
            string file = allFiles[index];
            try
            {
                string name = Path.GetFileName(file);
                if (request.Patterns.Any(pattern => pattern.IsMatch(name)))
                {
                    int matchStart, matchLength;
                    FindHighlight(name, request.Query, out matchStart, out matchLength);
                    found.Add(new Entry(name, file, matchStart, matchLength));
                }
            }
            catch (PathTooLongException) { }
            catch (ArgumentException) { }
            if ((index + 1) % 32 == 0 || index == allFiles.Count - 1)
            {
                string folder = GetSafeDirectoryName(file);
                worker.ReportProgress(0, "Dosyalar: " + (index + 1) + " / " + totalFiles + " · Klasörler: " + reportFolderCount +
                    " · Bulunan: " + found.Count + " · Şu an: " + folder);
            }
        }
        e.Result = found.OrderBy(x => x.FullPath, NaturalNameComparer.Instance).ToList();
    }

    private static string GetSafeFolderName(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return String.Empty;
        try { return Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)); }
        catch { return path; }
    }

    private static string GetSafeDirectoryName(string filePath)
    {
        try { return GetSafeFolderName(Path.GetDirectoryName(filePath)); }
        catch { return filePath; }
    }

    private void SearchWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        if (pendingSearch != null)
        {
            SearchRequest next = pendingSearch;
            pendingSearch = null;
            BeginDeepSearch(next);
            return;
        }
        searchProgress.Visible = false;
        cancelDeepSearchButton.Enabled = false;
        if (e.Error != null)
        {
            status.Text = "Arama sırasında hata oluştu.";
            MessageBox.Show(this, e.Error.Message, "Arama", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (e.Cancelled)
        {
            status.Text = "Arama iptal edildi.";
            return;
        }
        var results = e.Result as List<Entry> ?? new List<Entry>();
        deepSearchResults.BeginUpdate();
        var existingPaths = new HashSet<string>(deepSearchResults.Items.Cast<ListViewItem>()
            .Select(item => item.Tag as Entry).Where(entry => entry != null).Select(entry => entry.FullPath), StringComparer.OrdinalIgnoreCase);
        foreach (Entry result in results)
        {
            if (result == null || !existingPaths.Add(result.FullPath)) continue;
            try
            {
                string folder;
                try { folder = Path.GetDirectoryName(result.FullPath); }
                catch (PathTooLongException) { folder = "Yol uzun"; }
                var row = new ListViewItem(new string[] { result.ToString(), folder, FormatFileSize(result.FullPath) });
                row.Tag = result;
                row.ToolTipText = result.FullPath;
                deepSearchResults.Items.Add(row);
            }
            catch (PathTooLongException) { }
            catch (ArgumentException) { }
        }
        deepSearchResults.EndUpdate();
        ApplyDeepSearchSort();
        status.Text = results.Count + " sonuç bulundu.";
    }

    private void RevealSelectedSearchResult()
    {
        Entry selected = deepSearchResults.SelectedItems.Count > 0 ? deepSearchResults.SelectedItems[0].Tag as Entry : null;
        if (selected != null) RevealInExplorer(selected.FullPath, true);
    }

    private void SortDeepSearchResults(int column)
    {
        if (searchSortColumn == column) searchSortAscending = !searchSortAscending;
        else { searchSortColumn = column; searchSortAscending = true; }
        SaveListSortPreference("Arama", searchSortColumn, searchSortAscending);
        ApplyDeepSearchSort();
    }

    private void ApplyDeepSearchSort()
    {
        if (deepSearchResults.Columns.Count < 3) return;
        string[] headers = { "Dosya adı", "Klasör", "Boyut" };
        for (int i = 0; i < headers.Length; i++)
            deepSearchResults.Columns[i].Text = headers[i] + (searchSortColumn == i ? (searchSortAscending ? "  ▲" : "  ▼") : String.Empty);
        ListViewItem selected = deepSearchResults.SelectedItems.Count == 0 ? null : deepSearchResults.SelectedItems[0];
        IEnumerable<ListViewItem> rows = deepSearchResults.Items.Cast<ListViewItem>();
        if (searchSortColumn == 1)
            rows = searchSortAscending ? rows.OrderBy(row => row.SubItems.Count > 1 ? row.SubItems[1].Text : String.Empty, NaturalNameComparer.Instance)
                                       : rows.OrderByDescending(row => row.SubItems.Count > 1 ? row.SubItems[1].Text : String.Empty, NaturalNameComparer.Instance);
        else if (searchSortColumn == 2)
            rows = searchSortAscending ? rows.OrderBy(row => { Entry entry = row.Tag as Entry; return entry == null ? -1L : FileLength(entry.FullPath); })
                                       : rows.OrderByDescending(row => { Entry entry = row.Tag as Entry; return entry == null ? -1L : FileLength(entry.FullPath); });
        else
            rows = searchSortAscending ? rows.OrderBy(row => row.Text, NaturalNameComparer.Instance)
                                       : rows.OrderByDescending(row => row.Text, NaturalNameComparer.Instance);
        ListViewItem[] sorted = rows.ToArray();
        deepSearchResults.BeginUpdate();
        deepSearchResults.Items.Clear();
        deepSearchResults.Items.AddRange(sorted);
        if (selected != null) selected.Selected = true;
        deepSearchResults.EndUpdate();
    }

    private void ResizeDeepSearchColumns()
    {
        if (deepSearchColumnsCustomized) return;
        if (deepSearchResults.Columns.Count < 3 || deepSearchResults.ClientSize.Width <= 0) return;
        int available = Math.Max(500, deepSearchResults.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 4);
        int sizeWidth = 120;
        int nameWidth = Math.Max(360, (int)(available * 0.42));
        int folderWidth = Math.Max(220, available - nameWidth - sizeWidth);
        deepSearchResults.Columns[0].Width = nameWidth;
        deepSearchResults.Columns[1].Width = folderWidth;
        deepSearchResults.Columns[2].Width = sizeWidth;
    }

    private static void FindHighlight(string fileName, string query, out int start, out int length)
    {
        start = 0;
        length = 0;
        foreach (string raw in (query ?? String.Empty).Split(';'))
        {
            foreach (string literal in raw.Trim().Split(new char[] { '*', '?' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int found = fileName.IndexOf(literal, StringComparison.CurrentCultureIgnoreCase);
                if (found < 0) continue;
                start = found;
                length = literal.Length;
                return;
            }
        }
    }

    private void DrawSearchResultBold(object sender, DrawListViewSubItemEventArgs e)
    {
        bool selected = e.Item.Selected;
        Color back = selected ? SystemColors.Highlight : deepSearchResults.BackColor;
        Color fore = selected ? SystemColors.HighlightText : deepSearchResults.ForeColor;
        using (var brush = new SolidBrush(back)) e.Graphics.FillRectangle(brush, e.Bounds);
        if (e.ColumnIndex != 0)
        {
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text ?? String.Empty, deepSearchResults.Font,
                new Rectangle(e.Bounds.Left + 5, e.Bounds.Top, Math.Max(1, e.Bounds.Width - 8), e.Bounds.Height), fore,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
            return;
        }
        Entry entry = e.Item.Tag as Entry;
        if (entry == null) return;
        string text = e.SubItem.Text ?? String.Empty;
        int start = Math.Max(0, Math.Min(text.Length, entry.MatchStart));
        int length = Math.Max(0, Math.Min(entry.MatchLength, text.Length - start));
        string before = text.Substring(0, start);
        string matched = text.Substring(start, length);
        string after = text.Substring(start + length);
        TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter;
        int x = e.Bounds.Left + 5;
        Rectangle clip = e.Bounds;
        GraphicsState state = e.Graphics.Save();
        e.Graphics.SetClip(clip);
        Font bold = deepSearchBoldFont ?? deepSearchResults.Font;
        TextRenderer.DrawText(e.Graphics, before, deepSearchResults.Font, new Rectangle(x, clip.Top, clip.Right - x, clip.Height), fore, flags);
        x += TextRenderer.MeasureText(e.Graphics, before, deepSearchResults.Font, Size.Empty, flags).Width;
        TextRenderer.DrawText(e.Graphics, matched, bold, new Rectangle(x, clip.Top, Math.Max(1, clip.Right - x), clip.Height), fore, flags);
        x += TextRenderer.MeasureText(e.Graphics, matched, bold, Size.Empty, flags).Width;
        TextRenderer.DrawText(e.Graphics, after, deepSearchResults.Font, new Rectangle(x, clip.Top, Math.Max(1, clip.Right - x), clip.Height), fore, flags);
        e.Graphics.Restore(state);
    }

    private void InitializeSearchBox(TextBox box, string cue, Action<string> search, Action enter, ListView list)
    {
        box.Dock = DockStyle.Top;
        box.Height = 30;
        box.Font = new Font("Segoe UI", 10F);
        box.BorderStyle = BorderStyle.FixedSingle;
        box.TabStop = true;
        box.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Tab) e.IsInputKey = true; };
        box.TextChanged += (s, e) => search(box.Text);
        box.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Tab)
            {
                CycleSearchFocus(box, !e.Shift);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                enter();
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                MoveListViewSelection(list, e.KeyCode == Keys.Up ? -1 : 1);
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        };
        var tip = new ToolTip();
        tip.SetToolTip(box, cue + ": oklarla seç, Enter ile Gezgin'e git");
    }

    private void CycleSearchFocus(TextBox current, bool forward)
    {
        TextBox[] boxes = new TextBox[] { folderSearch, fileSearch, shortcutSearch, historySearch, deepSearchBox };
        int index = Array.IndexOf(boxes, current);
        int next = (index + (forward ? 1 : boxes.Length - 1)) % boxes.Length;
        boxes[next].Focus();
    }

    private static bool MatchesFileFilter(string path, string filter)
    {
        if (String.IsNullOrEmpty(filter) || filter == "Tüm dosyalar") return true;
        if (Directory.Exists(path)) return true;
        string extension = Path.GetExtension(path).ToLowerInvariant();
        switch (filter)
        {
            case "Word dosyaları": return new[] { ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf" }.Contains(extension);
            case "Excel dosyaları": return new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".csv" }.Contains(extension);
            case "Acrobat dosyaları": return extension == ".pdf";
            case "UDF dosyaları": return extension == ".udf";
            case "ZIP dosyaları": return extension == ".zip";
            default: return true;
        }
    }

    private bool MatchesFileSearch(string path)
    {
        string query = (fileSearch.Text ?? "").Trim();
        if (query.Length == 0) return true;
        string name = Path.GetFileName(path);
        return name.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
    }

    private void SelectNearestShortcut(string query)
    {
        SelectNearestListViewItem(shortcutList, query);
    }

    private void SelectNearestHistory(string query)
    {
        SelectNearestListViewItem(historyList, query);
    }

    private static void SelectNearestListViewItem(ListView list, string query)
    {
        string text = (query ?? "").Trim();
        if (text.Length == 0)
        {
            list.SelectedItems.Clear();
            return;
        }
        int prefixMatch = -1;
        int containsMatch = -1;
        int bestPrefixExtra = Int32.MaxValue;
        for (int i = 0; i < list.Items.Count; i++)
        {
            Entry item = list.Items[i].Tag as Entry;
            if (item == null) continue;
            string name = Path.GetFileName(item.FullPath);
            if (name.StartsWith(text, StringComparison.CurrentCultureIgnoreCase))
            {
                int extra = name.Length - text.Length;
                if (extra < bestPrefixExtra) { prefixMatch = i; bestPrefixExtra = extra; }
            }
            else if (containsMatch < 0 && name.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0)
                containsMatch = i;
        }
        int match = prefixMatch >= 0 ? prefixMatch : containsMatch;
        list.SelectedItems.Clear();
        if (match >= 0)
        {
            list.Items[match].Selected = true;
            list.FocusedItem = list.Items[match];
            list.EnsureVisible(match);
        }
    }

    private static void MoveListViewSelection(ListView list, int direction)
    {
        if (list.Items.Count == 0) return;
        int index = list.SelectedIndices.Count > 0 ? list.SelectedIndices[0] : -1;
        if (index < 0) index = direction > 0 ? -1 : list.Items.Count;
        index = Math.Max(0, Math.Min(list.Items.Count - 1, index + direction));
        list.SelectedItems.Clear();
        list.Items[index].Selected = true;
        list.FocusedItem = list.Items[index];
        list.EnsureVisible(index);
    }

    private void Navigate(string path, Entry clickedFolder = null, string selectFilePath = null, bool addNavigationHistory = true)
    {
        try
        {
            string full = Path.GetFullPath(Environment.ExpandEnvironmentVariables((path ?? "").Trim().Trim('"')));
            if (!Directory.Exists(full)) throw new DirectoryNotFoundException("Bu konum bulunamadı.");
            FolderTabState activeState = activeFolderTab == null ? null : activeFolderTab.Tag as FolderTabState;
            if (activeState != null) activeState.IsUnavailable = false;
            if (addNavigationHistory) AddNavigationHistory(full);
            currentPath = full;
            showingComputerDrives = false;
            files.LabelEdit = true;
            RememberActiveWorkProfilePath(full);
            SyncFolderTreeToPath(currentPath);
            ApplyFolderSortPreference(currentPath);
            up.Enabled = true;
            pendingFolder = clickedFolder;
            address.Text = currentPath;
            address.ReadOnly = false;
            breadcrumbBar.Visible = true;
            RefreshBreadcrumbs();
            folderSearch.Text = "";
            fileSearch.Text = "";
            folders.BeginUpdate();
            files.BeginUpdate();
            folders.Items.Clear();
            files.Items.Clear();
            fileRows.Clear();
            folders.Items.Add(new Entry("Bu Bilgisayar", ThisComputerPath));
            if (Directory.GetParent(currentPath) != null)
                folders.Items.Add(new Entry("↰  ..", Directory.GetParent(currentPath).FullName, true));
            foreach (string dir in Directory.GetDirectories(currentPath).OrderBy(p => Path.GetFileName(p), NaturalNameComparer.Instance))
                folders.Items.Add(new Entry(Path.GetFileName(dir), dir));
            foreach (string dir in Directory.GetDirectories(currentPath).OrderBy(p => Path.GetFileName(p), NaturalNameComparer.Instance))
            {
                DirectoryInfo info = new DirectoryInfo(dir);
                var folderRow = new ListViewItem(new string[]
                {
                    info.Name, info.LastWriteTime.ToString("dd.MM.yyyy HH:mm"), String.Empty, "Dosya klasörü", String.Empty
                });
                folderRow.Tag = new Entry(info.Name, dir);
                SetFileRowIcon(folderRow, dir);
                UpdateOneDriveStatusCell(folderRow, dir);
                fileRows.Add(folderRow);
            }
            foreach (string file in Directory.GetFiles(currentPath).OrderBy(p => Path.GetFileName(p), NaturalNameComparer.Instance))
            {
                string modified = File.GetLastWriteTime(file).ToString("dd.MM.yyyy HH:mm");
                var listItem = new ListViewItem(new string[]
                {
                Path.GetFileName(file), modified, FormatFileSize(file), GetFileTypeDescription(file), String.Empty
                });
                listItem.Tag = new Entry(Path.GetFileName(file), file);
                SetFileRowIcon(listItem, file);
                UpdateOneDriveStatusCell(listItem, file);
                fileRows.Add(listItem);
            }
            folders.EndUpdate();
            files.EndUpdate();
            RefreshFileRows();
            RefreshVisibleOneDriveStatuses();
            if (!String.IsNullOrEmpty(selectFilePath))
                for (int i = 0; i < files.Items.Count; i++)
                    if (String.Equals(((Entry)files.Items[i].Tag).FullPath, selectFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        files.Items[i].Selected = true;
                        files.FocusedItem = files.Items[i];
                        files.EnsureVisible(i);
                        break;
                    }
            status.Text = Math.Max(0, folders.Items.Count - 1) + " klasör   ·   " + fileRows.Count(row => { Entry entry = row.Tag as Entry; return entry != null && File.Exists(entry.FullPath); }) + " dosya";
            CaptureActiveFolderTab();
            SaveFolderTabSessions();
            ConfigureCurrentFolderWatcher();
        }
        catch (Exception ex)
        {
            folders.EndUpdate();
            files.EndUpdate();
            MessageBox.Show(this, ex.Message, "Konum açılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            address.Text = currentPath;
            address.SelectAll();
        }
    }

    private void ConfigureCurrentFolderWatcher()
    {
        try
        {
            if (currentFolderWatcher != null)
            {
                currentFolderWatcher.EnableRaisingEvents = false;
                currentFolderWatcher.Dispose();
                currentFolderWatcher = null;
            }
            if (showingComputerDrives || String.IsNullOrWhiteSpace(currentPath) || !Directory.Exists(currentPath)) return;
            currentFolderWatcher = new FileSystemWatcher(currentPath)
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.Size
            };
            currentFolderWatcher.Created += (s, e) => ScheduleDirectoryRefresh();
            currentFolderWatcher.Deleted += (s, e) => ScheduleDirectoryRefresh();
            currentFolderWatcher.Changed += (s, e) => ScheduleDirectoryRefresh();
            currentFolderWatcher.Renamed += (s, e) => ScheduleDirectoryRefresh();
            currentFolderWatcher.Error += (s, e) => ScheduleDirectoryRefresh();
            currentFolderWatcher.EnableRaisingEvents = true;
        }
        catch
        {
            if (currentFolderWatcher != null) { currentFolderWatcher.Dispose(); currentFolderWatcher = null; }
        }
    }

    private void ScheduleDirectoryRefresh()
    {
        if (IsDisposed || !IsHandleCreated) return;
        try
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (IsDisposed || refreshingDirectoryFromWatcher) return;
                if (directoryRefreshTimer == null)
                {
                    directoryRefreshTimer = new System.Windows.Forms.Timer();
                    directoryRefreshTimer.Interval = 500;
                    directoryRefreshTimer.Tick += (s, e) =>
                    {
                        directoryRefreshTimer.Stop();
                        RefreshCurrentDirectoryFromWatcher();
                    };
                }
                directoryRefreshTimer.Stop();
                directoryRefreshTimer.Start();
            }));
        }
        catch { }
    }

    private void RefreshCurrentDirectoryFromWatcher()
    {
        if (showingComputerDrives || refreshingDirectoryFromWatcher || String.IsNullOrWhiteSpace(currentPath) || !Directory.Exists(currentPath)) return;
        refreshingDirectoryFromWatcher = true;
        string path = currentPath;
        Entry selected = files.SelectedItems.Count > 0 ? files.SelectedItems[0].Tag as Entry : null;
        string selectedPath = selected == null ? null : selected.FullPath;
        if (String.IsNullOrEmpty(selectedPath) && fileSelectionToRestore != null && fileSelectionToRestore.Length > 0)
            selectedPath = fileSelectionToRestore[0];
        string folderQuery = folderSearch.Text;
        string fileQuery = fileSearch.Text;
        try
        {
            Navigate(path, null, selectedPath, false);
            folderSearch.Text = folderQuery;
            fileSearch.Text = fileQuery;
            RefreshFolderTreeChildrenAtPath(path);
        }
        finally { refreshingDirectoryFromWatcher = false; }
    }

    private void RefreshFolderTreeChildrenAtPath(string path)
    {
        TreeNode node = FindFolderTreeNode(path);
        if (node == null) return;
        bool wasExpanded = node.IsExpanded;
        synchronizingFolderTree = true;
        try
        {
            string selectedPath = folderTree.SelectedNode == null ? null : folderTree.SelectedNode.Tag as string;
            var existing = node.Nodes.Cast<TreeNode>().Where(child => child.Tag is string)
                .ToDictionary(child => (string)child.Tag, StringComparer.OrdinalIgnoreCase);
            string[] directories = Directory.GetDirectories(path)
                .OrderBy(item => Path.GetFileName(item), NaturalNameComparer.Instance).ToArray();
            node.Nodes.Clear();
            foreach (string directory in directories)
            {
                TreeNode child;
                if (!existing.TryGetValue(directory, out child)) child = CreateFolderTreeNode(Path.GetFileName(directory), directory);
                node.Nodes.Add(child);
            }
            if (wasExpanded) node.Expand();
            if (!String.IsNullOrEmpty(selectedPath))
            {
                TreeNode selectedNode = FindFolderTreeNode(selectedPath);
                if (selectedNode != null) folderTree.SelectedNode = selectedNode;
            }
        }
        catch { }
        finally { synchronizingFolderTree = false; }
    }

    private void AddNavigationHistory(string path)
    {
        if (navigationHistoryIndex >= 0 &&
            String.Equals(navigationHistory[navigationHistoryIndex], path, StringComparison.OrdinalIgnoreCase))
        {
            UpdateNavigationButtons();
            return;
        }
        if (navigationHistoryIndex + 1 < navigationHistory.Count)
            navigationHistory.RemoveRange(navigationHistoryIndex + 1, navigationHistory.Count - navigationHistoryIndex - 1);
        navigationHistory.Add(path);
        navigationHistoryIndex = navigationHistory.Count - 1;
        if (navigationHistory.Count > 20)
        {
            int removeCount = navigationHistory.Count - 20;
            navigationHistory.RemoveRange(0, removeCount);
            navigationHistoryIndex -= removeCount;
        }
        UpdateNavigationButtons();
    }

    private void NavigateHistory(int direction)
    {
        int candidate = navigationHistoryIndex + direction;
        while (candidate >= 0 && candidate < navigationHistory.Count)
        {
            string path = navigationHistory[candidate];
            if (Directory.Exists(path))
            {
                navigationHistoryIndex = candidate;
                Navigate(path, null, null, false);
                UpdateNavigationButtons();
                return;
            }
            candidate += direction;
        }
        UpdateNavigationButtons();
    }

    private void UpdateNavigationButtons()
    {
        back.Enabled = navigationHistoryIndex > 0;
        forward.Enabled = navigationHistoryIndex >= 0 && navigationHistoryIndex < navigationHistory.Count - 1;
    }

    private void ShowNavigationHistoryMenu(Control button, ContextMenuStrip menu, int direction)
    {
        menu.Items.Clear();
        if (navigationHistory.Count == 0) return;
        int currentIndex = Math.Max(0, Math.Min(navigationHistory.Count - 1, navigationHistoryIndex));
        int index = currentIndex;
        while (index >= 0 && index < navigationHistory.Count)
        {
            string path = navigationHistory[index];
            if (Directory.Exists(path))
            {
                string displayName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                if (String.IsNullOrEmpty(displayName)) displayName = Path.GetPathRoot(path);
                var item = new ToolStripMenuItem(displayName) { Tag = index };
                int targetIndex = index;
                string targetPath = path;
                item.Click += (s, e) => NavigateToHistoryItem(targetIndex, targetPath);
                menu.Items.Add(item);
            }
            index += direction < 0 ? -1 : 1;
        }
        if (menu.Items.Count > 0) menu.Show(button, new Point(0, button.Height));
    }

    private void NavigateToHistoryItem(int index, string path)
    {
        if (index < 0 || index >= navigationHistory.Count || !Directory.Exists(path)) return;
        navigationHistoryIndex = index;
        Navigate(path, null, null, false);
        UpdateNavigationButtons();
    }

    private void LoadNavigationHistory()
    {
        try
        {
            if (!File.Exists(navigationHistoryFile)) return;
            string[] lines = File.ReadAllLines(navigationHistoryFile);
            int savedIndex;
            if (lines.Length == 0 || !Int32.TryParse(lines[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out savedIndex))
                return;
            navigationHistory.Clear();
            navigationHistory.AddRange(lines.Skip(1).Where(line => !String.IsNullOrWhiteSpace(line))
                .Select(ResolvePortablePath).Where(path => !String.IsNullOrWhiteSpace(path) && Directory.Exists(path)).Take(20));
            navigationHistoryIndex = navigationHistory.Count == 0 ? -1 : Math.Max(0, Math.Min(navigationHistory.Count - 1, savedIndex));
        }
        catch
        {
            navigationHistory.Clear();
            navigationHistoryIndex = -1;
        }
    }

    private void SaveNavigationHistory()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(navigationHistoryFile));
            int removeCount = Math.Max(0, navigationHistory.Count - 20);
            List<string> savedPaths = navigationHistory.Skip(removeCount).ToList();
            int savedIndex = savedPaths.Count == 0 ? -1 : Math.Max(0, Math.Min(savedPaths.Count - 1, navigationHistoryIndex - removeCount));
            File.WriteAllLines(navigationHistoryFile,
                new string[] { savedIndex.ToString(CultureInfo.InvariantCulture) }.Concat(savedPaths.Select(StorePortablePath)).ToArray());
        }
        catch { }
    }

    private void SetAddressEditMode(bool editing)
    {
        addressEditMode = editing;
        address.Text = currentPath;
        address.Visible = editing;
        breadcrumbBar.Visible = !editing;
        if (editing)
        {
            address.BringToFront();
            address.Focus();
            address.SelectAll();
        }
        else
        {
            RefreshBreadcrumbs();
            breadcrumbBar.BringToFront();
        }
    }

    private void NavigateFromAddress()
    {
        string target = address.Text;
        bool valid = Directory.Exists(Environment.ExpandEnvironmentVariables(target.Trim().Trim('"')));
        Navigate(target);
        if (valid) SetAddressEditMode(false);
    }

    private void RefreshBreadcrumbs()
    {
        breadcrumbBar.SuspendLayout();
        foreach (Control child in breadcrumbBar.Controls.Cast<Control>().ToArray())
        {
            breadcrumbBar.Controls.Remove(child);
            child.Dispose();
        }
        string fullPath = currentPath;
        string root = Path.GetPathRoot(fullPath);
        var segments = new List<KeyValuePair<string, string>>();
        if (!String.IsNullOrEmpty(root))
        {
            segments.Add(new KeyValuePair<string, string>(root, root));
            string accumulated = root;
            string remainder = fullPath.Substring(root.Length);
            foreach (string part in remainder.Split(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries))
            {
                accumulated = Path.Combine(accumulated, part);
                segments.Add(new KeyValuePair<string, string>(part, accumulated));
            }
        }
        else segments.Add(new KeyValuePair<string, string>(fullPath, fullPath));

        // Show the most recent path components first when space is limited.
        // This keeps the current folder visible without a scrollbar.
        int available = Math.Max(40, breadcrumbBar.ClientSize.Width - 12);
        int firstVisible = segments.Count - 1;
        int used = 0;
        for (int i = segments.Count - 1; i >= 0; i--)
        {
            int segmentWidth = TextRenderer.MeasureText(segments[i].Key, breadcrumbBar.Font).Width + 12;
            if (i < segments.Count - 1) segmentWidth += 18;
            if (used + segmentWidth > available && i < segments.Count - 1) break;
            used += segmentWidth;
            firstVisible = i;
        }
        if (firstVisible > 0)
        {
            var ellipsis = new LinkLabel
            {
                Text = "…",
                AutoSize = true,
                LinkBehavior = LinkBehavior.HoverUnderline,
                LinkColor = ThemeColors.ForIndex(themeSelector.SelectedIndex).Text,
                Margin = new Padding(3, 3, 3, 0),
                Tag = segments[firstVisible - 1].Value,
                TextAlign = ContentAlignment.MiddleCenter
            };
            actionTooltips.SetToolTip(ellipsis, "Önceki klasörler: " + segments[firstVisible - 1].Value);
            ellipsis.LinkClicked += (s, e) => Navigate((string)((Control)s).Tag);
            breadcrumbBar.Controls.Add(ellipsis);
            var ellipsisSeparator = new Label { Text = "›", AutoSize = true, ForeColor = Color.FromArgb(130, 140, 150), Margin = new Padding(0, 3, 0, 0) };
            breadcrumbBar.Controls.Add(ellipsisSeparator);
        }
        for (int i = firstVisible; i < segments.Count; i++)
        {
            if (i > firstVisible)
            {
                breadcrumbBar.Controls.Add(new Label { Text = "›", AutoSize = true, ForeColor = Color.FromArgb(130, 140, 150), Margin = new Padding(0, 3, 0, 0) });
            }
            AddBreadcrumbSegment(segments[i].Key, segments[i].Value);
        }
        breadcrumbBar.ResumeLayout(true);
    }

    private Control AddBreadcrumbSegment(string label, string targetPath)
    {
        ThemeColors theme = ThemeColors.ForIndex(themeSelector.SelectedIndex);
        var segment = new LinkLabel
        {
            Text = label,
            AutoSize = true,
            LinkBehavior = LinkBehavior.HoverUnderline,
            LinkColor = theme.Text,
            ActiveLinkColor = theme.Muted,
            VisitedLinkColor = theme.Text,
            Margin = new Padding(3, 3, 3, 0),
            Padding = Padding.Empty,
            BackColor = Color.White,
            Tag = targetPath,
            TextAlign = ContentAlignment.MiddleCenter
        };
        actionTooltips.SetToolTip(segment, "Bu klasöre git: " + targetPath);
        bool leftClick = false;
        segment.MouseDown += (s, e) => leftClick = e.Button == MouseButtons.Left;
        segment.LinkClicked += (s, e) =>
        {
            if (leftClick) Navigate(targetPath);
            leftClick = false;
        };
        var menu = new ContextMenuStrip();
        var copyPath = new ToolStripMenuItem("Kopyala");
        copyPath.Click += (s, e) => CopyFolderPath(targetPath);
        menu.Items.Add(copyPath);
        segment.ContextMenuStrip = menu;
        breadcrumbBar.Controls.Add(segment);
        return segment;
    }

    private void CopyFolderPath(string targetPath)
    {
        try
        {
            if (String.IsNullOrWhiteSpace(targetPath) || !Directory.Exists(targetPath))
                throw new DirectoryNotFoundException("Klasör bulunamadı: " + targetPath);
            Clipboard.SetText(targetPath);
            if (!String.Equals(Clipboard.GetText(), targetPath, StringComparison.Ordinal))
                throw new IOException("Klasör yolu Windows panosuna aktarılamadı.");
            status.Text = "Klasörün tam yolu panoya kopyalandı. Metin kutusuna Ctrl+V ile yapıştırabilirsiniz.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Klasörün tam yolu panoya kopyalanamadı.\r\n\r\n" + ex.Message,
                "Kopyala", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void SelectNearestFolder(string query)
    {
        string text = (query ?? "").Trim();
        if (text.Length == 0)
        {
            folders.ClearSelected();
            return;
        }
        int prefixMatch = -1;
        int containsMatch = -1;
        int bestPrefixExtra = Int32.MaxValue;
        for (int i = 0; i < folders.Items.Count; i++)
        {
            Entry item = folders.Items[i] as Entry;
            if (item == null || item.IsParent) continue;
            string name = Path.GetFileName(item.FullPath);
            if (name.StartsWith(text, StringComparison.CurrentCultureIgnoreCase))
            {
                int extra = name.Length - text.Length;
                if (extra < bestPrefixExtra)
                {
                    prefixMatch = i;
                    bestPrefixExtra = extra;
                }
            }
            else if (containsMatch < 0 && name.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0)
                containsMatch = i;
        }
        int match = prefixMatch >= 0 ? prefixMatch : containsMatch;
        if (match >= 0)
        {
            folders.ClearSelected();
            folders.SetSelected(match, true);
            folders.TopIndex = match;
            Entry selected = folders.Items[match] as Entry;
            if (selected != null) SelectFolderTreePath(selected.FullPath);
        }
        else folders.ClearSelected();
    }

    private void SelectNearestFile(string query)
    {
        string text = (query ?? "").Trim();
        if (text.Length == 0)
        {
            ClearFileSelection();
            return;
        }
        int prefixMatch = -1;
        int containsMatch = -1;
        int bestPrefixExtra = Int32.MaxValue;
        for (int i = 0; i < files.Items.Count; i++)
        {
            Entry item = files.Items[i].Tag as Entry;
            if (item == null) continue;
            string name = Path.GetFileName(item.FullPath);
            if (name.StartsWith(text, StringComparison.CurrentCultureIgnoreCase))
            {
                int extra = name.Length - text.Length;
                if (extra < bestPrefixExtra)
                {
                    prefixMatch = i;
                    bestPrefixExtra = extra;
                }
            }
            else if (containsMatch < 0 && name.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0)
                containsMatch = i;
        }
        int match = prefixMatch >= 0 ? prefixMatch : containsMatch;
        if (match >= 0)
        {
            ClearFileSelection();
            files.Items[match].Selected = true;
            files.FocusedItem = files.Items[match];
            files.EnsureVisible(match);
        }
        else ClearFileSelection();
    }

    private void MoveFileSelection(int direction)
    {
        if (files.Items.Count == 0) return;
        int index = files.SelectedIndices.Count > 0 ? files.SelectedIndices[0] : -1;
        if (index < 0) index = direction > 0 ? -1 : files.Items.Count;
        index = Math.Max(0, Math.Min(files.Items.Count - 1, index + direction));
        ClearFileSelection();
        files.Items[index].Selected = true;
        files.FocusedItem = files.Items[index];
        files.EnsureVisible(index);
    }

    private void ClearFileSelection()
    {
        foreach (ListViewItem selected in files.SelectedItems) selected.Selected = false;
    }

    private void ApplyFolderViewMode()
    {
        folders.Visible = classicFolderView;
        folderTree.Visible = !classicFolderView;
    }

    private void PopulateFolderTreeRoots()
    {
        folderTree.BeginUpdate();
        folderTree.Nodes.Clear();
        try
        {
            var thisComputer = new TreeNode("Bu Bilgisayar") { Tag = "shell:thispc" };
            thisComputer.Nodes.Add(new TreeNode(String.Empty));
            folderTree.Nodes.Add(thisComputer);
        }
        catch { }
        folderTree.EndUpdate();
    }

    private TreeNode CreateFolderTreeNode(string label, string path)
    {
        var node = new TreeNode(label) { Tag = path };
        if (Directory.Exists(path)) node.Nodes.Add(new TreeNode(String.Empty));
        return node;
    }

    private void LoadFolderTreeChildren(TreeNode node)
    {
        if (node == null || node.Nodes.Count != 1 || node.Nodes[0].Tag != null) return;
        if (String.Equals(node.Tag as string, "shell:thispc", StringComparison.OrdinalIgnoreCase))
        {
            node.Nodes.Clear();
            try
            {
                foreach (DriveInfo drive in DriveInfo.GetDrives().OrderBy(item => item.Name, NaturalNameComparer.Instance))
                {
                    string drivePath = drive.Name;
                    string label = drivePath;
                    try
                    {
                        if (drive.IsReady && !String.IsNullOrWhiteSpace(drive.VolumeLabel))
                            label = drive.VolumeLabel + " (" + drivePath.TrimEnd(Path.DirectorySeparatorChar) + ")";
                    }
                    catch { }
                    node.Nodes.Add(CreateFolderTreeNode(label, drivePath));
                }
            }
            catch { }
            return;
        }
        string path = node.Tag as string;
        node.Nodes.Clear();
        if (String.IsNullOrWhiteSpace(path)) return;
        try
        {
            foreach (string directory in Directory.GetDirectories(path)
                .OrderBy(item => Path.GetFileName(item), NaturalNameComparer.Instance))
                node.Nodes.Add(CreateFolderTreeNode(Path.GetFileName(directory), directory));
        }
        catch { }
    }

    private void SyncFolderTreeToPath(string path)
    {
        if (folderTree.Nodes.Count == 0) PopulateFolderTreeRoots();
        string root = Path.GetPathRoot(path);
        TreeNode computer = folderTree.Nodes.Cast<TreeNode>()
            .FirstOrDefault(node => String.Equals(node.Tag as string, "shell:thispc", StringComparison.OrdinalIgnoreCase));
        TreeNode current = computer;
        if (computer != null)
        {
            computer.Expand();
            LoadFolderTreeChildren(computer);
            current = computer.Nodes.Cast<TreeNode>()
                .FirstOrDefault(node => String.Equals(node.Tag as string, root, StringComparison.OrdinalIgnoreCase));
        }
        if (current == null)
        {
            current = CreateFolderTreeNode(root, root);
            if (computer != null) computer.Nodes.Add(current);
            else folderTree.Nodes.Add(current);
        }

        synchronizingFolderTree = true;
        try
        {
            string remainder = path.Substring(root.Length);
            string[] parts = remainder.Split(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                StringSplitOptions.RemoveEmptyEntries);
            current.Expand();
            LoadFolderTreeChildren(current);
            string accumulated = root;
            foreach (string part in parts)
            {
                accumulated = Path.Combine(accumulated, part);
                TreeNode parent = current;
                current = parent.Nodes.Cast<TreeNode>()
                    .FirstOrDefault(node => String.Equals(node.Tag as string, accumulated, StringComparison.OrdinalIgnoreCase));
                if (current == null)
                {
                    current = CreateFolderTreeNode(part, accumulated);
                    parent.Nodes.Add(current);
                }
                if (!String.Equals(accumulated, path, StringComparison.OrdinalIgnoreCase))
                {
                    current.Expand();
                    LoadFolderTreeChildren(current);
                }
            }
            current.Expand();
            LoadFolderTreeChildren(current);
            folderTree.SelectedNode = current;
            current.EnsureVisible();
        }
        catch { }
        finally { synchronizingFolderTree = false; }
    }

    private void SelectFolderTreePath(string path)
    {
        TreeNode node = FindFolderTreeNode(path);
        if (node == null) return;
        synchronizingFolderTree = true;
        try
        {
            folderTree.SelectedNode = node;
            node.EnsureVisible();
        }
        finally { synchronizingFolderTree = false; }
    }

    private TreeNode FindFolderTreeNode(string path)
    {
        foreach (TreeNode root in folderTree.Nodes)
        {
            TreeNode found = FindFolderTreeNode(root, path);
            if (found != null) return found;
        }
        return null;
    }

    private TreeNode FindFolderTreeNode(TreeNode node, string path)
    {
        if (String.Equals(node.Tag as string, path, StringComparison.OrdinalIgnoreCase)) return node;
        foreach (TreeNode child in node.Nodes)
        {
            TreeNode found = FindFolderTreeNode(child, path);
            if (found != null) return found;
        }
        return null;
    }

    private void MoveFolderSelection(int direction)
    {
        if (folders.Items.Count == 0) return;
        int index = folders.SelectedIndex;
        if (index < 0) index = direction > 0 ? -1 : folders.Items.Count;
        index = Math.Max(0, Math.Min(folders.Items.Count - 1, index + direction));
        folders.ClearSelected();
        folders.SetSelected(index, true);
        folders.TopIndex = index;
        Entry selected = folders.Items[index] as Entry;
        if (selected != null) SelectFolderTreePath(selected.FullPath);
    }

    private void OpenFileAt(Point location)
    {
        ListViewItem hit = files.HitTest(location).Item;
        if (hit == null) return;
        Entry item = hit.Tag as Entry;
        if (item == null) return;
        SelectExplorerFileRow(item.FullPath);
        if (Directory.Exists(item.FullPath)) Navigate(item.FullPath, new Entry(Path.GetFileName(item.FullPath), item.FullPath));
        else if (File.Exists(item.FullPath)) OpenFile(item.FullPath);
    }

    private void SelectExplorerFileRow(string path)
    {
        if (tabs.SelectedIndex != 0 || String.IsNullOrWhiteSpace(path)) return;
        foreach (ListViewItem row in files.Items)
        {
            Entry entry = row.Tag as Entry;
            if (entry == null || !String.Equals(entry.FullPath, path, StringComparison.OrdinalIgnoreCase)) continue;
            row.Selected = true;
            row.Focused = true;
            files.FocusedItem = row;
            row.EnsureVisible();
            return;
        }
    }

    private void OpenFile(string path)
    {
        try
        {
            SelectExplorerFileRow(path);
            if (TryOpenEdefterXml(path)) return;
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            AddToHistory(path);
            if (minimizeAfterOpeningFile) MinimizeToFloatingIcon(path);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Dosya açılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private bool TryOpenEdefterXml(string path)
    {
        try
        {
            if (!String.Equals(Path.GetExtension(path), ".xml", StringComparison.OrdinalIgnoreCase)) return false;
        }
        catch (PathTooLongException) { return false; }
        XmlDocument document;
        try
        {
            document = EDefterViewerForm.LoadXmlDocument(path);
        }
        catch { return false; }
        XmlElement root = document.DocumentElement;
        if (root == null || root.NamespaceURI != "http://www.edefter.gov.tr" ||
            (root.LocalName != "berat" && root.LocalName != "defter")) return false;
        try
        {
            string stylesheet = EDefterViewerForm.ChooseStylesheet(path, document, root.LocalName);
            var viewer = new EDefterViewerForm(path, stylesheet, edefterTemplateFolder);
            AddToHistory(path);
            viewer.Show(this);
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "E-defter XML dosyası açılamadı.\r\n\r\n" + ex.Message,
                "E-defter görüntüleyici", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }
    }

    private void MinimizeToFloatingIcon(string openedFilePath = null)
    {
        if (miniRestoreIcon != null) return;
        fileSelectionToRestore = null;
        if (tabs.SelectedIndex == 0)
        {
            var selectedPaths = files.SelectedItems.Cast<ListViewItem>()
                .Select(item => item.Tag as Entry)
                .Where(entry => entry != null)
                .Select(entry => entry.FullPath)
                .ToList();
            if (!String.IsNullOrWhiteSpace(openedFilePath) && files.Items.Cast<ListViewItem>()
                .Any(item => { Entry entry = item.Tag as Entry; return entry != null && String.Equals(entry.FullPath, openedFilePath, StringComparison.OrdinalIgnoreCase); }) &&
                !selectedPaths.Contains(openedFilePath, StringComparer.OrdinalIgnoreCase))
                selectedPaths.Add(openedFilePath);
            fileSelectionToRestore = selectedPaths.ToArray();
        }
        savedWindowBounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        savedWindowState = WindowState == FormWindowState.Minimized ? lastNonMinimizedState : WindowState;
        savedTopMost = TopMost;
        WindowState = FormWindowState.Normal;
        Hide();
        Rectangle workingArea = Screen.FromRectangle(savedWindowBounds).WorkingArea;
        Point iconLocation = new Point(workingArea.Right - 113, workingArea.Top + 72);
        miniRestoreIcon = new MiniRestoreIcon(RestoreFromFloatingIcon, () => { exitRequested = true; Close(); }, iconLocation);
        miniRestoreIcon.Show();
    }

    private void RestoreFromFloatingIcon()
    {
        MiniRestoreIcon icon = miniRestoreIcon;
        miniRestoreIcon = null;
        if (icon != null) icon.Close();
        Bounds = savedWindowBounds;
        TopMost = savedTopMost;
        Show();
        WindowState = savedWindowState;
        Activate();
        RestoreExplorerFileSelection();
    }

    private void RestoreExplorerFileSelection()
    {
        string[] paths = fileSelectionToRestore;
        fileSelectionToRestore = null;
        if (paths == null || paths.Length == 0 || tabs.SelectedIndex != 0 || IsDisposed || !IsHandleCreated) return;
        try
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (IsDisposed || files.IsDisposed || tabs.SelectedIndex != 0) return;
                foreach (ListViewItem row in files.Items) row.Selected = false;
                ListViewItem first = null;
                foreach (ListViewItem row in files.Items)
                {
                    Entry entry = row.Tag as Entry;
                    if (entry == null || !paths.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase)) continue;
                    row.Selected = true;
                    row.Focused = first == null;
                    if (first == null) first = row;
                }
                if (first != null)
                {
                    files.FocusedItem = first;
                    first.EnsureVisible();
                    files.Focus();
                }
            }));
        }
        catch { }
    }

    private void AddSelectedShortcuts()
    {
        var selected = new List<Entry>();
        if (pendingFolder != null) selected.Add(pendingFolder);
        foreach (Entry item in folders.SelectedItems) selected.Add(item);
        foreach (ListViewItem fileItem in files.SelectedItems)
        {
            Entry item = fileItem.Tag as Entry;
            if (item != null) selected.Add(item);
        }
        foreach (Entry item in selected)
            if (!shortcuts.Any(x => String.Equals(x.FullPath, item.FullPath, StringComparison.OrdinalIgnoreCase)))
                shortcuts.Add(new Entry(DisplayName(item.FullPath), item.FullPath, DateTime.Now));
        pendingFolder = null;
        if (selected.Count == 0)
        {
            MessageBox.Show(this, "Kısayol eklemek için listeden bir klasör veya dosya seçin. Klasörleri seçmek için Ctrl+tıklayın.", "Kısayol ekle", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        SaveShortcuts();
        RefreshShortcutList();
        tabs.SelectedIndex = 1;
    }

    private void LoadShortcuts()
    {
        try
        {
            if (File.Exists(shortcutsFile))
            {
                bool migrated = false;
                foreach (string line in File.ReadAllLines(shortcutsFile))
                {
                    if (String.IsNullOrWhiteSpace(line)) continue;
                    int separator = line.LastIndexOf('\t');
                    string path = line;
                    DateTime added = DateTime.Now;
                    long binary;
                    if (separator > 0 && Int64.TryParse(line.Substring(separator + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out binary))
                    {
                        path = line.Substring(0, separator);
                        try { added = DateTime.FromBinary(binary); }
                        catch { added = DateTime.Now; }
                    }
                    else migrated = true;
                    path = ResolvePortablePath(path);
                    if (oneDriveRootEnabled && !String.IsNullOrWhiteSpace(path) &&
                        String.Equals(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                            oneDriveRoot.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                            StringComparison.OrdinalIgnoreCase)) continue;
                    if (!String.IsNullOrWhiteSpace(path)) shortcuts.Add(new Entry(DisplayName(path), path, added));
                }
                if (migrated) SaveShortcuts();
            }
        }
        catch { }
        LoadPortableShortcuts();
        RefreshShortcutList();
    }

    private void LoadPortableShortcuts()
    {
        try
        {
            if (!File.Exists(workProfilesFile)) return;
            var document = new XmlDocument();
            document.Load(workProfilesFile);
            foreach (XmlNode node in document.SelectNodes("/WorkProfiles/PortableShortcuts/Shortcut"))
            {
                Guid profileId;
                if (!Guid.TryParse(node.Attributes["profileId"] == null ? null : node.Attributes["profileId"].Value, out profileId)) continue;
                WorkProfile profile = workProfiles.FirstOrDefault(item => item.Id == profileId);
                if (profile == null || String.IsNullOrWhiteSpace(profile.RootPath)) continue;
                string relative = node.Attributes["path"] == null ? String.Empty : node.Attributes["path"].Value;
                string nativeRelative = relative.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
                if (String.IsNullOrWhiteSpace(relative) || Path.IsPathRooted(nativeRelative)) continue;
                string path = Path.GetFullPath(Path.Combine(profile.RootPath, nativeRelative));
                if (!IsPathWithin(path, profile.RootPath) || String.Equals(path, profile.RootPath, StringComparison.OrdinalIgnoreCase)) continue;
                DateTime added;
                if (!DateTime.TryParse(node.Attributes["added"] == null ? null : node.Attributes["added"].Value,
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out added)) added = DateTime.Now;
                if (!shortcuts.Any(item => String.Equals(item.FullPath, path, StringComparison.OrdinalIgnoreCase)))
                    shortcuts.Add(new Entry(DisplayName(path), path, added));
            }
        }
        catch { }
    }

    private void SaveShortcuts()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(shortcutsFile));
            var local = new List<Entry>();
            foreach (Entry entry in shortcuts)
            {
                if (oneDriveRootEnabled && !String.IsNullOrWhiteSpace(oneDriveRoot) &&
                    String.Equals(entry.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                        oneDriveRoot.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                        StringComparison.OrdinalIgnoreCase)) continue;
                bool isPortable = false;
                foreach (WorkProfile candidate in workProfiles.OrderByDescending(item => item.RootPath == null ? 0 : item.RootPath.Length))
                {
                    string candidateRelative;
                    if (!String.IsNullOrWhiteSpace(candidate.RootPath) &&
                        !String.Equals(entry.FullPath, candidate.RootPath, StringComparison.OrdinalIgnoreCase) &&
                        TryGetRelativeWorkPath(candidate.RootPath, entry.FullPath, out candidateRelative))
                    {
                        isPortable = true;
                        break;
                    }
                }
                if (!isPortable) local.Add(entry);
            }
            var mappedProfileIds = new HashSet<Guid>(workProfiles.Where(profile => !String.IsNullOrWhiteSpace(profile.RootPath)).Select(profile => profile.Id));
            portableShortcutRefs.RemoveAll(reference => mappedProfileIds.Contains(reference.Item1));
            File.WriteAllLines(shortcutsFile, local.Select(x => StorePortablePath(x.FullPath) + "\t" + x.AddedAt.ToBinary().ToString(CultureInfo.InvariantCulture)).ToArray());
            SaveWorkProfiles();
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Kısayollar kaydedilemedi", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }

    private void RefreshShortcutList()
    {
        shortcutList.BeginUpdate();
        shortcutList.Items.Clear();
        IEnumerable<Entry> items = shortcuts.Where(x => MatchesFileFilter(x.FullPath, shortcutFilter.SelectedItem as string));
        if (shortcutSortColumn == 0)
            items = shortcutSortAscending
                ? items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => Path.GetFileName(x.FullPath), NaturalNameComparer.Instance)
                : items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenByDescending(x => Path.GetFileName(x.FullPath), NaturalNameComparer.Instance);
        else if (shortcutSortColumn == 1)
            items = shortcutSortAscending ? items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => x.AddedAt) : items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenByDescending(x => x.AddedAt);
        else if (shortcutSortColumn == 2)
            items = shortcutSortAscending ? items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => FileLength(x.FullPath)) : items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenByDescending(x => FileLength(x.FullPath));
        else items = items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => Path.GetFileName(x.FullPath), NaturalNameComparer.Instance);
        foreach (Entry entry in items)
        {
            var row = new ListViewItem(new string[] { Path.GetFileName(entry.FullPath), entry.AddedAt.ToString("dd.MM.yyyy HH:mm"), FormatFileSize(entry.FullPath) });
            row.Tag = entry;
            row.ToolTipText = entry.FullPath;
            SetFileRowIcon(row, entry.FullPath);
            shortcutList.Items.Add(row);
        }
        shortcutList.EndUpdate();
    }

    private void SortShortcuts(int column, bool userClick)
    {
        if (userClick)
        {
            if (shortcutSortColumn == column) shortcutSortAscending = !shortcutSortAscending;
            else { shortcutSortColumn = column; shortcutSortAscending = true; }
        }
        else shortcutSortColumn = column;
        SetSortIndicators(shortcutList, "Ad", "Eklenme zamanı", shortcutSortColumn, shortcutSortAscending);
        if (userClick) SaveListSortPreference("Kısayollarım", shortcutSortColumn, shortcutSortAscending);
        RefreshShortcutList();
    }

    private void SortHistory(int column, bool userClick)
    {
        if (userClick)
        {
            if (historySortColumn == column) historySortAscending = !historySortAscending;
            else { historySortColumn = column; historySortAscending = true; }
        }
        else historySortColumn = column;
        SetSortIndicators(historyList, "Ad", "Eklenme zamanı", historySortColumn, historySortAscending);
        if (userClick) SaveListSortPreference("Geçmiş", historySortColumn, historySortAscending);
        RefreshHistoryList();
    }

    private void SortFiles(int column, bool userClick)
    {
        if (userClick)
        {
            if (fileSortColumn == column) fileSortAscending = !fileSortAscending;
            else { fileSortColumn = column; fileSortAscending = true; }
        }
        else if (fileSortColumn < 0) return;
        SetSortIndicators(files, "Ad", "Değiştirme tarihi", fileSortColumn, fileSortAscending);
        if (userClick) SaveListSortPreference("Gezgin", fileSortColumn, fileSortAscending);
        if (userClick && rememberFolderSort) RememberCurrentFolderSort();
        RefreshFileRows();
    }

    private void RefreshFileRows()
    {
        string[] selectedPaths = files.SelectedItems.Cast<ListViewItem>()
            .Select(item => item.Tag as Entry)
            .Where(entry => entry != null)
            .Select(entry => entry.FullPath)
            .ToArray();
        IEnumerable<ListViewItem> rows = fileRows.Where(x =>
        {
            string path = ((Entry)x.Tag).FullPath;
            return MatchesFileFilter(path, fileFilter.SelectedItem as string) && MatchesFileSearch(path);
        });
        if (fileSortColumn == 0)
            rows = fileSortAscending
                ? rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenBy(x => Path.GetFileName(((Entry)x.Tag).FullPath), NaturalNameComparer.Instance)
                : rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenByDescending(x => Path.GetFileName(((Entry)x.Tag).FullPath), NaturalNameComparer.Instance);
        else if (fileSortColumn == 1)
            rows = fileSortAscending
                ? rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenBy(x => File.GetLastWriteTime(((Entry)x.Tag).FullPath))
                : rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenByDescending(x => File.GetLastWriteTime(((Entry)x.Tag).FullPath));
        else if (fileSortColumn == 2)
            rows = fileSortAscending
                ? rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenBy(x => FileLength(((Entry)x.Tag).FullPath))
                : rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenByDescending(x => FileLength(((Entry)x.Tag).FullPath));
        else if (fileSortColumn == 3)
            rows = fileSortAscending
                ? rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? "Dosya klasörü" : GetFileTypeDescription(((Entry)x.Tag).FullPath), NaturalNameComparer.Instance)
                : rows.OrderBy(x => Directory.Exists(((Entry)x.Tag).FullPath) ? 0 : 1).ThenByDescending(x => Directory.Exists(((Entry)x.Tag).FullPath) ? "Dosya klasörü" : GetFileTypeDescription(((Entry)x.Tag).FullPath), NaturalNameComparer.Instance);
        ListViewItem[] items = rows.ToArray();
        foreach (ListViewItem row in items)
        {
            string path = ((Entry)row.Tag).FullPath;
            if (Directory.Exists(path))
            {
                bool isDrive = String.Equals(Path.GetPathRoot(path), path, StringComparison.OrdinalIgnoreCase);
                if (row.SubItems.Count > 3) row.SubItems[3].Text = isDrive ? "Sürücü" : "Dosya klasörü";
                if (!isDrive) UpdateOneDriveStatusCell(row, path);
                continue;
            }
            if (row.SubItems.Count > 2) row.SubItems[2].Text = FormatFileSize(path);
            if (row.SubItems.Count > 3) row.SubItems[3].Text = GetFileTypeDescription(path);
            UpdateOneDriveStatusCell(row, path);
        }
        files.BeginUpdate();
        files.Items.Clear();
        files.Groups.Clear();
        files.ShowGroups = fileGroupMode != 0;
        if (files.ShowGroups)
        {
            var groups = new Dictionary<string, ListViewGroup>(StringComparer.CurrentCultureIgnoreCase);
            foreach (ListViewItem row in items)
            {
                Entry entry = row.Tag as Entry;
                if (entry == null) continue;
                string groupName = GetFileGroupName(entry.FullPath, fileGroupMode);
                ListViewGroup group;
                if (!groups.TryGetValue(groupName, out group))
                {
                    group = new ListViewGroup(groupName, HorizontalAlignment.Left);
                    groups.Add(groupName, group);
                    files.Groups.Add(group);
                }
                row.Group = group;
            }
        }
        files.Items.AddRange(items);
        foreach (ListViewItem row in items)
        {
            Entry entry = row.Tag as Entry;
            if (entry != null && selectedPaths.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase))
                row.Selected = true;
        }
        files.EndUpdate();
        RefreshCutFileAppearance();
        status.Text = Math.Max(0, folders.Items.Count - 1) + " klasör   ·   " + fileRows.Count(row => { Entry entry = row.Tag as Entry; return entry != null && File.Exists(entry.FullPath); }) + " dosya";
    }

    private void UpdateSelectedFileCountStatus()
    {
        if (IsDisposed || !status.IsHandleCreated || showingComputerDrives) return;
        int count = files.Items.Cast<ListViewItem>()
            .Where(item => item.Selected || item.Checked)
            .Select(item => item.Tag as Entry)
            .Count(entry => entry != null && File.Exists(entry.FullPath));
        if (count > 0) status.Text = count + (count == 1 ? " dosya seçildi" : " dosya seçildi");
        else status.Text = Math.Max(0, folders.Items.Count - 1) + " klasör   ·   " + files.Items.Cast<ListViewItem>()
            .Count(item => { Entry entry = item.Tag as Entry; return entry != null && File.Exists(entry.FullPath); }) + " dosya";
    }

    private void SetFileGrouping(int mode)
    {
        fileGroupMode = mode;
        RefreshFileRows();
    }

    private string GetFileGroupName(string path, int mode)
    {
        if (Directory.Exists(path)) return "Klasörler";
        if (mode == 1) return GetFileTypeDescription(path);
        if (mode == 2)
        {
            DateTime modified = File.GetLastWriteTime(path).Date;
            if (modified == DateTime.Today) return "Bugün";
            if (modified >= DateTime.Today.AddDays(-7)) return "Bu hafta";
            if (modified >= DateTime.Today.AddMonths(-1)) return "Bu ay";
            return "Daha eski";
        }
        long size = FileLength(path);
        if (size < 1024L * 1024L) return "1 MB'den küçük";
        if (size < 100L * 1024L * 1024L) return "1–100 MB";
        return "100 MB ve üzeri";
    }

    private static void SetSortIndicators(ListView list, string nameHeader, string dateHeader, int column, bool ascending)
    {
        list.Columns[0].Text = nameHeader + (column == 0 ? (ascending ? "  ▲" : "  ▼") : "");
        list.Columns[1].Text = dateHeader + (column == 1 ? (ascending ? "  ▲" : "  ▼") : "");
        if (list.Columns.Count > 2)
            list.Columns[2].Text = "Boyut" + (column == 2 ? (ascending ? "  ▲" : "  ▼") : "");
        if (list.Columns.Count > 3)
            list.Columns[3].Text = "Tür" + (column == 3 ? (ascending ? "  ▲" : "  ▼") : "");
    }

    private static long FileLength(string path)
    {
        try { return File.Exists(path) ? new FileInfo(path).Length : -1L; }
        catch { return -1L; }
    }

    private static string FormatFileSize(string path)
    {
        long bytes = FileLength(path);
        return FormatFileSize(bytes);
    }

    private static string FormatFileSize(long bytes)
    {
        if (bytes < 0) return String.Empty;
        double kilobytes = bytes / 1024.0;
        if (bytes < 1024L * 1024L)
            return kilobytes.ToString(bytes < 10240 ? "0.00" : "0.#", CultureInfo.CurrentCulture) + " KB";
        return (kilobytes / 1024.0).ToString("0.##", CultureInfo.CurrentCulture) + " MB";
    }

    private string GetFileTypeDescription(string path)
    {
        string extension = Path.GetExtension(path);
        string cached;
        lock (fileTypeCache) if (fileTypeCache.TryGetValue(extension, out cached)) return cached;
        string description = String.Empty;
        try
        {
            var info = new ShellFileInfo();
            SHGetFileInfo(path, 0, ref info, (uint)Marshal.SizeOf(typeof(ShellFileInfo)), 0x400);
            description = (info.TypeName ?? String.Empty).Trim();
        }
        catch { }
        if (description.Length == 0)
            description = extension.Length == 0 ? "Dosya" : extension.TrimStart('.').ToUpper(CultureInfo.CurrentCulture) + " dosyası";
        lock (fileTypeCache) fileTypeCache[extension] = description;
        return description;
    }

    private void InitializeFileIconList()
    {
        fileIconImageList.ColorDepth = ColorDepth.Depth32Bit;
        fileIconImageList.ImageSize = new Size(16, 16);
        fileIconImageList.TransparentColor = Color.Transparent;
        folderIconKey = EnsureFileIcon("", true);
    }

    private void DrawFolderListItem(ListBox list, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= list.Items.Count) return;
        e.DrawBackground();
        Entry entry = list.Items[e.Index] as Entry;
        string label = entry == null ? Convert.ToString(list.Items[e.Index]) : entry.ToString();
        bool hasFolder = entry != null && !entry.IsParent && Directory.Exists(entry.FullPath);
        bool isComputer = entry != null && String.Equals(entry.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase);
        int textLeft = e.Bounds.Left + 7;
        if (isComputer)
        {
            int iconTop = e.Bounds.Top + Math.Max(3, (e.Bounds.Height - 16) / 2);
            Rectangle monitor = new Rectangle(textLeft + 1, iconTop, 15, 10);
            using (var fill = new SolidBrush(Color.FromArgb(214, 235, 250))) e.Graphics.FillRectangle(fill, monitor);
            using (var border = new Pen(Color.FromArgb(70, 115, 150), 1.4F))
            {
                e.Graphics.DrawRectangle(border, monitor);
                e.Graphics.DrawLine(border, textLeft + 8, iconTop + 10, textLeft + 8, iconTop + 13);
                e.Graphics.DrawLine(border, textLeft + 4, iconTop + 14, textLeft + 12, iconTop + 14);
            }
            textLeft += 23;
        }
        else if (hasFolder && !String.IsNullOrEmpty(folderIconKey) && fileIconImageList.Images.ContainsKey(folderIconKey))
        {
            Image icon = fileIconImageList.Images[folderIconKey];
            int y = e.Bounds.Top + Math.Max(0, (e.Bounds.Height - icon.Height) / 2);
            e.Graphics.DrawImage(icon, textLeft, y, icon.Width, icon.Height);
            textLeft += icon.Width + 7;
        }
        var textBounds = new Rectangle(textLeft, e.Bounds.Top, Math.Max(0, e.Bounds.Right - textLeft - 4), e.Bounds.Height);
        bool isCut = entry != null && pendingCutPaths.Contains(entry.FullPath, StringComparer.OrdinalIgnoreCase);
        Color color = (e.State & DrawItemState.Selected) != 0 ? SystemColors.HighlightText : (isCut ? Color.Gray : list.ForeColor);
        TextRenderer.DrawText(e.Graphics, label, list.Font, textBounds, color,
            TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        e.DrawFocusRectangle();
    }

    private string EnsureFileIcon(string path, bool isDirectory = false)
    {
        string extension = isDirectory ? "<folder>" : Path.GetExtension(path ?? String.Empty).ToLowerInvariant();
        string cachedKey;
        if (fileIconKeyCache.TryGetValue(extension, out cachedKey)) return cachedKey;
        string key = "icon:" + (extension.Length == 0 ? "<file>" : extension);
        try
        {
            string shellPath = isDirectory ? "folder" : "file" + extension;
            uint attributes = isDirectory ? 0x10u : 0x80u;
            var info = new ShellFileInfo();
            IntPtr result = SHGetFileInfo(shellPath, attributes, ref info, (uint)Marshal.SizeOf(typeof(ShellFileInfo)),
                0x00000100 | 0x00000001 | 0x00000010);
            if (result != IntPtr.Zero && info.Icon != IntPtr.Zero)
            {
                try
                {
                    using (Icon icon = Icon.FromHandle(info.Icon))
                    using (Bitmap bitmap = icon.ToBitmap()) fileIconImageList.Images.Add(key, (Image)bitmap.Clone());
                }
                finally { DestroyIcon(info.Icon); }
            }
        }
        catch { }
        if (!fileIconImageList.Images.ContainsKey(key))
        {
            using (Bitmap fallback = SystemIcons.Application.ToBitmap()) fileIconImageList.Images.Add(key, (Image)fallback.Clone());
        }
        fileIconKeyCache[extension] = key;
        return key;
    }

    private void SetFileRowIcon(ListViewItem row, string path)
    {
        if (row == null || String.IsNullOrWhiteSpace(path)) return;
        bool isDirectory = Directory.Exists(path);
        row.ImageKey = EnsureFileIcon(path, isDirectory);
    }

    private string QueryOneDriveStatus(string path)
    {
        if (!IsOneDrivePath(path)) return "—";
        string result = "Durum alınamadı";
        try
        {
            FileAttributes attributes = File.GetAttributes(path);
            int storageState = ReadStorageProviderState(path);
            int transferState = ReadShellIntProperty(path, "System.SyncTransferStatus");
            uint placeholderState = ReadCloudPlaceholderState(path);
            bool onlineOnly = (((int)attributes & 0x00400000) != 0) || (((int)attributes & 0x00040000) != 0);
            bool alwaysAvailable = (((int)attributes & 0x00080000) != 0);
            bool isInSync = (placeholderState & 0x00000008) != 0;
            bool isPartiallyOnDisk = (placeholderState & 0x00000020) != 0;
            if (transferState >= 0)
            {
                if ((transferState & 0x10) != 0) result = "Hata";
                else if ((transferState & 0x08) != 0) result = "Duraklatıldı";
                else if ((transferState & (0x01 | 0x02 | 0x04 | 0x20 | 0x40 | 0x200)) != 0) result = "Güncelleniyor";
                else if (isInSync) result = onlineOnly || isPartiallyOnDisk ? "Yalnızca çevrimiçi" : (alwaysAvailable ? "Bu cihazda her zaman" : "Güncel");
                else if ((transferState & 0x80) != 0) result = "Güncelleniyor";
                else if (transferState == 0)
                    result = onlineOnly || isPartiallyOnDisk ? "Yalnızca çevrimiçi" : (alwaysAvailable ? "Bu cihazda her zaman" : "Güncel");
                else result = "Güncelleniyor";
            }
            else if (isInSync)
                result = onlineOnly || isPartiallyOnDisk ? "Yalnızca çevrimiçi" : (alwaysAvailable ? "Bu cihazda her zaman" : "Güncel");
            else switch (storageState)
            {
                case 0: result = onlineOnly ? "Yalnızca çevrimiçi" : (alwaysAvailable ? "Bu cihazda her zaman" : "Güncel"); break;
                case 1: result = "Güncelleniyor"; break;
                case 2: result = "Duraklatıldı"; break;
                case 3: result = "Hata"; break;
                case 4: result = "Güncelleniyor"; break;
                case 5: result = "Çevrimdışı"; break;
                default:
                    if (onlineOnly) result = "Yalnızca çevrimiçi";
                    else if (alwaysAvailable) result = "Bu cihazda her zaman";
                    else if (((int)attributes & 0x00100000) != 0) result = "Yerel kullanılabilir";
                    break;
            }
        }
        catch { }
        return result;
    }

    private static uint ReadCloudPlaceholderState(string path)
    {
        IntPtr handle = IntPtr.Zero;
        try
        {
            const uint shareReadWriteDelete = 0x00000001 | 0x00000002 | 0x00000004;
            const uint openExisting = 3;
            const uint backupSemantics = 0x02000000;
            handle = CreateFile(path, 0, shareReadWriteDelete, IntPtr.Zero, openExisting, backupSemantics, IntPtr.Zero);
            if (handle == IntPtr.Zero || handle == new IntPtr(-1)) return 0;
            FileAttributeTagInformation info;
            uint size = (uint)Marshal.SizeOf(typeof(FileAttributeTagInformation));
            if (!GetFileInformationByHandleEx(handle, 9, out info, size)) return 0;
            return CfGetPlaceholderStateFromAttributeTag(info.FileAttributes, info.ReparseTag);
        }
        catch { return 0; }
        finally { if (handle != IntPtr.Zero && handle != new IntPtr(-1)) CloseHandle(handle); }
    }

    private void StartOneDriveStatusPolling()
    {
        if (oneDriveStatusWorker == null)
        {
            oneDriveStatusWorker = new BackgroundWorker { WorkerSupportsCancellation = true };
            oneDriveStatusWorker.DoWork += (s, e) =>
            {
                string[] paths = e.Argument as string[];
                var results = new List<KeyValuePair<string, string>>();
                if (paths != null)
                    foreach (string path in paths)
                    {
                        if (oneDriveStatusWorker.CancellationPending) { e.Cancel = true; break; }
                        results.Add(new KeyValuePair<string, string>(path, QueryOneDriveStatus(path)));
                    }
                e.Result = results;
            };
            oneDriveStatusWorker.RunWorkerCompleted += (s, e) =>
            {
                if (IsDisposed || Disposing || e.Cancelled || e.Error != null) return;
                var results = e.Result as List<KeyValuePair<string, string>>;
                if (results == null) return;
                foreach (KeyValuePair<string, string> result in results)
                {
                    oneDriveStatusCache[result.Key] = result.Value;
                    oneDriveStatusLastChecked[result.Key] = DateTime.UtcNow;
                    foreach (ListViewItem row in files.Items)
                    {
                        Entry entry = row.Tag as Entry;
                        if (entry != null && String.Equals(entry.FullPath, result.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            ApplyOneDriveStatusCell(row, result.Key, result.Value);
                            break;
                        }
                    }
                }
                files.Invalidate();
            };
        }
        if (oneDriveStatusTimer != null) oneDriveStatusTimer.Dispose();
        oneDriveStatusTimer = new System.Windows.Forms.Timer { Interval = 15000 };
        oneDriveStatusTimer.Tick += (s, e) => RefreshVisibleOneDriveStatuses();
        oneDriveStatusTimer.Start();
    }

    private void RefreshVisibleOneDriveStatuses()
    {
        if (IsDisposed || Disposing || files.IsDisposed || files.Columns.Count < 5 || files.Items.Count == 0 || oneDriveStatusWorker == null || oneDriveStatusWorker.IsBusy) return;
        int first = 0;
        try { if (files.TopItem != null) first = files.TopItem.Index; } catch { }
        int rowHeight = 22;
        try { if (first < files.Items.Count && files.Items[first].Bounds.Height > 0) rowHeight = files.Items[first].Bounds.Height; } catch { }
        int visibleCount = Math.Max(1, files.ClientSize.Height / Math.Max(1, rowHeight)) + 4;
        int last = Math.Min(files.Items.Count, first + visibleCount);
        DateTime now = DateTime.UtcNow;
        var paths = new List<string>();
        for (int i = first; i < last; i++)
        {
            ListViewItem item = files.Items[i];
            Entry entry = item.Tag as Entry;
            if (entry == null || item.SubItems.Count < 5 || !IsOneDrivePath(entry.FullPath)) continue;
            DateTime checkedAt;
            string status;
            bool hasStatus = oneDriveStatusCache.TryGetValue(entry.FullPath, out status);
            bool hasTime = oneDriveStatusLastChecked.TryGetValue(entry.FullPath, out checkedAt);
            int interval = String.Equals(status, "Güncelleniyor", StringComparison.Ordinal) ? 15000 : 60000;
            if (!hasStatus || !hasTime || (now - checkedAt).TotalMilliseconds >= interval)
                paths.Add(entry.FullPath);
        }
        if (paths.Count > 0) oneDriveStatusWorker.RunWorkerAsync(paths.ToArray());
    }

    private void UpdateOneDriveStatusCell(ListViewItem item, string path)
    {
        if (item == null || item.SubItems.Count < 5) return;
        string statusText;
        if (!IsOneDrivePath(path)) statusText = "—";
        else if (!oneDriveStatusCache.TryGetValue(path, out statusText)) statusText = "Denetleniyor";
        ApplyOneDriveStatusCell(item, path, statusText);
    }

    private void ApplyOneDriveStatusCell(ListViewItem item, string path, string statusText)
    {
        if (item == null || item.SubItems.Count < 5) return;
        string glyph;
        Color color;
        switch (statusText)
        {
            case "Güncelleniyor": glyph = "sync"; color = Color.FromArgb(0, 120, 212); break;
            case "Güncel": case "Yerel kullanılabilir": glyph = "check"; color = Color.FromArgb(16, 124, 16); break;
            case "Bu cihazda her zaman": glyph = "pinned"; color = Color.FromArgb(16, 124, 16); break;
            case "Duraklatıldı": glyph = "paused"; color = Color.FromArgb(160, 110, 0); break;
            case "Hata": glyph = "error"; color = Color.FromArgb(196, 43, 28); break;
            case "Çevrimdışı": case "Yalnızca çevrimiçi": glyph = "cloud"; color = Color.FromArgb(0, 120, 212); break;
            default: glyph = IsOneDrivePath(path) ? "cloud" : String.Empty; color = Color.FromArgb(0, 120, 212); break;
        }
        ListViewItem.ListViewSubItem cell = item.SubItems[4];
        // Keep a zero-width character so WinForms still raises DrawSubItem for this cell.
        cell.Text = "\u200B";
        cell.Tag = glyph;
        cell.ForeColor = color;
        item.ToolTipText = statusText;
    }

    private void DrawFileSubItem(object sender, DrawListViewSubItemEventArgs e)
    {
        bool isStatusColumn = oneDriveStatusColumn != null &&
            ((e.Header != null && Object.ReferenceEquals(e.Header, oneDriveStatusColumn)) ||
             (e.ColumnIndex >= 0 && e.ColumnIndex < files.Columns.Count &&
              Object.ReferenceEquals(files.Columns[e.ColumnIndex], oneDriveStatusColumn)));
        if (!isStatusColumn)
        {
            e.DrawDefault = true;
            return;
        }
        bool selected = e.Item.Selected;
        Color background = selected ? SystemColors.Highlight : files.BackColor;
        using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
        string glyph = e.SubItem.Tag as string;
        if (!String.IsNullOrEmpty(glyph))
        {
            Rectangle iconBounds = new Rectangle(e.Bounds.Left + (e.Bounds.Width - 20) / 2,
                e.Bounds.Top + (e.Bounds.Height - 16) / 2, 20, 16);
            Color iconColor = selected ? SystemColors.HighlightText : e.SubItem.ForeColor;
            DrawOneDriveStatusGlyph(e.Graphics, iconBounds, glyph, iconColor);
        }
        using (var pen = new Pen(selected ? SystemColors.Highlight : Color.FromArgb(232, 235, 239)))
        {
            e.Graphics.DrawLine(pen, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom);
            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }
    }

    private static void DrawOneDriveStatusGlyph(Graphics graphics, Rectangle bounds, string glyph, Color color)
    {
        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (var pen = new Pen(color, 1.7F))
        using (var fill = new SolidBrush(color))
        {
            pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
            pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
            float x = bounds.Left, y = bounds.Top;
            switch (glyph)
            {
                case "cloud":
                    using (var cloud = new System.Drawing.Drawing2D.GraphicsPath())
                    {
                        cloud.StartFigure();
                        cloud.AddBezier(x + 4, y + 13, x + 1, y + 13, x + 1, y + 9, x + 4, y + 8);
                        cloud.AddBezier(x + 4, y + 8, x + 4, y + 5, x + 8, y + 5, x + 9, y + 7);
                        cloud.AddBezier(x + 9, y + 7, x + 10, y + 2, x + 16, y + 3, x + 16, y + 8);
                        cloud.AddBezier(x + 16, y + 8, x + 19, y + 8, x + 19, y + 13, x + 15, y + 13);
                        cloud.AddLine(x + 4, y + 13, x + 15, y + 13);
                        graphics.DrawPath(pen, cloud);
                    }
                    break;
                case "sync":
                    DrawWindowsUpdateGlyph(graphics, bounds, color);
                    break;
                case "check":
                    graphics.DrawLines(pen, new PointF[] { new PointF(x + 3, y + 8), new PointF(x + 7, y + 12), new PointF(x + 16, y + 3) });
                    break;
                case "pinned":
                    graphics.DrawEllipse(pen, x + 5, y + 1, 9, 9);
                    graphics.DrawLine(pen, x + 9.5F, y + 10, x + 9.5F, y + 15);
                    graphics.DrawLine(pen, x + 6, y + 6, x + 13, y + 6);
                    break;
                case "paused":
                    graphics.FillRectangle(fill, x + 5, y + 2, 4, 12);
                    graphics.FillRectangle(fill, x + 12, y + 2, 4, 12);
                    break;
                case "error":
                    graphics.DrawEllipse(pen, x + 2, y + 1, 15, 15);
                    graphics.DrawLine(pen, x + 6, y + 5, x + 13, y + 12);
                    graphics.DrawLine(pen, x + 13, y + 5, x + 6, y + 12);
                    break;
            }
        }
    }

    private static void DrawWindowsUpdateGlyph(Graphics graphics, Rectangle bounds, Color color)
    {
        string glyph = "\uF0C5";
        string fontName = "Segoe Fluent Icons";
        try
        {
            bool hasFluent = IsFontInstalled(fontName);
            if (!hasFluent)
            {
                fontName = "Segoe MDL2 Assets";
                glyph = "\uE895";
            }
            using (var iconFont = new Font(fontName, 16F, FontStyle.Regular, GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(color))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                graphics.DrawString(glyph, iconFont, brush, (RectangleF)bounds, format);
        }
        catch
        {
            // Older Windows builds may not have the Windows Update glyph font; retain a legible sync symbol.
            using (var pen = new Pen(color, 1.7F))
            using (var brush = new SolidBrush(color))
            {
                pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                graphics.DrawArc(pen, bounds.Left + 4, bounds.Top + 2, 11, 11, 205, 225);
                graphics.DrawArc(pen, bounds.Left + 4, bounds.Top + 2, 11, 11, 25, 225);
                graphics.FillPolygon(brush, new PointF[] { new PointF(bounds.Left + 3, bounds.Top + 4), new PointF(bounds.Left + 7, bounds.Top + 4), new PointF(bounds.Left + 4, bounds.Top + 1) });
                graphics.FillPolygon(brush, new PointF[] { new PointF(bounds.Left + 16, bounds.Top + 12), new PointF(bounds.Left + 12, bounds.Top + 12), new PointF(bounds.Left + 15, bounds.Top + 15) });
            }
        }
    }

    private static bool? segoeFluentIconsAvailable;

    private static bool IsFontInstalled(string fontName)
    {
        if (String.Equals(fontName, "Segoe Fluent Icons", StringComparison.OrdinalIgnoreCase) && segoeFluentIconsAvailable.HasValue)
            return segoeFluentIconsAvailable.Value;
        bool installed;
        using (var fonts = new System.Drawing.Text.InstalledFontCollection())
            installed = fonts.Families.Any(family => String.Equals(family.Name, fontName, StringComparison.OrdinalIgnoreCase));
        if (String.Equals(fontName, "Segoe Fluent Icons", StringComparison.OrdinalIgnoreCase))
            segoeFluentIconsAvailable = installed;
        return installed;
    }

    private int ReadStorageProviderState(string path)
    {
        return ReadShellIntProperty(path, "System.StorageProviderState");
    }

    private int ReadShellIntProperty(string path, string propertyName)
    {
        object shellFolder = null;
        object shellItem = null;
        try
        {
            if (shellApplication == null)
            {
                Type shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null) return -1;
                shellApplication = Activator.CreateInstance(shellType);
            }
            string parent = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);
            if (String.IsNullOrEmpty(parent) || String.IsNullOrEmpty(fileName)) return -1;
            shellFolder = shellApplication.GetType().InvokeMember("NameSpace", BindingFlags.InvokeMethod, null, shellApplication, new object[] { parent });
            if (shellFolder == null) return -1;
            shellItem = shellFolder.GetType().InvokeMember("ParseName", BindingFlags.InvokeMethod, null, shellFolder, new object[] { fileName });
            if (shellItem == null) return -1;
            object value = shellItem.GetType().InvokeMember("ExtendedProperty", BindingFlags.InvokeMethod, null, shellItem, new object[] { propertyName });
            if (value == null) return -1;
            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }
        catch { return -1; }
        finally
        {
            if (shellItem != null && Marshal.IsComObject(shellItem)) { try { Marshal.ReleaseComObject(shellItem); } catch { } }
            if (shellFolder != null && Marshal.IsComObject(shellFolder)) { try { Marshal.ReleaseComObject(shellFolder); } catch { } }
        }
    }

    private bool IsOneDrivePath(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return false;
        if (oneDriveRoots == null)
        {
            var roots = new List<string>();
            foreach (string variable in new string[] { "OneDrive", "OneDriveConsumer", "OneDriveCommercial" })
            {
                string value = Environment.GetEnvironmentVariable(variable);
                if (!String.IsNullOrWhiteSpace(value) && Directory.Exists(value)) roots.Add(value);
            }
            try
            {
                using (RegistryKey accounts = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\OneDrive\\Accounts"))
                {
                    if (accounts != null)
                        foreach (string accountName in accounts.GetSubKeyNames())
                            using (RegistryKey account = accounts.OpenSubKey(accountName))
                            {
                                string value = account == null ? null : account.GetValue("UserFolder") as string;
                                if (!String.IsNullOrWhiteSpace(value) && Directory.Exists(value)) roots.Add(value);
                            }
                }
            }
            catch { }
            oneDriveRoots = roots.Select(root => Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }
        try
        {
            string fullPath = Path.GetFullPath(path);
            return oneDriveRoots.Any(root => String.Equals(fullPath, root, StringComparison.OrdinalIgnoreCase) ||
                fullPath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    private static string DisplayName(string path)
    {
        string name = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        if (String.IsNullOrEmpty(name)) name = path;
        return (Directory.Exists(path) ? "📁  " : "📄  ") + name;
    }

    private void RevealInExplorer(string path, bool keepCurrentTab = false)
    {
        if (!keepCurrentTab) tabs.SelectedIndex = 0;
        if (Directory.Exists(path)) Navigate(path);
        else if (File.Exists(path))
        {
            string parent = Path.GetDirectoryName(path);
            if (!String.IsNullOrEmpty(parent)) Navigate(parent, null, path);
        }
        else MessageBox.Show(this, "Bu öğenin hedefi bulunamadı:\n" + path, "Hedef bulunamadı", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void RevealSelectedItem(ListView list)
    {
        Entry entry = list.SelectedItems.Count > 0 ? list.SelectedItems[0].Tag as Entry : null;
        if (entry != null) RevealInExplorer(entry.FullPath, true);
    }

    private void RevealClickedItem(ListView list, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left || (ModifierKeys & (Keys.Control | Keys.Shift)) != Keys.None) return;
        ListViewItem row = list.HitTest(e.Location).Item;
        if (row == null) return;
        Entry entry = row.Tag as Entry;
        if (entry == null) return;
        RevealInExplorer(entry.FullPath, true);
    }

    private void OpenListItemInExplorer(ListView list, Point location)
    {
        ListViewItem row = list.HitTest(location).Item;
        Entry entry = row == null ? null : row.Tag as Entry;
        if (entry != null) RevealInExplorer(entry.FullPath);
    }

    private void ConfigureItemContextMenu(ListView list, bool historyMode, bool searchMode = false)
    {
        var menu = new ContextMenuStrip();
        var openFile = new ToolStripMenuItem("Dosya Aç");
        var openExplorerItem = new ToolStripMenuItem("Windows Explorer ile Aç");
        var preview = new ToolStripMenuItem("Ön izleme");
        var copy = new ToolStripMenuItem("Kopyala");
        var cut = new ToolStripMenuItem("Kes");
        var paste = new ToolStripMenuItem("Yapıştır");
        var rename = new ToolStripMenuItem("Yeniden adlandır");
        var delete = new ToolStripMenuItem("Sil");
        var removeFromList = new ToolStripMenuItem(historyMode ? "Geçmişten kaldır" : "Kısayollardan kaldır");
        var mail = new ToolStripMenuItem("Mail Gönder");
        var whatsappShare = new ToolStripMenuItem("WhatsApp için karakter uyumlu paylaş");
        menu.Items.Add(openFile);
        menu.Items.Add(openExplorerItem);
        menu.Items.Add(preview);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(copy);
        menu.Items.Add(cut);
        menu.Items.Add(paste);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(rename);
        menu.Items.Add(delete);
        if (!searchMode) menu.Items.Add(removeFromList);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(mail);
        menu.Items.Add(whatsappShare);
        list.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            ListViewItem row = list.HitTest(e.Location).Item;
            if (row != null && !row.Selected)
            {
                list.SelectedItems.Clear();
                row.Selected = true;
                list.FocusedItem = row;
            }
        };
        menu.Opening += (s, e) =>
        {
            string[] selectedPaths = GetSelectedListPaths(list);
            bool hasSelection = selectedPaths.Length > 0;
            Entry selected = GetSelectedEntry(list);
            e.Cancel = !hasSelection;
            openFile.Enabled = selected != null && (File.Exists(selected.FullPath) || Directory.Exists(selected.FullPath));
            openExplorerItem.Enabled = openFile.Enabled;
            preview.Enabled = selected != null && File.Exists(selected.FullPath);
            copy.Enabled = cut.Enabled = hasSelection;
            paste.Enabled = HasFileDropClipboard() || transferClipboardPaths.Count > 0;
            rename.Enabled = selected != null && File.Exists(selected.FullPath);
            delete.Enabled = hasSelection;
            removeFromList.Enabled = !searchMode && hasSelection;
            mail.Enabled = selected != null && File.Exists(selected.FullPath);
            whatsappShare.Enabled = mail.Enabled;
        };
        openFile.Click += (s, e) => OpenSelectedListFile(list);
        openExplorerItem.Click += (s, e) =>
        {
            Entry selected = GetSelectedEntry(list);
            if (selected != null) OpenInWindowsExplorer(selected.FullPath);
        };
        preview.Click += (s, e) => PreviewListSelection(list);
        copy.Click += (s, e) => CopyListSelection(list);
        cut.Click += (s, e) => CutListSelection(list);
        paste.Click += (s, e) => PasteFiles();
        rename.Click += (s, e) => RenameListSelection(list);
        delete.Click += (s, e) => DeleteListSelection(list);
        removeFromList.Click += (s, e) => RemoveSelectedListReferences(list, historyMode);
        mail.Click += (s, e) =>
        {
            Entry selected = GetSelectedEntry(list);
            if (selected != null && File.Exists(selected.FullPath)) SendByMail(selected.FullPath);
        };
        whatsappShare.Click += (s, e) =>
        {
            Entry selected = GetSelectedEntry(list);
            if (selected != null && File.Exists(selected.FullPath)) ShareWithWhatsAppSafeFileName(selected.FullPath);
        };
        SetContextMenuTooltips(menu.Items);
        list.ContextMenuStrip = menu;
        ConfigureFileListKeyboard(list,
            () => OpenSelectedListFile(list),
            () => DeleteListSelection(list),
            () => RenameListSelection(list),
            () => CutListSelection(list),
            () => CopyListSelection(list),
            () => PasteFiles());
    }

    private void PreviewListSelection(ListView list)
    {
        Entry selected = GetSelectedEntry(list);
        if (selected == null || !File.Exists(selected.FullPath)) return;
        if (String.Equals(Path.GetExtension(selected.FullPath), ".xml", StringComparison.OrdinalIgnoreCase))
        {
            if (!TryOpenEdefterXml(selected.FullPath)) new XmlSourceViewerForm(selected.FullPath).Show(this);
        }
        else new FilePreviewForm(selected.FullPath, true).Show(this);
    }

    private void RenameListSelection(ListView list)
    {
        Entry selected = GetSelectedEntry(list);
        if (selected == null || !File.Exists(selected.FullPath)) return;
        string oldPath = selected.FullPath;
        string oldName = Path.GetFileName(oldPath);
        using (var dialog = new RenameFileDialog(oldName))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string newName = dialog.NewName.Trim();
            if (String.IsNullOrWhiteSpace(newName) || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                newName == "." || newName == "..")
            {
                MessageBox.Show(this, "Geçerli bir dosya adı girin.", "Yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName);
            if (File.Exists(newPath) || Directory.Exists(newPath))
            {
                MessageBox.Show(this, "Bu adla başka bir dosya veya klasör zaten var.", "Yeniden adlandır", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                File.Move(oldPath, newPath);
                Entry oldShortcut = shortcuts.FirstOrDefault(item => String.Equals(item.FullPath, oldPath, StringComparison.OrdinalIgnoreCase));
                shortcuts.RemoveAll(item => String.Equals(item.FullPath, oldPath, StringComparison.OrdinalIgnoreCase));
                if (oldShortcut != null) shortcuts.Insert(0, new Entry(DisplayName(newPath), newPath, oldShortcut.AddedAt));
                HistoryEntry oldHistory = history.FirstOrDefault(item => String.Equals(item.FullPath, oldPath, StringComparison.OrdinalIgnoreCase));
                history.RemoveAll(item => String.Equals(item.FullPath, oldPath, StringComparison.OrdinalIgnoreCase));
                if (oldHistory != null) history.Insert(0, new HistoryEntry(newPath, oldHistory.OpenedAt));
                SaveShortcuts();
                SaveHistory();
                RefreshShortcutList();
                RefreshHistoryList();
                foreach (ListViewItem row in deepSearchResults.Items.Cast<ListViewItem>().Where(row =>
                {
                    Entry item = row.Tag as Entry;
                    return item != null && String.Equals(item.FullPath, oldPath, StringComparison.OrdinalIgnoreCase);
                }).ToArray()) deepSearchResults.Items.Remove(row);
                string parent = Path.GetDirectoryName(oldPath);
                if (String.Equals(currentPath, parent, StringComparison.OrdinalIgnoreCase)) Navigate(currentPath, null, newPath, false);
                if (!String.IsNullOrWhiteSpace(parent) && Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
                foreach (TransferPaneState pane in new TransferPaneState[] { transferLeftPane, transferRightPane })
                    if (pane != null && String.Equals(pane.CurrentPath, parent, StringComparison.OrdinalIgnoreCase)) TransferNavigate(pane, pane.CurrentPath);
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Yeniden adlandırılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }

    private void DeleteListSelection(ListView list)
    {
        string[] paths = GetSelectedListPaths(list);
        if (paths.Length == 0) return;
        DialogResult answer = MessageBox.Show(this, paths.Length + " seçili dosya/klasör Geri Dönüşüm Kutusu'na taşınsın mı?",
            "Silme onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes) return;
        try
        {
            MovePathsToRecycleBin(paths);
            lastRecycleBinDeletedPaths.Clear();
            lastRecycleBinDeletedPaths.AddRange(paths);
            shortcuts.RemoveAll(item => paths.Contains(item.FullPath, StringComparer.OrdinalIgnoreCase));
            history.RemoveAll(item => paths.Contains(item.FullPath, StringComparer.OrdinalIgnoreCase));
            SaveShortcuts();
            SaveHistory();
            foreach (ListView target in new ListView[] { shortcutList, historyList, deepSearchResults })
                foreach (ListViewItem row in target.Items.Cast<ListViewItem>().Where(row =>
                {
                    Entry item = row.Tag as Entry;
                    return item != null && paths.Contains(item.FullPath, StringComparer.OrdinalIgnoreCase);
                }).ToArray()) target.Items.Remove(row);
            RefreshShortcutList();
            RefreshHistoryList();
            string[] parents = paths.Select(Path.GetDirectoryName).Where(parent => !String.IsNullOrWhiteSpace(parent)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            foreach (string parent in parents) if (Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
            if (parents.Any(parent => String.Equals(parent, currentPath, StringComparison.OrdinalIgnoreCase))) Navigate(currentPath, null, null, false);
            status.Text = paths.Length + " öğe Geri Dönüşüm Kutusu'na taşındı.";
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void RemoveSelectedListReferences(ListView list, bool historyMode)
    {
        string[] paths = GetSelectedListPaths(list);
        if (paths.Length == 0) return;
        string label = historyMode ? "Geçmiş" : "Kısayollarım";
        if (MessageBox.Show(this, paths.Length + " öğe " + label + " listesinden kaldırılsın mı? Dosyalar silinmez.",
            label, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        if (historyMode) { history.RemoveAll(item => paths.Contains(item.FullPath, StringComparer.OrdinalIgnoreCase)); SaveHistory(); RefreshHistoryList(); }
        else { shortcuts.RemoveAll(item => paths.Contains(item.FullPath, StringComparer.OrdinalIgnoreCase)); SaveShortcuts(); RefreshShortcutList(); }
    }

    private void ConfigureMailContextMenu(ListBox list)
    {
        Entry contextTarget = null;
        var menu = new ContextMenuStrip();
        var openExplorerItem = new ToolStripMenuItem("Windows Explorer ile Aç");
        var enterFolder = new ToolStripMenuItem("Klasöre gir");
        var mail = new ToolStripMenuItem("Mail Gönder");
        var createFolder = new ToolStripMenuItem("Klasör ekle");
        var renameFolder = new ToolStripMenuItem("Değiştir...");
        var cutFolder = new ToolStripMenuItem("Kes");
        var copyFolder = new ToolStripMenuItem("Kopyala");
        var pasteFolder = new ToolStripMenuItem("Yapıştır");
        var deleteFolder = new ToolStripMenuItem("Sil");
        menu.Items.Add(openExplorerItem);
        menu.Items.Add(enterFolder);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(renameFolder);
        menu.Items.Add(cutFolder);
        menu.Items.Add(copyFolder);
        menu.Items.Add(pasteFolder);
        menu.Items.Add(deleteFolder);
        menu.Items.Add(mail);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(createFolder);
        list.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            int index = list.IndexFromPoint(e.Location);
            contextTarget = index >= 0 ? list.Items[index] as Entry : null;
            if (contextTarget != null && !list.GetSelected(index))
            {
                list.ClearSelected();
                list.SetSelected(index, true);
            }
        };
        menu.Opening += (s, e) =>
        {
            openExplorerItem.Enabled = contextTarget != null && Directory.Exists(contextTarget.FullPath);
            enterFolder.Enabled = contextTarget != null && Directory.Exists(contextTarget.FullPath);
            renameFolder.Enabled = contextTarget != null && !contextTarget.IsParent && Directory.Exists(contextTarget.FullPath);
            bool hasFolderSelection = GetSelectedFolderPaths(list).Length > 0;
            cutFolder.Enabled = copyFolder.Enabled = hasFolderSelection;
            pasteFolder.Enabled = HasFileDropClipboard() || (!String.IsNullOrWhiteSpace(pendingFolderClipboardPath) && Directory.Exists(pendingFolderClipboardPath));
            deleteFolder.Enabled = hasFolderSelection;
            mail.Enabled = contextTarget != null && File.Exists(contextTarget.FullPath);
            createFolder.Enabled = contextTarget == null || Directory.Exists(contextTarget.FullPath);
        };
        openExplorerItem.Click += (s, e) =>
        {
            if (contextTarget != null) OpenInWindowsExplorer(contextTarget.FullPath);
        };
        enterFolder.Click += (s, e) =>
        {
            if (contextTarget == null) return;
            if (String.Equals(contextTarget.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) ShowComputerDrives();
            else if (contextTarget.IsParent && showingComputerDrives) Navigate(contextTarget.FullPath, null, null, false);
            else if (Directory.Exists(contextTarget.FullPath)) Navigate(contextTarget.FullPath, contextTarget);
        };
        renameFolder.Click += (s, e) => RenameSelectedFolder(contextTarget);
        cutFolder.Click += (s, e) => SetFolderListClipboard(list, true);
        copyFolder.Click += (s, e) => SetFolderListClipboard(list, false);
        pasteFolder.Click += (s, e) => PasteClipboardIntoFolder(contextTarget != null && Directory.Exists(contextTarget.FullPath) ? contextTarget.FullPath : currentPath);
        deleteFolder.Click += (s, e) => DeleteFolderListSelection(list);
        mail.Click += (s, e) =>
        {
            if (contextTarget != null && File.Exists(contextTarget.FullPath))
                SendByMail(contextTarget.FullPath);
        };
        createFolder.Click += (s, e) => CreateFolderIn(contextTarget != null && Directory.Exists(contextTarget.FullPath)
            ? contextTarget.FullPath : currentPath);
        list.ContextMenuStrip = menu;
    }

    private void DeleteFolderListSelection(ListBox list)
    {
        string[] paths = GetSelectedFolderPaths(list);
        if (paths.Length == 0) return;
        if (MessageBox.Show(this, paths.Length + " seçili klasör Geri Dönüşüm Kutusu'na taşınsın mı?", "Silme onayı",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        try
        {
            MovePathsToRecycleBin(paths);
            foreach (string parent in paths.Select(Path.GetDirectoryName).Where(parent => !String.IsNullOrWhiteSpace(parent)).Distinct(StringComparer.OrdinalIgnoreCase))
                if (Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
            Navigate(currentPath, null, null, false);
            status.Text = paths.Length + " klasör Geri Dönüşüm Kutusu'na taşındı.";
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void DeleteFolderTreeSelection(Entry entry)
    {
        if (entry == null || !Directory.Exists(entry.FullPath) || String.Equals(entry.FullPath, ThisComputerPath, StringComparison.OrdinalIgnoreCase)) return;
        if (MessageBox.Show(this, "'" + Path.GetFileName(entry.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)) +
            "' klasörü Geri Dönüşüm Kutusu'na taşınsın mı?", "Silme onayı", MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        string parent = Path.GetDirectoryName(entry.FullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        try
        {
            MovePathsToRecycleBin(new[] { entry.FullPath });
            if (!String.IsNullOrWhiteSpace(parent) && Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
            Navigate(currentPath, null, null, false);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    private void ConfigureMailContextMenu(ListView list)
    {
        ListViewItem contextRow = null;
        var menu = new ContextMenuStrip();
        menu.BackColor = Color.FromArgb(250, 250, 250);
        menu.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point);
        menu.Padding = new Padding(1);
        var openFile = new ToolStripMenuItem("Dosyayı Aç");
        var openExplorerItem = new ToolStripMenuItem("Windows Explorer ile Aç");
        var preview = new ToolStripMenuItem("Ön izleme");
        var print = new ToolStripMenuItem("Seçili dosyaları yazdır");
        var mergePdfs = new ToolStripMenuItem("Seçili PDF'leri birleştir...");
        var ocrPdfToText = new ToolStripMenuItem("OCR ile TXT oluştur");
        var mail = new ToolStripMenuItem("Mail Gönder");
        var whatsappSafeShare = new ToolStripMenuItem("WhatsApp için karakter uyumlu paylaş");
        var cut = new ToolStripMenuItem("Kes");
        var copy = new ToolStripMenuItem("Kopyala");
        var paste = new ToolStripMenuItem("Yapıştır");
        var move = new ToolStripMenuItem("Taşı...");
        var undoMove = new ToolStripMenuItem("Geri Al Taşı");
        var rename = new ToolStripMenuItem("Yeniden adlandır");
        var delete = new ToolStripMenuItem("Sil");
        var undoDelete = new ToolStripMenuItem("Geri Al Silme");
        var zip = new ToolStripMenuItem("ZIP ile sıkıştır");
        var extractHere = new ToolStripMenuItem("Arşivi buraya ayıkla");
        var extractNamedFolder = new ToolStripMenuItem("Arşivi kendi klasörüne ayıkla");
        var extractChooseFolder = new ToolStripMenuItem("Arşivi seçilen klasöre ayıkla...");
        var compressionMenu = new ToolStripMenuItem("Sıkıştırma");
        compressionMenu.DropDownItems.Add(zip);
        compressionMenu.DropDownItems.Add(new ToolStripSeparator());
        compressionMenu.DropDownItems.Add(extractHere);
        compressionMenu.DropDownItems.Add(extractNamedFolder);
        compressionMenu.DropDownItems.Add(extractChooseFolder);
        var newMenu = new ToolStripMenuItem("Yeni");
        var viewMenu = new ToolStripMenuItem("Görünüm");
        var detailsView = new ToolStripMenuItem("Ayrıntılar") { Checked = true, Enabled = false };
        var previewPaneItem = new ToolStripMenuItem("Ön izleme paneli") { CheckOnClick = true };
        viewMenu.DropDownItems.Add(detailsView);
        viewMenu.DropDownItems.Add(previewPaneItem);
        var sortMenu = new ToolStripMenuItem("Sıralama ölçütü");
        var sortName = new ToolStripMenuItem("Ad");
        var sortDate = new ToolStripMenuItem("Değiştirme tarihi");
        var sortSize = new ToolStripMenuItem("Boyut");
        var sortType = new ToolStripMenuItem("Tür");
        sortMenu.DropDownItems.AddRange(new ToolStripItem[] { sortName, sortDate, sortSize, sortType });
        var groupMenu = new ToolStripMenuItem("Gruplandırma ölçütü");
        var groupNone = new ToolStripMenuItem("Gruplandırma yok");
        var groupType = new ToolStripMenuItem("Dosya türü");
        var groupDate = new ToolStripMenuItem("Değiştirme tarihine göre");
        var groupSize = new ToolStripMenuItem("Boyuta göre");
        groupMenu.DropDownItems.AddRange(new ToolStripItem[] { groupNone, groupType, groupDate, groupSize });
        var fileTasksMenu = new ToolStripMenuItem("Dosya işlemleri");
        fileTasksMenu.DropDownItems.AddRange(new ToolStripItem[] { print, mergePdfs, ocrPdfToText, new ToolStripSeparator(), mail });
        var manageMenu = new ToolStripMenuItem("Düzenle ve yönet");
        manageMenu.DropDownItems.AddRange(new ToolStripItem[] { rename, delete, undoDelete });
        var clipboardMenu = new ToolStripMenuItem("Pano ve taşıma");
        clipboardMenu.DropDownItems.AddRange(new ToolStripItem[] { cut, copy, paste, move });
        var moreMenu = new ToolStripMenuItem("Diğer seçenekler");
        var copyPath = new ToolStripMenuItem("Yolu kopyala");
        moreMenu.DropDownItems.AddRange(new ToolStripItem[] { openExplorerItem, copyPath });
        var properties = new ToolStripMenuItem("Özellikler");
        var terminal = new ToolStripMenuItem("Terminalde Aç");
        var codeEditor = new ToolStripMenuItem("Code ile Aç");
        var actionBar = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = false,
            Size = new Size(244, 38),
            Padding = new Padding(2, 1, 2, 1),
            Margin = Padding.Empty,
            BackColor = Color.White
        };
        Button cutButton = CreateExplorerMenuActionButton("Kes", "\uE8C6", 38);
        Button copyButton = CreateExplorerMenuActionButton("Kopyala", "\uE8C8", 54);
        Button pasteButton = CreateExplorerMenuActionButton("Yapıştır", "\uE77F", 45);
        Button shareButton = CreateExplorerMenuActionButton("Paylaş", "\uE72D", 40);
        Button deleteButton = CreateExplorerMenuActionButton("Sil", "\uE74D", 38);
        actionBar.Controls.Add(cutButton);
        actionBar.Controls.Add(CreateExplorerMenuSeparator());
        actionBar.Controls.Add(copyButton);
        actionBar.Controls.Add(CreateExplorerMenuSeparator());
        actionBar.Controls.Add(pasteButton);
        actionBar.Controls.Add(CreateExplorerMenuSeparator());
        actionBar.Controls.Add(shareButton);
        actionBar.Controls.Add(CreateExplorerMenuSeparator());
        actionBar.Controls.Add(deleteButton);
        var actionBarHost = new ToolStripControlHost(actionBar)
        {
            AutoSize = false,
            Size = new Size(244, 40),
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            BackColor = Color.White
        };
        menu.Items.Add(actionBarHost);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(openFile);
        var openWith = new ToolStripMenuItem("Birlikte aç...");
        menu.Items.Add(openWith);
        var pasteInEmptyArea = new ToolStripMenuItem("Yapıştır");
        var oneDriveMenu = new ToolStripMenuItem("OneDrive");
        var keepOffline = new ToolStripMenuItem("Her zaman bu cihazda sakla");
        var freeSpace = new ToolStripMenuItem("Yer aç");
        oneDriveMenu.DropDownItems.AddRange(new ToolStripItem[] { keepOffline, freeSpace });
        var winRarMenu = new ToolStripMenuItem("WinRAR");
        var rarOpen = new ToolStripMenuItem("WinRAR ile aç");
        var rarHere = new ToolStripMenuItem("Buraya ayıkla");
        var rarFolder = new ToolStripMenuItem("Klasöre ayıkla");
        winRarMenu.DropDownItems.AddRange(new ToolStripItem[] { rarOpen, rarHere, rarFolder });
        menu.Items.Add(clipboardMenu);
        menu.Items.Add(preview);
        menu.Items.Add(viewMenu);
        menu.Items.Add(sortMenu);
        menu.Items.Add(groupMenu);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(undoMove);
        menu.Items.Add(newMenu);
        menu.Items.Add(manageMenu);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(properties);
        menu.Items.Add(terminal);
        menu.Items.Add(codeEditor);
        menu.Items.Add(fileTasksMenu);
        menu.Items.Add(whatsappSafeShare);
        menu.Items.Add(compressionMenu);
        menu.Items.Add(oneDriveMenu);
        menu.Items.Add(winRarMenu);
        menu.Items.Add(moreMenu);
        menu.Items.Add(pasteInEmptyArea);
        menu.ShowItemToolTips = true;

        SetMenuGlyph(openFile, "\uE8A5");
        SetMenuGlyph(preview, "\uE890");
        SetMenuGlyph(viewMenu, "\uE7F4");
        SetMenuGlyph(sortMenu, "\uE8CB");
        SetMenuGlyph(groupMenu, "\uF168");
        SetMenuGlyph(undoMove, "\uE7A7");
        SetMenuGlyph(newMenu, "\uE710");
        SetMenuGlyph(manageMenu, "\uE70F");
        SetMenuGlyph(clipboardMenu, "\uE8C8");
        SetMenuGlyph(properties, "\uE90F");
        SetMenuGlyph(terminal, "\uE756");
        SetMenuGlyph(codeEditor, "\uE943");
        SetMenuGlyph(fileTasksMenu, "\uE8A5");
        SetMenuGlyph(moreMenu, "\uE712");
        SetMenuGlyph(copyPath, "\uE8C8");
        SetMenuGlyph(print, "\uE749");
        SetMenuGlyph(mergePdfs, "\uE8A7");
        SetMenuGlyph(ocrPdfToText, "\uE8D2");
        SetMenuGlyph(mail, "\uE715");
        SetMenuGlyph(whatsappSafeShare, "\uE72D");
        SetMenuGlyph(cut, "\uE8C6");
        SetMenuGlyph(copy, "\uE8C8");
        SetMenuGlyph(paste, "\uE77F");
        SetMenuGlyph(move, "\uE8DE");
        SetMenuGlyph(rename, "\uE8AC");
        SetMenuGlyph(delete, "\uE74D");
        SetMenuGlyph(undoDelete, "\uE7A7");
        SetMenuGlyph(zip, "\uE7B8");
        SetMenuGlyph(extractHere, "\uE8B7");
        SetMenuGlyph(extractNamedFolder, "\uE8B7");
        SetMenuGlyph(extractChooseFolder, "\uE8B7");
        SetMenuGlyph(compressionMenu, "\uE7B8");
        SetMenuGlyph(openExplorerItem, "\uE8B7");
        SetContextMenuTooltips(menu.Items);
        list.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            contextRow = list.HitTest(e.Location).Item;
            if (contextRow != null && !contextRow.Selected)
            {
                list.SelectedItems.Clear();
                contextRow.Selected = true;
                list.FocusedItem = contextRow;
            }
        };
        menu.Opening += (s, e) =>
        {
            if (showingComputerDrives) { e.Cancel = true; return; }
            Color menuBackground = menu.BackColor;
            actionBar.BackColor = menuBackground;
            actionBarHost.BackColor = menuBackground;
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            bool hasTarget = selected != null;
            oneDriveMenu.Visible = hasTarget && IsOneDrivePath(selected.FullPath);
            winRarMenu.Visible = hasTarget && !String.IsNullOrWhiteSpace(FindWinRarExecutable());
            bool isWinRarArchive = hasTarget && File.Exists(selected.FullPath) &&
                (String.Equals(Path.GetExtension(selected.FullPath), ".zip", StringComparison.OrdinalIgnoreCase) || String.Equals(Path.GetExtension(selected.FullPath), ".rar", StringComparison.OrdinalIgnoreCase));
            rarOpen.Enabled = rarHere.Enabled = rarFolder.Enabled = isWinRarArchive;
            actionBarHost.Visible = hasTarget;
            pasteInEmptyArea.Visible = !hasTarget;
            openFile.Visible = openWith.Visible = preview.Visible = clipboardMenu.Visible = manageMenu.Visible = fileTasksMenu.Visible = whatsappSafeShare.Visible = compressionMenu.Visible = hasTarget;
            viewMenu.Visible = sortMenu.Visible = groupMenu.Visible = newMenu.Visible = !hasTarget;
            if (hasTarget && contextRow.ImageIndex >= 0 && list.SmallImageList != null && contextRow.ImageIndex < list.SmallImageList.Images.Count)
                openFile.Image = list.SmallImageList.Images[contextRow.ImageIndex];
            openWith.Enabled = selected != null && File.Exists(selected.FullPath);
            previewPaneItem.Checked = previewPaneOpen;
            previewPaneItem.Enabled = selected != null && File.Exists(selected.FullPath);
            sortName.Checked = fileSortColumn == 0;
            sortDate.Checked = fileSortColumn == 1;
            sortSize.Checked = fileSortColumn == 2;
            sortType.Checked = fileSortColumn == 3;
            groupNone.Checked = fileGroupMode == 0;
            groupType.Checked = fileGroupMode == 1;
            groupDate.Checked = fileGroupMode == 2;
            groupSize.Checked = fileGroupMode == 3;
            openExplorerItem.Enabled = selected != null && (File.Exists(selected.FullPath) || Directory.Exists(selected.FullPath));
            openFile.Enabled = selected != null && File.Exists(selected.FullPath);
            preview.Enabled = selected != null && File.Exists(selected.FullPath);
            print.Enabled = GetSelectedOrCheckedFilePaths().Length > 0;
            mergePdfs.Enabled = GetSelectedPdfPaths().Length >= 2;
            ocrPdfToText.Enabled = GetSelectedPdfPaths().Length > 0 && (pdfOcrWorker == null || !pdfOcrWorker.IsBusy);
            mail.Enabled = selected != null && File.Exists(selected.FullPath);
            whatsappSafeShare.Enabled = selected != null && File.Exists(selected.FullPath);
            copy.Enabled = GetSelectedOrCheckedFileAndFolderPaths().Length > 0;
            cut.Enabled = copy.Enabled;
            move.Enabled = copy.Enabled;
            undoMove.Enabled = lastMoveRecords.Count > 0;
            rename.Enabled = selected != null && File.Exists(selected.FullPath);
            delete.Enabled = GetSelectedOrCheckedFileAndFolderPaths().Length > 0;
            undoDelete.Enabled = lastRecycleBinDeletedPaths.Count > 0;
            properties.Enabled = GetSelectedOrCheckedFileAndFolderPaths().Length > 0;
            terminal.Enabled = !String.IsNullOrWhiteSpace(currentPath) && Directory.Exists(currentPath);
            codeEditor.Enabled = selected == null || File.Exists(selected.FullPath) || Directory.Exists(selected.FullPath);
            copyPath.Enabled = selected != null || !String.IsNullOrWhiteSpace(currentPath);
            zip.Enabled = GetSelectedOrCheckedFilePaths().Length > 0;
            bool isArchive = selected != null && File.Exists(selected.FullPath) &&
                (String.Equals(Path.GetExtension(selected.FullPath), ".zip", StringComparison.OrdinalIgnoreCase) ||
                 String.Equals(Path.GetExtension(selected.FullPath), ".rar", StringComparison.OrdinalIgnoreCase));
            extractHere.Enabled = extractNamedFolder.Enabled = extractChooseFolder.Enabled = isArchive;
            compressionMenu.Enabled = zip.Enabled || isArchive;
            cutButton.Enabled = cut.Enabled;
            copyButton.Enabled = copy.Enabled;
            shareButton.Enabled = selected != null && (File.Exists(selected.FullPath) || Directory.Exists(selected.FullPath));
            PopulateShellNewMenu(newMenu);
            SetContextMenuTooltips(newMenu.DropDownItems);
            try { paste.Enabled = Clipboard.GetFileDropList().Count > 0 ||
                (!String.IsNullOrWhiteSpace(pendingFolderClipboardPath) && Directory.Exists(pendingFolderClipboardPath)); }
            catch { paste.Enabled = !String.IsNullOrWhiteSpace(pendingFolderClipboardPath) && Directory.Exists(pendingFolderClipboardPath); }
            pasteButton.Enabled = paste.Enabled;
            pasteInEmptyArea.Enabled = paste.Enabled;
            deleteButton.Enabled = delete.Enabled;
        };
        pasteInEmptyArea.Click += (s, e) => PasteFiles();
        Action<bool> setCloudAvailability = keep =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected == null || !IsOneDrivePath(selected.FullPath)) return;
            try
            {
                string arguments = (keep ? "+P -U " : "-P +U ") + QuoteProcessArgument(selected.FullPath);
                Process.Start(new ProcessStartInfo("attrib.exe", arguments) { UseShellExecute = false, CreateNoWindow = true });
                status.Text = keep ? "OneDrive çevrimdışı kullanım isteği gönderildi." : "OneDrive yer açma isteği gönderildi.";
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "OneDrive", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        };
        keepOffline.Click += (s, e) => setCloudAvailability(true);
        freeSpace.Click += (s, e) => setCloudAvailability(false);
        rarOpen.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            string exe = FindWinRarExecutable();
            if (selected == null || String.IsNullOrEmpty(exe)) return;
            try { Process.Start(new ProcessStartInfo(exe, QuoteProcessArgument(selected.FullPath)) { UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "WinRAR", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        };
        rarHere.Click += (s, e) => ExtractSelectedArchive(contextRow == null ? null : contextRow.Tag as Entry, ArchiveExtractionTarget.Here);
        rarFolder.Click += (s, e) => ExtractSelectedArchive(contextRow == null ? null : contextRow.Tag as Entry, ArchiveExtractionTarget.ArchiveFolder);
        openWith.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected == null || !File.Exists(selected.FullPath)) return;
            try { Process.Start(new ProcessStartInfo("rundll32.exe", "shell32.dll,OpenAs_RunDLL \"" + selected.FullPath + "\"") { UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Birlikte aç", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        };
        openExplorerItem.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected != null) OpenInWindowsExplorer(selected.FullPath);
        };
        openFile.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected != null && File.Exists(selected.FullPath)) OpenFile(selected.FullPath);
        };
        preview.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected == null || !File.Exists(selected.FullPath)) return;
            if (String.Equals(Path.GetExtension(selected.FullPath), ".xml", StringComparison.OrdinalIgnoreCase))
            {
                if (!TryOpenEdefterXml(selected.FullPath)) new XmlSourceViewerForm(selected.FullPath).Show(this);
                return;
            }
            new FilePreviewForm(selected.FullPath, true).Show(this);
        };
        previewPaneItem.Click += (s, e) => ToggleEmbeddedPreview();
        sortName.Click += (s, e) => SortFiles(0, true);
        sortDate.Click += (s, e) => SortFiles(1, true);
        sortSize.Click += (s, e) => SortFiles(2, true);
        sortType.Click += (s, e) => SortFiles(3, true);
        groupNone.Click += (s, e) => SetFileGrouping(0);
        groupType.Click += (s, e) => SetFileGrouping(1);
        groupDate.Click += (s, e) => SetFileGrouping(2);
        groupSize.Click += (s, e) => SetFileGrouping(3);
        properties.Click += (s, e) =>
        {
            string[] selected = GetSelectedOrCheckedFileAndFolderPaths();
            if (selected.Length == 0 && contextRow != null)
            {
                Entry contextEntry = contextRow.Tag as Entry;
                if (contextEntry != null) selected = new[] { contextEntry.FullPath };
            }
            ShowSelectedItemsProperties(selected);
        };
        terminal.Click += (s, e) => OpenTerminalAt(currentPath);
        codeEditor.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            OpenCodeEditor(selected == null ? currentPath : selected.FullPath);
        };
        copyPath.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            try { Clipboard.SetText(selected == null ? currentPath : selected.FullPath); } catch { }
        };
        print.Click += (s, e) => PrintSelectedFiles();
        mergePdfs.Click += (s, e) => MergeSelectedPdfFiles();
        ocrPdfToText.Click += (s, e) => OcrSelectedPdfFiles();
        mail.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected != null && File.Exists(selected.FullPath)) SendByMail(selected.FullPath);
        };
        whatsappSafeShare.Click += (s, e) =>
        {
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected != null) ShareWithWhatsAppSafeFileName(selected.FullPath);
        };
        copy.Click += (s, e) => CopySelectedExplorerItems(false);
        cut.Click += (s, e) => CopySelectedExplorerItems(true);
        paste.Click += (s, e) => PasteFiles();
        move.Click += (s, e) => MoveSelectedExplorerItems();
        undoMove.Click += (s, e) => UndoLastMove();
        rename.Click += (s, e) => RenameSelectedFile(contextRow == null ? null : contextRow.Tag as Entry);
        delete.Click += (s, e) => DeleteSelectedExplorerFiles();
        undoDelete.Click += (s, e) => UndoLastRecycleBinDelete();
        zip.Click += (s, e) => ZipSelectedFiles();
        extractHere.Click += (s, e) => ExtractSelectedArchive(contextRow == null ? null : contextRow.Tag as Entry, ArchiveExtractionTarget.Here);
        extractNamedFolder.Click += (s, e) => ExtractSelectedArchive(contextRow == null ? null : contextRow.Tag as Entry, ArchiveExtractionTarget.ArchiveFolder);
        extractChooseFolder.Click += (s, e) => ExtractSelectedArchive(contextRow == null ? null : contextRow.Tag as Entry, ArchiveExtractionTarget.ChooseFolder);
        cutButton.Click += (s, e) => { menu.Close(); CopySelectedExplorerItems(true); };
        copyButton.Click += (s, e) => { menu.Close(); CopySelectedExplorerItems(false); };
        pasteButton.Click += (s, e) => { menu.Close(); PasteFiles(); };
        shareButton.Click += (s, e) =>
        {
            menu.Close();
            Entry selected = contextRow == null ? null : contextRow.Tag as Entry;
            if (selected != null) ShareWithWindows(selected.FullPath);
        };
        deleteButton.Click += (s, e) => { menu.Close(); DeleteSelectedExplorerFiles(); };
        list.ContextMenuStrip = menu;
    }

    private Button CreateExplorerMenuActionButton(string text, string glyph, int width)
    {
        Color hoverColor = Color.FromArgb(221, 239, 232);
        var image = new Bitmap(18, 18);
        using (Graphics graphics = Graphics.FromImage(image))
        using (var font = new Font("Segoe MDL2 Assets", 13F, FontStyle.Regular, GraphicsUnit.Point))
        using (var brush = new SolidBrush(GetMenuGlyphColor(glyph)))
            graphics.DrawString(glyph, font, brush, new RectangleF(0, 0, 18, 18),
                new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
        var button = new Button
        {
            Text = text,
            Image = image,
            TextImageRelation = TextImageRelation.ImageAboveText,
            TextAlign = ContentAlignment.MiddleCenter,
            ImageAlign = ContentAlignment.MiddleCenter,
            Size = new Size(width, 42),
            Margin = new Padding(0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            UseVisualStyleBackColor = false,
            Font = new Font("Segoe UI", 7.5F)
        };
        button.FlatAppearance.BorderSize = 0;
        button.FlatAppearance.MouseOverBackColor = hoverColor;
        button.MouseEnter += (s, e) => { button.BackColor = hoverColor; button.ForeColor = SystemColors.MenuText; };
        button.MouseLeave += (s, e) => { button.BackColor = Color.White; button.ForeColor = SystemColors.MenuText; };
        button.Disposed += (s, e) => { if (button.Image != null) { button.Image.Dispose(); button.Image = null; } };
        return button;
    }

    private Control CreateExplorerMenuSeparator()
    {
        return new Panel
        {
            Width = 1,
            Height = 36,
            Margin = new Padding(2, 3, 2, 3),
            BackColor = Color.FromArgb(231, 233, 236)
        };
    }

    private void SetContextMenuTooltips(ToolStripItemCollection items)
    {
        var descriptions = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase)
        {
            { "Pano ve taşıma", "Kesme, kopyalama, yapıştırma ve taşıma komutları." },
            { "Kes", "Seçili öğeleri kesip hedef klasöre taşımak üzere panoya al." },
            { "Kopyala", "Seçili öğeleri panoya kopyala." },
            { "Yapıştır", "Panodaki öğeleri bu klasöre yapıştır." },
            { "Taşı...", "Seçili öğeleri başka bir klasöre taşı." },
            { "Dosyayı Aç", "Seçili dosyayı varsayılan uygulamasında aç." },
            { "Ön izleme", "Dosyayı açmadan ön izleme penceresinde görüntüle." },
            { "Görünüm", "Dosya listesinin görünüm ve ön izleme seçenekleri." },
            { "Ayrıntılar", "Dosyaları ayrıntı sütunlarıyla göster." },
            { "Ön izleme paneli", "Seçili dosyanın sağdaki ön izleme panelinde gösterimini aç veya kapat." },
            { "Sıralama ölçütü", "Dosyaları seçilen sütuna göre sırala." },
            { "Ad", "Dosyaları ada göre sırala." },
            { "Değiştirme tarihi", "Dosyaları son değiştirilme tarihine göre sırala." },
            { "Boyut", "Dosyaları boyuta göre sırala." },
            { "Tür", "Dosyaları dosya türüne göre sırala." },
            { "Gruplandırma ölçütü", "Dosyaları tür, tarih veya boyut grupları altında topla." },
            { "Gruplandırma yok", "Dosyaları gruplara ayırmadan listele." },
            { "Dosya türü", "Dosyaları uzantı ve türe göre grupla." },
            { "Değiştirme tarihine göre", "Dosyaları değiştirilme dönemine göre grupla." },
            { "Boyuta göre", "Dosyaları boyut aralıklarına göre grupla." },
            { "Geri Al Taşı", "Son taşıma işlemini geri al." },
            { "Yeni", "Bu klasörde yeni bir klasör veya dosya oluştur." },
            { "Düzenle ve yönet", "Yeniden adlandırma ve silme komutları." },
            { "Yeniden adlandır", "Seçili dosyanın adını değiştir." },
            { "Sil", "Seçili öğeleri Geri Dönüşüm Kutusu'na taşı." },
            { "Geri Al Silme", "Son Geri Dönüşüm Kutusu silme işlemini geri al." },
            { "Özellikler", "Seçili dosya veya klasörün Windows özelliklerini aç." },
            { "Terminalde Aç", "Geçerli klasörde bir terminal penceresi aç." },
            { "Code ile Aç", "Seçili öğeyi veya klasörü Visual Studio Code'da aç." },
            { "Dosya işlemleri", "Yazdırma, PDF, OCR ve mail işlemleri." },
            { "Seçili dosyaları yazdır", "Seçili dosyaları varsayılan yazdırma uygulamasına gönder." },
            { "Seçili PDF'leri birleştir...", "Seçili PDF dosyalarını tek bir PDF'te birleştir." },
            { "OCR ile TXT oluştur", "PDF'deki metni OCR ile tanıyıp TXT dosyasına kaydet." },
            { "Mail Gönder", "Seçili dosyayı varsayılan mail uygulamasına ekle." },
            { "WhatsApp için karakter uyumlu paylaş", "Özgün dosyayı değiştirmeden, Türkçe harfleri ASCII karşılıklarına çevrilmiş geçici kopyayı paylaş. Durum çubuğu gönderilen geçici dosyanın adını gösterir. 2 günden eski kopyalar sonraki kullanımda temizlenir." },
            { "Sıkıştırma", "Dosyaları ZIP'le veya arşivleri ayıkla." },
            { "ZIP ile sıkıştır", "Seçili dosyaları ZIP arşivine ekle." },
            { "Arşivi buraya ayıkla", "Arşiv içeriğini bu klasöre çıkar." },
            { "Arşivi kendi klasörüne ayıkla", "Arşivi adıyla oluşturulan klasöre çıkar." },
            { "Arşivi seçilen klasöre ayıkla...", "Arşiv için hedef klasör seç." },
            { "Diğer seçenekler", "Windows Explorer ve yol kopyalama seçenekleri." },
            { "Windows Explorer ile Aç", "Seçili öğeyi Windows Explorer'da aç veya konumunu göster." },
            { "Yolu kopyala", "Seçili dosya veya klasörün tam yolunu panoya kopyala." }
        };
        foreach (ToolStripItem item in items)
        {
            var menuItem = item as ToolStripMenuItem;
            if (menuItem == null) continue;
            string description;
            menuItem.ToolTipText = descriptions.TryGetValue(menuItem.Text, out description)
                ? description : menuItem.Text + " seçeneğini kullan.";
            if (menuItem.DropDownItems.Count > 0) SetContextMenuTooltips(menuItem.DropDownItems);
        }
    }

    private void ShowItemProperties(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return;
        try
        {
            if (!SHObjectProperties(Handle, 2, path, null))
                Process.Start(new ProcessStartInfo(path) { Verb = "properties", UseShellExecute = true });
        }
        catch
        {
            try { Process.Start(new ProcessStartInfo(path) { Verb = "properties", UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Özellikler açılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }

    private void ShowSelectedItemsProperties(IEnumerable<string> selectedPaths)
    {
        string[] paths = (selectedPaths ?? Enumerable.Empty<string>())
            .Where(path => !String.IsNullOrWhiteSpace(path) && (File.Exists(path) || Directory.Exists(path)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (paths.Length == 0) return;
        foreach (IGrouping<string, string> group in paths.GroupBy(path => Path.GetDirectoryName(path) ?? String.Empty, StringComparer.OrdinalIgnoreCase))
        {
            string[] groupPaths = group.ToArray();
            if (groupPaths.Length == 1 || !TryShowShellPropertiesForSelection(groupPaths))
                foreach (string path in groupPaths) ShowItemProperties(path);
        }
    }

    private bool TryShowShellPropertiesForSelection(string[] paths)
    {
        if (paths == null || paths.Length < 2) return false;
        string parentPath = Path.GetDirectoryName(paths[0]);
        if (String.IsNullOrWhiteSpace(parentPath) || paths.Any(path => !String.Equals(Path.GetDirectoryName(path), parentPath, StringComparison.OrdinalIgnoreCase))) return false;

        Guid shellFolderId = new Guid("000214E6-0000-0000-C000-000000000046");
        Guid contextMenuId = new Guid("000214E4-0000-0000-C000-000000000046");
        var absoluteItemIds = new List<IntPtr>();
        var childItemIds = new List<IntPtr>();
        IntPtr parentFolderPointer = IntPtr.Zero;
        IntPtr contextMenuPointer = IntPtr.Zero;
        IntPtr itemIdArray = IntPtr.Zero;
        IntPtr nativeMenu = IntPtr.Zero;
        IShellFolderNative shellFolder = null;
        IShellContextMenuNative contextMenu = null;
        try
        {
            for (int i = 0; i < paths.Length; i++)
            {
                IntPtr absoluteItemId;
                uint attributes;
                int parseResult = SHParseDisplayName(paths[i], IntPtr.Zero, out absoluteItemId, 0, out attributes);
                if (absoluteItemId != IntPtr.Zero) absoluteItemIds.Add(absoluteItemId);
                if (parseResult < 0 || absoluteItemId == IntPtr.Zero) return false;

                IntPtr itemParent;
                IntPtr childItemId;
                int bindResult = SHBindToParent(absoluteItemId, ref shellFolderId, out itemParent, out childItemId);
                if (bindResult < 0 || itemParent == IntPtr.Zero || childItemId == IntPtr.Zero)
                {
                    if (itemParent != IntPtr.Zero) Marshal.Release(itemParent);
                    return false;
                }
                childItemIds.Add(childItemId);
                if (i == 0)
                {
                    parentFolderPointer = itemParent;
                    shellFolder = Marshal.GetObjectForIUnknown(parentFolderPointer) as IShellFolderNative;
                    if (shellFolder == null) return false;
                }
                else Marshal.Release(itemParent);
            }

            itemIdArray = Marshal.AllocHGlobal(IntPtr.Size * childItemIds.Count);
            for (int i = 0; i < childItemIds.Count; i++)
                Marshal.WriteIntPtr(itemIdArray, i * IntPtr.Size, childItemIds[i]);

            int uiObjectResult = shellFolder.GetUIObjectOf(Handle, (uint)childItemIds.Count, itemIdArray,
                ref contextMenuId, IntPtr.Zero, out contextMenuPointer);
            if (uiObjectResult < 0 || contextMenuPointer == IntPtr.Zero) return false;
            contextMenu = Marshal.GetObjectForIUnknown(contextMenuPointer) as IShellContextMenuNative;
            if (contextMenu == null) return false;

            nativeMenu = CreatePopupMenu();
            if (nativeMenu == IntPtr.Zero) return false;
            int menuResult = contextMenu.QueryContextMenu(nativeMenu, 0, 1, 0x7FFF, 0);
            if (menuResult < 0) return false;
            int commandCount = menuResult & 0xFFFF;
            for (int commandOffset = 0; commandOffset < commandCount; commandOffset++)
            {
                IntPtr verbBuffer = Marshal.AllocHGlobal(256);
                string verb;
                int verbResult;
                try
                {
                    Marshal.WriteInt16(verbBuffer, 0);
                    verbResult = contextMenu.GetCommandString(new IntPtr(commandOffset), 4, IntPtr.Zero, verbBuffer, 128);
                    verb = Marshal.PtrToStringUni(verbBuffer) ?? String.Empty;
                }
                finally { Marshal.FreeHGlobal(verbBuffer); }
                if (verbResult < 0 || !String.Equals(verb, "properties", StringComparison.OrdinalIgnoreCase)) continue;
                var invoke = new ShellInvokeCommandInfo
                {
                    Size = (uint)Marshal.SizeOf(typeof(ShellInvokeCommandInfo)),
                    Window = Handle,
                    Verb = new IntPtr(commandOffset),
                    Show = 1
                };
                return contextMenu.InvokeCommand(ref invoke) >= 0;
            }
            return false;
        }
        catch { return false; }
        finally
        {
            if (nativeMenu != IntPtr.Zero) DestroyMenu(nativeMenu);
            if (itemIdArray != IntPtr.Zero) Marshal.FreeHGlobal(itemIdArray);
            foreach (IntPtr itemId in absoluteItemIds)
                if (itemId != IntPtr.Zero) Marshal.FreeCoTaskMem(itemId);
            if (contextMenuPointer != IntPtr.Zero) Marshal.Release(contextMenuPointer);
            if (parentFolderPointer != IntPtr.Zero) Marshal.Release(parentFolderPointer);
            try { if (contextMenu != null && Marshal.IsComObject(contextMenu)) Marshal.ReleaseComObject(contextMenu); } catch { }
            try { if (shellFolder != null && Marshal.IsComObject(shellFolder)) Marshal.ReleaseComObject(shellFolder); } catch { }
        }
    }

    private void OpenTerminalAt(string path)
    {
        if (String.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;
        try
        {
            Process.Start(new ProcessStartInfo("wt.exe", "-d \"" + path + "\"")
            {
                UseShellExecute = true,
                WorkingDirectory = path
            });
        }
        catch
        {
            try
            {
                Process.Start(new ProcessStartInfo("cmd.exe", "/K cd /d \"" + path + "\"")
                {
                    UseShellExecute = true,
                    WorkingDirectory = path
                });
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Terminal açılamadı", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }

    private void OpenCodeEditor(string path)
    {
        if (String.IsNullOrWhiteSpace(path)) return;
        try
        {
            Process.Start(new ProcessStartInfo("code", "\"" + path.Replace("\"", "\\\"") + "\"")
            {
                UseShellExecute = true,
                WorkingDirectory = Directory.Exists(path) ? path : Path.GetDirectoryName(path)
            });
        }
        catch
        {
            MessageBox.Show(this, "Visual Studio Code bulunamadı. Code komutunun PATH içinde olduğundan emin olun.",
                "Code ile Aç", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private static void SetMenuGlyph(ToolStripMenuItem item, string glyph)
    {
        var image = new Bitmap(18, 18);
        using (Graphics graphics = Graphics.FromImage(image))
        using (var font = new Font("Segoe MDL2 Assets", 12F, FontStyle.Regular, GraphicsUnit.Point))
        using (var brush = new SolidBrush(GetMenuGlyphColor(glyph)))
            graphics.DrawString(glyph, font, brush, new RectangleF(0, 0, 18, 18),
                new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
        item.Image = image;
        item.ImageScaling = ToolStripItemImageScaling.None;
        item.ImageAlign = ContentAlignment.MiddleCenter;
        item.Padding = new Padding(0, 0, 4, 0);
        item.Disposed += (s, e) => { if (item.Image != null) { item.Image.Dispose(); item.Image = null; } };
    }

    private static Color GetMenuGlyphColor(string glyph)
    {
        switch (glyph)
        {
            case "\uE8C6": return Color.FromArgb(224, 125, 42); // Kes
            case "\uE8C8": return Color.FromArgb(44, 119, 190); // Kopyala
            case "\uE8DE": return Color.FromArgb(126, 87, 194); // Taşı
            case "\uE77F": return Color.FromArgb(35, 145, 125); // Yapıştır
            case "\uE74D": return Color.FromArgb(204, 63, 72); // Sil
            case "\uE7A7": return Color.FromArgb(77, 133, 154); // Geri al
            case "\uE715": return Color.FromArgb(190, 77, 132); // Mail
            case "\uE749": return Color.FromArgb(67, 101, 154); // Yazdır
            case "\uE8A7": return Color.FromArgb(130, 93, 178); // Aç / PDF
            case "\uE8A5": return Color.FromArgb(41, 139, 91); // Dosya aç
            case "\uE72D": return Color.FromArgb(0, 120, 215); // Paylaş
            case "\uE8B7": return Color.FromArgb(207, 151, 46); // Klasör
            case "\uE890": return Color.FromArgb(40, 145, 158); // Ön izleme
            case "\uE8AC": return Color.FromArgb(186, 121, 48); // Yeniden adlandır
            case "\uE7B8": return Color.FromArgb(112, 126, 143); // Sıkıştırma
            case "\uE710": return Color.FromArgb(50, 132, 177); // Yeni
            default: return Color.FromArgb(69, 91, 112);
        }
    }

    private void PrintSelectedFiles()
    {
        string[] paths = GetSelectedOrCheckedFilePaths().Where(File.Exists).ToArray();
        if (paths.Length == 0) return;
        if (paths.Length > 1)
        {
            DialogResult answer = MessageBox.Show(this,
                paths.Length + " seçili dosya Windows'un varsayılan yazdırma işlemiyle yazdırılsın mı?",
                "Dosyaları yazdır", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (answer != DialogResult.Yes) return;
        }

        int started = 0;
        var failures = new List<string>();
        foreach (string path in paths)
        {
            try
            {
                var start = new ProcessStartInfo
                {
                    FileName = path,
                    Verb = "print",
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process process = Process.Start(start);
                if (process == null) failures.Add(Path.GetFileName(path) + " (yazdırma işlemi başlatılamadı)");
                else { started++; process.Dispose(); }
            }
            catch (Exception ex)
            {
                failures.Add(Path.GetFileName(path) + " — " + ex.Message);
            }
        }
        status.Text = started + " dosya yazdırma işlemine gönderildi.";
        if (failures.Count > 0)
            MessageBox.Show(this, "Bazı dosyalar için Windows yazdırma işlemi başlatılamadı:\r\n\r\n" + String.Join("\r\n", failures),
                "Yazdırma", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private string[] GetSelectedPdfPaths()
    {
        return GetSelectedOrCheckedFilePaths()
            .Where(path => String.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
            .ToArray();
    }

    private void MergeSelectedPdfFiles()
    {
        string[] sources = GetSelectedPdfPaths();
        if (sources.Length < 2)
        {
            MessageBox.Show(this, "Birleştirmek için en az iki PDF dosyası seçin.", "PDF'leri birleştir",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using (var dialog = new SaveFileDialog())
        {
            dialog.Title = "Birleştirilmiş PDF dosyasını kaydet";
            dialog.Filter = "PDF dosyası (*.pdf)|*.pdf";
            dialog.DefaultExt = "pdf";
            dialog.AddExtension = true;
            dialog.FileName = "Birlesik.pdf";
            dialog.InitialDirectory = Path.GetDirectoryName(sources[0]);
            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            string destination = Path.GetFullPath(dialog.FileName);
            if (sources.Any(path => String.Equals(Path.GetFullPath(path), destination, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(this, "Çıktı dosyası kaynak PDF'lerden biriyle aynı olamaz. Başka bir dosya adı seçin.",
                    "PDF'leri birleştir", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string temporaryOutput = destination + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var mergedDocument = new PdfDocument())
                {
                    foreach (string sourcePath in sources)
                    {
                        PdfDocument sourceDocument = null;
                        try
                        {
                            sourceDocument = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Import);
                            for (int pageIndex = 0; pageIndex < sourceDocument.PageCount; pageIndex++)
                                mergedDocument.AddPage(sourceDocument.Pages[pageIndex]);
                        }
                        finally
                        {
                            if (sourceDocument != null) sourceDocument.Close();
                        }
                    }
                    if (mergedDocument.PageCount == 0)
                        throw new InvalidOperationException("Seçilen PDF dosyalarında birleştirilecek sayfa bulunamadı.");
                    mergedDocument.Save(temporaryOutput);
                }

                if (File.Exists(destination)) File.Replace(temporaryOutput, destination, null);
                else File.Move(temporaryOutput, destination);

                MessageBox.Show(this, sources.Length + " PDF dosyası, seçili listedeki sırayla birleştirildi.\r\n\r\n" + destination,
                    "PDF'leri birleştir", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshFileRows();
            }
            catch (Exception ex)
            {
                Exception cause = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                MessageBox.Show(this, "PDF dosyaları birleştirilemedi. Dosya parola korumalı, bozuk veya bu PDF sürümü desteklenmiyor olabilir.\r\n\r\n" + cause.Message,
                    "PDF'leri birleştir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try { if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput); } catch { }
            }
        }
    }

    private void OcrSelectedPdfFiles()
    {
        string[] sources = GetSelectedPdfPaths();
        if (sources.Length == 0) return;
        string[] outputs = sources.Select(path => Path.ChangeExtension(path, ".txt")).ToArray();
        string[] existingOutputs = outputs.Where(File.Exists).ToArray();
        if (existingOutputs.Length > 0)
        {
            DialogResult overwrite = MessageBox.Show(this,
                existingOutputs.Length + " TXT dosyası zaten var. Üzerlerine yazılsın mı?",
                "OCR metin dosyası", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (overwrite != DialogResult.Yes) return;
        }

        pdfOcrWorker = new BackgroundWorker { WorkerReportsProgress = true };
        pdfOcrWorker.DoWork += (s, e) =>
        {
            string[] inputPaths = (string[])e.Argument;
            var created = new List<string>();
            for (int i = 0; i < inputPaths.Length; i++)
            {
                ((BackgroundWorker)s).ReportProgress((i * 100) / inputPaths.Length,
                    "OCR: " + Path.GetFileName(inputPaths[i]) + " (" + (i + 1) + "/" + inputPaths.Length + ")");
                string output = Path.ChangeExtension(inputPaths[i], ".txt");
                string recognizedText = RecognizePdfWithWindowsOcr(inputPaths[i]);
                File.WriteAllText(output, recognizedText, new UTF8Encoding(true));
                created.Add(output);
            }
            e.Result = created.ToArray();
        };
        pdfOcrWorker.ProgressChanged += (s, e) => status.Text = Convert.ToString(e.UserState, CultureInfo.CurrentCulture);
        pdfOcrWorker.RunWorkerCompleted += (s, e) =>
        {
            BackgroundWorker worker = pdfOcrWorker;
            pdfOcrWorker = null;
            if (worker != null) worker.Dispose();
            if (e.Error != null)
            {
                status.Text = "PDF OCR işlemi tamamlanamadı.";
                MessageBox.Show(this, "PDF'den metin çıkarılamadı.\r\n\r\n" + e.Error.GetBaseException().Message,
                    "PDF OCR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string[] results = (string[])e.Result;
            status.Text = results.Length + " PDF için OCR metni oluşturuldu.";
            RefreshFileRows();
            MessageBox.Show(this, results.Length + " TXT dosyası oluşturuldu:\r\n\r\n" + String.Join("\r\n", results),
                "PDF OCR tamamlandı", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        status.Text = "PDF OCR başlatılıyor...";
        pdfOcrWorker.RunWorkerAsync(sources);
    }

    private static string RecognizePdfWithWindowsOcr(string pdfPath)
    {
        object storageFile = null;
        object pdfDocument = null;
        object ocrEngine = null;
        var text = new StringBuilder();
        try
        {
            Type storageFileType = GetWindowsRuntimeType("Windows.Storage.StorageFile");
            Type pdfDocumentType = GetWindowsRuntimeType("Windows.Data.Pdf.PdfDocument");
            Type ocrEngineType = GetWindowsRuntimeType("Windows.Media.Ocr.OcrEngine");
            Type languageType = GetWindowsRuntimeType("Windows.Globalization.Language");
            storageFile = WaitForWindowsRuntimeOperation(InvokeWindowsRuntimeStatic(storageFileType, "GetFileFromPathAsync", pdfPath));
            pdfDocument = WaitForWindowsRuntimeOperation(InvokeWindowsRuntimeStatic(pdfDocumentType, "LoadFromFileAsync", storageFile));

            try
            {
                object language = Activator.CreateInstance(languageType, new object[] { "tr-TR" });
                ocrEngine = InvokeWindowsRuntimeStatic(ocrEngineType, "TryCreateFromLanguage", language);
            }
            catch { }
            if (ocrEngine == null)
                ocrEngine = InvokeWindowsRuntimeStatic(ocrEngineType, "TryCreateFromUserProfileLanguages");
            if (ocrEngine == null)
                throw new InvalidOperationException("Bu bilgisayarda kullanılabilir bir Windows OCR dili yüklü değil. Windows Ayarları > Zaman ve Dil > Dil ve Bölge bölümünden Türkçe için OCR/dil bileşenini yükleyip yeniden deneyin.");

            string languageTag = "";
            try
            {
                object recognizerLanguage = GetWindowsRuntimeProperty(ocrEngine, "RecognizerLanguage");
                languageTag = Convert.ToString(GetWindowsRuntimeProperty(recognizerLanguage, "LanguageTag"), CultureInfo.InvariantCulture);
            }
            catch { }

            int pageCount = Convert.ToInt32(GetWindowsRuntimeProperty(pdfDocument, "PageCount"), CultureInfo.InvariantCulture);
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                object page = null;
                object stream = null;
                object decoder = null;
                object softwareBitmap = null;
                try
                {
                    page = InvokeWindowsRuntimeMethod(pdfDocument, "GetPage", (uint)pageIndex);
                    Type streamType = GetWindowsRuntimeType("Windows.Storage.Streams.InMemoryRandomAccessStream");
                    stream = Activator.CreateInstance(streamType);
                    Type renderOptionsType = GetWindowsRuntimeType("Windows.Data.Pdf.PdfPageRenderOptions");
                    object renderOptions = Activator.CreateInstance(renderOptionsType);
                    object pageSize = GetWindowsRuntimeProperty(page, "Size");
                    double pageWidth = Convert.ToDouble(GetWindowsRuntimeProperty(pageSize, "Width"), CultureInfo.InvariantCulture);
                    double pageHeight = Convert.ToDouble(GetWindowsRuntimeProperty(pageSize, "Height"), CultureInfo.InvariantCulture);
                    double scale = Math.Min(2.0, 2400.0 / Math.Max(pageWidth, pageHeight));
                    scale = Math.Max(1.0, scale);
                    renderOptionsType.GetProperty("DestinationWidth").SetValue(renderOptions, (uint)Math.Max(1, Math.Round(pageWidth * scale)), null);
                    renderOptionsType.GetProperty("DestinationHeight").SetValue(renderOptions, (uint)Math.Max(1, Math.Round(pageHeight * scale)), null);

                    object renderOperation = InvokeWindowsRuntimeMethod(page, "RenderToStreamAsync", stream, renderOptions);
                    WaitForWindowsRuntimeOperation(renderOperation);
                    InvokeWindowsRuntimeMethod(stream, "Seek", (ulong)0);
                    Type decoderType = GetWindowsRuntimeType("Windows.Graphics.Imaging.BitmapDecoder");
                    decoder = WaitForWindowsRuntimeOperation(InvokeWindowsRuntimeStatic(decoderType, "CreateAsync", stream));
                    softwareBitmap = WaitForWindowsRuntimeOperation(InvokeWindowsRuntimeMethod(decoder, "GetSoftwareBitmapAsync"));
                    object result = WaitForWindowsRuntimeOperation(InvokeWindowsRuntimeMethod(ocrEngine, "RecognizeAsync", softwareBitmap));
                    string pageText = Convert.ToString(GetWindowsRuntimeProperty(result, "Text"), CultureInfo.InvariantCulture);
                    if (pageIndex > 0) text.Append("\r\n\f\r\n");
                    text.Append("----- Sayfa ").Append(pageIndex + 1).Append(" -----\r\n");
                    text.Append(pageText ?? "");
                    CloseWindowsRuntimeObject(result);
                }
                finally
                {
                    CloseWindowsRuntimeObject(softwareBitmap);
                    CloseWindowsRuntimeObject(decoder);
                    CloseWindowsRuntimeObject(stream);
                    CloseWindowsRuntimeObject(page);
                }
            }

            if (pageCount == 0) throw new InvalidOperationException("PDF dosyasında sayfa bulunamadı.");
            if (!String.IsNullOrEmpty(languageTag)) text.Insert(0, "OCR dili: " + languageTag + "\r\n\r\n");
            return text.ToString();
        }
        catch (Exception ex)
        {
            Exception cause = ex;
            while (cause is TargetInvocationException || cause is AggregateException)
            {
                if (cause is TargetInvocationException && cause.InnerException != null) cause = cause.InnerException;
                else if (cause is AggregateException && ((AggregateException)cause).Flatten().InnerExceptions.Count > 0)
                    cause = ((AggregateException)cause).Flatten().InnerExceptions[0];
                else break;
            }
            throw new InvalidOperationException("Windows PDF/OCR bileşeni işlemi tamamlayamadı: " + cause.Message, cause);
        }
        finally
        {
            CloseWindowsRuntimeObject(ocrEngine);
            CloseWindowsRuntimeObject(pdfDocument);
            CloseWindowsRuntimeObject(storageFile);
        }
    }

    private static Type GetWindowsRuntimeType(string typeName)
    {
        Type type = Type.GetType(typeName + ", Windows, ContentType=WindowsRuntime", false);
        if (type == null)
            throw new PlatformNotSupportedException("Windows'un PDF ve OCR altyapısı bulunamadı. Bu özellik Windows 10 veya Windows 11 gerektirir.");
        return type;
    }

    private static object InvokeWindowsRuntimeStatic(Type type, string methodName, params object[] arguments)
    {
        return InvokeWindowsRuntimeMethodCore(type, null, methodName, arguments);
    }

    private static object InvokeWindowsRuntimeMethod(object target, string methodName, params object[] arguments)
    {
        if (target == null) throw new InvalidOperationException("Windows OCR işlemi için beklenen nesne oluşturulamadı.");
        return InvokeWindowsRuntimeMethodCore(target.GetType(), target, methodName, arguments);
    }

    private static object InvokeWindowsRuntimeMethodCore(Type type, object target, string methodName, object[] arguments)
    {
        BindingFlags flags = BindingFlags.Public | (target == null ? BindingFlags.Static : BindingFlags.Instance);
        foreach (MethodInfo method in type.GetMethods(flags).Where(candidate => candidate.Name == methodName && candidate.GetParameters().Length == arguments.Length))
        {
            ParameterInfo[] parameters = method.GetParameters();
            bool compatible = true;
            for (int i = 0; i < parameters.Length; i++)
            {
                object value = arguments[i];
                Type expected = parameters[i].ParameterType;
                if (value != null && !expected.IsInstanceOfType(value) && !(expected.IsValueType && value.GetType() == expected))
                { compatible = false; break; }
            }
            if (compatible) return method.Invoke(target, arguments);
        }
        throw new MissingMethodException(type.FullName, methodName);
    }

    private static object GetWindowsRuntimeProperty(object target, string propertyName)
    {
        if (target == null) throw new InvalidOperationException("Windows OCR işlemi beklenen bilgiyi döndürmedi.");
        PropertyInfo property = target.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
        if (property == null) throw new MissingMemberException(target.GetType().FullName, propertyName);
        return property.GetValue(target, null);
    }

    private static object WaitForWindowsRuntimeOperation(object operation)
    {
        if (operation == null) throw new InvalidOperationException("Windows PDF/OCR işlemi başlatılamadı.");
        Type operationType = operation.GetType().GetInterfaces().FirstOrDefault(item =>
            item.FullName != null && (item.FullName.StartsWith("Windows.Foundation.IAsyncOperation`1", StringComparison.Ordinal) ||
                                      item.FullName == "Windows.Foundation.IAsyncAction"));
        if (operationType == null) throw new InvalidOperationException("Windows OCR işleminin sonucu alınamadı.");
        Type extensions = Type.GetType("System.WindowsRuntimeSystemExtensions, System.Runtime.WindowsRuntime", true);
        MethodInfo asTask = null;
        foreach (MethodInfo candidate in extensions.GetMethods(BindingFlags.Public | BindingFlags.Static).Where(item => item.Name == "AsTask" && item.GetParameters().Length == 1))
        {
            Type parameterType = candidate.GetParameters()[0].ParameterType;
            if (operationType.IsGenericType && candidate.IsGenericMethodDefinition && parameterType.IsGenericType &&
                parameterType.GetGenericTypeDefinition().FullName == "Windows.Foundation.IAsyncOperation`1")
            { asTask = candidate.MakeGenericMethod(operationType.GetGenericArguments()[0]); break; }
            if (!operationType.IsGenericType && !candidate.IsGenericMethod && parameterType.FullName == "Windows.Foundation.IAsyncAction")
            { asTask = candidate; break; }
        }
        if (asTask == null) throw new InvalidOperationException("Windows OCR için gerekli Windows Runtime desteği bu bilgisayarda bulunamadı.");
        Task task = (Task)asTask.Invoke(null, new object[] { operation });
        task.Wait();
        PropertyInfo result = task.GetType().GetProperty("Result", BindingFlags.Public | BindingFlags.Instance);
        return result == null ? null : result.GetValue(task, null);
    }

    private static void CloseWindowsRuntimeObject(object value)
    {
        if (value == null) return;
        try
        {
            MethodInfo close = value.GetType().GetMethod("Close", BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
            if (close != null) close.Invoke(value, null);
        }
        catch { }
    }

    private void DeleteSelectedExplorerFiles()
    {
        string[] paths = GetSelectedOrCheckedFileAndFolderPaths();
        if (paths.Length == 0) return;
        string[] folderParents = paths.Where(Directory.Exists).Select(Path.GetDirectoryName)
            .Where(parent => !String.IsNullOrWhiteSpace(parent)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        DialogResult answer = MessageBox.Show(this, paths.Length + " seçili dosya/klasör Geri Dönüşüm Kutusu'na taşınsın mı?", "Silme onayı",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes) return;
        int removed = 0;
        try
        {
            MovePathsToRecycleBin(paths);
            removed = paths.Length;
            lastRecycleBinDeletedPaths.Clear();
            lastRecycleBinDeletedPaths.AddRange(paths);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Silinemedi", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        Navigate(currentPath, null, null, false);
        foreach (string parent in folderParents)
            if (Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
        status.Text = removed + " öğe Geri Dönüşüm Kutusu'na taşındı.";
    }

    private static void MovePathsToRecycleBin(IEnumerable<string> paths)
    {
        string[] existing = paths.Where(path => File.Exists(path) || Directory.Exists(path)).ToArray();
        if (existing.Length == 0) return;
        var operation = new ShellFileOperation
        {
            wFunc = 3,
            pFrom = String.Join("\0", existing) + "\0\0",
            fFlags = (ushort)(FileOperationAllowUndo | FileOperationSilent | FileOperationNoConfirmation)
        };
        int result = SHFileOperation(ref operation);
        if (result != 0) throw new Win32Exception(result, "Dosyalar Geri Dönüşüm Kutusu'na taşınamadı.");
        if (operation.fAnyOperationsAborted) throw new IOException("Silme işlemi Windows tarafından iptal edildi.");
    }

    private void UndoLastRecycleBinDelete()
    {
        if (lastRecycleBinDeletedPaths.Count == 0) return;
        var restored = new List<string>();
        object shell = null, recycleBin = null, items = null;
        try
        {
            Type shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null) throw new InvalidOperationException("Windows Shell erişilebilir değil.");
            shell = Activator.CreateInstance(shellType);
            recycleBin = InvokeShellMember(shell, "NameSpace", BindingFlags.InvokeMethod, new object[] { 10 });
            if (recycleBin == null) throw new InvalidOperationException("Geri Dönüşüm Kutusu açılamadı.");
            items = InvokeShellMember(recycleBin, "Items", BindingFlags.GetProperty, new object[0]);
            int count = Convert.ToInt32(InvokeShellMember(items, "Count", BindingFlags.GetProperty, new object[0]), CultureInfo.InvariantCulture);
            var remaining = new List<string>(lastRecycleBinDeletedPaths);
            for (int index = 0; index < count && remaining.Count > 0; index++)
            {
                object item = null;
                try
                {
                    item = InvokeShellMember(items, "Item", BindingFlags.InvokeMethod, new object[] { index });
                    string itemName = Convert.ToString(InvokeShellMember(item, "Name", BindingFlags.GetProperty, new object[0]), CultureInfo.InvariantCulture);
                    string deletedFrom = Convert.ToString(InvokeShellMember(item, "ExtendedProperty", BindingFlags.InvokeMethod,
                        new object[] { "System.Recycle.DeletedFrom" }), CultureInfo.InvariantCulture);
                    if (String.IsNullOrWhiteSpace(itemName) || String.IsNullOrWhiteSpace(deletedFrom)) continue;
                    string originalPath = String.Equals(Path.GetFileName(deletedFrom.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)), itemName,
                        StringComparison.OrdinalIgnoreCase) ? deletedFrom : Path.Combine(deletedFrom, itemName);
                    string requestedPath = remaining.FirstOrDefault(path => String.Equals(path, originalPath, StringComparison.OrdinalIgnoreCase));
                    if (requestedPath == null) continue;
                    if (!RestoreRecycleBinItem(item)) continue;
                    restored.Add(requestedPath);
                    remaining.Remove(requestedPath);
                }
                catch { }
                finally { ReleaseShellObject(item); }
            }
            lastRecycleBinDeletedPaths.RemoveAll(path => restored.Contains(path, StringComparer.OrdinalIgnoreCase));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Geri Dönüşüm Kutusu'ndaki öğe geri yüklenemedi.\r\n\r\n" + ex.Message,
                "Geri Al Silme", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        finally
        {
            ReleaseShellObject(items);
            ReleaseShellObject(recycleBin);
            ReleaseShellObject(shell);
        }

        if (restored.Count == 0)
        {
            MessageBox.Show(this, "Silinen öğeler Geri Dönüşüm Kutusu'nda bulunamadı. Kutuyu boşalttıysanız veya öğeler taşındıysa geri yükleme yapılamaz.",
                "Geri Al Silme", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        string[] parents = restored.Select(Path.GetDirectoryName).Where(path => !String.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        Navigate(currentPath, null, restored.FirstOrDefault(path => String.Equals(Path.GetDirectoryName(path), currentPath, StringComparison.OrdinalIgnoreCase)), false);
        foreach (string parent in parents)
            if (Directory.Exists(parent)) RefreshFolderTreeChildrenAtPath(parent);
        status.Text = restored.Count + " öğe Geri Dönüşüm Kutusu'ndan geri yüklendi.";
    }

    private static object InvokeShellMember(object target, string member, BindingFlags operation, object[] arguments)
    {
        if (target == null) return null;
        return target.GetType().InvokeMember(member, operation | BindingFlags.Public | BindingFlags.Instance,
            null, target, arguments, CultureInfo.InvariantCulture);
    }

    private static bool RestoreRecycleBinItem(object item)
    {
        object verbs = null;
        try
        {
            verbs = InvokeShellMember(item, "Verbs", BindingFlags.GetProperty, new object[0]);
            int count = Convert.ToInt32(InvokeShellMember(verbs, "Count", BindingFlags.GetProperty, new object[0]), CultureInfo.InvariantCulture);
            for (int index = 0; index < count; index++)
            {
                object verb = null;
                try
                {
                    verb = InvokeShellMember(verbs, "Item", BindingFlags.InvokeMethod, new object[] { index });
                    string name = Convert.ToString(InvokeShellMember(verb, "Name", BindingFlags.GetProperty, new object[0]), CultureInfo.InvariantCulture);
                    string normalized = (name ?? String.Empty).Replace("&", String.Empty).Trim();
                    if (normalized.IndexOf("restore", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        normalized.IndexOf("geri yükle", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        InvokeShellMember(item, "InvokeVerb", BindingFlags.InvokeMethod, new object[] { name });
                        return true;
                    }
                }
                finally { ReleaseShellObject(verb); }
            }
            foreach (string verbName in new string[] { "RESTORE", "Restore", "Geri Yükle" })
            {
                try { InvokeShellMember(item, "InvokeVerb", BindingFlags.InvokeMethod, new object[] { verbName }); return true; }
                catch { }
            }
        }
        finally { ReleaseShellObject(verbs); }
        return false;
    }

    private static void ReleaseShellObject(object value)
    {
        try { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); } catch { }
    }

    private void PopulateShellNewMenu(ToolStripMenuItem newMenu)
    {
        newMenu.DropDownItems.Clear();
        var folderItem = new ToolStripMenuItem("Klasör");
        folderItem.Click += (s, e) => CreateFolderIn(currentPath);
        newMenu.DropDownItems.Add(folderItem);
        newMenu.DropDownItems.Add(new ToolStripSeparator());
        foreach (ShellNewTemplate template in GetShellNewTemplates())
        {
            var item = new ToolStripMenuItem(template.DisplayName) { Tag = template };
            item.Click += (s, e) => CreateShellNewFile((ShellNewTemplate)((ToolStripMenuItem)s).Tag);
            newMenu.DropDownItems.Add(item);
        }
        if (newMenu.DropDownItems.Count == 2)
        {
            newMenu.DropDownItems.Add(new ToolStripMenuItem("Bu bilgisayarda kayıtlı dosya şablonu yok") { Enabled = false });
        }
    }

    private IEnumerable<ShellNewTemplate> GetShellNewTemplates()
    {
        var result = new List<ShellNewTemplate>
        {
            new ShellNewTemplate(".txt", "Metin Belgesi (.txt)", null, new byte[0], true),
            new ShellNewTemplate(".csv", "CSV Belgesi (.csv)", null, new byte[0], true),
            new ShellNewTemplate(".xml", "XML Belgesi (.xml)", null, Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<root>\r\n</root>\r\n"), false),
            new ShellNewTemplate(".json", "JSON Belgesi (.json)", null, Encoding.UTF8.GetBytes("{}\r\n"), false),
            new ShellNewTemplate(".md", "Markdown Belgesi (.md)", null, new byte[0], true),
            new ShellNewTemplate(".html", "HTML Belgesi (.html)", null, Encoding.UTF8.GetBytes("<!doctype html>\r\n<html lang=\"tr\">\r\n<head><meta charset=\"utf-8\"><title>Yeni belge</title></head>\r\n<body>\r\n</body>\r\n</html>\r\n"), false),
            new ShellNewTemplate(".docx", "Word Belgesi (.docx)", null, null, false, "word"),
            new ShellNewTemplate(".xlsx", "Excel Çalışma Kitabı (.xlsx)", null, null, false, "excel")
        };
        var knownExtensions = new HashSet<string>(result.Select(item => item.Extension), StringComparer.OrdinalIgnoreCase);
        string[] extensions;
        try { extensions = Registry.ClassesRoot.GetSubKeyNames().Where(name => name.StartsWith(".", StringComparison.Ordinal) && name.Length > 1).ToArray(); }
        catch { return result; }
        foreach (string extension in extensions)
        {
            try
            {
                using (RegistryKey shellNew = Registry.ClassesRoot.OpenSubKey(extension + "\\ShellNew"))
                {
                    if (shellNew == null) continue;
                    object nullFile = shellNew.GetValue("NullFile");
                    object fileNameValue = shellNew.GetValue("FileName", null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    object dataValue = shellNew.GetValue("Data");
                    if (nullFile == null && fileNameValue == null && dataValue == null) continue;
                    if (knownExtensions.Contains(extension)) continue;
                    string templatePath = fileNameValue == null ? null : ResolveShellNewTemplatePath(Environment.ExpandEnvironmentVariables(fileNameValue.ToString()));
                    if (fileNameValue != null && String.IsNullOrWhiteSpace(templatePath)) continue;
                    byte[] data = dataValue as byte[];
                    if (data == null && dataValue is string) data = Encoding.Default.GetBytes((string)dataValue);
                    string display = GetFileTypeDescription("new" + extension);
                    if (String.IsNullOrWhiteSpace(display) || display.Equals(extension.TrimStart('.'), StringComparison.OrdinalIgnoreCase))
                        display = extension.TrimStart('.').ToUpper(CultureInfo.CurrentCulture) + " dosyası";
                    result.Add(new ShellNewTemplate(extension, "Yeni " + display, templatePath, data, nullFile != null));
                    knownExtensions.Add(extension);
                }
            }
            catch { }
        }
        return result.OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase);
    }

    private static string ResolveShellNewTemplatePath(string templateName)
    {
        if (String.IsNullOrWhiteSpace(templateName)) return null;
        var candidates = new List<string> { templateName };
        if (!Path.IsPathRooted(templateName))
        {
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "ShellNew", templateName));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Templates", templateName));
            candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Templates", templateName));
        }
        return candidates.FirstOrDefault(File.Exists);
    }

    private void CreateShellNewFile(ShellNewTemplate template)
    {
        if (template == null || !Directory.Exists(currentPath)) return;
        string suggestedName = template.BuiltInKind == "word" ? "Yeni Word Belgesi.docx" :
            template.BuiltInKind == "excel" ? "Yeni Excel Çalışma Kitabı.xlsx" :
            "Yeni " + template.Extension.TrimStart('.').ToUpper(CultureInfo.CurrentCulture) + template.Extension;
        using (var dialog = new NewFileNameDialog(template.DisplayName, suggestedName))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string name = dialog.FileName.Trim();
            if (name.Length == 0 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                MessageBox.Show(this, "Geçerli bir dosya adı girin.", "Yeni dosya", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Path.GetExtension(name).Length == 0) name += template.Extension;
            string destination = Path.Combine(currentPath, name);
            if (File.Exists(destination) || Directory.Exists(destination))
            {
                MessageBox.Show(this, "Bu adla bir dosya veya klasör zaten var.", "Yeni dosya", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                if (template.BuiltInKind == "word") CreateBlankWordDocument(destination);
                else if (template.BuiltInKind == "excel") CreateBlankExcelWorkbook(destination);
                else if (!String.IsNullOrWhiteSpace(template.TemplatePath)) File.Copy(template.TemplatePath, destination);
                else if (template.Data != null) File.WriteAllBytes(destination, template.Data);
                else File.WriteAllBytes(destination, new byte[0]);
                Navigate(currentPath, null, destination, false);
                status.Text = "Yeni dosya oluşturuldu: " + name;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Dosya oluşturulamadı.\r\n\r\n" + ex.Message,
                    "Yeni dosya", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    private static void CreateBlankWordDocument(string path)
    {
        using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
        using (var archive = new ZipArchive(file, ZipArchiveMode.Create))
        {
            AddOpenXmlPart(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>");
            AddOpenXmlPart(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            AddOpenXmlPart(archive, "word/document.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>");
            AddOpenXmlPart(archive, "word/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style></w:styles>");
            AddOpenXmlPart(archive, "word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
        }
    }

    private static void CreateBlankExcelWorkbook(string path)
    {
        using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
        using (var archive = new ZipArchive(file, ZipArchiveMode.Create))
        {
            AddOpenXmlPart(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>");
            AddOpenXmlPart(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            AddOpenXmlPart(archive, "xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sayfa1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            AddOpenXmlPart(archive, "xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
            AddOpenXmlPart(archive, "xl/worksheets/sheet1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><sheetData/></worksheet>");
            AddOpenXmlPart(archive, "xl/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"0\"/><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");
        }
    }

    private static void AddOpenXmlPart(ZipArchive archive, string name, string content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using (Stream stream = entry.Open())
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(content);
    }

    private void CreateFolderIn(string parentPath)
    {
        if (String.IsNullOrWhiteSpace(parentPath) || !Directory.Exists(parentPath)) return;
        using (var dialog = new NewFolderDialog())
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            string name = dialog.FolderName.Trim();
            if (name.Length == 0 || name == "." || name == ".." || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                MessageBox.Show(this, "Geçerli bir klasör adı girin.", "Klasör ekle", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string newPath = Path.Combine(parentPath, name);
            if (Directory.Exists(newPath) || File.Exists(newPath))
            {
                MessageBox.Show(this, "Bu adla bir dosya veya klasör zaten var.", "Klasör ekle", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                Directory.CreateDirectory(newPath);
                RefreshFolderTreeChildrenAtPath(parentPath);
                if (String.Equals(parentPath, currentPath, StringComparison.OrdinalIgnoreCase))
                    Navigate(currentPath, null, null, false);
                status.Text = "Klasör oluşturuldu: " + name;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Klasör oluşturulamadı.\r\n\r\n" + ex.Message,
                    "Klasör ekle", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    private void ToggleEmbeddedPreview()
    {
        if (previewPaneOpen)
        {
            previewPaneOpen = false;
            StopEmbeddedPreview();
            explorerLayout.Panel2Collapsed = true;
            try { folderFilesSplit.SplitterDistance = folderWidthBeforePreview; } catch { }
            actionTooltips.SetToolTip(togglePreview, "Seçili dosyanın ön izlemesini aç");
            return;
        }

        Entry selected = files.SelectedItems.Count > 0 ? files.SelectedItems[0].Tag as Entry : null;
        if (selected == null || !File.Exists(selected.FullPath))
        {
            status.Text = "Ön izleme için önce bir dosya seçin.";
            return;
        }
        folderWidthBeforePreview = folderFilesSplit.SplitterDistance;
        previewPaneOpen = true;
        explorerLayout.Panel2Collapsed = false;
        actionTooltips.SetToolTip(togglePreview, "Ön izleme panelini kapat");
        BeginInvoke((MethodInvoker)(() =>
        {
            if (!previewPaneOpen) return;
            int max = explorerLayout.ClientSize.Width - explorerLayout.Panel2MinSize - explorerLayout.SplitterWidth;
            int desired = (int)(explorerLayout.ClientSize.Width * 0.58);
            try { explorerLayout.SplitterDistance = Math.Max(explorerLayout.Panel1MinSize, Math.Min(max, desired)); } catch { }
            ShowEmbeddedPreview(selected.FullPath);
        }));
    }

    private void ShowEmbeddedPreview(string path)
    {
        if (!previewPaneOpen) return;
        if (embeddedPreview != null && String.Equals(embeddedPreview.PreviewPath, path, StringComparison.OrdinalIgnoreCase)) return;
        StopEmbeddedPreview();
        embeddedPreview = new FilePreviewForm(path);
        embeddedPreview.TopLevel = false;
        embeddedPreview.FormBorderStyle = FormBorderStyle.None;
        embeddedPreview.Dock = DockStyle.Fill;
        previewContent.Controls.Add(embeddedPreview);
        embeddedPreview.BringToFront();
        embeddedPreview.Show();
    }

    private void StopEmbeddedPreview()
    {
        if (embeddedPreview == null) return;
        FilePreviewForm old = embeddedPreview;
        embeddedPreview = null;
        try { old.Close(); } catch { }
        try { old.Dispose(); } catch { }
    }

    private void SendByMail(string path)
    {
        try
        {
            string subject = Uri.EscapeDataString("Dosya gönderimi: " + Path.GetFileName(path));
            string body = Uri.EscapeDataString("Lütfen bu dosyayı iletiye ekleyin:\r\n" + path);
            string mailto = "mailto:?subject=" + subject + "&body=" + body;
            Process.Start(new ProcessStartInfo(mailto) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Mail Gönder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void AddToHistory(string path)
    {
        history.RemoveAll(x => String.Equals(x.FullPath, path, StringComparison.OrdinalIgnoreCase));
        history.Insert(0, new HistoryEntry(path, DateTime.Now));
        if (history.Count > 500) history.RemoveRange(500, history.Count - 500);
        PruneHistoryByRetention();
        SaveHistory();
        RefreshHistoryList();
    }

    private void LoadHistory()
    {
        try
        {
            if (File.Exists(historyFile))
                foreach (string line in File.ReadAllLines(historyFile))
                {
                    int separator = line.LastIndexOf('\t');
                    if (separator < 1) continue;
                    long binary;
                    if (Int64.TryParse(line.Substring(separator + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out binary))
                    {
                        try
                        {
                            string path = ResolvePortablePath(line.Substring(0, separator));
                            if (!String.IsNullOrWhiteSpace(path)) history.Add(new HistoryEntry(path, DateTime.FromBinary(binary)));
                        }
                        catch { }
                    }
                }
        }
        catch { }
        PruneHistoryByRetention();
        RefreshHistoryList();
    }

    private void SaveHistoryRetentionFromControls()
    {
        int count;
        if (historyRetentionCountSelector.SelectedItem != null &&
            Int32.TryParse(historyRetentionCountSelector.SelectedItem.ToString(), out count))
            historyRetentionCount = Math.Max(1, Math.Min(30, count));
        if (historyRetentionUnitSelector.SelectedItem != null)
            historyRetentionUnit = historyRetentionUnitSelector.SelectedItem.ToString();
        SavePreferences();
        PruneHistoryByRetention();
        RefreshHistoryList();
    }

    private void PruneHistoryByRetention()
    {
        if (!preserveHistoryRetention) return;
        DateTime cutoff = GetHistoryRetentionCutoff();
        int oldCount = history.Count;
        history.RemoveAll(item => item.OpenedAt < cutoff);
        if (history.Count != oldCount) SaveHistory();
    }

    private DateTime GetHistoryRetentionCutoff()
    {
        DateTime cutoff = DateTime.Now;
        if (historyRetentionUnit == "Hafta") cutoff = cutoff.AddDays(-7 * historyRetentionCount);
        else if (historyRetentionUnit == "Ay") cutoff = cutoff.AddMonths(-historyRetentionCount);
        else if (historyRetentionUnit == "Yıl") cutoff = cutoff.AddYears(-historyRetentionCount);
        else cutoff = cutoff.AddDays(-historyRetentionCount);
        return cutoff;
    }

    private void ClearHistory()
    {
        if (history.Count == 0) return;
        DateTime cutoff = preserveHistoryRetention ? GetHistoryRetentionCutoff() : DateTime.MinValue;
        int removeCount = preserveHistoryRetention
            ? history.Count(item => item.OpenedAt < cutoff)
            : history.Count;
        if (removeCount == 0)
        {
            MessageBox.Show(this, "Seçili koruma döneminden daha eski bir geçmiş kaydı yok.",
                "Geçmişi temizle", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        string prompt = preserveHistoryRetention
            ? removeCount + " eski kayıt silinsin mi? Son " + historyRetentionCount + " " + historyRetentionUnit + " içindeki " + (history.Count - removeCount) + " kayıt korunacak."
            : "Geçmiş sekmesindeki " + removeCount + " kaydın tümü silinsin mi?";
        DialogResult answer = MessageBox.Show(this, prompt, "Geçmişi temizle", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (answer != DialogResult.Yes) return;
        if (preserveHistoryRetention) history.RemoveAll(item => item.OpenedAt < cutoff);
        else history.Clear();
        SaveHistory();
        RefreshHistoryList();
    }

    private void SaveHistory()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(historyFile));
            File.WriteAllLines(historyFile, history.Select(x => StorePortablePath(x.FullPath) + "\t" + x.OpenedAt.ToBinary().ToString(CultureInfo.InvariantCulture)).ToArray());
        }
        catch { }
    }

    private void RefreshHistoryList()
    {
        historyList.BeginUpdate();
        historyList.Items.Clear();
        IEnumerable<HistoryEntry> items = history.Where(x => MatchesFileFilter(x.FullPath, historyFilter.SelectedItem as string));
        if (historySortColumn == 0)
            items = historySortAscending
                ? items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => Path.GetFileName(x.FullPath), NaturalNameComparer.Instance)
                : items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenByDescending(x => Path.GetFileName(x.FullPath), NaturalNameComparer.Instance);
        else if (historySortColumn == 1)
            items = historySortAscending ? items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => x.OpenedAt) : items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenByDescending(x => x.OpenedAt);
        else if (historySortColumn == 2)
            items = historySortAscending ? items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => FileLength(x.FullPath)) : items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenByDescending(x => FileLength(x.FullPath));
        else items = items.OrderBy(x => Directory.Exists(x.FullPath) ? 0 : 1).ThenBy(x => Path.GetFileName(x.FullPath), NaturalNameComparer.Instance);
        foreach (HistoryEntry item in items)
        {
            var row = new ListViewItem(new string[] { Path.GetFileName(item.FullPath), item.OpenedAt.ToString("dd.MM.yyyy HH:mm"), FormatFileSize(item.FullPath) });
            row.Tag = new Entry(Path.GetFileName(item.FullPath), item.FullPath);
            row.ToolTipText = item.FullPath;
            SetFileRowIcon(row, item.FullPath);
            historyList.Items.Add(row);
        }
        historyList.EndUpdate();
    }

    private void SaveLastPath()
    {
        try
        {
            string directory = Path.GetDirectoryName(settingsFile);
            Directory.CreateDirectory(directory);
            File.WriteAllText(settingsFile, StorePortablePath(currentPath));
        }
        catch { }
    }

    private void SavePortablePathData()
    {
        if (!oneDriveRootEnabled || !Directory.Exists(oneDriveRoot)) return;
        SaveLastPath();
        SaveNavigationHistory();
        SaveFolderTabSessions();
        SaveShortcuts();
        SaveHistory();
        SaveTransferPanePaths();
        SaveWorkProfiles();
        SaveWorkProfileLocations();
    }

    private int LoadSplitterDistance()
    {
        try
        {
            int distance;
            if (File.Exists(splitterFile) && Int32.TryParse(File.ReadAllText(splitterFile).Trim(), out distance))
                return distance;
        }
        catch { }
        return 140;
    }

    private void SaveSplitterDistance(int distance)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(splitterFile));
            File.WriteAllText(splitterFile, distance.ToString(CultureInfo.InvariantCulture));
        }
        catch { }
    }

    private static void RestoreSplitterDistance(SplitContainer split, int distance)
    {
        int maxDistance = Math.Max(split.Panel1MinSize,
            split.ClientSize.Width - split.Panel2MinSize - split.SplitterWidth);
        int restoredDistance = Math.Max(split.Panel1MinSize, Math.Min(maxDistance, distance));
        try { split.SplitterDistance = restoredDistance; }
        catch { }
    }

    private bool LoadMaximizedState()
    {
        try { return File.Exists(windowStateFile) && File.ReadAllText(windowStateFile).Trim() == "1"; }
        catch { return false; }
    }

    private void SaveMaximizedState(bool maximized)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(windowStateFile));
            File.WriteAllText(windowStateFile, maximized ? "1" : "0");
        }
        catch { }
    }

    private void RestoreWindowBounds()
    {
        try
        {
            if (!File.Exists(windowBoundsFile)) return;
            string[] values = File.ReadAllText(windowBoundsFile).Split(',');
            int x, y, width, height;
            if (values.Length != 4 || !Int32.TryParse(values[0], out x) || !Int32.TryParse(values[1], out y) ||
                !Int32.TryParse(values[2], out width) || !Int32.TryParse(values[3], out height)) return;
            if (width < MinimumSize.Width || height < MinimumSize.Height) return;
            Rectangle candidate = new Rectangle(x, y, width, height);
            if (!Screen.AllScreens.Any(screen => screen.WorkingArea.IntersectsWith(candidate))) return;
            Bounds = candidate;
            StartPosition = FormStartPosition.Manual;
        }
        catch { }
    }

    private void SaveWindowBounds()
    {
        try
        {
            if (WindowState != FormWindowState.Normal || Width < MinimumSize.Width || Height < MinimumSize.Height) return;
            string directory = Path.GetDirectoryName(windowBoundsFile);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(windowBoundsFile, String.Join(",", new string[]
            {
                Left.ToString(CultureInfo.InvariantCulture), Top.ToString(CultureInfo.InvariantCulture),
                Width.ToString(CultureInfo.InvariantCulture), Height.ToString(CultureInfo.InvariantCulture)
            }));
        }
        catch { }
    }

    private sealed class FileSortPreference
    {
        public readonly int Column;
        public readonly bool Ascending;
        public FileSortPreference(int column, bool ascending) { Column = column; Ascending = ascending; }
    }

    private sealed class SyncWorkspaceProfile
    {
        public string Name { get; set; }
        public string LeftPath { get; set; }
        public string RightPath { get; set; }
        public int ComparisonMode { get; set; }
        public bool IncludeSubfolders { get; set; }
        public bool ShowIdentical { get; set; }
        public override string ToString() { return Name ?? "Senkron profili"; }
    }

    private sealed class SyncComparisonRequest
    {
        public string LeftPath;
        public string RightPath;
        public bool IncludeSubfolders;
        public bool ShowIdentical;
        public bool CompareContent;
    }

    private sealed class SyncComparisonRow
    {
        public string Category;
        public string RelativePath;
        public string LeftPath;
        public string RightPath;
        public string SizeText;
        public string CloudStatus;
        public string LookupStatus;
        public string LookupRootPath;
        public string LookupSizeText;
    }

    private sealed class SyncLookupRequest
    {
        public string LeftRoot;
        public string RightRoot;
        public List<SyncComparisonRow> Rows;
    }

    private sealed class SyncLookupResult
    {
        public string LeftRoot;
        public string RightRoot;
        public List<SyncComparisonRow> Rows;
        public Dictionary<string, List<SyncFileMetadata>> LeftFilesByName;
        public Dictionary<string, List<SyncFileMetadata>> RightFilesByName;
        public int Errors;
        public List<string> ErrorDetails;
    }

    private sealed class SyncFileMetadata
    {
        public string FullPath;
        public long Length;
        public DateTime LastWriteTimeUtc;
        public bool IsCloudPlaceholder;
        public bool IsOnlineOnly;
    }

    private sealed class SyncComparisonResult
    {
        public readonly List<SyncComparisonRow> Rows = new List<SyncComparisonRow>();
        public readonly List<string> ErrorDetails = new List<string>();
        public int LeftOnlyCount;
        public int RightOnlyCount;
        public int DifferentCount;
        public int IdenticalCount;
        public int CloudContentDeferredCount;
        public int CloudPlaceholderCount;
        public int Errors;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Win32FileTime
    {
        public int Low;
        public int High;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct Win32FindData
    {
        public FileAttributes FileAttributes;
        public Win32FileTime CreationTime;
        public Win32FileTime LastAccessTime;
        public Win32FileTime LastWriteTime;
        public uint FileSizeHigh;
        public uint FileSizeLow;
        public uint Reserved0;
        public uint Reserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string FileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string AlternateFileName;
    }

    private sealed class UpdateRelease
    {
        public readonly string Version;
        public readonly Uri DownloadUri;
        public readonly string Sha256;
        public readonly string Notes;
        public readonly Uri UpdaterUri;
        public readonly string UpdaterSha256;
        public UpdateRelease(string version, Uri downloadUri, string sha256, string notes, Uri updaterUri, string updaterSha256)
        {
            Version = version;
            DownloadUri = downloadUri;
            Sha256 = sha256;
            Notes = notes;
            UpdaterUri = updaterUri;
            UpdaterSha256 = updaterSha256;
        }
    }

    private sealed class PaneWidthPreference
    {
        public readonly int PrimaryWidth;
        public readonly int ExplorerFolderWidth;
        public PaneWidthPreference(int primaryWidth, int explorerFolderWidth)
        {
            PrimaryWidth = primaryWidth;
            ExplorerFolderWidth = explorerFolderWidth;
        }
    }

    private sealed class ColumnLayout
    {
        public readonly int[] DisplayIndexes;
        public readonly int[] Widths;
        public ColumnLayout(int count)
        {
            DisplayIndexes = new int[count];
            Widths = new int[count];
        }
    }

    private sealed class WorkProfile
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string RootPath { get; set; }
        public string LastPath { get; set; }
        public string Notes { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime LastOpenedAt { get; set; }
        public bool IsCompleted { get; set; }
        public readonly List<string> RelatedFolders = new List<string>();
        public readonly List<string> PortableRelatedPaths = new List<string>();
    }

    private sealed class DevelopmentTask
    {
        public Guid Id { get; set; }
        public string Text { get; set; }
        public bool IsCompleted { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    private sealed class DevelopmentVersion
    {
        public Guid Id { get; set; }
        public int Number { get; set; }
        public int Major { get; set; }
        public int Minor { get; set; }
        public int Patch { get; set; }
        public DateTime UsedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
        public readonly List<string> Changes = new List<string>();
    }

    private sealed class DevelopmentTaskDialog : Form
    {
        private readonly TextBox textBox = new TextBox();
        public string TaskText { get { return textBox.Text; } }

        public DevelopmentTaskDialog(string currentText)
        {
            Text = "İsteği düzelt";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 170);
            Font = new Font("Segoe UI", 9F);
            var label = new Label { Text = "İstek metni:", AutoSize = true, Location = new Point(14, 14) };
            textBox.SetBounds(14, 39, 492, 68);
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Vertical;
            textBox.Text = currentText ?? String.Empty;
            var save = new Button { Text = "Kaydet", DialogResult = DialogResult.OK };
            save.SetBounds(332, 125, 82, 30);
            var cancel = new Button { Text = "İptal", DialogResult = DialogResult.Cancel };
            cancel.SetBounds(424, 125, 82, 30);
            Controls.Add(label);
            Controls.Add(textBox);
            Controls.Add(save);
            Controls.Add(cancel);
            AcceptButton = save;
            CancelButton = cancel;
            Shown += (s, e) => { textBox.Focus(); textBox.SelectAll(); };
        }
    }

    private sealed class FolderTabNameDialog : Form
    {
        private readonly TextBox nameBox = new TextBox();
        public string TabName { get { return nameBox.Text; } }

        public FolderTabNameDialog(string currentName, string folderName)
        {
            Text = "Sekmeye isim ver";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(390, 138);
            Font = new Font("Segoe UI", 9F);
            var label = new Label
            {
                Text = "Sekme adı (boş bırakırsanız klasör adı kullanılır):",
                AutoSize = true,
                Location = new Point(14, 14)
            };
            nameBox.SetBounds(14, 40, 362, 27);
            nameBox.Text = String.IsNullOrWhiteSpace(currentName) ? folderName : currentName;
            var save = new Button { Text = "Kaydet", DialogResult = DialogResult.OK };
            save.SetBounds(204, 91, 82, 30);
            var cancel = new Button { Text = "İptal", DialogResult = DialogResult.Cancel };
            cancel.SetBounds(294, 91, 82, 30);
            Controls.Add(label);
            Controls.Add(nameBox);
            Controls.Add(save);
            Controls.Add(cancel);
            AcceptButton = save;
            CancelButton = cancel;
            Shown += (s, e) => { nameBox.Focus(); nameBox.SelectAll(); };
        }
    }

    private sealed class FolderTabState
    {
        public string Path;
        public string ShortcutPath;
        public List<string> History;
        public int HistoryIndex;
        public string CustomName;
        public bool IsUnavailable;

        public FolderTabState(string path, IEnumerable<string> history, int historyIndex)
        {
            Path = path;
            ShortcutPath = String.Empty;
            History = history == null ? new List<string>() : history.ToList();
            HistoryIndex = historyIndex;
            CustomName = String.Empty;
            IsUnavailable = false;
        }
    }

    private sealed class Entry
    {
        public readonly string FullPath;
        public readonly bool IsParent;
        public readonly DateTime AddedAt;
        public readonly int MatchStart;
        public readonly int MatchLength;
        private readonly string display;
        public Entry(string display, string path) : this(display, path, false, DateTime.MinValue, 0, 0) { }
        public Entry(string display, string path, int matchStart, int matchLength) : this(display, path, false, DateTime.MinValue, matchStart, matchLength) { }
        public Entry(string display, string path, bool isParent) : this(display, path, isParent, DateTime.MinValue, 0, 0) { }
        public Entry(string display, string path, DateTime addedAt) : this(display, path, false, addedAt, 0, 0) { }
        private Entry(string display, string path, bool isParent, DateTime addedAt, int matchStart, int matchLength)
        {
            this.display = display;
            FullPath = path;
            IsParent = isParent;
            AddedAt = addedAt;
            MatchStart = matchStart;
            MatchLength = matchLength;
        }
        public override string ToString() { return display; }
    }

    private sealed class TransferPaneState
    {
        public readonly int Side;
        public readonly string Title;
        public string CurrentPath;
        public bool ShowingComputerDrives;
        public string ComputerViewReturnPath;
        public int SortColumn = -1;
        public bool SortAscending = true;
        public string DropTargetPath;
        public int LastFolderPaneWidth = 175;
        public bool SyncingSelection;
        public bool Refreshing;
        public bool Loading;
        public int LoadGeneration;
        public bool FolderPaneHidden;
        public FileSystemWatcher Watcher;
        public System.Windows.Forms.Timer RefreshTimer;
        public string WatchedPath = String.Empty;
        public DateTime LastWatcherRefreshUtc = DateTime.MinValue;
        public TextBox Address;
        public Button LocationButton;
        public Button UpButton;
        public Button GoButton;
        public Button FolderWidthButton;
        public SplitContainer FolderSplit;
        public ListBox FolderList;
        public ListView FileList;
        public ContextMenuStrip QuickLocationsMenu;
        public TransferPaneState(int side, string title, string path)
        {
            Side = side;
            Title = title;
            CurrentPath = path;
        }
    }

    private sealed class TransferFileRow
    {
        public readonly Entry Entry;
        public readonly string Name, Size, Type, Modified;
        public readonly bool IsDirectory;
        public TransferFileRow(Entry entry, string name, string size, string type, string modified, bool isDirectory)
        { Entry = entry; Name = name; Size = size; Type = type; Modified = modified; IsDirectory = isDirectory; }
    }

    private sealed class TransferDragFeedbackForm : Form
    {
        private readonly int fileCount;
        private string actionText = "Taşınıyor";

        public TransferDragFeedbackForm(int fileCount)
        {
            this.fileCount = fileCount;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            Enabled = true;
            BackColor = Color.White;
            Opacity = 0.94;
            Size = new Size(136, 66);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams parameters = base.CreateParams;
                parameters.ExStyle |= 0x00000020 | 0x00000080 | 0x00080000;
                return parameters;
            }
        }

        public void SetDragInfo(string action, Point cursor)
        {
            if (!String.Equals(actionText, action, StringComparison.Ordinal))
            {
                actionText = action;
                Invalidate();
            }
            Rectangle workArea = Screen.FromPoint(cursor).WorkingArea;
            int x = Math.Min(cursor.X + 18, workArea.Right - Width);
            int y = Math.Min(cursor.Y + 20, workArea.Bottom - Height);
            Location = new Point(Math.Max(workArea.Left, x), Math.Max(workArea.Top, y));
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle card = new Rectangle(1, 1, Width - 3, Height - 3);
            using (var background = new SolidBrush(Color.FromArgb(250, 255, 255, 255)))
            using (var border = new Pen(Color.FromArgb(180, 94, 111, 131), 1F))
            {
                e.Graphics.FillRectangle(background, card);
                e.Graphics.DrawRectangle(border, card);
            }
            DrawTransferDocumentGlyph(e.Graphics, new Rectangle((Width - 34) / 2, 4, 34, 36));
            string caption = actionText + (fileCount > 1 ? " (" + fileCount.ToString(CultureInfo.CurrentCulture) + ")" : String.Empty);
            using (var font = new Font("Segoe UI", 8.5F, FontStyle.Bold, GraphicsUnit.Point))
            using (var brush = new SolidBrush(Color.FromArgb(36, 48, 64)))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter })
                e.Graphics.DrawString(caption, font, brush, new RectangleF(4, 42, Width - 8, 20), format);
        }
    }

    private sealed class TransferListViewComparer : System.Collections.IComparer
    {
        private readonly int column;
        private readonly bool ascending;
        public TransferListViewComparer(int column, bool ascending)
        {
            this.column = column;
            this.ascending = ascending;
        }
        public int Compare(object first, object second)
        {
            ListViewItem a = first as ListViewItem;
            ListViewItem b = second as ListViewItem;
            if (a == null || b == null) return 0;
            bool directoryA = a.SubItems.Count > 2 && String.Equals(a.SubItems[2].Text, "Dosya klasörü", StringComparison.OrdinalIgnoreCase);
            bool directoryB = b.SubItems.Count > 2 && String.Equals(b.SubItems[2].Text, "Dosya klasörü", StringComparison.OrdinalIgnoreCase);
            if (directoryA != directoryB) return directoryA ? -1 : 1;
            int result = 0;
            switch (column)
            {
                case 1: result = ParseTransferSize(a.SubItems[1].Text).CompareTo(ParseTransferSize(b.SubItems[1].Text)); break;
                case 2: result = StringComparer.CurrentCultureIgnoreCase.Compare(a.SubItems[2].Text, b.SubItems[2].Text); break;
                case 3:
                    DateTime dateA, dateB;
                    DateTime.TryParseExact(a.SubItems[3].Text, "dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture, DateTimeStyles.None, out dateA);
                    DateTime.TryParseExact(b.SubItems[3].Text, "dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture, DateTimeStyles.None, out dateB);
                    result = dateA.CompareTo(dateB);
                    break;
                default: result = NaturalNameComparer.Instance.Compare(a.Text, b.Text); break;
            }
            if (result == 0 && column != 0) result = NaturalNameComparer.Instance.Compare(a.Text, b.Text);
            return ascending ? result : -result;
        }

        private static double ParseTransferSize(string text)
        {
            if (String.IsNullOrWhiteSpace(text)) return -1;
            string number = text.Replace("KB", String.Empty).Replace("MB", String.Empty).Trim();
            double value;
            if (!Double.TryParse(number, NumberStyles.Float, CultureInfo.CurrentCulture, out value)) return -1;
            return text.IndexOf("MB", StringComparison.OrdinalIgnoreCase) >= 0 ? value * 1024.0 : value;
        }
    }

    private sealed class MoveRecord
    {
        public readonly string SourcePath;
        public readonly string DestinationPath;
        public readonly string BackupPath;
        public MoveRecord(string sourcePath, string destinationPath, string backupPath)
        {
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
            BackupPath = backupPath;
        }
    }

    private sealed class HistoryEntry
    {
        public readonly string FullPath;
        public readonly DateTime OpenedAt;
        public HistoryEntry(string path, DateTime openedAt) { FullPath = path; OpenedAt = openedAt; }
    }

    private sealed class SearchRequest
    {
        public readonly string Root;
        public readonly Regex[] Patterns;
        public readonly string Query;
        public readonly bool SearchSubfolders;
        public SearchRequest(string root, Regex[] patterns, string query, bool searchSubfolders)
        {
            Root = root; Patterns = patterns; Query = query; SearchSubfolders = searchSubfolders;
        }
    }

    private sealed class XmlSourceViewerForm : Form
    {
        public XmlSourceViewerForm(string path)
        {
            Text = "XML ön izleme — " + Path.GetFileName(path);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(650, 420);
            Size = new Size(1000, 700);
            WindowState = FormWindowState.Maximized;
            var source = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = false,
                DetectUrls = false,
                Font = new Font("Consolas", 10F),
                BackColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            Controls.Add(source);
            try
            {
                using (var reader = new StreamReader(path, true)) source.Text = reader.ReadToEnd();
            }
            catch (Exception ex) { source.Text = "XML dosyası okunamadı.\r\n\r\n" + ex.Message; }
        }
    }

    private sealed class EDefterViewerForm : Form
    {
        private const string ResourcePrefix = "DosyaGezgini.";
        private readonly string xmlPath;
        private readonly string stylesheetName;
        private readonly string configuredTemplateFolder;
        private readonly WebBrowser browser = new WebBrowser();
        private readonly Label state = new Label();
        private readonly ProgressBar progress = new ProgressBar();
        private readonly Button exportExcel = new Button();
        private readonly BackgroundWorker worker = new BackgroundWorker();
        private bool documentReady;
        private string transformedHtml = String.Empty;

        public EDefterViewerForm(string path, string selectedStylesheet, string templateFolder)
        {
            xmlPath = path;
            stylesheetName = selectedStylesheet;
            configuredTemplateFolder = templateFolder ?? String.Empty;
            Text = "E-defter — " + Path.GetFileName(path);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 500);
            Size = new Size(1180, 780);
            WindowState = FormWindowState.Maximized;
            var toolbar = new Panel { Dock = DockStyle.Top, Height = 42, BackColor = Color.White, Padding = new Padding(12, 8, 12, 6) };
            state.Text = "E-defter görünümü hazırlanıyor...";
            state.SetBounds(12, 8, 600, 26);
            state.TextAlign = ContentAlignment.MiddleLeft;
            state.ForeColor = Color.FromArgb(55, 65, 81);
            progress.SetBounds(0, 12, 150, 16);
            progress.Style = ProgressBarStyle.Marquee;
            progress.MarqueeAnimationSpeed = 25;
            exportExcel.Text = String.Equals(stylesheetName, "berat.xslt", StringComparison.OrdinalIgnoreCase)
                ? "PDF olarak kaydet" : "Excel'e kaydet";
            exportExcel.SetBounds(0, 5, 124, 31);
            exportExcel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            exportExcel.Enabled = false;
            exportExcel.FlatStyle = FlatStyle.Flat;
            exportExcel.FlatAppearance.BorderColor = Color.FromArgb(202, 212, 225);
            exportExcel.BackColor = Color.FromArgb(239, 245, 255);
            exportExcel.Click += (s, e) => ExportToExcel();
            toolbar.Controls.Add(state);
            toolbar.Controls.Add(progress);
            toolbar.Controls.Add(exportExcel);
            toolbar.Resize += (s, e) =>
            {
                exportExcel.Left = toolbar.ClientSize.Width - 136;
                progress.Left = exportExcel.Left - 174;
                state.Left = 12;
                state.Width = Math.Max(100, progress.Left - 24);
            };
            browser.Dock = DockStyle.Fill;
            browser.ScriptErrorsSuppressed = true;
            browser.AllowWebBrowserDrop = false;
            browser.IsWebBrowserContextMenuEnabled = true;
            browser.DocumentCompleted += (s, e) =>
            {
                if (browser.ReadyState != WebBrowserReadyState.Complete) return;
                documentReady = browser.Document != null;
                exportExcel.Enabled = documentReady;
            };
            Controls.Add(browser);
            Controls.Add(toolbar);
            worker.DoWork += (s, e) => e.Result = TransformFile(xmlPath, stylesheetName, configuredTemplateFolder);
            worker.RunWorkerCompleted += (s, e) =>
            {
                progress.Visible = false;
                if (e.Error != null)
                {
                    state.Text = "Görüntüleme başarısız oldu.";
                    browser.DocumentText = "<html><body style='font:14px Segoe UI;padding:24px'><h2>E-defter görüntülenemedi</h2><pre style='white-space:pre-wrap'>" +
                        System.Net.WebUtility.HtmlEncode(e.Error.GetBaseException().Message) + "</pre></body></html>";
                    return;
                }
                state.Text = "Şablon: " + stylesheetName + "     |     " + xmlPath;
                transformedHtml = e.Result as string ?? "<html><body>İçerik üretilemedi.</body></html>";
                browser.DocumentText = transformedHtml;
            };
            Shown += (s, e) => worker.RunWorkerAsync();
        }

        public static XmlDocument LoadXmlDocument(string path)
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreComments = false };
            var document = new XmlDocument { XmlResolver = null };
            using (XmlReader reader = XmlReader.Create(path, settings)) document.Load(reader);
            return document;
        }

        public static string ChooseStylesheet(string path, XmlDocument document, string rootName)
        {
            XmlProcessingInstruction instruction = document.SelectSingleNode("/processing-instruction('xml-stylesheet')") as XmlProcessingInstruction;
            if (instruction != null)
            {
                Match href = Regex.Match(instruction.Data, "\\bhref\\s*=\\s*(['\"])(.*?)\\1", RegexOptions.IgnoreCase);
                if (href.Success)
                {
                    string name = Path.GetFileName(Uri.UnescapeDataString(href.Groups[2].Value));
                    if (String.Equals(name, "yevmiye.xslt", StringComparison.OrdinalIgnoreCase)) return "yevmiye.xslt";
                    if (String.Equals(name, "kebir.xslt", StringComparison.OrdinalIgnoreCase)) return "kebir.xslt";
                    if (String.Equals(name, "berat.xslt", StringComparison.OrdinalIgnoreCase)) return "berat.xslt";
                }
            }
            if (rootName == "berat") return "berat.xslt";
            string filename = Path.GetFileNameWithoutExtension(path);
            if (filename.IndexOf("-KB-", StringComparison.OrdinalIgnoreCase) >= 0) return "kebir.xslt";
            return "yevmiye.xslt";
        }

        private static string TransformFile(string xmlFile, string styleName, string configuredFolder)
        {
            XmlDocument style = new XmlDocument { XmlResolver = null };
            using (Stream stream = OpenStylesheet(styleName, configuredFolder))
            {
                if (stream == null) throw new FileNotFoundException("E-defter biçim şablonu bulunamadı: " + styleName);
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (XmlReader reader = XmlReader.Create(stream, settings)) style.Load(reader);
            }
            XmlNamespaceManager ns = new XmlNamespaceManager(style.NameTable);
            ns.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");
            XmlNodeList characterMaps = style.SelectNodes("//xsl:character-map", ns);
            foreach (XmlNode node in characterMaps) node.ParentNode.RemoveChild(node);
            XmlNodeList outputs = style.SelectNodes("//xsl:output", ns);
            foreach (XmlNode node in outputs)
            {
                XmlAttribute characterMapsAttribute = node.Attributes["use-character-maps"];
                if (characterMapsAttribute != null) node.Attributes.Remove(characterMapsAttribute);
                XmlAttribute publicId = node.Attributes["doctype-public"];
                if (publicId != null) node.Attributes.Remove(publicId);
                XmlAttribute systemId = node.Attributes["doctype-system"];
                if (systemId != null) node.Attributes.Remove(systemId);
            }
            XmlElement root = style.DocumentElement;
            if (root != null && root.HasAttribute("version")) root.SetAttribute("version", "1.0");

            var transform = new XslCompiledTransform();
            var safeSettings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using (var styleText = new StringReader(style.OuterXml))
            using (XmlReader styleReader = XmlReader.Create(styleText, safeSettings))
                transform.Load(styleReader, new XsltSettings(false, false), null);

            var inputSettings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            var result = new StringBuilder();
            using (XmlReader input = XmlReader.Create(xmlFile, inputSettings))
            using (var writer = new StringWriter(result, CultureInfo.InvariantCulture))
                transform.Transform(input, null, writer);
            return result.ToString();
        }

        private static Stream OpenStylesheet(string styleName, string configuredFolder)
        {
            string[] folders = new string[]
            {
                configuredFolder,
                PortableDataPath("Şablonlar")
            };
            foreach (string folder in folders.Where(item => !String.IsNullOrWhiteSpace(item)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                string candidate = Path.Combine(folder, styleName);
                if (File.Exists(candidate)) return new FileStream(candidate, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            return typeof(EDefterViewerForm).Assembly.GetManifestResourceStream(ResourcePrefix + styleName);
        }

        private void ExportToExcel()
        {
            if (!documentReady || browser.Document == null) return;
            if (String.Equals(stylesheetName, "berat.xslt", StringComparison.OrdinalIgnoreCase))
            {
                ExportToPdf();
                return;
            }
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "E-defter tablolarını Excel'e aktar";
                dialog.Filter = "Excel çalışma kitabı (*.xlsx)|*.xlsx";
                dialog.FileName = Path.GetFileNameWithoutExtension(xmlPath) + ".xlsx";
                string sourceFolder = Path.GetDirectoryName(xmlPath);
                if (!String.IsNullOrEmpty(sourceFolder) && Directory.Exists(sourceFolder)) dialog.InitialDirectory = sourceFolder;
                dialog.AddExtension = true;
                dialog.DefaultExt = "xlsx";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    List<SheetData> sheets = ReadHtmlTables(browser.Document);
                    if (sheets.Count == 0) throw new InvalidOperationException("Görüntülenen e-defterde Excel'e aktarılabilecek tablo bulunamadı.");
                    if (String.Equals(stylesheetName, "yevmiye.xslt", StringComparison.OrdinalIgnoreCase) ||
                        String.Equals(stylesheetName, "kebir.xslt", StringComparison.OrdinalIgnoreCase))
                    {
                        var allRows = new List<List<string>>();
                        foreach (SheetData sheet in sheets)
                        {
                            if (allRows.Count > 0) allRows.Add(new List<string>());
                            allRows.AddRange(sheet.Rows);
                        }
                        sheets = new List<SheetData> { new SheetData(stylesheetName.StartsWith("kebir", StringComparison.OrdinalIgnoreCase) ? "Kebir" : "Yevmiye", allRows) };
                    }
                    XlsxWriter.Write(dialog.FileName, sheets);
                    MessageBox.Show(this, sheets.Count + " tablo Excel çalışma kitabına aktarıldı.", "Excel'e aktar", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Excel dosyası oluşturulamadı.\r\n\r\n" + ex.Message, "Excel'e aktar", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void ExportToPdf()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "E-defter beratını PDF olarak kaydet";
                dialog.Filter = "PDF dosyası (*.pdf)|*.pdf";
                dialog.FileName = Path.GetFileNameWithoutExtension(xmlPath) + ".pdf";
                string sourceFolder = Path.GetDirectoryName(xmlPath);
                if (!String.IsNullOrEmpty(sourceFolder) && Directory.Exists(sourceFolder)) dialog.InitialDirectory = sourceFolder;
                dialog.AddExtension = true;
                dialog.DefaultExt = "pdf";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                string edge = FindEdgeExecutable();
                if (String.IsNullOrEmpty(edge))
                {
                    MessageBox.Show(this, "PDF oluşturmak için Microsoft Edge bulunamadı.", "PDF olarak kaydet", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DosyaGezginiPdf-" + Guid.NewGuid().ToString("N"));
                try
                {
                    Directory.CreateDirectory(tempDirectory);
                    string htmlFile = Path.Combine(tempDirectory, "document.html");
                    string profileDirectory = Path.Combine(tempDirectory, "edge-profile");
                    File.WriteAllText(htmlFile, transformedHtml, new UTF8Encoding(false));
                    string arguments = "--headless --disable-gpu --no-first-run --no-default-browser-check --no-pdf-header-footer " +
                        "--user-data-dir=\"" + profileDirectory + "\" --print-to-pdf=\"" + dialog.FileName + "\" \"" + new Uri(htmlFile).AbsoluteUri + "\"";
                    var start = new ProcessStartInfo(edge, arguments) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
                    using (Process process = Process.Start(start))
                    {
                        if (process == null || !process.WaitForExit(120000))
                        {
                            if (process != null) try { process.Kill(); } catch { }
                            throw new TimeoutException("PDF dışa aktarma iki dakika içinde tamamlanmadı.");
                        }
                    }
                    if (!File.Exists(dialog.FileName) || new FileInfo(dialog.FileName).Length == 0)
                        throw new InvalidOperationException("Edge PDF dosyasını oluşturamadı.");
                    MessageBox.Show(this, "Berat PDF olarak kaydedildi.", "PDF olarak kaydet", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "PDF dosyası oluşturulamadı.\r\n\r\n" + ex.Message, "PDF olarak kaydet", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally { try { if (Directory.Exists(tempDirectory)) Directory.Delete(tempDirectory, true); } catch { } }
            }
        }

        private static string FindEdgeExecutable()
        {
            string[] candidates = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft", "Edge", "Application", "msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft", "Edge", "Application", "msedge.exe")
            };
            return candidates.FirstOrDefault(File.Exists);
        }

        private static List<SheetData> ReadHtmlTables(HtmlDocument document)
        {
            var sheets = new List<SheetData>();
            HtmlElementCollection tables = document.GetElementsByTagName("table");
            for (int i = 0; i < tables.Count; i++)
            {
                HtmlElement table = tables[i];
                var rows = new List<List<string>>();
                HtmlElementCollection htmlRows = table.GetElementsByTagName("tr");
                foreach (HtmlElement row in htmlRows)
                {
                    var cells = new List<string>();
                    foreach (HtmlElement cell in row.Children)
                    {
                        string tag = cell.TagName;
                        if (!String.Equals(tag, "TD", StringComparison.OrdinalIgnoreCase) && !String.Equals(tag, "TH", StringComparison.OrdinalIgnoreCase)) continue;
                        cells.Add(NormalizeCell(cell.InnerText));
                    }
                    if (cells.Count > 0) rows.Add(cells);
                }
                if (rows.Count == 0) continue;
                string name = "Tablo " + (sheets.Count + 1).ToString(CultureInfo.InvariantCulture);
                sheets.Add(new SheetData(name, rows));
            }
            return sheets;
        }

        private static string NormalizeCell(string value)
        {
            if (String.IsNullOrEmpty(value)) return String.Empty;
            string normalized = Regex.Replace(value.Replace('\u00a0', ' '), "[\\r\\n\\t ]+", " ").Trim();
            var clean = new StringBuilder(normalized.Length);
            foreach (char c in normalized)
                if (c == '\t' || c == '\n' || c == '\r' || c >= 0x20) clean.Append(c);
            return clean.ToString();
        }

        private sealed class SheetData
        {
            public readonly string Name;
            public readonly List<List<string>> Rows;
            public SheetData(string name, List<List<string>> rows) { Name = name; Rows = rows; }
        }

        private static class XlsxWriter
        {
            private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

            public static void Write(string file, List<SheetData> sheets)
            {
                using (FileStream output = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                using (var archive = new ZipArchive(output, ZipArchiveMode.Create, false))
                {
                    AddText(archive, "[Content_Types].xml", BuildContentTypes(sheets.Count));
                    AddText(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                    AddText(archive, "xl/workbook.xml", BuildWorkbook(sheets));
                    AddText(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelationships(sheets.Count));
                    for (int i = 0; i < sheets.Count; i++)
                        AddText(archive, "xl/worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml", BuildWorksheet(sheets[i]));
                }
            }

            private static void AddText(ZipArchive archive, string path, string text)
            {
                ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
                using (Stream stream = entry.Open())
                using (var writer = new StreamWriter(stream, Utf8)) writer.Write(text);
            }

            private static string BuildContentTypes(int count)
            {
                var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
                for (int i = 1; i <= count; i++) xml.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i.ToString(CultureInfo.InvariantCulture)).Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                return xml.Append("</Types>").ToString();
            }

            private static string BuildWorkbook(List<SheetData> sheets)
            {
                var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
                for (int i = 0; i < sheets.Count; i++)
                    xml.Append("<sheet name=\"").Append(XmlEscape(SanitizeSheetName(sheets[i].Name))).Append("\" sheetId=\"").Append(i + 1).Append("\" r:id=\"rId").Append(i + 1).Append("\"/>");
                return xml.Append("</sheets></workbook>").ToString();
            }

            private static string BuildWorkbookRelationships(int count)
            {
                var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 1; i <= count; i++)
                    xml.Append("<Relationship Id=\"rId").Append(i).Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet").Append(i).Append(".xml\"/>");
                return xml.Append("</Relationships>").ToString();
            }

            private static string BuildWorksheet(SheetData sheet)
            {
                var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                for (int rowIndex = 0; rowIndex < sheet.Rows.Count; rowIndex++)
                {
                    int excelRow = rowIndex + 1;
                    xml.Append("<row r=\"").Append(excelRow).Append("\">");
                    List<string> cells = sheet.Rows[rowIndex];
                    for (int col = 0; col < cells.Count; col++)
                    {
                        string value = cells[col] ?? String.Empty;
                        if (value.Length > 32767) value = value.Substring(0, 32767);
                        decimal numericValue;
                        string address = ColumnName(col + 1) + excelRow;
                        if (value.IndexOf(',') >= 0 && Decimal.TryParse(value, NumberStyles.Number, CultureInfo.GetCultureInfo("tr-TR"), out numericValue))
                            xml.Append("<c r=\"").Append(address).Append("\"><v>").Append(numericValue.ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
                        else
                            xml.Append("<c r=\"").Append(address).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">").Append(XmlEscape(value)).Append("</t></is></c>");
                    }
                    xml.Append("</row>");
                }
                return xml.Append("</sheetData></worksheet>").ToString();
            }

            private static string ColumnName(int column)
            {
                var name = new StringBuilder();
                while (column > 0)
                {
                    int digit = (column - 1) % 26;
                    name.Insert(0, (char)('A' + digit));
                    column = (column - 1) / 26;
                }
                return name.ToString();
            }

            private static string SanitizeSheetName(string value)
            {
                string name = Regex.Replace(value ?? String.Empty, "[\\\\/:*?\\[\\]]", "-").Trim('\'', ' ');
                if (name.Length == 0) name = "Sayfa";
                return name.Length > 31 ? name.Substring(0, 31) : name;
            }

            private static string XmlEscape(string value)
            {
                return System.Security.SecurityElement.Escape(value ?? String.Empty) ?? String.Empty;
            }
        }
    }

    private sealed class FilePreviewForm : Form
    {
        private readonly string path;
        private readonly Panel host = new Panel();
        private readonly Label message = new Label();
        private IPreviewHandler handler;
        private object handlerObject;
        private System.Runtime.InteropServices.ComTypes.IStream stream;
        private Image previewImage;

        public FilePreviewForm(string filePath, bool maximizeWindow = false)
        {
            path = filePath;
            Text = "Ön izleme — " + Path.GetFileName(filePath);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(560, 380);
            Size = new Size(900, 650);
            if (maximizeWindow) WindowState = FormWindowState.Maximized;
            host.Dock = DockStyle.Fill;
            host.BackColor = Color.White;
            message.Dock = DockStyle.Fill;
            message.TextAlign = ContentAlignment.MiddleCenter;
            message.Font = new Font("Segoe UI", 10F);
            message.ForeColor = Color.FromArgb(75, 85, 99);
            host.Controls.Add(message);
            Controls.Add(host);
            Shown += (s, e) => StartPreview();
            host.Resize += (s, e) => ResizePreview();
            FormClosed += (s, e) => StopPreview();
        }

        public string PreviewPath { get { return path; } }

        private void StartPreview()
        {
            try
            {
                string extension = Path.GetExtension(path);
                if (IsImagePreviewType(extension))
                {
                    ShowImagePreview();
                    return;
                }
                if (String.Equals(extension, ".udf", StringComparison.OrdinalIgnoreCase))
                {
                    ShowUdfPreview();
                    return;
                }
                if (IsTextPreviewType(extension))
                {
                    ShowTextPreview();
                    return;
                }
                string previewClsid;
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(extension + "\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}"))
                    previewClsid = key == null ? null : key.GetValue(null) as string;
                if (String.IsNullOrWhiteSpace(previewClsid))
                {
                    ShowMessage("Bu dosya türü için Windows ön izleme sağlayıcısı bulunamadı.");
                    return;
                }
                Type comType = Type.GetTypeFromCLSID(new Guid(previewClsid), true);
                handlerObject = Activator.CreateInstance(comType);
                handler = handlerObject as IPreviewHandler;
                if (handler == null) throw new InvalidOperationException("Kayıtlı sağlayıcı IPreviewHandler arabirimini desteklemiyor.");
                IInitializeWithStream initializeStream = handlerObject as IInitializeWithStream;
                if (initializeStream != null)
                {
                    IPreviewStreamFactory.Create(path, out stream);
                    initializeStream.Initialize(stream, 0);
                }
                else
                {
                    IInitializeWithFile initializeFile = handlerObject as IInitializeWithFile;
                    if (initializeFile == null) throw new InvalidOperationException("Ön izleme sağlayıcısı dosya veya akış ile başlatılamıyor.");
                    initializeFile.Initialize(path, 0);
                }
                message.Visible = false;
                NativeRect rect = GetRect();
                handler.SetWindow(host.Handle, ref rect);
                handler.DoPreview();
                ResizePreview();
                if (IsHandleCreated)
                    BeginInvoke((MethodInvoker)(() => ResizePreview()));
            }
            catch (Exception ex)
            {
                StopPreview();
                ShowMessage("Dosya ön izlenemedi. Bu dosya türü için uyumlu ve kayıtlı bir Windows ön izleme sağlayıcısı gerekli.\r\n\r\n" + ex.Message);
            }
        }

        private static bool IsImagePreviewType(string extension)
        {
            switch ((extension ?? String.Empty).ToLowerInvariant())
            {
                case ".jpg": case ".jpeg": case ".png": case ".gif": case ".bmp":
                case ".tif": case ".tiff": case ".ico": case ".jfif":
                    return true;
                default: return false;
            }
        }

        private void ShowImagePreview()
        {
            host.Controls.Clear();
            using (Image source = Image.FromFile(path)) previewImage = new Bitmap(source);
            var picture = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 247, 250),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = previewImage
            };
            host.Controls.Add(picture);
        }

        private static bool IsTextPreviewType(string extension)
        {
            switch ((extension ?? String.Empty).ToLowerInvariant())
            {
                case ".udf": case ".txt": case ".log": case ".csv": case ".xml": case ".xslt":
                case ".json": case ".html": case ".htm": case ".css": case ".js": case ".ini":
                case ".config": case ".yaml": case ".yml": case ".md": case ".sql": case ".xsd":
                    return true;
                default: return false;
            }
        }

        private void ShowTextPreview()
        {
            const long maxPreviewBytes = 20L * 1024L * 1024L;
            var info = new FileInfo(path);
            if (info.Length > maxPreviewBytes)
            {
                ShowMessage("Metin ön izlemesi 20 MB'a kadar olan dosyalar için kullanılabilir.\r\nDosya boyutu: " + info.Length.ToString("N0") + " bayt");
                return;
            }

            string content;
            using (var reader = new StreamReader(path, Encoding.UTF8, true)) content = reader.ReadToEnd();
            host.Controls.Clear();
            var text = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                WordWrap = false,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(35, 42, 52),
                Font = new Font("Consolas", 10F),
                Text = content
            };
            host.Controls.Add(text);
        }

        private void ShowUdfPreview()
        {
            const long maxPreviewBytes = 20L * 1024L * 1024L;
            string content = null;
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                byte[] signature = new byte[4];
                int signatureLength = input.Read(signature, 0, signature.Length);
                input.Position = 0;
                bool isZip = signatureLength >= 2 && signature[0] == 0x50 && signature[1] == 0x4B;
                if (isZip)
                {
                    using (var archive = new ZipArchive(input, ZipArchiveMode.Read, true))
                    {
                        ZipArchiveEntry entry = archive.Entries.FirstOrDefault(item =>
                            String.Equals(Path.GetFileName(item.FullName), "content.xml", StringComparison.OrdinalIgnoreCase));
                        if (entry == null)
                            throw new InvalidDataException("UDF paketi içinde content.xml bulunamadı.");
                        if (entry.Length > maxPreviewBytes)
                            throw new InvalidDataException("UDF belgesinin içeriği 20 MB ön izleme sınırını aşıyor.");
                        using (var entryStream = entry.Open())
                        using (var reader = new StreamReader(entryStream, Encoding.UTF8, true)) content = reader.ReadToEnd();
                    }
                }
                else
                {
                    if (input.Length > maxPreviewBytes)
                        throw new InvalidDataException("UDF dosyası 20 MB ön izleme sınırını aşıyor.");
                    using (var reader = new StreamReader(input, Encoding.UTF8, true)) content = reader.ReadToEnd();
                }
            }

            string readable = ExtractUdfText(content);
            if (String.IsNullOrWhiteSpace(readable))
                readable = "UDF paketi açıldı ancak görüntülenebilir metin bulunamadı. Bu belgeyi UYAP Doküman Editörü ile açabilirsiniz.";
            ShowPlainText(readable, true);
        }

        private static string ExtractUdfText(string content)
        {
            if (String.IsNullOrEmpty(content)) return String.Empty;
            try
            {
                var document = new XmlDocument { XmlResolver = null };
                document.LoadXml(content);
                XmlNode contentNode = document.SelectSingleNode("//*[local-name()='content']");
                string text = contentNode == null ? String.Empty : contentNode.InnerText;
                if (String.IsNullOrEmpty(text))
                {
                    var output = new StringBuilder();
                    AppendUdfReadableText(document.DocumentElement, output);
                    return output.ToString().Trim();
                }

                var paragraphStarts = new SortedSet<int>();
                foreach (XmlNode paragraph in document.SelectNodes("//*[local-name()='paragraph']"))
                {
                    XmlAttribute start = paragraph.Attributes == null ? null : paragraph.Attributes["startOffset"];
                    int offset;
                    if (start != null && Int32.TryParse(start.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out offset) && offset > 0 && offset < text.Length)
                        paragraphStarts.Add(offset);
                }
                if (paragraphStarts.Count == 0) return text.Trim();

                var formatted = new StringBuilder(text.Length + paragraphStarts.Count);
                for (int i = 0; i < text.Length; i++)
                {
                    if (paragraphStarts.Contains(i) && formatted.Length > 0 && formatted[formatted.Length - 1] != '\n' && formatted[formatted.Length - 1] != '\r')
                        formatted.AppendLine();
                    formatted.Append(text[i]);
                }
                return formatted.ToString().Trim();
            }
            catch (XmlException)
            {
                // Some older UDF files store readable text directly rather than XML.
                return content;
            }
        }

        private static void AppendUdfReadableText(XmlNode node, StringBuilder output)
        {
            if (node == null) return;
            if (node.NodeType == XmlNodeType.Text || node.NodeType == XmlNodeType.CDATA)
            {
                output.Append(node.Value);
                return;
            }
            bool block = node.NodeType == XmlNodeType.Element &&
                (node.Name.IndexOf("par", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 node.Name.IndexOf("line", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 node.Name.IndexOf("row", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 node.Name.IndexOf("page", StringComparison.OrdinalIgnoreCase) >= 0);
            int before = output.Length;
            foreach (XmlNode child in node.ChildNodes) AppendUdfReadableText(child, output);
            if (block && output.Length > before && output[output.Length - 1] != '\n') output.AppendLine();
        }

        private void ShowPlainText(string content, bool wordWrap = false)
        {
            host.Controls.Clear();
            var text = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                WordWrap = wordWrap,
                ScrollBars = wordWrap ? ScrollBars.Vertical : ScrollBars.Both,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(35, 42, 52),
                Font = new Font("Consolas", 10F),
                Text = content
            };
            host.Controls.Add(text);
        }

        private NativeRect GetRect()
        {
            return new NativeRect { Left = 0, Top = 0, Right = host.ClientSize.Width, Bottom = host.ClientSize.Height };
        }

        private void ResizePreview()
        {
            if (handler == null || host.IsDisposed || host.ClientSize.Width <= 0 || host.ClientSize.Height <= 0) return;
            try { NativeRect rect = GetRect(); handler.SetRect(ref rect); }
            catch { }
        }

        private void ShowMessage(string text)
        {
            message.Text = text;
            message.Visible = true;
        }

        private void StopPreview()
        {
            try { if (handler != null) handler.Unload(); } catch { }
            handler = null;
            if (previewImage != null) { previewImage.Dispose(); previewImage = null; }
            if (stream != null) { try { Marshal.ReleaseComObject(stream); } catch { } stream = null; }
            if (handlerObject != null && Marshal.IsComObject(handlerObject)) { try { Marshal.FinalReleaseComObject(handlerObject); } catch { } }
            handlerObject = null;
        }
    }

    private static class IPreviewStreamFactory
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateStreamOnFileEx(string file, uint mode, uint attributes, bool create, IntPtr template, out System.Runtime.InteropServices.ComTypes.IStream stream);

        public static void Create(string path, out System.Runtime.InteropServices.ComTypes.IStream stream)
        {
            int hr = SHCreateStreamOnFileEx(path, 0, 0x80, false, IntPtr.Zero, out stream);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        }
    }

    private enum FileConflictChoice
    {
        Replace,
        Skip,
        KeepBoth
    }

    private enum ArchiveExtractionTarget
    {
        Here,
        ArchiveFolder,
        ChooseFolder
    }

    private sealed class FileConflictDialog : Form
    {
        private readonly CheckBox applyToAll = new CheckBox();
        private FileConflictChoice choice = FileConflictChoice.Skip;

        private FileConflictDialog(string source, string destination)
        {
            Text = "Dosya çakışması";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(650, 300);
            Font = new Font("Segoe UI", 9F);

            var prompt = new Label
            {
                Text = "Bu klasörde aynı adlı dosya zaten var. Dosyaları karşılaştırıp bir işlem seçin."
            };
            prompt.SetBounds(16, 14, 618, 34);
            var compare = new TableLayoutPanel
            {
                Location = new Point(16, 55),
                Size = new Size(618, 150),
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0),
                BackColor = Color.FromArgb(235, 238, 242)
            };
            compare.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            compare.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            compare.Controls.Add(CreateFileSummary("KOPYALANACAK DOSYA", source), 0, 0);
            compare.Controls.Add(CreateFileSummary("MEVCUT DOSYA", destination), 1, 0);
            applyToAll.Text = "Bu seçimi kalan çakışmalara da uygula";
            applyToAll.SetBounds(18, 215, 330, 24);
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Location = new Point(16, 247),
                Size = new Size(618, 40)
            };
            AddChoiceButton(buttons, "İkisini de sakla", FileConflictChoice.KeepBoth);
            AddChoiceButton(buttons, "Atla", FileConflictChoice.Skip);
            AddChoiceButton(buttons, "Değiştir", FileConflictChoice.Replace);
            Controls.Add(prompt);
            Controls.Add(compare);
            Controls.Add(applyToAll);
            Controls.Add(buttons);
            CancelButton = buttons.Controls[1] as Button;
        }

        public static FileConflictChoice Ask(IWin32Window owner, string source, string destination, out bool useForAll)
        {
            using (var dialog = new FileConflictDialog(source, destination))
            {
                DialogResult result = dialog.ShowDialog(owner);
                useForAll = dialog.applyToAll.Checked;
                return result == DialogResult.OK ? dialog.choice : FileConflictChoice.Skip;
            }
        }

        private static Control CreateFileSummary(string title, string path)
        {
            var panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(10) };
            var heading = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(38, 51, 70)
            };
            var details = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                Text = Path.GetFileName(path) + "\r\n" + path + "\r\n\r\nBoyut: " + FormatFileSize(path) +
                    "\r\nDeğiştirilme: " + File.GetLastWriteTime(path).ToString("dd.MM.yyyy HH:mm:ss"),
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            panel.Controls.Add(details);
            panel.Controls.Add(heading);
            return panel;
        }

        private void AddChoiceButton(FlowLayoutPanel panel, string text, FileConflictChoice value)
        {
            var button = new Button { Text = text, AutoSize = true, Height = 30, Margin = new Padding(5, 3, 0, 3) };
            button.Click += (s, e) => { choice = value; DialogResult = DialogResult.OK; Close(); };
            panel.Controls.Add(button);
        }
    }

    private sealed class WorkProfileDialog : Form
    {
        private readonly TextBox nameBox = new TextBox();
        private readonly TextBox rootBox = new TextBox();
        private readonly TextBox notesBox = new TextBox();
        public string ProfileName { get { return nameBox.Text.Trim(); } }
        public string RootPath { get { return Environment.ExpandEnvironmentVariables(rootBox.Text.Trim().Trim('"')); } }
        public string Notes { get { return notesBox.Text; } }

        public WorkProfileDialog(string name = "", string rootPath = "", string notes = "")
        {
            Text = String.IsNullOrWhiteSpace(name) ? "Yeni iş profili" : "İş profilini düzenle";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(610, 405);
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;

            var intro = new Label
            {
                Text = "İşin adını ve ana klasörünü belirleyin. Profil kaydı dosyalarınızı değiştirmez.",
                Location = new Point(18, 15),
                Size = new Size(574, 24),
                ForeColor = Color.FromArgb(82, 94, 112)
            };
            var nameLabel = new Label { Text = "İş adı", AutoSize = true, Location = new Point(18, 52) };
            nameBox.SetBounds(18, 73, 574, 29);
            nameBox.Text = name;
            var rootLabel = new Label { Text = "Ana klasör", AutoSize = true, Location = new Point(18, 115) };
            rootBox.SetBounds(18, 136, 478, 29);
            rootBox.Text = rootPath;
            var browse = new Button { Text = "Gözat...", Location = new Point(504, 135), Size = new Size(88, 31) };
            browse.Click += (s, e) =>
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Bu işin ana klasörünü seçin.";
                    dialog.ShowNewFolderButton = false;
                    if (Directory.Exists(rootBox.Text)) dialog.SelectedPath = rootBox.Text;
                    else dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    if (dialog.ShowDialog(this) == DialogResult.OK) rootBox.Text = dialog.SelectedPath;
                }
            };
            var notesLabel = new Label { Text = "Açıklama (isteğe bağlı)", AutoSize = true, Location = new Point(18, 177) };
            notesBox.SetBounds(18, 198, 574, 146);
            notesBox.Multiline = true;
            notesBox.WordWrap = true;
            notesBox.ScrollBars = ScrollBars.Vertical;
            notesBox.Text = notes;
            var save = new Button { Text = "Kaydet", Location = new Point(414, 359), Size = new Size(84, 31) };
            var cancel = new Button { Text = "İptal", Location = new Point(508, 359), Size = new Size(84, 31), DialogResult = DialogResult.Cancel };
            save.Click += (s, e) =>
            {
                if (ProfileName.Length == 0)
                {
                    MessageBox.Show(this, "İş adı girin.", "İş profili", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    nameBox.Focus();
                    return;
                }
                if (RootPath.Length == 0 || !Directory.Exists(RootPath))
                {
                    MessageBox.Show(this, "Var olan bir ana klasör seçin.", "İş profili", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    rootBox.Focus();
                    return;
                }
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(intro);
            Controls.Add(nameLabel);
            Controls.Add(nameBox);
            Controls.Add(rootLabel);
            Controls.Add(rootBox);
            Controls.Add(browse);
            Controls.Add(notesLabel);
            Controls.Add(notesBox);
            Controls.Add(save);
            Controls.Add(cancel);
            AcceptButton = save;
            CancelButton = cancel;
            Shown += (s, e) =>
            {
                if (String.IsNullOrWhiteSpace(name)) nameBox.Focus();
                else { nameBox.Focus(); nameBox.SelectAll(); }
            };
        }
    }

    private sealed class ShellNewTemplate
    {
        public readonly string Extension;
        public readonly string DisplayName;
        public readonly string TemplatePath;
        public readonly byte[] Data;
        public readonly bool IsNullFile;
        public readonly string BuiltInKind;

        public ShellNewTemplate(string extension, string displayName, string templatePath, byte[] data, bool isNullFile, string builtInKind = "")
        {
            Extension = extension;
            DisplayName = displayName;
            TemplatePath = templatePath;
            Data = data;
            IsNullFile = isNullFile;
            BuiltInKind = builtInKind;
        }
    }

    private sealed class NewFileNameDialog : Form
    {
        private readonly TextBox nameBox = new TextBox();
        public string FileName { get { return nameBox.Text; } }

        public NewFileNameDialog(string typeName, string suggestedName)
        {
            Text = typeName;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(460, 130);
            Font = new Font("Segoe UI", 9F);
            var label = new Label { Text = "Yeni dosya adı:", AutoSize = true, Location = new Point(14, 14) };
            nameBox.SetBounds(14, 38, 432, 27);
            nameBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            nameBox.Text = suggestedName;
            var create = new Button { Text = "Oluştur", DialogResult = DialogResult.OK };
            create.SetBounds(264, 83, 82, 30);
            var cancel = new Button { Text = "İptal", DialogResult = DialogResult.Cancel };
            cancel.SetBounds(354, 83, 92, 30);
            Controls.Add(label);
            Controls.Add(nameBox);
            Controls.Add(create);
            Controls.Add(cancel);
            AcceptButton = create;
            CancelButton = cancel;
            Shown += (s, e) => { nameBox.Focus(); nameBox.SelectAll(); };
        }
    }

    private sealed class RenameFileDialog : Form
    {
        private readonly TextBox nameBox = new TextBox();
        public string NewName { get { return nameBox.Text; } }

        public RenameFileDialog(string currentName, bool folderName = false)
        {
            Text = "Yeniden adlandır";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(440, 125);
            Font = new Font("Segoe UI", 9F);
            var label = new Label { Text = folderName ? "Yeni klasör adı:" : "Yeni dosya adı:", AutoSize = true, Location = new Point(14, 14) };
            nameBox.SetBounds(14, 38, 412, 26);
            nameBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            nameBox.Text = currentName;
            var ok = new Button { Text = "Tamam", DialogResult = DialogResult.OK };
            ok.SetBounds(254, 79, 82, 30);
            var cancel = new Button { Text = "İptal", DialogResult = DialogResult.Cancel };
            cancel.SetBounds(344, 79, 82, 30);
            Controls.Add(label);
            Controls.Add(nameBox);
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            Shown += (s, e) =>
            {
                nameBox.Focus();
                int extensionLength = folderName ? 0 : Path.GetExtension(currentName).Length;
                nameBox.Select(0, Math.Max(0, nameBox.TextLength - extensionLength));
            };
        }
    }

    private sealed class NewFolderDialog : Form
    {
        private readonly TextBox nameBox = new TextBox();
        public string FolderName { get { return nameBox.Text; } }

        public NewFolderDialog()
        {
            Text = "Klasör ekle";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(420, 125);
            Font = new Font("Segoe UI", 9F);
            var label = new Label { Text = "Yeni klasör adı:", AutoSize = true, Location = new Point(14, 14) };
            nameBox.SetBounds(14, 38, 392, 26);
            nameBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            nameBox.Text = "Yeni klasör";
            var create = new Button { Text = "Oluştur", DialogResult = DialogResult.OK };
            create.SetBounds(224, 79, 82, 30);
            var cancel = new Button { Text = "İptal", DialogResult = DialogResult.Cancel };
            cancel.SetBounds(314, 79, 92, 30);
            Controls.Add(label);
            Controls.Add(nameBox);
            Controls.Add(create);
            Controls.Add(cancel);
            AcceptButton = create;
            CancelButton = cancel;
            Shown += (s, e) => { nameBox.Focus(); nameBox.SelectAll(); };
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileAttributeTagInformation
    {
        public uint FileAttributes;
        public uint ReparseTag;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left, Top, Right, Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage
    {
        public IntPtr HWnd;
        public uint Message;
        public UIntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public Point Point;
    }

    [ComImport, Guid("8895b1c6-b41f-4c1c-a562-0d564250836f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPreviewHandler
    {
        void SetWindow(IntPtr hwnd, ref NativeRect rect);
        void SetRect(ref NativeRect rect);
        void DoPreview();
        void Unload();
        void SetFocus();
        void QueryFocus(out IntPtr hwnd);
        void TranslateAccelerator(ref NativeMessage message);
    }

    [ComImport, Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IInitializeWithStream
    {
        void Initialize([MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream stream, uint mode);
    }

    [ComImport, Guid("b7d14566-0509-4cce-a71f-0a554233bd9b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IInitializeWithFile
    {
        void Initialize([MarshalAs(UnmanagedType.LPWStr)] string file, uint mode);
    }

    private sealed class MiniRestoreIcon : Form
    {
        private readonly Action restore;
        private readonly Action exit;
        private readonly System.Windows.Forms.Timer animationTimer;
        private readonly Stopwatch animationClock;
        private bool hovered;
        private double hoverBlend;
        private Color animatedColor;
        private long opacityPausedDuration;
        private long opacityPauseStartedAt;
        private Button iconButton;

        public MiniRestoreIcon(Action restoreAction, Action exitAction, Point location)
        {
            restore = restoreAction;
            exit = exitAction;
            Text = "Dosya Gezgini";
            FormBorderStyle = FormBorderStyle.None;
            AutoScaleMode = AutoScaleMode.None;
            AutoSize = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            Opacity = 0.30;
            Padding = Padding.Empty;
            MinimumSize = new Size(37, 37);
            MaximumSize = new Size(37, 37);
            ClientSize = new Size(37, 37);
            Size = new Size(37, 37);
            Location = location;
            animatedColor = Color.FromArgb(190, 148, 48);
            BackColor = animatedColor;
            ApplyCircularRegion();
            var button = new Button
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                Text = String.Empty,
                Font = new Font("Segoe UI Emoji", 13F),
                BackColor = animatedColor,
                UseVisualStyleBackColor = false,
                ForeColor = Color.FromArgb(38, 51, 70),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AccessibleName = "Dosya Gezginini geri getir"
            };
            iconButton = button;
            var menu = new ContextMenuStrip();
            menu.Items.Add("Dosya Gezgini'ni geri getir", null, (s, e) => restore());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Programdan çık", null, (s, e) => exit());
            ContextMenuStrip = menu;
            button.ContextMenuStrip = menu;
            button.FlatAppearance.BorderSize = 0;
            button.Paint += (s, e) => DrawFolderIcon(e.Graphics, button.ClientRectangle);
            button.Click += (s, e) => restore();
            button.MouseEnter += (s, e) => SetHovered(true);
            button.MouseLeave += (s, e) => SetHovered(false);
            var tip = new ToolTip();
            tip.SetToolTip(button, "Dosya Gezgini penceresini geri getir");
            Controls.Add(button);
            ApplyCircularRegion();
            Click += (s, e) => restore();

            animationClock = Stopwatch.StartNew();
            animationTimer = new System.Windows.Forms.Timer { Interval = 20 };
            animationTimer.Tick += (s, e) =>
            {
                UpdateAppearance();
            };
            Resize += (s, e) => ApplyCircularRegion();
            animationTimer.Start();
        }

        private void ApplyCircularRegion()
        {
            int diameter = Math.Max(1, Math.Min(ClientSize.Width, ClientSize.Height));
            using (var path = new GraphicsPath())
            {
                path.AddEllipse(0, 0, diameter, diameter);
                Region = new Region(path);
                if (iconButton != null)
                {
                    using (var buttonPath = new GraphicsPath())
                    {
                        buttonPath.AddEllipse(0, 0, Math.Max(1, iconButton.Width), Math.Max(1, iconButton.Height));
                        iconButton.Region = new Region(buttonPath);
                    }
                }
            }
        }

        private void UpdateAppearance()
        {
            long animationElapsed = animationClock.ElapsedMilliseconds;
            Color[] palette = new Color[]
            {
                Color.FromArgb(190, 148, 48),
                Color.FromArgb(174, 78, 82),
                Color.FromArgb(72, 122, 176),
                Color.FromArgb(76, 145, 96)
            };
            double scaledColorPosition = (animationElapsed % 13000) * palette.Length / 13000.0;
            int colorSegment = (int)scaledColorPosition;
            double colorProgress = SmoothStep(scaledColorPosition - colorSegment);
            Color color = Blend(palette[colorSegment], palette[(colorSegment + 1) % palette.Length], colorProgress);

            double targetHover = hovered ? 1.0 : 0.0;
            hoverBlend += (targetHover - hoverBlend) * 0.12;
            color = Blend(color, Color.FromArgb(72, 122, 176), hoverBlend);

            if (hovered)
            {
                Opacity = 1.0;
            }
            else
            {
                long opacityElapsed = Math.Max(0, animationElapsed - opacityPausedDuration);
                long phase = opacityElapsed % 13000;
                if (phase < 5000)
                {
                    double progress = SmoothStep(phase / 5000.0);
                    Opacity = 0.75 - progress * 0.55;
                }
                else if (phase < 8000)
                {
                    Opacity = 0.20;
                }
                else
                {
                    double progress = SmoothStep((phase - 8000) / 5000.0);
                    Opacity = 0.20 + progress * 0.55;
                }
            }
            animatedColor = color;
            BackColor = color;
            if (iconButton != null)
            {
                iconButton.BackColor = color;
                iconButton.Invalidate();
            }
        }

        private static double SmoothStep(double value)
        {
            value = Math.Max(0.0, Math.Min(1.0, value));
            return value * value * (3.0 - 2.0 * value);
        }

        private static Color Blend(Color first, Color second, double amount)
        {
            amount = Math.Max(0.0, Math.Min(1.0, amount));
            return Color.FromArgb(
                (int)(first.R + (second.R - first.R) * amount),
                (int)(first.G + (second.G - first.G) * amount),
                (int)(first.B + (second.B - first.B) * amount));
        }

        private static void DrawFolderIcon(Graphics graphics, Rectangle bounds)
        {
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            float width = 22F;
            float left = (bounds.Width - width) / 2F;
            float top = (bounds.Height - 17F) / 2F + 1F;
            using (var folder = new GraphicsPath())
            using (var fill = new SolidBrush(Color.FromArgb(255, 221, 105)))
            using (var outline = new Pen(Color.FromArgb(54, 47, 31), 1.5F))
            {
                folder.StartFigure();
                folder.AddLine(left + 1F, top + 4F, left + 8F, top + 4F);
                folder.AddLine(left + 8F, top + 4F, left + 10F, top + 1F);
                folder.AddLine(left + 10F, top + 1F, left + 20F, top + 1F);
                folder.AddLine(left + 20F, top + 1F, left + 21F, top + 15F);
                folder.AddLine(left + 21F, top + 15F, left + 1F, top + 15F);
                folder.CloseFigure();
                graphics.FillPath(fill, folder);
                graphics.DrawPath(outline, folder);
                using (var tab = new SolidBrush(Color.FromArgb(255, 231, 143)))
                    graphics.FillRectangle(tab, left + 3F, top + 2F, 6F, 2F);
            }
        }

        private void SetHovered(bool value)
        {
            if (value == hovered) return;
            long now = animationClock == null ? 0 : animationClock.ElapsedMilliseconds;
            if (value)
                opacityPauseStartedAt = now;
            else if (opacityPauseStartedAt > 0)
            {
                opacityPausedDuration += Math.Max(0, now - opacityPauseStartedAt);
                opacityPauseStartedAt = 0;
            }
            hovered = value;
            if (hovered) Opacity = 1.0;
        }

        private static Color FromHsv(double hue, double saturation, double value)
        {
            double chroma = value * saturation;
            double h = (hue % 360.0) / 60.0;
            double x = chroma * (1.0 - Math.Abs(h % 2.0 - 1.0));
            double r = 0, g = 0, b = 0;
            if (h < 1) { r = chroma; g = x; }
            else if (h < 2) { r = x; g = chroma; }
            else if (h < 3) { g = chroma; b = x; }
            else if (h < 4) { g = x; b = chroma; }
            else if (h < 5) { r = x; b = chroma; }
            else { r = chroma; b = x; }
            double m = value - chroma;
            return Color.FromArgb((int)((r + m) * 255), (int)((g + m) * 255), (int)((b + m) * 255));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (animationTimer != null) animationTimer.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

[ComVisible(true)]
[Guid("00000122-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IExplorerOleDropTarget
{
    [PreserveSig]
    int DragEnter([In] System.Runtime.InteropServices.ComTypes.IDataObject data, uint keyState, POINTL point, ref uint effect);
    [PreserveSig]
    int DragOver(uint keyState, POINTL point, ref uint effect);
    [PreserveSig]
    int DragLeave();
    [PreserveSig]
    int Drop([In] System.Runtime.InteropServices.ComTypes.IDataObject data, uint keyState, POINTL point, ref uint effect);
}

[ComVisible(true)]
[StructLayout(LayoutKind.Sequential)]
public struct POINTL
{
    public int x;
    public int y;
}

[ComImport]
[Guid("3D8B0590-F691-11d2-8EA9-006097DF5BD4")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDataObjectAsyncCapability
{
    [PreserveSig]
    int SetAsyncMode([In, MarshalAs(UnmanagedType.Bool)] bool asyncMode);
    [PreserveSig]
    int GetAsyncMode([Out, MarshalAs(UnmanagedType.Bool)] out bool asyncMode);
    [PreserveSig]
    int StartOperation([In] System.Runtime.InteropServices.ComTypes.IBindCtx reserved);
    [PreserveSig]
    int InOperation([Out, MarshalAs(UnmanagedType.Bool)] out bool inOperation);
    [PreserveSig]
    int EndOperation([In] int result, [In] System.Runtime.InteropServices.ComTypes.IBindCtx reserved, uint effects);
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class OutlookAsyncDropTarget : IExplorerOleDropTarget, IDisposable
{
    private const int S_OK = 0;
    private const int E_FAIL = unchecked((int)0x80004005);
    private const int E_UNEXPECTED = unchecked((int)0x8000FFFF);
    private const uint DROPEFFECT_NONE = 0;
    private const uint DROPEFFECT_COPY = 1;
    private const uint DROPEFFECT_MOVE = 2;
    private const uint MK_CONTROL = 0x0008;
    private readonly Control control;
    private readonly Action<string[], bool> dropHandler;
    private readonly Action<string> statusHandler;
    private IDataObjectAsyncCapability asyncCapability;
    private IntPtr registeredHandle;
    private bool isAsyncOperation;
    private bool isRegistered;
    private bool oleInitialized;
    private bool disposed;

    [DllImport("ole32.dll")]
    private static extern int RegisterDragDrop(IntPtr hwnd, IExplorerOleDropTarget target);
    [DllImport("ole32.dll")]
    private static extern int RevokeDragDrop(IntPtr hwnd);
    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr reserved);
    [DllImport("ole32.dll")]
    private static extern void OleUninitialize();
    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref System.Runtime.InteropServices.ComTypes.STGMEDIUM medium);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint DragQueryFile(IntPtr drop, uint index, StringBuilder fileName, uint fileNameLength);

    public OutlookAsyncDropTarget(Control target, Action<string[], bool> onDrop, Action<string> onStatus)
    {
        if (target == null) throw new ArgumentNullException("target");
        if (onDrop == null) throw new ArgumentNullException("onDrop");
        control = target;
        dropHandler = onDrop;
        statusHandler = onStatus;
        control.AllowDrop = false;
        control.HandleCreated += ControlHandleCreated;
        control.HandleDestroyed += ControlHandleDestroyed;
        control.Disposed += ControlDisposed;
        if (control.IsHandleCreated) Register();
    }

    private void ControlHandleCreated(object sender, EventArgs e) { Register(); }
    private void ControlHandleDestroyed(object sender, EventArgs e) { Unregister(); }
    private void ControlDisposed(object sender, EventArgs e) { Dispose(); }

    private void Register()
    {
        if (disposed || isRegistered || control.IsDisposed || !control.IsHandleCreated) return;
        int oleResult = OleInitialize(IntPtr.Zero);
        oleInitialized = oleResult == S_OK || oleResult == 1;
        if (!oleInitialized) return;
        int result = RegisterDragDrop(control.Handle, this);
        if (result == S_OK) { isRegistered = true; registeredHandle = control.Handle; }
        else
        {
            OleUninitialize();
            oleInitialized = false;
            if (statusHandler != null) statusHandler("Sürükle bırak hedefi kaydedilemedi (0x" + result.ToString("X8") + ").");
        }
    }

    private void Unregister()
    {
        if (isRegistered && registeredHandle != IntPtr.Zero)
        {
            RevokeDragDrop(registeredHandle);
            isRegistered = false;
            registeredHandle = IntPtr.Zero;
        }
        if (oleInitialized)
        {
            OleUninitialize();
            oleInitialized = false;
        }
        asyncCapability = null;
        isAsyncOperation = false;
    }

    public int DragEnter(System.Runtime.InteropServices.ComTypes.IDataObject data, uint keyState, POINTL point, ref uint effect)
    {
        if (disposed || data == null) { effect = DROPEFFECT_NONE; return E_UNEXPECTED; }
        asyncCapability = data as IDataObjectAsyncCapability;
        isAsyncOperation = false;
        if (asyncCapability != null)
        {
            try
            {
                asyncCapability.GetAsyncMode(out isAsyncOperation);
                if (isAsyncOperation) asyncCapability.SetAsyncMode(true);
            }
            catch { isAsyncOperation = false; }
        }
        effect = GetEffect(effect, keyState, isAsyncOperation);
        return S_OK;
    }

    public int DragOver(uint keyState, POINTL point, ref uint effect)
    {
        effect = GetEffect(effect, keyState, isAsyncOperation);
        return S_OK;
    }

    public int DragLeave()
    {
        asyncCapability = null;
        isAsyncOperation = false;
        return S_OK;
    }

    public int Drop(System.Runtime.InteropServices.ComTypes.IDataObject data, uint keyState, POINTL point, ref uint effect)
    {
        if (disposed || data == null) { effect = DROPEFFECT_NONE; return E_UNEXPECTED; }
        IDataObjectAsyncCapability capability = asyncCapability ?? data as IDataObjectAsyncCapability;
        bool asynchronous = isAsyncOperation && capability != null;
        bool move = !asynchronous && (keyState & MK_CONTROL) == 0 && (effect & DROPEFFECT_MOVE) != 0;
        asyncCapability = null;
        isAsyncOperation = false;

        if (asynchronous)
        {
            int start = capability.StartOperation(null);
            if (start < 0) { effect = DROPEFFECT_NONE; return start; }
            if (statusHandler != null) statusHandler("Outlook eki hazırlanıyor...");
            System.Threading.ThreadPool.QueueUserWorkItem(_ => CompleteAsyncDrop(data, capability));
            effect = DROPEFFECT_COPY;
            return S_OK;
        }

        try
        {
            string[] paths = ReadHDropPaths(data);
            if (paths.Length == 0) { effect = DROPEFFECT_NONE; return E_FAIL; }
            dropHandler(paths, move);
            effect = move ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
            return S_OK;
        }
        catch { effect = DROPEFFECT_NONE; return E_FAIL; }
    }

    private void CompleteAsyncDrop(System.Runtime.InteropServices.ComTypes.IDataObject data, IDataObjectAsyncCapability capability)
    {
        int result = S_OK;
        try
        {
            string[] paths = ReadHDropPaths(data);
            if (paths.Length == 0) throw new IOException("Outlook sürükleme sırasında dosya oluşturmadı.");
            using (var completed = new ManualResetEvent(false))
            {
                Exception callbackError = null;
                if (control.IsDisposed || !control.IsHandleCreated) throw new ObjectDisposedException("Dosyalar");
                control.BeginInvoke((MethodInvoker)(() =>
                {
                    try { dropHandler(paths, false); }
                    catch (Exception ex) { callbackError = ex; }
                    finally { completed.Set(); }
                }));
                if (!completed.WaitOne(120000)) throw new TimeoutException("Outlook eki beklenirken işlem zaman aşımına uğradı.");
                if (callbackError != null) throw callbackError;
            }
            if (statusHandler != null) statusHandler(paths.Length + " mail eki dosya olarak aktarıldı.");
        }
        catch (Exception ex)
        {
            result = E_FAIL;
            if (statusHandler != null) statusHandler("Outlook eki aktarılamadı: " + ex.GetBaseException().Message);
        }
        finally
        {
            try { capability.EndOperation(result, null, result == S_OK ? DROPEFFECT_COPY : DROPEFFECT_NONE); }
            catch { }
        }
    }

    private static uint GetEffect(uint allowed, uint keyState, bool asynchronous)
    {
        if (asynchronous) return (allowed & DROPEFFECT_COPY) != 0 ? DROPEFFECT_COPY : DROPEFFECT_NONE;
        if ((keyState & MK_CONTROL) != 0) return (allowed & DROPEFFECT_COPY) != 0 ? DROPEFFECT_COPY : DROPEFFECT_NONE;
        if ((allowed & DROPEFFECT_MOVE) != 0) return DROPEFFECT_MOVE;
        if ((allowed & DROPEFFECT_COPY) != 0) return DROPEFFECT_COPY;
        return DROPEFFECT_NONE;
    }

    private static string[] ReadHDropPaths(System.Runtime.InteropServices.ComTypes.IDataObject data)
    {
        var format = new System.Runtime.InteropServices.ComTypes.FORMATETC
        {
            cfFormat = 15,
            dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL
        };
        System.Runtime.InteropServices.ComTypes.STGMEDIUM medium;
        data.GetData(ref format, out medium);
        try
        {
            if (medium.tymed != System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero)
                throw new InvalidDataException("Outlook CF_HDROP verisi boş.");
            uint count = DragQueryFile(medium.unionmember, UInt32.MaxValue, null, 0);
            var paths = new List<string>();
            for (uint index = 0; index < count; index++)
            {
                uint length = DragQueryFile(medium.unionmember, index, null, 0);
                if (length == 0 || length > 32767) continue;
                var buffer = new StringBuilder((int)length + 1);
                if (DragQueryFile(medium.unionmember, index, buffer, (uint)buffer.Capacity) > 0)
                    paths.Add(buffer.ToString());
            }
            return paths.ToArray();
        }
        finally { ReleaseStgMedium(ref medium); }
    }

    public void Dispose()
    {
        if (disposed) return;
        if (control != null)
        {
            control.HandleCreated -= ControlHandleCreated;
            control.HandleDestroyed -= ControlHandleDestroyed;
            control.Disposed -= ControlDisposed;
        }
        Unregister();
        disposed = true;
    }
}
