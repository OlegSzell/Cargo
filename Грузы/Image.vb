Imports System.IO
Imports System.Runtime.CompilerServices
Imports AxSHDocVw

Public Class ImageForm

    Private ФормаЗапуска As Boolean = False
    Public Property Imagese As Byte() = Nothing
    Private IDp As Integer
    Private Read As Boolean = False

    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Sub New(ByVal _ФормаЗапуска As Boolean, Optional ID As Integer = 0, Optional _Read As Boolean = False)
        ФормаЗапуска = _ФормаЗапуска
        IDp = ID
        Read = _Read
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ReadImg()
    End Sub
    Private Sub ReadImg()
        Using db As New dbAllDataContext()
            Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = IDp).FirstOrDefault()
            If f IsNot Nothing Then
                Dim df = byteArrayToImage(f.ФотоДанные.ToArray)
                Clipboard.Clear()
                Clipboard.SetImage(df)
                RichTextBox1.Text = String.Empty
                RichTextBox1.Paste()
                Label2.Text = f.ДатаИзменения
                'IO.File.WriteAllBytes(RichTextBox1.Text, f.ФотоДанные.ToArray)
            End If
        End Using
    End Sub

    Public Shared Function imageToByteArray(ByVal imageIn As System.Drawing.Image) As Byte()
        Dim ms As MemoryStream = New MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
        Return ms.ToArray()
    End Function

    Public Shared Function byteArrayToImage(ByVal byteArrayIn As Byte()) As Image
        Dim ms As MemoryStream = New MemoryStream(byteArrayIn)
        Dim returnImage As Image = Image.FromStream(ms)
        Return returnImage
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If Clipboard.ContainsImage() = False Then Return

        Dim returnImage As Image = Nothing
        returnImage = My.Computer.Clipboard.GetImage()
        'Dim bt = IO.File.ReadAllBytes(RichTextBox1.Text)
        Dim mo As New AllUpd
        If ФормаЗапуска = False Then
            Using db As New dbAllDataContext()
                Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = IDp).FirstOrDefault()
                If f IsNot Nothing Then
                    f.ФотоДанные = imageToByteArray(returnImage)
                    f.ДатаИзменения = Now
                    db.SubmitChanges()
                    mo.ПеревозчикиБазаAllAsync()
                    MessageBox.Show("Фото принято!", Рик)
                    RichTextBox1.Text = String.Empty
                    Clipboard.Clear()
                End If
            End Using

        Else
            Imagese = imageToByteArray(returnImage)
            MessageBox.Show("Фото принято!", Рик)
            RichTextBox1.Text = String.Empty
            Clipboard.Clear()
        End If


    End Sub








    Private Sub ImageForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If IDp > 0 Then
            Button2.Enabled = True
        End If
        If Read = True Then
            ReadImg()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Clipboard.ContainsImage() = False Then
            MessageBox.Show("Нет данных для вставки!", Рик)
            Return
        End If
        RichTextBox1.Text = String.Empty
        RichTextBox1.Paste()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RichTextBox1.Text = String.Empty

    End Sub
    'Private Sub InitializeBrowserEvents(SourceBrowser As ExtendedWebBrowser)
    '    SourceBrowser.NewWindow2 += New EventHandler(Of NewWindow2EventArgs)(AddressOf SourceBrowser_NewWindow2)
    'End Sub
    'Private Sub SourceBrowser_NewWindow2(sender As Object, e As NewWindow2EventArgs)
    '    Dim NewTabPage As New TabPage() With { Key.Text = "Loading..."}
    '    Dim NewTabBrowser As New ExtendedWebBrowser() With { Key.Parent = NewTabPage, Key.Dock = DockStyle.Fill, Key.Tag = NewTabPage}
    '    '
    '    e.PPDisp = NewTabBrowser.Application
    '    InitializeBrowserEvents(NewTabBrowser)
    '    '
    '    Tabs.TabPages.Add(NewTabPage)
    '    Tabs.SelectedTab = NewTabPage
    'End Sub

    'Private Sub ImageForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    InitializeBrowserEvents(InitialTabBrowser)
    'End Sub

    'Public Class ExtendedWebBrowserConfiguration
    '    Private _emulationVersion As Integer

    '    Public Property EmulationVersion As Integer
    '        Get
    '            Return _emulationVersion
    '        End Get
    '        Set(ByVal value As Integer)

    '            If value >= 7 AndAlso value <= 11 Then
    '                _emulationVersion = value
    '                Return
    '            End If

    '            Throw New ArgumentException("The provided value is invalid. Valid versions are values from 7 through 11.")
    '        End Set
    '    End Property

    '    <FeatureControl("FEATURE_RESTRICT_ABOUT_PROTOCOL_IE7")>
    '    Public Property AboutProtocolRestriction As Boolean
    '    <FeatureControl("FEATURE_SAFE_BINDTOOBJECT")>
    '    Public Property ActiveXBindingSafetyChecks As Boolean
    '    <FeatureControl("FEATURE_OBJECT_CACHING")>
    '    Public Property ActiveXObjectCaching As Boolean
    '    <FeatureControl("FEATURE_RESTRICT_ACTIVEXINSTALL")>
    '    Public Property ActiveXUpdateRestriction As Boolean
    '    <FeatureControl("FEATURE_FORCE_ADDR_AND_STATUS")>
    '    Public Property AddressAndStatusBarDisplay As Boolean
    '    <FeatureControl("FEATURE_AJAX_CONNECTIONEVENTS")>
    '    Public Property AjaxConnectionEvents As Boolean
    '    <FeatureControl("FEATURE_SHOW_APP_PROTOCOL_WARN_DIALOG")>
    '    Public Property ApplicationProtocolConfirmation As Boolean
    '    <FeatureControl("FEATURE_BEHAVIORS")>
    '    Public Property BinaryBehaviorSecurity As Boolean

    '    <FeatureControl("FEATURE_BROWSER_EMULATION")>
    '    Public Property BrowserEmulation As UInteger
    '        Get
    '            Return CUInt(EmulationVersion) * 1000
    '        End Get
    '        Set(ByVal value As UInteger)
    '            EmulationVersion = CInt(value) / 1000
    '        End Set
    '    End Property

    '    <FeatureControl("FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION")>
    '    Public Property ChildWindowClipping As Boolean
    '    <FeatureControl("FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS")>
    '    Public Property CircularReferencesInScriptManagement As Boolean
    '    <FeatureControl("FEATURE_ENABLE_SCRIPT_PASTE_URLACTION_IF_PROMPT")>
    '    Public Property ClipboardScriptControl As Boolean
    '    <FeatureControl("FEATURE_BLOCK_SETCAPTURE_XDOMAIN")>
    '    Public Property CrossDomainCaptureEvent As Boolean
    '    <FeatureControl("FEATURE_CROSS_DOMAIN_REDIRECT_MITIGATION")>
    '    Public Property CrossDomainRedirection As Boolean
    '    <FeatureControl("FEATURE_DOWNLOAD_INITIATOR_HTTP_HEADER")>
    '    Public Property DebuggingNetworkTrafficRequests As Boolean
    '    <FeatureControl("FEATURE_DOMSTORAGE")>
    '    Public Property DomWebStorageApiSupport As Boolean
    '    <FeatureControl("FEATURE_CFSTR_INETURLW_DRAGDROP_FORMAT")>
    '    Public Property DragAndDropUrlFormat As Boolean
    '    <FeatureControl("FEATURE_FEEDS")>
    '    Public Property Feeds As Boolean
    '    <FeatureControl("FEATURE_RESTRICT_FILEDOWNLOAD")>
    '    Public Property FileDownloadRestrictions As Boolean
    '    <FeatureControl("FEATURE_BLOCK_CROSS_PROTOCOL_FILE_NAVIGATION")>
    '    Public Property FileProtocolNavigation As Boolean
    '    <FeatureControl("FEATURE_IE6_DEFAULT_FRAME_NAVIGATION_BEHAVIOR")>
    '    Public Property FrameContentModification As Boolean
    '    <FeatureControl("FEATURE_VIEWLINKEDWEBOC_IS_UNSAFE")>
    '    Public Property FrameContentSecurity As Boolean
    '    <FeatureControl("FEATURE_GPU_RENDERING")>
    '    Public Property GpuRendering As Boolean
    '    <FeatureControl("FEATURE_MAXCONNECTIONSPER1_0SERVER")>
    '    Public Property Http10ConnectionMaximum As UInteger
    '    <FeatureControl("MAXCONNECTIONSPERSERVER")>
    '    Public Property Http11ConnectionMaximum As UInteger
    '    <FeatureControl("FEATURE_IFRAME_MAILTO_THRESHOLD")>
    '    Public Property IFrameMailToThreshold As Boolean
    '    <FeatureControl("FEATURE_MIME_TREAT_IMAGE_AS_AUTHORITATIVE")>
    '    Public Property ImageMimeTypeDetermination As Boolean
    '    <FeatureControl("FEATURE_SECURITYBAND")>
    '    Public Property InformationBarHandling As Boolean
    '    <FeatureControl("FEATURE_BLOCK_INPUT_PROMPTS")>
    '    Public Property InputPromptBlocking As Boolean
    '    <FeatureControl("FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI")>
    '    Public Property IViewObjectLegacyDrawing As Boolean
    '    <FeatureControl("FEATURE_NINPUT_LEGACYMODE")>
    '    Public Property LegacyInputModel As Boolean
    '    <FeatureControl("FEATURE_DISABLE_LEGACY_COMPRESSION")>
    '    Public Property LegacyCompressionSupport As Boolean
    '    <FeatureControl("FEATURE_LOCALMACHINE_LOCKDOWN")>
    '    Public Property LocalMachineLockdown As Boolean
    '    <FeatureControl("FEATURE_BLOCK_LMZ_IMG")>
    '    Public Property LocalImageBlocking As Boolean
    '    <FeatureControl("FEATURE_BLOCK_LMZ_OBJECT")>
    '    Public Property LocalObjectBlocking As Boolean
    '    <FeatureControl("FEATURE_BLOCK_LMZ_SCRIPT")>
    '    Public Property LocalScriptBlocking As Boolean
    '    <FeatureControl("FEATURE_MIME_SNIFFING")>
    '    Public Property MimeTypeDetermination As Boolean
    '    <FeatureControl("FEATURE_MIME_HANDLING")>
    '    Public Property MimeTypeHandling As Boolean
    '    <FeatureControl("FEATURE_DISABLE_MK_PROTOCOL")>
    '    Public Property MKProtocolSupport As Boolean
    '    <FeatureControl("FEATURE_ISOLATE_NAMED_WINDOWS")>
    '    Public Property NamedWindowIsolation As Boolean
    '    <FeatureControl("FEATURE_DISABLE_NAVIGATION_SOUNDS")>
    '    Public Property NavigationSoundSupport As Boolean
    '    <FeatureControl("FEATURE_PROTOCOL_LOCKDOWN")>
    '    Public Property ProtocolLockdown As Boolean
    '    <FeatureControl("FEATURE_RESTRICT_ACTIVEXINSTALL")>
    '    Public Property ResourceProtocolRestriction As Boolean
    '    <FeatureControl("FEATURE_DOWNLOAD_PROMPT_META_CONTROL")>
    '    Public Property SaveDialogButtonHiding As Boolean
    '    <FeatureControl("FEATURE_SCRIPTURL_MITIGATION")>
    '    Public Property ScriptUrlMitigation As Boolean
    '    <FeatureControl("FEATURE_WARN_ON_SEC_CERT_REV_FAILED")>
    '    Public Property SecurityCertificateRevocationFailure As Boolean
    '    <FeatureControl("FEATURE_LOAD_SHDOCLC_RESOURCES")>
    '    Public Property ShdoclcDllResourceLoading As Boolean
    '    <FeatureControl("FEATURE_SPELLCHECKING")>
    '    Public Property SpellcheckAndAutoCorrectSupport As Boolean
    '    <FeatureControl("FEATURE_SSLUX")>
    '    Public Property SslSecurityAlertDisplay As Boolean
    '    <FeatureControl("FEATURE_STATUS_BAR_THROTTLING")>
    '    Public Property StatusBarUpdateFrequency As Boolean
    '    <FeatureControl("FEATURE_RESTRICT_CDL_CLSIDSNIFF")>
    '    Public Property StructuredStorageDetection As Boolean
    '    <FeatureControl("FEATURE_TABBED_BROWSING")>
    '    Public Property TabbedBrowsingShortcutsAndNotifications As Boolean
    '    <FeatureControl("FEATURE_DISABLE_TELNET_PROTOCOL")>
    '    Public Property TelnetProtocolSupport As Boolean
    '    <FeatureControl("FEATURE_UNC_SAVEDFILECHECK")>
    '    Public Property UncFileSupportForMotW As Boolean
    '    <FeatureControl("FEATURE_HTTP_USERNAME_PASSWORD_DISABLE")>
    '    Public Property UsernamesAndPasswordsInUrls As Boolean
    '    <FeatureControl("FEATURE_VALIDATE_NAVIGATE_URL")>
    '    Public Property ValidateUrlNavigation As Boolean
    '    <FeatureControl("FEATURE_SHIM_MSHELP_COMBINE")>
    '    Public Property VisualStudioLegacyHelpSupport As Boolean
    '    <FeatureControl("FEATURE_WEBOC_DOCUMENT_ZOOM")>
    '    Public Property WebBrowserControlDocumentZoom As Boolean
    '    <FeatureControl("FEATURE_WEBOC_POPUPMANAGEMENT")>
    '    Public Property WebBrowserControlPopupManagement As Boolean
    '    <FeatureControl("FEATURE_WEBOC_MOVESIZECHILD")>
    '    Public Property WebBrowserControlWindowControl As Boolean
    '    <FeatureControl("FEATURE_ENABLE_WEB_CONTROL_VISUALS")>
    '    Public Property WebControlVisuals As Boolean
    '    <FeatureControl("FEATURE_ADDON_MANAGEMENT")>
    '    Public Property WebOcAddonManagement As Boolean
    '    <FeatureControl("FEATURE_WEBSOCKET")>
    '    Public Property WebSocket As Boolean
    '    <FeatureControl("FEATURE_WEBSOCKET_AUTHPROMPT")>
    '    Public Property WebSocketAuthenticationPrompt As Boolean
    '    <FeatureControl("FEATURE_WEBSOCKET_CLOSETIMEOUT")>
    '    Public Property WebSocketCloseTimeout As UInteger
    '    <FeatureControl("FEATURE_WEBSOCKET_MAXCONNECTIONSPERSERVER")>
    '    Public Property WebSocketMaximumServerConnections As UInteger
    '    <FeatureControl("FEATURE_WEBSOCKET_FOLLOWHTTPREDIRECT")>
    '    Public Property WebSocketFollowRedirects As Boolean
    '    <FeatureControl("FEATURE_WINDOW_RESTRICTIONS")>
    '    Public Property WindowRestrictions As Boolean
    '    <FeatureControl("FEATURE_XDOMAINREQUEST")>
    '    Public Property XDomainRequestObjectSupport As Boolean
    '    <FeatureControl("FEATURE_XMLHTTP")>
    '    Public Property XmlHttpRequestObjectSupport As Boolean
    '    <FeatureControl("FEATURE_ZONE_ELEVATION")>
    '    Public Property ZoneElevation As Boolean
    '    <FeatureControl("FEATURE_RESTRICTED_ZONE_WHEN_FILE_NOT_FOUND")>
    '    Public Property ZoneHandlingForMissingFiles As Boolean
    '    <FeatureControl("FEATURE_READ_ZONE_STRINGS_FROM_REGISTRY")>
    '    Public Property ZoneStringLoading As Boolean
    '    Public Shared ReadOnly WebBrowserControlDefault As ExtendedWebBrowserConfiguration = New ExtendedWebBrowserConfiguration With {
    '    .AboutProtocolRestriction = False,
    '    .ActiveXBindingSafetyChecks = True,
    '    .ActiveXObjectCaching = True,
    '    .ActiveXUpdateRestriction = False,
    '    .AddressAndStatusBarDisplay = False,
    '    .AjaxConnectionEvents = False,
    '    .ApplicationProtocolConfirmation = False,
    '    .BinaryBehaviorSecurity = True,
    '    .BrowserEmulation = 7000UI,
    '    .ChildWindowClipping = True,
    '    .CircularReferencesInScriptManagement = True,
    '    .ClipboardScriptControl = True,
    '    .CrossDomainCaptureEvent = True,
    '    .CrossDomainRedirection = True,
    '    .DebuggingNetworkTrafficRequests = False,
    '    .DomWebStorageApiSupport = True,
    '    .DragAndDropUrlFormat = True,
    '    .Feeds = False,
    '    .FileDownloadRestrictions = False,
    '    .FileProtocolNavigation = False,
    '    .FrameContentModification = False,
    '    .FrameContentSecurity = False,
    '    .GpuRendering = False,
    '    .Http10ConnectionMaximum = 6,
    '    .Http11ConnectionMaximum = 6,
    '    .IFrameMailToThreshold = False,
    '    .ImageMimeTypeDetermination = True,
    '    .InformationBarHandling = False,
    '    .InputPromptBlocking = False,
    '    .IViewObjectLegacyDrawing = True,
    '    .LegacyInputModel = True,
    '    .LegacyCompressionSupport = True,
    '    .LocalMachineLockdown = False,
    '    .LocalImageBlocking = False,
    '    .LocalObjectBlocking = False,
    '    .LocalScriptBlocking = False,
    '    .MimeTypeDetermination = True,
    '    .MimeTypeHandling = False,
    '    .MKProtocolSupport = True,
    '    .NamedWindowIsolation = True,
    '    .NavigationSoundSupport = False,
    '    .ProtocolLockdown = False,
    '    .ResourceProtocolRestriction = False,
    '    .SaveDialogButtonHiding = True,
    '    .ScriptUrlMitigation = False,
    '    .SecurityCertificateRevocationFailure = False,
    '    .ShdoclcDllResourceLoading = False,
    '    .SpellcheckAndAutoCorrectSupport = False,
    '    .SslSecurityAlertDisplay = False,
    '    .StatusBarUpdateFrequency = False,
    '    .StructuredStorageDetection = False,
    '    .TabbedBrowsingShortcutsAndNotifications = False,
    '    .TelnetProtocolSupport = False,
    '    .UncFileSupportForMotW = False,
    '    .UsernamesAndPasswordsInUrls = False,
    '    .ValidateUrlNavigation = False,
    '    .VisualStudioLegacyHelpSupport = True,
    '    .WebBrowserControlDocumentZoom = False,
    '    .WebBrowserControlPopupManagement = True,
    '    .WebBrowserControlWindowControl = False,
    '    .WebControlVisuals = False,
    '    .WebOcAddonManagement = False,
    '    .WebSocket = True,
    '    .WebSocketAuthenticationPrompt = False,
    '    .WebSocketCloseTimeout = 15000,
    '    .WebSocketMaximumServerConnections = 6,
    '    .WebSocketFollowRedirects = False,
    '    .WindowRestrictions = True,
    '    .XDomainRequestObjectSupport = True,
    '    .XmlHttpRequestObjectSupport = True,
    '    .ZoneElevation = True,
    '    .ZoneHandlingForMissingFiles = False,
    '    .ZoneStringLoading = False
    '}
    '    Public Shared ReadOnly InternetExplorerDefault As ExtendedWebBrowserConfiguration = New ExtendedWebBrowserConfiguration With {
    '    .AboutProtocolRestriction = True,
    '    .ActiveXBindingSafetyChecks = True,
    '    .ActiveXObjectCaching = True,
    '    .ActiveXUpdateRestriction = False,
    '    .AddressAndStatusBarDisplay = True,
    '    .AjaxConnectionEvents = True,
    '    .ApplicationProtocolConfirmation = True,
    '    .BinaryBehaviorSecurity = True,
    '    .BrowserEmulation = 11000UI,
    '    .ChildWindowClipping = True,
    '    .CircularReferencesInScriptManagement = True,
    '    .ClipboardScriptControl = False,
    '    .CrossDomainCaptureEvent = True,
    '    .CrossDomainRedirection = True,
    '    .DebuggingNetworkTrafficRequests = False,
    '    .DomWebStorageApiSupport = True,
    '    .DragAndDropUrlFormat = True,
    '    .Feeds = True,
    '    .FileDownloadRestrictions = False,
    '    .FileProtocolNavigation = True,
    '    .FrameContentModification = False,
    '    .FrameContentSecurity = True,
    '    .GpuRendering = True,
    '    .Http10ConnectionMaximum = 6,
    '    .Http11ConnectionMaximum = 6,
    '    .IFrameMailToThreshold = True,
    '    .ImageMimeTypeDetermination = True,
    '    .InformationBarHandling = True,
    '    .InputPromptBlocking = True,
    '    .IViewObjectLegacyDrawing = True,
    '    .LegacyInputModel = False,
    '    .LegacyCompressionSupport = True,
    '    .LocalMachineLockdown = True,
    '    .LocalImageBlocking = True,
    '    .LocalObjectBlocking = True,
    '    .LocalScriptBlocking = True,
    '    .MimeTypeDetermination = True,
    '    .MimeTypeHandling = True,
    '    .MKProtocolSupport = True,
    '    .NamedWindowIsolation = True,
    '    .NavigationSoundSupport = False,
    '    .ProtocolLockdown = False,
    '    .ResourceProtocolRestriction = True,
    '    .SaveDialogButtonHiding = True,
    '    .ScriptUrlMitigation = True,
    '    .SecurityCertificateRevocationFailure = False,
    '    .ShdoclcDllResourceLoading = False,
    '    .SpellcheckAndAutoCorrectSupport = False,
    '    .SslSecurityAlertDisplay = False,
    '    .StatusBarUpdateFrequency = True,
    '    .StructuredStorageDetection = False,
    '    .TabbedBrowsingShortcutsAndNotifications = True,
    '    .TelnetProtocolSupport = True,
    '    .UncFileSupportForMotW = True,
    '    .UsernamesAndPasswordsInUrls = True,
    '    .ValidateUrlNavigation = True,
    '    .VisualStudioLegacyHelpSupport = False,
    '    .WebBrowserControlDocumentZoom = True,
    '    .WebBrowserControlPopupManagement = True,
    '    .WebBrowserControlWindowControl = False,
    '    .WebControlVisuals = False,
    '    .WebOcAddonManagement = False,
    '    .WebSocket = True,
    '    .WebSocketAuthenticationPrompt = False,
    '    .WebSocketCloseTimeout = 15000,
    '    .WebSocketMaximumServerConnections = 6,
    '    .WebSocketFollowRedirects = False,
    '    .WindowRestrictions = True,
    '    .XDomainRequestObjectSupport = True,
    '    .XmlHttpRequestObjectSupport = True,
    '    .ZoneElevation = True,
    '    .ZoneHandlingForMissingFiles = False,
    '    .ZoneStringLoading = False
    '}
    'End Class


End Class