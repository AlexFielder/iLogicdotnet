Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Autodesk.iLogic.Interfaces
Imports System.Windows.Forms
Imports System.Linq

Namespace iLogicdotnet
    <ProgIdAttribute("iLogicdotnet.StandardAddInServer"),
    GuidAttribute("71a7b2f1-19fe-43aa-ba19-dd3476802e3b")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents m_uiEvents As UserInterfaceEvents
        'Private WithEvents m_sampleButton As ButtonDefinition
        Private DumpiLogicRulesButton As ButtonDefinition
        Private ExtractiLogicRulesButton As ButtonDefinition
        Private SimplifyRulesButton As ButtonDefinition

        Public rulesListObject As iLogicRuleType = Nothing
        Public ruleslistFileName As String = String.Empty

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            DumpiLogicRulesButton = controlDefs.AddButtonDefinition(My.Settings.ButtonNameDumpiLogicRules,
                                                              My.Settings.ButtonInternalNameDumpiLogicRules,
                                                              CommandTypesEnum.kFileOperationsCmdType,
                                                              Guid.NewGuid().ToString(),
                                                              My.Settings.ButtonDescrDumpiLogicRules,
                                                              My.Settings.ButtonTooltipDumpiLogicRules,
                                                              GetICOResource(My.Settings.ButtonIconDumpiLogic),
                                                              GetICOResource(My.Settings.ButtonIconDumpiLogic))
            ExtractiLogicRulesButton = controlDefs.AddButtonDefinition(My.Settings.ButtonNameExternaliseiLogicRules,
                                                              My.Settings.ButtonInternalNameExternaliseiLogicRules,
                                                              CommandTypesEnum.kFileOperationsCmdType,
                                                              Guid.NewGuid().ToString(),
                                                              My.Settings.ButtonDescrExternaliseiLogicRules,
                                                              My.Settings.ButtonTooltipExternaliseiLogicRules,
                                                              GetICOResource(My.Settings.ButtonIconExtract),
                                                              GetICOResource(My.Settings.ButtonIconExtract))
            SimplifyRulesButton = controlDefs.AddButtonDefinition(My.Settings.ButtonNameSimplifyiLogicRules,
                                                              My.Settings.ButtonInternalNameSimplifyiLogicRules,
                                                              CommandTypesEnum.kFileOperationsCmdType,
                                                              Guid.NewGuid().ToString(),
                                                              My.Settings.ButtonDescrSimplifyiLogicRules,
                                                              My.Settings.ButtonTooltipSimplifyiLogicRules,
                                                              GetICOResource(My.Settings.ButtonIconDumpiLogic),
                                                              GetICOResource(My.Settings.ButtonIconDumpiLogic))
            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region
#Region "AddIn-specific"
        ''' <summary>
        ''' Returns the relevant resource from the compiled .dll file
        ''' Means you don't need to Copy local .ico (or any other resource files)!
        ''' </summary>
        ''' <param name="icoResourceName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetICOResource(
                  ByVal icoResourceName As String) As Object
            Dim assemblyNet As System.Reflection.Assembly =
              System.Reflection.Assembly.GetExecutingAssembly()
            Dim stream As System.IO.Stream =
              assemblyNet.GetManifestResourceStream(icoResourceName)
            Dim ico As System.Drawing.Icon =
              New System.Drawing.Icon(stream)
            Return PictureDispConverter.ToIPictureDisp(ico)
        End Function
#End Region
#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")
            Dim tabPart As Inventor.RibbonTab = partRibbon.RibbonTabs.Add(My.Settings.TabName,
                                                                             My.Settings.PartRibbonTabNameInternal,
                                                                             Guid.NewGuid().ToString())
            Dim panelPart As Inventor.RibbonPanel = tabPart.RibbonPanels.Add(My.Settings.panelPartName,
                                                                                   My.Settings.panelPartNameInternal,
                                                                                   Guid.NewGuid().ToString())
            panelPart.CommandControls.AddButton(DumpiLogicRulesButton, True)
            panelPart.CommandControls.AddButton(ExtractiLogicRulesButton, True)
            panelPart.CommandControls.AddButton(SimplifyRulesButton, True)
            'get the assembly ribbon
            Dim assemblyRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Assembly")
            Dim tabAssembly As Inventor.RibbonTab = assemblyRibbon.RibbonTabs.Add(My.Settings.TabName,
                                                                                     My.Settings.AssemblyRibbonTabNameInternal,
                                                                                     Guid.NewGuid().ToString())
            Dim panelAssembly As Inventor.RibbonPanel = tabAssembly.RibbonPanels.Add(My.Settings.panelAssemblyName,
                                                                                           My.Settings.panelAssemblyNameInternal,
                                                                                           Guid.NewGuid().ToString())
            panelAssembly.CommandControls.AddButton(DumpiLogicRulesButton, True)
            panelAssembly.CommandControls.AddButton(ExtractiLogicRulesButton, True)
            panelAssembly.CommandControls.AddButton(SimplifyRulesButton, True)
            'get the zero doc ribbon
            Dim zeroDocRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("ZeroDoc")
            Dim tabZeroDoc As Inventor.RibbonTab = zeroDocRibbon.RibbonTabs.Add(My.Settings.TabName,
                                                                                   My.Settings.ZeroDocRibbonTabNameInternal,
                                                                                   Guid.NewGuid().ToString())
            Dim panelZeroDoc As Inventor.RibbonPanel = tabZeroDoc.RibbonPanels.Add(My.Settings.panelZeroDocName,
                                                                                      My.Settings.panelZeroDocNameInternal,
                                                                                      Guid.NewGuid().ToString())
            'these two buttons will not function with zero documents open!
            panelZeroDoc.CommandControls.AddButton(DumpiLogicRulesButton, True)
            panelZeroDoc.CommandControls.AddButton(ExtractiLogicRulesButton, True)
            panelZeroDoc.CommandControls.AddButton(SimplifyRulesButton, True)
            'get the drawing doc ribbon
            Dim DrawingRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Drawing")
            Dim tabDrawing As Inventor.RibbonTab = DrawingRibbon.RibbonTabs.Add(My.Settings.TabName,
                                                                                   My.Settings.DrawingRibbonTabNameInternal,
                                                                                   Guid.NewGuid().ToString())
            Dim panelDrawing As Inventor.RibbonPanel = tabDrawing.RibbonPanels.Add(My.Settings.panelDrawingName,
                                                                                      My.Settings.panelDrawingNameInternal,
                                                                                      Guid.NewGuid().ToString())
            panelDrawing.CommandControls.AddButton(DumpiLogicRulesButton, True)
            panelDrawing.CommandControls.AddButton(ExtractiLogicRulesButton, True)
            panelDrawing.CommandControls.AddButton(SimplifyRulesButton, True)

            AddHandler ExtractiLogicRulesButton.OnExecute, AddressOf ExtractRules
            AddHandler DumpiLogicRulesButton.OnExecute, AddressOf RunRulesDump
            AddHandler SimplifyRulesButton.OnExecute, AddressOf SimplifyRules
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        'Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
        '    MsgBox("Button was clicked.")
        'End Sub
#End Region
#Region "RulesDump"
        ''' <summary>
        '''
        ''' </summary>
        ''' <param name="Context"></param>
        Private Sub RunRulesDump(Context As NameValueMap)
            MessageBox.Show("Hello World!") '<-debugging!
            'Dim Auto As Automation.iLogicAutomation
            'Auto = iLogic.Automation
            Dim iLogicAuto As Object = GetiLogicAutomation(g_inventorApplication)
            Dim oDoc As Document = g_inventorApplication.ActiveDocument
            'Dim ruleColl As List(Of iLogicRule) = New List(Of iLogicRule)
            ' Build collection of rules
            'ruleColl = New List(Of iLogicRule)
            rulesListObject = New iLogicRuleType
            BuildRuleColl(oDoc, iLogicAuto)
            ' Loop through all referenced docs in assembly
            For Each oSubDoc As Document In oDoc.AllReferencedDocuments
                ' Build collection of rules

                BuildRuleColl(oSubDoc, iLogicAuto)
            Next
            DumpExistingRules(oDoc)
            MessageBox.Show("Done dumping iLogic rules to: " + ruleslistFileName)
        End Sub

        ''' <summary>
        ''' Dumps any iLogic rules found to an XML document of the same name as the assembly.
        ''' </summary>
        ''' <param name="oDoc"></param>
        Private Sub DumpExistingRules(ByVal oDoc As Document)
            'Dim rulesList As New iLogicRuleType
            ruleslistFileName = IO.Path.GetDirectoryName(oDoc.FullFileName) + "\" +
                                  IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName) +
                                  My.Settings.RulesFileNameDelimiter +
                                  My.Settings.DefaultRulesFileName

            rulesListObject.SaveToFile(ruleslistFileName)
        End Sub

        ''' <summary>
        ''' Adds iLogic rules from each document passed to it to the global rulelist.
        ''' </summary>
        ''' <param name="oDoc"></param>
        ''' <param name="iLogicAuto"></param>
        Private Sub BuildRuleColl(oDoc As Document, iLogicAuto As Object)
            Dim ruleName As String = String.Empty
            Dim rules As Object = iLogicAuto.rules(oDoc)

            If Not (rules Is Nothing) Then
                For Each rule As iLogicRule In rules
                    Dim thisrule As RuleType = New RuleType
                    thisrule.Name = rule.Name
                    thisrule.IsActive = rule.IsActive
                    thisrule.FireDependentImmediately = rule.FireDependentImmediately
                    thisrule.AutomaticOnParameterChange = rule.AutomaticOnParamChange
                    thisrule.Silentoperation = rule.SilentOperation
                    thisrule.Text = rule.Text
                    thisrule.ParentFileName = oDoc.FullFileName
                    'this is removing files where the file already exists in the .xml output.
                    'what we should do instead is look for rules with the same name and then remove the duplicates
                    Dim foundexistingrule As RuleType = (From anEntry As RuleType In rulesListObject.Rule
                                                         Where anEntry.ParentFileName = oDoc.FullFileName And anEntry.Name = thisrule.Name
                                                         Select anEntry).FirstOrDefault()
                    If Not foundexistingrule Is Nothing Then
                        'remove existing entries for this document
                        rulesListObject.Rule.RemoveAll(Function(X) X.ParentFileName.Contains(oDoc.FullFileName))
                        rulesListObject.Rule.Add(thisrule)
                    Else
                        'remove existing entries for this document
                        rulesListObject.Rule.Add(thisrule)
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Gets the iLogic Automation interface from the current Inventor application.
        ''' </summary>
        ''' <param name="app"></param>
        ''' <returns></returns>
        Function GetiLogicAutomation(ByVal app As Inventor.Application) As Object

            Dim addIn As Inventor.ApplicationAddIn = Nothing
            Try
                addIn = app.ApplicationAddIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
            Catch ex As Exception
                Return Nothing
            End Try

            Return addIn.Automation
        End Function
#End Region
#Region "ExtractRules"
        Private Sub ExtractRules(Context As NameValueMap)
            MessageBox.Show("Hello World", "Rule Extractor")
        End Sub
#End Region
#Region "Simplify Rules"
        Private Sub SimplifyRules(Context As NameValueMap)
            MessageBox.Show("Hello World", "Rule Simplification")
        End Sub
#End Region
    End Class
End Namespace

#Region "Icon Tools"
Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(iLogicdotnet.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object,
            ByRef iid As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)>
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)>
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
#End Region