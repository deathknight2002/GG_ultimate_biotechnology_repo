Attribute VB_Name = "PluginHelper"
Option Explicit

' ========================================================================
' SeleniumVBA Plugin Helper Module
' Provides VBA integration with .NET Plugin System
' ========================================================================

Private m_pluginManager As Object
Private m_pluginPath As String

' Initialize the plugin system
Public Function InitializePlugins(Optional customPath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Create plugin manager
    Set m_pluginManager = CreateObject("SeleniumVBA.Plugins.Core.PluginManager")
    
    ' Set plugin path
    If customPath = "" Then
        m_pluginPath = ThisWorkbook.Path & "\plugins\.NET"
    Else
        m_pluginPath = customPath
    End If
    
    ' Add plugin directory and load plugins
    m_pluginManager.AddPluginDirectory m_pluginPath
    m_pluginManager.LoadPlugins
    
    InitializePlugins = True
    Debug.Print "Plugins initialized from: " & m_pluginPath
    Exit Function
    
ErrorHandler:
    Debug.Print "Error initializing plugins: " & Err.Description
    InitializePlugins = False
End Function

' Execute a plugin by ID
Public Function ExecutePlugin(pluginId As String, ParamArray params() As Variant) As Object
    On Error GoTo ErrorHandler
    
    If m_pluginManager Is Nothing Then
        If Not InitializePlugins() Then
            Err.Raise vbObjectError + 1, "PluginHelper", "Failed to initialize plugin system"
        End If
    End If
    
    ' Create parameters dictionary
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    
    ' Add parameters (expect pairs: key, value, key, value, ...)
    Dim i As Long
    For i = LBound(params) To UBound(params) Step 2
        If i + 1 <= UBound(params) Then
            parameters.Add CStr(params(i)), params(i + 1)
        End If
    Next i
    
    ' Execute plugin
    Set ExecutePlugin = m_pluginManager.ExecutePlugin(pluginId, parameters)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error executing plugin '" & pluginId & "': " & Err.Description
    Set ExecutePlugin = Nothing
End Function

' Get list of all loaded plugins
Public Function GetPluginList() As Collection
    On Error GoTo ErrorHandler
    
    If m_pluginManager Is Nothing Then
        If Not InitializePlugins() Then
            Err.Raise vbObjectError + 1, "PluginHelper", "Failed to initialize plugin system"
        End If
    End If
    
    Dim plugins As Object
    Dim result As New Collection
    Dim plugin As Object
    
    Set plugins = m_pluginManager.GetAllPlugins()
    
    For Each plugin In plugins
        result.Add plugin.Name & " (" & plugin.Id & ")"
    Next plugin
    
    Set GetPluginList = result
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting plugin list: " & Err.Description
    Set GetPluginList = New Collection
End Function

' DNA Sequencing Helper
Public Function AnalyzeDNA(sequence As String) As Object
    On Error GoTo ErrorHandler
    
    Set AnalyzeDNA = ExecutePlugin("biotech.dna.sequencing", _
        "action", "analyze", _
        "sequence", sequence)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error analyzing DNA: " & Err.Description
    Set AnalyzeDNA = Nothing
End Function

' Protein Analysis Helper
Public Function AnalyzeProtein(sequence As String) As Object
    On Error GoTo ErrorHandler
    
    Set AnalyzeProtein = ExecutePlugin("biotech.protein.analysis", _
        "action", "analyze", _
        "sequence", sequence)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error analyzing protein: " & Err.Description
    Set AnalyzeProtein = Nothing
End Function

' Apply UI Theme Helper
Public Function ApplyTheme(themeName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim result As Object
    Set result = ExecutePlugin("ui.theme.glass-oled", _
        "action", "applytheme", _
        "theme", themeName)
    
    If Not result Is Nothing Then
        ApplyTheme = result("Success")
    Else
        ApplyTheme = False
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error applying theme: " & Err.Description
    ApplyTheme = False
End Function

' Render UI Component Helper
Public Function RenderComponent(componentName As String, ParamArray dataParams() As Variant) As String
    On Error GoTo ErrorHandler
    
    ' Create data dictionary
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = LBound(dataParams) To UBound(dataParams) Step 2
        If i + 1 <= UBound(dataParams) Then
            data.Add CStr(dataParams(i)), dataParams(i + 1)
        End If
    Next i
    
    ' Create parameters for plugin
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "action", "rendercomponent"
    params.Add "component", componentName
    params.Add "data", data
    
    ' Execute plugin
    Dim result As Object
    Set result = m_pluginManager.ExecutePlugin("ui.theme.glass-oled", params)
    
    If Not result Is Nothing Then
        RenderComponent = result("HTML")
    Else
        RenderComponent = ""
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error rendering component: " & Err.Description
    RenderComponent = ""
End Function

' Cleanup plugin system
Public Sub CleanupPlugins()
    On Error Resume Next
    
    If Not m_pluginManager Is Nothing Then
        m_pluginManager.UnloadAllPlugins
        Set m_pluginManager = Nothing
    End If
End Sub

' Example: Process Excel data with biotech plugin
Public Sub ProcessExcelDataWithPlugin()
    On Error GoTo ErrorHandler
    
    ' Initialize plugins
    If Not InitializePlugins() Then
        MsgBox "Failed to initialize plugin system", vbCritical
        Exit Sub
    End If
    
    ' Get data from Excel (assuming sequence in column A)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Start from row 2 (assuming headers in row 1)
        Dim sequence As String
        sequence = ws.Cells(i, 1).Value
        
        If sequence <> "" Then
            ' Analyze DNA sequence
            Dim result As Object
            Set result = AnalyzeDNA(sequence)
            
            If Not result Is Nothing Then
                ' Write results to Excel
                ws.Cells(i, 2).Value = result("Length")
                ws.Cells(i, 3).Value = result("GCContent")
                ws.Cells(i, 4).Value = result("ComplexityScore")
                ws.Cells(i, 5).Value = "Success"
            Else
                ws.Cells(i, 5).Value = "Error"
            End If
        End If
    Next i
    
    MsgBox "Processing complete!", vbInformation
    CleanupPlugins
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing data: " & Err.Description, vbCritical
    CleanupPlugins
End Sub

' Example: Launch biotech dashboard with WebDriver
Public Sub LaunchBiotechDashboard()
    On Error GoTo ErrorHandler
    
    Dim driver As WebDriver
    Set driver = New WebDriver
    
    ' Start browser
    driver.StartChrome
    driver.OpenBrowser
    
    ' Navigate to dashboard
    Dim dashboardPath As String
    dashboardPath = "file:///" & Replace(ThisWorkbook.Path, "\", "/") & "/plugins/.NET/UI/templates/dashboard.html"
    driver.NavigateTo dashboardPath
    
    ' Wait for page load
    driver.Wait 2000
    
    ' You can now interact with the dashboard via driver
    MsgBox "Dashboard launched! Close this message to continue...", vbInformation
    
    ' Keep browser open for user interaction
    ' driver.CloseBrowser
    ' driver.Shutdown
    Exit Sub
    
ErrorHandler:
    MsgBox "Error launching dashboard: " & Err.Description, vbCritical
    If Not driver Is Nothing Then
        driver.CloseBrowser
        driver.Shutdown
    End If
End Sub
