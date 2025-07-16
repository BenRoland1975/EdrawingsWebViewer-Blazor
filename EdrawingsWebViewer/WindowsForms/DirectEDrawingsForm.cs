using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Linq; // Added for LINQ methods
using System.Collections.Generic; // Added for List

namespace EdrawingsWebViewer.WindowsForms
{
    // Custom AxHost wrapper for eDrawings
    public class EDrawingsAxHost : AxHost
    {
        public EDrawingsAxHost() : base("22945A69-1191-4DCF-9E6F-409BDE94D101")
        {
        }

        protected override void OnCreateControl()
        {
            base.OnCreateControl();
        }

        // Expose the ActiveX control directly - we'll handle casting elsewhere
        public object? ActiveXControl => this.GetOcx();
    }

    public partial class DirectEDrawingsForm : Form
    {
        private EDrawingsAxHost? eDrawingsControl;
        private Button? btnLoadFile;
        private Button? btnOpenFileDialog;
        private TextBox? txtFilePath;
        private Label? lblStatus;
        private string selectedFilePath = string.Empty;
        private dynamic? eDrawingsObject; // Store the dynamic reference

        public DirectEDrawingsForm()
        {
            InitializeComponent();
            SetupForm();
        }

        private void InitializeComponent()
        {
            this.btnLoadFile = new Button();
            this.btnOpenFileDialog = new Button();
            this.txtFilePath = new TextBox();
            this.lblStatus = new Label();
            this.SuspendLayout();

            // Form
            this.ClientSize = new System.Drawing.Size(1024, 768);
            this.Text = "eDrawings Direct ActiveX Host - Windows Forms";
            this.WindowState = FormWindowState.Maximized;

            // Panel for controls - make it taller to accommodate all controls
            var panel = new Panel();
            panel.Height = 80; // Increased from 60 to 80 to accommodate two rows
            panel.Dock = DockStyle.Top;
            panel.BackColor = System.Drawing.Color.LightGray;

            // File path textbox
            this.txtFilePath.Location = new System.Drawing.Point(12, 12);
            this.txtFilePath.Size = new System.Drawing.Size(300, 23);
            this.txtFilePath.PlaceholderText = "Enter path to eDrawings file...";

            // Open file dialog button
            this.btnOpenFileDialog.Location = new System.Drawing.Point(318, 12);
            this.btnOpenFileDialog.Size = new System.Drawing.Size(75, 23);
            this.btnOpenFileDialog.Text = "Browse...";
            this.btnOpenFileDialog.UseVisualStyleBackColor = true;
            this.btnOpenFileDialog.Click += BtnOpenFileDialog_Click;

            // Reload file button
            this.btnLoadFile.Location = new System.Drawing.Point(399, 12);
            this.btnLoadFile.Size = new System.Drawing.Size(75, 23);
            this.btnLoadFile.Text = "Reload";
            this.btnLoadFile.UseVisualStyleBackColor = true;
            this.btnLoadFile.Click += BtnLoadFile_Click;

            // Zoom fit button (keep this as it's useful)
            var btnZoomFit = new Button();
            btnZoomFit.Location = new System.Drawing.Point(480, 12);
            btnZoomFit.Size = new System.Drawing.Size(75, 23);
            btnZoomFit.Text = "Zoom Fit";
            btnZoomFit.UseVisualStyleBackColor = true;
            btnZoomFit.Click += BtnZoomFit_Click;



            // Emergency Restore button 
            var btnRestore = new Button();
            btnRestore.Location = new System.Drawing.Point(561, 12);
            btnRestore.Size = new System.Drawing.Size(80, 23);
            btnRestore.Text = "🔄 RESTORE";
            btnRestore.BackColor = System.Drawing.Color.Orange;
            btnRestore.ForeColor = System.Drawing.Color.White;
            btnRestore.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            btnRestore.UseVisualStyleBackColor = false;
            btnRestore.Click += BtnRestore_Click;

            // Diagnostic button
            var btnDiagnostic = new Button();
            btnDiagnostic.Location = new System.Drawing.Point(647, 12);
            btnDiagnostic.Size = new System.Drawing.Size(80, 23);
            btnDiagnostic.Text = "Diagnostics";
            btnDiagnostic.UseVisualStyleBackColor = true;
            btnDiagnostic.Click += BtnDiagnostic_Click;

            // Status label - move to second row and make VERY visible
            this.lblStatus.Location = new System.Drawing.Point(12, 45);
            this.lblStatus.Size = new System.Drawing.Size(1000, 25);
            this.lblStatus.Text = "✅ eDrawings Integration Ready - Browse to load files automatically";
            this.lblStatus.ForeColor = System.Drawing.Color.White;
            this.lblStatus.BackColor = System.Drawing.Color.Green;
            this.lblStatus.BorderStyle = BorderStyle.Fixed3D;
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblStatus.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Bold);
            this.lblStatus.AutoSize = false;
            this.lblStatus.Visible = true;

            // Add controls to panel
            panel.Controls.Add(this.txtFilePath);
            panel.Controls.Add(this.btnOpenFileDialog);
            panel.Controls.Add(this.btnLoadFile);
            panel.Controls.Add(btnZoomFit);
            panel.Controls.Add(btnRestore);
            panel.Controls.Add(btnDiagnostic);
            panel.Controls.Add(this.lblStatus);

            // Add panel to form
            this.Controls.Add(panel);
            
            // Ensure the panel is brought to front
            panel.BringToFront();

            this.ResumeLayout(false);
        }

        private void SetupForm()
        {
            try
            {
                // Create the eDrawings ActiveX control with proper sizing
                UpdateStatus("Creating eDrawings ActiveX control with proper layout...");
                
                // Create a container panel for the eDrawings control that DOESN'T overlap
                var eDrawingsPanel = new Panel();
                
                // CRITICAL: Position BELOW the control panel, not overlapping!
                eDrawingsPanel.Location = new System.Drawing.Point(0, 80); // Start below 80px control panel
                eDrawingsPanel.Size = new System.Drawing.Size(this.ClientSize.Width, this.ClientSize.Height - 80);
                eDrawingsPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
                
                eDrawingsPanel.Padding = new Padding(5); // Add padding to prevent edge cutoff
                eDrawingsPanel.BackColor = System.Drawing.Color.DarkGray;
                
                eDrawingsControl = new EDrawingsAxHost();
                eDrawingsControl.Dock = DockStyle.Fill;
                
                // Add the eDrawings control to its container panel
                eDrawingsPanel.Controls.Add(eDrawingsControl);
                
                // Add the container panel to the form
                this.Controls.Add(eDrawingsPanel);
                
                // Ensure the control panel stays on top
                eDrawingsPanel.SendToBack();
                
                // Set up events
                eDrawingsControl.HandleCreated += EDrawingsControl_HandleCreated;
                
                UpdateStatus("eDrawings control created with proper sizing!");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error creating eDrawings control: {ex.Message}");
                MessageBox.Show($"Failed to create eDrawings control: {ex.Message}\n\nThis usually means:\n1. eDrawings Viewer is not installed\n2. ActiveX control is not registered\n3. Application needs to run as administrator", 
                    "eDrawings Control Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EDrawingsControl_HandleCreated(object? sender, EventArgs e)
        {
            // Add a small delay to ensure the control is fully loaded
            System.Threading.Timer? timer = null;
            timer = new System.Threading.Timer(_ => {
                try
                {
                    this.Invoke(new Action(() => {
                        try
                        {
                            UpdateStatus("🔍 Starting eDrawings interface detection...");
                            
                            var activeXControl = eDrawingsControl?.ActiveXControl;
                            if (activeXControl != null)
                            {
                                UpdateStatus($"📋 Raw ActiveX Type: {activeXControl.GetType().FullName}");
                                
                                // Store the object directly - the key is using it as dynamic
                                eDrawingsObject = activeXControl;
                                
                                // Test if we can access eDrawings methods using reflection first
                                var type = activeXControl.GetType();
                                var methods = type.GetMethods();
                                var hasEDrawingsMethods = false;
                                
                                // Look for eDrawings-specific methods
                                foreach (var method in methods)
                                {
                                    if (method.Name.Contains("OpenDoc") || method.Name.Contains("CloseActiveDoc") || 
                                        method.Name.Contains("SetMeasureMode") || method.Name.Contains("Version"))
                                    {
                                        hasEDrawingsMethods = true;
                                        break;
                                    }
                                }
                                
                                if (hasEDrawingsMethods)
                                {
                                    UpdateStatus("✅ eDrawings methods detected via reflection!");
                                }
                                else
                                {
                                    UpdateStatus("⚠️ No eDrawings methods found via reflection");
                                }
                                
                                // Test version access using late binding (this often works even with __ComObject)
                                try
                                {
                                    var version = type.InvokeMember("Version", 
                                        System.Reflection.BindingFlags.GetProperty, 
                                        null, activeXControl, null);
                                    UpdateStatus($"✅ Version via InvokeMember: {version}");
                                }
                                catch (Exception vex)
                                {
                                    UpdateStatus($"❌ Version via InvokeMember failed: {vex.Message}");
                                    
                                    // Try dynamic version as last resort
                                    try
                                    {
                                        dynamic dynControl = activeXControl;
                                        var version2 = dynControl.Version;
                                        UpdateStatus($"✅ Version via dynamic: {version2}");
                                    }
                                    catch (Exception vex2)
                                    {
                                        UpdateStatus($"❌ Version via dynamic failed: {vex2.Message}");
                                    }
                                }
                                
                                // Try to initialize UI
                                UpdateStatus("🎨 Attempting UI initialization...");
                                InitializeEDrawingsUI();
                            }
                            else
                            {
                                UpdateStatus("❌ ActiveX control is null!");
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"❌ HandleCreated error: {ex.Message}");
                        }
                    }));
                }
                catch (Exception ex)
                {
                    // Can't update status here since we're not on UI thread
                    Console.WriteLine($"Timer error: {ex.Message}");
                }
                finally
                {
                    timer?.Dispose();
                }
            }, null, 1000, Timeout.Infinite); // 1 second delay
        }

        private void InitializeEDrawingsUI()
        {
            if (eDrawingsObject == null) return;

            var successfulMethods = new List<string>();

            // Try to enable UI and toolbar immediately when control is created
            var attempts = new[]
            {
                new { Name = "UserInterfaceEnabled", Value = true },
                new { Name = "ShowUI", Value = true },
                new { Name = "ToolbarVisible", Value = true }
            };

            foreach (var attempt in attempts)
            {
                try
                {
                    // Use dynamic property assignment
                    switch (attempt.Name)
                    {
                        case "UserInterfaceEnabled":
                            eDrawingsObject.UserInterfaceEnabled = attempt.Value;
                            break;
                        case "ShowUI":
                            eDrawingsObject.ShowUI = attempt.Value;
                            break;
                        case "ToolbarVisible":
                            eDrawingsObject.ToolbarVisible = attempt.Value;
                            break;
                    }
                    successfulMethods.Add(attempt.Name);
                }
                catch
                {
                    // Continue to next method
                }
            }

            // Try method calls
            var methodAttempts = new[]
            {
                "EnableUI",
                "SetUserInterfaceEnabled",
                "SetToolbarVisible",
                "ShowToolbar",
                "EnableToolbar"
            };

            foreach (var methodName in methodAttempts)
            {
                try
                {
                    switch (methodName)
                    {
                        case "EnableUI":
                            eDrawingsObject.EnableUI(true);
                            break;
                        case "SetUserInterfaceEnabled":
                            eDrawingsObject.SetUserInterfaceEnabled(true);
                            break;
                        case "SetToolbarVisible":
                            eDrawingsObject.SetToolbarVisible(true);
                            break;
                        case "ShowToolbar":
                            eDrawingsObject.ShowToolbar(true);
                            break;
                        case "EnableToolbar":
                            eDrawingsObject.EnableToolbar(true);
                            break;
                    }
                    successfulMethods.Add(methodName);
                }
                catch
                {
                    // Continue to next method
                }
            }

            if (successfulMethods.Count > 0)
            {
                UpdateStatus($"✅ UI initialized using: {string.Join(", ", successfulMethods)}");
            }
            else
            {
                UpdateStatus("⚠️ UI methods not accessible - using default eDrawings settings");
            }
        }

        private void BtnOpenFileDialog_Click(object? sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "eDrawings Files (*.edrw;*.eprt;*.easm;*.edwg)|*.edrw;*.eprt;*.easm;*.edwg|SOLIDWORKS Files (*.sldprt;*.sldasm;*.slddrw)|*.sldprt;*.sldasm;*.slddrw|All Files (*.*)|*.*";
                openFileDialog.Title = "Select eDrawings or SOLIDWORKS File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    txtFilePath!.Text = selectedFilePath;
                    UpdateStatus($"Selected: {Path.GetFileName(selectedFilePath)}");
                    
                    // Auto-load the selected file
                    LoadFileFromPath(selectedFilePath);
                }
            }
        }

        private void LoadFileFromPath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Please select or enter a file path first.", "No File Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"File not found: {filePath}", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                UpdateStatus($"Loading file: {Path.GetFileName(filePath)}");
                
                if (eDrawingsObject != null)
                {
                    // Call OpenDoc method using COM late binding (based on diagnostics)
                    var type = eDrawingsObject.GetType();
                    try
                    {
                        type.InvokeMember("OpenDoc", 
                            System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                            null, eDrawingsObject, new object[] { filePath, false, false, false, "" });
                        UpdateStatus($"✅ OpenDoc called successfully for: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception openEx)
                    {
                        UpdateStatus($"❌ OpenDoc failed: {openEx.Message}");
                        // Try simpler parameter combinations
                        try
                        {
                            type.InvokeMember("OpenDoc", 
                                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                null, eDrawingsObject, new object[] { filePath });
                            UpdateStatus($"✅ OpenDoc (simple) successful for: {Path.GetFileName(filePath)}");
                        }
                        catch (Exception simpleEx)
                        {
                            UpdateStatus($"❌ OpenDoc (simple) also failed: {simpleEx.Message}");
                            throw; // Re-throw to show error dialog
                        }
                    }
                    
                    UpdateStatus($"✅ File loaded successfully: {Path.GetFileName(filePath)}");
                    
                    // Automatically show toolbar after successful file load
                    UpdateStatus("🔄 Auto-showing toolbar...");
                    ShowToolbar();
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available");
                    MessageBox.Show("eDrawings control is not available. Please check if eDrawings Viewer is installed and try restarting the application.", 
                        "Control Not Available", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Error loading file: {ex.Message}");
                MessageBox.Show($"Error loading file: {ex.Message}\n\nFile: {filePath}\n\nThis could mean:\n1. File format not supported\n2. File is corrupted\n3. eDrawings control error", 
                    "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLoadFile_Click(object? sender, EventArgs e)
        {
            string filePath = !string.IsNullOrEmpty(txtFilePath?.Text) ? txtFilePath.Text : selectedFilePath;
            LoadFileFromPath(filePath);
        }

        private void BtnMeasure_Click(object? sender, EventArgs e)
        {
            try
            {
                UpdateStatus("📏 Measurement tool guidance...");
                
                // Don't send any keyboard shortcuts - they disrupt the eDrawings interface
                MessageBox.Show("🔍 HOW TO USE MEASUREMENT IN eDrawings:\n\n" +
                    "The measurement APIs are not available, but you can still measure:\n\n" +
                    "✅ TOOLBAR: Look for ruler/measure icons in the eDrawings toolbar\n" +
                    "✅ RIGHT-CLICK: Right-click on the 3D model and look for measurement options\n" +
                    "✅ DOUBLE-CLICK: Try double-clicking edges or faces to measure\n" +
                    "✅ MOUSE: Some eDrawings versions show measurements on hover\n\n" +
                    "💡 TIP: If toolbar is missing, try the 'Show Toolbar' button first!", 
                    "Measurement Instructions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                UpdateStatus("💡 Check eDrawings toolbar for measurement tools, or right-click the model");
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Error: {ex.Message}");
            }
        }



        private void BtnZoomFit_Click(object? sender, EventArgs e)
        {
            try
            {
                if (eDrawingsObject != null)
                {
                    UpdateStatus("🔍 Testing zoom methods...");
                    var type = eDrawingsObject.GetType();
                    bool success = false;
                    
                    // Try different zoom methods that might exist
                    var zoomMethods = new[]
                    {
                        new { Name = "ViewZoomToFit", Params = new object[0] },
                        new { Name = "ZoomToFit", Params = new object[0] },
                        new { Name = "FitToWindow", Params = new object[0] },
                        new { Name = "ViewZoomToFit", Params = new object[] { true } },
                        new { Name = "ZoomExtents", Params = new object[0] },
                        new { Name = "ViewFit", Params = new object[0] }
                    };
                    
                    foreach (var method in zoomMethods)
                    {
                        try
                        {
                            type.InvokeMember(method.Name, 
                                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                null, eDrawingsObject, method.Params);
                            UpdateStatus($"✅ SUCCESS: {method.Name} worked - view should be fitted");
                            success = true;
                            break;
                        }
                        catch (System.MissingMethodException)
                        {
                            // Method doesn't exist, try next
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"❌ {method.Name} failed: {ex.Message}");
                        }
                    }
                    
                    if (!success)
                    {
                        UpdateStatus("❌ No zoom methods worked - try mouse wheel or right-click for zoom options");
                    }
                    
                    // Also try to refresh the control to ensure proper sizing
                    try
                    {
                        type.InvokeMember("Refresh", 
                            System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                            null, eDrawingsObject, new object[0]);
                        this.Refresh();
                        if (eDrawingsControl != null)
                            eDrawingsControl.Refresh();
                    }
                    catch { }
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Zoom operation error: {ex.Message}");
            }
        }

        private void SetView(string viewName)
        {
            try
            {
                if (eDrawingsObject != null)
                {
                    // Try standard view orientation methods
                    try
                    {
                        switch (viewName.ToLower())
                        {
                            case "top":
                                eDrawingsObject.ViewOrientationTop();
                                break;
                            case "front":
                                eDrawingsObject.ViewOrientationFront();
                                break;
                            case "isometric":
                                eDrawingsObject.ViewOrientationIsometric();
                                break;
                        }
                        UpdateStatus($"✅ View set to {viewName}");
                        return;
                    }
                    catch { /* Try alternative methods */ }

                    // Try alternative method names
                    try
                    {
                        switch (viewName.ToLower())
                        {
                            case "top":
                                eDrawingsObject.ViewTop();
                                break;
                            case "front":
                                eDrawingsObject.ViewFront();
                                break;
                            case "isometric":
                                eDrawingsObject.ViewIsometric();
                                break;
                        }
                        UpdateStatus($"✅ View set to {viewName} using View{viewName}");
                        return;
                    }
                    catch { /* Try more alternatives */ }

                    try
                    {
                        switch (viewName.ToLower())
                        {
                            case "top":
                                eDrawingsObject.SetViewTop();
                                break;
                            case "front":
                                eDrawingsObject.SetViewFront();
                                break;
                            case "isometric":
                                eDrawingsObject.SetViewIsometric();
                                break;
                        }
                        UpdateStatus($"✅ View set to {viewName} using SetView{viewName}");
                        return;
                    }
                    catch { /* No more alternatives */ }

                    UpdateStatus($"⚠️ Could not set {viewName} view - try using eDrawings toolbar");
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ View change error: {ex.Message}");
            }
        }

        private void BtnShowToolbar_Click(object? sender, EventArgs e)
        {
            ShowToolbar();
        }

        private void ShowToolbar()
        {
            try
            {
                if (eDrawingsObject != null)
                {
                    UpdateStatus("🔍 Testing toolbar methods that actually exist...");
                    
                    var type = eDrawingsObject.GetType();
                    bool success = false;
                    string successMethod = "";

                    // From diagnostics, we know ShowToolbar exists but needs the right parameters
                    // Let's try COMPREHENSIVE parameter combinations for ShowToolbar
                    var toolbarAttempts = new[]
                    {
                        new { Name = "ShowToolbar()", Params = new object[0] },
                        new { Name = "ShowToolbar(true)", Params = new object[] { true } },
                        new { Name = "ShowToolbar(false)", Params = new object[] { false } },
                        new { Name = "ShowToolbar(1)", Params = new object[] { 1 } },
                        new { Name = "ShowToolbar(0)", Params = new object[] { 0 } },
                        new { Name = "ShowToolbar(-1)", Params = new object[] { -1 } },
                        new { Name = "ShowToolbar(2)", Params = new object[] { 2 } },
                        new { Name = "ShowToolbar(\"true\")", Params = new object[] { "true" } },
                        new { Name = "ShowToolbar(\"show\")", Params = new object[] { "show" } },
                        new { Name = "ShowToolbar(\"toolbar\")", Params = new object[] { "toolbar" } },
                        new { Name = "ShowToolbar(1, true)", Params = new object[] { 1, true } },
                        new { Name = "ShowToolbar(0, true)", Params = new object[] { 0, true } },
                        new { Name = "ShowToolbar(true, true)", Params = new object[] { true, true } }
                    };

                    foreach (var attempt in toolbarAttempts)
                    {
                        try
                        {
                            type.InvokeMember("ShowToolbar", 
                                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                null, eDrawingsObject, attempt.Params);
                            success = true;
                            successMethod = attempt.Name;
                            UpdateStatus($"✅ SUCCESS: {attempt.Name} worked!");
                            break;
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"❌ {attempt.Name} failed: {ex.Message}");
                        }
                    }

                    // If ShowToolbar didn't work, the toolbar might not be controllable in this version
                    if (!success)
                    {
                        UpdateStatus("⚠️ ShowToolbar method exists but all parameter combinations failed");
                        UpdateStatus("💡 This suggests the toolbar is not controllable via API in this eDrawings version");
                        
                        // Try a refresh to see if anything changed
                        try
                        {
                            type.InvokeMember("Refresh", 
                                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                null, eDrawingsObject, new object[0]);
                            UpdateStatus("🔄 Refreshed eDrawings control");
                        }
                        catch { }
                    }

                    // Force window refresh to make any changes visible
                    this.Refresh();
                    if (eDrawingsControl != null)
                        eDrawingsControl.Refresh();
                    
                    if (success)
                    {
                        UpdateStatus($"✅ SUCCESS: {successMethod} executed - check if toolbar appeared");
                    }
                    else
                    {
                        UpdateStatus("❌ CONCLUSION: ShowToolbar method exists but requires parameters we don't know");
                        UpdateStatus("💡 The toolbar might be controlled by eDrawings settings, not API");
                        UpdateStatus("🔍 Try: Right-click on eDrawings viewer → Toolbar options");
                    }
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Show toolbar error: {ex.Message}");
            }
        }

        private void BtnFixSize_Click(object? sender, EventArgs e)
        {
            try
            {
                UpdateStatus("🔧 Fixing eDrawings control size and layout...");
                
                if (eDrawingsControl != null)
                {
                    // Force control to refresh and resize
                    eDrawingsControl.SuspendLayout();
                    
                    // Get current form size
                    var formWidth = this.ClientSize.Width;
                    var formHeight = this.ClientSize.Height;
                    
                    UpdateStatus($"📐 Form size: {formWidth}x{formHeight}");
                    
                    // Calculate proper eDrawings control size (form minus control panel height and padding)
                    var controlPanelHeight = 80;
                    var padding = 10;
                    
                    var eDrawingsWidth = formWidth - (padding * 2);
                    var eDrawingsHeight = formHeight - controlPanelHeight - (padding * 2);
                    
                    // Fix the parent panel position first
                    var parentPanel = eDrawingsControl.Parent;
                    if (parentPanel != null)
                    {
                        parentPanel.Location = new System.Drawing.Point(0, controlPanelHeight);
                        parentPanel.Size = new System.Drawing.Size(formWidth, formHeight - controlPanelHeight);
                        UpdateStatus($"📐 Parent panel positioned at: Y={controlPanelHeight}, Size={formWidth}x{formHeight - controlPanelHeight}");
                    }
                    
                    // eDrawings control fills the parent panel (which is now correctly positioned)
                    eDrawingsControl.Dock = DockStyle.Fill;
                    
                    eDrawingsControl.ResumeLayout();
                    eDrawingsControl.Refresh();
                    
                    UpdateStatus($"✅ eDrawings control resized to: {eDrawingsWidth}x{eDrawingsHeight}");
                    
                    // Also try to fit the view if eDrawings object is available
                    if (eDrawingsObject != null)
                    {
                        try
                        {
                            var type = eDrawingsObject.GetType();
                            type.InvokeMember("Refresh", 
                                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                null, eDrawingsObject, new object[0]);
                            UpdateStatus("🔄 eDrawings viewer refreshed");
                        }
                        catch (Exception refreshEx)
                        {
                            UpdateStatus($"⚠️ eDrawings refresh failed: {refreshEx.Message}");
                        }
                    }
                    
                    // Force form refresh
                    this.Refresh();
                    UpdateStatus("✅ Layout fix complete - check if toolbar/edges are now visible");
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Fix size error: {ex.Message}");
            }
        }

        private void BtnRestore_Click(object? sender, EventArgs e)
        {
            try
            {
                UpdateStatus("🚨 EMERGENCY RESTORE: Attempting to restore eDrawings interface...");
                
                if (eDrawingsObject != null)
                {
                    var type = eDrawingsObject.GetType();
                    
                    // Get current file name
                    try
                    {
                        var currentFile = type.InvokeMember("FileName", 
                            System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.IgnoreCase,
                            null, eDrawingsObject, null);
                        
                        if (currentFile != null && !string.IsNullOrEmpty(currentFile.ToString()))
                        {
                            string fileName = currentFile.ToString() ?? "";
                            UpdateStatus($"📄 Found loaded file: {System.IO.Path.GetFileName(fileName)}");
                            UpdateStatus("🔄 Force-reloading file to restore toolbar and interface...");
                            
                            // First try to close any existing document
                            try
                            {
                                type.InvokeMember("CloseActiveDoc", 
                                    System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                    null, eDrawingsObject, new object[0]);
                                UpdateStatus("✅ Closed existing document");
                                System.Threading.Thread.Sleep(500); // Brief pause
                            }
                            catch (Exception closeEx)
                            {
                                UpdateStatus($"⚠️ Close failed (continuing): {closeEx.Message}");
                            }
                            
                            // Reload the same file with multiple parameter combinations
                            bool success = false;
                            var paramCombinations = new object[][]
                            {
                                new object[] { fileName, false, false, false, "" },
                                new object[] { fileName, true, false, false, "" },
                                new object[] { fileName },
                                new object[] { fileName, false },
                                new object[] { fileName, true }
                            };
                            
                            foreach (var paramSet in paramCombinations)
                            {
                                try
                                {
                                    type.InvokeMember("OpenDoc", 
                                        System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                        null, eDrawingsObject, paramSet);
                                    UpdateStatus($"✅ SUCCESS: File reloaded with {paramSet.Length} parameters");
                                    success = true;
                                    break;
                                }
                                catch (Exception openEx)
                                {
                                    UpdateStatus($"❌ Failed with {paramSet.Length} params: {openEx.Message}");
                                }
                            }
                            
                            if (success)
                            {
                                UpdateStatus("🎯 File successfully reloaded - eDrawings toolbar should be restored!");
                                
                                // Force multiple refreshes
                                try
                                {
                                    type.InvokeMember("Refresh", 
                                        System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                        null, eDrawingsObject, new object[0]);
                                    UpdateStatus("🔄 eDrawings object refreshed");
                                }
                                catch { }
                                
                                // Force control refreshes
                                if (eDrawingsControl != null)
                                {
                                    eDrawingsControl.Refresh();
                                    eDrawingsControl.Invalidate();
                                    UpdateStatus("🔄 ActiveX control refreshed");
                                }
                                
                                // Force form refresh
                                this.Refresh();
                                this.Invalidate();
                                UpdateStatus("🔄 Form refreshed");
                                
                                // Brief pause then final status
                                System.Threading.Thread.Sleep(1000);
                                UpdateStatus("🎉 EMERGENCY RESTORE COMPLETE! Check if toolbar is visible at the top of the viewer.");
                            }
                            else
                            {
                                UpdateStatus("❌ All reload attempts failed - toolbar may not be recoverable");
                                UpdateStatus("💡 Try manually resizing the window or restarting the application");
                            }
                        }
                        else
                        {
                            UpdateStatus("❌ No file loaded - cannot restore interface");
                            
                            // If no file loaded, try loading the last known good file
                            if (!string.IsNullOrEmpty(selectedFilePath) && System.IO.File.Exists(selectedFilePath))
                            {
                                UpdateStatus($"🔄 Attempting to load last known file: {System.IO.Path.GetFileName(selectedFilePath)}");
                                LoadFileFromPath(selectedFilePath);
                            }
                            else
                            {
                                MessageBox.Show("No file is currently loaded and no previous file found.\n\nTo restore the interface:\n1. Browse and load a file\n2. Then use the RESTORE button if toolbar disappears", 
                                    "No File Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    catch (Exception restoreEx)
                    {
                        UpdateStatus($"❌ Restore process failed: {restoreEx.Message}");
                        UpdateStatus("💡 SOLUTION: Close this window and launch a fresh eDrawings viewer");
                        
                        // Show helpful error dialog
                        MessageBox.Show($"Emergency restore failed: {restoreEx.Message}\n\nRecommended actions:\n1. Close this window\n2. Launch fresh eDrawings viewer from main page\n3. Load your file again\n\nThis will guarantee a clean interface with toolbar visible.", 
                            "Restore Failed - Manual Recovery Needed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available - major failure");
                    MessageBox.Show("eDrawings control is completely unavailable.\n\nYou must close this window and launch a new viewer from the main page.", 
                        "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ CRITICAL ERROR during restore: {ex.Message}");
                MessageBox.Show($"Critical error during emergency restore: {ex.Message}\n\nPlease close this window and launch a fresh viewer.", 
                    "Critical Restore Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDiagnostic_Click(object? sender, EventArgs e)
        {
            try
            {
                UpdateStatus("🔍 Running comprehensive eDrawings API diagnostics...");
                
                if (eDrawingsObject != null)
                {
                    var type = eDrawingsObject.GetType();
                    
                    var diagnostics = new System.Text.StringBuilder();
                    diagnostics.AppendLine("=== eDrawings COM Object Diagnostics (NEW VERSION!) ===\n");
                    
                    diagnostics.AppendLine($"Object Type: {type.Name}");
                    diagnostics.AppendLine($"Full Type Name: {type.FullName}");
                    diagnostics.AppendLine($"Assembly: {type.Assembly.FullName}\n");

                    // For __ComObject, we need to use COM late binding
                    if (type.Name == "__ComObject")
                    {
                        diagnostics.AppendLine("⚠️  DETECTED __ComObject - Testing eDrawings Methods via COM Late Binding\n");
                        
                        // Test common eDrawings methods using InvokeMember
                        var testMethods = new[]
                        {
                            "OpenDoc", "CloseActiveDoc", "CloseAllDocuments", "Print", "Save",
                            "SetMeasureMode", "ViewMeasure", "MeasureDistance", "GetMeasurement",
                            "EnableUI", "SetUserInterfaceEnabled", "ShowToolbar", "SetToolbarVisible",
                            "ZoomToFit", "SetViewOrientation", "Refresh", "GetVersion", "About"
                        };
                        
                        var testProperties = new[]
                        {
                            "Version", "UserInterfaceEnabled", "ToolbarVisible", "ShowUI", "MeasureMode",
                            "IsDocumentLoaded", "FileName", "DocumentType", "EnableMarkups", "EnableMeasurement"
                        };
                        
                        diagnostics.AppendLine("TESTING METHODS VIA COM LATE BINDING:");
                        foreach (var methodName in testMethods)
                        {
                            try
                            {
                                // Try to invoke the method with no parameters
                                type.InvokeMember(methodName, 
                                    System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.IgnoreCase,
                                    null, eDrawingsObject, new object[0]);
                                diagnostics.AppendLine($"  ✅ {methodName} - EXISTS (executed without parameters)");
                            }
                            catch (System.Reflection.TargetParameterCountException)
                            {
                                // Method exists but needs parameters
                                diagnostics.AppendLine($"  ✅ {methodName} - EXISTS (requires parameters)");
                            }
                            catch (System.MissingMethodException)
                            {
                                diagnostics.AppendLine($"  ❌ {methodName} - NOT FOUND");
                            }
                            catch (Exception ex)
                            {
                                diagnostics.AppendLine($"  ⚠️  {methodName} - EXISTS but error: {ex.Message}");
                            }
                        }
                        
                        diagnostics.AppendLine("\nTESTING PROPERTIES VIA COM LATE BINDING:");
                        foreach (var propName in testProperties)
                        {
                            try
                            {
                                var value = type.InvokeMember(propName, 
                                    System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.IgnoreCase,
                                    null, eDrawingsObject, null);
                                diagnostics.AppendLine($"  ✅ {propName}: {value}");
                            }
                            catch (System.MissingMethodException)
                            {
                                diagnostics.AppendLine($"  ❌ {propName} - PROPERTY NOT FOUND");
                            }
                            catch (Exception ex)
                            {
                                diagnostics.AppendLine($"  ⚠️  {propName} - Error: {ex.Message}");
                            }
                        }

                        // Now test setting properties
                        diagnostics.AppendLine("\nTESTING SETTING PROPERTIES:");
                        var settableProps = new[]
                        {
                            new { Name = "UserInterfaceEnabled", Value = true },
                            new { Name = "ToolbarVisible", Value = true },
                            new { Name = "ShowUI", Value = true },
                            new { Name = "MeasureMode", Value = true }
                        };

                        foreach (var prop in settableProps)
                        {
                            try
                            {
                                type.InvokeMember(prop.Name, 
                                    System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.IgnoreCase,
                                    null, eDrawingsObject, new object[] { prop.Value });
                                diagnostics.AppendLine($"  ✅ SET {prop.Name} = {prop.Value} - SUCCESS");
                            }
                            catch (System.MissingMethodException)
                            {
                                diagnostics.AppendLine($"  ❌ SET {prop.Name} - PROPERTY NOT SETTABLE");
                            }
                            catch (Exception ex)
                            {
                                diagnostics.AppendLine($"  ⚠️  SET {prop.Name} - Error: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        // Normal reflection for non-COM objects
                        var allMethods = type.GetMethods();
                        var allProperties = type.GetProperties();

                        diagnostics.AppendLine("ALL AVAILABLE METHODS:");
                        foreach (var method in allMethods)
                            if (!method.IsSpecialName)
                                diagnostics.AppendLine($"  - {method.Name}");
                        
                        diagnostics.AppendLine("\nALL AVAILABLE PROPERTIES:");
                        foreach (var prop in allProperties)
                            diagnostics.AppendLine($"  - {prop.Name}");
                    }

                    // Show in a scrollable message box
                    var diagForm = new Form();
                    diagForm.Text = "eDrawings COM API Diagnostics";
                    diagForm.Size = new System.Drawing.Size(800, 600);
                    diagForm.StartPosition = FormStartPosition.CenterParent;
                    
                    var textBox = new TextBox();
                    textBox.Multiline = true;
                    textBox.ScrollBars = ScrollBars.Both;
                    textBox.Dock = DockStyle.Fill;
                    textBox.Text = diagnostics.ToString();
                    textBox.Font = new System.Drawing.Font("Consolas", 9);
                    textBox.ReadOnly = true;
                    
                    diagForm.Controls.Add(textBox);
                    diagForm.ShowDialog(this);

                    UpdateStatus($"✅ COM Late Binding Diagnostics Complete!");
                }
                else
                {
                    UpdateStatus("❌ eDrawings control not available for diagnostics");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Diagnostics error: {ex.Message}");
                MessageBox.Show($"Error running diagnostics: {ex.Message}", "Diagnostics Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper method to update status consistently
        private void UpdateStatus(string message)
        {
            if (lblStatus != null)
            {
                // Update on UI thread if needed
                if (lblStatus.InvokeRequired)
                {
                    lblStatus.Invoke(new Action(() => {
                        lblStatus.Text = message;
                        lblStatus.Refresh();
                        lblStatus.Update();
                        this.Refresh(); // Also refresh the form
                    }));
                }
                else
                {
                    lblStatus.Text = message;
                    lblStatus.Refresh();
                    lblStatus.Update();
                    this.Refresh(); // Also refresh the form
                }
                
                // Force immediate painting
                Application.DoEvents();
            }
        }
    }
} 