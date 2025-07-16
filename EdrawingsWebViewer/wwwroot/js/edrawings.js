// eDrawings ActiveX Control Integration JavaScript

// Test if ActiveX is supported in the current browser
window.testActiveXSupport = function() {
    try {
        // Try to create an ActiveX object (only works in IE)
        var test = new ActiveXObject("htmlfile");
        return true;
    } catch (e) {
        // ActiveX not supported
        return false;
    }
};

// Try to create and initialize eDrawings ActiveX control
window.initializeEDrawingsControl = function() {
    try {
        // eDrawings ActiveX Control CLSID
        var edrawingsClassId = "22945A69-1191-4DCF-9E6F-409BDE94D101";
        
        // Create the object element for eDrawings
        var container = document.getElementById('edrawings-container');
        if (!container) {
            console.error('eDrawings container not found');
            return false;
        }
        
        // Create the ActiveX object element
        var objectHtml = `
            <object id="eDrawingsControl" 
                    classid="clsid:${edrawingsClassId}"
                    codebase="http://www.solidworks.com/edrawings/install/EdrawingsInstaller.cab"
                    width="100%" 
                    height="100%"
                    style="border: none;">
                <param name="toolbar" value="true">
                <param name="menubar" value="true">
                <param name="statusbar" value="true">
                <param name="UserInterfaceEnabled" value="true">
                <param name="ToolbarVisible" value="true">
                <param name="src" value="">
                <p>eDrawings ActiveX control could not be loaded. Please ensure:</p>
                <ul>
                    <li>You are using Internet Explorer</li>
                    <li>eDrawings viewer is installed</li>
                    <li>ActiveX controls are enabled</li>
                </ul>
            </object>
        `;
        
        container.innerHTML = objectHtml;
        
        // Try to get reference to the control
        setTimeout(function() {
            var control = document.getElementById('eDrawingsControl');
            if (control && typeof control.OpenDoc === 'function') {
                console.log('eDrawings control loaded successfully');
                window.eDrawingsControl = control;
                return true;
            } else {
                console.log('eDrawings control not available');
                return false;
            }
        }, 1000);
        
    } catch (e) {
        console.error('Error initializing eDrawings control:', e);
        return false;
    }
};

// Load an eDrawings file
window.loadEDrawingsFile = function(filePath) {
    try {
        if (window.eDrawingsControl && typeof window.eDrawingsControl.OpenDoc === 'function') {
            window.eDrawingsControl.OpenDoc(filePath, false, false, false, "");
            return true;
        } else {
            console.error('eDrawings control not available');
            return false;
        }
    } catch (e) {
        console.error('Error loading eDrawings file:', e);
        return false;
    }
};

// Get eDrawings control information
window.getEDrawingsInfo = function() {
    try {
        if (window.eDrawingsControl) {
            return {
                version: window.eDrawingsControl.Version || 'Unknown',
                isLoaded: typeof window.eDrawingsControl.OpenDoc === 'function'
            };
        }
        return null;
    } catch (e) {
        console.error('Error getting eDrawings info:', e);
        return null;
    }
};

// Event handlers for eDrawings control
window.setupEDrawingsEvents = function() {
    try {
        if (window.eDrawingsControl) {
            // Set up event handlers
            window.eDrawingsControl.OnFinishedLoadingDocument = function(fileName) {
                console.log('Document loaded:', fileName);
                // Notify Blazor component
                if (window.DotNet) {
                    DotNet.invokeMethodAsync('EdrawingsWebViewer', 'OnDocumentLoaded', fileName);
                }
            };
            
            window.eDrawingsControl.OnFailedLoadingDocument = function(fileName, errorCode, errorString) {
                console.error('Failed to load document:', fileName, errorCode, errorString);
                // Notify Blazor component
                if (window.DotNet) {
                    DotNet.invokeMethodAsync('EdrawingsWebViewer', 'OnDocumentLoadFailed', fileName, errorString);
                }
            };
        }
    } catch (e) {
        console.error('Error setting up eDrawings events:', e);
    }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    console.log('eDrawings JavaScript loaded');
}); 