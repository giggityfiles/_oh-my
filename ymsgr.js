window.attachEvent("onload", pageLoad);
window.attachEvent("onunload", pageUnLoad);

var host = new YahooGadgetHost();
host.Initialize();

function pageLoad()
{
	imgBackground.addImageObject("logo.png", 4, 0);
		
	// Make sure this code runs after all event handlers are processed
	setTimeout(
		function() {
			try
			{
				LoadMessenger();
				imgBackground.removeObjects();
			}
			catch(e)
			{
				// TODO: Make sure this works ...
				imgBackground.removeObjects();
				imgBackground.addImageObject("logo.png", 6, 2);
				var loadingText = imgBackground.addTextObject(strError, "Segoe UI", 12, "White", 10, 110);
				loadingText.addShadow("Color(255, 0, 0, 0)", 10, 100, 0, 0);
				wpfHost.style.display = "none";
			}
		}, 1000);
}

function pageUnLoad()
{
	// this does not work
	//host.CleanUp();
	window.detachEvent("onload", pageLoad);
}

////////////////////////////////////////////////////////////////////////////////
//
// Set up the gadget and subscribe to events
//
////////////////////////////////////////////////////////////////////////////////
function LoadMessenger()
{
	//System.Gadget.Flyout.file = "Flyout.html";
	//System.Gadget.Flyout.onShow = OnShowFlyout;
	//System.Gadget.settingsUI = "settings.html";
	//System.Gadget.onSettingsClosed = SettingsClosed;
	System.Gadget.onUndock = CheckDock;
	System.Gadget.onDock = CheckDock;
	
	// Resize based on dock state
	CheckDock();

	// Don't set the classid until now to prevent the gadget from trying to load right away
	wpfHost.setAttribute("classid", "clsid:BDDE678A-70B4-4fc1-B70B-B4BCFD9DED40");
	wpfHost.GadgetObject = System.Gadget;
	wpfHost.GadgetSettings = System.Gadget.Settings;
	wpfHost.style.visibility = "visible";
	
	self.focus;
	document.body.focus();
}
////////////////////////////////////////////////////////////////////////////////
//
// Check if the gadget is docked or undocked.  Change images and CSS based on 
// the current docked state.
//
////////////////////////////////////////////////////////////////////////////////
function CheckDock()
{
	if(!System.Gadget.docked) 
	{
		GadgetUndocked();
	} 
	else if (System.Gadget.docked)
	{
		GadgetDocked(); 
	}
}
////////////////////////////////////////////////////////////////////////////////
//
// styles for gadget when undocked
//
////////////////////////////////////////////////////////////////////////////////
function GadgetUndocked()
{
}
////////////////////////////////////////////////////////////////////////////////
//
// styles for gadget when docked
// 
////////////////////////////////////////////////////////////////////////////////
function GadgetDocked()
{   
}
