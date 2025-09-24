////////////////////////////////////////////////////////////////////////////////
//
// JavaScript library to load any .NET assembly/type and execute against its 
// public methods and properties
//
////////////////////////////////////////////////////////////////////////////////
function YahooGadgetHost()
{
	// ProgID of the Gadget Adapter for use in creating an ActiveX object instance
	var progID = "Yahoo.GadgetHost";
	// File name of the Gadget Interop DLL
	var assemblyName = "Yahoo.Gadget.Host.dll";
	// Gadget Adapter class GUID
	var guid = "{BDDE678A-70B4-4fc1-B70B-B4BCFD9DED40}";
	// Gadget Adapter assembly location (under the gadget's root directory)
	var assemblyStore = "\\bin\\" + System.Gadget.version + "\\";
	
	// Instance of the ActiveX gadget adapter object
	this._host;  
	
	// Method pointers
	this.Initialize = Initialize;
	this.InteropRegistered = InteropRegistered;
	this.GetActiveXObject = GetActiveXObject;
	
	////////////////////////////////////////////////////////////////////////////////
	//
	// Initializes this class by creating an instance of the Gadget Adapter.  If the
	// Gadget Adapter assembly is not yet retistered, the corresponding registry
	// values are added to register the Gadget Adapter.  (i.e. runtime registration)
	//
	////////////////////////////////////////////////////////////////////////////////
	function Initialize()
	{
		// Check if the gadget adapter needs to be registered
		if(InteropRegistered() == false)
		{
			// Register the adapter since an instance couldn't be created.
			RegisterGadgetInterop();
		}
		
		// Load an instance of the Gadget Adapter as an ActiveX object
		_host = GetActiveXObject();
	}
	
	function CleanUp()
	{
		UnregisterGadgetInterop();
	}
	
	////////////////////////////////////////////////////////////////////////////////
	//
	// Add the Gadget.Interop dll to the registry so it can be used by COM and
	// created in javascript as an ActiveX object
	//
	////////////////////////////////////////////////////////////////////////////////
	function RegisterGadgetInterop()
	{
		try
		{
			// Full path to the Gadget.Interop.dll assembly
			var fullPath = System.Gadget.path + assemblyStore;
			var asmPath = fullPath + assemblyName;
			
			// Register the interop assembly under the Current User registry key
			RegAsmInstall(progID, "Yahoo.Gadget.Host.GadgetHost", guid,
				"Yahoo.Gadget.Host, Version=1.0.0.0, Culture=neutral, PublicKeyToken=d1cd46393cf306ad",
				"1.0.0.0", asmPath);
		}
		catch(e)
		{}    
	}
	////////////////////////////////////////////////////////////////////////////////
	//
	// Remove the Gadget.Interop dll from the registry and the GAC
	//
	////////////////////////////////////////////////////////////////////////////////
	function UnregisterGadgetInterop()
	{
		try
		{
			RegAsmUninstall(progID, guid);
		}
		catch(e)
		{}
	}
	////////////////////////////////////////////////////////////////////////////////
	//
	// Returns true if a Gadget Adapter ActiveX object is successfully created,
	// otherwise returns false.
	//
	////////////////////////////////////////////////////////////////////////////////
	function InteropRegistered()
	{
		try
		{
			var proxy = GetActiveXObject();
			proxy = null;
			
			return true;
		}
		catch(e)
		{
			return false;
		}
	}
	////////////////////////////////////////////////////////////////////////////////
	//
	// Attempts to create and return an instance of the Gadget Adapter ActiveX object.
	//
	////////////////////////////////////////////////////////////////////////////////
	function GetActiveXObject()
	{
		return new ActiveXObject(progID);
	}
	////////////////////////////////////////////////////////////////////////////////
	//
	// Code to register a .NET type for COM interop
	//
	////////////////////////////////////////////////////////////////////////////////
	function RegAsmInstall(progId, cls, clsid, assembly, version, codebase) 
	{
		var wshShell;
		wshShell = new ActiveXObject("WScript.Shell"); 

		var root = "HKLM";
		
		try
		{
			wshShell.RegWrite(root + "\\Software\\Classes\\", progId);
		}
		catch(e)
		{
			root = "HKCU";
			wshShell.RegWrite(root + "\\Software\\Classes\\", progId);
		}
		
		wshShell.RegWrite(root + "\\Software\\Classes\\" + progId + "\\", cls);
		wshShell.RegWrite(root + "\\Software\\Classes\\" + progId + "\\CLSID\\", clsid);       
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\", cls); 

		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\", "mscoree.dll");
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\ThreadingModel", "Both");
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\Class", cls);
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\Assembly", assembly);
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\RuntimeVersion", "v2.0.50727");
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\CodeBase", codebase); 

		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\" + version + "\\Class", cls);
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\" + version + "\\Assembly", assembly);
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\" + version + "\\RuntimeVersion", "v2.0.50727");
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\InprocServer32\\" + version + "\\CodeBase", codebase); 

		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\ProgId\\", progId); 

		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\Implemented Categories\\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}\\", "");
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\Implemented Categories\\{7DD95801-9882-11CF-9FA9-00AA006C42C4}\\", "");
		wshShell.RegWrite(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\Implemented Categories\\{7DD95802-9882-11CF-9FA9-00AA006C42C4}\\", "");
	} 
	////////////////////////////////////////////////////////////////////////////////
	//
	// Unregister a component
	//
	////////////////////////////////////////////////////////////////////////////////
	function RegAsmUninstall(progId, clsid) 
	{
		var wshShell;
		wshShell = new ActiveXObject("WScript.Shell"); 

		var root = "HKLM";
		
		try
		{
			wshShell.RegDelete(root + "\\Software\\Classes\\", progId + "\\");
		}
		catch(e)
		{
			root = "HKCU";
			wshShell.RegDelete(root + "\\Software\\Classes\\", progId + "\\");
		}

		wshShell.RegDelete(root + "\\Software\\Classes\\CLSID\\" + clsid + "\\");
	} 
}