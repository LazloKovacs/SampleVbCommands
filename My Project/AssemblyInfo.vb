Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Rhino.PlugIns

' Plug-in Description Attributes - all of these are optional
' These will show in Rhino's option dialog, in the tab Plug-ins
<Assembly: PlugInDescription(DescriptionType.Address, "-")>
<Assembly: PlugInDescription(DescriptionType.Country, "-")>
<Assembly: PlugInDescription(DescriptionType.Email, "-")>
<Assembly: PlugInDescription(DescriptionType.Phone, "-")>
<Assembly: PlugInDescription(DescriptionType.Fax, "-")>
<Assembly: PlugInDescription(DescriptionType.Organization, "-")>
<Assembly: PlugInDescription(DescriptionType.UpdateUrl, "-")>
<Assembly: PlugInDescription(DescriptionType.WebSite, "-")>


' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.
<Assembly: AssemblyTitle("SampleVbCommands")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("Robert McNeel & Associates")> 
<Assembly: AssemblyProduct("SampleVbCommands")> 
<Assembly: AssemblyCopyright("Copyright © Robert McNeel & Associates 2013")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("83a6e05b-4483-4268-82d1-f08f2f359a60")> ' This will also be the Guid of the Rhino plug-in

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 
