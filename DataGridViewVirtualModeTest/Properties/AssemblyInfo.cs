using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Ssepan.Application.WinForms;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("DataGridViewVirtualModeTest")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("DataGridViewVirtualModeTest")]
[assembly: AssemblyCopyright("Copyright ©  2009")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("466af7a6-ef63-4894-b1be-37ac433a4053")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
[assembly: AssemblyVersion("0.13.0.*")]


#region " Helper class to get information for the About form. "
/// <summary>
/// This class uses the System.Reflection.Assembly class to
/// access assembly meta-data
/// This class is ! a normal feature of AssemblyInfo.cs
/// </summary>
public class AssemblyInfo :
    AssemblyInfoBase
{
    // Used by Helper Functions to access information from Assembly Attributes
    public AssemblyInfo()
    {
        base.myType = typeof(DataGridViewVirtualModeTest.Form1);
    }
}
#endregion
