using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Ssepan.Application;
using Ssepan.Application.WinForms;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("MultiDimensionalSpreadsheet (Alpha)")]
[assembly: AssemblyDescription("A research project for a multi-dimensional spreadsheet-like application. \n\nSteve Sepan takes credit for:\n~ designing and coding this application in C#.Net")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Free Software Foundation, Inc.")]
[assembly: AssemblyProduct("MultiDimensionalSpreadsheet")]
[assembly: AssemblyCopyright("Copyright (C) 1989, 1991 Free Software Foundation, Inc.  \n59 Temple Place - Suite 330, Boston, MA  02111-1307, USA")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("6b8206b1-1abe-4cd6-bf22-1d6e3295a4e9")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
[assembly: AssemblyVersion("0.14")]


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
        base.myType = typeof(MultiDimensionalSpreadsheet.ModelView);
    }
}
#endregion
