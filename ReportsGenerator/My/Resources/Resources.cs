using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ReportsGenerator.My.Resources;

[DebuggerNonUserCode]
[HideModuleName]
[GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
[CompilerGenerated]
[StandardModule]
internal sealed class Resources
{
    private static ResourceManager _resourceMan;

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    internal static ResourceManager ResourceManager
    {
        get
        {
            if (ReferenceEquals(_resourceMan, null))
            {
                var resourceManager = new ResourceManager("ReportsGenerator.Resources", typeof(Resources).Assembly);
                _resourceMan = resourceManager;
            }

            return _resourceMan;
        }
    }

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    internal static CultureInfo Culture { get; set; }

    internal static byte[] Help
    {
        get
        {
            var objectValue = RuntimeHelpers.GetObjectValue(ResourceManager.GetObject("help", Culture));
            return (byte[]) objectValue;
        }
    }

    internal static string Sortament => ResourceManager.GetString("Sortament", Culture);
}