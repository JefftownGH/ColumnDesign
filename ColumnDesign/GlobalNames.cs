using System.IO;
using System.Reflection;
using System.Reflection.Emit;

namespace ColumnDesign
{
    public static class GlobalNames
    {
        public static readonly string WtFilesLocationPrefix = @$"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\";
        public static readonly string WtFamiliesLocationPrefix = @$"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Families\";
    }
}