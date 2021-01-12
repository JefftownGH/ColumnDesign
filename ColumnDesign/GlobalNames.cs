using System.IO;
using System.Reflection;
using System.Reflection.Emit;

namespace ColumnDesign
{
    public static class GlobalNames
    {
        public static readonly string FilesLocationPrefix = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),@"Source\");
        public static readonly string FamiliesLocationPrefix =  Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),@"Families\");
    }
}