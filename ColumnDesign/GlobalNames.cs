using System.IO;
using System.Reflection;
using System.Reflection.Emit;

namespace ColumnDesign
{
    public static class GlobalNames
    {
        public static readonly string FilesLocationPrefix = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),@"Source\");
        public static readonly string FamiliesLocationPrefix =  Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),@"Families\");
        public static readonly string ConfigColumnLocation =  Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),@"Column Sheet Data.txt");
        public static readonly string ConfigScissorsLocation =  Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),@"Scissor Clamp Column Sheet Data.txt");
    }
}