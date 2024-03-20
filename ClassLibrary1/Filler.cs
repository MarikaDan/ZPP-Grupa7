using Org.BouncyCastle.Bcpg.OpenPgp;

namespace ClassLibrary1
{
    public class Filler
    {
        public static string GetOutPath(string inputPath, string fileName)
        {
            var dir = Path.GetDirectoryName(inputPath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var outPath = Path.Combine(dir, fileName);

            return outPath;
        }

        public static string GetTmpFileName(string extension)
        {
            var tmp = Path.GetRandomFileName();
            return $"{tmp}.{extension}";
        }

    }
}