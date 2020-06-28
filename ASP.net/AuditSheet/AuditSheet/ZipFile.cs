using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.IO;

namespace AuditSheet
{
    public class ZipFile
    {
        public static bool Compress(string dirSource, string dirDestination, string zipFilename)
        {
            try
            {
                if (!Directory.Exists(dirDestination))
                {
                    Directory.CreateDirectory(dirDestination);
                }

                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    string[] fileEntries = Directory.GetFiles(dirSource);
                    foreach (string fileName in fileEntries)
                    {
                        zip.AddFile(fileName, "");
                    }

                    string name = dirDestination + "\\" + zipFilename + ".zip";
                    if(File.Exists(name))
                    {
                        File.Delete(name);
                    }
                    zip.Save(name);
                }

                return true;
            }
            catch(Exception ex)
            {  }

            return false;
        }
    }
}