using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using ICSharpCode.SharpZipLib.Zip;

namespace M3Utils
{
    public static class IOHelper
    {
        public static void ZipFile(string filePath, Stream stream)
        {
            int size;
            byte[] buffer = new byte[8000];

            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (ZipOutputStream zipOutStream = new ZipOutputStream(fileStream))
                {
                    zipOutStream.SetLevel(9);
                    zipOutStream.PutNextEntry(new ZipEntry(Path.GetFileName(filePath)));

                    if (stream.CanRead)
                    {
                        stream.Seek(0, SeekOrigin.Begin);

                        do
                        {
                            size = stream.Read(buffer, 0, buffer.Length);
                            zipOutStream.Write(buffer, 0, size);
                        }
                        while (size > 0);

                        zipOutStream.Seek(0, SeekOrigin.Begin);
                    }
                }
            }

            GC.Collect();
        }

        public static void ZipFile(string innputFilePath, string outputFilePath)
        {
            byte[] buffer = new byte[8000];

            using (FileStream readFileStream = new FileStream(innputFilePath, FileMode.Open, FileAccess.Read))
            {
                using (FileStream writeFileStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                {
                    using (ZipOutputStream zipOutStream = new ZipOutputStream(writeFileStream))
                    {
                        zipOutStream.SetLevel(9);
                        zipOutStream.PutNextEntry(new ZipEntry(Path.GetFileName(outputFilePath)));

                        if (readFileStream.CanRead)
                        {
                            readFileStream.Seek(0, SeekOrigin.Begin);

                            int size;
                            do
                            {
                                size = readFileStream.Read(buffer, 0, buffer.Length);
                                zipOutStream.Write(buffer, 0, size);
                            }
                            while (size > 0);
                        }
                    }
                }
            }

            new FileInfo(innputFilePath).Delete();

            GC.Collect();
        }

        public static void UnzipFile(string filePath, Stream stream)
        {
            byte[] buffer = new byte[8000];

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (ZipInputStream zipInputStream = new ZipInputStream(fileStream))
                {
                    ZipEntry entry = zipInputStream.GetNextEntry();

                    int size;
                    do
                    {
                        size = zipInputStream.Read(buffer, 0, buffer.Length);
                        stream.Write(buffer, 0, size);
                    } while (size > 0);

                    stream.Seek(0, SeekOrigin.Begin);
                }
            }

            GC.Collect();
        }

        public static void UnzipFile(string innputFilePath, string outputFilePath)
        {
            byte[] buffer = new byte[8000];

            using (FileStream fileStream = new FileStream(innputFilePath, FileMode.Open, FileAccess.Read))
            {
                using (FileStream writeFileStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                {
                    using (ZipInputStream zipInputStream = new ZipInputStream(writeFileStream))
                    {
                        ZipEntry entry = zipInputStream.GetNextEntry();

                        int size;
                        do
                        {
                            size = zipInputStream.Read(buffer, 0, buffer.Length);
                            writeFileStream.Write(buffer, 0, size);
                        } while (size > 0);
                    }
                }
            }

            FileInfo fileInfo = new FileInfo(innputFilePath);

            if (fileInfo != null)
                fileInfo.Delete();

            GC.Collect();
        }

        public static string UnzipBytes(byte[] bytes)
        {
            StringBuilder result = new StringBuilder();

            try
            {
                byte[] buffer = new byte[8000];

                using (var stream = new ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream(new MemoryStream(bytes)))
                {
                    int size;
                    do
                    {
                        size = stream.Read(buffer, 0, buffer.Length);
                        result.Append(Encoding.Default.GetString(buffer, 0, size));
                    } while (size > 0);
                }
            }
            catch (Exception exp)
            {
                Log.Instance.Info("Utils.IOHelper.UnzipBytes(...) exception:");
                Log.Instance.Info(exp.Message);
                Log.Instance.Info(exp.Source);
                Log.Instance.Info(exp.StackTrace);
            }

            return result.ToString();
        }

        public static void CreateDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(path);

                    FileInfo[] filesInfo = directoryInfo.GetFiles();

                    for (int i = 0; i < filesInfo.Length; i++)
                        filesInfo[i].Delete();
                }
                else
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch (Exception exp)
            {
                Log.Instance.Info("Utils.IOHelper.CreateDirectory(...) exception:");
                Log.Instance.Info(exp.Message);
                Log.Instance.Info(exp.Source);
                Log.Instance.Info(exp.StackTrace);
            }
        }

        public static void DeleteDirectory(string path)
        {
            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(path);

                DirectoryInfo[] directoriesInfo = directoryInfo.GetDirectories();

                for (int i = 0; i < directoriesInfo.Length; i++)
                    DeleteDirectory(directoriesInfo[i].FullName);

                FileInfo[] filesInfo = directoryInfo.GetFiles();

                for (int i = 0; i < filesInfo.Length; i++)
                    filesInfo[i].Delete();

                directoryInfo.Delete();
            }
            catch (Exception exp)
            {
                Log.Instance.Info("Utils.IOHelper.DeleteDirectory(...) exception:");
                Log.Instance.Info(exp.Message);
                Log.Instance.Info(exp.Source);
                Log.Instance.Info(exp.StackTrace);
            }
        }
    }
}
