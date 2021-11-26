using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoUpdate
{
    internal class Common
    {
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="fileFullPath"></param>
        /// <returns></returns>
        private bool DeleteFile(string fileFullPath)
        {
            if (File.Exists(fileFullPath))
            {
                // 2、根据路径字符串判断是文件还是文件夹
                FileAttributes attr = File.GetAttributes(fileFullPath);
                // 3、根据具体类型进行删除
                if (attr == FileAttributes.Directory)
                {
                    // 3.1、删除文件夹
                    Directory.Delete(fileFullPath, true);
                }
                else
                {
                    // 3.2、删除文件
                    File.Delete(fileFullPath);

                }
                File.Delete(fileFullPath);
            }
            return true;
        }
    }
}
