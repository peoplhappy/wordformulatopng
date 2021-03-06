# wordimage
word中所有公式转为图片

转换思路：
1.先将docx文件转为html，其中公式或者图片按照base64编码直接写入html中（因如果不使用此种方式，直接通过 python-docx提取图片数量可能会小于实际含有的图片数量
导致转换失败）
2.提取html中的所有img标签的内容，将src重新写入文件中，调用c#程序直接转为png后转成base64编码，替换回html中
3.将html转化为docx文件，就是最后的结果了，pydocx转换的html有如下规律
1）所有内容都在body中
2）paragraph对应为p标签，遇到直接再新建的文档中添加paragraph即可
3）所有文字以run类型直接添加到paragraph
4）遇到图片直接再run中添加图片即可，注意大小，大小的计算方式（获取html中的width和height，并除以80转为Inches就是较合适的大小）
5）table处理较为特别，table是不会存在于p标签中，而是与p标签并列，遇到table标签后，获取tr和td的数量并创建table，通过循环遍历方式获取cell，
再cell中通过添加paragraph来插入内容或图片（处理与上述一致）

C# code：
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;

namespace Project1
{
    class Class2
    {
        static void Main(string[] args)
         {
             String FileName = args[0];
             using (System.Drawing.Imaging.Metafile img = new System.Drawing.Imaging.Metafile(FileName))
             {
                 System.Drawing.Imaging.MetafileHeader header = img.GetMetafileHeader();
                 float scale = header.DpiX / 96f;
                 using (System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap((int)(scale * img.Width / header.DpiX * 100), (int)(scale * img.Height / header.DpiY * 100)))
                 {
                     using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bitmap))
                     {
                         g.Clear(System.Drawing.Color.White);
                         g.ScaleTransform(scale, scale);
                         g.DrawImage(img, 0, 0);
                     }
                     bitmap.Save(@args[1], System.Drawing.Imaging.ImageFormat.Png);
                 }
             }
             Console.WriteLine("转换完成");
         }
    }
}
