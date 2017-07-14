using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace winfrm_OpenXML_Word_Lab
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGenDocx_Click(object sender, EventArgs e)
        {
            //# 預計產生文件內容如下：
            // 其 WordprocessingML 如下所示:
            // < w:document xmlns:w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
            //   < w:body >
            //     < w:p >
            //       < w:r >
            //         < w:t > 在 body 本文內容產生 text 文字</ w:t >
            //       </ w:r >
            //     </ w:p >
            //   </ w:body >
            // </ w:document >

            // 檔案路徑檔名、請注意副檔名與 WordprocessingDocumentType.Document 一致
            string filepath = txtOutFile.Text.Trim(); // @"C:\temp\test.docx";

            // 建立 WordprocessingDocument 類別，透過 WordprocessingDocument 類別中的 Create 方法建立 Word 文件
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // 建立 MainDocumentPart 類別物件 mainPart，加入主文件部分 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                // 實例化 Document(w) 部分
                mainPart.Document = new Document();
                // 建立 Body 類別物件，於加入 Doucment(w) 中加入 Body 內文
                Body body = mainPart.Document.AppendChild(new Body());
                // 建立 Paragraph 類別物件，於 Body 本文中加入段落 Paragraph(p)
                Paragraph paragraph = body.AppendChild(new Paragraph());
                // 建立 Run 類別物件，於 段落 Paragraph(p) 中加入文字屬性 Run(r) 範圍
                Run run = paragraph.AppendChild(new Run());
                // 在文字屬性 Run(r) 範圍中加入文字內容
                run.AppendChild(new Text("在 body 本文內容產生 text 文字"));
            }

            MessageBox.Show("執行完成。");
        }

        private void btnApplyDocx_Click(object sender, EventArgs e)
        {
            string template = txtTplFile.Text.Trim(); // @"C:\Temp\Word套用範本檔案測試.docx";
            string outputFile = txtOutFile.Text.Trim();

            Dictionary<string, string> dct = new Dictionary<string, string>()
            {
                { "idToday", txtToday.Text.Trim() }, // DateTime.Today.ToString("yyyy-MM-dd")
                { "idName",  txtName.Text.Trim()  },
                { "idAddress", txtAddress.Text.Trim() },
                { "idAmount",  txtAmount.Text.Trim() }
            };

            // GO
            File.Copy(template, outputFile, true);

            // 建立 WordprocessingDocument 類別，透過 WordprocessingDocument 類別中的 Create 方法建立 Word 文件
            using (WordprocessingDocument wd = WordprocessingDocument.Open(outputFile, true))
            {
                // 建立 MainDocumentPart 類別物件 mainPart，加入主文件部分 
                MainDocumentPart mainPart = wd.MainDocumentPart;
                Document docx = mainPart.Document;

                var fields4 = docx.Body.Descendants<SdtElement>();
                foreach (SdtElement field in fields4)
                {
                    var title = field.Descendants<SdtAlias>().First().Val;
                    var tag = field.Descendants<Tag>().First().Val;
                    Debug.WriteLine(string.Format("{0} : {1} : {2}", tag, title, field.GetType().Name));
                    Debug.WriteLine(field.InnerText);

                    // 取套表的內容
                    string newContentText = dct[tag];

                    // 套入目標欄位 replace field content
                    field.Descendants<Text>().First().Text = newContentText;

                    //※ 注意：仍需注意Document下node的階層關係。
                }

                docx.Save();
            }

            MessageBox.Show("執行完成。");
        }

        private void btnApplyDocx_Click_old(object sender, EventArgs e)
        {
            string template = txtTplFile.Text.Trim(); // @"C:\Temp\Word套用範本檔案測試.docx";
            string outputFile = txtOutFile.Text.Trim();

            Dictionary<string, string> dct = new Dictionary<string, string>()
            {
                { "idToday", txtToday.Text.Trim() }, // DateTime.Today.ToString("yyyy-MM-dd")
                { "idName",  txtName.Text.Trim()  },
                { "idAddress", txtAddress.Text.Trim() },
                { "idAmount",  txtAmount.Text.Trim() }
            };

            // GO
            File.Copy(template, outputFile, true);

            // 建立 WordprocessingDocument 類別，透過 WordprocessingDocument 類別中的 Create 方法建立 Word 文件
            using (WordprocessingDocument wd = WordprocessingDocument.Open(outputFile, true))
            {
                // 建立 MainDocumentPart 類別物件 mainPart，加入主文件部分 
                MainDocumentPart mainPart = wd.MainDocumentPart;
                Document docx = mainPart.Document;

                var fields4 = docx.Body.Descendants<SdtElement>();
                foreach (SdtElement field in fields4)
                {
                    var title = field.Descendants<SdtAlias>().First().Val;
                    var tag = field.Descendants<Tag>().First().Val;
                    Debug.WriteLine(string.Format("{0} : {1} : {2}", tag, title, field.GetType().Name));
                    Debug.WriteLine(field.InnerText);

                    // 取套表的內容
                    string contentText = dct[tag];

                    // 套入目標欄位 replace field content
                    if (field is SdtRun)
                    {

                        SdtContentRun oldContent = field.GetFirstChild<SdtContentRun>();

                        //String newContentXml = "<w:sdtContent xmlns:w = \"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>我是天才。</w:t></w:r></w:sdtContent>";
                        StringBuilder newContentXml = new StringBuilder("<w:sdtContent xmlns:w = \"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>");
                        newContentXml.Append(contentText);
                        newContentXml.Append("</w:t></w:r></w:sdtContent>"); // postfix 

                        SdtContentRun newContent = new SdtContentRun(newContentXml.ToString());
                        field.ReplaceChild<SdtContentRun>(newContent, oldContent);
                    }
                    else if (field is SdtCell)
                    {
                        SdtContentCell oldContent = field.GetFirstChild<SdtContentCell>();
                        //String newContentXml = "<w:sdtContent xmlns:w = \"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tc><w:p><w:r><w:t>我是天才。</w:t></w:r></w:p></w:tc></w:sdtContent>";
                        StringBuilder newContentXml = new StringBuilder("<w:sdtContent xmlns:w = \"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tc><w:p><w:r><w:t>");
                        newContentXml.Append(contentText);
                        newContentXml.Append("</w:t></w:r></w:p></w:tc></w:sdtContent>");

                        SdtContentCell newContent = new SdtContentCell(newContentXml.ToString());
                        field.ReplaceChild<SdtContentCell>(newContent, oldContent);
                    }
                }

                docx.Save();
            }

            MessageBox.Show("執行完成。");
        }


    }
}
