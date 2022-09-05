
using ConsoleApp3.Models;

using Spire.Doc;
using Spire.Doc.Documents;
using System;
using Spire.Doc.Fields;

using System.Drawing;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xceed.Words.NET;

namespace ConsoleApp3
{
    internal class Program
    {
        static void Main(string[] args)
        {
           
            Document doc = new Document();
            doc.LoadFromFile(@"C:\Users\PIZONE\Documents\copy_template.docx");
            Section s1=doc.Sections[0];
            Paragraph p0=s1.Paragraphs[0];
            p0.Format.HorizontalAlignment=HorizontalAlignment.Center;
            DocPicture picture = new DocPicture(doc);
            string txtUrl = "https://www.bing.com/th?id=OIP.7acp2NExpYoVMo6rbOUYNQHaHS&w=150&h=147&c=8&rs=1&qlt=90&o=6&pid=3.1&rm=2";
            Image imag = DownloadImageFromUrl(txtUrl);
            picture.LoadImage(imag);
            p0.AppendHyperlink("https://www.omdbapi.com/", picture, HyperlinkType.WebLink);
            Paragraph p = s1.Paragraphs[3];
            p.Text = "28/7/2022,Yamun Palace";

                 Image DownloadImageFromUrl(string imageUrl)
                {
                    Image image = null;

                    try
                    {
                        System.Net.HttpWebRequest webRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(imageUrl);
                        webRequest.AllowWriteStreamBuffering = true;
                        webRequest.Timeout = 30000;

                        System.Net.WebResponse webResponse = webRequest.GetResponse();

                        System.IO.Stream stream = webResponse.GetResponseStream();

                        image = System.Drawing.Image.FromStream(stream);

                        webResponse.Close();
                    }
                    catch (Exception ex)
                    {
                        return null;
                    }

                    return image;
                }

            
            String[] Header = { "Name", "Email" };
            String[][] data =
                { 
                new String[]{ "abhi", "xyz@gmail.com"},
                new String[]{ "Priya","abx@gmail.com"},
                new String[]{ "Nabin","plz@gmail.com"},
                 };
            Table table = s1.Tables[0] as Table;
            Console.WriteLine(table.TableDescription);

            int i = 1;
            for (int r = 0; r < data.Length; r++)

            {
                TableRow row = table.AddRow();
                table.Rows.Insert(i, row);

          
                for (int c = 0; c < data[r].Length; c++)

                {
                    TableRow DataRow = table.Rows[i];
                    Paragraph p4 = DataRow.Cells[c].AddParagraph();
                    Spire.Doc.Fields.TextRange TR2 = p4.AppendText(data[r][c]);
                    DataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                   
                    p4.Format.HorizontalAlignment = HorizontalAlignment.Center;


                    TR2.CharacterFormat.FontName = "Calibri";

                    TR2.CharacterFormat.FontSize = 11;

                }
                i++;

            }




            doc.SaveToFile("updated_fil.docx", FileFormat.Docx);

            Document newdoc = new Document();




            newdoc.LoadFromFile(@"C:\Users\PIZONE\Desktop\web\ConsoleApp3\ConsoleApp3\bin\Debug\updated_fil.docx");
            newdoc.SaveToFile("Convert.PDF", FileFormat.PDF);
            //Launch Document  
            System.Diagnostics.Process.Start("Convert.PDF");

           
        }
    }
}
