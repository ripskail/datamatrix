using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using ZXing.Datamatrix;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing.Imaging;
using System.IO;

namespace datamatrix
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string[] q;
        private void button1_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();

            String[] s = textBox1.Text.Split(new String[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            List<PictureBox> picturebox = new List<PictureBox>();
            QRCodeWriter qrEncode = new QRCodeWriter(); //создание QR кода
            BarcodeWriter data = new BarcodeWriter();
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            q = new string[500];
            var y = 15;
            for (int i = 0; i < s.Length; i++)
            {
                
                string strRUS = textBox1.Text;  //строка на русском языке
               
                Dictionary<EncodeHintType, object> hints = new Dictionary<EncodeHintType, object>();    //для колекции поведений
                hints.Add(EncodeHintType.CHARACTER_SET, "utf-8");   //добавление в коллекцию кодировки utf-8

                var bw = new BarcodeWriter
                {
                    Format = BarcodeFormat.DATA_MATRIX,
                    Options = new EncodingOptions { Width = 80, Height = 80 }
                };
                var img = bw.Write(s[i]);

                var pb = new PictureBox();
                pb.Location = new Point(picturebox.Count * 60 + 50, y);
                pb.Size = new Size(50, 50);
                try
                {
                    pb.Image = img;
                }
                catch (OutOfMemoryException) { continue; }
                pb.SizeMode = PictureBoxSizeMode.StretchImage;
                pb.Name = "pic" + i;
                flowLayoutPanel1.Controls.Add(pb);
                picturebox.Add(pb);
                string put = @"C:\Users\Public\Documents\" + i+".jpg";
                pb.Image.Save(put, ImageFormat.Jpeg);
                q[i] = put;

                //pictureBox1.Image = img;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Получить объект приложения Word.
            Word._Application word_app = new Word.Application();

            // Сделать Word видимым (необязательно).
            word_app.Visible = false;

            // Создаем документ Word.
            object missing = Type.Missing;
            Word._Document word_doc = word_app.Documents.Add(
                ref missing, ref missing, ref missing, ref missing);
            for (int i = 0; i < q.Length; i++)
            {
                Object oMissed = word_doc.Paragraphs[1].Range;
                Object oLinkToFile = false;
                Object oSaveWithDocument = true;
                if (q[i] != null)
                {
                    word_doc.InlineShapes.AddPicture(q[i], ref oLinkToFile, ref oSaveWithDocument, ref oMissed);
                    //  object fileName = saveFileDialog1.ToString();// @"C:\Test\NewDocument.docx";
                    File.Delete(q[i]);
                }
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "|*.docx";
            saveFileDialog1.Title = "Save the Word Document";
            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                string docName = saveFileDialog1.FileName;
                if (docName.Length > 0)
                {
                    object oDocName = (object)docName;
                    word_doc.SaveAs(ref oDocName, ref missing, ref missing, ref missing, ref missing, ref missing,
                                 ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                 ref missing, ref missing, ref missing, ref missing);
                }
            }
            word_doc.Close();
            word_app.Quit();
        }
    }
}
