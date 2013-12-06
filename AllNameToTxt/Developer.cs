using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AllNameToTxt
{
    public partial class Developer : Form
    {
        public Developer()
        {
            InitializeComponent();
        }

        private void textBoxName_Click(object sender, EventArgs e)
        {
            try
            {
                if (sender is TextBox)
                {
                    switch (((TextBox)sender).Name)
                    {
                        case "textBoxName":
                            {
                                Clipboard.SetText(textBoxName.Text);
                                lbl.Text = "ФИО скопированы в буфер."; 
                            }
                            break;
                        case "textBoxPhone":
                            {
                                Clipboard.SetText(textBoxPhone.Text);
                                lbl.Text = "Телефон скопирован в буфер.";
                            }
                            break;
                        case "textBoxMail":
                            {
                                Clipboard.SetText(textBoxMail.Text);
                                lbl.Text = "E-mail скопирован в буфер.";
                            }
                            break;
                        case "textBoxICQ":
                            {
                                Clipboard.SetText(textBoxICQ.Text);
                                lbl.Text = "ICQ скопирован в буфер.";
                            }
                            break;
                        case "textBoxVK":
                            {
                                Clipboard.SetText(textBoxVK.Text);
                                lbl.Text = "Ссылка скопирована в буфер.";
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void SendMail(string otpravitel, string poluchatel, string title, string text)// отправляет письмо
        {
            System.Net.Mail.MailMessage mm = new System.Net.Mail.MailMessage();
            mm.From = new System.Net.Mail.MailAddress(otpravitel);
            mm.To.Add(new System.Net.Mail.MailAddress(poluchatel));
            mm.Subject = title;
            //mm.IsBodyHtml = true;//письмо в html формате (если надо)
            mm.Body = text;
            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(System.Net.Dns.GetHostByName("LocalHost").HostName);
            client.Send(mm);//поехало
        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            SendMail(textBoxFor.Text, textBoxTo.Text, textBoxTitle.Text, richTextBoxText.Text);
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            panelMail.Visible = false;
        }

        private void labelName_Click(object sender, EventArgs e)
        {
            try
            {
                if (sender is Label)
                {
                    switch (((Label)sender).Name)
                    {
                        case "labelName":
                            {
                                Clipboard.SetText(textBoxName.Text);
                                lbl.Text = "ФИО скопированы в буфер.";
                            }
                            break;
                        case "labelPhone":
                            {
                                Clipboard.SetText(textBoxPhone.Text);
                                lbl.Text = "Телефон скопирован в буфер.";
                            }
                            break;
                        case "labelMail":
                            {
                                Clipboard.SetText(textBoxMail.Text);
                                lbl.Text = "E-mail скопирован в буфер.";
                            }
                            break;
                        case "labelICQ":
                            {
                                Clipboard.SetText(textBoxICQ.Text);
                                lbl.Text = "ICQ скопирован в буфер.";
                            }
                            break;
                        case "labelVK":
                            {
                                Clipboard.SetText(textBoxVK.Text);
                                lbl.Text = "Ссылка скопирована в буфер.";
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBoxVK.Text);
        }



        //Для преобразования телефонного номера получателя и SMS-центра можно воспользоваться вот такой функцией
        //private string EncodeNumber(string Number)
        //{
        //    string result = "";

        //    Number = Number.Replace("+", "");
        //    if ((Number.Length % 2) > 0)
        //        Number += "F";
        //    int i = 0;
        //    while (i < Number.Length)
        //    {
        //        result += Number[i + 1].ToString() + Number[i].ToString();
        //        i += 2;
        //    }
        //    return result;
        //}

        ////Нам остается преобразовать текст, и вот как это делается в C#.
        ////String7To8 функция перекодирует ASCII символы.
        //public string String7To8(string str)
        //{
        //    string result = "";
        //    ASCIIEncoding enc = new ASCIIEncoding();
        //    byte[] arr =  enc.GetBytes(str);
        //    int i = 1;
        //    while (i < arr.Length)
        //    {
        //        int j = arr.Length - 1;
        //        while (j >= i)
        //        {
        //            byte firstBit = (arr[j] % 2 > 0) ? (byte)0x80 : (byte)0x00;
        //            arr[j - 1] = (byte)((arr[j - 1] & 0x7f) | firstBit);
        //            arr[j] = (byte)(arr[j] >> 1);
        //            j--;
        //        }
        //        i++;
        //    }
        //    i = 0;
        //    while ((i < arr.Length) && (arr[i] != 0))
        //    {
        //        result += arr[i].ToString("X2");
        //        i++;
        //    }
        //    return result;
        //}

        ////Для кодировки UCS2  можно воспользоватся функцией StringToUCS2.
        //public string StringToUCS2(string str)
        //{
        //    UnicodeEncoding ue = new UnicodeEncoding();
        //    byte[] ucs = ue.GetBytes(str);
        //    int i = 0;
        //    while (i < ucs.Length)
        //    {
        //        byte b = ucs[i + 1];
        //        ucs[i + 1] = ucs[i];
        //        ucs[i] = b;
        //        i += 2;
        //    }
        //    return BitConverter.ToString(ucs2).Replace("-", "");
        //}
        ////Для перекодировки из PDU  в нормальный вид номер абонента и SMS-центра надо воспользоваться функцией DecodeNumber.
        //private string DecodeNumber(string Number)
        //{
        //    string result = "";
        //    int i = 0;
        //    while (i < Number.Length)
        //    {
        //        result += Number[i + 1].ToString() + Number[i].ToString();
        //        i += 2;
        //    }

        //    result = result.Replace("F", "");

        //    return result;
        //}

    }
}
