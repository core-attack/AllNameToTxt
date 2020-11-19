using System;
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
    }
}
