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
    public partial class Manual : Form
    {
        //<b></b> - выделить текст жирным
        //<i></i> - выделить текст курсивом
        //<u></u> - выделить текст подчеркиванием
        //<с> - центрировать
        //<l> - по левому краю
        //<r> - по правому краю
        string bold = "<b>";
        string boldE = "</b>";
        string italic = "<i>";
        string italicE = "</i>";
        string under = "<u>";
        string underE = "</u>";
        string middle = "<m>";
        string middleE = "</m>";
        string center = "<c>";
        string left = "<l>";
        string right = "<r>";
        string defaultText = "";

        public Manual()
        {
            InitializeComponent();
            richTextBox1.ScrollBars = RichTextBoxScrollBars.Vertical;
            richTextBox1.BackColor = Color.White;
            разрешитьРедактированиеToolStripMenuItem.Checked = true;
            setFont();
        }

        void setFont()
        {
            Graphics gr = CreateGraphics();
            FontFamily[] fonts = FontFamily.GetFamilies(gr);
            foreach (FontFamily font in fonts)
            {
                if (font.IsStyleAvailable(FontStyle.Regular))
                {
                    fontBox.Items.Add(font.Name);
                    if (font.Name == this.Font.Name)
                    {
                        fontBox.Text = this.Font.Name;
                    }
                }
            }
            int[] intt = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            for (int i = 0; i < intt.Length; i++)
                fontSize.Items.Add(intt[i]);
            fontSize.Text = "12";
        }

        public void downloadText(string s)
        {
            richTextBox1.Text = s;
            defaultText = s;
        }

        public void downloadText(string[] s)
        {

            foreach (string str in s)
            { 
                richTextBox1.Text += str + "\n";
                defaultText += str + "\n";
            }
            selectText();
        }

        private void manual_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripItem)
            {
                switch (((ToolStripItem)sender).Text)
                {
                    case "Жирный": {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Bold);
                    }
                        break;
                    case "Курсив":
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Italic);
                        }
                        break;
                    case "Подчеркивание":
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Underline);
                        }
                        break;
                    case "Зачеркивание":
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Strikeout);
                        }
                        break;
                    case "Центрировать": {
                        richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
                    }
                        break;
                    case "По левому краю": { richTextBox1.SelectionAlignment = HorizontalAlignment.Left; }
                        break;
                    case "По правому краю": { richTextBox1.SelectionAlignment = HorizontalAlignment.Right; }
                        break;
                    case "По обоим краям": {}
                        break;
                    case "Выделить всё": { richTextBox1.SelectAll(); }
                        break;
                    case "Отменить": { richTextBox1.Undo(); }
                        break;
                    case "Вернуть": { richTextBox1.Redo(); }
                        break;
                    case "Вырезать": {
                        //Clipboard.SetText(richTextBox1.SelectedText);
                        //richTextBox1.Text = richTextBox1.Text.Remove(richTextBox1.SelectionStart, richTextBox1.SelectionLength);
                        richTextBox1.Cut();
                    }
                        break;
                    case "Копировать": { 
                        //Clipboard.SetText(richTextBox1.SelectedText); 
                        richTextBox1.Copy();
                    }
                        break;
                    case "Вставить":
                        { 
                            //richTextBox1.SelectedText = Clipboard.GetText();
                            richTextBox1.Paste();
                        }
                        break;
                    case "Удалить": { 
                        //Удаления у меня не будет, потому что сбрасыва.тся все выделения из-за него
                        //richTextBox1.Text = richTextBox1.Text.Remove(richTextBox1.SelectionStart, richTextBox1.SelectionLength);
                        //richTextBox1.AutoWordSelection
                    }
                        break;
                    case "Только чтение": { разрешитьРедактированиеToolStripMenuItem.Checked = !разрешитьРедактированиеToolStripMenuItem.Checked; }
                        break;
                }
            }
        }

        class selections
        {
            //позиция начала выделения жирного
            public int selectionStart = 0;
            //позиция окончания выделения жирного
            public int selectionEnd = 0;

            //какое выделение
            public bool bold = false;
            public bool italic = false;
            public bool under = false;
            public bool middle = false;
            public bool center = false;
            public bool left = false;
            public bool right = false;
            //индекс строки выделения в массиве всех строк
            public int index = 0;
        }

        
        //каждая строка может иметь несколько выделений
        List<List<selections>> listListSelections = new List<List<selections>>();

        void selectText()
        {
            try
            {
                List<string> rtbList = new List<string>();
                for (int i = 0; i < richTextBox1.Lines.Length; i++)
                {
                    rtbList.Add(richTextBox1.Lines[i]);
                }
                for (int i = 0; i < rtbList.Count; i++)
                {
                    List<selections> listSelections = new List<selections>();
                    string s = rtbList[i];
                    string sWhisOnlyTag = "";
                    sWhisOnlyTag = deleteTags(s, bold, boldE);
                    while (sWhisOnlyTag.IndexOf(bold) != -1)
                    {
                        selections sel = new selections();
                        sel.bold = true;
                        int Start = sWhisOnlyTag.IndexOf(bold);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(Start, bold.Length);
                        int End = sWhisOnlyTag.IndexOf(boldE);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(End, boldE.Length);
                        if (End == -1)
                            End = Start;
                        sel.index = i;
                        sel.selectionStart = Start;
                        sel.selectionEnd = End;
                        listSelections.Add(sel);
                    }
                    sWhisOnlyTag = deleteTags(s, italic, italicE);
                    while (sWhisOnlyTag.IndexOf(italic) != -1)
                    {
                        selections sel = new selections();
                        sel.italic = true;
                        int Start = sWhisOnlyTag.IndexOf(italic);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(Start, italic.Length);
                        int End = sWhisOnlyTag.IndexOf(italicE);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(End, italicE.Length);
                        if (End == -1)
                            End = Start;
                        sel.index = i;
                        sel.selectionStart = Start;
                        sel.selectionEnd = End;
                        listSelections.Add(sel);
                    }
                    sWhisOnlyTag = deleteTags(s, under, underE);
                    while (sWhisOnlyTag.IndexOf(under) != -1)
                    {
                        selections sel = new selections();
                        sel.under = true;
                        int Start = sWhisOnlyTag.IndexOf(under);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(Start, under.Length);
                        int End = sWhisOnlyTag.IndexOf(underE);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(End, underE.Length);
                        if (End == -1)
                            End = Start;
                        sel.index = i;
                        sel.selectionStart = Start;
                        sel.selectionEnd = End;
                        listSelections.Add(sel);
                    }
                    sWhisOnlyTag = deleteTags(s, middle, middleE);
                    while (sWhisOnlyTag.IndexOf(middle) != -1)
                    {
                        selections sel = new selections();
                        sel.middle = true;
                        int Start = sWhisOnlyTag.IndexOf(middle);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(Start, middle.Length);
                        int End = sWhisOnlyTag.IndexOf(middleE);
                        sWhisOnlyTag = sWhisOnlyTag.Remove(End, middleE.Length);
                        if (End == -1)
                            End = Start;
                        sel.index = i;
                        sel.selectionStart = Start;
                        sel.selectionEnd = End;
                        listSelections.Add(sel);
                    }
                    if (s.IndexOf(center) != -1)
                    {
                        selections sel = new selections();
                        sel.center = true;
                        sel.index = i;
                        listSelections.Add(sel);
                    }
                    if (s.IndexOf(left) != -1)
                    {
                        selections sel = new selections();
                        sel.left = true;
                        sel.index = i;
                        listSelections.Add(sel);
                    }
                    if (s.IndexOf(right) != -1)
                    {
                        selections sel = new selections();
                        sel.right = true;
                        sel.index = i;
                        listSelections.Add(sel);
                    }
                    listListSelections.Add(listSelections);
                    rtbList[i] = deleteTags(s);

                }
                richTextBox1.Text = "";
                int length = 0;
                for (int i = 0; i < rtbList.Count; i++)
                {
                    richTextBox1.AppendText(rtbList[i] + "\n");
                    for (int j = 0; j < listListSelections[i].Count; j++)
                    {
                        int end = Math.Abs(listListSelections[i][j].selectionEnd - listListSelections[i][j].selectionStart);
                        if (i != 0)
                            richTextBox1.Select(listListSelections[i][j].selectionStart + length, end);
                        else
                            richTextBox1.Select(listListSelections[i][j].selectionStart, Math.Abs(listListSelections[i][j].selectionEnd - listListSelections[i][j].selectionStart));
                        if (listListSelections[i][j].bold)
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Bold);
                        }
                        if (listListSelections[i][j].italic)
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Italic);
                        }
                        if (listListSelections[i][j].under)
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Underline);
                        }
                        if (listListSelections[i][j].middle)
                        {
                            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Strikeout);
                        }
                        if (listListSelections[i][j].center)
                            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
                        if (listListSelections[i][j].left)
                            richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                        if (listListSelections[i][j].right)
                            richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
                    }
                    length += rtbList[i].Length + 1;
                    int lrtb = richTextBox1.Text.Length;
                }
                delteTags();
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace, "Проверьте синтаксис использованных тегов!");

            }
        }

        
        void delteTags()
        {
            string[] tags = { bold, boldE, italic, italicE, under, underE, middle, middleE, center, left, right };
            foreach (string tag in tags)
                richTextBox1.Text.Replace(tag, "");
        }

        string deleteTags(string s)
        {
            string[] tags = { bold, boldE, italic, italicE, under, underE, middle, middleE, center, left, right };
            foreach (string tag in tags)
                    s = s.Replace(tag, "");
            return s;
        }

        //удаляет все теги, кроме двух указаных
        string deleteTags(string s, string exceptionTag1, string exceptionTag2)
        {
            string[] tags = { bold, boldE, italic, italicE, under, underE, middle, middleE, center, left, right };
            foreach (string tag in tags)
                if (exceptionTag1 != tag && exceptionTag2 != tag)
                    s = s.Replace(tag, "");
            return s;
        }

        private void всёВОднуСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = richTextBox1.Text.Replace("\n","");
        }

        private void fontSize_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(Convert.ToChar(e.KeyChar)) && e.KeyChar != ',' && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void richTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //e.Handled = true;
            
        }

        private void fontBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                listListSelections.Clear();
                richTextBox1.Text = defaultText;
                richTextBox1.Font = new Font(fontBox.Text, (float)Convert.ToDouble(fontSize.Text));
                selectText();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void разрешитьРедактированиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void разрешитьРедактированиеToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if (разрешитьРедактированиеToolStripMenuItem.Checked)
            {
                richTextBox1.ContextMenuStrip.Items[toolStripMenuItem3.Name].Enabled = false;
                richTextBox1.ContextMenuStrip.Items[toolStripMenuItem5.Name].Enabled = false;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem7.Name].Enabled = false;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem9.Name].Enabled = false;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem12.Name].Enabled = false;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem13.Name].Enabled = false;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem14.Name].Enabled = false;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem15.Name].Enabled = false;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[отменитьToolStripMenuItem.Name].Enabled = false;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[вернутьToolStripMenuItem.Name].Enabled = false;
                richTextBox1.ContextMenuStrip.Items[жирныйToolStripMenuItem.Name].Enabled = false;
                richTextBox1.ContextMenuStrip.Items[курсивToolStripMenuItem.Name].Enabled = false;
                richTextBox1.ContextMenuStrip.Items[подчеркиваниеToolStripMenuItem.Name].Enabled = false;
                richTextBox1.ContextMenuStrip.Items[зачеркиваниеToolStripMenuItem.Name].Enabled = false;

                richTextBox1.ReadOnly = true;
            }
            else
            {
                richTextBox1.ContextMenuStrip.Items[toolStripMenuItem3.Name].Enabled = true;
                richTextBox1.ContextMenuStrip.Items[toolStripMenuItem5.Name].Enabled = true;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem7.Name].Enabled = true;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem9.Name].Enabled = true;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem12.Name].Enabled = true;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem13.Name].Enabled = true;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem14.Name].Enabled = true;
                ((ToolStripMenuItem)((ToolStripMenuItem)menuStrip1.Items[текстToolStripMenuItem.Name]).DropDownItems[выделенияToolStripMenuItem.Name]).DropDownItems[toolStripMenuItem15.Name].Enabled = true;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[отменитьToolStripMenuItem.Name].Enabled = true;
                ((ToolStripMenuItem)menuStrip1.Items[правкаToolStripMenuItem.Name]).DropDownItems[вернутьToolStripMenuItem.Name].Enabled = true;
                richTextBox1.ContextMenuStrip.Items[жирныйToolStripMenuItem.Name].Enabled = true;
                richTextBox1.ContextMenuStrip.Items[курсивToolStripMenuItem.Name].Enabled = true;
                richTextBox1.ContextMenuStrip.Items[подчеркиваниеToolStripMenuItem.Name].Enabled = true;
                richTextBox1.ContextMenuStrip.Items[зачеркиваниеToolStripMenuItem.Name].Enabled = true;
                richTextBox1.ReadOnly = false;
            }
        }

    }
}
