using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using AllNameToTxt.Models;
using AllNameToTxt.Models.Audition;
using CueSharp;
using Word = Microsoft.Office.Interop.Word; 


namespace AllNameToTxt
{
    public partial class GeneralForm : Form
    {
        public GeneralForm()
        {
            InitializeComponent();
            setMask();
            setTimeFormat();
        }

        private string _filename = "";
        private int _mouseX = 0;
        private int _mouseY = 0;
        //для восстановления удаленного значения
        private List<string> _oldDeleteValue = new List<string>();
        //для восстановления к первоначальному виду
        private List<string> _defaultView = new List<string>();
        //индекс удаленного значения
        private List<int> _oldDeleteValueIndex = new List<int>();
        private List<FileName> listFileNames = new List<FileName>();
        private string _foldersName = "";
        //мой буфер
        private FileName _clipboard = new FileName();

        private void myOpen_Click(object sender, EventArgs e)
        {
            MyselfOpen();
        }

        private void MyselfOpen()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.DefaultExt = ".txt";
            //не смещать cue файлы в списке со второй позиции! от этого зависит корректность выбора
            dialog.Filter = "Cue files(*.cue)|*.cue|MPEG layer 3(*.mp3)|*.mp3|Текстовые файлы(*.txt)|*.txt|Audition Session(*.sesx)|*.sesx";//"Текстовые файлы(*.txt)|*.txt|Playlist (*.m3u)|*.m3u|Cue (*.cue)|*.cue|Все файлы(*.*)|*.*";
            _filename = dialog.FileName;
            dialog.Multiselect = true;
            char[] sep = { '\\' };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                switch (dialog.FilterIndex)
                {
                    case 1:
                    {
                        try
                        {
                            foreach (string file in dialog.FileNames)
                            {
                                openCue(file);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Ошибка открытия cue-файла");
                        }
                    }
                    break;
                    case 2:
                    {
                        string[] s = dialog.FileNames[0].Split(sep);

                        if (s.Length > 1)
                            _foldersName = s[s.Length - 2];

                        foreach (string file in dialog.FileNames)
                        {
                            try
                            {
                                FileName mfn = new FileName();
                                mfn.OldFileName = file;
                                string sfn = shortName(file);

                                if (toolsOformList.Checked)
                                {
                                    listBoxName.Items.Add(begin.Text + sfn + end.Text);
                                    mfn.CurrentFileName = begin.Text + sfn + end.Text;
                                }
                                else
                                {
                                    listBoxName.Items.Add(sfn);
                                    mfn.CurrentFileName = sfn;

                                }

                                if (_defaultView.Count < 1000000)
                                    _defaultView.Add(sfn);

                                listFileNames.Add(mfn);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                                return;
                            }
                        }
                    }
                    break;
                    case 3:
                    {
                        try
                        {
                            foreach (string file in dialog.FileNames)
                            {
                                openTxt(file);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Ошибка открытия текстового");
                        }
                    }
                    break;
                    case 4:
                    {
                        try
                        {
                            foreach (string file in dialog.FileNames)
                            {
                                openSesx(file);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Ошибка открытия Audition Session ");
                        }
                    }
                    break;
                }
            }
        }

        private bool _defaultSave = true;
        private void mySave_Click(object sender, EventArgs e)
        {
            _defaultSave = true;
            MySave();
        }

        private void myCurrentSave_Click(object sender, EventArgs e)
        {
            _defaultSave = false;
            mySaveCurrentTxt();
        }

        private void MySave()
        {
            try
            {
                SaveFileDialog savedialog = saveFileDialog1;
                savedialog.Title = "Сохранить как ...";
                savedialog.OverwritePrompt = true;
                savedialog.CheckPathExists = true;
                savedialog.Filter =
                    "Cue (*.cue)|*.cue|PromoDJ Cue (*.pue)|*.pue|Playlist (*.m3u)|*.m3u|Текстовые файлы(*.txt)|*.txt|Все файлы(*.*)|*.*";
                savedialog.ShowHelp = true;
                char[] sep = { '\\' };
                // If selected, save
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = savedialog.FileName;
                    string strFilExtn = fileName.Remove(0, fileName.Length - 3);
                    
                    switch (strFilExtn)
                    {
                        case Constants.Txt:
                            saveToTxt(fileName);
                            break;
                        case Constants.M3u:
                            saveToM3u(fileName);
                            break;
                        case Constants.Cue:
                            saveToCue(fileName);
                            break;
                        case Constants.Pue:
                            saveToPromodjCue(fileName);
                            break;
                    }

                    textBoxChange.Text = "Данные сохранены в файл " + shortName(fileName);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void mySaveCurrentTxt()
        {
            try
            {
                //_filename = Application.StartupPath + "\\mySavedFiles\\" + ".txt";
                SaveFileDialog savedialog = saveFileDialog1;
                savedialog.Title = "Сохранить как ...";
                savedialog.OverwritePrompt = false;
                savedialog.CheckPathExists = true;
                savedialog.Filter =
                    "Текстовые файлы(*.txt)|*.txt|Все файлы(*.*)|*.*";
                savedialog.ShowHelp = true;
                char[] sep = { '\\' };
                // If selected, save
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the user-selected file name
                    string fileName = savedialog.FileName;
                    //fileName = fileName.Remove(fileName.Length - 4);
                    // Get the extension
                    string strFilExtn =
                        fileName.Remove(0, fileName.Length - 3);
                    // Save file
                    switch (strFilExtn)
                    {
                        case Constants.Txt:
                            saveToCurrentTxt(fileName);
                            break;
                        default:
                            break;
                    }
                    textBoxChange.Text = "Данные сохранены в файл " + shortName(fileName);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void saveToTxt(string fn)
        {
            string filename = shortName(fn);

            using (var outStream = new StreamWriter(filename))
            {
                if (именоватьСписокtoolStripMenuItem.Checked)
                {
                    outStream.WriteLine(listTitle.Text);
                    outStream.WriteLine("--------------------------------------------------------------------");
                }

                foreach (var t in listBoxName.Items)
                    outStream.WriteLine(t);

                outStream.Close();
            }
        }

        private void saveToM3u(string fn)
        {
            string filename = shortName(fn);

            using (var outStream = new StreamWriter(filename))
            {
                foreach (var t in listBoxName.Items)
                    outStream.WriteLine(t);

                outStream.Close();
            }
        }

        private void openCue(string filePath)
        {
            CueSheet cue = new CueSheet(filePath);
            listTitle.Text = cue.Performer + " - " + getTitle(cue.Title);
            listTitle.ToolTipText = listTitle.Text;

            foreach (var t in cue.Tracks)
            {
                string time = buildTime(t.Indices[0].Minutes.ToString(), t.Indices[0].Seconds.ToString(), t.Indices[0].Frames.ToString());
                
                if (leftTime.Checked)
                    listBoxName.Items.Add(time + " " + t.Performer + " - " + t.Title);
                else
                    listBoxName.Items.Add(t.Performer + " - " + t.Title + " " + time);
                
                FileName myfn = new FileName();
                myfn.CurrentFileName = "";
                myfn.OldFileName = "";
                myfn.Performer = t.Performer;
                myfn.Title = t.Title;
                myfn.Time = time;
                listFileNames.Add(myfn);
            }

            listBoxName.Items.Add("");
            FileName mf = new FileName();
            mf.CurrentFileName = "";
            mf.OldFileName = "";
            mf.Performer = "";
            mf.Title = "";
            mf.Time = "";
            listFileNames.Add(mf);
        }

        private void openTxt(string file)
        {
            string[] arr = File.ReadAllLines(file, Encoding.Default);
            listBoxName.Items.Add(file);

            foreach (string s in arr)
            {
                listBoxName.Items.Add(s);
                FileName myfn = new FileName();
                myfn.CurrentFileName = "";
                myfn.OldFileName = "";
                myfn.Performer = getPerformer(s);
                myfn.Title = getTitle(s);
                myfn.Time = "00:00:00";
                listFileNames.Add(myfn);
            }
        }

        void openSesx(string file)
        {
            var document = new XmlDocument();

            document.Load(file);

            var nodes = document.SelectNodes($"sesx/session/tracks/audioTrack/audioClip");

            if (nodes == null)
            {
                return;
            }

            var clips = new List<AudioClip>();

            foreach (XmlNode node in nodes)
            {
                var clip = new AudioClip();

                var nameAttr = node.Attributes?.GetNamedItem("name");
                var startAttr = node.Attributes?.GetNamedItem("startPoint");
                var endAttr = node.Attributes?.GetNamedItem("endPoint");

                if (nameAttr != null)
                {
                    clip.Name = nameAttr.Value;
                }

                if (startAttr != null)
                {
                    long.TryParse(startAttr.Value, out var start);
                    clip.Start = start;

                }

                if (endAttr != null)
                {
                    long.TryParse(endAttr.Value, out var end);
                    clip.End = end;
                }

                clips.Add(clip);
            }

            var ordered = clips.OrderBy(x => x.Start);

            foreach (var x in ordered)
            {
                listBoxName.Items.Add(x.Name);

                FileName myfn = new FileName();
                myfn.CurrentFileName = "";
                myfn.OldFileName = "";
                myfn.Performer = getPerformer(x.Name);
                myfn.Title = getTitle(x.Name);
                myfn.Time = "00:00:00";
                listFileNames.Add(myfn);
            }
        }

        private string buildTime(string m, string s, string ms)
        {
            string min = m;
            string sec = s;
            string msec = ms;

            if (m.Length == 1)
                min = "0" + min;

            if (s.Length == 1)
                sec = "0" + sec;

            if (ms.Length == 1)
                msec = "0" + msec;

            return min + ":" + sec + ":" + msec;
        }

        /// <summary>
        /// помещает время в конец или в начало строки
        /// </summary>
        private void rebiuldTime()
        {
            if (leftTime.Checked)
                for (int i = 0; i < listBoxName.Items.Count; i++)
                {
                    if (listFileNames[i].Time != "")
                        listBoxName.Items[i] = listFileNames[i].Time + " " + listFileNames[i].Performer + " - " + listFileNames[i].Title;
                }
            else
                for (int i = 0; i < listBoxName.Items.Count; i++)
                {
                    if (listFileNames[i].Time != "")
                        listBoxName.Items[i] = listFileNames[i].Performer + " - " + listFileNames[i].Title + " " + listFileNames[i].Time;
                }
        }

        private void saveToCue(string cuename)
        {
            try
            {
                OpenFileDialog od = openFileDialog1;
                od.Title = "Выберите файл, для которого следует создать cue-файл";
                od.FileName = "noname.mp3";

                if (od.ShowDialog() == DialogResult.OK)
                {
                    //сам cue-файл
                    CueSheet cue = new CueSheet(cuename);

                    if (!выключитьАвтоматическоеЗаполнениеToolStripMenuItem.Checked)
                    {
                        cue.Performer = toolStripTextBoxPerformer.Text;
                        cue.Title = toolStripTextBoxTitle.Text;
                        if (именоватьСписокtoolStripMenuItem.Checked)
                        {
                            cue.Title = listTitle.Text;
                        }
                    }
                    else
                    {
                        cue.Performer = getPerformer(shortName(od.FileName));
                        cue.Title = getTitle(shortName(od.FileName));
                    }

                    cue.Songwriter = toolStripTextBoxPerformer.Text;
                    Track trk = new Track(1, DataType.AUDIO);
                    trk.DataFile = new AudioFile(shortName(od.FileName), FileType.MP3);
                    trk.Title = "_filename";
                    trk.Performer = "_filename";
                    trk.AddIndex(1, 0, 0, 0);
                    cue.AddTrack(trk);

                    for (int i = 0; i < listFileNames.Count; i++)
                    {
                        FileInfo fi = new FileInfo(cuename);
                        if (listFileNames[i].CurrentFileName != "")
                        {
                            //перед тем, как создать cue-файл нужно сохранить в исходники измененные названия файлов
                            fi = new FileInfo(listFileNames[i].CurrentFileName);
                        }
                        if (!fi.Exists && listFileNames[i].OldFileName != "")
                        {
                            fi = new FileInfo(listFileNames[i].OldFileName);
                        }
                        string time = listFileNames[i].Time;
                        char[] c = {':'};
                        string[] subs = time.Split(c);
                        int min = 0;
                        int sec = 0;
                        int fra = 0;
                        if (subs.Length > 0 && subs[0] != "")
                            min = Convert.ToInt32(subs[0]);
                        if (subs.Length > 1 && subs[1] != "")
                            sec = Convert.ToInt32(subs[1]);
                        if (subs.Length > 2 && subs[2] != "")
                            fra = Convert.ToInt32(subs[2]);
                        //позиция и время начала
                        Index index = new Index(1, min, sec, fra); 
                        //трек
                        Track track = new Track(i + 1, DataType.AUDIO);
                        track.Performer = listFileNames[i].Performer;
                        track.Title = listFileNames[i].Title;
                        track.AddIndex(1, min, sec, fra);
                        if (track.Performer != "" && track.Title != "")
                            cue.AddTrack(track);
                    }

                    cue.SaveCue(cuename);
                    //не могу разобраться с тем, как добавить только ссылку на файл mp3, без добавления Performer и Title
                    //поэтому просто удаляю лишние три  строки и всё работает
                    string line = "";

                    using (var lines = new StreamReader(cuename))
                    {
                        line = lines.ReadToEnd();
                        lines.Close();
                    }

                    char[] sep = { '\n' };
                    string[] cueLines = line.Split(sep);

                    for (int j = 0; j < cueLines.Length; j++)
                    {
                        if (cueLines[j].Length > 1)
                            cueLines[j] = cueLines[j].Remove(cueLines[j].Length - 1);
                        if (j == 4 || j == 5 || j == 6 || j == 7)
                            cueLines[j] = "";
                    }

                    using (var deletedLines = new StreamWriter(cuename))
                    {

                        foreach (var t in cueLines)
                        {
                            if (t != "")
                                deletedLines.WriteLine(t);
                        }

                        deletedLines.Close();
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void saveToPromodjCue(string cueName)
        {
            try
            {
                
                List<string> list = new List<string>();

                for (int i = 0; i < listFileNames.Count; i++)
                {
                    char[] c = { ':' };
                    string[] subs = listFileNames[i].Time.Split(c);
                    if (subs.Length > 2)
                    {
                        string min = subs[0];
                        string sec = subs[1];
                        list.Add(min + ":" + sec + " " + listFileNames[i].Performer + " - " + listFileNames[i].Title);
                    }
                    else
                    {
                        if (i != 0)
                            list.Add("00:00" + " " + listFileNames[i].Performer + " - " + listFileNames[i].Title);
                        else
                            list.Add(listFileNames[i].Performer + " - " + listFileNames[i].Title);
                    }   
                }

                File.WriteAllLines(cueName, list, Encoding.UTF8);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private int getMin(string s)
        {
            char[] c = {':'};
            string[] subs = s.Split(c);

            if (subs[0][0] == '0')
                return Convert.ToInt32(subs[0][1]);
            else
                return Convert.ToInt32(subs[0]);
        }

        private int getSec(string s)
        {
            char[] c = { ':' };
            string[] subs = s.Split(c);

            if (subs[1][0] == '0')
                return Convert.ToInt32(subs[1][1]);
            else
                return Convert.ToInt32(subs[1]);
        }

        private int getFrames(string s)
        {
            char[] c = { ':' };
            string[] subs = s.Split(c);

            if (subs[2][0] == '0')
                return Convert.ToInt32(subs[2][1]);
            else
                return Convert.ToInt32(subs[2]);
        }

        private string getPerformer(string s)
        {
            string str = "";
            int i = -1;

            if (s.IndexOf("–") != -1)
                s = s.Replace("–", "-");

            if (s.IndexOf("-") != -1)
                i = s.IndexOf("-");

            if (i != -1)
                str = s.Substring(0, i);

            if (isNumerated)
                str = str.Remove(0, numMask.SelectedIndex + 2);

            return str;
        }

        private string getTitle(string s)
        {
            string str = "";
            int i = -1;

            if (s.IndexOf("–") != -1)
                s = s.Replace("–", "-");

            if (s.IndexOf("-") != -1)
                i = s.IndexOf("-");

            if (i != -1)
                str = s.Substring(i + 1);
            else
                str = s;

            string mp3 = str.ToLower();

            if (mp3.IndexOf(".mp3") != -1)
                str = str.Remove(mp3.IndexOf(".mp3"));

            return str.TrimStart();
        }

        private void saveToCurrentTxt(string fn)
        {
            try
            {
                string filename = shortName(fn);
                //для дописывания в конец файла
                FileInfo fi = new FileInfo(filename);

                using (var outStream = fi.AppendText())
                {
                    outStream.WriteLine();
                    if (именоватьСписокtoolStripMenuItem.Checked)
                    {
                        outStream.WriteLine(listTitle.Text);
                        outStream.WriteLine("--------------------------------------------------------------------");
                    }

                    for (int i = 0; i < listBoxName.Items.Count; i++)
                        outStream.WriteLine(listBoxName.Items[i]);
                    outStream.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }

        }

        /// <summary>
        /// сохраняет только имя файла
        /// </summary>
        string shortName(string file)
        {
            char[] sep = { '\\' };
            string[] shortName = file.Split(sep);
            return shortName[shortName.Length - 1];
        }

        private void listBoxName_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBoxName.Items.Count != 0)
                {
                    if (listBoxName.SelectedIndex != -1)
                    {
                        textBoxChange.Text = "" + listBoxName.Items[listBoxName.SelectedIndex];
                    }
                    else
                    {
                        textBoxChange.Text = "" + listBoxName.Items[listBoxName.Items.Count - 1];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void открытьФайлыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyselfOpen();
            listCleared = false;
            _specCharsDeleted = false;
            isNumerated = false;
            _isRemove = false;
        }

        private void сохранитьВtxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MySave();
        }

        bool isRepeatAdd = false;
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            addSubs();
        }

        /// <summary>
        /// добавляет в начало и конец пункта списка подстроки
        /// </summary>
        void addSubs()
        {
            autosave();
            string s = "";
            string beg = begin.Text;
            string en = end.Text;

            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                s = listBoxName.Items[i].ToString();

                if (begin.Text.IndexOf("#Time") != -1)
                {
                    begin.Text = addTime(i) + " " + begin.Text.Replace("#Time", "");
                }

                if (end.Text.IndexOf("#Time") != -1)
                {
                    end.Text = end.Text.Replace("#Time", "") + " " + addTime(i);
                }
                s = begin.Text + s + end.Text;
                begin.Text = beg;
                end.Text = en;
                listBoxName.Items[i] = s;
            }
            
            isRepeatAdd = true;
        }

        private string addTime(int itemIndex)
        {
            string time = "";

            switch (timeFormat.Text)
            {
                case "mm:ss": {
                    if (listFileNames[itemIndex].Time != "")
                        time = listFileNames[itemIndex].Time.Remove(listFileNames[itemIndex].Time.Length - 3, listFileNames[itemIndex].Time.Length);
                    else
                        time = "00:00";
                }
                    break;
                case "mm:ss:msms": {
                    if (listFileNames[itemIndex].Time != "")
                        time = listFileNames[itemIndex].Time;
                    else
                        time = "00:00:00";
                }
                    break;
                default: time = "00:00:01";
                    break;
            }

            return time;
        }

        private void заменитьПодчеркиваниеНаПробелToolStripMenuItem_Click(object sender, EventArgs e)
        {
            replace_ToSpase();
        }

        //заменяет подчеркивание на пробел во всех строках списка
        private void replace_ToSpase()
        {
            autosave();

            string s = "";
            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                s = listBoxName.Items[i].ToString();
                s = s.Replace("_", " ");
                listBoxName.Items[i] = s;
            }
            _isRemove = true;
        }

        private void listBoxName_KeyDown(object sender, KeyEventArgs e)
        {
            if (listBoxName.Items.Count != 0)
            {
                if (e.KeyCode == Keys.Delete)
                {
                    myDelete();
                }
            }
        }

        private void undo_Click(object sender, EventArgs e)
        {
            abort();
        }

        private void abort()
        {
            try
            {

                if (!listCleared && !_specCharsDeleted && !isNumerated && !_isRemove && !foldersNameInsert && !isDeleteNumeration && !isRepeatAdd)
                {
                    if ((_oldDeleteValueIndex.Count != 0 && _oldDeleteValue.Count != 0))
                    {
                        listBoxName.Items.Insert(_oldDeleteValueIndex[_oldDeleteValueIndex.Count - 1], _oldDeleteValue[_oldDeleteValue.Count - 1]);
                        _oldDeleteValueIndex.RemoveAt(_oldDeleteValueIndex.Count - 1);
                        _oldDeleteValue.RemoveAt(_oldDeleteValue.Count - 1);
                    }
                }
                else if (foldersNameInsert)
                {
                    listBoxName.Items.RemoveAt(0);
                    //listFileNames.RemoveAt(0);
                    foldersNameInsert = false;
                }
                else
                {
                    listBoxName.Items.Clear();
                    //listFileNames.Clear();
                    while (_oldDeleteValue.Count != 0)
                        if ((_oldDeleteValueIndex.Count != 0 && _oldDeleteValue.Count != 0))
                        {
                            listBoxName.Items.Insert(_oldDeleteValueIndex[_oldDeleteValueIndex.Count - 1], _oldDeleteValue[_oldDeleteValue.Count - 1]);
                            //FileName mfn = new FileName();
                            //mfn.CurrentFileName = _oldDeleteValue[_oldDeleteValue.Count - 1];
                            //mfn.OldFileName = listFileNames[_oldDeleteValueIndex[_oldDeleteValueIndex.Count - 1]].OldFileName;
                            //listFileNames.Insert(_oldDeleteValueIndex[_oldDeleteValueIndex.Count - 1], mfn);
                            _oldDeleteValueIndex.RemoveAt(_oldDeleteValueIndex.Count - 1);
                            _oldDeleteValue.RemoveAt(_oldDeleteValue.Count - 1);

                        }
                    listCleared = _specCharsDeleted = isNumerated = _isRemove = isDeleteNumeration = isRepeatAdd = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void autosave()
        {
            _oldDeleteValue.Clear();
            _oldDeleteValueIndex.Clear();
            for (int i = listBoxName.Items.Count - 1; i >= 0; i--)
            {
                _oldDeleteValue.Add(listBoxName.Items[i].ToString());
                _oldDeleteValueIndex.Add(i);
            }
        }

        private void разработчикToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Developer d = new Developer();
            d.Text = Text + " © " + d.Text;
            d.ShowDialog();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            deleteSpecSymb();
        }

        private bool _specCharsDeleted = false;
        /// <summary>
        /// удаляет специальные символы из всех строк списка
        /// </summary>
        void deleteSpecSymb()
        {
            autosave();
            string str = textBoxDeleteChars.Text;
            string s = "";
            char[] chars = new char[str.Length];

            for (int i = 0; i < str.Length; i++)
                chars[i] = str[i];

            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                s = listBoxName.Items[i].ToString();

                while (s.IndexOfAny(chars) != -1)
                {
                    s = s.Remove(s.IndexOfAny(chars), 1);
                    listBoxName.Items[i] = s;
                    listFileNames[i].CurrentFileName = s;
                }
            }
            _specCharsDeleted = true;
        }


        private void очиститьСписокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearList();
        }

        bool listCleared = false;
        private void clearList()
        {
            autosave();
            listBoxName.Items.Clear();
            listFileNames.Clear();
            listCleared = true;
        }


        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            numeration();
        }

        //для отмены нумерации
        bool isNumerated = false;
        //нумерует список
        private void numeration()
        {
            autosave();
            string sep = numSeparator.Text;

            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                //сначала просто нумеруем все пункты списка
                listBoxName.Items[i] = (i + 1).ToString() + "$$$$" + sep + listBoxName.Items[i];
            }

            int idx = listBoxName.Items[listBoxName.Items.Count - 1].ToString().IndexOf("$$$$");
            //получаем количество цифр последнего числа нумерации списка
            int kol = listBoxName.Items[listBoxName.Items.Count - 1].ToString().Remove(idx).Length;
            numMask.Text = numMask.Items[kol - 1].ToString();

            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                listBoxName.Items[i] = listBoxName.Items[i].ToString().Replace("$$$$", "");
                switch (kol)
                {
                    case 1:
                        {
                            if (i < 9)
                            {
                                listBoxName.Items[i] = listBoxName.Items[i];
                            }
                        }
                        break;
                    case 2:
                        {
                            if (i < 9)
                            {
                                listBoxName.Items[i] = "0" + listBoxName.Items[i];
                            }
                            else if (i >= 9 && i < 99)
                            {
                                listBoxName.Items[i] = listBoxName.Items[i];
                            }
                        }
                        break;
                    case 3:
                        {
                            if (i < 9)
                            {
                                listBoxName.Items[i] = "00" + listBoxName.Items[i];
                            }
                            else if (i >= 9 && i < 99)
                            {
                                listBoxName.Items[i] = "0" + listBoxName.Items[i];
                            }
                            else if (i >= 99 && i < 999)
                            {
                                listBoxName.Items[i] = listBoxName.Items[i];
                            }
                        }
                        break;
                    case 4:
                        {
                            if (i < 9)
                            {
                                listBoxName.Items[i] = "000" + listBoxName.Items[i];
                            }
                            else if (i >= 9 && i < 99)
                            {
                                listBoxName.Items[i] = "00" + listBoxName.Items[i];
                            }
                            else if (i >= 99 && i < 999)
                            {
                                listBoxName.Items[i] = "0" + listBoxName.Items[i];
                            }
                            else if (i >= 999 && i < 9999)
                            {
                                listBoxName.Items[i] = listBoxName.Items[i];
                            }
                        }
                        break;

                }


            }

            isNumerated = true;
        }

        bool isDeleteNumeration = false;
        private void удалитьНумерациюИзСпискаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteNumeration();
        }

        private void deleteNumeration()
        {
            autosave();
            char[] c = { '|' };
            string[] s = numMask.Text.Split(c);
            int kol = s[0].Length;
            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                if (listBoxName.Items[i].ToString() != "" && listBoxName.Items[i].ToString().Length > kol)
                    listBoxName.Items[i] = listBoxName.Items[i].ToString().Remove(0, kol);
            }
            isDeleteNumeration = true;
            isNumerated = false;
        }

        //задает маску для удаления нумерации списка
        private void setMask()
        {
            string[] mask = { "0.|0)|0-", "00.|00)|00-", "000.|000)|000-", "0000.|0000)|0000-" };
            foreach (string s in mask)
                numMask.Items.Add(s);
            numMask.Text = numMask.Items[1].ToString();
        }
        //задает формат вывода времени для треков
        private void setTimeFormat()
        {
            string[] format = { "mm:ss", "mm:ss:msms"};
            foreach (string s in format)
                timeFormat.Items.Add(s);
            timeFormat.Text = timeFormat.Items[0].ToString();
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            replaseSubs();
        }

        private bool _isRemove = false;
        //заменяет одну подстроку на другую
        private void replaseSubs()
        {
            autosave();
            string s = "";
            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                s = listBoxName.Items[i].ToString();
                s = s.Replace(subs1.Text, subs2.Text);
                listBoxName.Items[i] = s;
            }
            _isRemove = true;
        }

        private bool foldersNameInsert = false;

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            myToUpper();
        }

        /// <summary>
        /// заменяет начальные буквы слов на те же в верхнем регистре
        /// </summary>
        private void myToUpper()
        {
            string s = "";
            for (int i = 0; i < listBoxName.Items.Count; i++)
            {
                s = listBoxName.Items[i].ToString();
                s = AllToUpRegister(s);
                listBoxName.Items.RemoveAt(i);
                listBoxName.Items.Insert(i, s);
            }
        }

        /// <summary>
        /// заменяет все первые буквы слов в строке на верхний регистр 
        /// </summary>
        private string AllToUpRegister(string s)
        {
            char[] sep = { ' ', '_', '.', '(', '-' };
            string[] str = s.Split(sep);

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] != "")
                {
                    string firstChar = str[i][0].ToString();
                    string other = str[i].Substring(1);
                    firstChar = firstChar.ToUpper();
                    str[i] = firstChar + other;
                }
            }

            string result = s.ToLower();

            foreach (string word in str)
            {
                string w = word.ToLower();

                if (word != "")
                    if (result.ToLower().IndexOf(w) != -1 && w.Length != 1)
                        result = result.Replace(w, word);
                    else if (w.Length == 1)
                    {
                        result = result.Replace(" " + w, " " + word);
                        result = result.Replace("(" + w, "(" + word);
                        result = result.Replace("-" + w, "-" + word);
                        result = result.Replace("_" + w, "_" + word);
                    }

            }

            return result;
        }

        private void сохранитьВИсходникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveToFirst();
        }

        /// <summary>
        /// сохраняет в исходники
        /// </summary>
        private void saveToFirst()
        {
            try
            {

                for (int i = 0; i < listFileNames.Count; i++)
                {
                    if (listFileNames[i].CurrentFileName != listFileNames[i].OldFileName || listFileNames[i].CurrentFileName != "#_foldersName#")
                    {
                        listFileNames[i].CurrentFileName = listBoxName.Items[i].ToString();
                    }
                }

                foreach (FileName mfn in listFileNames)
                {
                    FileInfo fi = new FileInfo(mfn.OldFileName);

                    if (mfn.CurrentFileName != mfn.OldFileName || mfn.CurrentFileName != "#_foldersName#")
                    {
                        var audioFile = TagLib.File.Create(mfn.OldFileName);
                        char[] badChars = { '|', '\\', '/', ':', '*', '<', '>', '?', '\"' };
                        string file = "";
                        char[] sep = { '\\' };
                        string[] words = mfn.OldFileName.Split(sep);

                        foreach (char c in badChars)
                        {
                            mfn.CurrentFileName.Replace(c, ' ');
                        }
                        words[words.Length - 1] = mfn.CurrentFileName;

                        for (int i = 0; i < words.Length; i++)
                        {
                            if (i != words.Length - 1)
                                file += words[i] + "\\";
                            else
                                file += words[i];
                        }

                        audioFile.Tag.Album = mfn.Album;
                        string[] sep2 = { ", " };
                        audioFile.Tag.Performers = mfn.Performer.Split(sep2, 10, StringSplitOptions.None);
                        audioFile.Tag.Genres = mfn.Genres.Split(sep2, 10, StringSplitOptions.None);
                        audioFile.Tag.Title = mfn.Title;
                        audioFile.Save();
                        fi.MoveTo(file);
                        mfn.OldFileName = file;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void toolsOformList_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem)
                ((ToolStripMenuItem)sender).Checked = !((ToolStripMenuItem)sender).Checked;
        }

        private void textBoxChange_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyCode == Keys.Enter)
            {
                setAllFileAtributes();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                textBoxChange.ReadOnly = true;

                for (int i = 0; i < menuStrip1.Items.Count - 1; i++)
                {
                    menuStrip1.Items[i].Enabled = true;
                }

                panelCue.Visible = false;
                listBoxName.Focus();
            }
        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chngeItem();
        }

        private bool _changeItem = false;
        private int _selectedIdx = 0;
        //редактирует пункт списка
        //использовать библиотеку  Taglib-sharp
        private void chngeItem()
        {
            try
            {
                for (int i = 0; i < menuStrip1.Items.Count - 1; i++)
                {
                    menuStrip1.Items[i].Enabled = false;
                }
                _selectedIdx = listBoxName.SelectedIndex;
                textBoxChange.ReadOnly = false;

                if (listBoxName.SelectedIndex != -1)
                {
                    textBoxChange.Text = listBoxName.Items[listBoxName.SelectedIndex].ToString();
                    panelCue.Visible = true;

                        FileInfo fi = new FileInfo(listFileNames[listBoxName.SelectedIndex].CurrentFileName);
                        if (!fi.Exists)
                        {
                            var audioFile = TagLib.File.Create(listFileNames[listBoxName.SelectedIndex].OldFileName);
                            if (audioFile.Tag.Artists.Length != 0)
                                textBoxPerformer.Text = string.Join(", ", audioFile.Tag.Artists);
                            else
                                textBoxPerformer.Text = getPerformer(textBoxChange.Text);

                            if (!audioFile.Tag.Title.Equals(""))
                                textBoxTitle.Text = audioFile.Tag.Title;
                            else
                                textBoxTitle.Text = getTitle(textBoxChange.Text);

                            if (audioFile.Tag.Genres.Length != 0)
                                textBoxGenre.Text = string.Join(", ", audioFile.Tag.Genres);
                            else
                                textBoxGenre.Text = "Hard";

                            textBoxTime.Text = "00:00:00";

                            if (!audioFile.Tag.Album.Equals(""))
                                textBoxAlbum.Text = audioFile.Tag.Album;
                            else
                                textBoxAlbum.Text = "hard mixes";
                        }
                        else
                        {
                            if (textBoxPerformer.Text == "")
                                textBoxPerformer.Text = getPerformer(textBoxChange.Text);
                            else
                            {
                                if (listFileNames[listBoxName.SelectedIndex].Performer != "")
                                    textBoxPerformer.Text = listFileNames[listBoxName.SelectedIndex].Performer;
                                else
                                {
                                    textBoxPerformer.Text = getPerformer(textBoxChange.Text);
                                }
                            }
                            if (textBoxTitle.Text == "")
                                textBoxTitle.Text = getTitle(textBoxChange.Text);
                            else
                            {
                                if (listFileNames[listBoxName.SelectedIndex].Title != "")
                                    textBoxTitle.Text = listFileNames[listBoxName.SelectedIndex].Title;
                                else
                                {
                                    textBoxTitle.Text = getTitle(textBoxChange.Text);
                                }
                            }
                            if (textBoxTime.Text == "")
                                textBoxTime.Text = "00:00:00";
                            else
                            {
                                if (listFileNames[listBoxName.SelectedIndex].Time != "")
                                    textBoxTime.Text = listFileNames[listBoxName.SelectedIndex].Time;
                                else
                                {
                                    textBoxTime.Text = "00:00:00";
                                }
                            }
                        }
                }

                textBoxChange.Focus();
                _changeItem = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void numMask_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void test()
        {
            for (int i = 0; i < 9999; i++)
            {
                var filename = randomstring(i) + ".txt";
                FileInfo fi = new FileInfo("C:\\Users\\Николай\\Desktop\\" + "test\\" + filename);
                StreamWriter sw = new StreamWriter(fi.FullName);
                sw.Write(i.ToString());
                sw.Close();
            }
        }

        private string randomstring(int j)
        {
            string[] array = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "!", "-", "_", 
                             "q", "w", "e", "r", "t", "y", "u", "i", "o", "p", "a", "s", "d", "f", "g", "h", "j", "k", "l", 
                             "z", "x", "c", "v", "b", "n", "m", "~"};
            Random r = new Random();
            string s = "";

            for (int i = 0; i < 50; i++)
            {
                s += array[r.Next(0, array.Length)];
            }

            return s + j;
        }

        private void тестированиеНумерацииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            test();
        }

        private void cut_Click(object sender, EventArgs e)
        {
            myCut();
            setLB();
        }

        private void myCut()
        {
            try
            {
                if (listBoxName.SelectedIndex != -1 && !panelCue.Visible)
                {
                    int index = listBoxName.SelectedIndex;
                    _oldDeleteValue.Add(listBoxName.Items[index].ToString());
                    _oldDeleteValueIndex.Add(index);
                    _clipboard.CurrentFileName = listFileNames[index].CurrentFileName;
                    _clipboard.OldFileName = listFileNames[index].OldFileName;
                    listBoxName.Items.RemoveAt(index);
                    listFileNames.RemoveAt(index);
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }

        }

        private void copy_Click(object sender, EventArgs e)
        {
            myCopy();
            setLB();
        }

        private void myCopy()
        {
            try
            {
                if (listBoxName.SelectedIndex != -1 && !panelCue.Visible)
                {
                    _clipboard.CurrentFileName = listFileNames[listBoxName.SelectedIndex].CurrentFileName;
                    _clipboard.OldFileName = listFileNames[listBoxName.SelectedIndex].OldFileName;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void insert_Click(object sender, EventArgs e)
        {
            myInsert();
            setLB();
        }

        private void myInsert()
        {
            try 
            {
                if (listBoxName.SelectedIndex != -1 && !panelCue.Visible)
                {
                    if (_clipboard.CurrentFileName != "" && _clipboard.OldFileName != "")
                    {
                        int index = listBoxName.SelectedIndex;
                        listBoxName.Items.Insert(index, _clipboard.CurrentFileName);
                        FileName mfn = new FileName();
                        mfn.CurrentFileName = _clipboard.CurrentFileName;
                        mfn.OldFileName = _clipboard.OldFileName;
                        listFileNames.Insert(index, mfn);
                        //_clipboard.CurrentFileName = "";
                        //_clipboard.OldFileName = "";
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void delete_Click(object sender, EventArgs e)
        {
            myDelete();
            setLB();
        }

        private void myDelete()
        {
            try
            {
                if (listBoxName.SelectedIndex != -1 && !panelCue.Visible)
                {
                    int index = listBoxName.SelectedIndex;
                    _oldDeleteValue.Add(listBoxName.Items[index].ToString());
                    _oldDeleteValueIndex.Add(index);
                    listBoxName.Items.RemoveAt(index);
                    listFileNames.RemoveAt(index);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void программныйБуферОбменаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Имя файла: " + _clipboard.CurrentFileName +"\n" + "Полное имя файла: " + _clipboard.OldFileName, "Программный буфер обмена");
        }

        void setLB()
        {
            listBox1.Items.Clear();
            foreach (FileName mfn in listFileNames)
            {
                listBox1.Items.Add(mfn.OldFileName);
            }
        }

        //для перетаскиваемых объектов
        private object _objForDrop = new object();
        private void listBoxName_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
            string[] formats = e.Data.GetFormats();
            string type = e.Data.GetType().Name;
            _objForDrop = e.Data.GetData("FileDrop");//DataFormats.UnicodeText
            
        }

        private void listBoxName_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                allFileNames = new List<string>();
                for (int i = 0; i < ((string[])_objForDrop).Length; i++)
                {
                    getDir(((string[])_objForDrop)[i]);
                }
                foreach (string s in allFileNames)
                    addToList(s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        List<string> allFileNames;
        void getDir(string path)
        {
            if (path[path.Length - 4] != '.')
            {
                DirectoryInfo parentDirectory = new DirectoryInfo(path);
                if (parentDirectory.GetFiles().Length != 0)
                    foreach (FileInfo s in parentDirectory.GetFiles())
                        allFileNames.Add(s.FullName);
                foreach (DirectoryInfo dir in parentDirectory.GetDirectories())
                {
                    getDir(dir.FullName);
                }
            }
            else
            {
                allFileNames.Add(path);
            }
            
            
        }

        private void addToList(string path)
        {
            FileName mfn = new FileName();
            mfn.OldFileName = path;
            string sfn = shortName(mfn.OldFileName);
            if (toolsOformList.Checked)
            {
                listBoxName.Items.Add(begin.Text + sfn + end.Text);
                mfn.CurrentFileName = begin.Text + sfn + end.Text;
                _defaultView.Add(begin.Text + sfn + end.Text);
            }
            else
            {
                listBoxName.Items.Add(sfn);
                mfn.CurrentFileName = sfn;
                _defaultView.Add(sfn);
            }
            mfn.CurrentFileName = shortName(path);


            listFileNames.Add(mfn);
        }

        private void change_Click(object sender, EventArgs e)
        {
            chngeItem();
        }

        //перетаскивание мышью
        private FileName _mouseDragItem = new FileName();
        private bool _isMouseDown = false;
        private int _mouseDragIndex = -1;

        /// <summary>
        /// вырезание перетаскиваемого объекта
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxName_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                _isMouseDown = true;
                _mouseDragIndex = listBoxName.SelectedIndex;
                if (_mouseDragIndex != -1)
                {
                    if (_mouseDragIndex < listFileNames.Count)
                    {
                        _mouseDragItem.CurrentFileName = listFileNames[_mouseDragIndex].CurrentFileName;
                        _mouseDragItem.OldFileName = listFileNames[_mouseDragIndex].OldFileName;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }

        }

        private int _oldMouseDragIndex = -2;
        /// <summary>
        /// вставка перетаскиваемого объекта
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxName_MouseUp(object sender, MouseEventArgs e)
        {
            if (_isMouseDown)
            {
                if (_mouseDragIndex != -1)
                {
                    if (_mouseDragItem.CurrentFileName != "" && _mouseDragItem.OldFileName != "")
                    {
                        if (_oldMouseDragIndex != _mouseDragIndex)
                        {
                            
                            listBoxName.Items.RemoveAt(_oldMouseDragIndex);
                            listFileNames.RemoveAt(_oldMouseDragIndex);
                            listBoxName.Items.Insert(_mouseDragIndex, _mouseDragItem.CurrentFileName);
                            FileName mfn = new FileName();
                            mfn.CurrentFileName = _mouseDragItem.CurrentFileName;
                            mfn.OldFileName = _mouseDragItem.OldFileName;
                            listFileNames.Insert(_mouseDragIndex, mfn);
                        }
                    }
                    _oldMouseDragIndex = _mouseDragIndex;
                }
                _isMouseDown = false;
            }
            setLB();
        }

        private void listBoxName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isMouseDown)
            {
                _oldMouseDragIndex = _mouseDragIndex;
                _mouseDragIndex = listBoxName.SelectedIndex;
                label1.Text = "_oldMouseDragIndex = " + _oldMouseDragIndex.ToString();
                label2.Text = "_mouseDragIndex = " + _mouseDragIndex.ToString();
            }
        }
        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Продолжить? Изменения сохранены не будут!", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                listBoxName.Items.Clear();

                foreach (string s in _defaultView)
                    listBoxName.Items.Add(s);
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            setAllFileAtributes();
        }

        /// <summary>
        /// задает все поля для файла
        /// </summary>
        void setAllFileAtributes()
        {
            try
            {
                if (isTimeOk(textBoxTime.Text))
                {
                    if (listBoxName.SelectedIndex != -1)
                    {
                        listFileNames[listBoxName.SelectedIndex].Performer = textBoxPerformer.Text;
                        listFileNames[listBoxName.SelectedIndex].Title = textBoxTitle.Text;
                        listFileNames[listBoxName.SelectedIndex].Album = textBoxAlbum.Text;
                        listFileNames[listBoxName.SelectedIndex].Time = textBoxTime.Text;
                        listFileNames[listBoxName.SelectedIndex].Genres = textBoxGenre.Text;
                        listFileNames[listBoxName.SelectedIndex].CurrentFileName = textBoxChange.Text;
                        listBoxName.Items[listBoxName.SelectedIndex] = textBoxChange.Text;

                        
                    }
                    else
                    {
                        listBoxName.Items.Add(textBoxChange.Text);
                        listFileNames[listBoxName.Items.Count - 1].Performer = textBoxPerformer.Text;
                        listFileNames[listBoxName.Items.Count - 1].Title = textBoxTitle.Text;
                        listFileNames[listBoxName.SelectedIndex].Album = textBoxAlbum.Text;
                        listFileNames[listBoxName.SelectedIndex].Time = textBoxTime.Text;
                        listFileNames[listBoxName.SelectedIndex].Genres = textBoxGenre.Text;
                    }

                    textBoxChange.ReadOnly = true;
                    for (int i = 0; i < menuStrip1.Items.Count - 1; i++)
                    {
                        menuStrip1.Items[i].Enabled = true;
                    }

                    panelCue.Visible = false;
                    labelToolTip.Visible = false;
                    listBoxName.Focus();

                    if (listBoxName.SelectedIndex != -1 && listBoxName.SelectedIndex + 1 < listBoxName.Items.Count)
                        listBoxName.SelectedIndex++;
                }
                else
                {
                    labelToolTip.Visible = true;
                    labelToolTip.Text = "Формат ввода времени \"00:00:00\"";
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void panelCue_VisibleChanged(object sender, EventArgs e)
        {
            if (panelCue.Visible)
                listBoxName.Enabled = false;
            else
                listBoxName.Enabled = true;
        }

        private bool _deleteKeyPressed;
        private void textBoxPerformer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (sender is TextBox)
                {
                    if (((TextBox)sender).Name == "textBoxPerformer")
                    {
                        textBoxTitle.Focus();
                    }
                    else if (((TextBox)sender).Name == "textBoxTitle")
                    {
                        textBoxTime.Focus();
                    }
                    else if (((TextBox)sender).Name == "textBoxTime")
                    {
                        string s = ((TextBox)sender).Text;

                        if (isTimeOk(s))
                            setAllFileAtributes();
                        else
                        {
                            labelToolTip.Visible = true;
                            labelToolTip.Location = new Point(((TextBox)sender).Location.X, ((TextBox)sender).Location.Y + 20);
                            labelToolTip.Size = new Size(((TextBox)sender).Width, ((TextBox)sender).Height);
                            labelToolTip.Text = "Формат ввода времени \"00:00:00\"";
                        }
                    }
                }
            }
            if (e.KeyCode == Keys.Back)
            {
                _deleteKeyPressed = true;
            }
            else
            {
                _deleteKeyPressed = false;
            }
        }

        private bool isTimeOk(string s)
        {
            if (s.Length == 8)
            {
                if (s.IndexOf(":") == 2)
                {
                    s = s.Remove(2, 1);
                    if (s.IndexOf(":") == 4)
                    {
                        s = s.Remove(4, 1);
                        if (s.IndexOf(":") == -1)
                        {
                            int k = 0;

                            foreach (char c in s)
                            {
                                if ("0123456789".IndexOf(c) != -1)
                                    k++;
                            }

                            if (k == s.Length)
                                return true;
                        }
                    }
                }
            }

            return false;
        }

        private void textBoxPerformer_MouseHover(object sender, EventArgs e)
        {
            if (sender is TextBox)
            {
                labelToolTip.Visible = true;
                labelToolTip.Location = new Point(_mouseX, _mouseY);
                labelToolTip.Size = new Size(((TextBox)sender).Width, ((TextBox)sender).Height);
                labelToolTip.Text = ((TextBox)sender).Text;
            }
            else if (sender is Label)
            {
                labelToolTip.Visible = true;
                labelToolTip.Location = new Point(_mouseX, _mouseY);
                labelToolTip.Size = new Size(((Label)sender).Width, ((Label)sender).Height);
                labelToolTip.Text = ((Label)sender).Text;
            }
        }

        private void buttonCansel_Click(object sender, EventArgs e)
        {
            panelCue.Visible = false;
            textBoxChange.ReadOnly = true;

            for (int i = 0; i < menuStrip1.Items.Count - 1; i++)
            {
                menuStrip1.Items[i].Enabled = true;
            }
        }

        private void panelCue_MouseMove(object sender, MouseEventArgs e)
        {
            _mouseX = e.X;
            _mouseY = e.Y;
        }

        private void textBoxPerformer_MouseLeave(object sender, EventArgs e)
        {
            if (sender is TextBox)
            {
                labelToolTip.Visible = false;
            }
            else if (sender is Label)
            {
                labelToolTip.Visible = true;
            }
        }

        private void panelCue_Enter(object sender, EventArgs e)
        {
            if (panelCue.Visible)
            {
                textBoxPerformer.Focus();
            }
        }

        private void информацияОФайлеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myFileInfo();
        }

        /// <summary>
        /// показывает информацию о файле, использовать библиотеку  Taglib-sharp
        /// </summary>
        void myFileInfo()
        {
            try
            {
                if (listBoxName.SelectedIndex != -1)
                {
                    if (listFileNames[listBoxName.SelectedIndex].CurrentFileName != "")
                    {
                        FileInfo fi = new FileInfo(listFileNames[listBoxName.SelectedIndex].CurrentFileName);

                        if (!fi.Exists)
                        {
                            var audioFile = TagLib.File.Create(listFileNames[listBoxName.SelectedIndex].OldFileName);
                            MessageBox.Show(string.Format("Полное имя: {0}\nИсполнитель: {1}\nАльбом: {2}\nНаименование: {3}\nЖанр: {4}\nВремя начала: {5}\nТекущее имя: {6}",
                                listFileNames[listBoxName.SelectedIndex].OldFileName, string.Join(", ", audioFile.Tag.Artists), audioFile.Tag.Album, audioFile.Tag.Title,
                                string.Join(", ", audioFile.Tag.Genres), listFileNames[listBoxName.SelectedIndex].Time, listFileNames[listBoxName.SelectedIndex].CurrentFileName));
                        }
                    }
                    else
                        MessageBox.Show("Полное имя: " + "\n" +
                                        "Исполнитель: " + listFileNames[listBoxName.SelectedIndex].Performer + "\n" +
                                        "Наименование: " + listFileNames[listBoxName.SelectedIndex].Title + "\n" +
                                        "Время начала: " + listFileNames[listBoxName.SelectedIndex].Time,
                                        "Информация о файле \"" + listFileNames[listBoxName.SelectedIndex].CurrentFileName + "\"");
                }
                else
                    MessageBox.Show("Выберите файл из списка.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void getTextBoxFocus(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem)
            {
                switch (((ToolStripMenuItem)sender).Name)
                {
                    case "добавитьВНачалоToolStripMenuItem": { begin.Focus(); }
                        break;
                    case "добавитьВКонецToolStripMenuItem": { end.Focus(); }
                        break;
                    case "toolStripMenuItemFolderNameAdd": { listTitle.Focus(); }
                        break;
                    case "заменяемаяПодстрокаToolStripMenuItem": { subs1.Focus(); }
                        break;
                    case "заменяющаяПодстрокаToolStripMenuItem": { subs2.Focus(); }
                        break;
                    case "переченьСпециальныхСимволовToolStripMenuItem": { textBoxDeleteChars.Focus(); }
                        break;
                    case "маскаДляНумерацииToolStripMenuItem": { numMask.Focus(); }
                        break;
                    case "разделительНумерацииToolStripMenuItem": { numSeparator.Focus(); }
                        break;
                    case "исполнительToolStripMenuItem": { toolStripTextBoxPerformer.Focus(); }
                        break;
                    case "наименованиеToolStripMenuItem": { toolStripTextBoxTitle.Focus(); }
                        break;
                    case "файлToolStripMenuItem": { toolStripTextBoxFile.Focus(); }
                        break;

                }
            }
        }

        private bool _wasDeletePressed = false;
        private void textBoxTime_TextChanged(object sender, EventArgs e)
        {
            if (!_deleteKeyPressed)
            {
                if (textBoxTime.Text.Length == 2 || textBoxTime.Text.Length == 5)
                {
                    textBoxTime.Text += ":";
                    _deleteKeyPressed = false;
                }
                if (textBoxTime.Text.Length == 3 && textBoxTime.Text[2] != ':' && _wasDeletePressed ||
                    textBoxTime.Text.Length == 6 && textBoxTime.Text[5] != ':' && _wasDeletePressed)
                {
                    textBoxTime.Text = textBoxTime.Text.Substring(0, textBoxTime.Text.Length - 1) + ":" + textBoxTime.Text[textBoxTime.Text.Length - 1];
                    _wasDeletePressed = false; 
                }
            }
            else
                _wasDeletePressed = true;
            textBoxTime.SelectionStart = textBoxTime.Text.Length;
        }

        private void textBoxTime_Click(object sender, EventArgs e)
        {
            if (sender is TextBox)
            {
                ((TextBox)sender).SelectAll();
            }
        }

        private Word.Application ThisApplication = new Word.Application();
        private Word.Document doc = null;
        private void manual_Click(object sender, EventArgs e)
        {
            try
            {
                string filename = System.Windows.Forms.Application.StartupPath + "\\manual.docx";
                string[] content = {
                                       /*"для тестирования",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<c><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<l><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",
                                       "<r><b>q</b>w<i>e</i>r<b>t</b>y<i>u</i>i<b>o</b>p<i>[</i>]<b>a</b>s<i>d</i>f<b>g</b>h<i>j</i>k<b>l</b>a<i>s</i>",*/
                                   "Приложение <b>" + Text + "</b> представляет собой редактор наименований файлов, предназначенный  для быстрого создания плейлистов и cue-файлов.",
                                   "<c><b>Руководство использования приложения " + Text +"</b>",
                                   "Открыть файл(-ы) можно в меню <i>\"Файл - Открыть\"</i> (горячие клавиши Ctrl+O) или при помощи мыши перетащить в окно приложения необходимые файлы.",
                                   "<b><c>Имеется несколько видов сохранения файлов:</b>",
                                       "<l>1. Обычное сохранение. Пользователю предлагается выбрать необходимое расширение из \".txt\", \".m3u\", \".cue\". ",
                                       "<l>Сохраняя список в формате \".txt\" (иными словами в обычный текстовый файл), пользователь получит список всех загруженных и отредактированных в приложении строк в сохраненном файле.",
                                       "<l>В случае сохранения в формате \".m3u\" (плейлист) пользователь сгенерирует плейлист из всех сохраненных файлов. Следует отметить, что если имена файлов редактировались, необходимо их <i>\"Сохранить в исходники\"</i> (об этом ниже), иначе плейлист сгенерируется с ошибкой (Точнее просто будет пустым). ",
                                       "<l>Последний предложенный вариант сохранения - это формат \".cue\", который используется для создания плейлиста непрерывной аудио-дрожки.", 
                                       "<b><c>Чтобы cue-файл сгенерировался корректно необходимо учесть следующие факторы:</b>",
                                       "<l>а) Сохранять cue-файл только в папку с аудио-файлом, для которого он создается;",
                                       "<l>б) Выбрать айдио-файл, для которого создается cue-файл;",
                                       "<l>в) Указать время начала каждой композиции (об этом ниже)",
                                       "г) Отредактировать названия исполнителя и композиции. В приложение встроен автоматический поиск этих названий, но, к сожалению, из-за нестандартизованного именования файлов оно может допустить ошибки. Для безошибочной работы необходимо именовать файлы следующим образом: <i>\"исполнитель - наименование композиции\"</i>.",
                                        "<l>2. Сохранение в существующий файл. Особенности: 1) формат файла должен быть \".txt\" и 2) сохраняемое содержимое будет добавлено в конец выбранного файла.",
                                        "<l>3. Сохранение в исходники. В данном случае никакого выходного файла создано не будет, а все отредактированные наименования сохранятся в наименования файлов, откуда были взяты (удобное средство для переименовывания большого количества файлов).",
                                    "<b><c>Редактирование имён загруженных файлов - основная задача приложения "+ Text + ".</b>",
                                        "<c>Все загруженное содержимое приложения составляет список файловых имён, поэтому редактирование разделено на редактирование пунктов списка и редактирование списка целиком.",
                                        "<l>К редактированию пунктов списка относятся команда <i>\"Правка - Изменить\"</i> (горячие клавиши Alt+C). В случае нажатия на неё пользователь получает возможность изменения как названия пункта целиком, так и специальных полей, необходимых для создания cue-файла (поля \"Исполнитель\", \"Наименование\", \"Начало трека\"). Обязательным для заполнения является поле \"Начало трека\", где пользователь в формате \"ММ:СС:МСМС\" должен записать время начала композиции. Для сохранения внесённых изменений пункта списка нужно нажать появившуюся вместе с окном редактирования кнопку \"Сохранить\" или нажать клавишу \"Enter\" (в зависимости от поля, в котором находится курсор ввода пользователя возможно различное поведение: при нахождении курсора в поле измненеия текста пункта списка - произойдет сохранение, в иных случаях - переходы по окнам.)",
                                        "<l>К редактированию списка целиком относятся все остальные команды в пунктах меню \"Правка\", \"Дополнительно\" и \"Список\".",
                                        "<l>В пункте меню <i>\"Правка\"</i> находятся команды, позволяющие менять местами пункты меню (<i>\"Вырезать\"</i>, <i>\"Копировать\"</i>, <i>\"Вставить\"</i>, <i>\"Удалить\"</i>). Также возможно перетаскивание пунктов меню по списку.",
                                        "<b><c>Значение команд в пункте меню <i>\"Дополнительно\"</i> трактуется следующим образом:</b>",
                                            "<l><i>\"Добавить при заполнении списка в начало\"</i>, <i>\"Добавить при заполнении списка в конец\"</i> - содержат поля для ввода строк, которые будут добавлены в начало и конец каждого пункта списка по нажатию на пункт <i>\"Добавить указанные строки\"</i>. В пункте меню <i>\"Настройки\"</i> имеется опция <i>\"Включить предоформление списка\"</i>, которая добавляет упомянутые строки сразу при загрузке имён файлов. По умолчанию она выключена.",
                                            "<l><i>\"Удалить специальные символы\"</i> - удаляет из пунктов символы, указанные в пункте меню <i>\"Настройки - Перечень специальных символов\"</i>.",
                                            "<l><i>\"Наименование списка\"</i> - содержит поле, в котором можно задать строку, которой будет озаглавливаться список при сохранении. Включение и выключение этой опции регулируется пунктом <i>\"Настройки - Именовать список\"</i>.",
                                            "<l><i>\"Заменить подчеркивание на пробел\"</i> (горячие клавиши Ctrl+Space) - заменяет каждое подчеркивание (\"_\") на пробел.",
                                            "<l><i>\"Заменить все начальные буквы слов заглавными\"</i> (горячие клавиши Ctrl+U) - заменяет каждую первую букву каждого слова на то же значение, но в верхнем регистре.",
                                            "<l><i>\"Заменить подстроки на другие подстроки\"</i> (горячие клавиши Ctrl+R) - заменяет каждое вхождение подстроки, указанной в пункте <i>\"Заменяемая подстрока\"</i>, на строку, указанную в пунтке <i>\"Заменяющая подстрока\"</i>.",
                                        "<b><c>Значение команд в пункте меню <i>\"Список\"</i>:</b>",
                                            "<l><i>\"Пронумеровать список\"</i> - нумерует список, руководствуясь маской в пункте меню <i>\"Настройки - Маска нумерации\"</i>. При необходимости корректной пятизначной нумерации обратитесь к разработчику. Разделителем нумерации является по умолчанию точка (\".\"). Для изменения разделителя следует перейти к пункту <i>\"Настройки - Разделитель нумерации\"</i> и поменять на необходимый.",
                                            "<l><i>\"Удалить нумерацию\"</i> - удаляет нумерацию, руководствуясь маской в пункте меню <i>\"Настройки - Маска нумерации\"</i>.",
                                            "<l><i>\"Вернуться к первоначальному виду\"</i> - отменяет все изменения пунктов списка и самого списка. Все удаленные пункты списка за время сеанса работы будут восстановлены. Максимальное количество хранимых записей равно 1000000.",
                                            "<l><i>\"Очистить список\"</i> (горячие клавиши Shift+Del) - очищает список.",
                                            "<l><b>Теги</b>",
                                            "#Time - применяется в полях <i>\"Добавить при заполнении списка в начало\"</i>, <i>\"Добавить при заполнении списка в конец\"</i>, находящихся во вкладке <i>\"Дополнительно\"</i>. При его наличии ко всем пунктам списка в начало или (и) конец будет добавлено время начала композиции. Формат времени указывается в пунтке <i>\"Настройки - Формат времени\"</i>.",
                                            "<b>Обновления:</b>",
                                            "<b>Сборка 3.8:</b> исправлено пропадание окна редактирования наименования пункта списка, исправлена загрузка cue-файла (теперь его имя будет отображаться в Дополнительно - Наименование списка, а не в начале списка), добавлены \"горячие клавиши\" на все виды сохранения.",
                                            "<b>Сборка 3.9:</b> исправлен поиск имя исполнителя и наименование трека",
                                            "<b>Сборка 3.10:</b> добавлена оценка треков через контекстное меню и с помощью горячих клавиш",
                                            "<b>Сборка 3.11:</b> Добавлена библиотека Taglib-sharp, с помощью которой представляется информация о файлах (надо будет ещё битрейт выводить). Теперь есть возможность открывать всё содержимое любой папки.",
                                            "<r><i>Приложение " + Text + " будет и дальше совершенствоваться. </i>",
                                            "<r><i>Следите за обновлениями.</i>",
                                            "<r><i>Рад конструктивной критике.</i>",
                                            "<r><i>Разработчик.</i> "
                };
                Manual m = new Manual();
                m.downloadText(content);
                m.Text = Text + " - " + m.Text;
                m.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Ошибка открытия мануала");
            }
        }

        private void leftTime_CheckedChanged(object sender, EventArgs e)
        {
            rebiuldTime();
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog savedialog = saveFileDialog1;
                savedialog.Title = "Сохранить как ...";
                savedialog.OverwritePrompt = true;
                savedialog.CheckPathExists = true;
                savedialog.Filter =
                    "Cue (*.cue)|*.cue|Все файлы(*.*)|*.*";
                savedialog.ShowHelp = true;
                char[] sep = { '\\' };
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = savedialog.FileName;
                    saveToCue(fileName);
                    textBoxChange.Text = "Данные сохранены в файл " + shortName(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog savedialog = saveFileDialog1;
                savedialog.Title = "Сохранить как ...";
                savedialog.OverwritePrompt = true;
                savedialog.CheckPathExists = true;
                savedialog.Filter =
                    "Playlist (*.m3u)|*.m3u|Все файлы(*.*)|*.*";
                savedialog.ShowHelp = true;
                char[] sep = { '\\' };
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = savedialog.FileName;
                    saveToM3u(fileName);
                    textBoxChange.Text = "Данные сохранены в файл " + shortName(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog savedialog = saveFileDialog1;
                savedialog.Title = "Сохранить как ...";
                savedialog.OverwritePrompt = true;
                savedialog.CheckPathExists = true;
                savedialog.Filter =
                    "PromoDJ Cue (*.pue)|*.pue|Все файлы(*.*)|*.*";
                savedialog.ShowHelp = true;
                char[] sep = { '\\' };
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = savedialog.FileName;
                    saveToPromodjCue(fileName);
                    textBoxChange.Text = "Данные сохранены в файл " + shortName(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void toolStripMenuItem21_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog savedialog = saveFileDialog1;
                savedialog.Title = "Сохранить как ...";
                savedialog.OverwritePrompt = true;
                savedialog.CheckPathExists = true;
                savedialog.Filter =
                    "Текстовые файлы(*.txt)|*.txt|Все файлы(*.*)|*.*";
                savedialog.ShowHelp = true;
                char[] sep = { '\\' };
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = savedialog.FileName;
                    saveToTxt(fileName);
                    textBoxChange.Text = "Данные сохранены в файл " + shortName(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void клёвыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBoxName.SelectedIndex != -1)
                listBoxName.Items[listBoxName.SelectedIndex] = "!" + listBoxName.Items[listBoxName.SelectedIndex];
        }

        private void мощныйБассToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBoxName.SelectedIndex != -1)
                listBoxName.Items[listBoxName.SelectedIndex] = "%" + listBoxName.Items[listBoxName.SelectedIndex];
        }

        private void прекрасныйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBoxName.SelectedIndex != -1)
                listBoxName.Items[listBoxName.SelectedIndex] = "&" + listBoxName.Items[listBoxName.SelectedIndex];
        }

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            copyList();
        }

        private void copyList()
        {
            string s = "";
            foreach (string ss in listBoxName.Items)
                s += ss + "\n";
            Clipboard.SetText(s);
        }
    }
}
