using System.Windows.Controls;
using System.Windows;
using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.sldworks;
using CADBooster.SolidDna;
using Newtonsoft.Json;
using System;
using System.Diagnostics;

using System.Windows.Media;

namespace SolidDNA.Memo.AddIn
{
    /// <summary>
    /// Interaction logic for MyAddinControl.xaml
    /// </summary>
    public partial class MyAddinControl : UserControl
    {
        public bool universalNote = false;
        public bool blockGenerator = false;             // Zastavica generiranja Blocka iz beležke

        public string assamblyPath;                     // Direktorij programa
        public static string swDocName;                 // Ime trenutno odprtega Solidworks dokumenta
        public static string swDocPath;                 // Direktorij trenutno odprtega Solidworks dokumenta
        public static string fileNotePath;              // Default Direktorij za iskanje beležk

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        /// 
        SolidWorksApplication ActiveSwApp;

        public MyAddinControl()
        {
            // Najdi lokacijo programa in njegovih datotek
            assamblyPath = GetAssamblyPath();

            // Inicializacija okna
            InitializeComponent();
            // Iz cfg datoteke razberi lokacijo za shranjevanje beležk
            fileNotePath = ReadCfgFile();

            // Postavi in odreagiraj na dogodke v SW
            ActiveSwApp = SolidWorksEnvironment.Application;
            SwEventsInit(ActiveSwApp);

            FileLocationBox.Text = fileNotePath;
            if (fileNotePath.IsNullOrEmpty())
            {
                NotepadBox.IsEnabled = false;
                FileNameBox.IsEnabled = false;
                SaveNotepad_btn.IsEnabled = false;
                SwitchNotes_btn.IsEnabled = false;
                Timestamp_btn.IsEnabled = false;
                MakeDrawingBlock_btn.IsEnabled = false;
            }
            UpdateLayout();
        }

        #endregion

        #region SOLIDWORKS EWENTS HANDLING

        /// <summary>
        /// Funkcije ki se izvedejo ob Solidworksovih eventih
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void SwEventsInit(SolidWorksApplication swApp)
        {        

            swApp.ActiveModelInformationChanged += OnActiveModelInformationChanged;
            swApp.FileCreated += OnFileCreated;
            swApp.ActiveFileSaved += OnActiveFileSaved;

        }

        private void OnFileCreated(Model obj)
        {
            OnActiveModelInformationChanged(obj);
        }

        private void OnActiveFileSaved(string arg1, Model arg2)
        {
            string swDocNameFull = Path.GetFileName(arg1);
            swDocName = swDocNameFull.Remove(swDocNameFull.Length - 7, 7);
            FileNameBox.Text = swDocName;

            if (UseNotepadBox.IsChecked == true)
            {
                SaveNotepad(swDocName);
            }
            FileNameBox.Text = swDocName;

        }

        private void OnActiveModelInformationChanged(Model arg2)
        {

            // ČE se funkcija izvede md generiranjem drawing bloka, jo preskoči
            if (blockGenerator != true)
            {
  
                string swDocNameFull;

                // Preveri tip datoteke, da nastaviš pravilen event notify
                if (arg2 != null)
                {
                    swDocNameFull = Path.GetFileName(arg2.FilePath);
                    if (swDocNameFull.Length > 0) { swDocName = swDocNameFull.Remove(swDocNameFull.Length - 7, 7); }
                    

                    if (arg2.IsPart)
                    {
                        var swPart = arg2.AsPart();
                        if (swPart != null)
                        {

                            swPart.DestroyNotify += SwPart_DestroyNotify;
                            swPart = null;
                        }
                        else { MessageBox.Show("Failed to open part properly!"); }
                    }
                    else if (arg2.IsAssembly)
                    {
                        var swAssembly = arg2.AsAssembly();
                        if (swAssembly != null)
                        {
                            swAssembly.DestroyNotify += OnAsemmblyDestroyNotify;
                            swAssembly.NewSelectionNotify += OnAsemblySelectionChanged;
                            swAssembly = null;
                        }
                        else { MessageBox.Show("Failed to open assambly properly!"); }

                    }
                    else if (arg2.IsDrawing)
                    {
                        var swDrawing = arg2.AsDrawing();
                       
                        if (swDrawing != null)
                        {
                            swDrawing.DestroyNotify += OnDrawingDestroyNotify;
                            swDrawing = null;
                        }
                        else { MessageBox.Show("Failed to open drawing properly!"); }
                    }

                    
                   
                }
                LoadNotepad(swDocName);
                blockGenerator = false;
            }
        }

        private int OnAsemblySelectionChanged()
        {
 
            ModelDoc2 swModel = null;
            Model sWorks = ActiveSwApp.ActiveModel;
            if(sWorks.IsAssembly == true) {
                AssemblyDoc swAsm = sWorks.AsAssembly();
                swModel = swAsm as ModelDoc2;
            }

            if (swModel != null)
            {
                SelectionMgr swSelectionMgr = (SelectionMgr)swModel.SelectionManager;
                Component2 selectedComponent = null;
                try { selectedComponent = (Component2)swSelectionMgr.GetSelectedObject6(1, -1); }
                catch { }
                if (selectedComponent != null)
                {

                    string componentName = selectedComponent.Name;

                    // poišči znak "-" pri part-n
                    char line = componentName[componentName.Length - 2];
                    string a = line.ToString();
                    if (a.Equals("-"))
                    {
                        componentName = componentName.Remove(componentName.Length - 2, 2);
                    }
                    // poišči znak "-" pri part-nn
                    line = componentName[componentName.Length - 3];
                    a = line.ToString();
                    if (a.Equals("-"))
                    {
                        componentName = componentName.Remove(componentName.Length - 3, 3);
                    }

                    LoadNotepad(componentName);
                }



            }

            return 1;
        }

        private int SwPart_DestroyNotify()
        {
            SaveOnDestroy();
            FileNameBox.Text = null;
            NotepadBox.Text = null;
            UpdateLayout();
            return 1;
        }

        private int OnDrawingDestroyNotify()
        {
            SaveOnDestroy();
            FileNameBox.Text = null;
            NotepadBox.Text = null;
            UpdateLayout();
            return 1;
        }

        private int OnAsemmblyDestroyNotify()
        {
            SaveOnDestroy();
            FileNameBox.Text = null;
            NotepadBox.Text = null;
            UpdateLayout();
            return 1;
        }

        /// <summary>
        /// Shrani beležko ko uporabnik zapre datoteko
        /// </summary>


        #endregion

        #region USER INPUT HANDLING
        /// <summary>
        /// Dogodki, ki jih sproži uporabnik
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void SaveNotepad_btn_Click(object sender, RoutedEventArgs e)
        {
            if (FileNameBox.Text.Length > 0)
            {
                SaveNotepad(FileNameBox.Text);
                MessageBox.Show("Notepad saved.");
            }
            else
            {
                MessageBox.Show("Please give your notepad a name.");
            }
           
        }

        private void DirectoryBtn_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string json;
            /// Dialog box za izbiro direktorija za shranjevanje beležk
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            try
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result.ToString() == "OK")
                {
                    FileLocationBox.Text = dialog.SelectedPath;
                    fileNotePath = dialog.SelectedPath;

                    using (StreamReader sr = File.OpenText(assamblyPath + "\\" + "mtrConfig.json"))
                    {
                        var myObject = JsonConvert.DeserializeObject<List<RootObject>>(sr.ReadToEnd());
                        myObject[0].Pot = fileNotePath;
                        json = JsonConvert.SerializeObject(myObject.ToArray());

                        //write string to file


                    }
                    if (json != null)
                    {
                        File.WriteAllText(assamblyPath + "\\" + "mtrConfig.json", json);
                        NotepadBox.IsEnabled = true;
                        FileNameBox.IsEnabled = true;
                        SaveNotepad_btn.IsEnabled = true;
                        SwitchNotes_btn.IsEnabled = true;
                        Timestamp_btn.IsEnabled = true;
                        MakeDrawingBlock_btn.IsEnabled = true;
                        UpdateLayout();
                    }
                    else
                    {
                        MessageBox.Show("mtrConfig.json file missing from add-in root folder");
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Directory could not be set, because of:"+ex.ToString());
            }
        }

        private void FileNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {
           
        }

        private void FileLocationBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (FileLocationBox.Text != "")
            {
                fileNotePath = FileLocationBox.Text;
            }
        }

        private void RefreshNotes_btn_Click(object sender, RoutedEventArgs e)
        {
            LoadNotepad(FileNameBox.Text);
        }

        private void NotepadBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void UseNotepadBox_Checked(object sender, System.Windows.RoutedEventArgs e)
        {

        }

        private void CollectNotes_btn_Click(object sender, RoutedEventArgs e)
        {


            /// ModelDoc2 je v SOlidDNA = Part doc
            /// Tukaj pridobiš aktiven model
            Model model = ActiveSwApp.ActiveModel;
            
            if (model != null)
            {

                if (model.IsPart)
                {
                    string preeText = "\n\n >>>>>>>>>>>>>>>> IMPORTED COMMENTS <<<<<<<<<<<<<<<<\n\n";
                    string lines = NotepadBox.Text + preeText;

                    PartDoc part = model.AsPart();
                    ModelDoc2 swModelDoc = part as ModelDoc2;

                    /// Poišči feature v modelu
                    Feature swFeat = (Feature)swModelDoc.FirstFeature();

                    /// Preveri če model vsebuje komentarje
                    CommentFolder swCommentFolder = (CommentFolder)swFeat.GetSpecificFeature2();
                    int nbrComments = swCommentFolder.GetCommentCount();

                    object[] vComments = (object[])swCommentFolder.GetComments();
                    if (nbrComments > 0)
                    {
                        for (int i = 0; i <= (nbrComments - 1); i++)
                        {
                            Comment swComment = (Comment)vComments[i];
                            lines = lines + swComment.Name + ": " + swComment.Text + "\n";
                        }

                        NotepadBox.Text = lines;
                        UpdateLayout();

                    }
                    else
                    {
                        MessageBox.Show("Model contains no coments.");
                    }
                }
            }
            else
            {
                MessageBox.Show("There is no active model.");
            }

        }

        /// <summary>
        /// Zamenjaj med beležko ki je vezana na model in univerzalno beležko
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        
        private void SwitchNotes_btn_Click(object sender, RoutedEventArgs e)
        {


            var curentNote = FileNameBox.Text;
            SaveNotepad(curentNote);
            if (universalNote == false)
            {

                universalNote = true;
                NotepadBox.BorderBrush = Brushes.Green;
                LoadNotepad("Universal");
            }
            else
            {
                NotepadBox.BorderBrush = Brushes.Red;
                universalNote = false;
                LoadNotepad(swDocName);
            }
        }

        private void MakeBlock_btn_Click(object sender, RoutedEventArgs e)
        {
            blockGenerator = true;

            int errors = 0;


            IModelDoc2 swModelDoc = null;
            ISldWorks sWorks = ActiveSwApp.UnsafeObject;
            string blankFilePath = assamblyPath + "\\Resources\\blank.SLDDRW";

            if (NotepadBox.Text.Length > 0)
            {
                if (FileNameBox.Text.Length > 0)
                {

                    try
                    {
                        swModelDoc = sWorks.IOpenDoc5(blankFilePath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors);
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("Loading file !FAILED: " + "\n" + blankFilePath + "\n" + "Install may be corupted.");
                        MessageBox.Show(ex.Message);
                    }

                    if (swModelDoc != null)
                    {
                        // Po defoltu je direktorij bloka enak direktoriju beležke
                        string blockDir = fileNotePath;          
                        string potModela = swModelDoc.GetPathName();
                        swDocName = FileNameBox.Text;

                        IDrawingDoc drawing = (IDrawingDoc)swModelDoc;



                        /// Naredi drawing in na njega pripopaj besedilo
                        drawing.ICreateText2(NotepadBox.Text, 0, 0, 0, 0.0024, 0);

                        // poišči ID tega besedila in ga izberi
                       
                        swModelDoc.SelectAt(0, 0, 0, 0);
                        

                        SketchManager swSketchMgr = swModelDoc.SketchManager;
                        
                        SketchBlockDefinition swSketchBlockDef = swSketchMgr.MakeSketchBlockFromSelected(null);

                        // Če je odprta datoteka, shrani blok k datoteki

                        
                        
                        bool result = swSketchBlockDef.Save(blockDir +"\\"+ swDocName + ".SLDBLK");
                        
                        if (result == true)
                        {
                            MessageBox.Show("Block saved to " + blockDir);
                        }
                        else
                        {
                            MessageBox.Show("Saving block failed!");
                        }

                        swModelDoc.ClearSelection();
                        drawing.DeleteAllCosmeticThreads();
                        sWorks.CloseDoc(potModela);



                        // Popucaj za sabo
                        //Marshal.ReleaseComObject(sWorks);
                        ActiveSwApp.Dispose();
                       
                        
                    }
                }
                else
                {
                    MessageBox.Show("File name window is empty. Save your model or insert a name.");
                }
            
                                    
            }
            else
            {
                MessageBox.Show("Empty block can not be created. Please add text to your notepad.");
            }
            blockGenerator = false;
        }

        private void EditTemplate_btn_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("notepad.exe", assamblyPath+"\\template.txt");
        }

        /// <summary>
        /// Dodaj datum in čas v beležko
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timestamp_btn_Click(object sender, RoutedEventArgs e)
        {
            var dt = DateTime.Now;
            var dateTime = dt.ToString("dd/MM/yyyy HH:mm:ss");
            NotepadBox.Text = NotepadBox.Text + "\n" + dateTime;

        }

        #endregion

        #region SUPORT METHODS
        /// <summary>
        /// Metode s katerimi delamo različne stvart
        /// </summary>
        /// <returns></returns>
        private string ReadCfgFile()
        {
            // Nastavitve
            string defaultPath = "";
            using (StreamReader sr = File.OpenText(assamblyPath + "\\" + "mtrConfig.json"))
            {
                var myObject = JsonConvert.DeserializeObject<List<RootObject>>(sr.ReadToEnd());
                defaultPath = myObject[0].Pot;
            }
            return defaultPath;
        }
        private string GetAssamblyPath()
        {
            // Pot do install folderja
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = Path.GetDirectoryName(path);

            return path;
        }

        public class RootObject
        {
            public string Pot { get; set; }

        }

        /// <summary>
        /// Posodobitev podatkov v vseh oknih beležke.
        /// FileName - Ime beležke, ki jo želiš prikazati
        /// </summary>
        private void LoadNotepad(string FileName)
        {

            string lines="";          // String z besedilom za v beležko
            bool NoteExist = false;     // Zastavica ki pove ali beležka obstaja
            string[] fileEntries;       // Seznam datotek v direktoriju
            string FullName;            // Direktorij + ime beležke + končnica

            

            if (FileName != "File Name")
            {
                if (!fileNotePath.IsNullOrEmpty())
                {
                    // 1. Preveri direktorij ali beležka obstaja
                    FullName = fileNotePath + "\\" + FileName + ".txt";
                    fileEntries = Directory.GetFiles(fileNotePath);

                    foreach (string fileEntry in fileEntries)
                    {
                        if (fileEntry == FullName)
                        {
                            lines = File.ReadAllText(FullName);
                            NoteExist = true;
                            FileNameBox.Text = FileName;
                            NotepadBox.Text = lines;
                            break;
                        }

                    }

                    // 2. Če obstaja naloži obstoječo, drugače naloži template.
                    if (NoteExist == false)
                    {
                        FullName = assamblyPath + "\\" + "template" + ".txt";
                        lines = File.ReadAllText(FullName);
                        NotepadBox.Text = lines;
                        FileNameBox.Text = FileName;
                    }
                }
                else
                {
                    MessageBox.Show("Please specify directory with notes.");
                }
               
               
                
            }
            else
            {

                FullName = assamblyPath + "\\" + "template" + ".txt";
                lines = File.ReadAllText(FullName);
                NotepadBox.Text = lines;
            }
            UpdateLayout();
        }

        /// <summary>
        /// Shrani beležko ODPRTEGA dokumenta - če je odprt assambly to shrani samo beležko assamblija
        /// </summary>
        private void SaveOnDestroy()
        {
            // Shrani samo, če je obljukana uporaba beležke
            if (UseNotepadBox.IsChecked == true)
            {
                string fullName = fileNotePath + "\\" + swDocName + ".txt";
                File.WriteAllText(fullName, NotepadBox.Text);
            }

        }

        /// <summary>
        /// Določi beležko ki jo želiš shraniti.
        /// </summary>
        /// <param name="FileName"></param>
        private void SaveNotepad(string FileName)
        {

            string fullName = fileNotePath + "\\" + FileName + ".txt";
            try
            {
                File.WriteAllText(fullName, NotepadBox.Text);

            }
            catch { MessageBox.Show("Saving notepad FAILED."); }
            

            
        }








        #endregion


    }
}
